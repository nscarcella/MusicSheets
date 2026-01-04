import { describe, it, expect, vi } from "vitest"

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// MOCKS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

type MockRange = GoogleAppsScript.Spreadsheet.Range
type MockSheet = GoogleAppsScript.Spreadsheet.Sheet

const mockSheetProto = {}
const mockRangeProto = {}

const mockSheet: MockSheet = Object.create(mockSheetProto)
const mockRange: MockRange = Object.create(mockRangeProto)

mockSheet.getRange = (): ReturnType<MockSheet["getRange"]> => mockRange

vi.stubGlobal("SpreadsheetApp", {
  getActiveSpreadsheet: () => ({
    getSheets: () => [mockSheet]
  })
})


await import("../src/Range")


function createMockSheet(name: string, maxRows: number, maxCols: number, frozenRows: number = 0, frozenCols: number = 0): MockSheet {
  const ranges = new Map<string, MockRange>()

  const sheet: MockSheet = Object.create(mockSheetProto)

  sheet.getName = () => name
  sheet.getMaxRows = () => maxRows
  sheet.getMaxColumns = () => maxCols
  sheet.getFrozenRows = () => frozenRows
  sheet.getFrozenColumns = () => frozenCols
  sheet.getRange = (rowOrA1: number | string, col?: number, numRows?: number, numCols?: number): MockRange => {
    if (typeof rowOrA1 === "string") throw new Error("A1 notation not implemented in mock")

    const row = rowOrA1
    const column = col!
    const rows = numRows ?? 1
    const cols = numCols ?? 1
    const key = `${row},${column},${rows},${cols}`

    if (!ranges.has(key)) ranges.set(key, createMockRange(sheet, row, column, rows, cols))

    return ranges.get(key)!
  }

  return sheet
}


function createMockRange(sheet: MockSheet, row: number, col: number, numRows: number, numCols: number): MockRange {
  const range: MockRange = Object.create(mockRangeProto)

  range.getSheet = () => sheet
  range.getRow = () => row
  range.getColumn = () => col
  range.getNumRows = () => numRows
  range.getNumColumns = () => numCols
  range.getLastRow = () => row + numRows - 1
  range.getLastColumn = () => col + numCols - 1
  range.getA1Notation = () => `R${row}C${col}:R${row + numRows - 1}C${col + numCols - 1}`

  return range
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// TESTS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

describe("Sheet extensions", () => {
  describe("fullRange", () => {
    it("should return range covering entire sheet", () => {
      const sheet = createMockSheet("Test", 100, 50)
      const range = sheet.fullRange()

      expect(range.getRow()).toBe(1)
      expect(range.getColumn()).toBe(1)
      expect(range.getNumRows()).toBe(100)
      expect(range.getNumColumns()).toBe(50)
    })

    it("should work with different sheet sizes", () => {
      const sheet = createMockSheet("Large", 1000, 26)
      const range = sheet.fullRange()

      expect(range.getNumRows()).toBe(1000)
      expect(range.getNumColumns()).toBe(26)
    })
  })
})

describe("Range extensions", () => {
  const sheet = createMockSheet("Test", 1000, 1000)

  describe("translate", () => {
    it("should move range by offset", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const moved = range.translate(3, 2)

      expect(moved.getRow()).toBe(7)
      expect(moved.getColumn()).toBe(8)
      expect(moved.getNumRows()).toBe(10)
      expect(moved.getNumColumns()).toBe(10)
    })

    it("should handle negative offsets", () => {
      const range = sheet.getRange(10, 10, 5, 5)
      const moved = range.translate(-3, -2)

      expect(moved.getRow()).toBe(8)
      expect(moved.getColumn()).toBe(7)
    })

    it("should handle zero offset", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const moved = range.translate(0, 0)

      expect(moved.getRow()).toBe(5)
      expect(moved.getColumn()).toBe(5)
    })

    it("should throw on out of bounds translation", () => {
      const range = sheet.getRange(2, 2, 5, 5)
      expect(() => range.translate(-2, 0)).toThrow("Translation out of bounds")
      expect(() => range.translate(0, -2)).toThrow("Translation out of bounds")
    })
  })

  describe("translateTo", () => {
    it("should move range to absolute position", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const moved = range.translateTo(20, 15)

      expect(moved.getRow()).toBe(15)
      expect(moved.getColumn()).toBe(20)
      expect(moved.getNumRows()).toBe(10)
      expect(moved.getNumColumns()).toBe(10)
    })

    it("should throw on invalid coordinates", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      expect(() => range.translateTo(0, 5)).toThrow("Translation out of bounds")
      expect(() => range.translateTo(5, 0)).toThrow("Translation out of bounds")
      expect(() => range.translateTo(-1, 5)).toThrow("Translation out of bounds")
    })
  })

  describe("scale", () => {
    it("should scale range dimensions", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const scaled = range.scale(2, 3)

      expect(scaled.getRow()).toBe(5)
      expect(scaled.getColumn()).toBe(5)
      expect(scaled.getNumRows()).toBe(30)
      expect(scaled.getNumColumns()).toBe(20)
    })

    it("should round up fractional results", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const scaled = range.scale(1.5, 1.5)

      expect(scaled.getNumRows()).toBe(15)
      expect(scaled.getNumColumns()).toBe(15)
    })

    it("should handle scale of 1", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const scaled = range.scale(1, 1)

      expect(scaled.getNumRows()).toBe(10)
      expect(scaled.getNumColumns()).toBe(10)
    })

    it("should throw on invalid multipliers", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      expect(() => range.scale(0, 1)).toThrow("Invalid scale multiplier")
      expect(() => range.scale(1, 0)).toThrow("Invalid scale multiplier")
      expect(() => range.scale(-1, 1)).toThrow("Invalid scale multiplier")
    })
  })

  describe("resize", () => {
    it("should resize range by offset", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const resized = range.resize(5, 3)

      expect(resized.getRow()).toBe(5)
      expect(resized.getColumn()).toBe(5)
      expect(resized.getNumRows()).toBe(13)
      expect(resized.getNumColumns()).toBe(15)
    })

    it("should handle negative offsets", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const resized = range.resize(-5, -3)

      expect(resized.getNumRows()).toBe(7)
      expect(resized.getNumColumns()).toBe(5)
    })

    it("should throw on invalid dimensions", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      expect(() => range.resize(-10, 0)).toThrow("non-positive dimensions")
      expect(() => range.resize(0, -10)).toThrow("non-positive dimensions")
      expect(() => range.resize(-100, -100)).toThrow("non-positive dimensions")
    })
  })

  describe("resizeTo", () => {
    it("should resize range to absolute dimensions", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const resized = range.resizeTo(20, 15)

      expect(resized.getRow()).toBe(5)
      expect(resized.getColumn()).toBe(5)
      expect(resized.getNumRows()).toBe(15)
      expect(resized.getNumColumns()).toBe(20)
    })

    it("should throw on invalid dimensions", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      expect(() => range.resizeTo(0, 5)).toThrow("positive dimensions")
      expect(() => range.resizeTo(5, 0)).toThrow("positive dimensions")
      expect(() => range.resizeTo(-1, 5)).toThrow("positive dimensions")
    })
  })

  describe("projectInto", () => {
    it("should project range into different sheet", () => {
      const sheet1 = createMockSheet("Sheet1", 100, 100)
      const sheet2 = createMockSheet("Sheet2", 200, 200)
      const range = sheet1.getRange(5, 5, 10, 10)
      const projected = range.projectInto(sheet2)

      expect(projected.getSheet().getName()).toBe("Sheet2")
      expect(projected.getRow()).toBe(5)
      expect(projected.getColumn()).toBe(5)
      expect(projected.getNumRows()).toBe(10)
      expect(projected.getNumColumns()).toBe(10)
    })
  })

  describe("intersect", () => {
    it("should return intersection of overlapping ranges", () => {
      const range1 = sheet.getRange(5, 5, 10, 10)
      const range2 = sheet.getRange(10, 10, 10, 10)
      const intersection = range1.intersect(range2)

      expect(intersection).toBeDefined()
      expect(intersection!.getRow()).toBe(10)
      expect(intersection!.getColumn()).toBe(10)
      expect(intersection!.getLastRow()).toBe(14)
      expect(intersection!.getLastColumn()).toBe(14)
    })

    it("should return undefined for non-overlapping ranges", () => {
      const range1 = sheet.getRange(5, 5, 5, 5)
      const range2 = sheet.getRange(20, 20, 5, 5)
      const intersection = range1.intersect(range2)

      expect(intersection).toBeUndefined()
    })

    it("should handle adjacent ranges", () => {
      const range1 = sheet.getRange(5, 5, 5, 5)
      const range2 = sheet.getRange(10, 10, 5, 5)
      const intersection = range1.intersect(range2)

      expect(intersection).toBeUndefined()
    })

    it("should handle fully contained ranges", () => {
      const range1 = sheet.getRange(5, 5, 20, 20)
      const range2 = sheet.getRange(10, 10, 5, 5)
      const intersection = range1.intersect(range2)

      expect(intersection).toBeDefined()
      expect(intersection!.getRow()).toBe(10)
      expect(intersection!.getColumn()).toBe(10)
      expect(intersection!.getNumRows()).toBe(5)
      expect(intersection!.getNumColumns()).toBe(5)
    })
  })

  describe("overlapsWith", () => {
    it("should return true for overlapping ranges", () => {
      const range1 = sheet.getRange(5, 5, 10, 10)
      const range2 = sheet.getRange(10, 10, 10, 10)

      expect(range1.overlapsWith(range2)).toBe(true)
      expect(range2.overlapsWith(range1)).toBe(true)
    })

    it("should return false for non-overlapping ranges", () => {
      const range1 = sheet.getRange(5, 5, 5, 5)
      const range2 = sheet.getRange(20, 20, 5, 5)

      expect(range1.overlapsWith(range2)).toBe(false)
    })

    it("should return false for adjacent ranges", () => {
      const range1 = sheet.getRange(5, 5, 5, 5)
      const range2 = sheet.getRange(10, 10, 5, 5)

      expect(range1.overlapsWith(range2)).toBe(false)
    })
  })

  describe("print", () => {
    it("should return readable string representation", () => {
      const range = sheet.getRange(5, 5, 10, 10)
      const str = range.print()

      expect(str).toContain("Test")
      expect(str).toContain("10×10")
    })
  })

  describe("chaining", () => {
    it("should support chaining multiple operations", () => {
      const range = sheet.getRange(1, 1, 10, 10)
      const result = range
        .translate(5, 5)
        .resize(10, 10)
        .scale(2, 2)

      expect(result.getRow()).toBe(6)
      expect(result.getColumn()).toBe(6)
      expect(result.getNumRows()).toBe(40)
      expect(result.getNumColumns()).toBe(40)
    })
  })
})
