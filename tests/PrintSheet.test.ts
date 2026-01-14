import { describe, it, expect, vi } from "vitest"

const mockSheetProto = {}
const mockRangeProto = {}
const mockSheet = Object.create(mockSheetProto)
const mockRange = Object.create(mockRangeProto)
mockSheet.getRange = () => mockRange

vi.stubGlobal("SpreadsheetApp", {
  getActiveSpreadsheet: () => ({ getSheets: () => [mockSheet] }),
  getActive: () => ({ getSheetByName: () => mockSheet })
})

const { splitIntoSectionColumns, splitIntoSections, calculateSectionWidth, calculateLayout } = await import("../src/MusicSheet")

type Range = GoogleAppsScript.Spreadsheet.Range

function mockSection(width: number, height: number): Range {
  return {
    getNumColumns: () => width,
    getNumRows: () => height
  } as Range
}

describe("splitIntoSectionColumns", () => {
  it("should return empty array for empty input", () => {
    expect(splitIntoSectionColumns([[]])).toEqual([])
  })

  it("should return single column when no lyric boundaries found", () => {
    const values = [
      ["C", "", "", ""],
      ["Hello world", "", "", ""]
    ]
    const result = splitIntoSectionColumns(values)
    expect(result).toHaveLength(1)
    expect(result[0]).toEqual(values)
  })

  it("should split on lyric start in odd row", () => {
    const values = [
      ["C", "", "", "G"],
      ["Hello", "", "World", ""]
    ]
    const result = splitIntoSectionColumns(values)
    expect(result).toHaveLength(2)
    expect(result[0]).toEqual([["C", ""], ["Hello", ""]])
    expect(result[1]).toEqual([["", "G"], ["World", ""]])
  })

  it("should handle three section columns", () => {
    const values = [
      ["C", "", "G", "", "", "Am"],
      ["One", "", "Two", "", "Three", ""]
    ]
    const result = splitIntoSectionColumns(values)
    expect(result).toHaveLength(3)
    expect(result[0]).toEqual([["C", ""], ["One", ""]])
    expect(result[1]).toEqual([["G", ""], ["Two", ""]])
    expect(result[2]).toEqual([["", "Am"], ["Three", ""]])
  })

  it("should not split on chord-only columns (even rows)", () => {
    const values = [
      ["C", "D", "", ""],
      ["Hello", "", "", ""]
    ]
    const result = splitIntoSectionColumns(values)
    expect(result).toHaveLength(1)
  })

  it("should discard leading empty columns", () => {
    const values = [
      ["", "", "C", ""],
      ["", "", "Hello", ""]
    ]
    const result = splitIntoSectionColumns(values)
    expect(result).toHaveLength(1)
    expect(result[0]).toEqual([["C", ""], ["Hello", ""]])
  })

  it("should handle multiple row pairs", () => {
    const values = [
      ["C", "", "", "G"],
      ["Line 1", "", "Chorus 1", ""],
      ["", "D", "Am", ""],
      ["Line 2", "", "Chorus 2", ""]
    ]
    const result = splitIntoSectionColumns(values)
    expect(result).toHaveLength(2)
    expect(result[0]).toEqual([
      ["C", ""],
      ["Line 1", ""],
      ["", "D"],
      ["Line 2", ""]
    ])
    expect(result[1]).toEqual([
      ["", "G"],
      ["Chorus 1", ""],
      ["Am", ""],
      ["Chorus 2", ""]
    ])
  })
})

describe("splitIntoSections", () => {
  it("should return empty array for all-empty input", () => {
    const values = [
      ["", ""],
      ["", ""]
    ]
    expect(splitIntoSections(values)).toEqual([])
  })

  it("should return single section for continuous content", () => {
    const values = [
      ["C", "", ""],
      ["Hello", "", ""],
      ["", "D", ""],
      ["World", "", ""]
    ]
    const result = splitIntoSections(values)
    expect(result).toHaveLength(1)
    expect(result[0]).toEqual({ startRow: 0, endRow: 4, width: 3 })
  })

  it("should split on empty row pair", () => {
    const values = [
      ["C", ""],
      ["Hello", ""],
      ["", ""],
      ["", ""],
      ["", "D"],
      ["World", ""]
    ]
    const result = splitIntoSections(values)
    expect(result).toHaveLength(2)
    expect(result[0]).toEqual({ startRow: 0, endRow: 2, width: 3 })
    expect(result[1]).toEqual({ startRow: 4, endRow: 6, width: 3 })
  })

  it("should skip leading empty pairs", () => {
    const values = [
      ["", ""],
      ["", ""],
      ["", "C"],
      ["Hello", ""]
    ]
    const result = splitIntoSections(values)
    expect(result).toHaveLength(1)
    expect(result[0]).toEqual({ startRow: 2, endRow: 4, width: 3 })
  })

  it("should skip trailing empty pairs", () => {
    const values = [
      ["C", ""],
      ["Hello", ""],
      ["", ""],
      ["", ""]
    ]
    const result = splitIntoSections(values)
    expect(result).toHaveLength(1)
    expect(result[0]).toEqual({ startRow: 0, endRow: 2, width: 3 })
  })

  it("should handle chord-only row as non-empty", () => {
    const values = [
      ["", "C"],
      ["", ""],
      ["", ""],
      ["", ""],
      ["D", ""],
      ["Hello", ""]
    ]
    const result = splitIntoSections(values)
    expect(result).toHaveLength(2)
    // C at col 1: end = 1 + 1 + ceil(1/2) = 3
    expect(result[0]).toEqual({ startRow: 0, endRow: 2, width: 3 })
    // D at col 0: end = 2, lyric "Hello" = 3
    expect(result[1]).toEqual({ startRow: 4, endRow: 6, width: 3 })
  })

  it("should handle lyric-only row as non-empty", () => {
    const values = [
      ["", ""],
      ["Hello", ""],
      ["", ""],
      ["", ""],
      ["", "D"],
      ["World", ""]
    ]
    const result = splitIntoSections(values)
    expect(result).toHaveLength(2)
    expect(result[0]).toEqual({ startRow: 0, endRow: 2, width: 3 })
    expect(result[1]).toEqual({ startRow: 4, endRow: 6, width: 3 })
  })

  it("should handle three sections", () => {
    const values = [
      ["C", ""],
      ["One", ""],
      ["", ""],
      ["", ""],
      ["", "D"],
      ["Two", ""],
      ["", ""],
      ["", ""],
      ["E", ""],
      ["Three", ""]
    ]
    const result = splitIntoSections(values)
    expect(result).toHaveLength(3)
    // C at col 0: end = 2, lyric "One" = 2
    expect(result[0]).toEqual({ startRow: 0, endRow: 2, width: 2 })
    // D at col 1: end = 1 + 1 + 1 = 3, lyric "Two" = 2
    expect(result[1]).toEqual({ startRow: 4, endRow: 6, width: 3 })
    // E at col 0: end = 2, lyric "Three" = 3
    expect(result[2]).toEqual({ startRow: 8, endRow: 10, width: 3 })
  })
})

describe("calculateSectionWidth", () => {
  it("should return 1 for empty section", () => {
    expect(calculateSectionWidth([[""]])).toBe(1)
  })

  it("should calculate width from lyric length (2 chars per cell)", () => {
    const section = [
      ["C", ""],
      ["Hello", ""]
    ]
    expect(calculateSectionWidth(section)).toBe(3)
  })

  it("should calculate width from chord position plus chord length", () => {
    const section = [
      ["", "", "", "C", ""],
      ["Hi", "", "", "", ""]
    ]
    // C at col 3, length 1: end = 3 + 1 + ceil(1/2) = 5
    expect(calculateSectionWidth(section)).toBe(5)
  })

  it("should use max of lyric width and chord position", () => {
    const section = [
      ["C", "", "", "", "", "", ""],
      ["Hello World!", "", "", "", "", "", ""]
    ]
    expect(calculateSectionWidth(section)).toBe(6)
  })

  it("should handle multiple row pairs and use max", () => {
    const section = [
      ["C", "D", "", ""],
      ["Hi", "", "", ""],
      ["", "", "", "G"],
      ["Hello World", "", "", ""]
    ]
    expect(calculateSectionWidth(section)).toBe(6)
  })

  it("should handle long lyric in first cell", () => {
    const section = [
      ["", "C"],
      ["This is a very long lyric line", ""]
    ]
    expect(calculateSectionWidth(section)).toBe(15)
  })

  it("should handle chord at far right with chord length", () => {
    const section = [
      ["", "", "", "", "", "", "", "", "", "Am"],
      ["Short", "", "", "", "", "", "", "", "", ""]
    ]
    // Am at col 9, length 2: end = 9 + 1 + ceil(2/2) = 11
    expect(calculateSectionWidth(section)).toBe(11)
  })

  it("should account for long chord names overflowing their cell", () => {
    const section = [
      ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "B#aug"],
      ["Short lyric", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
    ]
    // B#aug at col 20, length 5: end = 20 + 1 + ceil(5/2) = 24
    // Lyric "Short lyric" = ceil(11/2) = 6
    expect(calculateSectionWidth(section)).toBe(24)
  })
})

describe("calculateLayout", () => {
  it("should return empty array for no sections", () => {
    expect(calculateLayout([], 45, 50, 5)).toEqual([])
  })

  it("should place single section on first page", () => {
    const section = mockSection(10, 20)
    const result = calculateLayout([section], 45, 50, 5)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(1)
    expect(result[0][0]).toEqual([section])
  })

  it("should stack sections vertically in same column", () => {
    const s1 = mockSection(10, 10)
    const s2 = mockSection(10, 10)
    const result = calculateLayout([s1, s2], 45, 50, 5)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(1)
    expect(result[0][0]).toEqual([s1, s2])
  })

  it("should start new column when height exceeded", () => {
    const s1 = mockSection(10, 30)
    const s2 = mockSection(10, 20)
    const result = calculateLayout([s1, s2], 45, 50, 5)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2])
  })

  it("should start new page when width exceeded", () => {
    const s1 = mockSection(25, 30)
    const s2 = mockSection(25, 30)
    const result = calculateLayout([s1, s2], 45, 50, 5)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2])
  })

  it("should respect first page header height", () => {
    const s1 = mockSection(10, 44)
    const s2 = mockSection(10, 2)
    const result = calculateLayout([s1, s2], 45, 50, 5)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
  })

  it("should place sections on second page when first is full", () => {
    const s1 = mockSection(45, 45)
    const s2 = mockSection(10, 45)
    const result = calculateLayout([s1, s2], 45, 50, 5)
    expect(result).toHaveLength(2)
    expect(result[1][0]).toEqual([s2])
  })

  it("should have more space on second page (no header)", () => {
    // First page: 50 - 5 = 45 rows available
    // Second page: 50 rows available (no header)
    // s1 takes full first page width, height 45 (fits exactly on first page)
    // s2 is 48 rows - too tall for first page (45) but fits on second page (50)
    const s1 = mockSection(45, 45)
    const s2 = mockSection(10, 48)
    const result = calculateLayout([s1, s2], 45, 50, 5)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2])
  })

  it("should stack more sections on second page due to extra height", () => {
    // First page: 45 rows available (50 - 5 header)
    // Second page: 50 rows available
    // s1 (45 rows) fills first page
    // s2 + s3 (24 + 24 = 48 rows) should fit on second page but not on first
    const s1 = mockSection(45, 45)
    const s2 = mockSection(10, 24)
    const s3 = mockSection(10, 24)
    const result = calculateLayout([s1, s2, s3], 45, 50, 5)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2, s3])
  })

  it("should fill columns before starting new page", () => {
    const s1 = mockSection(15, 30)
    const s2 = mockSection(15, 30)
    const s3 = mockSection(15, 30)
    const result = calculateLayout([s1, s2, s3], 45, 50, 5)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(3)
  })

  it("should handle complex layout", () => {
    const s1 = mockSection(20, 25)
    const s2 = mockSection(20, 25)
    const s3 = mockSection(20, 20)
    const s4 = mockSection(20, 20)
    const result = calculateLayout([s1, s2, s3, s4], 45, 50, 5)
    expect(result).toHaveLength(2)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2, s3])
    expect(result[1][0]).toEqual([s4])
  })

  it("should account for horizontal padding between columns", () => {
    const s1 = mockSection(20, 30)
    const s2 = mockSection(20, 30)
    const result = calculateLayout([s1, s2], 45, 50, 5, 6, 0)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2])
  })

  it("should account for vertical padding between sections", () => {
    const s1 = mockSection(10, 22)
    const s2 = mockSection(10, 22)
    const result = calculateLayout([s1, s2], 45, 50, 5, 0, 5)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2])
  })

  it("should fit sections without padding when padding is zero", () => {
    const s1 = mockSection(20, 30)
    const s2 = mockSection(20, 30)
    const result = calculateLayout([s1, s2], 45, 50, 5, 0, 0)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
  })

  it("should not add padding before first column or first section", () => {
    const s1 = mockSection(43, 43)
    const result = calculateLayout([s1], 45, 50, 5, 2, 2)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(1)
    expect(result[0][0]).toEqual([s1])
  })

  it("should account for column width growth when checking if new column fits", () => {
    // Column 1: s1 (10w, 40h) then s2 (15w, 4h) stacks → column width grows to 15
    // Column 2: s3 (15w, 44h) alone (no room for s4)
    // Column 3: s4 (15w) → should NOT fit because 15 + 2 + 15 + 2 + 15 = 49 > 45
    // Bug: currentRowWidth only tracks completed columns, but check didn't include currentColumnWidth
    const s1 = mockSection(10, 40)
    const s2 = mockSection(15, 4)   // stacks, grows column 1 to 15w, total height 44
    const s3 = mockSection(15, 44)  // starts column 2 (can't stack: 44+4=48>45)
    const s4 = mockSection(15, 44)  // should NOT fit as column 3 (15+2+15+2+15=49>45)
    const result = calculateLayout([s1, s2, s3, s4], 45, 50, 5, 2, 0)
    // Expected: 2 pages (s4 doesn't fit on page 1)
    // Bug: 1 page with 3 columns (if currentColumnWidth not included in check)
    expect(result).toHaveLength(2)
    expect(result[0]).toHaveLength(2)  // columns 1 and 2
    expect(result[0][0]).toEqual([s1, s2])
    expect(result[0][1]).toEqual([s3])
    expect(result[1][0]).toEqual([s4])
  })

  it("should not allow stacking when it would cause column width to overflow page", () => {
    // Page width: 46, already have columns totaling 31w (14 + 2 + 15)
    // Column 3 starts with 13w section, then tries to stack 15w section
    // 31 + 2 + 15 = 48 > 46, so the 15w section should NOT stack
    const s1 = mockSection(14, 44)  // column 1
    const s2 = mockSection(15, 30)  // column 2: doesn't stack (44+30 > 46)
    const s3 = mockSection(13, 16)  // column 3: fits (14+2+15+2+13=46)
    const s4 = mockSection(15, 8)   // would grow column 3 to 15w: 14+2+15+2+15=48 > 46
    const result = calculateLayout([s1, s2, s3, s4], 46, 51, 5, 2, 2)
    // s4 should NOT stack into column 3, should go to new page
    expect(result).toHaveLength(2)
    expect(result[0]).toHaveLength(3)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2])
    expect(result[0][2]).toEqual([s3])
    expect(result[1][0]).toEqual([s4])
  })
})
