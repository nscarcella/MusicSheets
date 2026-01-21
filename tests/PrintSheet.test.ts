import { describe, it, expect, vi } from "vitest"
import { Area, Point } from "../src/Area"
import type { LayoutConfig } from "../src/MusicSheet"

const mockSheetProto = {}
const mockRangeProto = {}
const mockSheet = Object.create(mockSheetProto)
const mockRange = Object.create(mockRangeProto)
mockSheet.getRange = () => mockRange

vi.stubGlobal("SpreadsheetApp", {
  getActiveSpreadsheet: () => ({ getSheets: () => [mockSheet] }),
  getActive: () => ({ getSheetByName: () => mockSheet })
})

const { splitIntoSectionColumns, splitIntoSections, calculateSectionWidth, calculatePositions } = await import("../src/MusicSheet")

function mockSection(width: number, height: number): Area {
  return new Area(0, 0, width, height)
}

function testConfig(
  availableWidth: number,
  availableHeight: number,
  firstPageHeaderHeight: number,
  horizontalPadding = 0,
  verticalPadding = 0
): LayoutConfig {
  return {
    fullPageWidth: availableWidth,
    availableWidthPerPage: availableWidth,
    availableHeightPerPage: availableHeight,
    firstPageHeaderHeight,
    horizontalContentMargin: 0,
    verticalContentMargin: 0,
    horizontalPadding,
    verticalPadding,
  }
}

function deriveStructure(sections: Area[], positions: Point[], pageWidth: number): Area[][][] {
  const pages = new Map<number, Map<number, Area[]>>()

  sections.forEach((section, i) => {
    const pos = positions[i]
    const pageIndex = Math.floor(pos.x / pageWidth)
    const columnX = pos.x

    if (!pages.has(pageIndex)) pages.set(pageIndex, new Map())
    const page = pages.get(pageIndex)!
    if (!page.has(columnX)) page.set(columnX, [])
    page.get(columnX)!.push(section)
  })

  return Array.from(pages.entries())
    .sort(([a], [b]) => a - b)
    .map(([, page]) =>
      Array.from(page.entries())
        .sort(([a], [b]) => a - b)
        .map(([, sections]) => sections)
    )
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

describe("calculatePositions", () => {
  it("should return empty array for no sections", () => {
    expect(calculatePositions([], testConfig(45, 50, 5))).toEqual([])
  })

  it("should place single section on first page", () => {
    const section = mockSection(10, 20)
    const sections = [section]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(1)
    expect(result[0][0]).toEqual([section])
  })

  it("should stack sections vertically in same column", () => {
    const s1 = mockSection(10, 10)
    const s2 = mockSection(10, 10)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(1)
    expect(result[0][0]).toEqual([s1, s2])
  })

  it("should start new column when height exceeded", () => {
    const s1 = mockSection(10, 30)
    const s2 = mockSection(10, 20)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2])
  })

  it("should start new page when width exceeded", () => {
    const s1 = mockSection(25, 30)
    const s2 = mockSection(25, 30)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2])
  })

  it("should respect first page header height", () => {
    const s1 = mockSection(10, 44)
    const s2 = mockSection(10, 2)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
  })

  it("should place sections on second page when first is full", () => {
    const s1 = mockSection(45, 45)
    const s2 = mockSection(10, 45)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(2)
    expect(result[1][0]).toEqual([s2])
  })

  it("should have more space on second page (no header)", () => {
    const s1 = mockSection(45, 45)
    const s2 = mockSection(10, 48)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2])
  })

  it("should stack more sections on second page due to extra height", () => {
    const s1 = mockSection(45, 45)
    const s2 = mockSection(10, 24)
    const s3 = mockSection(10, 24)
    const sections = [s1, s2, s3]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2, s3])
  })

  it("should fill columns before starting new page", () => {
    const s1 = mockSection(15, 30)
    const s2 = mockSection(15, 30)
    const s3 = mockSection(15, 30)
    const sections = [s1, s2, s3]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(3)
  })

  it("should handle complex layout", () => {
    const s1 = mockSection(20, 25)
    const s2 = mockSection(20, 25)
    const s3 = mockSection(20, 20)
    const s4 = mockSection(20, 20)
    const sections = [s1, s2, s3, s4]
    const positions = calculatePositions(sections, testConfig(45, 50, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(2)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2, s3])
    expect(result[1][0]).toEqual([s4])
  })

  it("should account for horizontal padding between columns", () => {
    const s1 = mockSection(20, 30)
    const s2 = mockSection(20, 30)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5, 6, 0))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[1][0]).toEqual([s2])
  })

  it("should account for vertical padding between sections", () => {
    const s1 = mockSection(10, 22)
    const s2 = mockSection(10, 22)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5, 0, 5))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2])
  })

  it("should fit sections without padding when padding is zero", () => {
    const s1 = mockSection(20, 30)
    const s2 = mockSection(20, 30)
    const sections = [s1, s2]
    const positions = calculatePositions(sections, testConfig(45, 50, 5, 0, 0))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(2)
  })

  it("should not add padding before first column or first section", () => {
    const s1 = mockSection(43, 43)
    const sections = [s1]
    const positions = calculatePositions(sections, testConfig(45, 50, 5, 2, 2))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(1)
    expect(result[0]).toHaveLength(1)
    expect(result[0][0]).toEqual([s1])
  })

  it("should account for column width growth when checking if new column fits", () => {
    const s1 = mockSection(10, 40)
    const s2 = mockSection(15, 4)
    const s3 = mockSection(15, 44)
    const s4 = mockSection(15, 44)
    const sections = [s1, s2, s3, s4]
    const positions = calculatePositions(sections, testConfig(45, 50, 5, 2, 0))
    const result = deriveStructure(sections, positions, 45)
    expect(result).toHaveLength(2)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1, s2])
    expect(result[0][1]).toEqual([s3])
    expect(result[1][0]).toEqual([s4])
  })

  it("should move entire column to next page when it doesn't fit width-wise", () => {
    const s1 = mockSection(14, 44)
    const s2 = mockSection(15, 30)
    const s3 = mockSection(13, 16)
    const s4 = mockSection(15, 8)
    const sections = [s1, s2, s3, s4]
    const positions = calculatePositions(sections, testConfig(46, 51, 5, 2, 2))
    const result = deriveStructure(sections, positions, 46)
    expect(result).toHaveLength(2)
    expect(result[0]).toHaveLength(2)
    expect(result[0][0]).toEqual([s1])
    expect(result[0][1]).toEqual([s2])
    expect(result[1][0]).toEqual([s3, s4])
  })
})
