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

const { detectChanges, applyStructuralColumnChanges: applyStructuralColumnChanges, applyStructuralRowChanges: applyStructuralRowChanges } = await import("../src/MusicSheet")

describe("detectChanges", () => {
  describe("no changes", () => {
    it("should return empty array for sequential indexes", () => {
      expect(detectChanges([1, 2, 3, 4, 5], 1, 5)).toEqual([])
    })

    it("should return empty array for empty input", () => {
      expect(detectChanges([], 1, 0)).toEqual([])
    })

    it("should return empty array for single element", () => {
      expect(detectChanges([1], 1, 1)).toEqual([])
    })
  })

  describe("insertions", () => {
    it("should detect single insertion at the start", () => {
      const changes = detectChanges([null, 1, 2, 3], 1, 4)
      expect(changes).toEqual([
        { position: 1, span: 1 }
      ])
    })

    it("should detect single insertion in the middle", () => {
      const changes = detectChanges([1, 2, null, 3, 4], 1, 5)
      expect(changes).toEqual([
        { position: 3, span: 1 }
      ])
    })

    it("should detect single insertion at the end", () => {
      const changes = detectChanges([1, 2, 3, null], 1, 4)
      expect(changes).toEqual([
        { position: 4, span: 1 }
      ])
    })

    it("should detect multiple consecutive insertions as single change", () => {
      const changes = detectChanges([1, 2, null, null, null, 3, 4], 1, 7)
      expect(changes).toEqual([
        { position: 3, span: 3 }
      ])
    })

    it("should detect multiple separate insertions (ordered for right-to-left processing)", () => {
      const changes = detectChanges([1, null, 2, null, 3], 1, 5)
      expect(changes).toEqual([
        { position: 3, span: 1 },
        { position: 2, span: 1 }
      ])
    })

    it("should detect consecutive insertions at the start", () => {
      const changes = detectChanges([null, null, 1, 2], 1, 4)
      expect(changes).toEqual([
        { position: 1, span: 2 }
      ])
    })

    it("should detect insertion at first working area column (with frozen columns)", () => {
      // Full indexes: [1, 2, null, 3, 4, 5], frozen=2, working area columns 3-6
      const changes = detectChanges([1, 2, null, 3, 4, 5], 3, 6)
      expect(changes).toEqual([
        { position: 1, span: 1 }
      ])
    })

    it("should detect multiple insertions at first working area column (with frozen columns)", () => {
      // Full indexes: [1, 2, null, null, 5, 6], frozen=2, working area columns 3-6
      const changes = detectChanges([1, 2, null, null, 5, 6], 3, 6)
      expect(changes).toEqual([
        { position: 1, span: 2 }
      ])
    })
  })

  describe("deletions", () => {
    it("should detect single deletion at the start", () => {
      const changes = detectChanges([2, 3, 4], 1, 3)
      expect(changes).toEqual([
        { position: 1, span: -1 }
      ])
    })

    it("should detect single deletion in the middle", () => {
      const changes = detectChanges([1, 2, 4, 5], 1, 4)
      expect(changes).toEqual([
        { position: 3, span: -1 }
      ])
    })

    it("should detect multiple consecutive deletions as single change", () => {
      const changes = detectChanges([1, 2, 6, 7], 1, 7)
      expect(changes).toEqual([
        { position: 3, span: -3 }
      ])
    })

    it("should detect multiple separate deletions (ordered for right-to-left processing)", () => {
      const changes = detectChanges([1, 3, 5], 1, 5)
      expect(changes).toEqual([
        { position: 4, span: -1 },
        { position: 2, span: -1 }
      ])
    })

    it("should detect deletion of all but first element", () => {
      const changes = detectChanges([1, 10], 1, 10)
      expect(changes).toEqual([
        { position: 2, span: -8 }
      ])
    })
  })

  describe("filtering by working area", () => {
    it("should ignore insertions in frozen area", () => {
      // Insert in frozen column 2, working area starts at 3
      const changes = detectChanges([1, null, 2, 3, 4], 3, 5)
      expect(changes).toEqual([])
    })

    it("should ignore deletions in frozen area", () => {
      // Delete frozen column 2, working area starts at 3
      const changes = detectChanges([1, 3, 4, 5], 3, 5)
      expect(changes).toEqual([])
    })

    it("should ignore insertions in tray area", () => {
      // Insert at position 6 which is in tray (working area ends at 4)
      const changes = detectChanges([1, 2, 3, 4, null, 5], 1, 4)
      expect(changes).toEqual([])
    })

    it("should ignore deletions in tray area", () => {
      // Delete at position 5 which is in tray (working area ends at 4)
      const changes = detectChanges([1, 2, 3, 4, 6], 1, 4)
      expect(changes).toEqual([])
    })
  })

  describe("edge cases", () => {
    it("should handle all nulls (all new columns)", () => {
      const changes = detectChanges([null, null, null], 1, 3)
      expect(changes).toEqual([
        { position: 1, span: 3 }
      ])
    })

    it("should handle single null", () => {
      const changes = detectChanges([null], 1, 1)
      expect(changes).toEqual([
        { position: 1, span: 1 }
      ])
    })
  })
})

describe("applyStructuralColumnChanges", () => {
  const sampleValues = [
    ["A", "B", "C", "D"],
    ["1", "2", "3", "4"]
  ]

  describe("insertions", () => {
    it("should insert column at the start", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 1, span: 1 }
      ])
      expect(result).toEqual([
        ["", "A", "B", "C", "D"],
        ["", "1", "2", "3", "4"]
      ])
    })

    it("should insert column in the middle", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 3, span: 1 }
      ])
      expect(result).toEqual([
        ["A", "B", "", "C", "D"],
        ["1", "2", "", "3", "4"]
      ])
    })

    it("should insert column at the end", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 5, span: 1 }
      ])
      expect(result).toEqual([
        ["A", "B", "C", "D", ""],
        ["1", "2", "3", "4", ""]
      ])
    })

    it("should insert multiple consecutive columns", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 2, span: 3 }
      ])
      expect(result).toEqual([
        ["A", "", "", "", "B", "C", "D"],
        ["1", "", "", "", "2", "3", "4"]
      ])
    })
  })

  describe("deletions", () => {
    it("should delete column at the start", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 1, span: -1 }
      ])
      expect(result).toEqual([
        ["B", "C", "D"],
        ["2", "3", "4"]
      ])
    })

    it("should delete column in the middle", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 2, span: -1 }
      ])
      expect(result).toEqual([
        ["A", "C", "D"],
        ["1", "3", "4"]
      ])
    })

    it("should delete column at the end", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 4, span: -1 }
      ])
      expect(result).toEqual([
        ["A", "B", "C"],
        ["1", "2", "3"]
      ])
    })

    it("should delete multiple consecutive columns", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 2, span: -2 }
      ])
      expect(result).toEqual([
        ["A", "D"],
        ["1", "4"]
      ])
    })
  })

  describe("multiple changes", () => {
    it("should apply multiple changes right-to-left", () => {
      const result = applyStructuralColumnChanges(sampleValues, [
        { position: 4, span: 1 },
        { position: 2, span: -1 }
      ])
      expect(result).toEqual([
        ["A", "C", "", "D"],
        ["1", "3", "", "4"]
      ])
    })
  })

  describe("edge cases", () => {
    it("should return unchanged copy for empty changes", () => {
      const result = applyStructuralColumnChanges(sampleValues, [])
      expect(result).toEqual(sampleValues)
      expect(result).not.toBe(sampleValues)
    })

    it("should handle empty values array", () => {
      const result = applyStructuralColumnChanges([], [
        { position: 1, span: 1 }
      ])
      expect(result).toEqual([])
    })

    it("should not mutate original values", () => {
      const original = [["A", "B"], ["1", "2"]]
      applyStructuralColumnChanges(original, [
        { position: 1, span: -1 }
      ])
      expect(original).toEqual([["A", "B"], ["1", "2"]])
    })
  })
})

describe("applyStructuralRowChanges", () => {
  const sampleValues = [
    ["A", "B", "C"],
    ["1", "2", "3"],
    ["D", "E", "F"],
    ["4", "5", "6"]
  ]

  describe("insertions (with 2x scaling)", () => {
    it("should insert 2 rows at the start for 1 lyrics row", () => {
      const result = applyStructuralRowChanges(sampleValues, [
        { position: 1, span: 1 }
      ])
      expect(result).toEqual([
        ["", "", ""],
        ["", "", ""],
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["D", "E", "F"],
        ["4", "5", "6"]
      ])
    })

    it("should insert 2 rows in the middle for 1 lyrics row", () => {
      const result = applyStructuralRowChanges(sampleValues, [
        { position: 2, span: 1 }
      ])
      expect(result).toEqual([
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["", "", ""],
        ["", "", ""],
        ["D", "E", "F"],
        ["4", "5", "6"]
      ])
    })

    it("should insert 4 rows for 2 lyrics rows", () => {
      const result = applyStructuralRowChanges(sampleValues, [
        { position: 2, span: 2 }
      ])
      expect(result).toEqual([
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["D", "E", "F"],
        ["4", "5", "6"]
      ])
    })

    it("should insert 2 rows at the end for 1 lyrics row", () => {
      const result = applyStructuralRowChanges(sampleValues, [
        { position: 3, span: 1 }
      ])
      expect(result).toEqual([
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["D", "E", "F"],
        ["4", "5", "6"],
        ["", "", ""],
        ["", "", ""]
      ])
    })
  })

  describe("deletions (with 2x scaling)", () => {
    it("should delete 2 rows at the start for 1 lyrics row", () => {
      const result = applyStructuralRowChanges(sampleValues, [
        { position: 1, span: -1 }
      ])
      expect(result).toEqual([
        ["D", "E", "F"],
        ["4", "5", "6"]
      ])
    })

    it("should delete 2 rows in the middle for 1 lyrics row", () => {
      const sixRowValues = [
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["D", "E", "F"],
        ["4", "5", "6"],
        ["G", "H", "I"],
        ["7", "8", "9"]
      ]
      const result = applyStructuralRowChanges(sixRowValues, [
        { position: 2, span: -1 }
      ])
      expect(result).toEqual([
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["G", "H", "I"],
        ["7", "8", "9"]
      ])
    })

    it("should delete 4 rows for 2 lyrics rows", () => {
      const sixRowValues = [
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["D", "E", "F"],
        ["4", "5", "6"],
        ["G", "H", "I"],
        ["7", "8", "9"]
      ]
      const result = applyStructuralRowChanges(sixRowValues, [
        { position: 1, span: -2 }
      ])
      expect(result).toEqual([
        ["G", "H", "I"],
        ["7", "8", "9"]
      ])
    })
  })

  describe("multiple changes", () => {
    it("should apply multiple changes right-to-left with 2x scaling", () => {
      const sixRowValues = [
        ["A", "B", "C"],
        ["1", "2", "3"],
        ["D", "E", "F"],
        ["4", "5", "6"],
        ["G", "H", "I"],
        ["7", "8", "9"]
      ]
      const result = applyStructuralRowChanges(sixRowValues, [
        { position: 3, span: 1 },
        { position: 1, span: -1 }
      ])
      expect(result).toEqual([
        ["D", "E", "F"],
        ["4", "5", "6"],
        ["", "", ""],
        ["", "", ""],
        ["G", "H", "I"],
        ["7", "8", "9"]
      ])
    })
  })

  describe("edge cases", () => {
    it("should return unchanged copy for empty changes", () => {
      const result = applyStructuralRowChanges(sampleValues, [])
      expect(result).toEqual(sampleValues)
      expect(result).not.toBe(sampleValues)
    })

    it("should handle empty values array", () => {
      const result = applyStructuralRowChanges([], [
        { position: 1, span: 1 }
      ])
      expect(result).toEqual([])
    })

    it("should not mutate original values", () => {
      const original = [["A", "B"], ["1", "2"]]
      applyStructuralRowChanges(original, [
        { position: 1, span: -1 }
      ])
      expect(original).toEqual([["A", "B"], ["1", "2"]])
    })
  })
})
