import { describe, it, expect, beforeEach, vi } from "vitest"
import { Area, Origin } from "../src/Area"

type CellValue = string | number | boolean | Date | null

type SheetConfig = {
  frozen: [number, number]
  tray?: number
}

type SheetData = {
  values: CellValue[][]
  weights: string[][]
}

class MockSubSpace {
  constructor(
    private data: SheetData,
    readonly area: Area,
  ) { }

  get x() { return this.area.x }
  get y() { return this.area.y }
  get width() { return this.area.width }
  get height() { return this.area.height }
  get start() { return this.area.start }
  get isEmpty() { return this.area.isEmpty }

  sub(f: (area: Area) => Area): MockSubSpace {
    return new MockSubSpace(this.data, f(this.area))
  }

  getValues(): CellValue[][] {
    if (this.area.isEmpty) return []
    return this.data.values
      .slice(this.y, this.y + this.height)
      .map(row => row.slice(this.x, this.x + this.width))
  }

  setValues(values: CellValue[][]): void {
    if (this.area.isEmpty) return
    values.forEach((row, i) => {
      row.forEach((cell, j) => {
        this.data.values[this.y + i][this.x + j] = cell
      })
    })
  }

  getFontWeights(): string[][] {
    if (this.area.isEmpty) return []
    return this.data.weights
      .slice(this.y, this.y + this.height)
      .map(row => row.slice(this.x, this.x + this.width))
  }

  setFontWeights(weights: string[][]): void {
    if (this.area.isEmpty) return
    weights.forEach((row, i) => {
      row.forEach((weight, j) => {
        this.data.weights[this.y + i][this.x + j] = weight
      })
    })
  }
}

class MockSheetSpace {
  private data: SheetData
  readonly area: Area
  private frozenRowCount: number
  private frozenColCount: number
  private trayWidth: number

  constructor(values: CellValue[][], config: SheetConfig) {
    this.data = {
      values: values.map(row => [...row]),
      weights: values.map(row => row.map(() => "normal")),
    }
    this.frozenRowCount = config.frozen[0]
    this.frozenColCount = config.frozen[1]
    this.trayWidth = config.tray ?? 0
    this.area = Origin.by({ x: values[0]?.length ?? 0, y: values.length })
  }

  get main(): MockSubSpace {
    return this.sub(area => area.crop({
      left: this.frozenColCount,
      top: this.frozenRowCount,
      right: this.trayWidth,
    }))
  }

  get frozenRows(): MockSubSpace {
    return this.sub(area => area.rows(this.frozenRowCount))
  }

  get frozenColumns(): MockSubSpace {
    return this.sub(area => area.columns(this.frozenColCount))
  }

  sub(f: (area: Area) => Area): MockSubSpace {
    return new MockSubSpace(this.data, f(this.area))
  }

  getData(): CellValue[][] {
    return this.data.values.map(row => [...row])
  }

  getWeights(): string[][] {
    return this.data.weights.map(row => [...row])
  }
}

function sheet(data: CellValue[][], config: SheetConfig): MockSheetSpace {
  return new MockSheetSpace(data, config)
}

let mockLyrics: MockSheetSpace
let mockChords: MockSheetSpace

vi.mock("../src/Range", () => ({}))

vi.mock("../src/Spaces", () => ({
  $: {
    get Lyrics() { return mockLyrics },
    get Chords() { return mockChords },
  },
  LYRICS_SHEET_NAME: "Letra",
  CHORDS_SHEET_NAME: "Acordes",
  PRINT_SHEET_NAME: "Impresión",
  LYRICS_RIGHT_TRAY_RANGE_NAME: "Ideas_Sueltas",
}))

import { syncLyricsToChordSheet, syncLyricsFromChordSheet } from "../src/MusicSheet"

describe("syncLyricsToChordSheet", () => {
  beforeEach(() => {
    mockLyrics = sheet([
      ["1", "2", "3", "4", "5", "6", "7", "8"],
      ["2", " ", " ", " ", " ", " ", " ", " "],
      ["3", " ", " ", " ", " ", " ", " ", " "],
      ["4", " ", " ", "A", "B", "C", "t", "x"],
      ["5", " ", " ", "D", "E", "F", "u", "y"],
    ], { frozen: [3, 3], tray: 2 })

    mockChords = sheet([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", " ", " ", " "],
      ["4", " ", "a", "b", "c"],
      ["5", " ", " ", " ", " "],
      ["6", " ", "d", "e", "f"],
    ], { frozen: [2, 2] })
  })

  it("syncs edit range from lyrics main to chords lyric rows", () => {
    syncLyricsToChordSheet(new Area(3, 3, 3, 2))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", " ", " ", " "],
      ["4", " ", "A", "B", "C"],
      ["5", " ", " ", " ", " "],
      ["6", " ", "D", "E", "F"],
    ])
  })

  it("does not affect cells outside the target range", () => {
    syncLyricsToChordSheet(new Area(3, 3, 2, 1))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", " ", " ", " "],
      ["4", " ", "A", "B", "c"],
      ["5", " ", " ", " ", " "],
      ["6", " ", "d", "e", "f"],
    ])
  })

  it("preserves chord rows (even indices in main)", () => {
    mockChords = sheet([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "X", "Y", "Z"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "P", "Q", "R"],
    ], { frozen: [2, 2] })

    syncLyricsToChordSheet(new Area(3, 3, 3, 2))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "D", "E", "F"],
    ])
  })

  it("produces no change when edit is outside lyrics main (frozen area)", () => {
    const originalChords = mockChords.getData()

    syncLyricsToChordSheet(new Area(0, 0, 3, 3))

    expect(mockChords.getData()).toEqual(originalChords)
  })

  it("produces no change when edit is outside lyrics main (tray area)", () => {
    const originalChords = mockChords.getData()

    syncLyricsToChordSheet(new Area(6, 3, 2, 2))

    expect(mockChords.getData()).toEqual(originalChords)
  })

  it("syncs edit that projects outside chords main (chords resizes)", () => {
    mockLyrics = sheet([
      ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
      ["2", " ", " ", " ", " ", " ", " ", " ", " ", " "],
      ["3", " ", " ", " ", " ", " ", " ", " ", " ", " "],
      ["4", " ", " ", "A", "B", "C", "G", "H", "t", "x"],
      ["5", " ", " ", "D", "E", "F", "I", "J", "u", "y"],
    ], { frozen: [3, 3], tray: 2 })

    syncLyricsToChordSheet(new Area(6, 3, 2, 2))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", " ", " ", " "],
      ["4", " ", "a", "b", "c", "G", "H"],
      ["5", " ", " ", " ", " "],
      ["6", " ", "d", "e", "f", "I", "J"],
    ])
  })

  it("syncs only the valid intersection when edit partially overlaps main", () => {
    syncLyricsToChordSheet(new Area(2, 2, 3, 3))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", " ", " ", " "],
      ["4", " ", "A", "B", "c"],
      ["5", " ", " ", " ", " "],
      ["6", " ", "D", "E", "f"],
    ])
  })
})

describe("syncLyricsFromChordSheet", () => {
  beforeEach(() => {
    mockLyrics = sheet([
      ["1", "2", "3", "4", "5", "6", "7", "8"],
      ["2", " ", " ", " ", " ", " ", " ", " "],
      ["3", " ", " ", " ", " ", " ", " ", " "],
      ["4", " ", " ", "A", "B", "C", "t", "x"],
      ["5", " ", " ", "D", "E", "F", "u", "y"],
    ], { frozen: [3, 3], tray: 2 })

    mockChords = sheet([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "a", "b", "c"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "d", "e", "f"],
    ], { frozen: [2, 2] })
  })

  it("height=0 produces no change", () => {
    const originalChords = mockChords.getData()

    syncLyricsFromChordSheet(new Area(2, 2, 3, 0))

    expect(mockChords.getData()).toEqual(originalChords)
  })

  it("height=1 at even row (chord row) syncs first lyrics row", () => {
    syncLyricsFromChordSheet(new Area(2, 2, 3, 1))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "d", "e", "f"],
    ])
  })

  it("height=1 at odd row (lyric row) syncs first lyrics row", () => {
    syncLyricsFromChordSheet(new Area(2, 3, 3, 1))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "d", "e", "f"],
    ])
  })

  it("height=2 at even row syncs first lyrics row", () => {
    syncLyricsFromChordSheet(new Area(2, 2, 3, 2))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "d", "e", "f"],
    ])
  })

  it("height=2 at odd row syncs both lyrics rows", () => {
    syncLyricsFromChordSheet(new Area(2, 3, 3, 2))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "D", "E", "F"],
    ])
  })

  it("height=3 at even row syncs both lyrics rows", () => {
    syncLyricsFromChordSheet(new Area(2, 2, 3, 3))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "D", "E", "F"],
    ])
  })

  it("height=3 at odd row syncs both lyrics rows", () => {
    syncLyricsFromChordSheet(new Area(2, 3, 3, 3))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "D", "E", "F"],
    ])
  })

  it("height=4 at even row syncs both lyrics rows", () => {
    syncLyricsFromChordSheet(new Area(2, 2, 3, 4))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "C"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "D", "E", "F"],
    ])
  })

  it("produces no change when edit is outside chords main (frozen area)", () => {
    const originalChords = mockChords.getData()

    syncLyricsFromChordSheet(new Area(0, 0, 2, 2))

    expect(mockChords.getData()).toEqual(originalChords)
  })

  it("syncs only the valid intersection when edit partially overlaps main", () => {
    syncLyricsFromChordSheet(new Area(1, 1, 3, 3))

    expect(mockChords.getData()).toEqual([
      ["1", "2", "3", "4", "5"],
      ["2", " ", " ", " ", " "],
      ["3", " ", "x", "y", "z"],
      ["4", " ", "A", "B", "c"],
      ["5", " ", "p", "q", "r"],
      ["6", " ", "d", "e", "f"],
    ])
  })
})
