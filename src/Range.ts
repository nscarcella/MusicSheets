type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet


declare global {
  namespace GoogleAppsScript.Spreadsheet {
    interface Sheet {
      fullRange(): Range
    }

    interface Range {
      translate(x?: number, y?: number): Range
      translateTo(x?: number, y?: number): Range
      scale(x?: number, y?: number): Range
      resize(x?: number, y?: number): Range
      resizeTo(x?: number, y?: number): Range
      projectInto(targetSheet: Sheet): Range
      overlapsWith(other: Range): boolean
      intersect(other: Range): Range | undefined
      print(): string
    }
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// SHEET
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

(() => {
  const SheetPrototype = Object.getPrototypeOf(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0])

  SheetPrototype.fullRange = function (this: Sheet): Range {
    return this.getRange(1, 1, this.getMaxRows(), this.getMaxColumns())
  }
})();

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// RANGE
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

(() => {
  const RangePrototype = Object.getPrototypeOf(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange("A1"))

  RangePrototype.translate = function (this: Range, x: number = 0, y: number = 0): Range {
    return this.translateTo(this.getColumn() + x, this.getRow() + y)
  }

  RangePrototype.translateTo = function (this: Range, x: number = this.getColumn(), y: number = this.getRow()): Range {
    if (x < 1 || y < 1) throw new Error(`Translation out of bounds: column ${x}, row ${y}`)

    return this.getSheet().getRange(y, x, this.getNumRows(), this.getNumColumns())
  }

  RangePrototype.scale = function (this: Range, x: number = 1, y: number = 1): Range {
    if (x <= 0 || y <= 0) throw new Error(`Invalid scale multiplier: x=${x}, y=${y}`)

    return this.getSheet().getRange(this.getRow(), this.getColumn(), Math.ceil(this.getNumRows() * y), Math.ceil(this.getNumColumns() * x))
  }

  RangePrototype.resize = function (this: Range, x: number = 0, y: number = 0): Range {
    return this.resizeTo(this.getNumColumns() + x, this.getNumRows() + y)
  }

  RangePrototype.resizeTo = function (this: Range, x: number = this.getNumColumns(), y: number = this.getNumRows()): Range {
    if (x < 1 || y < 1) throw new Error(`Resize to non-positive dimensions: ${x} columns, ${y} rows`)

    return this.getSheet().getRange(this.getRow(), this.getColumn(), y, x)
  }


  RangePrototype.projectInto = function (this: Range, targetSheet: Sheet): Range {
    return targetSheet.getRange(this.getRow(), this.getColumn(), this.getNumRows(), this.getNumColumns())
  }

  RangePrototype.overlapsWith = function (this: Range, other: Range): boolean {
    return this.intersect(other) !== undefined
  }

  RangePrototype.intersect = function (this: Range, other: Range): Range | undefined {
    const startRow = Math.max(this.getRow(), other.getRow())
    const startCol = Math.max(this.getColumn(), other.getColumn())
    const endRow = Math.min(this.getLastRow(), other.getLastRow())
    const endCol = Math.min(this.getLastColumn(), other.getLastColumn())

    if (startRow > endRow || startCol > endCol) return undefined

    return this.getSheet().getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1)
  }

  RangePrototype.print = function (this: Range): string {
    return `${this.getSheet().getName()}!${this.getA1Notation()} [${this.getNumRows()}×${this.getNumColumns()}]`
  }
})()

export { }