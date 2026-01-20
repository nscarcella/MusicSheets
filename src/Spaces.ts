import { lazy, raise } from "./Utils"
import { Area, Empty, Origin, Point } from "./Area"

type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
export type CellValue = string | number | boolean | Date | null


const LYRICS_SHEET_NAME = "Letra"
const CHORDS_SHEET_NAME = "Acordes"
const PRINT_SHEET_NAME = "Impresión"

const LYRICS_RIGHT_TRAY_RANGE_NAME = "Ideas_Sueltas"
const CHORDS_HEADER_RANGE_NAME = "Encabezado_Acordes"
const KEY_RANGE_NAME = "Tonalidad"
const AUTOTRANSPOSE_RANGE_NAME = "Auto_Trasponer"


export abstract class Space {
  abstract area: Area
  abstract sheet: Sheet

  @lazy get range(): Range | undefined {
    return this.area.isEmpty ? undefined : this.sheet.getRange(this.y + 1, this.x + 1, this.height, this.width)
  }
  @lazy get x(): number { return this.area.x }
  @lazy get y(): number { return this.area.y }
  @lazy get width(): number { return this.area.width }
  @lazy get height(): number { return this.area.height }
  @lazy get start(): Point { return this.area.start }
  @lazy get end(): Point { return this.area.end }
  @lazy get isEmpty(): boolean { return this.area.isEmpty }

  getValues(): CellValue[][] { return this.range?.getValues() ?? [] }
  setValues(newValues: CellValue[][]): void { this.range?.setValues(newValues) }

  sub(f: (area: Area) => Area): SubSpace { return new SubSpace(this, f(this.area)) }
  cell<T extends CellValue>(adapter: (value: CellValue) => T) {
    return (f: (area: Area) => Area) => { return new CellSpace(this, adapter, f(this.area)) }
  }
}


export class SheetSpace extends Space {
  readonly sheet: Sheet

  constructor(name: string) {
    super()
    this.sheet = $.SPREADSHEET.getSheetByName(name) ?? raise(new Error(`Sheet "${name}" not found`))
  }

  @lazy get name(): string { return this.sheet.getName() }
  @lazy get area(): Area { return Origin.by({ x: this.sheet.getMaxColumns(), y: this.sheet.getMaxRows() }) }

  @lazy get main(): SubSpace { return this.sub(area => area.crop({ left: this.frozenColumns.width, top: this.frozenRows.height })) }
  @lazy get frozenRows(): SubSpace { return this.sub(area => area.rows(this.sheet.getFrozenRows())) }
  @lazy get frozenColumns(): SubSpace { return this.sub(area => area.columns(this.sheet.getFrozenColumns())) }

  getNamedArea(name: string, defaultArea?: Area): Area {
    const range = $.SPREADSHEET.getRangeByName(name)
    return range ? Area.fromRange(range) : defaultArea ?? raise(new Error(`Named range "${name}" not found`))
  }

  getLastRowWithContent(): number { return this.sheet.getLastRow() - this.frozenRows.height }
}


export class SubSpace extends Space {
  constructor(
    readonly parent: Space,
    readonly area: Area,
  ) {
    super()
    if (area.x < 0 || area.y < 0) {
      throw new Error(`Invalid area for Range: ${area}`)
    }
  }

  @lazy get sheet(): Sheet { return this.parent.sheet }
}


export class CellSpace<T extends CellValue> extends Space {
  constructor(
    readonly parent: Space,
    readonly adapter: (value: CellValue) => T,
    readonly area: Area,
  ) {
    super()
    if (area.width !== 1 || area.height !== 1) {
      throw new Error(`CellSpace must be 1x1, got ${area}`)
    }
  }

  @lazy get sheet(): Sheet { return this.parent.sheet }

  getValue(): T { return this.adapter(this.range?.getValue() ?? null) }
  setValue(value: T): void { this.range?.setValue(value) }
}



export class $ {
  @lazy static get SPREADSHEET() { return SpreadsheetApp.getActive() }
  @lazy static get ALL(): SheetSpace[] { return Object.values($).filter((v): v is SheetSpace => v instanceof SheetSpace) }

  static get(name: string): SheetSpace {
    return $.ALL.find(space => space.name === name) ?? raise(new Error(`Unknown sheet: "${name}"`))
  }

  @lazy static get Lyrics() { return new LyricsSheet() }
  @lazy static get Chords() { return new ChordsSheet() }
  @lazy static get Print() { return new SheetSpace(PRINT_SHEET_NAME) }
}

class LyricsSheet extends SheetSpace {
  constructor() { super(LYRICS_SHEET_NAME) }

  @lazy override get main(): SubSpace { return super.main.sub(area => area.crop({ right: this.sideTray.width })) }
  @lazy get indexColumn(): SubSpace { return this.sub(area => area.columns(1)) }
  @lazy get indexRow(): SubSpace { return this.sub(area => area.rows(1)) }
  @lazy get sideTray(): SubSpace {
    return this.sub(() => super.main.area.columns(-this.getNamedArea(LYRICS_RIGHT_TRAY_RANGE_NAME, Empty).width))
  }
}

class ChordsSheet extends SheetSpace {
  constructor() { super(CHORDS_SHEET_NAME) }

  @lazy get header(): SubSpace { return this.sub(() => this.getNamedArea(CHORDS_HEADER_RANGE_NAME)) }
  @lazy get key(): CellSpace<string> { return this.header.cell(String)(() => this.getNamedArea(KEY_RANGE_NAME)) }
  @lazy get autotranspose(): CellSpace<boolean> { return this.header.cell(Boolean)(() => this.getNamedArea(AUTOTRANSPOSE_RANGE_NAME)) }
}