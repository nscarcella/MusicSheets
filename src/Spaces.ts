import { lazy, raise } from "./Utils"
import { Area, Empty, Origin, Point } from "./Area"

type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type RichTextValue = GoogleAppsScript.Spreadsheet.RichTextValue
type FontWeight = GoogleAppsScript.Spreadsheet.FontWeight
export type CellValue = string | number | boolean | Date | null


const LYRICS_SHEET_NAME = "Letra"
const CHORDS_SHEET_NAME = "Acordes"
const PRINT_SHEET_NAME = "Impresión"
const CONTROL_PANEL_SHEET_NAME = "Panel de Control"

const LYRICS_RIGHT_TRAY_RANGE_NAME = "Ideas_Sueltas"
const CHORDS_HEADER_RANGE_NAME = "Encabezado_Acordes"
const KEY_RANGE_NAME = "Tonalidad"
const AUTOTRANSPOSE_RANGE_NAME = "Auto_Trasponer"
const TEMPO_RANGE_NAME = "Tempo"
const NOTES_RANGE_NAME = "Notas"
const AUTHOR_RANGE_NAME = "Autor"
const CONTENT_MARGIN_H_RANGE_NAME = "Margen_H"
const CONTENT_MARGIN_V_RANGE_NAME = "Margen_V"
const HORIZONTAL_PADDING_RANGE_NAME = "Separación_H"
const VERTICAL_PADDING_RANGE_NAME = "Separación_V"

const DEFAULT_CONTENT_MARGIN_H = 1
const DEFAULT_CONTENT_MARGIN_V = 1
const DEFAULT_HORIZONTAL_PADDING = 2
const DEFAULT_VERTICAL_PADDING = 2

export const PRINT_PAGE_WIDTH = 46
export const PRINT_PAGE_HEIGHT = 51
export const PRINT_HEADER_HEIGHT = 4
export const PRINT_FOOTER_HEIGHT = 1
const PRINT_HEADER_MARGIN = 1


export interface FormatOptions {
  fontFamily?: string
  fontSize?: number
  fontWeight?: "bold" | "normal"
  verticalAlignment?: "top" | "middle" | "bottom"
  horizontalAlignment?: "left" | "center" | "right"
  wrapStrategy?: GoogleAppsScript.Spreadsheet.WrapStrategy
  border?: boolean
}

export abstract class Space {
  abstract area: Area
  abstract sheet: Sheet

  @lazy protected get range(): Range {
    if (this.area.isEmpty) throw new Error("Cannot access range of empty space")
    return this.sheet.getRange(this.y + 1, this.x + 1, this.height, this.width)
  }

  get x(): number { return this.area.x }
  get y(): number { return this.area.y }
  get width(): number { return this.area.width }
  get height(): number { return this.area.height }
  get start(): Point { return this.area.start }
  get end(): Point { return this.area.end }
  get isEmpty(): boolean { return this.area.isEmpty }

  getValues(): CellValue[][] { return this.range.getValues() }
  setValues(newValues: CellValue[][]): void { this.range.setValues(newValues) }

  getRichTextValues(): (RichTextValue | null)[][] { return this.range.getRichTextValues() }
  setRichTextValues(values: RichTextValue[][]): void { this.range.setRichTextValues(values) }

  getFontWeights(): (FontWeight | null)[][] { return this.range.getFontWeights() }
  setFontWeights(weights: (FontWeight | null)[][]): void { this.range.setFontWeights(weights) }

  format(options: FormatOptions): this {
    if (options.fontFamily !== undefined) this.range.setFontFamily(options.fontFamily)
    if (options.fontSize !== undefined) this.range.setFontSize(options.fontSize)
    if (options.fontWeight !== undefined) this.range.setFontWeight(options.fontWeight)
    if (options.verticalAlignment !== undefined) this.range.setVerticalAlignment(options.verticalAlignment)
    if (options.horizontalAlignment !== undefined) this.range.setHorizontalAlignment(options.horizontalAlignment)
    if (options.wrapStrategy !== undefined) this.range.setWrapStrategy(options.wrapStrategy)
    if (options.border) this.range.setBorder(true, true, true, true, null, null)
    return this
  }

  sub(f: (area: Area) => Area): SubSpace { return new SubSpace(this, f(this.area)) }
  cell<T extends CellValue>(adapter: (value: CellValue) => T, defaultValue?: T) {
    return (f: (area: Area) => Area) => new CellSpace(this, adapter, f(this.area), defaultValue)
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
  private merged = false

  constructor(
    readonly parent: Space,
    readonly adapter: (value: CellValue) => T,
    readonly area: Area,
    readonly defaultValue?: T,
  ) {
    super()
    if (area.isEmpty) raise(new Error("Cannot create CellSpace with empty area"))
  }

  @lazy get sheet(): Sheet { return this.parent.sheet }

  private ensureMerged(): void {
    if (!this.merged && (this.area.width > 1 || this.area.height > 1)) {
      this.range?.merge()
      this.merged = true
    }
  }

  getValue(): T {
    this.ensureMerged()
    const value = this.adapter(this.range!.getValue())
    if (!value && this.defaultValue !== undefined) return this.defaultValue
    return value
  }

  setValue(value: T): void {
    this.ensureMerged()
    this.range!.setValue(value)
  }

  getRichTextValue(): RichTextValue | null { return this.range!.getRichTextValue() }
  setRichTextValue(value: RichTextValue): void { this.range!.setRichTextValue(value) }
}



export class $ {
  @lazy static get SPREADSHEET() { return SpreadsheetApp.getActive() }
  static get ALL(): SheetSpace[] { return [$.Lyrics, $.Chords, $.Print, $.ControlPanel] }

  static get(name: string): SheetSpace {
    return $.ALL.find(space => space.name === name) ?? raise(new Error(`Unknown sheet: "${name}"`))
  }

  @lazy static get Lyrics() { return new LyricsSheet() }
  @lazy static get Chords() { return new ChordsSheet() }
  @lazy static get Print() { return new PrintSheet() }
  @lazy static get ControlPanel() { return new ControlPanelSheet() }
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
  @lazy get tempo(): CellSpace<number> { return this.header.cell(Number)(() => this.getNamedArea(TEMPO_RANGE_NAME)) }
  @lazy get notes(): CellSpace<string> { return this.header.cell(String)(() => this.getNamedArea(NOTES_RANGE_NAME)) }
}

export class PrintSheet extends SheetSpace {
  constructor() { super(PRINT_SHEET_NAME) }

  private readonly headerWidth = PRINT_PAGE_WIDTH - 2 * PRINT_HEADER_MARGIN

  page(index: number): SubSpace {
    return this.sub(() => new Area(index * PRINT_PAGE_WIDTH, 0, PRINT_PAGE_WIDTH, PRINT_PAGE_HEIGHT))
  }

  pageContent(index: number): SubSpace {
    return this.sub(() => new Area(index * PRINT_PAGE_WIDTH, 0, PRINT_PAGE_WIDTH, PRINT_PAGE_HEIGHT - PRINT_FOOTER_HEIGHT))
  }

  pageFooter(index: number): CellSpace<string> {
    return this.cell(String)(() => new Area(index * PRINT_PAGE_WIDTH, PRINT_PAGE_HEIGHT - PRINT_FOOTER_HEIGHT, PRINT_PAGE_WIDTH, PRINT_FOOTER_HEIGHT))
  }

  @lazy get header(): SubSpace { return this.sub(() => Origin.to({ x: PRINT_PAGE_WIDTH, y: PRINT_HEADER_HEIGHT })) }
  @lazy get title(): CellSpace<string> { return this.cell(String)(() => new Area(PRINT_HEADER_MARGIN, 0, this.headerWidth, 2)) }
  @lazy get author(): CellSpace<string> { return this.cell(String)(() => new Area(PRINT_HEADER_MARGIN, 2, Math.floor(this.headerWidth / 2), 1)) }
  @lazy get timestamp(): CellSpace<string> { return this.cell(String)(() => new Area(PRINT_HEADER_MARGIN + Math.floor(this.headerWidth / 2), 2, Math.ceil(this.headerWidth / 2), 1)) }
  @lazy get info(): CellSpace<string> { return this.cell(String)(() => new Area(PRINT_HEADER_MARGIN, 3, this.headerWidth, 1)) }
}

class ControlPanelSheet extends SheetSpace {
  constructor() { super(CONTROL_PANEL_SHEET_NAME) }

  @lazy get author(): CellSpace<string> { return this.cell(String)(() => this.getNamedArea(AUTHOR_RANGE_NAME)) }
  @lazy get contentMarginH(): CellSpace<number> { return this.cell(Number, DEFAULT_CONTENT_MARGIN_H)(() => this.getNamedArea(CONTENT_MARGIN_H_RANGE_NAME)) }
  @lazy get contentMarginV(): CellSpace<number> { return this.cell(Number, DEFAULT_CONTENT_MARGIN_V)(() => this.getNamedArea(CONTENT_MARGIN_V_RANGE_NAME)) }
  @lazy get horizontalPadding(): CellSpace<number> { return this.cell(Number, DEFAULT_HORIZONTAL_PADDING)(() => this.getNamedArea(HORIZONTAL_PADDING_RANGE_NAME)) }
  @lazy get verticalPadding(): CellSpace<number> { return this.cell(Number, DEFAULT_VERTICAL_PADDING)(() => this.getNamedArea(VERTICAL_PADDING_RANGE_NAME)) }
}