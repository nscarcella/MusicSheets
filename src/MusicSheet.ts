import { Chord } from "./Chords"
import "./Range"
import { $, CellValue } from "./Spaces"
import { Area } from "./Area"

type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type Range = GoogleAppsScript.Spreadsheet.Range
type OnEdit = GoogleAppsScript.Events.SheetsOnEdit
type OnChange = GoogleAppsScript.Events.SheetsOnChange


const VERSION = "1.0"

const _LYRICS_SHEET_NAME = "Letra"
const _CHORDS_SHEET_NAME = "Acordes"
const _PRINT_SHEET_NAME = "Impresión"

const _LYRICS_RIGHT_TRAY_RANGE_NAME = "Ideas_Sueltas"
const _KEY_RANGE_NAME = "Tonalidad"
const AUTHOR_RANGE_NAME = "Autor"
const TEMPO_RANGE_NAME = "Tempo"
const NOTES_RANGE_NAME = "Notas"
const CONTENT_MARGIN_H_RANGE_NAME = "Margen_H"
const CONTENT_MARGIN_V_RANGE_NAME = "Margen_V"
const HEADER_MARGIN_H_RANGE_NAME = "Margen_Encabezado"
const HORIZONTAL_PADDING_RANGE_NAME = "Separación_H"
const VERTICAL_PADDING_RANGE_NAME = "Separación_V"

const FONT_FAMILY = "Space Mono"
const FONT_SIZE = 10

const ROW_HEIGHT = 21
const NORMAL_COLUMN_WIDTH = 15
const WIDE_COLUMN_WIDTH = 17
const WIDE_COLUMN_PERIODICITY = 6
const PADDING = 3


const TRIGGERS_INSTALLED_PROPERTY = "triggers_installed"

const PRINT_PAGE_WIDTH = 46
const PRINT_PAGE_HEIGHT = 51
const PRINT_HEADER_HEIGHT = 4
const PRINT_FOOTER_HEIGHT = 1
const DEFAULT_PRINT_HORIZONTAL_PADDING = 2
const DEFAULT_PRINT_VERTICAL_PADDING = 2
const DEFAULT_PRINT_CONTENT_MARGIN_H = 1
const DEFAULT_PRINT_CONTENT_MARGIN_V = 1
const DEFAULT_PRINT_HEADER_MARGIN_H = 1

function getNumericSetting(rangeName: string, defaultValue: number): number {
  const value = Number(SPREADSHEET().getRangeByName(rangeName)?.getValue())
  return Number.isFinite(value) && value >= 0 ? value : defaultValue
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// HOOKS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function onOpen(): void {
  try {
    if (!PropertiesService.getDocumentProperties().getProperty(TRIGGERS_INSTALLED_PROPERTY))
      SpreadsheetApp.getUi()
        .createMenu("⚠️")
        .addItem("Autorizar la extensión para habilitar la sincronización de texto avanzada", setupTriggers.name)
        .addToUi()

    SpreadsheetApp.getUi()
      .createMenu("🖨️ Impresión")
      .addItem("Regenerar hoja de impresión", regeneratePrint.name)
      .addToUi()
  } catch (error) {
    warn("Unexpected error in onOpen hook", error instanceof Error ? error.message : undefined)
  }
}

export function onEdit(event: OnEdit): void {
  try {
    const sheet = $.get(event.range.getSheet().getName())
    const changed = Area.fromRange(event.range)

    switch (sheet.name) {
      case $.Lyrics.name:
        restoreIndexes(changed)
        syncLyricsToChordSheet(changed)
        enforceChordHeight()
        break

      case $.Chords.name:
        syncLyricsFromChordSheet(changed)
        handleKeyChange(changed, event.oldValue)
        disableAutoTransposeIfKeyIsInvalid(changed)
        break
    }
  } catch (error) {
    warn("Unexpected error in onEdit hook", error instanceof Error ? error.message : undefined)
  }
}

// TODO: PENDING REFACTOR
export function onChange(event: OnChange): void {
  try {
    if (event.changeType === "INSERT_COLUMN" || event.changeType === "REMOVE_COLUMN") {
      const lyricsSheet = LYRICS_SHEET()
      const trayWidth = SPREADSHEET().getRangeByName(_LYRICS_RIGHT_TRAY_RANGE_NAME)?.getNumColumns() ?? 0
      const rangeStart = lyricsSheet.getFrozenColumns() + 1
      const rangeEnd = lyricsSheet.getMaxColumns() - trayWidth

      syncStructuralColumnChanges(detectChanges(popColumnIndexes(), rangeStart, rangeEnd))
      enforceChordWidth()
    }
    else if (event.changeType === "INSERT_ROW" || event.changeType === "REMOVE_ROW") {
      const lyricsSheet = LYRICS_SHEET()
      const rangeStart = lyricsSheet.getFrozenRows() + 1
      const rangeEnd = lyricsSheet.getMaxRows()

      syncStructuralRowChanges(detectChanges(popRowIndexes(), rangeStart, rangeEnd))
      enforceChordHeight()
    }
  } catch (error) {
    warn("Unexpected error in onChange hook", error instanceof Error ? error.message : undefined)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// SYNC LYRICS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function syncLyricsToChordSheet(changed: Area): void {
  const source = $.Lyrics.main.sub(area => area.intersect(changed))
  const target = $.Chords.main.sub(area => source.area
    .relativeTo($.Lyrics.main)
    .scale({ y: 2 })
    .translate(area)
  )

  const targetValues = target.getValues()
  source.getValues().forEach((row, i) => {
    targetValues[i * 2 + 1] = [...row]
  })
  target.setValues(targetValues)
}


export function syncLyricsFromChordSheet(changed: Area): void {
  syncLyricsToChordSheet(
    $.Chords.main.area
      .intersect(changed)
      .relativeTo($.Chords.main)
      .scale({ y: 0.5 })
      .translate($.Lyrics.main)
  )
}

function restoreIndexes(changed: Area): void {
  const changedColumns = $.Lyrics.indexColumn.sub(area => area.intersect(changed))
  if (!changedColumns.isEmpty) {
    changedColumns.setValues(Array.from({ length: changedColumns.height }, (_, i) => [changedColumns.y + i + 1]))
  }

  const changedRows = $.Lyrics.indexRow.sub(area => area.intersect(changed))
  if (!changedRows.isEmpty) {
    changedRows.setValues([Array.from({ length: changedRows.width }, (_, i) => changedRows.x + i + 1)])
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// KEY CHANGE
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

function handleKeyChange(changed: Area, oldValue: CellValue): void {
  const key = $.Chords.key
  if (!key.area.overlapsWith(changed)) return

  if ($.Chords.autotranspose.getValue()) {
    const newKey = Chord.parse(key.getValue() ?? "")
    const oldKey = Chord.parse(String(oldValue))
    newKey && oldKey && transposeAll(oldKey.semitonesTo(newKey), false)
  }
}

function disableAutoTransposeIfKeyIsInvalid(changed: Area): void {
  const key = $.Chords.key
  const autotranspose = $.Chords.autotranspose
  if (!key.area.overlapsWith(changed) && !autotranspose.area.overlapsWith(changed)) return
  if (autotranspose.getValue() && !Chord.parse(key.getValue() ?? "")) autotranspose.setValue(false)
}

export function transposeUp(): void { transposeAll(1) }

export function transposeDown(): void { transposeAll(-1) }

function markAsInvalid(value: string): string { return value.startsWith("!") ? value : "!" + value }

function transposeAll(semitones: number, updateKey: boolean = true): void {
  if (semitones === 0) return

  if (updateKey) {
    const key = $.Chords.key.getValue()
    $.Chords.key.setValue(Chord.parse(key)?.transpose(semitones).toString() ?? markAsInvalid(key))
  }

  const values = $.Chords.main.getValues()
  values.forEach((row, rowIndex) => {
    if (rowIndex % 2 === 1) return
    values[rowIndex] = row.map(cell => {
      if (!cell) return null
      const chord = String(cell)
      return Chord.parse(chord)?.transpose(semitones).toString() ?? markAsInvalid(chord)
    })
  })
  $.Chords.main.setValues(values)
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function resetFormatting(): void {
  $.ALL.forEach(space => {
    const mainStart = space.main.x + 1

    for (let columnIndex = 1; columnIndex <= space.width; columnIndex++) {
      const columnWidth = columnIndex < mainStart
        ? WIDE_COLUMN_WIDTH + 2 * PADDING
        : columnIndex === mainStart
          ? WIDE_COLUMN_WIDTH + PADDING
          : (columnIndex - mainStart) % WIDE_COLUMN_PERIODICITY === 0
            ? WIDE_COLUMN_WIDTH
            : NORMAL_COLUMN_WIDTH

      space.sheet.setColumnWidth(columnIndex, columnWidth)
    }

    space.range?.setFontFamily(FONT_FAMILY)
    if (space !== $.Print) {
      if (space.height >= 1) {
        space.sheet.setRowHeights(1, space.height, ROW_HEIGHT)
      }
      space.range?.setFontSize(FONT_SIZE)
    }
  })
}

function setupTriggers(): void {
  try {
    if (
      ScriptApp.getProjectTriggers().some(trigger =>
        trigger.getHandlerFunction() === onChange.name && trigger.getEventType() === ScriptApp.EventType.ON_CHANGE
      )
    ) return info("Trigger setup", `${onChange.name} trigger already installed`)

    ScriptApp.newTrigger(onChange.name)
      .forSpreadsheet($.SPREADSHEET)
      .onChange()
      .create()

    PropertiesService.getDocumentProperties().setProperty(TRIGGERS_INSTALLED_PROPERTY, "true")

    success("Trigger setup", `${onChange.name} trigger installed successfully`)

  } catch (e) {
    error("Trigger setup failed", e instanceof Error ? e.message : `${e}`)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// STRUCTURAL SYNC
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export type StructuralChange = {
  position: number
  span: number
}

export function popColumnIndexes(): (number | null)[] {
  const currentValues = $.Lyrics.indexRow.getValues()[0]
  $.Lyrics.indexRow.setValues([currentValues.map((_, i) => i + 1)])
  return currentValues.map(v => typeof v === "number" && v > 0 ? v : null)
}

export function popRowIndexes(): (number | null)[] {
  const currentValues = $.Lyrics.indexColumn.getValues()
  $.Lyrics.indexColumn.setValues(currentValues.map((_, i) => [i + 1]))
  return currentValues.map(([v]) => typeof v === "number" && v > 0 ? v : null)
}

export function detectChanges(indexes: (number | null)[], from: number, to: number): StructuralChange[] {
  const changes: StructuralChange[] = []
  let i = indexes.length - 1

  while (i >= 0) {
    const current = indexes[i]

    if (current === null) {
      const nullEnd = i
      while (i > 0 && indexes[i - 1] === null) i--
      const span = nullEnd - i + 1

      const position = i > 0
        ? indexes[i - 1]! + 1
        : indexes[nullEnd + 1] ?? 1

      if (position >= from && position <= to + span) {
        const clampedStart = Math.max(position, from)
        const clampedEnd = Math.min(position + span - 1, to + span - 1)
        if (clampedEnd >= clampedStart) {
          changes.push({ position: clampedStart - from + 1, span: clampedEnd - clampedStart + 1 })
        }
      }
    } else {
      const previous = indexes[i - 1]
      const expected = i > 0 ? (previous ?? current) + 1 : 1
      const gap = current - expected

      if (gap > 0 && expected <= to && expected + gap > from) {
        const clampedStart = Math.max(expected, from)
        const clampedEnd = Math.min(expected + gap - 1, to)
        if (clampedEnd >= clampedStart) {
          changes.push({ position: clampedStart - from + 1, span: -(clampedEnd - clampedStart + 1) })
        }
      }
    }

    i--
  }

  return changes
}

export function applyStructuralColumnChanges(values: CellValue[][], changes: StructuralChange[]): CellValue[][] {
  const result = values.map(row => [...row])

  for (const { position, span } of changes) {
    const index = position - 1
    if (span > 0) {
      result.forEach(row => row.splice(index, 0, ...Array(span).fill("")))
    } else {
      result.forEach(row => row.splice(index, -span))
    }
  }

  return result
}

export function applyStructuralRowChanges(values: CellValue[][], changes: StructuralChange[]): CellValue[][] {
  const result = values.map(row => [...row])
  if (result.length === 0) return result

  const rowWidth = result[0].length

  for (const { position, span } of changes) {
    const index = (position - 1) * 2
    const scaledSpan = span * 2
    if (scaledSpan > 0) {
      result.splice(index, 0, ...Array.from({ length: scaledSpan }, () => Array(rowWidth).fill("")))
    } else {
      result.splice(index, -scaledSpan)
    }
  }

  return result
}

function syncStructuralColumnChanges(changes: StructuralChange[]): void {
  if (!changes.length) return

  const newValues = applyStructuralColumnChanges($.Chords.main.getValues(), changes)

  $.Chords.main
    .sub(area => area.resizeTo({ x: newValues[0].length, y: newValues.length }))
    .setValues(newValues)
}

function syncStructuralRowChanges(changes: StructuralChange[]): void {
  if (!changes.length) return

  const newValues = applyStructuralRowChanges($.Chords.main.getValues(), changes)
  const targetHeight = $.Lyrics.getLastRowWithContent() * 2

  $.Chords.main
    .sub(area => area.resizeTo({ x: newValues[0].length, y: targetHeight }))
    .setValues(newValues.slice(0, targetHeight))
}

function enforceChordWidth(): void {
  const targetMaxColumns = Math.max($.Chords.frozenColumns.width + $.Lyrics.main.width, $.Chords.header.width)
  const excess = $.Chords.width - targetMaxColumns

  if (excess > 0) {
    $.Chords.sheet.deleteColumns(targetMaxColumns + 1, excess)
  }
}

function enforceChordHeight(): void {
  const targetMaxRows = $.Chords.frozenRows.height + Math.max($.Lyrics.getLastRowWithContent() * 2, 1)
  const excess = $.Chords.height - targetMaxRows

  if (excess > 0) {
    $.Chords.sheet.deleteRows(targetMaxRows + 1, excess)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// PRINT SHEET GENERATION
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function detectSectionRanges(range: Range): Range[] {
  const values = range.getValues()
  if (values.length === 0 || values[0].length === 0) return []

  const sectionColumns = splitIntoSectionColumns(values)

  let columnOffset = 0
  return sectionColumns.flatMap(sectionColumn => {
    const columnWidth = sectionColumn[0].length
    const sections = splitIntoSections(sectionColumn).map(({ startRow, endRow, width }) =>
      range
        .translateTo(range.getColumn() + columnOffset, range.getRow() + startRow)
        .resizeTo(width, endRow - startRow)
    )
    columnOffset += columnWidth
    return sections
  })
}

export function splitIntoSectionColumns(values: unknown[][]): unknown[][][] {
  const columns: unknown[][][] = []
  let currentEndCol = values[0].length

  for (let col = values[0].length - 1; col >= 0; col--) {
    const hasLyricStart = values.some((row, rowIndex) => rowIndex % 2 === 1 && row[col] !== "")

    if (hasLyricStart && col < currentEndCol - 1) {
      columns.unshift(values.map(row => row.slice(col, currentEndCol)))
      currentEndCol = col
    }
  }

  if (currentEndCol > 0) {
    const hasAnyContent = values.some(row => row.slice(0, currentEndCol).some(cell => cell !== ""))
    if (hasAnyContent) {
      columns.unshift(values.map(row => row.slice(0, currentEndCol)))
    }
  }

  return columns
}

export type SectionBounds = { startRow: number, endRow: number, width: number }

export type PageLayout = Range[][][]

export function regeneratePrint(): void {
  const chordsSheet = CHORDS_SHEET()
  const spreadsheet = SPREADSHEET()

  const contentMarginH = getNumericSetting(CONTENT_MARGIN_H_RANGE_NAME, DEFAULT_PRINT_CONTENT_MARGIN_H)
  const contentMarginV = getNumericSetting(CONTENT_MARGIN_V_RANGE_NAME, DEFAULT_PRINT_CONTENT_MARGIN_V)
  const horizontalPadding = getNumericSetting(HORIZONTAL_PADDING_RANGE_NAME, DEFAULT_PRINT_HORIZONTAL_PADDING)
  const verticalPadding = getNumericSetting(VERTICAL_PADDING_RANGE_NAME, DEFAULT_PRINT_VERTICAL_PADDING)

  const availableWidth = PRINT_PAGE_WIDTH - 2 * contentMarginH
  const availableHeight = PRINT_PAGE_HEIGHT - PRINT_FOOTER_HEIGHT - contentMarginV

  const sectionRanges = detectSectionRanges(getWorkingArea(chordsSheet))

  const oversizedSection = sectionRanges.find(
    s => s.getNumRows() > availableHeight || s.getNumColumns() > availableWidth
  )
  if (oversizedSection) {
    error("Sección demasiado grande", "Una sección excede los límites de la página")
    return
  }

  const layout = calculateLayout(
    sectionRanges,
    availableWidth,
    availableHeight,
    PRINT_HEADER_HEIGHT + contentMarginV,
    horizontalPadding,
    verticalPadding
  )

  const totalPages = layout.length || 1
  const requiredColumns = totalPages * PRINT_PAGE_WIDTH

  const emptyRichText = SpreadsheetApp.newRichTextValue().setText("").build()
  const grid: GoogleAppsScript.Spreadsheet.RichTextValue[][] = Array.from(
    { length: PRINT_PAGE_HEIGHT },
    () => Array(requiredColumns).fill(emptyRichText)
  )

  for (let pageIndex = 0; pageIndex < layout.length; pageIndex++) {
    const page = layout[pageIndex]
    const pageColumnOffset = pageIndex * PRINT_PAGE_WIDTH + contentMarginH
    const pageContentStartRow = pageIndex === 0 ? PRINT_HEADER_HEIGHT + contentMarginV : contentMarginV

    const columnWidths = page.map(column => Math.max(...column.map(s => s.getNumColumns())))
    const totalColumnsWidth = columnWidths.reduce((sum, w) => sum + w, 0)
    const gaps = page.length - 1
    const extraSpace = availableWidth - totalColumnsWidth - gaps * horizontalPadding
    const effectivePadding = gaps > 0
      ? horizontalPadding + Math.floor(extraSpace / gaps)
      : 0

    let columnOffset = 0
    for (let colIndex = 0; colIndex < page.length; colIndex++) {
      const column = page[colIndex]
      let rowOffset = 0
      for (let sectionIndex = 0; sectionIndex < column.length; sectionIndex++) {
        const section = column[sectionIndex]
        if (sectionIndex > 0) rowOffset += verticalPadding
        const sourceData = section.getRichTextValues()
        for (let r = 0; r < sourceData.length; r++) {
          for (let c = 0; c < sourceData[r].length; c++) {
            grid[pageContentStartRow + rowOffset + r][pageColumnOffset + columnOffset + c] =
              sourceData[r][c] ?? emptyRichText
          }
        }
        rowOffset += section.getNumRows()
      }
      columnOffset += columnWidths[colIndex] + (colIndex < page.length - 1 ? effectivePadding : 0)
    }
  }

  let printSheet = spreadsheet.getSheetByName(_PRINT_SHEET_NAME)
  if (printSheet) {
    const rows = printSheet.getMaxRows()
    const cols = printSheet.getMaxColumns()
    if (rows > 1) printSheet.deleteRows(2, rows - 1)
    if (cols > 1) printSheet.deleteColumns(2, cols - 1)
    printSheet.clear()
  } else {
    printSheet = spreadsheet.insertSheet(_PRINT_SHEET_NAME, spreadsheet.getNumSheets())
  }

  printSheet.insertRows(1, PRINT_PAGE_HEIGHT - 1)
  printSheet.insertColumns(1, requiredColumns - 1)

  printSheet.getRange(1, 1, PRINT_PAGE_HEIGHT, requiredColumns).setRichTextValues(grid)

  for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
    const pageRange = printSheet.getRange(1, pageIndex * PRINT_PAGE_WIDTH + 1, PRINT_PAGE_HEIGHT - PRINT_FOOTER_HEIGHT, PRINT_PAGE_WIDTH)
    pageRange.setBorder(true, true, true, true, false, false)
  }

  generatePrintHeader(printSheet)
  for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
    generatePrintFooter(printSheet, pageIndex, totalPages)
  }

  resetFormatting()
}

function generatePrintHeader(printSheet: Sheet): void {
  const author = String(SPREADSHEET().getRangeByName(AUTHOR_RANGE_NAME)?.getValue() ?? "")
  const key = String(SPREADSHEET().getRangeByName(_KEY_RANGE_NAME)?.getValue() ?? "")
  const tempo = String(SPREADSHEET().getRangeByName(TEMPO_RANGE_NAME)?.getValue() ?? "")
  const notes = String(SPREADSHEET().getRangeByName(NOTES_RANGE_NAME)?.getValue() ?? "")

  const headerMarginH = getNumericSetting(HEADER_MARGIN_H_RANGE_NAME, DEFAULT_PRINT_HEADER_MARGIN_H)
  const headerWidth = PRINT_PAGE_WIDTH - 2 * headerMarginH
  const headerStart = 1 + headerMarginH

  const titleRange = printSheet.getRange(1, headerStart, 2, headerWidth)
  titleRange.merge()
  titleRange.setValue(SPREADSHEET().getName())
  titleRange.setFontSize(14)
  titleRange.setFontWeight("bold")
  titleRange.setVerticalAlignment("bottom")

  const authorRange = printSheet.getRange(3, headerStart, 1, Math.floor(headerWidth / 2))
  authorRange.merge()
  authorRange.setValue(author)
  authorRange.setFontSize(8)

  const timestamp = `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}:${VERSION}`
  const timestampRange = printSheet.getRange(3, headerStart + Math.floor(headerWidth / 2), 1, Math.ceil(headerWidth / 2))
  timestampRange.merge()
  timestampRange.setValue(timestamp)
  timestampRange.setHorizontalAlignment("right")
  timestampRange.setFontSize(8)

  const keyTempo = [key ? `[${key}]` : "", tempo ? `${tempo} bpm` : ""].filter(Boolean).join(" ")
  const prefix = keyTempo && notes ? `${keyTempo} | ` : keyTempo
  const fullText = prefix + notes
  const infoRange = printSheet.getRange(4, headerStart, 1, headerWidth)
  infoRange.merge()
  if (notes) {
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(fullText)
      .setTextStyle(prefix.length, fullText.length, SpreadsheetApp.newTextStyle().setItalic(true).build())
      .build()
    infoRange.setRichTextValue(richText)
  } else {
    infoRange.setValue(fullText)
  }
  infoRange.setFontSize(9)
  infoRange.setVerticalAlignment("top")
  infoRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)

  const headerRange = printSheet.getRange(1, 1, PRINT_HEADER_HEIGHT, PRINT_PAGE_WIDTH)
  headerRange.setBorder(true, true, true, true, false, false)
}

function generatePrintFooter(printSheet: Sheet, pageIndex: number, totalPages: number): void {
  const footerText = `${SPREADSHEET().getName()}  ${pageIndex + 1}/${totalPages}`

  const footerRange = printSheet.getRange(PRINT_PAGE_HEIGHT, pageIndex * PRINT_PAGE_WIDTH + 1, 1, PRINT_PAGE_WIDTH)
  footerRange.merge()
  footerRange.setValue(footerText)
  footerRange.setHorizontalAlignment("right")
  footerRange.setFontSize(8)
  printSheet.setRowHeight(PRINT_PAGE_HEIGHT, 13)
}

export function calculateLayout(
  sections: Range[],
  pageWidth: number,
  pageHeight: number,
  firstPageHeaderHeight: number,
  horizontalPadding: number = 0,
  verticalPadding: number = 0
): PageLayout {
  const pages: PageLayout = []
  let currentPage: Range[][] = []
  let currentColumnWidth = 0
  let currentColumnHeight = 0
  let currentRowWidth = 0

  const getAvailableHeight = () =>
    pages.length === 0 ? pageHeight - firstPageHeaderHeight : pageHeight

  const canStackInCurrentColumn = (sectionHeight: number) =>
    currentPage.length > 0 &&
    currentColumnHeight + verticalPadding + sectionHeight <= getAvailableHeight()

  const canFitNewColumn = (newColumnWidth: number) => {
    const hPadding = currentPage.length > 0 ? horizontalPadding : 0
    return currentRowWidth + currentColumnWidth + hPadding + newColumnWidth <= pageWidth
  }

  const canGrowCurrentColumn = (newWidth: number) =>
    currentRowWidth + newWidth <= pageWidth

  for (const section of sections) {
    const sectionWidth = section.getNumColumns()
    const sectionHeight = section.getNumRows()

    if (currentPage.length === 0) {
      currentPage.push([section])
      currentColumnWidth = sectionWidth
      currentColumnHeight = sectionHeight
    } else if (canStackInCurrentColumn(sectionHeight)) {
      const grownWidth = Math.max(currentColumnWidth, sectionWidth)
      if (canGrowCurrentColumn(grownWidth)) {
        currentPage[currentPage.length - 1].push(section)
        currentColumnHeight += verticalPadding + sectionHeight
        currentColumnWidth = grownWidth
      } else if (canFitNewColumn(sectionWidth)) {
        currentRowWidth += currentColumnWidth + horizontalPadding
        currentPage.push([section])
        currentColumnWidth = sectionWidth
        currentColumnHeight = sectionHeight
      } else {
        pages.push(currentPage)
        currentPage = [[section]]
        currentColumnWidth = sectionWidth
        currentRowWidth = 0
        currentColumnHeight = sectionHeight
      }
    } else if (canFitNewColumn(sectionWidth)) {
      currentRowWidth += currentColumnWidth + horizontalPadding
      currentPage.push([section])
      currentColumnWidth = sectionWidth
      currentColumnHeight = sectionHeight
    } else {
      pages.push(currentPage)
      currentPage = [[section]]
      currentColumnWidth = sectionWidth
      currentRowWidth = 0
      currentColumnHeight = sectionHeight
    }
  }

  if (currentPage.length > 0) {
    pages.push(currentPage)
  }

  return pages
}

export function splitIntoSections(values: unknown[][]): SectionBounds[] {
  const sections: SectionBounds[] = []
  let sectionStartRow: number | null = null

  for (let pairIndex = 0; pairIndex < values.length / 2; pairIndex++) {
    const chordRow = values[pairIndex * 2]
    const lyricRow = values[pairIndex * 2 + 1]
    const isEmpty = chordRow.every(cell => cell === "") && (!lyricRow || lyricRow.every(cell => cell === ""))

    if (isEmpty) {
      if (sectionStartRow !== null) {
        sections.push({
          startRow: sectionStartRow,
          endRow: pairIndex * 2,
          width: calculateSectionWidth(values.slice(sectionStartRow, pairIndex * 2))
        })
        sectionStartRow = null
      }
    } else {
      if (sectionStartRow === null) {
        sectionStartRow = pairIndex * 2
      }
    }
  }

  if (sectionStartRow !== null) {
    sections.push({
      startRow: sectionStartRow,
      endRow: values.length,
      width: calculateSectionWidth(values.slice(sectionStartRow))
    })
  }

  return sections
}

export function calculateSectionWidth(sectionData: unknown[][]): number {
  let maxLyricWidth = 0
  let maxChordEnd = 0

  for (let row = 0; row < sectionData.length; row++) {
    if (row % 2 === 1) {
      const lyricText = String(sectionData[row][0] ?? "")
      maxLyricWidth = Math.max(maxLyricWidth, Math.ceil(lyricText.length / 2))
    } else {
      for (let col = sectionData[row].length - 1; col >= 0; col--) {
        const chordText = String(sectionData[row][col] ?? "")
        if (chordText !== "") {
          const chordEnd = col + 1 + Math.ceil(chordText.length / 2)
          maxChordEnd = Math.max(maxChordEnd, chordEnd)
          break
        }
      }
    }
  }

  return Math.max(maxLyricWidth, maxChordEnd, 1)
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

const SPREADSHEET = (): Spreadsheet => SpreadsheetApp.getActive()

const getSheet = (sheetName: string) => (): Sheet => {
  const sheet = SPREADSHEET().getSheetByName(sheetName)
  if (!sheet) throw new Error(`${sheetName} sheet not found`)
  return sheet
}

const LYRICS_SHEET = getSheet(_LYRICS_SHEET_NAME)
const CHORDS_SHEET = getSheet(_CHORDS_SHEET_NAME)


function getWorkingArea(sheet: Sheet): Range {
  const frozenRows = sheet.getFrozenRows()
  const frozenColumns = sheet.getFrozenColumns()
  const rightTrayWidth = sheet.getName() === _LYRICS_SHEET_NAME
    ? SPREADSHEET().getRangeByName(_LYRICS_RIGHT_TRAY_RANGE_NAME)?.getNumColumns() ?? 0
    : 0

  return sheet.fullRange()
    .resize(-frozenColumns - rightTrayWidth, -frozenRows)
    .translate(frozenColumns, frozenRows)
}

const toast = (icon: string) => (title = "", message = "") => {
  SPREADSHEET().toast(message, icon + " " + title, 10)
}

const info = toast("ℹ️")
const success = toast("✅")
const warn = toast("⚠️")
const error = toast("❌")