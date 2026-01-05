import { Chord } from "./Chords"
import "./Range"

type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type Range = GoogleAppsScript.Spreadsheet.Range
type OnEdit = GoogleAppsScript.Events.SheetsOnEdit
type OnChange = GoogleAppsScript.Events.SheetsOnChange


const LYRICS_SHEET_NAME = "Letra"
const CHORDS_SHEET_NAME = "Acordes"
const PRINT_SHEET_NAME = "Impresión"

const LYRICS_RIGHT_TRAY_RANGE_NAME = "Ideas_Sueltas"
const PRINT_HEADER_RANGE_NAME = "Encabezado"
const PRINT_FOOTER_RANGE_NAME = "Pie_de_Página"
const DOCUMENT_TITLE_RANGE_NAME = "Título"
const KEY_RANGE_NAME = "Tonalidad"
const AUTOTRANSPOSE_RANGE_NAME = "Auto_Trasponer"

const FONT_FAMILY = "Space Mono"
const FONT_SIZE = 10

const ROW_HEIGHT = 21
const NORMAL_COLUMN_WIDTH = 15
const WIDE_COLUMN_WIDTH = 17
const WIDE_COLUMN_PERIODICITY = 6
const PADDING = 3


const STRUCTURAL_CHANGES = ["INSERT_ROW", "INSERT_COLUMN", "REMOVE_ROW", "REMOVE_COLUMN", "OTHER"]

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// HOOKS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function onOpen(): void {
  try {
    updateDocumentTitle()
  } catch (error) {
    warn("Unexpected error in onOpen hook", error instanceof Error ? error.message : undefined)
  }
}

export function onEdit(event: OnEdit): void {
  try {
    const editedRange = event.range

    switch (editedRange.getSheet().getName()) {
      case LYRICS_SHEET_NAME:
        syncLyricsToChordSheet(editedRange)
        break

      case CHORDS_SHEET_NAME:
        handleKeyChange(editedRange, event.oldValue)
        disableAutoTransposeIfKeyIsInvalid(editedRange)
        syncLyricsFromChordSheet(editedRange)
        break
    }
  } catch (error) {
    warn("Unexpected error in onEdit hook", error instanceof Error ? error.message : undefined)
  }
}

export function onChange(event: OnChange): void {
  try {
    if (STRUCTURAL_CHANGES.includes(event.changeType)) {
      const sourceSheet = LYRICS_SHEET()
      const effectiveRange = getWorkingArea(sourceSheet).resizeTo(undefined, sourceSheet.getDataRange().getNumRows())

      syncLyricsToChordSheet(effectiveRange)
    }
  } catch (error) {
    warn("Unexpected error in onChange hook", error instanceof Error ? error.message : undefined)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// SYNC LOGIC
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

function syncLyricsToChordSheet(range: Range): void {
  const sourceSheet = range.getSheet()
  const sourceWorkingArea = getWorkingArea(sourceSheet)

  const targetSheet = CHORDS_SHEET()
  const targetWorkingArea = getWorkingArea(targetSheet)

  const sourceRange = sourceWorkingArea.intersect(range)
  if (!sourceRange) return

  const targetRange = sourceRange
    .projectInto(targetSheet)
    .scale(1, 2)
    .translate(
      targetWorkingArea.getColumn() - sourceWorkingArea.getColumn(),
      targetWorkingArea.getRow() - sourceWorkingArea.getRow() + (sourceRange.getRow() - sourceWorkingArea.getRow())
    )

  const targetValues = targetRange.getValues()
  sourceRange.getValues().forEach((sourceRow, rowOffset) => {
    targetValues[rowOffset * 2 + 1] = sourceRow
  })
  targetRange.setValues(targetValues)
}


function syncLyricsFromChordSheet(range: Range): void {
  const sourceSheet = LYRICS_SHEET()
  const sourceWorkingArea = getWorkingArea(sourceSheet)

  const targetSheet = range.getSheet()
  const targetWorkingArea = getWorkingArea(targetSheet)

  const intersected = targetWorkingArea.intersect(range)
  if (!intersected) return

  const targetStartsAtChordRow = (intersected.getRow() - targetWorkingArea.getRow()) % 2 === 0
  if (targetStartsAtChordRow && intersected.getNumRows() === 1) return

  const targetRange = targetStartsAtChordRow
    ? intersected.translate(0, 1).resize(0, -1)
    : intersected

  const sourceRange = targetRange
    .projectInto(sourceSheet)
    .scale(1, 0.5)
    .translate(
      sourceWorkingArea.getColumn() - targetWorkingArea.getColumn(),
      sourceWorkingArea.getRow() - targetWorkingArea.getRow() - Math.ceil((targetRange.getRow() - targetWorkingArea.getRow()) / 2)
    )

  const targetValues = targetRange.getValues()
  sourceRange.getValues().forEach((sourceRow, rowOffset) => {
    targetValues[rowOffset * 2] = sourceRow
  })
  targetRange.setValues(targetValues)
}

function handleKeyChange(range: Range, oldValue: string | undefined): void {
  const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
  if (!keyRange?.overlapsWith(range)) return

  const autoTranspose = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)?.getValue()
  if (autoTranspose) {
    const newKey = Chord.parse(keyRange.getValue() ?? "")
    const oldKey = Chord.parse(oldValue ?? "")

    newKey && oldKey && transposeAllChords(oldKey.semitonesTo(newKey), false)
  }
}

function disableAutoTransposeIfKeyIsInvalid(editedRange: Range): void {
  const autoTransposeRange = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)
  if (!autoTransposeRange?.overlapsWith(editedRange) || !autoTransposeRange?.getValue()) return

  const key = Chord.parse(SPREADSHEET().getRangeByName(KEY_RANGE_NAME)?.getValue() ?? "")
  if (!key) autoTransposeRange.setValue(false)
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// ACTIONS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function transposeUp(): void { transposeAllChords(1) }

export function transposeDown(): void { transposeAllChords(-1) }

function markAsInvalid(value: unknown): string {
  const str = String(value)
  return str.startsWith("!") ? str : "!" + str
}


function transposeAllChords(semitones: number, updateKey: boolean = true): void {
  if (semitones === 0) return

  if (updateKey) {
    const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
    const key = keyRange && Chord.parse(keyRange.getValue())

    keyRange?.setValue(key?.transpose(semitones) ?? markAsInvalid(keyRange.getValue()))
  }

  const range = getWorkingArea(CHORDS_SHEET()).intersect(CHORDS_SHEET().getDataRange())
  if (!range) return

  const values = range.getValues()
  values.forEach((row, rowIndex) => {
    if (rowIndex % 2 === 1) return
    values[rowIndex] = row.map(cell => Chord.parse(cell)?.transpose(semitones) ?? markAsInvalid(cell))
  })
  range.setValues(values)
}


export function setupTriggers(): void {
  const triggers = ScriptApp.getProjectTriggers()

  const existingOnChange = triggers.find(trigger =>
    trigger.getHandlerFunction() === "onChange" &&
    trigger.getEventType() === ScriptApp.EventType.ON_CHANGE
  )

  if (existingOnChange) {
    info("Trigger setup", "onChange trigger already installed")
    return
  }

  try {
    ScriptApp.newTrigger("onChange")
      .forSpreadsheet(SPREADSHEET())
      .onChange()
      .create()

    success("Trigger setup", "onChange trigger installed successfully")
  } catch (e) {
    error("Trigger setup failed", e instanceof Error ? e.message : `${e}`)
  }
}

export function resetFormatting(): void {
  const printHeaderRange = SPREADSHEET().getRangeByName(PRINT_HEADER_RANGE_NAME)
  const printFooterRange = SPREADSHEET().getRangeByName(PRINT_FOOTER_RANGE_NAME)
  const printHeaderWidth = printHeaderRange?.getNumColumns()

  SPREADSHEET().getSheets().forEach(sheet => {
    const workingArea = getWorkingArea(sheet)
    const isPrintSheet = sheet.getName() === PRINT_SHEET_NAME

    for (let columnIndex = 1; columnIndex <= sheet.getMaxColumns(); columnIndex++) {
      const cellRange = sheet.getRange(1, columnIndex)
      if (isPrintSheet && (printHeaderRange?.overlapsWith(cellRange) || printFooterRange?.overlapsWith(cellRange))) continue

      const effectiveColumnIndex = isPrintSheet && printHeaderWidth
        ? ((columnIndex - 1) % printHeaderWidth) + 1
        : columnIndex

      const columnWidth = effectiveColumnIndex < workingArea.getColumn()
        ? WIDE_COLUMN_WIDTH + 2 * PADDING
        : effectiveColumnIndex === workingArea.getColumn()
          ? WIDE_COLUMN_WIDTH + PADDING
          : (effectiveColumnIndex - workingArea.getColumn()) % WIDE_COLUMN_PERIODICITY === 0
            ? WIDE_COLUMN_WIDTH
            : NORMAL_COLUMN_WIDTH

      sheet.setColumnWidth(columnIndex, columnWidth)
    }

    const maxRow = isPrintSheet && printFooterRange
      ? printFooterRange.getRow() - 1
      : sheet.getMaxRows()

    if (maxRow >= 1) {
      sheet.setRowHeights(1, maxRow, ROW_HEIGHT)
    }

    if (isPrintSheet && printHeaderRange) {
      const headerRow = printHeaderRange.getRow()
      const headerLastRow = printHeaderRange.getLastRow()
      const footerRow = printFooterRange?.getRow() ?? sheet.getMaxRows() + 1

      if (headerRow > 1) {
        sheet.fullRange().resizeTo(undefined, headerRow - 1)
          .setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE)
      }

      if (headerLastRow < footerRow - 1) {
        sheet.fullRange().resizeTo(undefined, footerRow - headerLastRow - 1).translateTo(undefined, headerLastRow + 1)
          .setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE)
      }
    } else {
      sheet.getDataRange().setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE)
    }
  })

  success("Format reset", "Done")
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

const LYRICS_SHEET = getSheet(LYRICS_SHEET_NAME)
const CHORDS_SHEET = getSheet(CHORDS_SHEET_NAME)


function getWorkingArea(sheet: Sheet): Range {
  const frozenRows = sheet.getFrozenRows()
  const frozenColumns = sheet.getFrozenColumns()
  const rightTrayWidth = sheet.getName() === LYRICS_SHEET_NAME
    ? SPREADSHEET().getRangeByName(LYRICS_RIGHT_TRAY_RANGE_NAME)?.getNumColumns() ?? 0
    : 0

  return sheet.fullRange()
    .resize(-frozenColumns - rightTrayWidth, -frozenRows)
    .translate(frozenColumns, frozenRows)
}


function updateDocumentTitle(): void {
  const titleRange = SPREADSHEET().getRangeByName(DOCUMENT_TITLE_RANGE_NAME)
  titleRange?.setValue(SPREADSHEET().getName())
}


const toast = (icon: string) => (title = "", message = "") => {
  SPREADSHEET().toast(message, icon + " " + title, 10)
}

const info = toast("ℹ️")
const success = toast("✅")
const warn = toast("⚠️")
const error = toast("❌")
