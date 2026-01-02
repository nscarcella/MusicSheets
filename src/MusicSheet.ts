import { parseChord, semitoneDistance, transpose } from "./Chords"

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

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// HOOKS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function onOpen(): void {
  try {
    updateDocumentTitle()
  } catch (error: unknown) {
    warn("Unexpected error in onOpen hook", (error as Error).message)
  }
}

export function onEdit(event: GoogleAppsScript.Events.SheetsOnEdit): void {
  try {
    const editedRange = event.range
    const sourceSheetName = editedRange.getSheet().getName()

    switch (sourceSheetName) {
      case LYRICS_SHEET_NAME:
        syncLyrics(editedRange, CHORDS_SHEET())
        break

      case CHORDS_SHEET_NAME:
        handleKeyChange(editedRange, event.oldValue)
        validateAutoTranspose(editedRange)
        revertEvenChordsEdits(editedRange)
        break
    }
  } catch (error: unknown) {
    warn("Unexpected error in onEdit hook", (error as Error).message)
  }
}

export function onChange(event: GoogleAppsScript.Events.SheetsOnChange): void {
  try {
    const structuralChanges = ["INSERT_ROW", "INSERT_COLUMN", "REMOVE_ROW", "REMOVE_COLUMN", "OTHER"] as const
    if (structuralChanges.includes(event.changeType as typeof structuralChanges[number])) {
      const lyricsSheet = LYRICS_SHEET()
      const fullRange = lyricsSheet.getRange(1, 1, lyricsSheet.getMaxRows(), lyricsSheet.getMaxColumns())
      syncLyrics(fullRange, CHORDS_SHEET())
    }
  } catch (error: unknown) {
    warn("Unexpected error in onChange hook", (error as Error).message)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// SYNC LOGIC
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

function syncLyrics(range: GoogleAppsScript.Spreadsheet.Range, targetSheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const effectiveRange = workingArea(range)
  if (!effectiveRange) return

  const sourceSheet = effectiveRange.getSheet()
  const targetFrozenRows = targetSheet.getFrozenRows()
  const targetFrozenColumns = targetSheet.getFrozenColumns()

  const lastRowWithContent = sourceSheet.getDataRange().getLastRow()
  const unfrozenRowsWithContent = Math.max(0, lastRowWithContent - sourceSheet.getFrozenRows())

  ensureMinDimensions(
    targetSheet,
    targetFrozenRows + unfrozenRowsWithContent * 2,
    targetFrozenColumns +
    effectiveRange.getLastColumn() -
    sourceSheet.getFrozenColumns()
  )

  const sourceValues = effectiveRange.getValues()
  const rowsToSync = Math.min(sourceValues.length, unfrozenRowsWithContent - (effectiveRange.getRow() - sourceSheet.getFrozenRows() - 1))

  if (rowsToSync === 0) return

  const targetStartRow = targetFrozenRows + (effectiveRange.getRow() - sourceSheet.getFrozenRows()) * 2
  const targetStartCol = targetFrozenColumns + effectiveRange.getColumn() - sourceSheet.getFrozenColumns()
  const targetRowCount = rowsToSync * 2
  const targetColCount = sourceValues[0].length

  const targetRange = targetSheet.getRange(
    targetStartRow,
    targetStartCol,
    targetRowCount,
    targetColCount
  )
  const targetValues = targetRange.getValues()

  for (let rowOffset = 0; rowOffset < rowsToSync; rowOffset++) {
    targetValues[rowOffset * 2] = sourceValues[rowOffset]
  }

  targetRange.setValues(targetValues)
}

interface LyricsRowToRead {
  rowOffset: number
  absoluteLyricsRow: number
}

function revertEvenChordsEdits(editedRange: GoogleAppsScript.Spreadsheet.Range): void {
  const effectiveChordsRange = workingArea(editedRange)
  if (!effectiveChordsRange) return

  const chordsSheet = effectiveChordsRange.getSheet()
  const lyricsSheet = LYRICS_SHEET()

  const chordsFrozenRows = chordsSheet.getFrozenRows()
  const chordsFrozenColumns = chordsSheet.getFrozenColumns()
  const lyricsFrozenRows = lyricsSheet.getFrozenRows()
  const lyricsFrozenColumns = lyricsSheet.getFrozenColumns()

  const chordsValues = effectiveChordsRange.getValues()

  const lyricsRowsToRead: LyricsRowToRead[] = []

  for (let rowOffset = 0; rowOffset < chordsValues.length; rowOffset++) {
    const absoluteChordsRow = effectiveChordsRange.getRow() + rowOffset
    const relativeRowIndex = absoluteChordsRow - chordsFrozenRows - 1

    if (relativeRowIndex % 2 === 0) continue

    const absoluteLyricsRow = lyricsFrozenRows + Math.floor(relativeRowIndex / 2) + 1
    lyricsRowsToRead.push({ rowOffset, absoluteLyricsRow })
  }

  if (lyricsRowsToRead.length === 0) return

  const minLyricsRow = Math.min(...lyricsRowsToRead.map(r => r.absoluteLyricsRow))
  const maxLyricsRow = Math.max(...lyricsRowsToRead.map(r => r.absoluteLyricsRow))
  const relativeChordsColumn = effectiveChordsRange.getColumn() - chordsFrozenColumns - 1
  const absoluteLyricsColumn = lyricsFrozenColumns + relativeChordsColumn + 1

  const lyricsRange = lyricsSheet.getRange(
    minLyricsRow,
    absoluteLyricsColumn,
    maxLyricsRow - minLyricsRow + 1,
    chordsValues[0].length
  )
  const lyricsValues = lyricsRange.getValues()

  for (const { rowOffset, absoluteLyricsRow } of lyricsRowsToRead) {
    const lyricsRowOffset = absoluteLyricsRow - minLyricsRow
    chordsValues[rowOffset] = lyricsValues[lyricsRowOffset]
  }

  effectiveChordsRange.setValues(chordsValues)
}

function handleKeyChange(editedRange: GoogleAppsScript.Spreadsheet.Range, oldValue: string | undefined): void {
  const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
  if (!keyRange) return

  const editedSheet = editedRange.getSheet()
  const keySheet = keyRange.getSheet()

  if (editedSheet.getName() !== keySheet.getName()) return

  const editedRow = editedRange.getRow()
  const editedCol = editedRange.getColumn()
  const keyRow = keyRange.getRow()
  const keyCol = keyRange.getColumn()

  if (editedRow !== keyRow || editedCol !== keyCol) return

  const autoTransposeRange = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)
  if (!autoTransposeRange) return

  const newKey = keyRange.getValue() as string
  const isNewKeyValid = newKey && newKey !== "" && parseChord(newKey)

  if (!isNewKeyValid) {
    if (autoTransposeRange.getValue()) {
      autoTransposeRange.setValue(false)
    }
    return
  }

  if (!autoTransposeRange.getValue()) return

  if (!oldValue || oldValue === "") return

  const semitones = semitoneDistance(oldValue, newKey)
  if (semitones === undefined) return

  if (semitones === 0) return

  transposeChords(semitones, false)
}

function validateAutoTranspose(editedRange: GoogleAppsScript.Spreadsheet.Range): void {
  const autoTransposeRange = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)
  if (!autoTransposeRange) return

  const editedSheet = editedRange.getSheet()
  const autoTransposeSheet = autoTransposeRange.getSheet()

  if (editedSheet.getName() !== autoTransposeSheet.getName()) return

  const editedRow = editedRange.getRow()
  const editedCol = editedRange.getColumn()
  const autoTransposeRow = autoTransposeRange.getRow()
  const autoTransposeCol = autoTransposeRange.getColumn()

  if (editedRow !== autoTransposeRow || editedCol !== autoTransposeCol) return

  if (!autoTransposeRange.getValue()) return

  const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
  if (!keyRange) {
    autoTransposeRange.setValue(false)
    return
  }

  const keyValue = keyRange.getValue() as string
  const isKeyValid = keyValue && keyValue !== "" && parseChord(keyValue)

  if (!isKeyValid) {
    autoTransposeRange.setValue(false)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// ACTIONS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function transposeUp(): void { transposeChords(1, true) }

export function transposeDown(): void { transposeChords(-1, true) }

function transposeChords(semitones: number, updateKey: boolean = true): void {
  if (updateKey) {
    const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
    if (keyRange) {
      const keyValue = keyRange.getValue() as string
      if (keyValue && keyValue !== "") {
        let transposedKey: string
        try {
          transposedKey = transpose(keyValue, semitones)
        } catch {
          transposedKey = keyValue.toString().startsWith("!") ? keyValue : "!" + keyValue
        }
        keyRange.setValue(transposedKey)
      }
    }
  }

  const chordsSheet = CHORDS_SHEET()
  const frozenRows = chordsSheet.getFrozenRows()
  const frozenColumns = chordsSheet.getFrozenColumns()
  const lastRow = chordsSheet.getDataRange().getLastRow()
  const lastColumn = chordsSheet.getDataRange().getLastColumn()

  if (lastRow <= frozenRows || lastColumn <= frozenColumns) return

  const workingRows = lastRow - frozenRows
  const workingColumns = lastColumn - frozenColumns

  const workingRange = chordsSheet.getRange(
    frozenRows + 1,
    frozenColumns + 1,
    workingRows,
    workingColumns
  )
  const values = workingRange.getValues()

  const transposedValues = values.map((row, rowIndex) => {
    if (rowIndex % 2 === 1) return row

    return row.map(cell => {
      if (cell === "" || cell === null) return cell

      try {
        return transpose(cell as string, semitones)
      } catch {
        const cellStr = cell.toString()
        return cellStr.startsWith("!") ? cell : "!" + cellStr
      }
    })
  })

  workingRange.setValues(transposedValues)
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
    error("Trigger setup failed", (e as Error).message)
  }
}

export function resetFormatting(): void {
  const printHeaderRange = SPREADSHEET().getRangeByName(PRINT_HEADER_RANGE_NAME)
  const printFooterRange = SPREADSHEET().getRangeByName(PRINT_FOOTER_RANGE_NAME)
  const printHeaderWidth = printHeaderRange ? printHeaderRange.getNumColumns() : null

  for (const sheet of SPREADSHEET().getSheets()) {
    const frozenColumnCount = sheet.getFrozenColumns()
    const totalColumnCount = sheet.getMaxColumns()
    const isPrintSheet = sheet.getName() === PRINT_SHEET_NAME
    const isInPrintHeader = isInRange(isPrintSheet ? printHeaderRange : null)
    const isInPrintFooter = isInRange(isPrintSheet ? printFooterRange : null)

    for (let columnIndex = 1; columnIndex <= totalColumnCount; columnIndex++) {
      if (isInPrintHeader(columnIndex, 1) || isInPrintFooter(columnIndex, 1)) continue

      let columnWidth: number
      let effectiveColumnIndex = columnIndex

      if (isPrintSheet && printHeaderWidth) {
        effectiveColumnIndex = ((columnIndex - 1) % printHeaderWidth) + 1
      }

      columnWidth =
        effectiveColumnIndex <= frozenColumnCount
          ? (columnWidth = WIDE_COLUMN_WIDTH + 2 * PADDING)
          : effectiveColumnIndex === frozenColumnCount + 1
            ? (columnWidth = WIDE_COLUMN_WIDTH + PADDING)
            : (effectiveColumnIndex - (frozenColumnCount + 1)) % WIDE_COLUMN_PERIODICITY === 0
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

    const dataRange = sheet.getDataRange()
    if (printHeaderRange && isPrintSheet) {
      const headerRow = printHeaderRange.getRow()
      const headerLastRow = printHeaderRange.getLastRow()
      const footerRow = printFooterRange ? printFooterRange.getRow() : sheet.getMaxRows() + 1

      if (headerRow > 1) {
        sheet.getRange(1, 1, headerRow - 1, sheet.getMaxColumns())
          .setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE)
      }

      if (headerLastRow < footerRow - 1) {
        sheet.getRange(headerLastRow + 1, 1, footerRow - headerLastRow - 1, sheet.getMaxColumns())
          .setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE)
      }
    } else {
      dataRange.setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE)
    }
  }

  success("Format reset", "Done")
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

const SPREADSHEET = (): GoogleAppsScript.Spreadsheet.Spreadsheet => SpreadsheetApp.getActive()

const getSheet = (sheetName: string) => (): GoogleAppsScript.Spreadsheet.Sheet => {
  const sheet = SPREADSHEET().getSheetByName(sheetName)
  if (!sheet) throw new Error(`${sheetName} sheet not found`)
  return sheet
}

const LYRICS_SHEET = getSheet(LYRICS_SHEET_NAME)
const CHORDS_SHEET = getSheet(CHORDS_SHEET_NAME)

const isInRange = (range: GoogleAppsScript.Spreadsheet.Range | null) =>
  (colIndex: number, rowIndex: number): boolean => {
    if (!range) return false
    return (
      rowIndex >= range.getRow() &&
      rowIndex <= range.getLastRow() &&
      colIndex >= range.getColumn() &&
      colIndex <= range.getLastColumn()
    )
  }

function workingArea(range: GoogleAppsScript.Spreadsheet.Range): GoogleAppsScript.Spreadsheet.Range | null {
  if (!range) return range

  const sheet = range.getSheet()
  const startRow = Math.max(range.getRow(), sheet.getFrozenRows() + 1)
  const startColumn = Math.max(range.getColumn(), sheet.getFrozenColumns() + 1)
  const endRow = range.getLastRow()
  let endColumn = range.getLastColumn()

  if (sheet.getName() === LYRICS_SHEET_NAME) {
    const namedRange = SPREADSHEET().getRangeByName(LYRICS_RIGHT_TRAY_RANGE_NAME)
    if (namedRange) {
      const rightTrayStartColumn = namedRange.getColumn()
      endColumn = Math.min(endColumn, rightTrayStartColumn - 1)
    }
  }

  if (startRow > endRow || startColumn > endColumn) return null

  return sheet.getRange(
    startRow,
    startColumn,
    endRow - startRow + 1,
    endColumn - startColumn + 1
  )
}

function ensureMinDimensions(sheet: GoogleAppsScript.Spreadsheet.Sheet, minRows: number, minColumns: number): void {
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows())
  }

  if (sheet.getMaxColumns() < minColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      minColumns - sheet.getMaxColumns()
    )
  }
}

function updateDocumentTitle(): void {
  const titleRange = SPREADSHEET().getRangeByName(DOCUMENT_TITLE_RANGE_NAME)
  if (titleRange) {
    titleRange.setValue(SPREADSHEET().getName())
  }
}

type ToastFunction = (title?: string, message?: string) => void

const toast = (icon: string): ToastFunction => (title = "", message = "") => {
  SPREADSHEET().toast(message, icon + " " + title, 10)
}

const info = toast("ℹ️")
const success = toast("✅")
const warn = toast("⚠️")
const error = toast("❌")
