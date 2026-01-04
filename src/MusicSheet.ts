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
        validateAutoTranspose(editedRange)
        revertEvenChordsEdits(editedRange)
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
      targetWorkingArea.getRow() - sourceWorkingArea.getRow() * 2 + sourceRange.getRow()
    )

  const targetValues = targetRange.getValues()
  sourceRange.getValues().forEach((sourceRow, rowOffset) => {
    targetValues[rowOffset * 2 + 1] = sourceRow
  })
  targetRange.setValues(targetValues)
}


function revertEvenChordsEdits(editedRange: Range): void {
  const lyricsSheet = LYRICS_SHEET()

  const effectiveChordsRange = getWorkingArea(CHORDS_SHEET()).intersect(editedRange)
  if (!effectiveChordsRange) return

  const chordsSheet = effectiveChordsRange.getSheet()

  const chordsFrozenRows = chordsSheet.getFrozenRows()
  const chordsFrozenColumns = chordsSheet.getFrozenColumns()
  const lyricsFrozenRows = lyricsSheet.getFrozenRows()
  const lyricsFrozenColumns = lyricsSheet.getFrozenColumns()

  const chordsValues = effectiveChordsRange.getValues()

  const lyricsRowsToRead: {
    rowOffset: number
    absoluteLyricsRow: number
  }[] = []

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

function handleKeyChange(editedRange: Range, oldValue: string | undefined): void {
  const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
  if (!keyRange || !editedRange.overlapsWith(keyRange)) return

  const autoTransposeRange = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)
  if (!autoTransposeRange) return

  const newKey = keyRange.getValue()
  const isNewKeyValid = newKey && Chord.parse(newKey)

  if (!isNewKeyValid) {
    if (autoTransposeRange.getValue()) {
      autoTransposeRange.setValue(false)
    }
    return
  }

  if (!autoTransposeRange.getValue()) return

  if (!oldValue || oldValue === "") return

  const oldChord = Chord.parse(oldValue)
  const newChord = Chord.parse(newKey)
  if (!oldChord || !newChord) return

  const semitones = oldChord.semitonesTo(newChord)
  if (semitones === 0) return

  transposeChords(semitones, false)
}

function validateAutoTranspose(editedRange: Range): void {
  const autoTransposeRange = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)
  if (!autoTransposeRange || !editedRange.overlapsWith(autoTransposeRange)) return

  if (!autoTransposeRange.getValue()) return

  const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
  if (!keyRange) {
    autoTransposeRange.setValue(false)
    return
  }

  const keyValue = keyRange.getValue()
  const isKeyValid = keyValue && Chord.parse(keyValue)

  if (!isKeyValid) {
    autoTransposeRange.setValue(false)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// ACTIONS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function transposeUp(): void { transposeChords(1, true) }

export function transposeDown(): void { transposeChords(-1, true) }

function markAsInvalid(value: unknown): string {
  const str = String(value)
  return str.startsWith("!") ? str : "!" + str
}

function transposeChords(semitones: number, updateKey: boolean = true): void {
  if (updateKey) {
    const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
    if (keyRange) {
      const keyValue = keyRange.getValue() as string
      if (keyValue && keyValue !== "") {
        let transposedKey: string
        const chord = Chord.parse(keyValue)
        if (chord) {
          try {
            transposedKey = chord.transpose(semitones).toString()
          } catch {
            transposedKey = markAsInvalid(keyValue)
          }
        } else {
          transposedKey = markAsInvalid(keyValue)
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

      const chord = Chord.parse(cell as string)
      if (chord) {
        try {
          return chord.transpose(semitones).toString()
        } catch {
          return markAsInvalid(cell)
        }
      } else {
        return markAsInvalid(cell)
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

    for (let columnIndex = 1; columnIndex <= totalColumnCount; columnIndex++) {
      const cellRange = sheet.getRange(1, columnIndex, 1, 1)
      if ((isPrintSheet && printHeaderRange && cellRange.overlapsWith(printHeaderRange)) ||
        (isPrintSheet && printFooterRange && cellRange.overlapsWith(printFooterRange))) continue

      let columnWidth: number
      let effectiveColumnIndex = columnIndex

      if (isPrintSheet && printHeaderWidth) {
        effectiveColumnIndex = ((columnIndex - 1) % printHeaderWidth) + 1
      }

      columnWidth =
        effectiveColumnIndex <= frozenColumnCount
          ? columnWidth = WIDE_COLUMN_WIDTH + 2 * PADDING
          : effectiveColumnIndex === frozenColumnCount + 1
            ? columnWidth = WIDE_COLUMN_WIDTH + PADDING
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
  if (titleRange) {
    titleRange.setValue(SPREADSHEET().getName())
  }
}

const toast = (icon: string) => (title = "", message = "") => {
  SPREADSHEET().toast(message, icon + " " + title, 10)
}

const info = toast("ℹ️")
const success = toast("✅")
const warn = toast("⚠️")
const error = toast("❌")
