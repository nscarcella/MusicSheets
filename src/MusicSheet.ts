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
const CHORDS_HEADER_RANGE_NAME = "Encabezado_Acordes"

const FONT_FAMILY = "Space Mono"
const FONT_SIZE = 10

const ROW_HEIGHT = 21
const NORMAL_COLUMN_WIDTH = 15
const WIDE_COLUMN_WIDTH = 17
const WIDE_COLUMN_PERIODICITY = 6
const PADDING = 3


const TRIGGERS_INSTALLED_PROPERTY = "triggers_installed"

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// HOOKS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function onOpen(): void {
  try {
    updateDocumentTitle()
    createMissingTriggerWarning()
  } catch (error) {
    warn("Unexpected error in onOpen hook", error instanceof Error ? error.message : undefined)
  }
}

export function onEdit(event: OnEdit): void {
  try {
    const editedRange = event.range
    const editedSheet = editedRange.getSheet()

    switch (editedSheet.getName()) {
      case LYRICS_SHEET_NAME:
        syncLyricsToChordSheet(editedRange)
        enforceChordHeight()
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
    if (event.changeType === "INSERT_COLUMN" || event.changeType === "REMOVE_COLUMN") {
      const lyricsSheet = LYRICS_SHEET()
      const trayWidth = SPREADSHEET().getRangeByName(LYRICS_RIGHT_TRAY_RANGE_NAME)?.getNumColumns() ?? 0
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

  const blackText = SpreadsheetApp.newTextStyle().setForegroundColor("#000000").build()

  const sourceRichText = sourceRange.getRichTextValues()
  const targetRichText = targetRange.getRichTextValues()

  const emptyRichText = SpreadsheetApp.newRichTextValue().setText("").build()

  sourceRichText.forEach((sourceRow, rowOffset) => {
    targetRichText[rowOffset * 2 + 1] = sourceRow.map(rt =>
      rt?.getText() ? rt.copy().setTextStyle(0, rt.getText().length, blackText).build() : emptyRichText
    )
  })

  targetRange.setRichTextValues(targetRichText.map(row => row.map(rt => rt ?? emptyRichText)))
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
    .intersect(sourceWorkingArea)

  if (!sourceRange) return

  const blackText = SpreadsheetApp.newTextStyle().setForegroundColor("#000000").build()

  const sourceRichText = sourceRange.getRichTextValues()
  const targetRichText = targetRange.getRichTextValues()

  const emptyRichText = SpreadsheetApp.newRichTextValue().setText("").build()

  sourceRichText.forEach((sourceRow, rowOffset) => {
    targetRichText[rowOffset * 2] = sourceRow.map(rt =>
      rt?.getText() ? rt.copy().setTextStyle(0, rt.getText().length, blackText).build() : emptyRichText
    )
  })

  targetRange
    .resizeTo(sourceRange.getNumColumns(), targetRange.getNumRows())
    .setRichTextValues(targetRichText.map(row => row.slice(0, sourceRange.getNumColumns()).map(rt => rt ?? emptyRichText)))
}

function handleKeyChange(range: Range, oldValue: string | undefined): void {
  const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
  if (!keyRange?.overlapsWith(range)) return

  const autoTranspose = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)?.getValue()
  if (autoTranspose) {
    const newKey = Chord.parse(keyRange.getValue() ?? "")
    const oldKey = Chord.parse(oldValue ?? "")

    newKey && oldKey && transposeAll(oldKey.semitonesTo(newKey), false)
  }
}

function disableAutoTransposeIfKeyIsInvalid(editedRange: Range): void {
  const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
  const autoTransposeRange = SPREADSHEET().getRangeByName(AUTOTRANSPOSE_RANGE_NAME)

  if (!keyRange?.overlapsWith(editedRange) && !autoTransposeRange?.overlapsWith(editedRange)) return
  if (!autoTransposeRange?.getValue()) return

  const key = Chord.parse(keyRange?.getValue() ?? "")
  if (!key) autoTransposeRange.setValue(false)
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// ACTIONS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function transposeUp(): void { transposeAll(1) }

export function transposeDown(): void { transposeAll(-1) }

function markAsInvalid(value: unknown): string {
  const str = String(value)
  return str.startsWith("!") ? str : "!" + str
}


function transposeAll(semitones: number, updateKey: boolean = true): void {
  if (semitones === 0) return

  if (updateKey) {
    const keyRange = SPREADSHEET().getRangeByName(KEY_RANGE_NAME)
    const key = keyRange && Chord.parse(keyRange.getValue())

    keyRange?.setValue(key?.transpose(semitones).toString() ?? markAsInvalid(keyRange.getValue()))
  }

  const range = getWorkingArea(CHORDS_SHEET()).intersect(CHORDS_SHEET().getDataRange())
  if (!range) return

  const values = range.getValues()
  values.forEach((row, rowIndex) => {
    if (rowIndex % 2 === 1) return
    values[rowIndex] = row.map(cell => cell && (Chord.parse(`${cell}`)?.transpose(semitones).toString() ?? markAsInvalid(cell)))
  })
  range.setValues(values)
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
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// TRIGGER SETUP
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

function createMissingTriggerWarning(): void {
  if (PropertiesService.getDocumentProperties().getProperty(TRIGGERS_INSTALLED_PROPERTY)) return

  SpreadsheetApp.getUi()
    .createMenu("⚠️")
    .addItem("Autorizar la extensión para habilitar la sincronización de texto avanzada", setupTriggers.name)
    .addToUi()
}

function setupTriggers(): void {
  if (
    ScriptApp.getProjectTriggers().some(trigger =>
      trigger.getHandlerFunction() === onChange.name && trigger.getEventType() === ScriptApp.EventType.ON_CHANGE
    )
  ) return info("Trigger setup", `${onChange.name} trigger already installed`)

  try {
    ScriptApp.newTrigger(onChange.name)
      .forSpreadsheet(SPREADSHEET())
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
  type: "insertion" | "deletion"
  position: number
  span: number
}

export function popColumnIndexes(): (number | null)[] {
  const sheet = LYRICS_SHEET()
  const indexRange = sheet.getRange(1, 1, 1, sheet.getMaxColumns())

  const currentValues = indexRange.getValues()[0]
  indexRange.setValues([currentValues.map((_, i) => i + 1)])

  return currentValues.map(v => typeof v === "number" && v > 0 ? v : null)
}

export function popRowIndexes(): (number | null)[] {
  const sheet = LYRICS_SHEET()
  const indexRange = sheet.getRange(1, 1, sheet.getMaxRows(), 1)

  const currentValues = indexRange.getValues()
  indexRange.setValues(currentValues.map((_, i) => [i + 1]))

  return currentValues.map(([v]) => typeof v === "number" && v > 0 ? v : null)
}

export function detectChanges(indexes: (number | null)[], workingAreaStart: number, workingAreaEnd: number): StructuralChange[] {
  const changes: StructuralChange[] = []

  for (let i = indexes.length - 1; i >= 0; i--) {
    const current = indexes[i]
    const previous = indexes[i - 1]

    if (current === null) {
      const afterNulls = i
      let span = 1
      while (i > 0 && indexes[i - 1] === null) {
        span++
        i--
      }
      const firstAfterNulls = indexes[afterNulls + 1]
      const position = i > 0
        ? indexes[i - 1]! + 1
        : firstAfterNulls ?? 1

      if (position >= workingAreaStart && position <= workingAreaEnd + span) {
        const clampedStart = Math.max(position, workingAreaStart)
        const clampedEnd = Math.min(position + span - 1, workingAreaEnd + span - 1)
        const clampedSpan = clampedEnd - clampedStart + 1
        if (clampedSpan > 0) {
          changes.push({ type: "insertion", position: clampedStart - workingAreaStart + 1, span: clampedSpan })
        }
      }
    } else {
      const expected = i > 0 ? (previous ?? current) + 1 : 1
      const gap = current - expected
      if (gap > 0 && expected <= workingAreaEnd && expected + gap > workingAreaStart) {
        const clampedStart = Math.max(expected, workingAreaStart)
        const clampedEnd = Math.min(expected + gap - 1, workingAreaEnd)
        const clampedSpan = clampedEnd - clampedStart + 1
        if (clampedSpan > 0) {
          changes.push({ type: "deletion", position: clampedStart - workingAreaStart + 1, span: clampedSpan })
        }
      }
    }
  }

  return changes
}

export function applyStructuralColumnChanges(values: unknown[][], changes: StructuralChange[]): unknown[][] {
  const result = values.map(row => [...row])

  for (const change of changes) {
    const index = change.position - 1
    if (change.type === "insertion") {
      result.forEach(row => row.splice(index, 0, ...Array(change.span).fill("")))
    } else {
      result.forEach(row => row.splice(index, change.span))
    }
  }

  return result
}

export function applyStructuralRowChanges(values: unknown[][], changes: StructuralChange[]): unknown[][] {
  const result = values.map(row => [...row])
  if (result.length === 0) return result

  const rowWidth = result[0].length

  for (const change of changes) {
    const index = (change.position - 1) * 2
    const span = change.span * 2
    if (change.type === "insertion") {
      const emptyRows = Array.from({ length: span }, () => Array(rowWidth).fill(""))
      result.splice(index, 0, ...emptyRows)
    } else {
      result.splice(index, span)
    }
  }

  return result
}

function syncStructuralColumnChanges(changes: StructuralChange[]): void {
  if (!changes.length) return

  const workingArea = getWorkingArea(CHORDS_SHEET())

  const newValues = applyStructuralColumnChanges(workingArea.getValues(), changes)
  const newWidth = newValues[0]?.length ?? 0

  if (newWidth > 0) {
    workingArea.resizeTo(newWidth, newValues.length)
      .setValues(newValues)
  }
}

function syncStructuralRowChanges(changes: StructuralChange[]): void {
  if (!changes.length) return

  const workingArea = getWorkingArea(CHORDS_SHEET())

  const newValues = applyStructuralRowChanges(workingArea.getValues(), changes)
  const newWidth = newValues[0]?.length ?? 0

  const lyricsContentHeight = findLastRowWithContent(getWorkingArea(LYRICS_SHEET()).getValues())
  const targetContentHeight = lyricsContentHeight * 2

  if (targetContentHeight > 0 && newWidth > 0) {
    workingArea
      .resizeTo(newWidth, targetContentHeight)
      .setValues(newValues.slice(0, targetContentHeight))
  }
}

function enforceChordWidth(): void {
  const chordsSheet = CHORDS_SHEET()
  const lyricsWorkingAreaWidth = getWorkingArea(LYRICS_SHEET()).getNumColumns()

  const frozenColumns = chordsSheet.getFrozenColumns()
  const headerWidth = SPREADSHEET().getRangeByName(CHORDS_HEADER_RANGE_NAME)?.getNumColumns() ?? frozenColumns
  const targetMaxColumns = Math.max(frozenColumns + lyricsWorkingAreaWidth, headerWidth)
  const currentMaxColumns = chordsSheet.getMaxColumns()

  if (currentMaxColumns > targetMaxColumns) {
    chordsSheet.deleteColumns(targetMaxColumns + 1, currentMaxColumns - targetMaxColumns)
  }
}

function enforceChordHeight(): void {
  const chordsSheet = CHORDS_SHEET()
  const contentHeight = findLastRowWithContent(getWorkingArea(LYRICS_SHEET()).getValues())

  const targetMaxRows = chordsSheet.getFrozenRows() + Math.max(contentHeight * 2, 1)
  const currentMaxRows = chordsSheet.getMaxRows()

  if (currentMaxRows > targetMaxRows) {
    chordsSheet.deleteRows(targetMaxRows + 1, currentMaxRows - targetMaxRows)
  }
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

function findLastRowWithContent(values: unknown[][]): number {
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i].some(cell => cell !== "")) return i + 1
  }
  return 0
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