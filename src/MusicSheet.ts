import { Chord } from "./Chords"
import { $, CellValue, PrintSheet, SheetSpace, PRINT_PAGE_WIDTH, PRINT_PAGE_HEIGHT, PRINT_HEADER_HEIGHT, PRINT_FOOTER_HEIGHT } from "./Spaces"
import { Area, Point } from "./Area"

type OnEdit = GoogleAppsScript.Events.SheetsOnEdit
type OnChange = GoogleAppsScript.Events.SheetsOnChange


const VERSION = "2.0"

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
        syncLyricsToChordSheet()
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

export function onChange(event: OnChange): void {
  try {
    if (event.changeType === "INSERT_COLUMN" || event.changeType === "REMOVE_COLUMN") {
      syncStructuralColumnChanges(detectChanges(popColumnIndexes(), $.Lyrics.main.x + 1, $.Lyrics.main.end.x))
      enforceChordWidth()
      syncLyricsToChordSheet()
    }
    else if (event.changeType === "INSERT_ROW" || event.changeType === "REMOVE_ROW") {
      syncStructuralRowChanges(detectChanges(popRowIndexes(), $.Lyrics.main.y + 1, $.Lyrics.height))
      enforceChordHeight()
      syncLyricsToChordSheet()
    }
  } catch (error) {
    warn("Unexpected error in onChange hook", error instanceof Error ? error.message : undefined)
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// SYNC LYRICS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export function syncLyricsToChordSheet(changed: Area = $.Lyrics.main.area): void {
  const source = $.Lyrics.main.sub(area => area.intersect(changed))
  if (source.isEmpty) return

  const target = $.Chords.main.sub(area => source.area
    .relativeTo($.Lyrics.main)
    .scale({ y: 2 })
    .translate(area)
  )

  const targetValues = target.getValues()
  const targetWeights = target.getFontWeights()
  source.getValues().forEach((row, i) => {
    targetValues[i * 2 + 1] = [...row]
  })
  source.getFontWeights().forEach((row, i) => {
    targetWeights[i * 2 + 1] = [...row]
  })
  target.setValues(targetValues)
  target.setFontWeights(targetWeights)
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

export function resetFormatting(...spaces: SheetSpace[]): void {
  (spaces.length ? spaces : $.ALL).forEach(space => {
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

    space.format({ fontFamily: FONT_FAMILY })
    if (space !== $.Print) {
      if (space.height >= 1) {
        space.sheet.setRowHeights(1, space.height, ROW_HEIGHT)
      }
      space.format({ fontSize: FONT_SIZE })
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
      result.forEach(row => {
        while (row.length < index) row.push("")
        row.splice(index, 0, ...Array<CellValue>(span).fill(""))
      })
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
  const emptyRow = () => Array<CellValue>(rowWidth).fill("")

  for (const { position, span } of changes) {
    const index = (position - 1) * 2
    const scaledSpan = span * 2
    if (scaledSpan > 0) {
      while (result.length < index) result.push(emptyRow())
      result.splice(index, 0, ...Array.from({ length: scaledSpan }, emptyRow))
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
  const rowWidth = newValues[0]?.length ?? $.Chords.main.width

  while (newValues.length < targetHeight) newValues.push(Array(rowWidth).fill(""))
  const finalValues = newValues.slice(0, targetHeight)

  $.Chords.main
    .sub(area => area.resizeTo({ x: rowWidth, y: targetHeight }))
    .setValues(finalValues)
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

export type SectionBounds = { startRow: number, endRow: number, width: number }

export function regeneratePrint(): void {
  const horizontalContentMargin = $.ControlPanel.contentMarginH.getValue()
  const verticalContentMargin = $.ControlPanel.contentMarginV.getValue()
  const horizontalPadding = $.ControlPanel.horizontalPadding.getValue()
  const verticalPadding = $.ControlPanel.verticalPadding.getValue()

  const availableWidthPerPage = PRINT_PAGE_WIDTH - 2 * horizontalContentMargin
  const availableHeightPerPage = PRINT_PAGE_HEIGHT - PRINT_FOOTER_HEIGHT - verticalContentMargin

  const source = $.Chords.main.getValues()

  const sections = detectSectionAreas(source)
  if (sections.some(s => s.height > availableHeightPerPage - PRINT_HEADER_HEIGHT || s.width > availableWidthPerPage)) {
    error("Sección demasiado grande", "Una sección excede los límites de la página")
    return
  }

  const positions = calculatePositions(sections, {
    fullPageWidth: PRINT_PAGE_WIDTH,
    firstPageHeaderHeight: PRINT_HEADER_HEIGHT,
    availableWidthPerPage,
    availableHeightPerPage,
    horizontalContentMargin,
    verticalContentMargin,
    horizontalPadding,
    verticalPadding,
  })

  const totalPages = positions.length > 0 ? Math.floor(positions[positions.length - 1].x / PRINT_PAGE_WIDTH) + 1 : 1
  const totalColumns = totalPages * PRINT_PAGE_WIDTH
  const target: string[][] = Array.from({ length: PRINT_PAGE_HEIGHT }, () => Array(totalColumns).fill(""))

  sections.forEach((section, i) => {
    const position = positions[i]
    for (let y = 0; y < section.height; y++) {
      for (let x = 0; x < section.width; x++) {
        target[position.y + y][position.x + x] = String(source[section.y + y]?.[section.x + x] ?? "")
      }
    }
  })

  resetPrintSheet(PRINT_PAGE_HEIGHT, totalColumns)
  $.Print.main.setValues(target)

  generatePrintHeader()

  for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
    $.Print.pageContent(pageIndex).format({ border: true })
    generatePrintFooter(pageIndex, totalPages)
  }

  resetFormatting($.Print)
}

function resetPrintSheet(rows: number, columns: number): void {
  const sheet = $.Print.sheet
  if ($.Print.height > 1) sheet.deleteRows(2, $.Print.height - 1)
  if ($.Print.width > 1) sheet.deleteColumns(2, $.Print.width - 1)
  sheet.clear()
  sheet.insertRows(1, rows - 1)
  sheet.insertColumns(1, columns - 1)
  Object.defineProperty($, "Print", { value: new PrintSheet(), writable: false, configurable: true })
}

function generatePrintHeader(): void {
  const author = $.ControlPanel.author.getValue()
  const key = $.Chords.key.getValue()
  const tempo = $.Chords.tempo.getValue()
  const notes = $.Chords.notes.getValue()

  $.Print.title.setValue($.SPREADSHEET.getName())
  $.Print.title.format({ fontSize: 14, fontWeight: "bold", verticalAlignment: "bottom" })

  $.Print.author.setValue(author)
  $.Print.author.format({ fontSize: 8 })

  const timestamp = `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}:${VERSION}`
  $.Print.timestamp.setValue(timestamp)
  $.Print.timestamp.format({ horizontalAlignment: "right", fontSize: 8 })

  const keyTempo = [key ? `[${key}]` : "", tempo ? `${tempo} bpm` : ""].filter(Boolean).join(" ")
  const prefix = keyTempo && notes ? `${keyTempo} | ` : keyTempo
  const fullText = prefix + notes
  if (notes) {
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(fullText)
      .setTextStyle(prefix.length, fullText.length, SpreadsheetApp.newTextStyle().setItalic(true).build())
      .build()
    $.Print.info.setValue(fullText)
    $.Print.info.setRichTextValue(richText)
  } else {
    $.Print.info.setValue(fullText)
  }
  $.Print.info.format({ fontSize: 9, verticalAlignment: "top", wrapStrategy: SpreadsheetApp.WrapStrategy.CLIP })

  $.Print.header.format({ border: true })
}

function generatePrintFooter(pageIndex: number, totalPages: number): void {
  const footer = $.Print.pageFooter(pageIndex)
  footer.setValue(`${$.SPREADSHEET.getName()}  ${pageIndex + 1}/${totalPages}`)
  footer.format({ horizontalAlignment: "right", fontSize: 8 })
  $.Print.sheet.setRowHeight(PRINT_PAGE_HEIGHT, 13)
}

function buildColumn(sections: Area[], maxHeight: number, vPadding: number): Area[] {
  const column: Area[] = []
  let usedHeight = 0

  for (const section of sections) {
    const needed = column.length > 0 ? usedHeight + vPadding + section.height : section.height
    if (needed > maxHeight && column.length > 0) break
    column.push(section)
    usedHeight = needed
  }

  return column
}

export type LayoutConfig = {
  fullPageWidth: number
  availableWidthPerPage: number
  availableHeightPerPage: number
  firstPageHeaderHeight: number
  horizontalContentMargin: number
  verticalContentMargin: number
  horizontalPadding: number
  verticalPadding: number
}

export function calculatePositions(sections: Area[], config: LayoutConfig): Point[] {
  const {
    fullPageWidth,
    availableWidthPerPage: availableWidth,
    availableHeightPerPage: availableHeight,
    firstPageHeaderHeight,
    horizontalContentMargin: contentMarginH,
    verticalContentMargin: contentMarginV,
    horizontalPadding,
    verticalPadding,
  } = config

  const positions: Point[] = []
  let remaining = [...sections]
  let sectionIndex = 0
  let pageIndex = 0

  while (remaining.length > 0) {
    const isFirstPage = pageIndex === 0
    const pageAvailableHeight = isFirstPage ? availableHeight - firstPageHeaderHeight : availableHeight
    const pageStartY = isFirstPage ? firstPageHeaderHeight + contentMarginV : contentMarginV

    const pageColumns: Area[][] = []
    let usedWidth = 0

    while (remaining.length > 0) {
      const column = buildColumn(remaining, pageAvailableHeight, verticalPadding)
      const columnWidth = Math.max(...column.map(s => s.width))
      const fitsWidth = usedWidth === 0 || usedWidth + horizontalPadding + columnWidth <= availableWidth

      if (!fitsWidth) break

      pageColumns.push(column)
      remaining = remaining.slice(column.length)
      usedWidth = usedWidth === 0 ? columnWidth : usedWidth + horizontalPadding + columnWidth
    }

    const columnWidths = pageColumns.map(col => Math.max(...col.map(s => s.width)))
    const totalColumnsWidth = columnWidths.reduce((sum, w) => sum + w, 0)
    const gaps = pageColumns.length - 1
    const extraSpace = availableWidth - totalColumnsWidth - gaps * horizontalPadding
    const effectivePadding = gaps > 0 ? horizontalPadding + Math.floor(extraSpace / gaps) : 0

    let columnX = pageIndex * fullPageWidth + contentMarginH
    for (let colIdx = 0; colIdx < pageColumns.length; colIdx++) {
      let sectionY = pageStartY
      for (const section of pageColumns[colIdx]) {
        positions[sectionIndex++] = new Point(columnX, sectionY)
        sectionY += section.height + verticalPadding
      }
      columnX += columnWidths[colIdx] + effectivePadding
    }

    pageIndex++
  }

  return positions
}

export function splitIntoSections(values: CellValue[][]): SectionBounds[] {
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

export function calculateSectionWidth(sectionData: CellValue[][]): number {
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

export function detectSectionAreas(values: CellValue[][]): Area[] {
  if (values.length === 0 || values[0].length === 0) return []

  const sectionColumns = splitIntoSectionColumns(values)

  let columnOffset = 0
  return sectionColumns.flatMap(sectionColumn => {
    const columnWidth = sectionColumn[0].length
    const sections = splitIntoSections(sectionColumn).map(({ startRow, endRow, width }) =>
      new Area(columnOffset, startRow, width, endRow - startRow)
    )
    columnOffset += columnWidth
    return sections
  })
}

export function splitIntoSectionColumns(values: CellValue[][]): CellValue[][][] {
  const columns: CellValue[][][] = []
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


// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

const toast = (icon: string) => (title = "", message = "") => {
  $.SPREADSHEET.toast(message, icon + " " + title, 10)
}

const info = toast("ℹ️")
const success = toast("✅")
const warn = toast("⚠️")
const error = toast("❌")