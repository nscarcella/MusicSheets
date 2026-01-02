const LETTER_PITCH: Record<string, number> = {
  C: 0,
  D: 2,
  E: 4,
  F: 5,
  G: 7,
  A: 9,
  B: 11
}

const CANONICAL_NOTES = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]

const CHORD_REGEX = /^([A-G])([#b]*)([^/]*)(?:\/([A-G])([#b]*))?$/

// -----------------------------------------------------------------------------
// CHORD MANIPULATION
// -----------------------------------------------------------------------------

interface ParsedChord {
  rootPitch: number
  bassPitch: number | null
  suffix: string
}

export function transpose(chord: string, semitones: number): string {
  if (!Number.isInteger(semitones)) throw new Error("Invalid argument: semitones must be an integer")

  const parsed = parseChord(chord)
  if (!parsed) throw new Error(`Chord "${chord}" is not syntactically valid`)

  return buildChord(
    transposePitch(parsed.rootPitch, semitones),
    parsed.suffix,
    parsed.bassPitch !== null ? transposePitch(parsed.bassPitch, semitones) : null
  )
}

export function semitoneDistance(fromChord: string, toChord: string): number | undefined {
  const fromParsed = parseChord(fromChord)
  if (!fromParsed) return undefined

  const toParsed = parseChord(toChord)
  if (!toParsed) return undefined

  return (toParsed.rootPitch - fromParsed.rootPitch + 12) % 12
}

// -----------------------------------------------------------------------------
// UTILS
// -----------------------------------------------------------------------------

export function parseChord(chord: string): ParsedChord | null {
  const match = chord.match(CHORD_REGEX)
  if (!match) return null

  const [, rootLetter, rootAcc, suffix, bassLetter, bassAcc] = match

  return {
    rootPitch: letterPitch(rootLetter, rootAcc),
    bassPitch: bassLetter ? letterPitch(bassLetter, bassAcc) : null,
    suffix
  }
}

function letterPitch(letter: string, acc: string): number {
  return LETTER_PITCH[letter] + accidentalOffset(acc)
}

function accidentalOffset(acc: string): number {
  let sum = 0
  for (const c of acc) sum += c === "#" ? 1 : -1
  return sum
}

function transposePitch(pitch: number, semitones: number): number {
  return ((pitch + semitones) % 12 + 12) % 12
}

function buildChord(rootPitch: number, suffix: string, bassPitch: number | null): string {
  let result = CANONICAL_NOTES[rootPitch] + suffix
  if (bassPitch !== null) result += "/" + CANONICAL_NOTES[bassPitch]
  return result
}
