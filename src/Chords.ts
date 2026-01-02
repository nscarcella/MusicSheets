const CHROMATIC_PITCH_CLASSES = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]

const CHORD_REGEX = /^([A-G])([#b]*)([^/]*)(?:\/([A-G])([#b]*))?$/


interface Chord {
  rootPitch: number
  suffix: string
  bassPitch: number | null
}

// -----------------------------------------------------------------------------
// CHORD MANIPULATION
// -----------------------------------------------------------------------------

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

export function parseChord(chord: string): Chord | null {
  const match = chord.match(CHORD_REGEX)
  if (!match) return null

  const [, rootLetter, rootAcc, suffix, bassLetter, bassAcc] = match

  return {
    rootPitch: buildPitch(rootLetter, rootAcc),
    bassPitch: bassLetter ? buildPitch(bassLetter, bassAcc) : null,
    suffix
  }
}


function transposePitch(pitch: number, semitones: number): number {
  const totalPitchClasses = CHROMATIC_PITCH_CLASSES.length
  return ((pitch + semitones) % totalPitchClasses + totalPitchClasses) % totalPitchClasses
}


function buildPitch(letter: string, acc: string): number {
  const offset = [...acc].reduce((sum, c) => sum + (c === "#" ? 1 : -1), 0)
  return CHROMATIC_PITCH_CLASSES.indexOf(letter) + offset
}


function buildChord(rootPitch: number, suffix: string, bassPitch: number | null): string {
  let result = CHROMATIC_PITCH_CLASSES[rootPitch] + suffix
  if (bassPitch !== null) result += "/" + CHROMATIC_PITCH_CLASSES[bassPitch]
  return result
}
