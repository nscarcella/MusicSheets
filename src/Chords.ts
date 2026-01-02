const CHROMATIC_PITCH_CLASSES = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]

const CHORD_REGEX = /^([A-G])([#b]*)([^/]*)(?:\/([A-G])([#b]*))?$/


type DiatonicPitchClass = "A" | "B" | "C" | "D" | "E" | "F" | "G"
type Chord = string
type Pitch = number

interface ChordInfo {
  rootPitch: Pitch
  suffix: string
  bassPitch: Pitch | undefined
}

// -----------------------------------------------------------------------------
// CHORD MANIPULATION
// -----------------------------------------------------------------------------

export function transpose(chord: Chord, semitones: number): Chord {
  if (!Number.isInteger(semitones)) throw new Error("Invalid argument: semitones must be an integer")

  const parsed = parseChord(chord)
  if (!parsed) throw new Error(`Chord "${chord}" is not syntactically valid`)

  const rootPitch = transposePitch(parsed.rootPitch, semitones)
  const bassPitch = parsed.bassPitch !== undefined ? transposePitch(parsed.bassPitch, semitones) : undefined

  let result = CHROMATIC_PITCH_CLASSES[rootPitch] + parsed.suffix
  if (bassPitch !== undefined) result += "/" + CHROMATIC_PITCH_CLASSES[bassPitch]
  return result
}


export function semitoneDistance(from: Chord, to: Chord): number | undefined {
  const fromParsed = parseChord(from)
  if (!fromParsed) return undefined

  const toParsed = parseChord(to)
  if (!toParsed) return undefined

  return (toParsed.rootPitch - fromParsed.rootPitch + CHROMATIC_PITCH_CLASSES.length) % CHROMATIC_PITCH_CLASSES.length
}

// -----------------------------------------------------------------------------
// UTILS
// -----------------------------------------------------------------------------

export function parseChord(chord: Chord): ChordInfo | undefined {
  const match = chord.match(CHORD_REGEX)
  if (!match) return undefined

  const [, diatonicPitchClass, accidentals, suffix, bassDiatonicPitchClass, bassAccidentals] = match

  return {
    rootPitch: buildPitch(diatonicPitchClass as DiatonicPitchClass, accidentals),
    bassPitch: bassDiatonicPitchClass ? buildPitch(bassDiatonicPitchClass as DiatonicPitchClass, bassAccidentals) : undefined,
    suffix
  }
}


function transposePitch(pitch: Pitch, semitones: number): Pitch {
  const totalPitchClasses = CHROMATIC_PITCH_CLASSES.length
  return ((pitch + semitones) % totalPitchClasses + totalPitchClasses) % totalPitchClasses
}


function buildPitch(letter: DiatonicPitchClass, accidentals: string): Pitch {
  const offset = [...accidentals].reduce((sum, c) => sum + (c === "#" ? 1 : -1), 0)
  return CHROMATIC_PITCH_CLASSES.indexOf(letter) + offset
}
