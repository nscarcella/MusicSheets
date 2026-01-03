import { count, modulo } from "./Utils"

const { fromEntries } = Object


const CHORD_REGEX = /^([A-G])([#b]*)([^/]*)(?:\/([A-G])([#b]*))?$/

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// CHORDS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export class Chord {

  private constructor(
    private readonly pitch: Pitch,
    private readonly suffix: string,
    private readonly bass?: Pitch
  ) { }


  static parse(str: string): Chord | undefined {
    const [, pitchClass = '', accidentals = '', suffix = '', bassPitchClass = '', bassAccidentals = ''] = str.match(CHORD_REGEX) ?? []
    const pitch = Pitch.parse(pitchClass, accidentals)
    const bassPitch = Pitch.parse(bassPitchClass, bassAccidentals)

    return pitch && new Chord(pitch, suffix, bassPitch)
  }


  transpose(semitones: number): Chord {
    return new Chord(this.pitch.transpose(semitones), this.suffix, this.bass?.transpose(semitones))
  }

  semitonesTo(other: Chord): number {
    return this.pitch.semitonesTo(other.pitch)
  }

  toString(preferFlats: boolean = false): string {
    const root = this.pitch.toString(preferFlats)
    const bass = this.bass ? `/${this.bass.toString(preferFlats)}` : ""
    return `${root}${this.suffix}${bass}`
  }
}

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// PITCH
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export class Pitch {

  private constructor(
    private readonly sharpLabel: string,
    private readonly flatLabel: string = sharpLabel
  ) { }


  private static readonly INSTANCES: readonly Pitch[] = [
    new Pitch("C"),
    new Pitch("C#", "Db"),
    new Pitch("D"),
    new Pitch("D#", "Eb"),
    new Pitch("E"),
    new Pitch("F"),
    new Pitch("F#", "Gb"),
    new Pitch("G"),
    new Pitch("G#", "Ab"),
    new Pitch("A"),
    new Pitch("A#", "Bb"),
    new Pitch("B"),
  ]

  private static readonly BY_NAME: Readonly<Record<string, Pitch | undefined>> = fromEntries(
    Pitch.INSTANCES.flatMap(pitch => [
      [pitch.sharpLabel, pitch],
      [pitch.flatLabel, pitch]
    ])
  )

  static fromValue(value: number): Pitch {
    return Pitch.INSTANCES[modulo(value, Pitch.INSTANCES.length)]
  }

  static parse(diatonalClass: string, accidentals: string): Pitch | undefined {
    return Pitch.BY_NAME[diatonalClass]?.transpose(count(accidentals, '#') - count(accidentals, 'b'))
  }


  private get index(): number {
    return Pitch.INSTANCES.indexOf(this)
  }

  transpose(semitones: number): Pitch {
    if (!Number.isInteger(semitones)) throw new Error("Invalid argument: semitones must be an integer")

    return Pitch.fromValue(this.index + semitones)
  }

  semitonesTo(other: Pitch): number {
    return other.index - this.index
  }

  toString(preferFlats: boolean = false): string {
    return preferFlats ? this.flatLabel : this.sharpLabel
  }
}