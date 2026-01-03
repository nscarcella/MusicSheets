import { describe, it, expect } from 'vitest'
import { Chord } from '../src/Chords'

describe('Chord.transpose', () => {
  it('should transpose a simple chord up by 1 semitone', () => {
    const chord = Chord.parse('C')!
    expect(chord.transpose(1).toString()).toBe('C#')
  })

  it('should transpose a simple chord down by 1 semitone', () => {
    const chord = Chord.parse('C#')!
    expect(chord.transpose(-1).toString()).toBe('C')
  })

  it('should handle wrapping from B to C', () => {
    const chord = Chord.parse('B')!
    expect(chord.transpose(1).toString()).toBe('C')
  })

  it('should handle wrapping from C to B', () => {
    const chord = Chord.parse('C')!
    expect(chord.transpose(-1).toString()).toBe('B')
  })

  it('should transpose complex chords with suffixes', () => {
    expect(Chord.parse('Cmaj7')!.transpose(2).toString()).toBe('Dmaj7')
    expect(Chord.parse('Am7')!.transpose(5).toString()).toBe('Dm7')
  })

  it('should transpose slash chords', () => {
    expect(Chord.parse('C/G')!.transpose(2).toString()).toBe('D/A')
    expect(Chord.parse('Am/E')!.transpose(-3).toString()).toBe('F#m/C#')
  })

  it('should handle chords with flats', () => {
    expect(Chord.parse('Bb')!.transpose(2).toString()).toBe('C')
    expect(Chord.parse('Eb')!.transpose(-1).toString()).toBe('D')
  })

  it('should throw error for non-integer semitones', () => {
    const chord = Chord.parse('C')!
    expect(() => chord.transpose(1.5)).toThrow('semitones must be an integer')
  })

  it('should handle chords with multiple accidentals', () => {
    expect(Chord.parse('C####')!.transpose(1).toString()).toBe('F')
    expect(Chord.parse('Cbbb')!.transpose(2).toString()).toBe('B')
  })

  it('should transpose and render with preferFlats', () => {
    expect(Chord.parse('C')!.transpose(1).toString()).toBe('C#')
    expect(Chord.parse('C')!.transpose(1).toString(false)).toBe('C#')
    expect(Chord.parse('C')!.transpose(1).toString(true)).toBe('Db')
    expect(Chord.parse('A')!.transpose(1).toString(true)).toBe('Bb')
  })
})

describe('Chord.semitoneDistance', () => {
  it('should calculate distance between two chords', () => {
    expect(Chord.parse('C')!.semitonesTo(Chord.parse('D')!)).toBe(2)
    expect(Chord.parse('C')!.semitonesTo(Chord.parse('G')!)).toBe(7)
  })

  it('should calculate negative distance', () => {
    expect(Chord.parse('D')!.semitonesTo(Chord.parse('C')!)).toBe(-2)
    expect(Chord.parse('G')!.semitonesTo(Chord.parse('C')!)).toBe(-7)
  })

  it('should calculate distance wrapping around', () => {
    expect(Chord.parse('B')!.semitonesTo(Chord.parse('C')!)).toBe(-11)
    expect(Chord.parse('C')!.semitonesTo(Chord.parse('B')!)).toBe(11)
  })

  it('should return 0 for same chord', () => {
    expect(Chord.parse('C')!.semitonesTo(Chord.parse('C')!)).toBe(0)
  })

  it('should work with complex chords', () => {
    expect(Chord.parse('Cmaj7')!.semitonesTo(Chord.parse('Dmaj7')!)).toBe(2)
  })

  it('should handle chords with multiple accidentals', () => {
    expect(Chord.parse('C####')!.semitonesTo(Chord.parse('F')!)).toBe(1)
    expect(Chord.parse('Cbbb')!.semitonesTo(Chord.parse('A')!)).toBe(0)
  })
})

describe('Chord.parse', () => {
  it('should parse simple chords', () => {
    const chord = Chord.parse('C')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('C')
  })

  it('should parse chords with sharps', () => {
    const chord = Chord.parse('C#')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('C#')
  })

  it('should parse chords with flats', () => {
    const chord = Chord.parse('Db')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('C#')
    expect(chord!.toString(true)).toBe('Db')
  })

  it('should parse chords with suffixes', () => {
    const chord = Chord.parse('Cmaj7')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('Cmaj7')
  })

  it('should parse slash chords', () => {
    const chord = Chord.parse('C/G')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('C/G')
  })

  it('should return undefined for invalid chords', () => {
    expect(Chord.parse('H')).toBeUndefined()
    expect(Chord.parse('1')).toBeUndefined()
    expect(Chord.parse('')).toBeUndefined()
  })

  it('should handle multiple sharps with wrapping', () => {
    const chord = Chord.parse('C####')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('E')
  })

  it('should handle multiple flats with wrapping', () => {
    const chord = Chord.parse('Cbbb')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('A')
    expect(chord!.toString(true)).toBe('A')
  })

  it('should handle extreme accidentals', () => {
    const chord = Chord.parse('G#####')
    expect(chord).toBeDefined()
    expect(chord!.toString()).toBe('C')
  })

  it('should render with preferFlats parameter', () => {
    const chord = Chord.parse('C#')!
    expect(chord.toString()).toBe('C#')
    expect(chord.toString(false)).toBe('C#')
    expect(chord.toString(true)).toBe('Db')
  })
})
