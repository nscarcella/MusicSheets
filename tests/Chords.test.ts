import { describe, it, expect } from 'vitest'
import { parseChord, semitoneDistance, transpose } from '../src/Chords'

describe('transpose', () => {
  it('should transpose a simple chord up by 1 semitone', () => {
    expect(transpose('C', 1)).toBe('C#')
  })

  it('should transpose a simple chord down by 1 semitone', () => {
    expect(transpose('C#', -1)).toBe('C')
  })

  it('should handle wrapping from B to C', () => {
    expect(transpose('B', 1)).toBe('C')
  })

  it('should handle wrapping from C to B', () => {
    expect(transpose('C', -1)).toBe('B')
  })

  it('should transpose complex chords with suffixes', () => {
    expect(transpose('Cmaj7', 2)).toBe('Dmaj7')
    expect(transpose('Am7', 5)).toBe('Dm7')
  })

  it('should transpose slash chords', () => {
    expect(transpose('C/G', 2)).toBe('D/A')
    expect(transpose('Am/E', -3)).toBe('F#m/C#')
  })

  it('should handle chords with flats', () => {
    expect(transpose('Bb', 2)).toBe('C')
    expect(transpose('Eb', -1)).toBe('D')
  })

  it('should throw error for invalid chord', () => {
    expect(() => transpose('H', 1)).toThrow('not syntactically valid')
  })

  it('should throw error for non-integer semitones', () => {
    expect(() => transpose('C', 1.5)).toThrow('semitones must be an integer')
  })
})

describe('semitoneDistance', () => {
  it('should calculate distance between two chords', () => {
    expect(semitoneDistance('C', 'D')).toBe(2)
    expect(semitoneDistance('C', 'G')).toBe(7)
  })

  it('should calculate distance wrapping around', () => {
    expect(semitoneDistance('B', 'C')).toBe(1)
    expect(semitoneDistance('G', 'C')).toBe(5)
  })

  it('should return 0 for same chord', () => {
    expect(semitoneDistance('C', 'C')).toBe(0)
  })

  it('should return undefined for invalid chords', () => {
    expect(semitoneDistance('H', 'C')).toBeUndefined()
    expect(semitoneDistance('C', 'H')).toBeUndefined()
  })

  it('should work with complex chords', () => {
    expect(semitoneDistance('Cmaj7', 'Dmaj7')).toBe(2)
  })
})

describe('parseChord', () => {
  it('should parse simple chords', () => {
    const result = parseChord('C')
    expect(result).toEqual({ rootPitch: 0, bassPitch: undefined, suffix: '' })
  })

  it('should parse chords with sharps', () => {
    const result = parseChord('C#')
    expect(result).toEqual({ rootPitch: 1, bassPitch: undefined, suffix: '' })
  })

  it('should parse chords with flats', () => {
    const result = parseChord('Db')
    expect(result).toEqual({ rootPitch: 1, bassPitch: undefined, suffix: '' })
  })

  it('should parse chords with suffixes', () => {
    const result = parseChord('Cmaj7')
    expect(result).toEqual({ rootPitch: 0, bassPitch: undefined, suffix: 'maj7' })
  })

  it('should parse slash chords', () => {
    const result = parseChord('C/G')
    expect(result).toEqual({ rootPitch: 0, bassPitch: 7, suffix: '' })
  })

  it('should return undefined for invalid chords', () => {
    expect(parseChord('H')).toBeUndefined()
    expect(parseChord('1')).toBeUndefined()
    expect(parseChord('')).toBeUndefined()
  })
})
