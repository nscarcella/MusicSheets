import { describe, it, expect } from 'vitest'
import { modulo, count } from '../src/Utils'

describe('modulo', () => {
  it('should handle positive values', () => {
    expect(modulo(5, 12)).toBe(5)
    expect(modulo(13, 12)).toBe(1)
    expect(modulo(24, 12)).toBe(0)
  })

  it('should handle negative values correctly', () => {
    expect(modulo(-1, 12)).toBe(11)
    expect(modulo(-13, 12)).toBe(11)
    expect(modulo(-24, 12)).toBe(0)
  })

  it('should handle zero', () => {
    expect(modulo(0, 12)).toBe(0)
  })

  it('should work with different divisors', () => {
    expect(modulo(7, 5)).toBe(2)
    expect(modulo(-3, 5)).toBe(2)
    expect(modulo(10, 3)).toBe(1)
  })
})

describe('count', () => {
  it('should count occurrences of an element', () => {
    expect(count([1, 2, 3, 2, 4, 2], 2)).toBe(3)
    expect(count([1, 2, 3, 4], 5)).toBe(0)
    expect(count(['a', 'b', 'a', 'c'], 'a')).toBe(2)
  })

  it('should handle empty arrays', () => {
    expect(count([], 1)).toBe(0)
  })

  it('should handle single element arrays', () => {
    expect(count([5], 5)).toBe(1)
    expect(count([5], 3)).toBe(0)
  })

  it('should work with strings', () => {
    const chars = ['#', 'b', '#', '#', 'b']
    expect(count(chars, '#')).toBe(3)
    expect(count(chars, 'b')).toBe(2)
  })
})
