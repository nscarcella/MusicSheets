import { describe, it, expect } from "vitest"
import { modulo, count } from "../src/Utils"

describe("modulo", () => {
  it("should handle positive values", () => {
    expect(modulo(5, 12)).toBe(5)
    expect(modulo(13, 12)).toBe(1)
    expect(modulo(24, 12)).toBe(0)
  })

  it("should handle negative values correctly", () => {
    expect(modulo(-1, 12)).toBe(11)
    expect(modulo(-13, 12)).toBe(11)
    expect(modulo(-24, 12)).toBe(0)
  })

  it("should handle zero", () => {
    expect(modulo(0, 12)).toBe(0)
  })

  it("should work with different divisors", () => {
    expect(modulo(7, 5)).toBe(2)
    expect(modulo(-3, 5)).toBe(2)
    expect(modulo(10, 3)).toBe(1)
  })
})

describe("count", () => {
  it("should count occurrences of an element in arrays", () => {
    expect(count([1, 2, 3, 2, 4, 2], 2)).toBe(3)
    expect(count([1, 2, 3, 4], 5)).toBe(0)
    expect(count(["a", "b", "a", "c"], "a")).toBe(2)
  })

  it("should count occurrences in strings", () => {
    expect(count("hello", "l")).toBe(2)
    expect(count("mississippi", "s")).toBe(4)
    expect(count("test", "x")).toBe(0)
  })

  it("should handle empty iterables", () => {
    expect(count([], 1)).toBe(0)
    expect(count("", "a")).toBe(0)
  })

  it("should handle single element", () => {
    expect(count([5], 5)).toBe(1)
    expect(count([5], 3)).toBe(0)
    expect(count("a", "a")).toBe(1)
  })

  it("should work with accidentals in music notation", () => {
    expect(count("###b", "#")).toBe(3)
    expect(count("###b", "b")).toBe(1)
    expect(count("bbb", "b")).toBe(3)
  })
})
