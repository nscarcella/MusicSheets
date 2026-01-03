import { describe, it, expect } from "vitest"
import { Chord, Pitch } from "../src/Chords"

describe("Chord", () => {
  describe("parse", () => {
    it("should parse simple chords", () => {
      const chord = Chord.parse("C")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("C")
    })

    it("should parse chords with sharps", () => {
      const chord = Chord.parse("C#")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("C#")
    })

    it("should parse chords with flats", () => {
      const chord = Chord.parse("Db")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("C#")
      expect(chord!.toString(true)).toBe("Db")
    })

    it("should parse chords with suffixes", () => {
      const chord = Chord.parse("Cmaj7")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("Cmaj7")
    })

    it("should parse slash chords", () => {
      const chord = Chord.parse("C/G")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("C/G")
    })

    it("should return undefined for invalid chords", () => {
      expect(Chord.parse("H")).toBeUndefined()
      expect(Chord.parse("1")).toBeUndefined()
      expect(Chord.parse("")).toBeUndefined()
    })

    it("should handle multiple sharps with wrapping", () => {
      const chord = Chord.parse("C####")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("E")
    })

    it("should handle multiple flats with wrapping", () => {
      const chord = Chord.parse("Cbbb")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("A")
      expect(chord!.toString(true)).toBe("A")
    })

    it("should handle extreme accidentals", () => {
      const chord = Chord.parse("G#####")
      expect(chord).toBeDefined()
      expect(chord!.toString()).toBe("C")
    })
  })

  describe("transpose", () => {
    it("should transpose a simple chord up by 1 semitone", () => {
      const chord = Chord.parse("C")!
      expect(chord.transpose(1).toString()).toBe("C#")
    })

    it("should transpose a simple chord down by 1 semitone", () => {
      const chord = Chord.parse("C#")!
      expect(chord.transpose(-1).toString()).toBe("C")
    })

    it("should handle wrapping from B to C", () => {
      const chord = Chord.parse("B")!
      expect(chord.transpose(1).toString()).toBe("C")
    })

    it("should handle wrapping from C to B", () => {
      const chord = Chord.parse("C")!
      expect(chord.transpose(-1).toString()).toBe("B")
    })

    it("should transpose complex chords with suffixes", () => {
      expect(Chord.parse("Cmaj7")!.transpose(2).toString()).toBe("Dmaj7")
      expect(Chord.parse("Am7")!.transpose(5).toString()).toBe("Dm7")
    })

    it("should transpose slash chords", () => {
      expect(Chord.parse("C/G")!.transpose(2).toString()).toBe("D/A")
      expect(Chord.parse("Am/E")!.transpose(-3).toString()).toBe("F#m/C#")
    })

    it("should handle chords with flats", () => {
      expect(Chord.parse("Bb")!.transpose(2).toString()).toBe("C")
      expect(Chord.parse("Eb")!.transpose(-1).toString()).toBe("D")
    })

    it("should throw error for non-integer semitones", () => {
      const chord = Chord.parse("C")!
      expect(() => chord.transpose(1.5)).toThrow("semitones must be an integer")
    })

    it("should handle chords with multiple accidentals", () => {
      expect(Chord.parse("C####")!.transpose(1).toString()).toBe("F")
      expect(Chord.parse("Cbbb")!.transpose(2).toString()).toBe("B")
    })
  })

  describe("semitonesTo", () => {
    it("should calculate distance between two chords", () => {
      expect(Chord.parse("C")!.semitonesTo(Chord.parse("D")!)).toBe(2)
      expect(Chord.parse("C")!.semitonesTo(Chord.parse("G")!)).toBe(7)
    })

    it("should calculate negative distance", () => {
      expect(Chord.parse("D")!.semitonesTo(Chord.parse("C")!)).toBe(-2)
      expect(Chord.parse("G")!.semitonesTo(Chord.parse("C")!)).toBe(-7)
    })

    it("should calculate distance wrapping around", () => {
      expect(Chord.parse("B")!.semitonesTo(Chord.parse("C")!)).toBe(-11)
      expect(Chord.parse("C")!.semitonesTo(Chord.parse("B")!)).toBe(11)
    })

    it("should return 0 for same chord", () => {
      expect(Chord.parse("C")!.semitonesTo(Chord.parse("C")!)).toBe(0)
    })

    it("should work with complex chords", () => {
      expect(Chord.parse("Cmaj7")!.semitonesTo(Chord.parse("Dmaj7")!)).toBe(2)
    })

    it("should handle chords with multiple accidentals", () => {
      expect(Chord.parse("C####")!.semitonesTo(Chord.parse("F")!)).toBe(1)
      expect(Chord.parse("Cbbb")!.semitonesTo(Chord.parse("A")!)).toBe(0)
    })
  })

  describe("toString", () => {
    it("should render with default (sharps)", () => {
      expect(Chord.parse("C#")!.toString()).toBe("C#")
      expect(Chord.parse("C")!.transpose(1).toString()).toBe("C#")
    })

    it("should render with preferFlats parameter", () => {
      const chord = Chord.parse("C#")!
      expect(chord.toString(false)).toBe("C#")
      expect(chord.toString(true)).toBe("Db")
    })

    it("should render complex chords with preferFlats", () => {
      expect(Chord.parse("C")!.transpose(1).toString(true)).toBe("Db")
      expect(Chord.parse("A")!.transpose(1).toString(true)).toBe("Bb")
    })

    it("should preserve suffix in output", () => {
      expect(Chord.parse("Cmaj7")!.toString()).toBe("Cmaj7")
      expect(Chord.parse("Am7")!.toString()).toBe("Am7")
    })

    it("should render slash chords correctly", () => {
      expect(Chord.parse("C/G")!.toString()).toBe("C/G")
      expect(Chord.parse("C/G")!.toString(true)).toBe("C/G")
    })
  })
})

describe("Pitch", () => {
  describe("parse", () => {
    it("should parse simple pitch", () => {
      const pitch = Pitch.parse("C")
      expect(pitch).toBeDefined()
      expect(pitch!.toString()).toBe("C")
    })

    it("should parse pitch with sharp", () => {
      const pitch = Pitch.parse("C#")
      expect(pitch).toBeDefined()
      expect(pitch!.toString()).toBe("C#")
    })

    it("should parse pitch with flat", () => {
      const pitch = Pitch.parse("Db")
      expect(pitch).toBeDefined()
      expect(pitch!.toString()).toBe("C#")
      expect(pitch!.toString(true)).toBe("Db")
    })

    it("should parse pitch with multiple accidentals", () => {
      expect(Pitch.parse("C##")!.toString()).toBe("D")
      expect(Pitch.parse("Dbb")!.toString()).toBe("C")
    })

    it("should handle wrapping with extreme accidentals", () => {
      expect(Pitch.parse("C####")!.toString()).toBe("E")
      expect(Pitch.parse("Cbbb")!.toString()).toBe("A")
    })

    it("should return undefined for invalid pitch", () => {
      expect(Pitch.parse("H")).toBeUndefined()
      expect(Pitch.parse("")).toBeUndefined()
    })

    it("should return undefined for non-letter first character", () => {
      expect(Pitch.parse("1")).toBeUndefined()
      expect(Pitch.parse("#")).toBeUndefined()
    })
  })

  describe("fromValue", () => {
    it("should create pitch from semitone value", () => {
      expect(Pitch.fromValue(0).toString()).toBe("C")
      expect(Pitch.fromValue(1).toString()).toBe("C#")
      expect(Pitch.fromValue(11).toString()).toBe("B")
    })

    it("should wrap values above 11", () => {
      expect(Pitch.fromValue(12).toString()).toBe("C")
      expect(Pitch.fromValue(13).toString()).toBe("C#")
      expect(Pitch.fromValue(24).toString()).toBe("C")
    })

    it("should handle negative values with wrapping", () => {
      expect(Pitch.fromValue(-1).toString()).toBe("B")
      expect(Pitch.fromValue(-12).toString()).toBe("C")
      expect(Pitch.fromValue(-13).toString()).toBe("B")
    })
  })

  describe("transpose", () => {
    it("should transpose up by semitones", () => {
      const pitch = Pitch.parse("C")!
      expect(pitch.transpose(1).toString()).toBe("C#")
      expect(pitch.transpose(7).toString()).toBe("G")
    })

    it("should transpose down by semitones", () => {
      const pitch = Pitch.parse("C")!
      expect(pitch.transpose(-1).toString()).toBe("B")
      expect(pitch.transpose(-7).toString()).toBe("F")
    })

    it("should handle wrapping", () => {
      expect(Pitch.parse("B")!.transpose(1).toString()).toBe("C")
      expect(Pitch.parse("C")!.transpose(-1).toString()).toBe("B")
    })

    it("should throw error for non-integer semitones", () => {
      const pitch = Pitch.parse("C")!
      expect(() => pitch.transpose(1.5)).toThrow("semitones must be an integer")
    })

    it("should handle large transpositions", () => {
      expect(Pitch.parse("C")!.transpose(24).toString()).toBe("C")
      expect(Pitch.parse("C")!.transpose(25).toString()).toBe("C#")
    })
  })

  describe("semitonesTo", () => {
    it("should calculate distance to another pitch", () => {
      const c = Pitch.parse("C")!
      expect(c.semitonesTo(Pitch.parse("D")!)).toBe(2)
      expect(c.semitonesTo(Pitch.parse("G")!)).toBe(7)
    })

    it("should calculate negative distance", () => {
      const d = Pitch.parse("D")!
      expect(d.semitonesTo(Pitch.parse("C")!)).toBe(-2)
    })

    it("should return 0 for same pitch", () => {
      const c = Pitch.parse("C")!
      expect(c.semitonesTo(Pitch.parse("C")!)).toBe(0)
    })

    it("should handle wrapping", () => {
      expect(Pitch.parse("B")!.semitonesTo(Pitch.parse("C")!)).toBe(-11)
      expect(Pitch.parse("C")!.semitonesTo(Pitch.parse("B")!)).toBe(11)
    })
  })

  describe("toString", () => {
    it("should return sharp notation by default", () => {
      expect(Pitch.parse("C#")!.toString()).toBe("C#")
      expect(Pitch.fromValue(1).toString()).toBe("C#")
    })

    it("should return flat notation when preferFlats is true", () => {
      expect(Pitch.parse("C#")!.toString(true)).toBe("Db")
      expect(Pitch.fromValue(1).toString(true)).toBe("Db")
    })

    it("should handle natural notes (same output for both)", () => {
      expect(Pitch.parse("C")!.toString()).toBe("C")
      expect(Pitch.parse("C")!.toString(true)).toBe("C")
    })
  })
})
