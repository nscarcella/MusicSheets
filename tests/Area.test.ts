import { describe, it, expect } from "vitest"
import { Point, Area, Origin, X, Y } from "../src/Area"

describe("Point", () => {
  const p = (x: number, y: number) => new Point(x, y)

  describe("constructor and properties", () => {
    it("should create a point with x and y", () => {
      const point = p(3, 4)
      expect(point.x).toBe(3)
      expect(point.y).toBe(4)
    })
  })

  describe("copy", () => {
    it("should copy with partial changes", () => {
      expect(p(1, 2).copy({ x: 5 })).toEqual(p(5, 2))
      expect(p(1, 2).copy({ y: 5 })).toEqual(p(1, 5))
      expect(p(1, 2).copy({ x: 3, y: 4 })).toEqual(p(3, 4))
      expect(p(1, 2).copy({})).toEqual(p(1, 2))
    })
  })

  describe("tx", () => {
    it("should transform x and y with functions", () => {
      expect(p(2, 3).tx(x => x * 2, y => y + 1)).toEqual(p(4, 4))
    })
  })

  describe("arithmetic", () => {
    it("should add points", () => {
      expect(p(1, 2).add(p(3, 4))).toEqual(p(4, 6))
      expect(p(1, 2).add({ x: 3, y: 4 })).toEqual(p(4, 6))
    })

    it("should subtract points", () => {
      expect(p(5, 7).sub(p(2, 3))).toEqual(p(3, 4))
      expect(p(5, 7).sub({ x: 2, y: 3 })).toEqual(p(3, 4))
    })

    it("should scale by factor", () => {
      expect(p(2, 3).scale(2)).toEqual(p(4, 6))
      expect(p(2, 3).scale(-1)).toEqual(p(-2, -3))
      expect(p(2, 3).scale(0)).toEqual(p(0, 0))
    })

    it("should negate", () => {
      expect(p(2, 3).neg).toEqual(p(-2, -3))
      expect(p(-1, -2).neg).toEqual(p(1, 2))
    })
  })

  describe("equals", () => {
    it("should return true for equal points", () => {
      expect(p(1, 2).equals(p(1, 2))).toBe(true)
    })

    it("should return false for different points", () => {
      expect(p(1, 2).equals(p(1, 3))).toBe(false)
      expect(p(1, 2).equals(p(2, 2))).toBe(false)
    })

    it("should return false for non-Point objects", () => {
      expect(p(0, 0).equals(new Area(0, 0, 1, 1) as unknown as Point)).toBe(false)
    })
  })

  describe("toString", () => {
    it("should return formatted string", () => {
      expect(p(3, 4).toString()).toBe("(3, 4)")
      expect(p(-1, 0).toString()).toBe("(-1, 0)")
    })
  })

  describe("to", () => {
    it("should create area from two corners", () => {
      const area = p(1, 2).to(p(4, 6))
      expect(area).toEqual(new Area(1, 2, 3, 4))
    })

    it("should handle reversed corners", () => {
      const area = p(4, 6).to(p(1, 2))
      expect(area).toEqual(new Area(1, 2, 3, 4))
    })
  })

  describe("by", () => {
    it("should create area with positive size", () => {
      const area = p(1, 2).by(p(3, 4))
      expect(area).toEqual(new Area(1, 2, 3, 4))
    })

    it("should adjust origin for negative size", () => {
      expect(p(5, 5).by(p(-3, -2))).toEqual(new Area(2, 3, 3, 2))
      expect(p(5, 5).by(p(-3, 2))).toEqual(new Area(2, 5, 3, 2))
      expect(p(5, 5).by(p(3, -2))).toEqual(new Area(5, 3, 3, 2))
    })
  })

  describe("constants", () => {
    it("Origin should be (0, 0)", () => {
      expect(Origin).toEqual(p(0, 0))
    })

    it("X should create horizontal points", () => {
      expect(X(5)).toEqual(p(5, 0))
    })

    it("Y should create vertical points", () => {
      expect(Y(5)).toEqual(p(0, 5))
    })
  })
})

describe("Area", () => {
  const a = (x: number, y: number, w: number, h: number) => new Area(x, y, w, h)
  const p = (x: number, y: number) => new Point(x, y)

  describe("constructor and properties", () => {
    it("should create area with position and size", () => {
      const area = a(1, 2, 3, 4)
      expect(area.x).toBe(1)
      expect(area.y).toBe(2)
      expect(area.width).toBe(3)
      expect(area.height).toBe(4)
    })

    it("should treat negative width as x being the right edge", () => {
      expect(a(10, 5, -3, 4)).toEqual(a(7, 5, 3, 4))
      expect(a(10, 5, -10, 4)).toEqual(a(0, 5, 10, 4))
    })

    it("should treat negative height as y being the bottom edge", () => {
      expect(a(5, 10, 3, -4)).toEqual(a(5, 6, 3, 4))
      expect(a(5, 10, 3, -10)).toEqual(a(5, 0, 3, 10))
    })

    it("should allow negative x and y", () => {
      expect(a(-5, -3, 10, 10)).toEqual(a(-5, -3, 10, 10))
      expect(a(-5, -3, 10, 10).x).toBe(-5)
      expect(a(-5, -3, 10, 10).y).toBe(-3)
    })

    it("should adjust position when using negative size", () => {
      expect(a(5, 5, -3, -2)).toEqual(a(2, 3, 3, 2))
      expect(a(2, 2, -5, -5)).toEqual(a(-3, -3, 5, 5))
    })
  })

  describe("lazy getters", () => {
    it("start should return top-left corner", () => {
      expect(a(1, 2, 3, 4).start).toEqual(p(1, 2))
    })

    it("end should return bottom-right corner", () => {
      expect(a(1, 2, 3, 4).end).toEqual(p(4, 6))
    })

    it("size should return dimensions as point", () => {
      expect(a(1, 2, 3, 4).size).toEqual(p(3, 4))
    })

    it("isEmpty should return true for zero width", () => {
      expect(a(1, 2, 0, 4).isEmpty).toBe(true)
    })

    it("isEmpty should return true for zero height", () => {
      expect(a(1, 2, 3, 0).isEmpty).toBe(true)
    })

    it("isEmpty should return false for non-empty area", () => {
      expect(a(1, 2, 3, 4).isEmpty).toBe(false)
      expect(a(0, 0, 1, 1).isEmpty).toBe(false)
    })
  })

  describe("copy", () => {
    it("should copy with partial changes", () => {
      expect(a(1, 2, 3, 4).copy({ x: 10 })).toEqual(a(10, 2, 3, 4))
      expect(a(1, 2, 3, 4).copy({ width: 10, height: 20 })).toEqual(a(1, 2, 10, 20))
      expect(a(1, 2, 3, 4).copy({})).toEqual(a(1, 2, 3, 4))
    })
  })

  describe("translate", () => {
    it("should move area by delta", () => {
      expect(a(1, 2, 3, 4).translate(p(5, 6))).toEqual(a(6, 8, 3, 4))
      expect(a(1, 2, 3, 4).translate({ x: -1, y: -2 })).toEqual(a(0, 0, 3, 4))
    })

    it("should allow translation to negative coordinates", () => {
      expect(a(1, 2, 3, 4).translate({ x: -5, y: -10 })).toEqual(a(-4, -8, 3, 4))
    })

    it("should use defaults for partial arguments", () => {
      expect(a(1, 2, 3, 4).translate({ x: 5 })).toEqual(a(6, 2, 3, 4))
      expect(a(1, 2, 3, 4).translate({ y: 5 })).toEqual(a(1, 7, 3, 4))
    })
  })

  describe("translateTo", () => {
    it("should move area to position", () => {
      expect(a(1, 2, 3, 4).translateTo(p(10, 20))).toEqual(a(10, 20, 3, 4))
    })

    it("should allow negative positions", () => {
      expect(a(1, 2, 3, 4).translateTo({ x: -5, y: -3 })).toEqual(a(-5, -3, 3, 4))
    })

    it("should use current position for partial arguments", () => {
      expect(a(1, 2, 3, 4).translateTo({ x: 10 })).toEqual(a(10, 2, 3, 4))
      expect(a(1, 2, 3, 4).translateTo({ y: 20 })).toEqual(a(1, 20, 3, 4))
    })
  })

  describe("scale", () => {
    it("should scale position and dimensions by factor", () => {
      expect(a(1, 2, 3, 4).scale(p(2, 3))).toEqual(a(2, 6, 6, 12))
      expect(a(2, 4, 6, 8).scale({ x: 0.5, y: 0.5 })).toEqual(a(1, 2, 3, 4))
    })

    it("should handle negative scale factors (flips area)", () => {
      expect(a(2, 3, 4, 5).scale({ x: -1, y: -1 })).toEqual(a(-6, -8, 4, 5))
      expect(a(1, 2, 3, 4).scale({ x: -2 })).toEqual(a(-8, 2, 6, 4))
      expect(a(1, 2, 3, 4).scale({ y: -2 })).toEqual(a(1, -12, 3, 8))
    })

    it("should handle zero scale (collapses to empty)", () => {
      expect(a(1, 2, 3, 4).scale({ x: 0 })).toEqual(a(0, 2, 0, 4))
      expect(a(1, 2, 3, 4).scale({ y: 0 })).toEqual(a(1, 0, 3, 0))
    })

    it("should use 1 as default for partial arguments", () => {
      expect(a(1, 2, 3, 4).scale({ x: 2 })).toEqual(a(2, 2, 6, 4))
      expect(a(1, 2, 3, 4).scale({ y: 3 })).toEqual(a(1, 6, 3, 12))
    })
  })

  describe("resizeBy", () => {
    it("should scale dimensions only by factor", () => {
      expect(a(1, 2, 3, 4).resizeBy(p(2, 3))).toEqual(a(1, 2, 6, 12))
      expect(a(2, 4, 6, 8).resizeBy({ x: 0.5, y: 0.5 })).toEqual(a(2, 4, 3, 4))
    })

    it("should handle negative factors (normalizes via constructor)", () => {
      expect(a(5, 5, 4, 6).resizeBy({ x: -1 })).toEqual(a(1, 5, 4, 6))
      expect(a(5, 5, 4, 6).resizeBy({ y: -1 })).toEqual(a(5, -1, 4, 6))
    })

    it("should handle zero factor (collapses to empty)", () => {
      expect(a(1, 2, 3, 4).resizeBy({ x: 0 })).toEqual(a(1, 2, 0, 4))
      expect(a(1, 2, 3, 4).resizeBy({ y: 0 })).toEqual(a(1, 2, 3, 0))
    })

    it("should use 1 as default for partial arguments", () => {
      expect(a(1, 2, 3, 4).resizeBy({ x: 2 })).toEqual(a(1, 2, 6, 4))
      expect(a(1, 2, 3, 4).resizeBy({ y: 3 })).toEqual(a(1, 2, 3, 12))
    })
  })

  describe("resize", () => {
    it("should add to dimensions", () => {
      expect(a(1, 2, 3, 4).resize(p(2, 3))).toEqual(a(1, 2, 5, 7))
      expect(a(1, 2, 3, 4).resize({ x: -1, y: -2 })).toEqual(a(1, 2, 2, 2))
    })

    it("should handle resize to negative dimensions (normalizes via constructor)", () => {
      expect(a(5, 5, 3, 4).resize({ x: -10 })).toEqual(a(-2, 5, 7, 4))
      expect(a(5, 5, 3, 4).resize({ y: -10 })).toEqual(a(5, -1, 3, 6))
    })

    it("should use 0 as default for partial arguments", () => {
      expect(a(1, 2, 3, 4).resize({ x: 2 })).toEqual(a(1, 2, 5, 4))
      expect(a(1, 2, 3, 4).resize({ y: 3 })).toEqual(a(1, 2, 3, 7))
    })
  })

  describe("resizeTo", () => {
    it("should set dimensions", () => {
      expect(a(1, 2, 3, 4).resizeTo(p(10, 20))).toEqual(a(1, 2, 10, 20))
    })

    it("should handle negative dimensions (normalizes via constructor)", () => {
      expect(a(5, 5, 3, 4).resizeTo({ x: -2 })).toEqual(a(3, 5, 2, 4))
      expect(a(5, 5, 3, 4).resizeTo({ y: -3 })).toEqual(a(5, 2, 3, 3))
    })

    it("should use current size for partial arguments", () => {
      expect(a(1, 2, 3, 4).resizeTo({ x: 10 })).toEqual(a(1, 2, 10, 4))
      expect(a(1, 2, 3, 4).resizeTo({ y: 20 })).toEqual(a(1, 2, 3, 20))
    })
  })

  describe("columns", () => {
    it("should select first N columns with positive count", () => {
      expect(a(0, 0, 10, 5).columns(3)).toEqual(a(0, 0, 3, 5))
      expect(a(5, 5, 10, 5).columns(3)).toEqual(a(5, 5, 3, 5))
    })

    it("should select last N columns with negative count", () => {
      expect(a(0, 0, 10, 5).columns(-3)).toEqual(a(7, 0, 3, 5))
      expect(a(5, 5, 10, 5).columns(-3)).toEqual(a(12, 5, 3, 5))
    })

    it("should cap at available width", () => {
      expect(a(0, 0, 5, 5).columns(10)).toEqual(a(0, 0, 5, 5))
      expect(a(0, 0, 5, 5).columns(-10)).toEqual(a(0, 0, 5, 5))
    })
  })

  describe("rows", () => {
    it("should select first N rows with positive count", () => {
      expect(a(0, 0, 5, 10).rows(3)).toEqual(a(0, 0, 5, 3))
      expect(a(5, 5, 5, 10).rows(3)).toEqual(a(5, 5, 5, 3))
    })

    it("should select last N rows with negative count", () => {
      expect(a(0, 0, 5, 10).rows(-3)).toEqual(a(0, 7, 5, 3))
      expect(a(5, 5, 5, 10).rows(-3)).toEqual(a(5, 12, 5, 3))
    })

    it("should cap at available height", () => {
      expect(a(0, 0, 5, 5).rows(10)).toEqual(a(0, 0, 5, 5))
      expect(a(0, 0, 5, 5).rows(-10)).toEqual(a(0, 0, 5, 5))
    })
  })

  describe("relativeTo", () => {
    it("should return area relative to new origin", () => {
      expect(a(0, 0, 5, 4).relativeTo(p(2, 2))).toEqual(a(-2, -2, 5, 4))
      expect(a(5, 5, 3, 3).relativeTo(p(2, 3))).toEqual(a(3, 2, 3, 3))
    })

    it("should handle origin at current position", () => {
      expect(a(5, 5, 3, 3).relativeTo(p(5, 5))).toEqual(a(0, 0, 3, 3))
    })
  })

  describe("overlapsWith", () => {
    it("should return true for overlapping areas", () => {
      expect(a(0, 0, 10, 10).overlapsWith(a(5, 5, 10, 10))).toBe(true)
      expect(a(0, 0, 10, 10).overlapsWith(a(9, 9, 10, 10))).toBe(true)
    })

    it("should return false for non-overlapping areas", () => {
      expect(a(0, 0, 10, 10).overlapsWith(a(10, 0, 10, 10))).toBe(false)
      expect(a(0, 0, 10, 10).overlapsWith(a(0, 10, 10, 10))).toBe(false)
      expect(a(0, 0, 10, 10).overlapsWith(a(20, 20, 10, 10))).toBe(false)
    })

    it("should return false for adjacent areas", () => {
      expect(a(0, 0, 5, 5).overlapsWith(a(5, 0, 5, 5))).toBe(false)
      expect(a(0, 0, 5, 5).overlapsWith(a(0, 5, 5, 5))).toBe(false)
    })
  })

  describe("intersect", () => {
    it("should return intersection area", () => {
      expect(a(0, 0, 10, 10).intersect(a(5, 5, 10, 10))).toEqual(a(5, 5, 5, 5))
      expect(a(0, 0, 10, 10).intersect(a(2, 3, 4, 5))).toEqual(a(2, 3, 4, 5))
    })

    it("should return Empty for non-overlapping areas", () => {
      expect(a(0, 0, 10, 10).intersect(a(10, 10, 10, 10))).toEqual(a(0, 0, 0, 0))
      expect(a(0, 0, 5, 5).intersect(a(10, 10, 5, 5))).toEqual(a(0, 0, 0, 0))
    })
  })

  describe("equals", () => {
    it("should return true for equal areas", () => {
      expect(a(1, 2, 3, 4).equals(a(1, 2, 3, 4))).toBe(true)
    })

    it("should return false for different areas", () => {
      expect(a(1, 2, 3, 4).equals(a(1, 2, 3, 5))).toBe(false)
      expect(a(1, 2, 3, 4).equals(a(0, 2, 3, 4))).toBe(false)
    })

    it("should return false for non-Area objects", () => {
      expect(a(1, 2, 3, 4).equals({ x: 1, y: 2, width: 3, height: 4 } as Area)).toBe(false)
    })
  })

  describe("toString", () => {
    it("should return formatted string", () => {
      expect(a(1, 2, 3, 4).toString()).toBe("Area(1, 2, 3, 4)")
    })
  })

  describe("crop", () => {
    it("should crop from left", () => {
      expect(a(0, 0, 10, 10).crop({ left: 2 })).toEqual(a(2, 0, 8, 10))
    })

    it("should crop from top", () => {
      expect(a(0, 0, 10, 10).crop({ top: 3 })).toEqual(a(0, 3, 10, 7))
    })

    it("should crop from right", () => {
      expect(a(0, 0, 10, 10).crop({ right: 2 })).toEqual(a(0, 0, 8, 10))
    })

    it("should crop from bottom", () => {
      expect(a(0, 0, 10, 10).crop({ bottom: 3 })).toEqual(a(0, 0, 10, 7))
    })

    it("should crop from multiple edges", () => {
      expect(a(0, 0, 10, 10).crop({ left: 1, top: 2, right: 3, bottom: 4 })).toEqual(a(1, 2, 6, 4))
    })

    it("should preserve position when cropping non-origin area", () => {
      expect(a(5, 5, 10, 10).crop({ left: 2, top: 3 })).toEqual(a(7, 8, 8, 7))
    })

    it("should handle empty crop", () => {
      expect(a(1, 2, 3, 4).crop({})).toEqual(a(1, 2, 3, 4))
    })
  })
})
