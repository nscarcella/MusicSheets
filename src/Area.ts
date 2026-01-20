import { lazy } from "./Utils"

type Range = GoogleAppsScript.Spreadsheet.Range

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// POINT
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export type PointPayload = Pick<Point, "x" | "y">

export class Point {

  constructor(
    readonly x: number,
    readonly y: number,
  ) { }


  copy(payload: Partial<PointPayload>): Point {
    return new Point(payload.x ?? this.x, payload.y ?? this.y)
  }

  tx(dx: (x: number) => number, dy: (y: number) => number): Point {
    return this.copy({ x: dx(this.x), y: dy(this.y) })
  }


  add(other: PointPayload): Point {
    return this.copy({ x: this.x + other.x, y: this.y + other.y })
  }

  sub(other: PointPayload): Point {
    return this.copy({ x: this.x - other.x, y: this.y - other.y })
  }

  scale(factor: number): Point {
    return this.copy({ x: this.x * factor, y: this.y * factor })
  }

  @lazy get neg(): Point {
    return this.scale(-1)
  }


  equals(other: Point): boolean {
    return other instanceof Point && this.x === other.x && this.y === other.y
  }

  toString(): string {
    return `(${this.x}, ${this.y})`
  }


  to(other: PointPayload): Area {
    return this.by({ x: other.x - this.x, y: other.y - this.y })
  }

  by(size: PointPayload): Area {
    return new Area(
      size.x < 0 ? this.x + size.x : this.x,
      size.y < 0 ? this.y + size.y : this.y,
      Math.abs(size.x),
      Math.abs(size.y),
    )
  }
}

export const Origin = new Point(0, 0)
export const X = (x: number) => new Point(x, 0)
export const Y = (y: number) => new Point(0, y)

// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// AREA
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

export class Area {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number

  static fromRange(range: Range): Area {
    return new Area(range.getColumn() - 1, range.getRow() - 1, range.getNumColumns(), range.getNumRows())
  }

  constructor(x: number, y: number, width: number, height: number) {
    if (width < 0) x -= width = -width
    if (height < 0) y -= height = -height

    this.x = x
    this.y = y
    this.width = width
    this.height = height
  }

  @lazy get start(): Point { return new Point(this.x, this.y) }
  @lazy get end(): Point { return this.start.add(this.size) }
  @lazy get size(): Point { return new Point(this.width, this.height) }
  @lazy get isEmpty(): boolean { return this.width === 0 || this.height === 0 }

  copy({ x = this.x, y = this.y, width = this.width, height = this.height }: Partial<Area>): Area {
    return new Area(x, y, width, height)
  }

  translate({ x = 0, y = 0 }: Partial<PointPayload>): Area {
    return this.copy({ x: this.x + x, y: this.y + y })
  }

  translateTo({ x = this.x, y = this.y }: Partial<PointPayload>): Area {
    return this.copy({ x, y })
  }

  scale({ x = 1, y = 1 }: Partial<PointPayload>): Area {
    const newX = Math.floor(this.x * x)
    const newY = Math.floor(this.y * y)
    return new Area(
      newX,
      newY,
      Math.ceil((this.x + this.width) * x) - newX,
      Math.ceil((this.y + this.height) * y) - newY,
    )
  }

  resizeBy({ x = 1, y = 1 }: Partial<PointPayload>): Area {
    return this.copy({ width: this.width * x, height: this.height * y })
  }

  resize({ x = 0, y = 0 }: Partial<PointPayload>): Area {
    return this.copy({ width: this.width + x, height: this.height + y })
  }

  resizeTo({ x: width = this.width, y: height = this.height }: Partial<PointPayload>): Area {
    return this.copy({ width, height })
  }

  columns(count: number): Area {
    return count >= 0
      ? this.copy({ width: Math.min(count, this.width) })
      : this.copy({ x: this.x + Math.max(0, this.width + count), width: Math.min(-count, this.width) })
  }

  rows(count: number): Area {
    return count >= 0
      ? this.copy({ height: Math.min(count, this.height) })
      : this.copy({ y: this.y + Math.max(0, this.height + count), height: Math.min(-count, this.height) })
  }

  relativeTo(origin: PointPayload): Area {
    return this.translate({ x: -origin.x, y: -origin.y })
  }

  overlapsWith(other: Area): boolean {
    return this.x < other.x + other.width
      && this.x + this.width > other.x
      && this.y < other.y + other.height
      && this.y + this.height > other.y
  }

  intersect(other: Area): Area {
    if (!this.overlapsWith(other)) return Empty
    const x = Math.max(this.x, other.x)
    const y = Math.max(this.y, other.y)
    return new Area(
      x, y,
      Math.min(this.x + this.width, other.x + other.width) - x,
      Math.min(this.y + this.height, other.y + other.height) - y,
    )
  }

  crop({ left = 0, top = 0, right = 0, bottom = 0 }: { left?: number, top?: number, right?: number, bottom?: number }): Area {
    return new Area(
      this.x + left,
      this.y + top,
      this.width - left - right,
      this.height - top - bottom,
    )
  }

  equals(other: Area): boolean {
    return other instanceof Area
      && this.x === other.x
      && this.y === other.y
      && this.width === other.width
      && this.height === other.height
  }

  toString(): string {
    return `Area(${this.x}, ${this.y}, ${this.width}, ${this.height})`
  }
}

export const Empty = new Area(0, 0, 0, 0)