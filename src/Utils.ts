export function modulo(value: number, divisor: number): number {
  return ((value % divisor) + divisor) % divisor
}

export function count<T>(iterable: Iterable<T>, element: T): number {
  let count = 0
  for (const item of iterable) {
    if (item === element) count++
  }
  return count
}

export function compose<T>(...fns: Array<(value: T) => T>): (value: T) => T {
  return (value: T) => fns.reduceRight((acc, fn) => fn(acc), value)
}
