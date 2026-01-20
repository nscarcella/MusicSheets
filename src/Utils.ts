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

export function raise(error: Error): never {
  throw error
}

export function lazy<T>(_target: object, propertyKey: string | symbol, descriptor: TypedPropertyDescriptor<T>): TypedPropertyDescriptor<T> {
  const originalGetter = descriptor.get!
  descriptor.get = function (this: object): T {
    const value = originalGetter.call(this)
    Object.defineProperty(this, propertyKey, { value, writable: false, configurable: true })
    return value
  }
  return descriptor
}