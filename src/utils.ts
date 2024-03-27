type Entries<T> = {
  [K in keyof T]: [K, T[K]];
}[keyof T][];

export function entriesFromObject<T extends object>(object: T): Entries<T> {
  return Object.entries(object) as Entries<T>;
}