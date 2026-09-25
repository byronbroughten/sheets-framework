export function lazy<T>(make: () => T): () => T {
  let ready = false;
  let value: T;
  return () => {
    if (!ready) {
      value = make();
      ready = true;
    }
    return value;
  };
}
