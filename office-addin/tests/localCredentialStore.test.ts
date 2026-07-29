import { describe, expect, it } from "vitest";

import { LocalCredentialStore } from "../src/overleaf/localCredentialStore";

class MemoryStorage implements Storage {
  private readonly values = new Map<string, string>();

  get length(): number {
    return this.values.size;
  }

  clear(): void {
    this.values.clear();
  }

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  key(index: number): string | null {
    return [...this.values.keys()][index] ?? null;
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }
}

describe("LocalCredentialStore", () => {
  it("stores one token per normalized endpoint and Office partition", () => {
    const storage = new MemoryStorage();
    const store = new LocalCredentialStore(storage, "powerpoint-partition");

    expect(store.save("https://overleaf.example/", "  olp_example  ")).toBe(true);

    expect(store.get("https://overleaf.example")).toBe("olp_example");
    expect(store.get("https://other.example")).toBeNull();
    expect(storage.key(0)).toContain("powerpoint-partition");
  });

  it("removes a remembered token", () => {
    const storage = new MemoryStorage();
    const store = new LocalCredentialStore(storage);
    store.save("https://overleaf.example", "olp_example");

    expect(store.remove("https://overleaf.example")).toBe(true);
    expect(store.get("https://overleaf.example")).toBeNull();
  });

  it("falls back without throwing when browser storage is unavailable", () => {
    const storage = {
      getItem: () => {
        throw new DOMException("blocked", "SecurityError");
      },
      removeItem: () => {
        throw new DOMException("blocked", "SecurityError");
      },
      setItem: () => {
        throw new DOMException("blocked", "SecurityError");
      },
    } as unknown as Storage;
    const store = new LocalCredentialStore(storage);

    expect(store.get("https://overleaf.example")).toBeNull();
    expect(store.save("https://overleaf.example", "olp_example")).toBe(false);
    expect(store.remove("https://overleaf.example")).toBe(false);
  });
});
