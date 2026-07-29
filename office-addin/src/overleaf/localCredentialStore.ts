const STORAGE_PREFIX = "slide2pdf:overleaf-token:v1";

export class LocalCredentialStore {
  constructor(
    private readonly storage: Storage,
    private readonly partitionKey?: string,
  ) {}

  get(endpoint: string): string | null {
    try {
      return this.storage.getItem(this.key(endpoint));
    } catch {
      return null;
    }
  }

  save(endpoint: string, token: string): boolean {
    try {
      this.storage.setItem(this.key(endpoint), token.trim());
      return true;
    } catch {
      return false;
    }
  }

  remove(endpoint: string): boolean {
    try {
      this.storage.removeItem(this.key(endpoint));
      return true;
    } catch {
      return false;
    }
  }

  private key(endpoint: string): string {
    const origin = new URL(endpoint).origin;
    const partition = this.partitionKey
      ? `:${encodeURIComponent(this.partitionKey)}`
      : "";
    return `${STORAGE_PREFIX}${partition}:${encodeURIComponent(origin)}`;
  }
}
