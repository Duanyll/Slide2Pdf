const SETTINGS_KEY = "slide2pdf.overleaf.v1";

export interface SavedOverleafTarget {
  remoteUrl: string;
  filePath: string;
}

interface SavedDocumentTargets {
  version: 1;
  targets: Record<string, SavedOverleafTarget>;
}

export interface DocumentSettingsAdapter {
  get(name: string): unknown;
  set(name: string, value: unknown): void;
  save(): Promise<void>;
}

export function createOfficeDocumentSettingsAdapter(
  settings: Office.Settings,
): DocumentSettingsAdapter {
  return {
    get: (name) => settings.get(name),
    set: (name, value) => settings.set(name, value),
    save: () =>
      new Promise<void>((resolve, reject) => {
        settings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(result.error.message));
            return;
          }
          resolve();
        });
      }),
  };
}

export class DocumentTargetStore {
  constructor(private readonly settings: DocumentSettingsAdapter) {}

  get(slideId: string): SavedOverleafTarget | null {
    return this.read().targets[slideId] ?? null;
  }

  async save(slideId: string, target: SavedOverleafTarget): Promise<void> {
    const documentTargets = this.read();
    documentTargets.targets[slideId] = { ...target };
    this.settings.set(SETTINGS_KEY, documentTargets);
    await this.settings.save();
  }

  private read(): SavedDocumentTargets {
    const value = this.settings.get(SETTINGS_KEY);
    if (!isSavedDocumentTargets(value)) {
      return { version: 1, targets: {} };
    }

    return {
      version: 1,
      targets: Object.fromEntries(
        Object.entries(value.targets).map(([slideId, target]) => [
          slideId,
          { ...target },
        ]),
      ),
    };
  }
}

function isSavedDocumentTargets(value: unknown): value is SavedDocumentTargets {
  if (!isRecord(value) || value.version !== 1 || !isRecord(value.targets)) {
    return false;
  }

  return Object.values(value.targets).every(
    (target) =>
      isRecord(target) &&
      typeof target.remoteUrl === "string" &&
      typeof target.filePath === "string",
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
