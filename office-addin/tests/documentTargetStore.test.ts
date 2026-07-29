import { afterEach, describe, expect, it, vi } from "vitest";

import {
  createOfficeDocumentSettingsAdapter,
  DocumentTargetStore,
  type DocumentSettingsAdapter,
} from "../src/overleaf/documentTargetStore";

afterEach(() => {
  vi.unstubAllGlobals();
});

function createSettings(initialValue: unknown = null) {
  let value = initialValue;
  const settings: DocumentSettingsAdapter = {
    get: vi.fn(() => value),
    set: vi.fn((_name, nextValue) => {
      value = nextValue;
    }),
    save: vi.fn(async () => undefined),
  };
  return { settings, value: () => value };
}

describe("DocumentTargetStore", () => {
  it("saves targets by stable slide ID and preserves other slides", async () => {
    const { settings, value } = createSettings({
      version: 1,
      targets: {
        "slide-1": {
          remoteUrl: "https://overleaf.example/git/one",
          filePath: "figures/one.pdf",
        },
      },
    });
    const store = new DocumentTargetStore(settings);

    await store.save("slide-2", {
      remoteUrl: "https://overleaf.example/git/two",
      filePath: "figures/two.pdf",
    });

    expect(value()).toEqual({
      version: 1,
      targets: {
        "slide-1": {
          remoteUrl: "https://overleaf.example/git/one",
          filePath: "figures/one.pdf",
        },
        "slide-2": {
          remoteUrl: "https://overleaf.example/git/two",
          filePath: "figures/two.pdf",
        },
      },
    });
    expect(settings.save).toHaveBeenCalledOnce();
    expect(store.get("slide-2")).toEqual({
      remoteUrl: "https://overleaf.example/git/two",
      filePath: "figures/two.pdf",
    });
  });

  it("ignores malformed document data", () => {
    const { settings } = createSettings({ version: 99, targets: "broken" });
    const store = new DocumentTargetStore(settings);

    expect(store.get("slide-1")).toBeNull();
  });
});

describe("createOfficeDocumentSettingsAdapter", () => {
  it("reports a document settings save failure", async () => {
    vi.stubGlobal("Office", {
      AsyncResultStatus: { Failed: "failed" },
    });
    const officeSettings = {
      get: vi.fn(),
      set: vi.fn(),
      saveAsync: vi.fn(
        (
          callback: (result: {
            status: Office.AsyncResultStatus;
            error: { message: string };
          }) => void,
        ) => {
          callback({
            status: Office.AsyncResultStatus.Failed,
            error: { message: "presentation is read-only" },
          });
        },
      ),
    };

    await expect(
      createOfficeDocumentSettingsAdapter(
        officeSettings as unknown as Office.Settings,
      ).save(),
    ).rejects.toThrow("presentation is read-only");
  });
});
