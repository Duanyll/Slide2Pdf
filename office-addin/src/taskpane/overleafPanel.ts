import {
  createCurrentSlidePdf,
  type ExportMode,
  type ExportProgress,
} from "../export/exportCurrentSlide";
import {
  createOfficeDocumentSettingsAdapter,
  DocumentTargetStore,
} from "../overleaf/documentTargetStore";
import { LocalCredentialStore } from "../overleaf/localCredentialStore";
import {
  parseOverleafTarget,
  type OverleafTarget,
} from "../overleaf/overleafTarget";
import type { GitSyncProgress } from "../overleaf/overleafGitClient";
import { getCurrentSlide } from "../powerpoint/getCurrentSlide";

type StatusState = "idle" | "working" | "success" | "error";

interface OverleafPanelOptions {
  isTransparentBackgroundEnabled: () => boolean;
  setBusy: (busy: boolean) => void;
  showPdfProgress: (progress: ExportProgress) => void;
  showStatus: (message: string, state: StatusState) => void;
}

const gitProgressMessages: Record<GitSyncProgress, string> = {
  cloning: "首次连接，正在读取 Overleaf 项目…",
  pulling: "正在同步 Overleaf 项目的最新版本…",
  writing: "正在更新项目中的 PDF…",
  pushing: "正在推送到 Overleaf；完成前请不要同时在网页端编辑…",
  verifying: "正在核对 Overleaf 中的 PDF…",
};

export function initializeOverleafPanel(options: OverleafPanelOptions): void {
  const panel = getElement<HTMLDetailsElement>("#overleaf-panel");
  const slideLabel = getElement<HTMLElement>("#overleaf-slide");
  const remoteInput = getElement<HTMLInputElement>("#overleaf-remote");
  const pathInput = getElement<HTMLInputElement>("#overleaf-path");
  const tokenInput = getElement<HTMLInputElement>("#overleaf-token");
  const rememberToken = getElement<HTMLInputElement>("#remember-overleaf-token");
  const exportMode = getElement<HTMLSelectElement>("#overleaf-export-mode");
  const saveButton = getElement<HTMLButtonElement>("#save-overleaf-target");
  const pushButton = getElement<HTMLButtonElement>("#push-overleaf");

  const targetStore = new DocumentTargetStore(
    createOfficeDocumentSettingsAdapter(Office.context.document.settings),
  );
  const credentialStore = createCredentialStore();
  let activeSlideId: string | null = null;

  const loadCurrentSlide = async (force = false): Promise<void> => {
    try {
      const slide = await getCurrentSlide(false);
      if (!force && slide.slideId === activeSlideId) return;

      activeSlideId = slide.slideId;
      slideLabel.textContent = `当前幻灯片：第 ${slide.slideIndex + 1} 页`;
      const target = targetStore.get(slide.slideId);
      remoteInput.value = target?.remoteUrl ?? "";
      pathInput.value = target?.filePath ?? `figures/slide-${slide.slideIndex + 1}.pdf`;
      loadRememberedToken();
    } catch (error) {
      activeSlideId = null;
      slideLabel.textContent = getErrorMessage(error);
    }
  };

  const loadRememberedToken = (): void => {
    tokenInput.value = "";
    rememberToken.checked = false;
    if (!credentialStore || !remoteInput.value.trim()) return;

    try {
      const target = parseOverleafTarget(
        remoteInput.value,
        pathInput.value || "figure.pdf",
      );
      const token = credentialStore.get(target.endpoint);
      if (token) {
        tokenInput.value = token;
        rememberToken.checked = true;
      }
    } catch {
      // The complete validation error is shown when the user saves or pushes.
    }
  };

  const persistForm = async (
    slideId: string,
    requireToken: boolean,
  ): Promise<{ target: OverleafTarget; token: string }> => {
    const target = parseOverleafTarget(remoteInput.value, pathInput.value);
    const token = tokenInput.value.trim();
    if (requireToken && !token) {
      throw new Error("请输入 Overleaf Git Token。");
    }

    await targetStore.save(slideId, {
      remoteUrl: target.remoteUrl,
      filePath: target.filePath,
    });

    if (credentialStore) {
      if (rememberToken.checked && token) {
        if (!credentialStore.save(target.endpoint, token)) {
          throw new Error("无法在此设备上保存 Token。仍可取消“记住 Token”后继续使用。");
        }
      } else if (!rememberToken.checked) {
        credentialStore.remove(target.endpoint);
      }
    } else if (rememberToken.checked) {
      throw new Error("此 Office 环境不允许使用本机存储。请取消“记住 Token”。");
    }

    return { target, token };
  };

  saveButton.addEventListener("click", async () => {
    if (!activeSlideId) {
      options.showStatus("请先选择一张幻灯片。", "error");
      return;
    }

    options.setBusy(true);
    try {
      await persistForm(activeSlideId, false);
      options.showStatus("已保存当前页的目标。请保存演示文稿以保留这项设置。", "success");
    } catch (error) {
      options.showStatus(getErrorMessage(error), "error");
    } finally {
      options.setBusy(false);
    }
  });

  pushButton.addEventListener("click", async () => {
    if (!activeSlideId) {
      options.showStatus("请先选择一张幻灯片。", "error");
      return;
    }
    const requestedSlideId = activeSlideId;

    options.setBusy(true);
    try {
      const { target, token } = await persistForm(requestedSlideId, true);
      const pdf = await createCurrentSlidePdf(
        exportMode.value as ExportMode,
        {
          transparentBackground: options.isTransparentBackgroundEnabled(),
        },
        options.showPdfProgress,
      );
      if (pdf.slideId !== requestedSlideId) {
        throw new Error("生成 PDF 时切换了幻灯片。请确认当前页后重新推送。");
      }

      const { pushPdfFromBrowser } = await import(
        "../overleaf/browserGitClient"
      );
      const result = await pushPdfFromBrowser({
        data: pdf.data,
        filePath: target.filePath,
        remoteUrl: target.remoteUrl,
        token,
        onProgress: (progress) =>
          options.showStatus(gitProgressMessages[progress], "working"),
      });

      options.showStatus(
        result.changed
          ? `已推送到 Overleaf：${target.filePath}`
          : "PDF 没有变化，Overleaf 已是最新版本。",
        "success",
      );
    } catch (error) {
      options.showStatus(formatPushError(error), "error");
    } finally {
      options.setBusy(false);
    }
  });

  remoteInput.addEventListener("change", loadRememberedToken);
  panel.addEventListener("toggle", () => {
    if (panel.open) void loadCurrentSlide(true);
  });
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => void loadCurrentSlide(),
  );
  void loadCurrentSlide(true);
}

function createCredentialStore(): LocalCredentialStore | null {
  try {
    return new LocalCredentialStore(
      window.localStorage,
      Office.context.partitionKey,
    );
  } catch {
    return null;
  }
}

function getElement<T extends Element>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) {
    throw new Error(`Missing task pane element: ${selector}`);
  }
  return element;
}

function formatPushError(error: unknown): string {
  const message = getErrorMessage(error);
  const lowerMessage = message.toLowerCase();

  if (
    lowerMessage.includes("failed to fetch") ||
    lowerMessage.includes("networkerror") ||
    lowerMessage.includes("cors")
  ) {
    return "无法连接 Git 仓库。请检查地址，并确认 Overleaf 已允许 Slide2Pdf 的网页来源。";
  }
  if (
    lowerMessage.includes("401") ||
    lowerMessage.includes("authorization") ||
    lowerMessage.includes("authentication")
  ) {
    return "Overleaf 拒绝了登录。请检查 Git Token 后重试。";
  }
  if (
    lowerMessage.includes("non-fast-forward") ||
    lowerMessage.includes("fast-forward") ||
    lowerMessage.includes("fetch first")
  ) {
    return "Overleaf 项目刚刚发生了变化。请重新推送；Slide2Pdf 不会强制覆盖远端版本。";
  }
  return `推送失败：${message}`;
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
