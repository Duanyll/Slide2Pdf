import "./taskpane.css";

import {
  exportCurrentSlide,
  type ExportMode,
  type ExportProgress,
} from "../export/exportCurrentSlide";

const progressMessages: Record<ExportProgress, string> = {
  "reading-slide": "正在读取当前幻灯片…",
  "creating-pdf": "正在生成 PDF…",
  "processing-pdf": "正在处理当前页…",
  saving: "正在准备下载…",
};

Office.onReady((info) => {
  const app = document.querySelector<HTMLElement>("#app");
  if (!app) return;

  if (info.host !== Office.HostType.PowerPoint) {
    showStatus("Slide2Pdf 只能在 PowerPoint 中运行。", "error");
    return;
  }

  const supportsPowerPoint = Office.context.requirements.isSetSupported(
    "PowerPointApi",
    "1.10",
  );
  const supportsPdf = Office.context.requirements.isSetSupported("File", "1.1");
  if (!supportsPowerPoint || !supportsPdf) {
    showStatus("当前 PowerPoint 版本不支持导出。请更新 PowerPoint 后重试。", "error");
    return;
  }

  app.removeAttribute("hidden");
  document.querySelector("#loading")?.setAttribute("hidden", "");

  bindExportButton("#export-slide", "slide");
  bindExportButton("#export-content", "content");

  const saveHint = document.querySelector<HTMLElement>("#save-hint");
  if (saveHint) {
    saveHint.textContent =
      "文件名会自动添加递增序号；演示文稿和 PDF 不会上传。";
  }
});

function bindExportButton(selector: string, mode: ExportMode): void {
  const button = document.querySelector<HTMLButtonElement>(selector);
  button?.addEventListener("click", async () => {
    await runExport(mode);
  });
}

async function runExport(mode: ExportMode): Promise<void> {
  setBusy(true);

  try {
    const fileName = await exportCurrentSlide(
      mode,
      { transparentBackground: isTransparentBackgroundEnabled() },
      (progress) => showStatus(progressMessages[progress], "working"),
    );
    showStatus(`已开始下载：${fileName}`, "success");
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      showStatus("已取消。", "idle");
    } else {
      const message = error instanceof Error ? error.message : String(error);
      showStatus(`导出失败：${message}`, "error");
    }
  } finally {
    setBusy(false);
  }
}

function setBusy(busy: boolean): void {
  document
    .querySelectorAll<HTMLButtonElement>("button[data-export]")
    .forEach((button) => {
      button.disabled = busy;
    });
  const transparentBackground =
    document.querySelector<HTMLInputElement>("#transparent-background");
  if (transparentBackground) transparentBackground.disabled = busy;
}

function isTransparentBackgroundEnabled(): boolean {
  return (
    document.querySelector<HTMLInputElement>("#transparent-background")
      ?.checked ?? false
  );
}

function showStatus(
  message: string,
  state: "idle" | "working" | "success" | "error",
): void {
  const status = document.querySelector<HTMLElement>("#status");
  if (!status) return;
  status.removeAttribute("hidden");
  status.textContent = message;
  status.dataset.state = state;
}
