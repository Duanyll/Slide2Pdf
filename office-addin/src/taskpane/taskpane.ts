import "./taskpane.css";

import {
  exportCurrentSlide,
  type ExportMode,
  type ExportProgress,
} from "../export/exportCurrentSlide";

const progressMessages: Record<ExportProgress, string> = {
  "reading-slide": "正在读取当前幻灯片…",
  "creating-pdf": "正在让 PowerPoint 生成整份 PDF…",
  "processing-pdf": "正在抽取并处理当前页…",
  saving: "正在保存 PDF…",
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
    showStatus("当前 PowerPoint 版本缺少导出所需的 Office.js API。", "error");
    return;
  }

  app.removeAttribute("hidden");
  document.querySelector("#loading")?.setAttribute("hidden", "");

  bindExportButton("#export-slide", "slide");
  bindExportButton("#export-content", "content");

  const directSave =
    window.location.hostname === "localhost" || "showSaveFilePicker" in window;
  const saveHint = document.querySelector<HTMLElement>("#save-hint");
  if (saveHint) {
    saveHint.textContent = directSave
      ? "首次选择文件后，同一页会自动覆盖保存；按住 Shift 点击可另存。"
      : "PDF 会保存到浏览器下载位置；所有演示文稿内容仍只在本机处理。";
  }
});

function bindExportButton(selector: string, mode: ExportMode): void {
  const button = document.querySelector<HTMLButtonElement>(selector);
  button?.addEventListener("click", async (event) => {
    await runExport(mode, event.shiftKey);
  });
}

async function runExport(mode: ExportMode, forceNewPath: boolean): Promise<void> {
  setBusy(true);

  try {
    const result = await exportCurrentSlide(
      mode,
      forceNewPath,
      (progress) => showStatus(progressMessages[progress], "working"),
    );
    showStatus(
      result.method === "file"
        ? `已保存 ${result.fileName}`
        : `已下载 ${result.fileName}`,
      "success",
    );
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      showStatus("已取消。", "idle");
    } else {
      showStatus(error instanceof Error ? error.message : String(error), "error");
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
}

function showStatus(
  message: string,
  state: "idle" | "working" | "success" | "error",
): void {
  const status = document.querySelector<HTMLElement>("#status");
  if (!status) return;
  status.textContent = message;
  status.dataset.state = state;
}
