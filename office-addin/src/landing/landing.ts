import "./landing.css";

const copyButton = document.querySelector<HTMLButtonElement>("#copy-install");
const installCommand = document.querySelector<HTMLElement>("#install-command");
const copyStatus = document.querySelector<HTMLElement>("#copy-status");

copyButton?.addEventListener("click", async () => {
  if (!installCommand) {
    return;
  }

  const command = installCommand.textContent?.trim() || "";
  let copied = false;
  try {
    await navigator.clipboard.writeText(command);
    copied = true;
  } catch {
    copied = copyWithTemporaryTextArea(command);
  }

  const label = copyButton.querySelector("span");
  if (label) {
    label.textContent = copied ? "已复制" : "复制失败";
  }
  copyButton.dataset.copyState = copied ? "success" : "error";
  if (copyStatus) {
    copyStatus.textContent = copied
      ? "安装命令已复制到剪贴板。"
      : "无法自动复制，请手动选择安装命令。";
  }

  window.setTimeout(() => {
    if (label) {
      label.textContent = "复制";
    }
    delete copyButton.dataset.copyState;
  }, 1800);
});

function copyWithTemporaryTextArea(value: string): boolean {
  const textArea = document.createElement("textarea");
  textArea.value = value;
  textArea.setAttribute("readonly", "");
  textArea.style.position = "fixed";
  textArea.style.opacity = "0";
  document.body.append(textArea);
  textArea.select();
  try {
    return document.execCommand("copy");
  } catch {
    return false;
  } finally {
    textArea.remove();
  }
}
