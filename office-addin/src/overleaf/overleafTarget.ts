export interface OverleafTarget {
  endpoint: string;
  remoteUrl: string;
  filePath: string;
}

export function parseOverleafTarget(
  remoteUrlInput: string,
  filePathInput: string,
): OverleafTarget {
  let remoteUrl: URL;
  try {
    remoteUrl = new URL(remoteUrlInput.trim());
  } catch {
    throw new Error("Git 仓库地址无效。请粘贴完整的 HTTPS 地址。");
  }

  if (remoteUrl.protocol !== "https:") {
    throw new Error("Git 仓库地址必须使用 HTTPS。");
  }
  if (remoteUrl.password) {
    throw new Error("Git 仓库地址不能包含密码。请在 Token 输入框中填写凭据。");
  }
  if (remoteUrl.search) {
    throw new Error("Git 仓库地址不能包含查询参数。");
  }
  if (remoteUrl.hash) {
    throw new Error("Git 仓库地址不能包含片段标识。");
  }

  remoteUrl.username = "";
  remoteUrl.pathname = remoteUrl.pathname.replace(/\/+$/, "");

  const filePath = filePathInput.trim().replaceAll("\\", "/");
  if (filePath.startsWith("/")) {
    throw new Error("PDF 路径不能以 / 开头。");
  }

  const pathParts = filePath.split("/");
  if (pathParts.includes("..")) {
    throw new Error("PDF 路径不能包含 ..。");
  }
  if (pathParts.includes(".git")) {
    throw new Error("PDF 不能写入 .git 目录。");
  }
  if (!filePath.toLowerCase().endsWith(".pdf")) {
    throw new Error("PDF 路径必须以 .pdf 结尾。");
  }

  return {
    endpoint: remoteUrl.origin,
    remoteUrl: remoteUrl.toString().replace(/\/$/, ""),
    filePath,
  };
}
