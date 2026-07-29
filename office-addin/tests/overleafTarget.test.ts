import { describe, expect, it } from "vitest";

import { parseOverleafTarget } from "../src/overleaf/overleafTarget";

describe("parseOverleafTarget", () => {
  it("normalizes a pasted Overleaf Git URL and PDF path", () => {
    expect(
      parseOverleafTarget(
        " https://git@overleaf.example/git/0123456789abcdef01234567/ ",
        " figures/result.pdf ",
      ),
    ).toEqual({
      endpoint: "https://overleaf.example",
      filePath: "figures/result.pdf",
      remoteUrl:
        "https://overleaf.example/git/0123456789abcdef01234567",
    });
  });

  it.each([
    ["http://overleaf.example/git/project", "figure.pdf", "必须使用 HTTPS"],
    [
      "https://git:secret@overleaf.example/git/project",
      "figure.pdf",
      "不能包含密码",
    ],
    [
      "https://overleaf.example/git/project?token=secret",
      "figure.pdf",
      "不能包含查询参数",
    ],
    ["https://overleaf.example/git/project", "../figure.pdf", "不能包含 .."],
    ["https://overleaf.example/git/project", "/figure.pdf", "不能以 / 开头"],
    ["https://overleaf.example/git/project", ".git/figure.pdf", "不能写入 .git"],
    ["https://overleaf.example/git/project", "figure.png", "必须以 .pdf 结尾"],
  ])("rejects an unsafe target", (remoteUrl, filePath, message) => {
    expect(() => parseOverleafTarget(remoteUrl, filePath)).toThrow(message);
  });
});
