# Slide2Pdf

[中文说明](#中文说明) | [English](#english)

![Slide2Pdf buttons in the PowerPoint ribbon](./screenshots/ribbon.png)

## 中文说明

Slide2Pdf 可以将当前幻灯片单独导出为 PDF，也可以自动裁到可见内容边界，方便把 PowerPoint 制作的图示插入 LaTeX 或其他文档。

| 版本 | 运行环境 | 主要特点 |
| --- | --- | --- |
| Office Add-in | Windows 或 macOS、新版 Microsoft 365 PowerPoint | 跨平台主要维护版本；可下载到本机或推送到 Overleaf |
| Windows VSTO | Windows 10/11、PowerPoint 2013 及以上版本 | Windows 兼容版本；一键导出整页或按内容裁切 |

### Windows 安装

1. 从 [Releases](https://github.com/duanyll/Slide2Pdf/releases) 下载最新版本。
2. 解压下载的文件。
3. 双击 `setup.exe`，按提示完成安装。
4. 重新打开 PowerPoint。在“开始”选项卡中可以找到 Slide2Pdf 的两个按钮。

也可以使用 [Scoop](https://scoop.sh/) 安装：

```powershell
scoop bucket add duanyll https://github.com/duanyll/scoop-bucket
scoop install duanyll/slide2pdf
```

Windows 版提供两种导出方式：

- `Export Full Slide`：按幻灯片原始尺寸导出当前页。
- `Crop to Content`：导出当前页，并裁到可见内容边界。

对于已经保存的演示文稿，Slide2Pdf 会记住每一页的导出位置。按住 `Shift` 再点击导出按钮，可以重新选择位置。

### macOS 安装

1. 完全退出 PowerPoint。
2. 打开 [Slide2Pdf 安装页](https://slide2pdf.duanyll.com)，复制安装命令。
3. 在“终端”中运行命令，然后重新打开 PowerPoint。
4. 进入“开始”选项卡，点击“加载项”，再选择 Slide2Pdf。

macOS 版提供以下选项：

- `导出整页`：保留幻灯片原始尺寸。
- `按内容裁切`：只保留当前页可见对象的范围。
- `透明背景`：将整页纯白背景改为透明。

导出文件名形如 `Presentation_Slide3_1.pdf`。再次导出同一页时，末尾序号会依次递增。

本机下载不会上传演示文稿或 PDF。只有点击“生成并推送”时，生成的 PDF 才会发送到所填写的 Overleaf Git 仓库。

#### 推送到 Overleaf

Office Add-in 可以为每张幻灯片保存独立的 Overleaf Git 仓库和 PDF 路径：

1. 展开“推送到 Overleaf”。
2. 粘贴项目的 HTTPS Git 地址，并填写项目中的 PDF 路径和 Git Token。
3. 选择输出范围，然后点击“生成并推送”。

仓库地址和 PDF 路径随演示文稿保存；选择“记住此设备上的 Token”后，Token 只保存在当前设备的 Office 网页存储中，不会写入演示文稿。

直接连接要求 Overleaf Git 服务允许 Slide2Pdf 的网页来源。`overleaf-pro` 管理员可以在部署配置中加入：

```dotenv
GIT_BRIDGE_ALLOWED_CORS_ORIGINS=https://slide2pdf.duanyll.com,https://localhost:3000
```

#### 裁切说明

裁切范围按当前页可见对象的外框计算。母版或版式中的图形不会计入，阴影、光晕等超出外框的效果可能被截掉；遇到这种情况时，请改用“导出整页”。

## English

Slide2Pdf exports the current PowerPoint slide as an individual PDF. It can also crop the PDF to the bounds of visible content, making PowerPoint diagrams ready to insert into LaTeX or other documents.

| Version | Requirements | Highlights |
| --- | --- | --- |
| Office Add-in | Windows or macOS with a recent Microsoft 365 PowerPoint release | Primary cross-platform version; download locally or push to Overleaf |
| Windows VSTO | Windows 10/11 and PowerPoint 2013 or later | Windows compatibility version; export at full size or crop to content |

### Install on Windows

1. Download the latest version from [Releases](https://github.com/duanyll/Slide2Pdf/releases).
2. Extract the downloaded archive.
3. Run `setup.exe` and follow the installer.
4. Reopen PowerPoint. The two Slide2Pdf buttons appear on the Home tab.

You can also install the Windows version with [Scoop](https://scoop.sh/):

```powershell
scoop bucket add duanyll https://github.com/duanyll/scoop-bucket
scoop install duanyll/slide2pdf
```

The Windows version provides two export actions:

- `Export Full Slide` exports the current slide at its original size.
- `Crop to Content` exports the current slide and crops it to visible content.

For saved presentations, Slide2Pdf remembers the export location for each slide. Hold `Shift` while clicking an export button to choose a different location.

### Install on macOS

1. Quit PowerPoint completely.
2. Open the [Slide2Pdf installation page](https://slide2pdf.duanyll.com) and copy the installation command.
3. Run the command in Terminal, then reopen PowerPoint.
4. Open the Home tab, select Add-ins, and choose Slide2Pdf.

The macOS version provides these options:

- `导出整页` exports the current slide at its original size.
- `按内容裁切` crops the PDF to the bounds of visible objects.
- `透明背景` makes a solid white slide background transparent.

Exported files are named like `Presentation_Slide3_1.pdf`. Exporting the same slide again increments the final number.

Local downloads don't upload the presentation or PDF. The generated PDF is sent to the configured Overleaf Git repository only when you select **Generate and push**.

#### Push to Overleaf

The Office Add-in can remember a separate Overleaf Git repository and PDF path for each slide:

1. Expand **Push to Overleaf**.
2. Paste the project's HTTPS Git URL, then enter the PDF path and Git token.
3. Choose the output bounds and select **Generate and push**.

The repository URL and PDF path are stored in the presentation. If you choose to remember the token, it is stored only in the Office web storage on that device and isn't written to the presentation.

Direct connections require the Overleaf Git service to allow the Slide2Pdf web origin. An `overleaf-pro` administrator can add:

```dotenv
GIT_BRIDGE_ALLOWED_CORS_ORIGINS=https://slide2pdf.duanyll.com,https://localhost:3000
```

#### Cropping details

Cropping uses the bounds of visible objects on the current slide. Master and layout graphics are not included, and effects outside an object's bounds, such as shadows or glows, may be clipped. Use the full-slide export when this happens.

## Development

To build the Windows VSTO add-in, install the Office/SharePoint Development workload for Visual Studio 2022, then open `Slide2Pdf.sln`.

For the macOS Office.js add-in:

```bash
cd office-addin
npm install
npm run sideload:local
npm start
```

Run the project checks with:

```bash
npm run typecheck
npm test
npm run build
npm run validate
```

Deploy the static Office.js frontend with `npm run deploy`.

## Changelog

- v1.0.0.3
  - Remember the export location for each slide. Hold `Shift` while clicking an export button to select a new location.
  - Place the buttons on the Home tab by default.
- v1.0.0.2
  - Initial release.
