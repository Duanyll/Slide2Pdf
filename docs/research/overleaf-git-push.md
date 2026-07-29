# Office Add-in 直推 Overleaf PDF 可行性调查

调查日期：2026-07-29

目标实例：`https://overleaf.villa.moe`

目标代码：[`ayaka-notes/overleaf-pro`](https://github.com/ayaka-notes/overleaf-pro/tree/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee)

## 结论

这项功能可行，而且不需要复用 Overleaf 网页端的 session、内部 HTTP API 或 WebSocket。Office Add-in 可以在浏览器环境里运行 `isomorphic-git`，把生成的 PDF 写入一个浅克隆的工作树，提交后通过 Overleaf Git Bridge 推送。

本次实现采用 direct-first：用户为每张幻灯片填写任意 HTTPS Git remote 和 PDF 路径，这些目标随演示文稿保存；PAT 按 endpoint 保存在 Office webview 的 `localStorage`，不会写入 PPT。Add-in 使用 `isomorphic-git` 和 LightningFS 管理完整浅克隆，push 前先 fast-forward pull，且永远不 force。

Cloudflare Worker 不再是默认依赖，也不由 Slide2Pdf 提供公共额度。目标 fork 已经支持精确来源的 CORS 白名单；实例管理员配置 `GIT_BRIDGE_ALLOWED_CORS_ORIGINS` 后即可直连。当前 `overleaf.villa.moe` 尚未放行生产和开发来源，实测预检仍返回 `403`；在实例无法配置 CORS 时，使用者可以在自己的 Cloudflare 账户中部署可选的无状态代理。

## 已核实的接口与认证方式

目标 fork 显示的克隆地址是：

```text
https://git@overleaf.villa.moe/git/<24 位项目 ID>
```

这是标准 HTTPS Basic Authentication：用户名固定为 `git`，密码是 `olp_...` Git authentication token。实际用程序调用时，建议把 URL 写成不含用户信息的形式，再通过认证回调提供用户名和密码：

```text
remote:   https://overleaf.villa.moe/git/<project-id>
username: git
password: <Git authentication token>
```

依据如下：

- fork 的编辑器直接拼出 `https://git@<host>/git/<project-id>`，[见 Git Bridge 对话框源码](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/git-bridge/frontend/js/components/git-bridge-modal.tsx#L27-L50)；Nginx 把 `/git/` 前缀去掉后转发给 `git-bridge:8000`，[见反向代理配置](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/server-ce/nginx/nginx.conf.template#L24-L38)。
- Git Bridge 只接受用户名 `git`，并把密码当作 access token 验证，[见 `Oauth2Filter`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/git-bridge/src/main/java/uk/ac/ic/wlgitbridge/server/Oauth2Filter.java#L78-L139)。Overleaf 官方文档也规定用户名为 `git`、token 作为密码：[Git integration authentication tokens](https://docs.overleaf.com/integrations-and-add-ons/git-integration-and-github-synchronization/git/git-integration-authentication-tokens)。
- token 是用户级凭据，不是单项目 token。fork 给它分配 `git_bridge` scope、设置一年有效期并在数据库中保存哈希，[见 token manager](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/oauth2-server/app/src/OAuthPersonalAccessTokenManager.mjs#L52-L87)；项目读写权仍由各个 API 路由检查，[见 Git Bridge router](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/git-bridge/app/src/GitBridgeRouter.mjs#L12-L35)。

不要把 token 放进 URL、查询参数或 Git remote 配置。应使用 `isomorphic-git` 的 `onAuth: () => ({ username: "git", password: token })`，让 token 只进入 `Authorization` 请求头。[`isomorphic-git` 的认证文档](https://isomorphic-git.org/docs/en/onAuth)说明浏览器 Git 也采用 HTTPS Basic Authentication。

## Smart HTTP 流程

Overleaf 暴露的是标准 Git smart HTTP。一次同步大致包含四类请求：

```text
读取：GET  <remote>/info/refs?service=git-upload-pack
      POST <remote>/git-upload-pack

推送：GET  <remote>/info/refs?service=git-receive-pack
      POST <remote>/git-receive-pack
```

Git 先通过 `info/refs` 发现引用，再用相应的 POST 交换 pkt-line 和 packfile；请求及响应的 MIME 类型是 `application/x-git-*-request/result`。[Git 官方 HTTP protocol](https://git-scm.com/docs/http-protocol)给出了完整流程。Worker 若参与，只需逐字节转发这些 GET/POST，不需要理解 Git 对象。

仓库中的端到端测试已经证明 `isomorphic-git/http/web` + `LightningFS` 能在浏览器里完成 Git Bridge 的 clone、修改、commit、push 和 pull，[测试的客户端配置](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/server-ce/test/git-bridge.spec.ts#L13-L32)及[实际读写流程](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/server-ce/test/git-bridge.spec.ts#L254-L415)都在 fork 自己的测试套件中。`isomorphic-git` 官方也明确支持浏览器中的 clone、commit 和 push，并建议用 LightningFS 提供浏览器文件系统：[项目主页](https://isomorphic-git.org/en/)、[浏览器文件系统](https://isomorphic-git.org/docs/en/fs)。

Office Add-in 不会改变这套网络模型。任务窗格的 HTML/JavaScript 运行在浏览器或 webview 中，Windows 使用 WebView2，macOS/iOS 使用 WKWebView，Android 使用 Chrome；因此能够运行纯 JavaScript Git 客户端，同时也会受到浏览器同源策略约束。[Office Add-ins 平台概览](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)、[各平台 webview 说明](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)、[同源策略说明](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/addressing-same-origin-policy-limitations?view=office-js)。

## CORS 实测与 Worker 是否必要

fork 的 Git Bridge 已有 CORS handler：

- 来源必须和白名单中的字符串完全一致；
- 允许 `GET, HEAD, PUT, POST, DELETE`；
- 允许请求头 `Authorization, Content-Type`；
- 预检命中白名单时返回 `200`，否则返回 `403`。

实现见 [`CORSHandler`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/git-bridge/src/main/java/uk/ac/ic/wlgitbridge/server/CORSHandler.java#L12-L52)。白名单来自 `GIT_BRIDGE_ALLOWED_CORS_ORIGINS`，默认只有 `https://localhost`，[见运行配置模板](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/git-bridge/conf/envsubst_template.json#L1-L15)。多个来源用逗号连接，源码不会清理拆分后每一项的空格，所以配置值中不要在逗号后加空格。

当前 handler 没有把 `Git-Protocol` 列入允许的请求头。`isomorphic-git` 的常规 clone/fetch/push 使用兼容流程，不需要主动添加这个 header；实现时不要额外启用 protocol v2。若使用 `listServerRefs`，应显式选择 `protocolVersion: 1`，该方法默认使用 v2，[见官方说明](https://isomorphic-git.org/docs/en/listServerRefs)。将来若确实需要 v2，应同时给 fork 的 CORS handler 增加 `Git-Protocol`，不能只改来源白名单。

建议部署配置：

```dotenv
GIT_BRIDGE_ALLOWED_CORS_ORIGINS=https://slide2pdf.duanyll.com,https://localhost:3000
```

2026-07-29 对目标实例的探测与端到端验证结果：

| 探测 | 结果 |
| --- | --- |
| 使用 `.env` 凭据执行 `git ls-remote --symref` | 成功；当前默认分支为 `master` |
| `git-upload-pack` 的认证 `info/refs` | `200`，返回正确 smart-Git MIME |
| `git-receive-pack` 的认证 `info/refs` | `200`，返回正确 smart-Git MIME |
| `Origin: https://slide2pdf.duanyll.com` 的预检 | `403`，没有 CORS allow headers |
| 同一来源的认证 `git-upload-pack info/refs` | 服务端返回 `200`，但没有 `Access-Control-Allow-Origin`，浏览器仍会拦截 |
| `Origin: https://localhost:3000` 的预检 | `403`，本地 manifest 的真实来源也未放行 |
| `Origin: https://localhost` 的预检 | `200`，精确回显该来源 |
| 原生 Git clone、添加 PDF、push、删除 PDF、再次 push | 全部成功 |
| `isomorphic-git` 浅 clone、二进制 PDF push、清理 push | 全部成功 |
| `isomorphic-git` 经本地 Worker 流式代理完成同一流程 | 全部成功；PAT 由 Worker Secret 注入，客户端未持有 PAT |
| 本次实现的 `OverleafGitClient` 在目标实例添加并清理 PDF | 全部成功；最终 tree 与测试前一致 |
| 生产使用的 `isomorphic-git/http/web` + LightningFS | 在 fake IndexedDB 与本地 smart-HTTP 服务上通过完整 push 测试 |

验证使用 `.env` 中的测试凭据，但没有输出或复制 token。可行性调查最初产生了 6 个测试 commit；实现完成后又运行了一次 opt-in live test，产生 2 个“添加 PDF / 清理 PDF”commit。最终远端 tree 与测试前完全一致，项目内容已恢复；共 8 个测试 commit 仍保留在 Git 历史中。`.env` 中可见的变量名为 `OVERLEAF_ENDPOINT`、`OVERLEAF_GIT_KEY` 和 `OVERLEAF_GIT_REPO`；本报告没有记录对应值。

由此可得：

1. **默认使用 direct Git。** 实例在做快照后加入精确来源并重启 Git Bridge；Add-in 从本机存储读取每位用户自己的 PAT。
2. **无法配置实例 CORS 时，由用户自部署代理。** Worker 额度和上游 allowlist 都归使用者自己的 Cloudflare 账户管理，Slide2Pdf 不承担公共代理成本。
3. **不提供中央公共代理。** 接受任意 endpoint 的公共代理既容易滥用额度，也会成为 SSRF 入口。

## 可选 Worker 应该做什么，不应该做什么

### 用户自部署：无状态 smart-HTTP 代理

Worker 只需：

1. 允许 `GET`、`POST`、`OPTIONS`；
2. 只放行 `info/refs`、`git-upload-pack`、`git-receive-pack` 三种路径，通过部署变量固定上游 host 或 allowlist，避免变成开放代理或 SSRF 入口；
3. 原样转发查询参数、`Content-Type`、`Accept` 和请求 body；
4. 默认原样转发客户端的 `Authorization`，不在 Worker 中保存 PAT；
5. 流式返回上游 status、body、Git MIME 及 `WWW-Authenticate`，禁止缓存；
6. 不记录 `Authorization`、请求体和上游 URL 中可能含有的用户信息。

Workers 的 `Request` 可以用新 URL 复用原请求的方法、headers 和 body，[官方 Request API](https://developers.cloudflare.com/workers/runtime-apis/request/)给出了这种写法；直接转发 `response.body` 会保持流式处理，无需把 packfile 放进内存，[官方 Streams 文档](https://developers.cloudflare.com/workers/runtime-apis/streams/)也明确建议代理场景这样处理。当前项目只有静态资产配置；如果采用代理，需要给 [`wrangler.jsonc`](../../office-addin/wrangler.jsonc) 增加 `main` 和 Worker 脚本，并让 `/api/*` 优先进入脚本。Cloudflare 的[静态资产路由文档](https://developers.cloudflare.com/workers/static-assets/binding/)支持按路径配置 `run_worker_first`。

代理应由使用者部署到自己的账户，并用环境变量配置允许的上游。若部署者选择在 Worker Secret 中保存单用户 PAT，则必须额外限制调用者和项目；这属于私有部署选项，不是 Slide2Pdf 的默认认证模型。[Worker Secret](https://developers.cloudflare.com/workers/configuration/secrets/)可用于这种场景。

### 不推荐：在 Worker 中运行完整 Git 客户端

纯 JavaScript Git 或预编译 WebAssembly 理论上可以在 Workers 运行，但相对首版需求，收益有限、代价很高。Git 客户端还需要可写文件系统、完整当前树、对象库和并发控制；若每次请求临时 clone，速度和资源占用都很差，若把仓库持久化到 R2 或 Durable Objects，又要自行实现文件系统适配和事务。

Cloudflare 虽支持预编译 Wasm，[见 WebAssembly 文档](https://developers.cloudflare.com/workers/runtime-apis/webassembly/)，但每个 isolate 只有 128 MB 内存，请求体上限还取决于账户套餐，[见 Workers limits](https://developers.cloudflare.com/workers/platform/limits/)。把 packfile 原样流式代理不会触碰这类复杂度；在 Worker 中解包、改树、重新打包则会迅速碰到内存、CPU 和临时文件系统问题。

把 PAT 放进 Worker Secret **不要求 Worker 自己运行 Git**。本次 POC 已验证：Add-in 侧的 `isomorphic-git` 仍然管理完整工作树，Worker 只注入认证头并转发字节流。只有当 API 的输入变成“单独一个 PDF”，而不是 Git smart-HTTP 请求时，Worker 才必须承担 clone、修改、commit 和 push，此方案首版不采用。

## 最小实现建议

### 1. 先拆开“生成 PDF”和“保存目的地”

本次已将 [`exportCurrentSlide.ts`](../../office-addin/src/export/exportCurrentSlide.ts) 拆成可复用的 PDF 生成结果和下载包装：

```text
generateCurrentSlidePdf() -> { data: Uint8Array, suggestedFileName: string }
saveToDownload(result)
pushToOverleaf(result, destination)
```

这样下载和 Overleaf 共用同一份 PDF 生成逻辑，后续也容易增加“同时下载并推送”。

### 2. Add-in 内维护一个浅克隆

依赖建议：

```text
isomorphic-git
@isomorphic-git/lightning-fs
```

首次连接某个项目时执行 `clone({ singleBranch: true, depth: 1 })`，让 LightningFS 把仓库存到 IndexedDB。深度可以是 1，但必须保留当前提交的**完整文件树**，不能创建只含目标 PDF 的孤立仓库。`isomorphic-git` 的 bundler 用法见[官方 quick start](https://isomorphic-git.org/docs/en/quickstart-with-bundlers)，浅克隆参数见 [`clone`](https://isomorphic-git.org/docs/en/clone)。

不要硬编码 `main` 或 `master`。当前被测项目的 HEAD 是 `master`，fork 新项目的初始化代码却使用 `main`，[见 `GitProjectRepo`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/git-bridge/src/main/java/uk/ac/ic/wlgitbridge/bridge/repo/GitProjectRepo.java#L64-L87)。应以 clone/fetch 返回的默认分支为准。

### 3. 每次推送都基于远端最新状态

推荐顺序：

```text
获取项目级互斥锁
pull --fast-forward-only
把 PDF 写到最终远端路径
git add <new-path>
如用户明确选择“移动”，git remove <old-path>
commit
push（永远不 force）
释放锁
```

`isomorphic-git pull` 支持 `fastForwardOnly`，[见官方文档](https://isomorphic-git.org/docs/en/pull)；`push` 默认不 force，并返回每个 ref 的服务端结果，[见官方文档](https://isomorphic-git.org/docs/en/push.html)。

如果 pull 之后、push 之前有人在 Overleaf 网页端保存了修改，push 会被拒绝。此时 fetch 最新 HEAD，丢弃本次尚未推上的本地 commit，在最新树上重新写入同一个 PDF、重新 commit，再重试一到两次；不要自动 force。fork 的 receive hook 会拒绝非 fast-forward/force push，[见 `WriteLatexPutHook`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/git-bridge/src/main/java/uk/ac/ic/wlgitbridge/git/handler/hook/WriteLatexPutHook.java#L112-L141)，并明确说明 Overleaf 在 push 期间发生更新时客户端应重试，[见 `WLReceivePackFactory`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/git-bridge/src/main/java/uk/ac/ic/wlgitbridge/git/handler/WLReceivePackFactory.java#L45-L48)。

目标 fork 还有一个比普通 Git 冲突更窄、但确实存在的竞态窗口：服务端检查 history version 后立即返回 `202`，随后才异步逐个 upsert 和删除文件，[见接收与异步处理流程](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/git-bridge/app/src/GitBridgeApiController.mjs#L322-L365)。如果用户恰好在这个窗口里通过网页端编辑，Git push 可能覆盖这次修改；若网页端刚创建了一个不在被推 Git 树里的新文件，异步删除阶段还可能把它删掉。处理中途报错也没有事务回滚，理论上会留下部分更新。

首版应把这个限制写进操作约束：推送的几秒内不要同时在网页端编辑；每次 push 成功后再 fetch 一次核对远端 HEAD 和目标 PDF；遇到异常时停止自动重试并提示用户检查项目。长期可在 fork 侧把 version check 与整批快照应用改成原子操作，但这不是 Slide2Pdf 客户端能单独解决的问题。

### 4. 文件名采用稳定路径，不再追加下载序号

Overleaf 目标路径应由用户第一次配置，例如：

```text
figures/experiment-overview.pdf
```

后续导出始终覆盖这个路径；Git 会记录每次版本，LaTeX 引用也不用改。若目标路径变化：

- “另存为”只添加新路径，旧文件保留；
- “移动/改名”在同一 commit 中删除旧路径并添加新路径。Git 没有必须显式调用的 rename 操作，服务端会按最终树执行删除和新增。

PDF 会被 fork 判断为 binary file，并通过 `upsertFileWithPath` 创建或替换，[见 `processFileUpdate`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/git-bridge/app/src/GitBridgeApiController.mjs#L489-L529)。路径不能含 `..`、不能以 `/` 开头，也不能进入 `.git`，[见路径校验](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/git-bridge/app/src/GitBridgeApiController.mjs#L610-L637)。

Git Bridge 默认限制单文件不超过 50 MiB、项目文件数不超过 2000，[见配置模板](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/git-bridge/conf/envsubst_template.json#L12-L15)；部署方可以覆盖这些值。客户端应在 push 前检查 PDF 大小并给出明确错误，不要等服务端处理完整个 packfile 后才失败。

## 必须防住的两个数据风险

### 推送的是完整快照，不是“上传单个文件”

Git Bridge 会把新 commit 的整棵树转换成 Overleaf 快照。fork 明确把新树中缺失的所有现有文件删除，[见 `processSnapshotPush`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/git-bridge/app/src/GitBridgeApiController.mjs#L376-L449)。因此绝不能从空仓库创建一个只含 PDF 的 commit 再推送，否则可能删除项目中的 `.tex`、图片和参考文献。

首版必须通过正常 clone 得到完整当前树，只修改一个指定路径；集成测试也应先使用一次性项目，验证其他文件在 push 后仍然存在。

### token 及本地凭据管理

Git token 能访问该用户有权限的项目。本次实现允许用户选择把 PAT 按 endpoint 保存到分区后的 `localStorage`；远端地址和 PDF 路径按 `slide.id` 保存到 Office Document Settings，PAT 不进入文档。关闭“记住此设备上的 Token”时，PAT 只存在于当前任务窗格。浏览器存储可能被用户设置阻止或清理，此时 Add-in 会要求重新输入。[Office 状态持久化文档](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings)说明了这些平台差异。

还有两项应在开发前处理：

1. 本次调查已把 `.env` 加入 [`.gitignore`](../../.gitignore)，避免误提交真实 token；后续凭据仍不得写入源码、普通 Worker 变量或日志。
2. 目标 fork 的 token 校验代码会在 debug 级别记录明文 `accessToken`，[见 `verifyToken`](https://github.com/ayaka-notes/overleaf-pro/blob/49b8243d9ad5d9c29b0b1c348a0bb7a3b8e71eee/services/web/modules/oauth2-server/app/src/OAuthPersonalAccessTokenManager.mjs#L107-L139)。生产环境应保持 debug 日志关闭，并尽快删除这两条敏感日志；否则一旦开启 debug，Git token 会进入服务日志。

如果使用自部署 Worker，token 仍由 Add-in 提供并随请求转发；不要运营共享 PAT 的公共代理。Worker 必须限制上游域名和允许的路径、禁止记录认证头，并给所有 Git 响应设置 `Cache-Control: no-store`。

## 推荐落地顺序

1. **已完成：** 拆分 PDF 生成与下载，为每张幻灯片保存 remote/path，并按 endpoint 保存可选 PAT。
2. **已完成：** 使用 `isomorphic-git` + LightningFS 完成浅 clone、fast-forward pull、覆盖 PDF、commit 和非 force push。
3. **已完成：** 本地 smart-HTTP、生产 web HTTP 适配器、完整树保留、缓存 pull 和目标实例 live test。
4. **部署前：** 为 `overleaf.villa.moe` 配置生产与本地开发来源的 CORS 白名单；按全局运维约定先做实例快照。
5. **已完成：** push 被非 fast-forward 拒绝后恢复到干净基线，下一次手动推送会先拉取远端更新；后续可再加入同一次操作内的自动重试。
6. **后续增强：** 为不能修改 CORS 的用户提供自部署 Worker 模板。

第一版让 Add-in 管理工作树和逐页配置，Overleaf Git Bridge 负责同步，不依赖 Slide2Pdf 托管的 Worker。这既支持任意用户配置的 Overleaf endpoint，也避免公共代理额度和滥用风险。
