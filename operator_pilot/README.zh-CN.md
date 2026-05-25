# operator_pilot — 真实文件 learn-only 测试投放区

本目录是给非技术运营同学准备的「真实历史 Excel + 成品 Word 报告」**第一次
learn-only 匹配测试**投放区。它只做 learn-mode 巡检；不做 confirm、不做
run-preview、不做 render-docx、不做 validate-render，也不会把任何文件提交
到 git。

> ⚠️ 这是 **prototype**。运行的目的是观察 learn-mode 能否扫到一份真实的
> 报告对，并产出可复核的中间产物。不要把任何产物直接发给 stakeholder，
> 也不要由此跳过 `docs/real_file_pilot.md` 的完整流程。

## 为什么 drop-zone 在仓库内？

正常的真实文件 pilot 路径（见 `docs/real_file_pilot.md` §1）要求把真实
文件放在仓库**外**（例如 `~/pilot_data/`）。这里的 `operator_pilot/` 是
该规则的一个**故意收窄的例外**，专门给「第一次 learn-only 测试」用，
完整说明见 `docs/real_file_pilot.md` §1a「Operator first-time learn-only
drop-zone (narrow exception)」。

例外被三条不变量约束，违反任何一条都要立即停下来按 §1 重新走外部路径：

- **仅限 learn 模式。** `PROMPT_FOR_AGENT.md` 里固定 prompt 只允许
  `learn --strict` / `validate-artifacts` / `pilot-summary` 三条命令，
  并明确禁止 `confirm-mapping` / `run-preview` / `render-docx` /
  `validate-render`。一旦超出 learn 模式，必须切回 `docs/real_file_pilot.md`
  §2 的仓库外流程。
- **`.gitignore` 兜底。** `operator_pilot/input/` 和
  `operator_pilot/output/` 下除 `.gitkeep` 外的所有内容都被忽略，跑完
  之后必须做一次 `git status` 自检。
- **四文件不变量。** 仓库里 `operator_pilot/` 下只允许这四个被追踪的
  文件存在：`README.zh-CN.md`、`PROMPT_FOR_AGENT.md`、
  `input/.gitkeep`、`output/.gitkeep`。其它任何文件被 commit 都会被
  仓库 privacy-boundary 测试在 CI 拦截。

`pilot-preflight` 在这个例外下**不**使用：它的 inside-repo 闸门（退出
码 `12`）按设计会拒绝任何仓库内路径——这对正常 pilot 是正确行为，
但和本 drop-zone 的用途冲突。后续所有 confirm / render / validate-render
阶段仍然按 `docs/real_file_pilot.md` §2、§3 走仓库外流程，并由
`pilot-preflight` 充当只读预检闸门。

## 工作流程（四步）

1. 把**一份**历史 Excel 拷贝（建议用拷贝而非剪切，保留原文件不被改动）成：

   ```
   operator_pilot/input/historical.xlsx
   ```

2. 把**对应的**那份成品 Word 报告拷贝成：

   ```
   operator_pilot/input/finished_report.docx
   ```

3. 打开

   ```
   operator_pilot/PROMPT_FOR_AGENT.md
   ```

   把里面 `~~~` 之间的整段固定 prompt **原样**复制、粘贴给 Agent（不要
   删改、不要附加上下文、不要分批粘贴）。

4. 等 Agent 跑完。learn-mode 产物会落在：

   ```
   operator_pilot/output/
   ```

   你可以在本机直接打开 `mapping_review.xlsx` 和 `confidence_report.md`，
   按 `docs/artifact_review_guide.md` 的检查表逐项核对。

## 安全边界（请严格遵守）

- 不要在 `operator_pilot/input/` 之外放真实文件。其它路径**没有**
  `.gitignore` 兜底，一次手滑就可能进入 commit。
- 不要让 Agent 跑 `confirm-mapping` / `run-preview` / `render-docx` /
  `validate-render`，也不要跑除 learn-mode 三件套以外的任何 src.main
  子命令。本测试**只**做 learn-mode 巡检。
- 不要让 Agent 执行 `git add` / `git commit` / `git push` /
  `git stash` 等会改动 git 状态的命令。`operator_pilot/input/` 和
  `operator_pilot/output/` 已被 `.gitignore` 兜底，跑完后 `git status`
  应该看不到这两个目录下的真实/生成文件。
- 不要让 Agent 在聊天里粘贴：Word 报告中的具体数字、Excel 单元格的原始
  值、source sheet 名 / cell 地址、绝对路径（包括 `/Users/...`、
  `C:\...`、`~/...` 展开形式）、公司名/产品名/客户名等业务标识。**只**
  让 Agent 报告：每一步的退出码、各 artifact 是否落盘（只列 basename）、
  以及 `pilot-summary` 命令本身的 stdout（该命令是 redacted-by-contract，
  可以原样转述）。
- 跑完之后请运行一次 `git status` 自检：`operator_pilot/input/` 和
  `operator_pilot/output/` 下的真实/生成文件**不应该**出现在
  「Changes to be committed」或「Untracked files」中。如果出现了，说明
  `.gitignore` 兜底失效，立刻停止操作并联系工程同学。

## 出错怎么办？

- `learn --strict` 退出码非 `0`（典型是 `3`）：说明这份历史对里至少有
  一个 Word 数字没法在 Excel 中匹配（UNRESOLVED）或匹配置信度太低
  （LOW）。**不要**强行往下走；按 `docs/real_file_pilot.md` §4 的
  「Handling unresolved / LOW / EXCLUDED / non-renderable / tampered
  rows」处理。本次测试**不**要求做 confirm，所以一个非 0 的 strict
  退出码本身就是有价值的信号，不是失败。
- `validate-artifacts` 退出码非 `0`（典型是 `4`）：说明 learn 写出的
  四份产物自己对不齐。这是 bug 信号，请记录退出码和 Agent 的简短描述
  （不要把产物原文贴出来），联系工程同学。
- `pilot-summary` 退出码非 `0`（典型是 `11`）：说明 learn 没产出
  `auto_mapping.yml`，通常意味着上一步 learn 实际上没跑成功。
- 其它任何退出码（包括 `2`）：让 Agent 直接停下来报告码，不要继续。

## 这套测试**不会**做的事

- 不会调用任何 LLM 或云服务。
- 不会启动 GUI。
- 不会调用 Microsoft Word / Excel 自动化；底层只用 `openpyxl` +
  `python-docx`。
- 不会生成新的 Word 报告（`render-docx` 是后续阶段，本测试关闭）。
- 不会把真实文件或产物提交到 git。

## 在仓库里**会**保留的文件

只有四个：

- `operator_pilot/README.zh-CN.md`（本文件）
- `operator_pilot/PROMPT_FOR_AGENT.md`（固定 prompt）
- `operator_pilot/input/.gitkeep`
- `operator_pilot/output/.gitkeep`

其它所有出现在 `operator_pilot/input/` 或 `operator_pilot/output/` 下的
文件都会被 `.gitignore` 兜底，**不会**被追踪。这是设计；不要试图
「整理」掉那些未追踪文件。
