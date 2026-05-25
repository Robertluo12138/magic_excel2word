# 固定 Prompt（请勿修改，整段复制给 Agent）

下面 `~~~` 标记之间的内容就是要原样复制粘贴给 Agent 的固定 prompt。
本次测试是 **learn-only 第一次真实文件巡检**，规则不可商量。**不要**
对 prompt 做任何删改、改写、压缩或翻译，也不要附加你自己的上下文；如果
有疑问，先和工程同学确认，再决定是否粘贴。

~~~
你正在帮一个内部 Python 工具（`magic_excel2word`，仓库根目录即当前
工作目录）做**第一次真实文件 learn-only 匹配测试**。这是一次受限的
learn-mode 巡检；下列约束**不可商量**，违反任何一条都立即停下并报告。

## 仅允许的操作

1. 先确认输入文件已经被操作员放好：
   - `operator_pilot/input/historical.xlsx`（历史 Excel）
   - `operator_pilot/input/finished_report.docx`（对应的成品 Word 报告）
   如果其中任意一个不存在，停下并要求操作员补齐；**不要**继续，也
   **不要**用任何其它路径替代。

2. 按顺序、且仅限以下三个 `python -m src.main` 子命令，把产物全部
   写入 `operator_pilot/output/`：

   - `python -m src.main learn --excel operator_pilot/input/historical.xlsx --word operator_pilot/input/finished_report.docx --out operator_pilot/output --strict`
   - `python -m src.main validate-artifacts --out operator_pilot/output`
   - `python -m src.main pilot-summary --out operator_pilot/output`

   每一步分别记录退出码。即便某一步以预期内的非 `0` 退出（例如
   `learn --strict` 退出 `3`），仍然继续运行后两步——产物仍会落盘，
   后两步的检查仍然有信息量。但一旦某一步以**预期外**的码退出
   （`2` 或任何不在下面合约里的码），立刻停下来汇报。

3. 报告时**仅**使用以下三类聚合信息：
   a. 每一步的退出码；
   b. `operator_pilot/output/` 目录里出现了哪些文件（**只**列
      basename，例如 `mapping_review.xlsx`，不要列绝对路径或父
      目录链）；
   c. `pilot-summary` 这一步的 stdout 原文（该命令本身就是
      redacted-by-contract：不会泄露 Word 数字、Excel 值、sheet/cell
      内容、`word_id`、绝对路径或公司标识，所以可以原样转述）。

## 严格禁止的操作

- **禁止**运行 `confirm-mapping`、`run-preview`、`render-docx`、
  `validate-render`，以及上面那三个之外的任何 `python -m src.main`
  子命令（包括 `pilot-preflight`、`generate-synthetic`）。本次仅做
  learn-mode 巡检。
- **禁止**所有会改动 git 状态的操作：`git add`、`git commit`、
  `git push`、`git stash`、`git rm`、`git reset`、`git checkout`、
  `git restore`、`git clean`。`operator_pilot/input/` 和
  `operator_pilot/output/` 已被仓库根 `.gitignore` 兜底，跑完后
  `git status` 不会出现这两个目录里的真实/生成文件；如果意外出现，
  **停下来**告诉操作员，不要试图「整理」它。
- **禁止**用 `cat`、`head`、`tail`、`less`、`Read` 工具或任何其它
  方式打开 `operator_pilot/input/` 下的真实 Excel 或 Word 文件；
  也**禁止**打开或摘抄 `operator_pilot/output/` 下任一 artifact 的
  具体内容，包括但不限于 `mapping_review.xlsx`、`auto_mapping.yml`、
  `confidence_report.md`、`converted_template.docx`。仅允许用 `ls`
  确认 basename 是否存在。
- **禁止**在聊天里粘贴或转述以下任何一类内容（包括摘要、片段、
  脱敏后看起来还能反推业务的版本）：
  · Word 报告中的具体数字、句子、段落、表头；
  · Excel 中任何单元格的原始值；
  · source sheet 名、cell 地址（如 `Sheet1!B12`）、行号、列号；
  · 任何绝对路径（包括 `/Users/...`、`C:\...`、`~/...` 展开形式）；
  · 公司名、产品名、人名、客户名、订单号或其它能定位到具体业务的
    标识。
- **禁止**调用 LLM、网络请求、GUI、Microsoft Office 自动化或任何
  云服务。本工具不需要这些，强行调用都属于越权。
- **禁止**尝试「修复」匹配结果。learn-mode 是巡检，不是修改流程。
  即便看到 UNRESOLVED / LOW，也只**汇报聚合计数**；不要手动改任何
  artifact，也不要建议操作员手动改。

## 退出码合约（按这个口径汇报）

- `learn --strict`：`0` = 没有 eligible UNRESOLVED/LOW；
  `3` = 至少一个 eligible 数字是 UNRESOLVED 或 LOW（产物仍落盘，
  允许下一步 `validate-artifacts` 跑，但**不允许**升级到 confirm）。
- `validate-artifacts`：`0` = learn 写出的四份产物互相一致；
  `4` = 不一致（bug 信号，立即停下来）。
- `pilot-summary`：`0` = 已输出 redacted 汇总；`11` = 缺
  `auto_mapping.yml`，说明 learn 实际上没跑成功。
- 任何其它退出码（包括 `2`）都按「停下来、报告码、不要继续」处理。

## 报告模板

严格按这个模板回复；缺什么就写 `(missing)`，不要补描述：

```
Stage 1 — learn --strict
  exit code: <number>
  artifacts under operator_pilot/output/ (basenames only):
    - <basename>
    - <basename>
    ...

Stage 2 — validate-artifacts
  exit code: <number>

Stage 3 — pilot-summary
  exit code: <number>
  pilot-summary stdout (redacted-by-contract, paste verbatim):
    <stdout>

Working tree check:
  `git status --short` mentions paths under operator_pilot/input/
  or operator_pilot/output/: <"none" | count only — no path content>

Guarantee — I did NOT:
  - run confirm-mapping / run-preview / render-docx / validate-render
    / pilot-preflight / generate-synthetic
  - run git add / git commit / git push / git stash / git rm /
    git reset / git checkout / git restore / git clean
  - open, cat, head, tail, or quote any real Excel / Word / artifact
    content
  - paste any raw Word number, Excel value, sheet/cell address,
    absolute path, or company identifier
```
~~~
