---
name: pdf-text-replacer
description: "Replace searchable text in PDF files by page number while preserving the original font, size, color, and baseline as much as possible. Use when the user asks to modify a PDF such as replacing page N's XXX with YYY, changing names/titles in invitations or notices, batch-generating PDF variants from a roster, or keeping replacement text visually consistent with the source PDF."
---

# PDF 文字替换工具

在可搜索文字的 PDF 中替换指定文本，尽量保留原文字体、字号、颜色和基线位置。始终优先使用本 skill 提供的脚本，因为 PDF 擦除和文字重写容易在细节上出错。

## 脚本

| 脚本 | 用途 |
|------|------|
| `scripts/replace_pdf_text.py` | 单次或批量编辑：按页替换 PDF 文字（命令行传参 / JSON 文件 / Excel 文件三种模式） |
| `scripts/batch_replace.py` | 批量生成：一份模板 PDF + 一张数据 Excel → 多份个性化 PDF |

---

## 1. 单次/多项替换

用 `scripts/replace_pdf_text.py` 做一次性的或少量指定位置的文字替换。

### 基本用法

```powershell
python skills/pdf-text-replacer/scripts/replace_pdf_text.py `
  --input "输入.pdf" `
  --output "输出.pdf" `
  --page 1 `
  --find "张三院长：" `
  --replace "李四科长：" `
  --verify
```

### 通过 JSON 文件批量替换

```powershell
python skills/pdf-text-replacer/scripts/replace_pdf_text.py `
  --input "输入.pdf" `
  --output "输出.pdf" `
  --edits-json "edits.json" `
  --verify
```

`edits.json` 格式：
```json
[
  {"page": 1, "find": "张三院长：", "replace": "李四科长："},
  {"page": 3, "find": "张三", "replace": "李四", "occurrence": 1}
]
```

### 通过 Excel 文件批量替换

```powershell
python skills/pdf-text-replacer/scripts/replace_pdf_text.py `
  --input "输入.pdf" `
  --output "输出.pdf" `
  --excel "替换表.xlsx" `
  --verify
```

Excel 列结构（第 1 行为表头）：

| A: 页码（从1开始） | B: 搜索文字 | C: 替换文字 | D: 字号（可选） | E: 字体（可选） | F: 颜色（可选） |

### 操作流程

1. 在工作区中找到源 PDF。如有多个可能的 PDF，选用户提到的那份；否则向用户确认。
2. 解析用户的编辑需求：
   - 用户说的页码从 1 开始。给脚本传 `--page 1` 表示第一页。
   - 需要连带移动的标点符号要一并纳入搜索。例如原文是 `张三院长：`，新抬头是 `李四科长：`，搜索时写完整的 `张三院长：` 而不是 `张三院长`。
   - 如果只想替换某一处而非全页同文字，用 `--occurrence N` 指定第几次出现，或向用户确认是哪一处。
3. 始终指定新的输出路径，不覆盖源 PDF。
4. 默认不加字体、字号、颜色参数，让脚本自动继承原文样式。只在用户明确要求时才传覆盖参数。
5. 用 `--verify` 验证替换结果；排版敏感的场合再肉眼检查生成的 PDF 或渲染成图片确认。
6. 报告输出路径和任何局限，特别是源文字不可搜索时。

### 字体覆盖

脚本支持可选样式覆盖：

```powershell
python skills/pdf-text-replacer/scripts/replace_pdf_text.py `
  --input "输入.pdf" --output "输出.pdf" `
  --page 1 --find "旧文字" --replace "新文字" `
  --font-name fang --font-size 14 --color "255,0,0"
```

`--font-name` 可选值：`fang`（仿宋）、`song`（宋体）、`hei`（黑体）、`kai`（楷体）、`li`（隶书）。

---

## 2. 批量生成（模板 + Excel → 多份 PDF）

当你有一份**含占位文字的模板 PDF** 和一张**人员数据 Excel 表**时，用 `scripts/batch_replace.py` 批量生成个性化 PDF — 适用于邀请函、证书、通知、合同等场景。

### 基本用法

```powershell
python skills/pdf-text-replacer/scripts/batch_replace.py `
  -i "邀请函模板.pdf" `
  -e "人员名单.xlsx" `
  -d "输出/" `
  --find "尊敬的来宾：" `
  --replace "尊敬的{{姓名}}{{职务|short}}："
```

### 模板语法

| 语法 | 含义 |
|------|------|
| `{{姓名}}` | 替换为"姓名"列的值 |
| `{{职务\|short}}` | 替换为职务简称（如 设备科科长 → 科长） |

`|short` 过滤器从完整职务中提取末级称谓。目前支持的称谓包括：科长、处长、院长、主任、教授、部长、书记、秘书长、研究员、副主任医师 等。

### 全部参数

```
--input, -i          模板 PDF 路径（必填）
--excel, -e          数据 Excel 路径（必填）
--output-dir, -d     输出目录（必填）
--find, -f           查找的固定文字（必填，不支持模板变量）
--replace, -r        替换模板，支持 {{列名}} / {{列名|short}}（必填）
--filename, -n       输出文件名模板（默认：{{姓名}}.pdf）
--filename-suffix-from  从指定 PDF 文件名推导输出文件名后缀
--sheet, -s          Excel 工作表名（默认自动发现）
--page, -p           替换目标页码，从 1 开始（默认：1）
--verify, -v         生成后验证替换文字是否存在
--overwrite          覆盖已存在的同名文件
--dry-run            预览模式：只打印将要执行的操作，不实际生成文件
```

### Excel 智能识别

脚本自动发现数据所在的工作表和表头行：
- 在前 20 行内扫描包含 ≥2 个关键词匹配列的表头行
- 可识别的关键词：姓名、职务、单位、电话、邮箱、地址、编号 等
- 多 sheet 时优先匹配含"专家""人员""嘉宾""邀请明细""参会""数据""名单"等关键词的工作表
- 不匹配的列会被安全忽略

### 文件名推导

省略 `--filename` 时默认用 `{{姓名}}.pdf`。也可以加 `--filename-suffix-from` 从模板文件名推导后缀：

```
模板文件名："张三院长邀请函【会议通知】研讨会.pdf"
--find "张三院长："
--filename-suffix-from "张三院长邀请函【会议通知】研讨会.pdf"
→ 输出：李四科长邀请函【会议通知】研讨会.pdf、王五主任邀请函【会议通知】研讨会.pdf ……
```

### 操作流程

1. 向用户确认：模板 PDF、数据 Excel、要查找的占位文字、替换模板。
2. 先 `--dry-run` 预览将生成的文件。
3. 用户确认后，去掉 `--dry-run` 实际生成 PDF。
4. 报告输出目录和生成文件数量。

---

## 样式保留机制

脚本会尽量保留原文样式：

- 字体名或就近可用的本地 CJK 字体文件（Windows + macOS 双平台路径）
- 字号
- 文字颜色
- 原文字基线位置

字体解析优先级：用户指定覆盖 → 匹配的本地系统字体文件 → PyMuPDF 内置 CJK 字体 → 回退链。

内置 CJK 字体一览：

| 别名 | PyMuPDF 字体名 | 风格 |
|------|---------------|------|
| `fang` / `仿宋` | `china-t` | 仿宋 |
| `song` / `宋体` | `china-s` | 宋体 |
| `hei` / `黑体` | `china-ss` | 黑体 |
| `kai` / `楷体` | `china-ts` | 楷体 |
| `li` / `隶书` | `china-cs` | 隶书 |

PDF 内嵌的子集字体不一定能直接复用于新文字。若精确复用失败，脚本会自动回退到本机常见中文字体（Windows 下 simfang.ttf、simsun.ttc、simhei.ttf、simkai.ttf；macOS 下 Songti.ttc、STHeiti、Kaiti.ttc），最后回退到 PyMuPDF 内置 CJK 字体。

---

## 局限

- 源文字必须可搜索/可复制。扫描版 PDF 需要先做 OCR。
- 替换文字以绘制方式写入 PDF，周围排版不会自动重排。
- 若替换文字比原文长很多，需要检查是否会与后续文字重叠。
- 擦除操作会永久删除输出副本中匹配到的原文。
- 批量脚本的 `--find` 是固定文字，不支持模板变量——只有 `--replace` 和 `--filename` 支持 `{{列名}}` 语法。
