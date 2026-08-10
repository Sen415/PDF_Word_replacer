# PDF 文字替换工具

按页替换 PDF 可搜索文字，尽量保留原文字体、字号、颜色和基线位置。支持 Excel 批量替换 — **一份模板 + 一张数据表 → 多份个性化 PDF**。

基于 PyMuPDF 实现，核心逻辑为"先擦后写"：

1. `page.search_for(text)` → 定位旧文字位置
2. `page.add_redact_annot()` → 标记擦除区域
3. `page.apply_redactions()` → 永久删除旧文字
4. `page.insert_text()` → 在原位写入新文字

## 适用场景

- **邀请函 / 证书 / 通知**：模板抬头、署名按人员名单批量替换
- **合同 / 协议**：将模板占位符替换为实际数据
- 任何"一份模板 + 一张数据表 → 多份 PDF"的场景

## 环境要求

- Python 3.8+
- PyMuPDF (`pip install pymupdf`)
- openpyxl (`pip install openpyxl`)，仅在从 Excel 读取替换列表时需要

```bash
pip install pymupdf openpyxl
```

## 文件说明

| 文件 | 用途 |
|------|------|
| `pdf_replacer.py` | 核心模块：文字替换、Excel 读取、验证 |
| `main.py` | 命令行入口：单次替换 / Excel 驱动多编辑替换 |
| `batch_replacer.py` | 核心模块：模板引擎、Excel 智能识别、批量生成 |
| `batch_invitation.py` | 命令行入口：批量生成（模板 + Excel → 多份 PDF） |

## 用法一：单次或少量替换

### 代码调用

```python
from pdf_replacer import replace_from_list

edits = [
    # (页码, 搜索文字, 替换文字, 字号, 字体别名, 颜色)
    (0, "张三院长：", "李四科长：", None, None, None),
]

stats = replace_from_list("输入.pdf", "输出.pdf", edits)
print(stats)  # {"total": 1, "ok": 1, "fail": 0}
```

- 页码从 0 开始（第 1 页 = 0）
- 字号、字体、颜色为 `None` 时自动继承原文样式

### 命令行

```bash
python main.py -i 输入.pdf -o 输出.pdf
```

默认会执行示例替换（第 1 页"张三"→"李四"），修改 `main.py` 底部的 `edits` 列表即可自定义。

### 从 Excel 读取替换列表

```bash
python main.py --excel 替换表.xlsx -i 输入.pdf -o 输出.pdf
```

Excel 格式（第 1 行为表头）：

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| 页码 | 搜索文字 | 替换文字 | 字号（可选） | 字体（可选） | 颜色（可选） |

---

## 用法二：批量生成（模板 + Excel → 多份 PDF）

适用于：一张人员表、一份邀请函模板，批量生成每人一份的个性化 PDF。

### 命令行

```bash
python batch_invitation.py \
  -i "邀请函模板.pdf" \
  -e "人员名单.xlsx" \
  -d "输出目录/" \
  --find "尊敬的来宾：" \
  --replace "尊敬的{{姓名}}{{职务|short}}："
```

### 模板语法

| 语法 | 含义 |
|------|------|
| `{{姓名}}` | 替换为"姓名"列的值 |
| `{{职务\|short}}` | 替换为职务简称（如 设备科科长 → 科长） |

`|short` 过滤器支持的称谓：科长、处长、院长、主任、教授、部长、书记、秘书长、研究员、副主任医师 等。

### 全部参数

```
-i, --input          模板 PDF 路径
-e, --excel          数据 Excel 路径
-d, --output-dir     输出目录
-f, --find           查找的固定文字（不支持模板变量）
-r, --replace        替换模板，支持 {{列名}} / {{列名|short}}
-n, --filename       输出文件名模板（默认: {{姓名}}.pdf）
--filename-suffix-from  从 PDF 文件名推导后缀
-s, --sheet          Excel 工作表名（默认自动发现）
-p, --page           替换目标页码，从 1 开始（默认: 1）
-v, --verify         生成后验证替换文字是否存在
--overwrite          覆盖已存在文件
--dry-run            预览模式，不实际生成文件
```

### Excel 智能识别

脚本自动发现数据所在的工作表和表头行：
- 在前 20 行扫描包含 ≥2 个关键词匹配列的表头行
- 可识别关键词：姓名、职务、单位、电话、邮箱、地址、编号 等
- 多 Sheet 时优先匹配含"专家""人员""嘉宾""名单"等关键词的工作表
- 不关心的列会被安全忽略

### 文件名推导

省略 `--filename` 时默认用 `{{姓名}}.pdf`。添加 `--filename-suffix-from` 可从模板文件名推导后缀：

```
模板文件名：张三院长邀请函【会议通知】研讨会.pdf
--find "张三院长："
--filename-suffix-from "张三院长邀请函【会议通知】研讨会.pdf"
→ 输出：李四科长邀请函【会议通知】研讨会.pdf
```

### 预览模式

建议先用 `--dry-run` 预览，确认无误后再正式生成：

```bash
python batch_invitation.py -i 模板.pdf -e 人员.xlsx -d 输出/ \
    --find "XXX" --replace "{{姓名}}" --dry-run
```

### 配置文件直接运行

也可以直接修改 `batch_invitation.py` 顶部的 `DEFAULTS` 字典，然后无需命令行参数直接运行：

```python
DEFAULTS = {
    "input_pdf":    "我的模板.pdf",
    "excel":        "人员名单.xlsx",
    "output_dir":   "输出/",
    "find":         "尊敬的来宾：",
    "replace":      "尊敬的{{姓名}}{{职务|short}}：",
    "filename":     "",
    "page":         1,
    "dry_run":      False,
}
```

```bash
python batch_invitation.py
```

## 字体支持

### 内置 CJK 字体

脚本自带 PyMuPDF 内置 CJK 字体，无需额外安装中文字体即可使用：

| 别名 | 字体名 | 风格 |
|------|--------|------|
| `fang` / `仿宋` | `china-t` | 仿宋 |
| `song` / `宋体` | `china-s` | 宋体 |
| `hei` / `黑体` | `china-ss` | 黑体 |
| `kai` / `楷体` | `china-ts` | 楷体 |
| `li` / `隶书` | `china-cs` | 隶书 |

### 字体回退策略

当需要尽量复用原文样式时，字体解析优先级为：

1. 用户手动指定的字体
2. 系统已安装的匹配中文字体文件
3. PyMuPDF 内置 CJK 字体
4. 回退到仿宋 (`china-t`)

## 样式保留

脚本会尽量保留原文字的：
- 字体（或就近可用的 CJK 字体）
- 字号
- 文字颜色
- 基线位置

## 局限

- 源文字必须**可搜索/可复制**。扫描版 PDF 需要先 OCR 处理
- 替换文字以绘制方式写入，**周围排版不会自动重排**
- 若替换文字比原文长很多，可能与后续文字重叠
- 批量脚本的 `--find` 是固定文字，不支持模板变量 — 只有 `--replace` 和 `--filename` 支持 `{{列名}}` 语法

## 许可

MIT License
