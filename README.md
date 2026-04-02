# docx2md-with-comments

An agent skill for converting Word (.docx) files to clean Markdown — with Word comments preserved as inline blockquotes.

English | [中文](#中文)

## What This Does

**docx2md-with-comments** converts `.docx` files to well-formatted Markdown while extracting Word comments and placing them as `> 💬` blockquotes right after the paragraph they reference.

This is especially useful for collaborative review documents (like MMR reports) where reviewer comments carry important context that shouldn't be lost in conversion.

## Key Features

- **Comments Preserved** — Word comments become `> 💬 **Author** (date): text` blockquotes inline with content
- **Full Formatting** — Headings, bold, italic, underline, strikethrough, lists, tables, hyperlinks
- **Selective Extraction** — `--start-from` flag to extract from a specific heading onward
- **Zero Dependencies** — Uses only Python standard library, no `pip install` needed

## Installation

### For Kiro Users

#### Option 1: Clone and copy

```bash
git clone https://github.com/Janet2023Nov/docx2md-with-comments.git
cp -r docx2md-with-comments/docx2md-with-comments ~/.kiro/skills/docx2md-with-comments
```

#### Option 2: Clone and symlink (easy to update)

```bash
git clone https://github.com/Janet2023Nov/docx2md-with-comments.git ~/docx2md-with-comments
ln -s ~/docx2md-with-comments/docx2md-with-comments ~/.kiro/skills/docx2md-with-comments
```

### Install Python dependency

Python 3.8+ is all you need — no external packages required.

## Usage

In Kiro chat, just say:

```
帮我把 report.docx 转成 md
```

or

```
Convert meeting-notes.docx to markdown with comments
```

The skill will automatically use the bundled `docx2md_v2.py` script to convert.

## Output Example

```markdown
## Business Trends

### FSI DNB (Lei Kong)

**Security + Agility:** Customer is migrating from IDC to AWS...

> 💬 **Reviewer** (2026-04-01): Great progress on this migration.
> 💬 **Manager** (2026-04-01): Please add MRR numbers.
```

## Standalone CLI Usage

You can also use the script directly:

```bash
# Basic conversion
python docx2md_v2.py input.docx output.md

# Extract from a specific heading onward
python docx2md_v2.py input.docx output.md --start-from "Business Trends"

# Auto-generate output filename
python docx2md_v2.py input.docx
```

## Requirements

- Python 3.8+

## License

MIT

---

<a name="中文"></a>

# docx2md-with-comments

一个 Agent 技能，用于将 Word (.docx) 文件转换为干净的 Markdown —— 同时保留 Word 批注作为内联引用块。

## 这是什么

**docx2md-with-comments** 将 `.docx` 文件转换为格式良好的 Markdown，同时提取 Word 批注并以 `> 💬` 引用块的形式放在对应段落之后。

特别适合协作审阅文档（如 MMR 报告），审阅者的批注在转换过程中不会丢失。

## 核心特性

- **批注保留** — Word 批注变为 `> 💬 **作者** (日期): 内容` 引用块，内联在正文中
- **完整格式** — 标题、加粗、斜体、下划线、删除线、列表、表格、超链接
- **选择性提取** — `--start-from` 参数可从指定标题开始提取
- **零依赖** — 仅使用 Python 标准库，无需 `pip install` 任何包

## 安装

### Kiro 用户

```bash
git clone https://github.com/Janet2023Nov/docx2md-with-comments.git
cp -r docx2md-with-comments/docx2md-with-comments ~/.kiro/skills/docx2md-with-comments
```

### 安装 Python 依赖

Python 3.8+ 即可，无需安装任何外部包。

## 使用方法

在 Kiro 聊天中直接说：

```
帮我把 report.docx 转成 md
```

技能会自动调用内置的 `docx2md_v2.py` 脚本完成转换。

## 环境要求

- Python 3.8+
