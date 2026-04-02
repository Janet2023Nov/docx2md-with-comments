---
name: docx2md-with-comments
description: "Convert Word (.docx) files to clean Markdown with Word comments preserved as inline blockquotes. Use when the user mentions 'docx to md', 'docx to markdown', 'convert docx', 'docx with comments', '转成md', '转成markdown', or requests to extract content from .docx files while keeping reviewer comments. Do NOT use for creating or editing .docx files."
---
# docx2md-with-comments

Convert Word (.docx) files to clean Markdown with inline comments preserved.

## What This Does

This skill converts `.docx` files to well-formatted Markdown, with a key feature: **Word comments are extracted and placed as blockquotes** (💬) right after the paragraph they reference. This is especially useful for collaborative documents like MMR (Monthly Management Review) reports where reviewer comments carry important context.

## Features

- Headings, paragraphs, lists, tables, hyperlinks → clean Markdown
- Word comments → `> 💬 **Author** (date): comment text` blockquotes
- Bold, italic, underline, strikethrough formatting preserved
- `--start-from` option to extract from a specific heading onward
- Handles structured document tags (SDT) and nested content

## Dependencies

- Python 3.8+ (no external packages needed — uses only the standard library)

## How to Use

When the user asks to convert a `.docx` file to Markdown (with comments), follow these steps:

### Step 1: Confirm the input file

Ask the user for the `.docx` file path if not provided. Verify the file exists.

### Step 2: Determine output path

If the user specifies an output path, use it. Otherwise, generate one from the input filename:
- `"202603 MMR SA.docx"` → `"202603-MMR-SA.md"`

### Step 3: Run the conversion

Execute the conversion script bundled with this skill:

```bash
python <skill_path>/docx2md_v2.py "<input.docx>" "<output.md>"
```

Optional: use `--start-from "heading text"` to skip content before a specific heading.

### Step 4: Report results

Tell the user:
- Output file path
- Number of lines
- Number of comments extracted

## Example Interactions

User: "帮我把 report.docx 转成 md"
→ Run: `python <skill_path>/docx2md_v2.py "report.docx" "report.md"`

User: "Convert meeting-notes.docx to markdown, start from Business Trends"
→ Run: `python <skill_path>/docx2md_v2.py "meeting-notes.docx" "meeting-notes.md" --start-from "Business Trends"`

User: "提取 /path/to/file.docx 到 /output/path.md 要带 comments"
→ Run: `python <skill_path>/docx2md_v2.py "/path/to/file.docx" "/output/path.md"`

## Output Format

```markdown
## Business Trends

### FSI DNB (Lei Kong)

**Security + Agility:** Customer is migrating from IDC to AWS...

> 💬 **Reviewer** (2026-04-01): Great progress on this migration.
> 💬 **Manager** (2026-04-01): Please add MRR numbers.

### Strategic (Lili Liu)

Next section content here...
```

## Notes

- The script uses `lxml` to parse the raw XML inside `.docx` (which is a zip file), so it doesn't need `python-docx` for comment extraction.
- Comments are matched to paragraphs via `commentRangeStart` and `commentReference` markers in the Word XML.
- Empty comments (no text) are silently skipped.
