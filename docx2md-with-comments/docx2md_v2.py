#!/usr/bin/env python3
"""
docx2md.py - Word (.docx) to clean Markdown converter
Zero external dependencies — uses only Python standard library.

Usage:
    python docx2md_v2.py input.docx [output.md] [--start-from 'heading text']

If output.md is not specified, it will be generated from the input filename.

Features:
    - Converts headings, paragraphs, lists, tables, hyperlinks
    - Extracts Word comments and places them as blockquotes after related paragraphs
    - Handles bold, italic, underline, strikethrough formatting
    - Cleans up excessive whitespace and HTML artifacts
"""
import sys
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Word XML namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
}

# Register namespaces so ElementTree can resolve prefixes in findall
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def _local_tag(element):
    """Get the local tag name (without namespace) from an element."""
    tag = element.tag
    if tag and '}' in tag:
        return tag.split('}', 1)[1]
    return tag or ''


def parse_comments(zip_file):
    """Extract comments from word/comments.xml."""
    comments = {}
    try:
        with zip_file.open('word/comments.xml') as f:
            tree = ET.parse(f)
    except KeyError:
        return comments  # No comments in this document

    root = tree.getroot()
    for comment_el in root.findall('.//w:comment', NS):
        cid = comment_el.get(f'{{{NS["w"]}}}id')
        author = comment_el.get(f'{{{NS["w"]}}}author', 'Unknown')
        date = comment_el.get(f'{{{NS["w"]}}}date', '')
        # Clean date
        date_clean = date.split('T')[0] if 'T' in date else date
        if date_clean == '1900-01-01':
            date_clean = ''
        # Get all text content from the comment
        text_parts = []
        for p in comment_el.findall('.//w:p', NS):
            p_text = ''.join(r.text or '' for r in p.findall('.//w:t', NS))
            if p_text.strip():
                text_parts.append(p_text.strip())
        text = ' '.join(text_parts)

        if cid is not None:
            comments[cid] = {
                'author': author,
                'date': date_clean,
                'text': text,
            }
    return comments


def parse_hyperlinks(zip_file):
    """Extract hyperlink relationships from document.xml.rels."""
    rels = {}
    try:
        with zip_file.open('word/_rels/document.xml.rels') as f:
            tree = ET.parse(f)
    except KeyError:
        return rels
    for rel in tree.getroot():
        rid = rel.get('Id')
        target = rel.get('Target', '')
        rel_type = rel.get('Type', '')
        if 'hyperlink' in rel_type:
            rels[rid] = target
    return rels


def get_run_raw(run):
    """Extract raw text and format tuple from a w:r element.
    Returns (format_tuple, text) where format_tuple is (bold, italic, strike)."""
    rpr = run.find('w:rPr', NS)
    bold = rpr is not None and rpr.find('w:b', NS) is not None
    italic = rpr is not None and rpr.find('w:i', NS) is not None
    strike = rpr is not None and rpr.find('w:strike', NS) is not None

    texts = []
    for child in run:
        tag = _local_tag(child)
        if tag == 't':
            texts.append(child.text or '')
        elif tag == 'tab':
            texts.append('\t')
        elif tag == 'br':
            texts.append('\n')
        elif tag == 'sym':
            texts.append('')  # skip symbols

    return ((bold, italic, strike), ''.join(texts))


def merge_runs_to_md(raw_runs):
    """Merge adjacent runs with same formatting, then convert to markdown.
    Also absorbs whitespace-only runs into adjacent bold/italic groups,
    and moves leading/trailing whitespace outside ** markers."""
    if not raw_runs:
        return ''

    # Pass 1: absorb whitespace-only plain runs into adjacent formatted groups
    absorbed = []
    plain = (False, False, False)
    for i, (fmt, text) in enumerate(raw_runs):
        if text.strip() == '' and fmt == plain:
            prev_fmt = absorbed[-1][0] if absorbed else None
            next_fmt = raw_runs[i + 1][0] if i + 1 < len(raw_runs) else None
            if prev_fmt and next_fmt and prev_fmt == next_fmt and prev_fmt != plain:
                absorbed[-1] = (prev_fmt, absorbed[-1][1] + text)
                continue
        absorbed.append((fmt, text))

    # Pass 2: merge consecutive runs with same format
    groups = []
    cur_fmt = None
    cur_texts = []
    for fmt, text in absorbed:
        if fmt == cur_fmt:
            cur_texts.append(text)
        else:
            if cur_texts:
                groups.append((cur_fmt, ''.join(cur_texts)))
            cur_fmt = fmt
            cur_texts = [text]
    if cur_texts:
        groups.append((cur_fmt, ''.join(cur_texts)))

    # Pass 3: convert to markdown with whitespace outside markers
    parts = []
    for (bold, italic, strike), text in groups:
        if not text:
            continue
        if bold or italic:
            stripped = text.strip()
            if not stripped:
                parts.append(text)
                continue
            leading = text[:len(text) - len(text.lstrip())]
            trailing = text[len(text.rstrip()):]
            if bold and italic:
                inner = f'***{stripped}***'
            elif bold:
                inner = f'**{stripped}**'
            else:
                inner = f'*{stripped}*'
            if strike:
                inner = f'~~{inner}~~'
            parts.append(f'{leading}{inner}{trailing}')
        elif strike:
            parts.append(f'~~{text}~~')
        else:
            parts.append(text)

    return ''.join(parts)


def get_run_text(run, hyperlinks):
    """Extract text from a w:r (run) element with formatting (legacy single-run helper)."""
    fmt, text = get_run_raw(run)
    if not text:
        return ''
    bold, italic, strike = fmt
    if bold and italic:
        text = f'***{text}***'
    elif bold:
        text = f'**{text}**'
    elif italic:
        text = f'*{text}*'
    if strike:
        text = f'~~{text}~~'
    return text


def get_heading_level(paragraph):
    """Determine if paragraph is a heading and return its level (0 = not a heading)."""
    ppr = paragraph.find('w:pPr', NS)
    if ppr is None:
        return 0
    pstyle = ppr.find('w:pStyle', NS)
    if pstyle is None:
        return 0
    style_val = pstyle.get(f'{{{NS["w"]}}}val', '')
    m = re.match(r'[Hh]eading\s*(\d+)', style_val)
    if m:
        return int(m.group(1))
    if style_val == 'Title':
        return 1
    if style_val == 'Subtitle':
        return 2
    return 0


def get_list_info(paragraph):
    """Check if paragraph is a list item. Returns (level, numId) or None."""
    ppr = paragraph.find('w:pPr', NS)
    if ppr is None:
        return None
    num_pr = ppr.find('w:numPr', NS)
    if num_pr is None:
        return None
    ilvl = num_pr.find('w:ilvl', NS)
    num_id = num_pr.find('w:numId', NS)
    if ilvl is not None and num_id is not None:
        level = int(ilvl.get(f'{{{NS["w"]}}}val', '0'))
        nid = num_id.get(f'{{{NS["w"]}}}val', '0')
        return (level, nid)
    return None


def process_table(table, hyperlinks):
    """Convert a w:tbl element to markdown table."""
    rows = table.findall('.//w:tr', NS)
    if not rows:
        return ''

    md_rows = []
    for row in rows:
        cells = row.findall('w:tc', NS)
        cell_texts = []
        for cell in cells:
            paras = cell.findall('w:p', NS)
            cell_text_parts = []
            for p in paras:
                p_text = ''
                for child in p:
                    tag = _local_tag(child)
                    if tag == 'r':
                        p_text += get_run_text(child, hyperlinks)
                    elif tag == 'hyperlink':
                        rid = child.get(f'{{{NS["r"]}}}id', '')
                        link_text = ''.join(
                            get_run_text(r, hyperlinks)
                            for r in child.findall('w:r', NS)
                        )
                        if rid in hyperlinks:
                            p_text += f'[{link_text}]({hyperlinks[rid]})'
                        else:
                            p_text += link_text
                if p_text.strip():
                    cell_text_parts.append(p_text.strip())
            cell_texts.append(' '.join(cell_text_parts))
        md_rows.append(cell_texts)

    if not md_rows:
        return ''

    max_cols = max(len(r) for r in md_rows)
    for row in md_rows:
        while len(row) < max_cols:
            row.append('')

    lines = []
    lines.append('| ' + ' | '.join(md_rows[0]) + ' |')
    lines.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
    for row in md_rows[1:]:
        lines.append('| ' + ' | '.join(row) + ' |')

    return '\n'.join(lines)


def process_paragraph(paragraph, hyperlinks):
    """Process a single w:p element. Returns (markdown_line, [comment_ids])."""
    heading_level = get_heading_level(paragraph)
    list_info = get_list_info(paragraph)

    # Collect comment IDs referenced in this paragraph
    comment_ids = []
    for el in paragraph.iter():
        tag = _local_tag(el)
        if tag == 'commentRangeStart':
            cid = el.get(f'{{{NS["w"]}}}id')
            if cid and cid not in comment_ids:
                comment_ids.append(cid)
        if tag == 'commentReference':
            cid = el.get(f'{{{NS["w"]}}}id')
            if cid and cid not in comment_ids:
                comment_ids.append(cid)

    # Build paragraph text using run merging for clean bold/italic
    raw_runs = []
    for child in paragraph:
        tag = _local_tag(child)
        if tag == 'r':
            raw_runs.append(get_run_raw(child))
        elif tag == 'hyperlink':
            rid = child.get(f'{{{NS["r"]}}}id', '')
            link_runs = child.findall('w:r', NS)
            link_text = ''.join(
                get_run_raw(r)[1] for r in link_runs
            )
            if rid in hyperlinks:
                # Insert link as a no-format run so merge_runs_to_md keeps it
                raw_runs.append(((False, False, False), f'[{link_text}]({hyperlinks[rid]})'))
            else:
                # Preserve formatting of hyperlink runs
                for r in link_runs:
                    raw_runs.append(get_run_raw(r))

    text = merge_runs_to_md(raw_runs).strip()

    if not text and not heading_level:
        return ('', comment_ids) if comment_ids else (None, [])

    if heading_level > 0:
        prefix = '#' * heading_level
        return (f'\n{prefix} {text}\n', comment_ids)

    if list_info:
        level, _ = list_info
        indent = '  ' * level
        return (f'{indent}- {text}', comment_ids)

    return (text, comment_ids)


def convert_docx_to_md(docx_path, start_from=None):
    """Main conversion: docx -> markdown string with comments."""
    with zipfile.ZipFile(docx_path) as zf:
        comments = parse_comments(zf)
        hyperlinks = parse_hyperlinks(zf)

        with zf.open('word/document.xml') as f:
            tree = ET.parse(f)

    body = tree.getroot().find('.//w:body', NS)
    if body is None:
        return '# Error: No body found in document'

    output_lines = []
    started = start_from is None

    for element in body:
        tag = _local_tag(element)

        if tag == 'p':
            line, comment_ids = process_paragraph(element, hyperlinks)
            if line is not None:
                if not started and start_from:
                    if start_from.lower() in line.lower():
                        started = True
                if not started:
                    continue
                output_lines.append(line)
                if comment_ids:
                    comment_block = []
                    for cid in comment_ids:
                        if cid in comments and comments[cid]['text']:
                            c = comments[cid]
                            date_str = f" ({c['date']})" if c['date'] else ''
                            comment_block.append(
                                f"> 💬 **{c['author']}**{date_str}: {c['text']}"
                            )
                    if comment_block:
                        output_lines.append('')
                        output_lines.extend(comment_block)
                        output_lines.append('')

        elif tag == 'tbl':
            if not started:
                continue
            table_md = process_table(element, hyperlinks)
            if table_md:
                output_lines.append('')
                output_lines.append(table_md)
                output_lines.append('')

        elif tag == 'sdt':
            if not started:
                continue
            sdt_content = element.find('.//w:sdtContent', NS)
            if sdt_content is not None:
                for child in sdt_content:
                    child_tag = _local_tag(child)
                    if child_tag == 'p':
                        line, comment_ids = process_paragraph(child, hyperlinks)
                        if line is not None:
                            output_lines.append(line)

    result = '\n'.join(output_lines)
    result = re.sub(r'\n{4,}', '\n\n\n', result)
    result = '\n'.join(line.rstrip() for line in result.split('\n'))

    return result


def main():
    if len(sys.argv) < 2:
        print("Usage: python docx2md_v2.py <input.docx> [output.md] [--start-from 'heading text']")
        print("  If output.md is omitted, generates from input filename.")
        print("  --start-from: Skip content before this heading/text (inclusive)")
        sys.exit(1)

    args = sys.argv[1:]
    start_from = None
    if '--start-from' in args:
        idx = args.index('--start-from')
        if idx + 1 < len(args):
            start_from = args[idx + 1]
            args = args[:idx] + args[idx + 2:]
        else:
            print("Error: --start-from requires a value")
            sys.exit(1)

    input_path = args[0]
    if not os.path.exists(input_path):
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    if len(args) >= 2:
        output_path = args[1]
    else:
        stem = Path(input_path).stem
        clean_name = re.sub(r'[^\w\s-]', '', stem).strip()
        clean_name = re.sub(r'\s+', '-', clean_name)
        output_path = f'{clean_name}.md'

    print(f"Converting: {input_path}")
    print(f"Output:     {output_path}")
    if start_from:
        print(f"Start from: '{start_from}'")

    md_content = convert_docx_to_md(input_path, start_from=start_from)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(md_content)

    lines = md_content.count('\n')
    comment_count = md_content.count('> 💬')
    print(f"Done! {lines} lines, {comment_count} comments extracted.")


if __name__ == '__main__':
    main()
