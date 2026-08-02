# VS Code Markdown profile

Apply these rules to every generated document. Treat the profile as CommonMark core plus two VS Code-compatible extensions: pipe tables and Mermaid fenced blocks.

## Document structure

Use this order unless the user's template requires another:

1. one unnumbered H1 title;
2. one unnumbered `## 目次` heading;
3. a linked table of contents;
4. numbered body sections;
5. optional references or appendices, numbered as body sections.

Use exactly one unnumbered H1 title. Leave only the `## 目次` heading unnumbered; number every other H2, H3, and H4 by depth:

| Level | Format | Example |
|---|---|---|
| H1 | Unnumbered | `# 運用手順書` |
| TOC | Unnumbered H2 | `## 目次` |
| H2 | `N.` | `## 1. 目的` |
| H3 | `N.N.` | `### 1.1. 背景` |
| H4 | `N.N.N.` | `#### 1.1.1. 対象範囲` |

Do not use H5 or H6 by default. Do not skip heading levels. Keep numbering sequential within each parent section. Renumber the affected headings and links together after inserting, moving, or removing a section.

## Linked table of contents

List all H2 and H3 body headings. Add H4 only when it materially improves navigation in a long document. Mirror heading hierarchy with nested lists.

Use VS Code's generated heading fragments; do not add raw HTML anchors. For example:

```markdown
# 運用手順書

## 目次

1. [1. 目的](#1-目的)
   1. [1.1. 背景](#11-背景)
2. [2. 処理フロー](#2-処理フロー)

## 1. 目的

### 1.1. 背景

## 2. 処理フロー
```

Use unique heading text within the document. Avoid symbols and decorative punctuation in headings because they produce fragile fragments. Validate every table-of-contents target after changing a heading.

## Prose and lists

Use CommonMark syntax for paragraphs, emphasis, links, lists, block quotes, code spans, and fenced code. Keep paragraphs focused. Introduce a list with a complete sentence and a blank line. Use ordered lists for sequence and unordered lists for unordered sets.

Do not use raw HTML for spacing, layout, anchors, collapsible sections, or line breaks. Prefer native Markdown structure.

## Tables

Use VS Code-compatible pipe tables for exact mappings, comparisons, and compact structured data. Do not use Mermaid to imitate a table.

```markdown
| 項目 | 内容 |
|---|---|
| 表示先 | VS Code |
| 図 | Mermaid |
```

Keep cells short. Move paragraphs, lists, and code blocks outside the table. Escape a literal pipe as `\|`. Use alignment markers only when alignment conveys meaning, such as right-aligning numeric columns.

## Code, XML, and SVG

Use a fenced code block with an accurate language identifier for source content:

````markdown
```xml
<employee>
  <name>山田太郎</name>
</employee>
```
````

Use `svg` or `xml` fencing to show SVG source as text. To display SVG graphics, save a separate `.svg` file and use a relative image reference:

```markdown
![システム構成図](./assets/system-architecture.svg)
```

Use inline image destinations as shown above; do not use reference-style image syntax. Do not generate inline `<svg>`, arbitrary XML as display markup, scripts, event handlers, `foreignObject`, or remote resources. Confirm that every relative image path resolves from the Markdown file.

## Compatibility checks

Keep VS Code preview as the primary rendering target. Use features outside this profile only when the user names another renderer and accepts reduced portability. Enable VS Code Markdown link validation when possible. After previewing, check heading links, table wrapping, code fencing, image paths, and Mermaid rendering.
