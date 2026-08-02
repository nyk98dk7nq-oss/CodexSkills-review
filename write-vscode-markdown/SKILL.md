---
name: write-vscode-markdown
description: Create and revise Markdown documents optimized for VS Code preview, using a CommonMark base, VS Code-compatible pipe tables, numbered headings with a linked table of contents, and purpose-selected Mermaid diagrams. Use for `.md` documentation such as procedures, specifications, architecture notes, workflows, data models, and explanatory guides that must be structured, navigable, visually clear, and mechanically validated.
---

# Write VS Code Markdown

## Follow the workflow

1. Identify the document's purpose, audience, output path, and required depth.
2. Read [references/vscode-markdown-profile.md](references/vscode-markdown-profile.md) before creating or restructuring a document.
3. Read [references/mermaid-design-rules.md](references/mermaid-design-rules.md) when the document needs a diagram or chart.
4. Inspect existing content and referenced assets before editing. Preserve correct material unless the user requests a rewrite.
5. Outline the document, then create the title, linked table of contents, numbered headings, prose, tables, and diagrams.
6. Prefer prose for simple explanations, pipe tables for exact comparisons, and Mermaid only when relationships, sequence, state, structure, or responsibility become clearer visually.
7. Write XML and SVG source as fenced code. Reference a displayable SVG as an external file. Do not emit raw HTML or inline SVG.
8. Run the validator and fix every reported error:

   ```bash
   python3 <skill-dir>/scripts/validate_markdown.py <document.md>
   ```

9. Preview the result in VS Code when available. Treat validator success as structural validation, not proof that every Mermaid diagram renders correctly.
10. Re-run validation after every heading, link, asset-path, table, or Mermaid change.

## Review at the right level

Complete normal work with the main agent and `validate_markdown.py`.

For a complex, important, or large document, ask a subagent for an independent content review after mechanical validation passes, unless the user opts out or subagents are unavailable. Give the reviewer only the completed artifact and the minimum task-local context needed to identify its audience and purpose. Do not disclose intended findings, suspected defects, or proposed fixes. Ask it to assess organization, clarity, missing context, diagram choice, visual readability, and consistency between prose and diagrams. Do not use a subagent merely to recheck heading numbers, links, fences, prohibited HTML, or file paths.

Have the main agent evaluate the findings, make all accepted fixes, then validate and preview again.

## Deliver the document

Confirm that:

- the H1 title and `## 目次` heading are unnumbered;
- every other H2 through H4 heading is numbered consistently;
- table-of-contents links resolve in VS Code;
- tables remain readable without raw HTML;
- Mermaid diagrams use an appropriate diagram type and add information;
- SVG images use relative file references and their files exist;
- the validator completes successfully.

Return the completed Markdown file and any external SVG assets it references. Mention any preview limitation that could not be checked.
