#!/usr/bin/env python3
"""Validate Markdown documents against the write-vscode-markdown profile."""

from __future__ import annotations

import argparse
import html
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote, urlsplit


ATX_HEADING_RE = re.compile(
    r"^ {0,3}(?P<marks>#{1,6})(?:[ \t]+|$)(?P<text>.*?)(?:[ \t]+#+[ \t]*)?$"
)
FENCE_OPEN_RE = re.compile(r"^ {0,3}(?P<marker>`{3,}|~{3,})(?P<info>.*)$")
SETEXT_RE = re.compile(r"^ {0,3}(?:=+|-+)[ \t]*$")
LINK_RE = re.compile(r"(?<!!)\[[^\]]+\]\(\s*(?P<target><#[^>]+>|#[^)\s]+)")
REFERENCE_IMAGE_RE = re.compile(r"!\[[^\]]*\](?!\s*\()")
RAW_TAG_START_RE = re.compile(
    r"</?[A-Za-z][A-Za-z0-9-]*(?=\s|/?>|$)", re.IGNORECASE
)
NUMBER_PREFIX_RE = re.compile(r"^\d+(?:\.\d+){0,2}\.\s+")

MERMAID_HEADER_RE = re.compile(
    r"^(?:"
    r"(?:flowchart|graph)\s+(?:TB|TD|BT|RL|LR)\b|"
    r"sequenceDiagram\b|stateDiagram(?:-v2)?\b|classDiagram\b|"
    r"erDiagram\b|gantt\b|journey\b|requirementDiagram\b|"
    r"pie\b|quadrantChart\b|mindmap\b|timeline\b|gitGraph\b|"
    r"C4(?:Context|Container|Component|Dynamic|Deployment)\b"
    r")",
    re.IGNORECASE,
)


@dataclass(frozen=True, order=True)
class Diagnostic:
    path_text: str
    line: int
    code: str
    message: str

    def render(self) -> str:
        return f"{self.path_text}:{self.line}:{self.code}: {self.message}"


@dataclass(frozen=True)
class Heading:
    level: int
    text: str
    line: int
    slug: str


@dataclass(frozen=True)
class FenceBlock:
    info: str
    start_line: int
    end_line: int | None
    content: tuple[tuple[int, str], ...]


@dataclass(frozen=True)
class ScanResult:
    visible_lines: tuple[str, ...]
    code_lines: frozenset[int]
    fences: tuple[FenceBlock, ...]
    diagnostics: tuple[Diagnostic, ...]


def display_path(path: Path) -> str:
    try:
        return path.relative_to(Path.cwd()).as_posix()
    except ValueError:
        return path.as_posix()


def diagnostic(path: Path, line: int, code: str, message: str) -> Diagnostic:
    return Diagnostic(display_path(path), max(line, 1), code, message)


def scan_blocks(path: Path, lines: list[str]) -> ScanResult:
    code_lines: set[int] = set()
    fences: list[FenceBlock] = []
    diagnostics: list[Diagnostic] = []
    current: dict[str, object] | None = None

    for line_number, line in enumerate(lines, start=1):
        if current is not None:
            code_lines.add(line_number)
            marker = str(current["marker"])
            closing = re.match(
                rf"^ {{0,3}}{re.escape(marker[0])}{{{len(marker)},}}[ \t]*$", line
            )
            if closing:
                fences.append(
                    FenceBlock(
                        info=str(current["info"]),
                        start_line=int(current["start_line"]),
                        end_line=line_number,
                        content=tuple(current["content"]),  # type: ignore[arg-type]
                    )
                )
                current = None
            else:
                current["content"].append((line_number, line))  # type: ignore[union-attr]
            continue

        match = FENCE_OPEN_RE.match(line)
        if not match:
            continue
        marker = match.group("marker")
        info = match.group("info").strip()
        if marker.startswith("`") and "`" in info:
            continue
        code_lines.add(line_number)
        current = {
            "marker": marker,
            "info": info,
            "start_line": line_number,
            "content": [],
        }

    if current is not None:
        start_line = int(current["start_line"])
        fences.append(
            FenceBlock(
                info=str(current["info"]),
                start_line=start_line,
                end_line=None,
                content=tuple(current["content"]),  # type: ignore[arg-type]
            )
        )
        diagnostics.append(
            diagnostic(path, start_line, "FENCE001", "fenced code block is not closed")
        )

    visible_lines: list[str] = []
    in_comment = False
    comment_start = 1
    for line_number, original in enumerate(lines, start=1):
        if line_number in code_lines:
            visible_lines.append(original)
            continue

        chars = list(original)
        cursor = 0
        while cursor < len(chars):
            if in_comment:
                end = original.find("-->", cursor)
                if end < 0:
                    for index in range(cursor, len(chars)):
                        chars[index] = " "
                    cursor = len(chars)
                else:
                    for index in range(cursor, end + 3):
                        chars[index] = " "
                    cursor = end + 3
                    in_comment = False
            else:
                start = original.find("<!--", cursor)
                if start < 0:
                    break
                end = original.find("-->", start + 4)
                if end < 0:
                    for index in range(start, len(chars)):
                        chars[index] = " "
                    in_comment = True
                    comment_start = line_number
                    cursor = len(chars)
                else:
                    for index in range(start, end + 3):
                        chars[index] = " "
                    cursor = end + 3
        visible_lines.append("".join(chars))

    if in_comment:
        diagnostics.append(
            diagnostic(path, comment_start, "HTML001", "HTML comment is not closed")
        )

    return ScanResult(
        visible_lines=tuple(visible_lines),
        code_lines=frozenset(code_lines),
        fences=tuple(fences),
        diagnostics=tuple(diagnostics),
    )


def mask_inline_code(line: str) -> str:
    chars = list(line)
    cursor = 0
    while cursor < len(line):
        if line[cursor] != "`":
            cursor += 1
            continue
        end_of_run = cursor
        while end_of_run < len(line) and line[end_of_run] == "`":
            end_of_run += 1
        marker = line[cursor:end_of_run]
        closing = line.find(marker, end_of_run)
        if closing < 0:
            cursor = end_of_run
            continue
        for index in range(cursor, closing + len(marker)):
            chars[index] = " "
        cursor = closing + len(marker)
    return "".join(chars)


def heading_slug(text: str) -> str:
    value = html.unescape(text.strip())
    value = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", value)
    value = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", value)
    value = re.sub(r"`+([^`]*)`+", r"\1", value)
    value = re.sub(r"<[^>]+>", "", value)
    value = value.lower()

    result: list[str] = []
    for character in value:
        if character.isspace():
            result.append("-")
            continue
        if character in "-_":
            result.append(character)
            continue
        category = unicodedata.category(character)
        if category[0] in {"L", "N", "M"}:
            result.append(character)
        elif ord(character) >= 128 and category[0] == "S":
            result.append(character)
    return "".join(result).strip("-")


def collect_headings(
    path: Path, scan: ScanResult
) -> tuple[list[Heading], list[Diagnostic]]:
    headings: list[Heading] = []
    diagnostics: list[Diagnostic] = []
    lines = scan.visible_lines

    for index, line in enumerate(lines):
        line_number = index + 1
        if line_number in scan.code_lines:
            continue
        match = ATX_HEADING_RE.match(line)
        if match:
            text = match.group("text").strip()
            headings.append(
                Heading(
                    level=len(match.group("marks")),
                    text=text,
                    line=line_number,
                    slug=heading_slug(text),
                )
            )
            continue

        if index == 0 or not SETEXT_RE.match(line):
            continue
        previous_line_number = line_number - 1
        previous = lines[index - 1]
        if (
            previous_line_number not in scan.code_lines
            and previous.strip()
            and not previous.lstrip().startswith((">", "<!--"))
        ):
            diagnostics.append(
                diagnostic(
                    path,
                    previous_line_number,
                    "HEAD001",
                    "Setext heading is unsupported; use an ATX heading beginning with #",
                )
            )

    return headings, diagnostics


def normalized_heading_label(heading: Heading) -> str:
    text = NUMBER_PREFIX_RE.sub("", heading.text).strip()
    return re.sub(r"\s+", " ", text).casefold()


def validate_heading_structure(
    path: Path, headings: list[Heading]
) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    h1_headings = [heading for heading in headings if heading.level == 1]
    if len(h1_headings) != 1:
        line = h1_headings[1].line if len(h1_headings) > 1 else 1
        diagnostics.append(
            diagnostic(path, line, "HEAD002", "document must contain exactly one H1 title")
        )
    elif NUMBER_PREFIX_RE.match(h1_headings[0].text):
        diagnostics.append(
            diagnostic(path, h1_headings[0].line, "HEAD003", "H1 title must be unnumbered")
        )

    if headings and headings[0].level != 1:
        diagnostics.append(
            diagnostic(path, headings[0].line, "HEAD004", "the first heading must be the H1 title")
        )

    toc_headings = [
        heading
        for heading in headings
        if heading.level == 2 and heading.text.strip() == "目次"
    ]
    if len(toc_headings) != 1:
        line = toc_headings[1].line if len(toc_headings) > 1 else 1
        diagnostics.append(
            diagnostic(path, line, "HEAD005", "document must contain exactly one unnumbered '## 目次'")
        )

    seen_labels: dict[str, Heading] = {}
    seen_slugs: dict[str, Heading] = {}
    for heading in headings:
        label = normalized_heading_label(heading)
        if label in seen_labels:
            diagnostics.append(
                diagnostic(
                    path,
                    heading.line,
                    "HEAD006",
                    f"duplicate heading text also used on line {seen_labels[label].line}",
                )
            )
        else:
            seen_labels[label] = heading

        if not heading.slug:
            diagnostics.append(
                diagnostic(path, heading.line, "HEAD007", "heading produces an empty anchor")
            )
        elif heading.slug in seen_slugs:
            diagnostics.append(
                diagnostic(
                    path,
                    heading.line,
                    "HEAD008",
                    f"heading anchor collides with line {seen_slugs[heading.slug].line}",
                )
            )
        else:
            seen_slugs[heading.slug] = heading

    expected_h2 = 1
    current_h2: int | None = None
    expected_h3 = 1
    current_h3: int | None = None
    expected_h4 = 1
    first_body_line: int | None = None

    for heading in headings:
        if heading.level == 1:
            continue
        if heading.level > 4:
            diagnostics.append(
                diagnostic(path, heading.line, "HEAD009", "use no heading deeper than H4")
            )
            continue
        if heading.level == 2 and heading.text.strip() == "目次":
            continue

        if heading.level == 2:
            match = re.match(r"^(\d+)\.\s+(.+)$", heading.text)
            if not match:
                diagnostics.append(
                    diagnostic(path, heading.line, "NUM001", "H2 must use 'N. Heading' numbering")
                )
                current_h2 = None
                current_h3 = None
                continue
            value = int(match.group(1))
            if value != expected_h2:
                diagnostics.append(
                    diagnostic(
                        path,
                        heading.line,
                        "NUM002",
                        f"expected H2 number {expected_h2}, found {value}",
                    )
                )
            expected_h2 = value + 1
            current_h2 = value
            expected_h3 = 1
            current_h3 = None
            expected_h4 = 1
            first_body_line = first_body_line or heading.line
            continue

        if heading.level == 3:
            match = re.match(r"^(\d+)\.(\d+)\.\s+(.+)$", heading.text)
            if not match:
                diagnostics.append(
                    diagnostic(path, heading.line, "NUM003", "H3 must use 'N.N. Heading' numbering")
                )
                current_h3 = None
                continue
            parent, value = int(match.group(1)), int(match.group(2))
            if current_h2 is None:
                diagnostics.append(
                    diagnostic(path, heading.line, "NUM004", "H3 requires a preceding numbered H2")
                )
            elif parent != current_h2:
                diagnostics.append(
                    diagnostic(
                        path,
                        heading.line,
                        "NUM005",
                        f"H3 parent number must be {current_h2}, found {parent}",
                    )
                )
            if value != expected_h3:
                diagnostics.append(
                    diagnostic(
                        path,
                        heading.line,
                        "NUM006",
                        f"expected H3 child number {expected_h3}, found {value}",
                    )
                )
            expected_h3 = value + 1
            current_h3 = value
            expected_h4 = 1
            continue

        match = re.match(r"^(\d+)\.(\d+)\.(\d+)\.\s+(.+)$", heading.text)
        if not match:
            diagnostics.append(
                diagnostic(path, heading.line, "NUM007", "H4 must use 'N.N.N. Heading' numbering")
            )
            continue
        parent, child, value = map(int, match.group(1, 2, 3))
        if current_h2 is None or current_h3 is None:
            diagnostics.append(
                diagnostic(path, heading.line, "NUM008", "H4 requires a preceding numbered H3")
            )
        else:
            if parent != current_h2 or child != current_h3:
                diagnostics.append(
                    diagnostic(
                        path,
                        heading.line,
                        "NUM009",
                        f"H4 parent number must be {current_h2}.{current_h3}",
                    )
                )
        if value != expected_h4:
            diagnostics.append(
                diagnostic(
                    path,
                    heading.line,
                    "NUM010",
                    f"expected H4 child number {expected_h4}, found {value}",
                )
            )
        expected_h4 = value + 1

    if toc_headings and first_body_line is not None and toc_headings[0].line > first_body_line:
        diagnostics.append(
            diagnostic(path, toc_headings[0].line, "HEAD010", "table of contents must precede body sections")
        )

    return diagnostics


def validate_toc(
    path: Path, scan: ScanResult, headings: list[Heading]
) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    toc_headings = [
        heading
        for heading in headings
        if heading.level == 2 and heading.text.strip() == "目次"
    ]
    if len(toc_headings) != 1:
        return diagnostics

    toc = toc_headings[0]
    end_line = len(scan.visible_lines) + 1
    for heading in headings:
        if heading.line > toc.line and heading.level <= 2:
            end_line = heading.line
            break

    anchors = {heading.slug: heading for heading in headings if heading.slug}
    seen_targets: dict[str, int] = {}
    for line_number in range(toc.line + 1, end_line):
        if line_number in scan.code_lines:
            continue
        line = mask_inline_code(scan.visible_lines[line_number - 1])
        for match in LINK_RE.finditer(line):
            target = match.group("target").strip("<>")
            slug = unquote(target[1:])
            if slug not in anchors:
                diagnostics.append(
                    diagnostic(path, line_number, "TOC001", f"table-of-contents target '#{slug}' does not exist")
                )
            if slug in seen_targets:
                diagnostics.append(
                    diagnostic(
                        path,
                        line_number,
                        "TOC002",
                        f"table-of-contents target '#{slug}' duplicates line {seen_targets[slug]}",
                    )
                )
            else:
                seen_targets[slug] = line_number

    required = [
        heading
        for heading in headings
        if heading.line > toc.line
        and heading.level in {2, 3}
        and not (heading.level == 2 and heading.text.strip() == "目次")
    ]
    for heading in required:
        if heading.slug and heading.slug not in seen_targets:
            diagnostics.append(
                diagnostic(
                    path,
                    heading.line,
                    "TOC003",
                    f"heading is missing from the table of contents: '#{heading.slug}'",
                )
            )
    return diagnostics


def validate_raw_html(path: Path, scan: ScanResult) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    for line_number, original in enumerate(scan.visible_lines, start=1):
        if line_number in scan.code_lines:
            continue
        line = mask_inline_code(original)
        declaration = re.search(r"<![A-Za-z][^>]*>", line)
        tag = RAW_TAG_START_RE.search(line)
        match = declaration or tag
        if not match:
            continue
        value = match.group(0)
        code = "HTML003" if re.match(r"</?svg\b", value, re.IGNORECASE) else "HTML002"
        message = "inline SVG is prohibited; fence the source or reference an external SVG file"
        if code == "HTML002":
            message = "raw HTML or XML markup is prohibited outside fenced code"
        diagnostics.append(diagnostic(path, line_number, code, message))
    return diagnostics


def validate_mermaid(path: Path, scan: ScanResult) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    for block in scan.fences:
        language = block.info.split(maxsplit=1)[0].casefold() if block.info else ""
        if language != "mermaid" or block.end_line is None:
            continue

        meaningful = [(line_no, text) for line_no, text in block.content if text.strip()]
        if not meaningful:
            diagnostics.append(
                diagnostic(path, block.start_line, "MER001", "Mermaid block is empty")
            )
            continue
        first_line_number, first_line = meaningful[0]
        if not MERMAID_HEADER_RE.match(first_line.strip()):
            diagnostics.append(
                diagnostic(
                    path,
                    first_line_number,
                    "MER002",
                    "Mermaid block must begin with a supported diagram declaration",
                )
            )

        has_acc_title = False
        has_acc_descr = False
        subgraphs = 0
        subgraph_ends = 0
        is_flowchart = bool(re.match(r"^(?:flowchart|graph)\b", first_line.strip(), re.IGNORECASE))
        for line_number, text in block.content:
            stripped = text.strip()
            if re.match(r"^accTitle\s*:", stripped, re.IGNORECASE):
                has_acc_title = True
            if re.match(r"^accDescr\s*:", stripped, re.IGNORECASE):
                has_acc_descr = True
            if re.match(r"^click\b", stripped, re.IGNORECASE):
                diagnostics.append(
                    diagnostic(path, line_number, "MER003", "Mermaid click directives are prohibited")
                )
            if re.match(r"^%%\s*\{", stripped):
                diagnostics.append(
                    diagnostic(path, line_number, "MER004", "Mermaid init/config directives are prohibited")
                )
            if re.search(r"javascript\s*:", stripped, re.IGNORECASE):
                diagnostics.append(
                    diagnostic(path, line_number, "MER005", "javascript URLs are prohibited in Mermaid")
                )
            if RAW_TAG_START_RE.search(stripped) or re.search(r"<![A-Za-z]", stripped):
                diagnostics.append(
                    diagnostic(path, line_number, "MER006", "HTML/XML markup is prohibited in Mermaid")
                )
            if is_flowchart and re.match(r"^subgraph\b", stripped, re.IGNORECASE):
                subgraphs += 1
            if is_flowchart and re.match(r"^end\s*$", stripped, re.IGNORECASE):
                subgraph_ends += 1

        if not has_acc_title:
            diagnostics.append(
                diagnostic(path, block.start_line, "MER007", "Mermaid block requires accTitle")
            )
        if not has_acc_descr:
            diagnostics.append(
                diagnostic(path, block.start_line, "MER008", "Mermaid block requires accDescr")
            )
        if is_flowchart and subgraphs != subgraph_ends:
            diagnostics.append(
                diagnostic(
                    path,
                    block.start_line,
                    "MER009",
                    f"flowchart has {subgraphs} subgraph declaration(s) but {subgraph_ends} matching end line(s)",
                )
            )
    return diagnostics


SVG_UNSAFE_PATTERNS = (
    (re.compile(r"<\s*script\b", re.IGNORECASE), "script element"),
    (re.compile(r"<\s*foreignObject\b", re.IGNORECASE), "foreignObject element"),
    (re.compile(r"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE), "DTD or entity declaration"),
    (re.compile(r"\son[a-z]+\s*=", re.IGNORECASE), "event handler attribute"),
    (
        re.compile(
            r"(?:href|xlink:href|src)\s*=\s*['\"]\s*(?:https?:|//|data:|javascript:)",
            re.IGNORECASE,
        ),
        "external or executable resource reference",
    ),
    (re.compile(r"url\(\s*['\"]?\s*(?:https?:|//|data:)", re.IGNORECASE), "external CSS resource"),
)


def unsafe_svg_reason(path: Path, cache: dict[Path, str | None]) -> str | None:
    if path in cache:
        return cache[path]
    try:
        source = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        cache[path] = f"cannot read SVG safely: {exc.__class__.__name__}"
        return cache[path]
    decoded_source = source
    for _ in range(3):
        unescaped = html.unescape(decoded_source)
        if unescaped == decoded_source:
            break
        decoded_source = unescaped
    for pattern, reason in SVG_UNSAFE_PATTERNS:
        if pattern.search(decoded_source):
            cache[path] = reason
            return reason
    cache[path] = None
    return None


def extract_inline_image_targets(line: str) -> list[str]:
    """Extract inline image destinations while preserving balanced parentheses."""
    targets: list[str] = []
    cursor = 0
    while cursor < len(line):
        start = line.find("![", cursor)
        if start < 0:
            break

        index = start + 2
        bracket_depth = 1
        while index < len(line) and bracket_depth:
            if line[index] == "\\":
                index += 2
                continue
            if line[index] == "[":
                bracket_depth += 1
            elif line[index] == "]":
                bracket_depth -= 1
            index += 1
        if bracket_depth:
            break

        while index < len(line) and line[index].isspace():
            index += 1
        if index >= len(line) or line[index] != "(":
            cursor = index
            continue
        index += 1
        while index < len(line) and line[index].isspace():
            index += 1

        if index < len(line) and line[index] == "<":
            index += 1
            destination_start = index
            while index < len(line):
                if line[index] == "\\":
                    index += 2
                    continue
                if line[index] == ">":
                    targets.append(line[destination_start:index])
                    index += 1
                    break
                index += 1
            cursor = index
            continue

        destination_start = index
        parenthesis_depth = 0
        while index < len(line):
            character = line[index]
            if character == "\\":
                index += 2
                continue
            if character == "(":
                parenthesis_depth += 1
            elif character == ")":
                if parenthesis_depth == 0:
                    targets.append(line[destination_start:index])
                    index += 1
                    break
                parenthesis_depth -= 1
            elif character.isspace() and parenthesis_depth == 0:
                targets.append(line[destination_start:index])
                break
            index += 1
        cursor = max(index, start + 2)
    return [target for target in targets if target]


def validate_image_target(
    path: Path,
    line_number: int,
    target: str,
    svg_cache: dict[Path, str | None],
) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    target = target.replace(r"\(", "(").replace(r"\)", ")")
    if re.match(r"^[A-Za-z]:[\\/]", target) or target.startswith(("/", "\\")):
        diagnostics.append(
            diagnostic(path, line_number, "IMG001", f"image path must be relative: '{target}'")
        )
        return diagnostics
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or target.startswith("//"):
        diagnostics.append(
            diagnostic(path, line_number, "IMG002", f"remote or URI image is prohibited: '{target}'")
        )
        return diagnostics
    relative_path = unquote(parsed.path)
    asset = (path.parent / relative_path).resolve()
    if not asset.is_file():
        diagnostics.append(
            diagnostic(path, line_number, "IMG003", f"referenced image does not exist: '{relative_path}'")
        )
        return diagnostics
    if asset.suffix.casefold() == ".svg":
        reason = unsafe_svg_reason(asset, svg_cache)
        if reason:
            diagnostics.append(
                diagnostic(path, line_number, "SVG001", f"referenced SVG is unsafe: {reason}")
            )
    return diagnostics


def validate_images(path: Path, scan: ScanResult) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    svg_cache: dict[Path, str | None] = {}
    for line_number, original in enumerate(scan.visible_lines, start=1):
        if line_number in scan.code_lines:
            continue
        line = mask_inline_code(original)
        if REFERENCE_IMAGE_RE.search(line):
            diagnostics.append(
                diagnostic(
                    path,
                    line_number,
                    "IMG004",
                    "reference-style images are unsupported; use an inline relative destination",
                )
            )
        for target in extract_inline_image_targets(line):
            diagnostics.extend(validate_image_target(path, line_number, target, svg_cache))
    return diagnostics


def validate_file(path: Path) -> list[Diagnostic]:
    if path.is_symlink():
        return [diagnostic(path, 1, "INPUT001", "symbolic-link Markdown files are unsupported")]
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return [diagnostic(path, 1, "INPUT002", "file is not valid UTF-8")]
    except OSError as exc:
        return [diagnostic(path, 1, "INPUT003", f"cannot read file: {exc.__class__.__name__}")]

    lines = text.splitlines()
    scan = scan_blocks(path, lines)
    headings, heading_diagnostics = collect_headings(path, scan)
    diagnostics = list(scan.diagnostics)
    diagnostics.extend(heading_diagnostics)
    diagnostics.extend(validate_heading_structure(path, headings))
    diagnostics.extend(validate_toc(path, scan, headings))
    diagnostics.extend(validate_raw_html(path, scan))
    diagnostics.extend(validate_mermaid(path, scan))
    diagnostics.extend(validate_images(path, scan))
    return sorted(set(diagnostics))


def collect_targets(arguments: list[str]) -> tuple[list[Path], list[Diagnostic]]:
    files: set[Path] = set()
    diagnostics: list[Diagnostic] = []
    for argument in arguments:
        target = Path(argument)
        if not target.exists():
            diagnostics.append(diagnostic(target, 1, "INPUT004", "target does not exist"))
        elif target.is_dir():
            files.update(
                candidate
                for candidate in target.rglob("*.md")
                if candidate.is_file() or candidate.is_symlink()
            )
        elif target.suffix.casefold() != ".md":
            diagnostics.append(diagnostic(target, 1, "INPUT005", "target must be a Markdown (.md) file or directory"))
        else:
            files.add(target)
    if not files and not diagnostics:
        diagnostics.append(diagnostic(Path("."), 1, "INPUT006", "no Markdown files were found"))
    return sorted(files, key=lambda item: item.as_posix()), diagnostics


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate VS Code-targeted Markdown structure, links, Mermaid, HTML, and image paths."
    )
    parser.add_argument("targets", nargs="+", help="Markdown file(s) or directories to validate")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    files, diagnostics = collect_targets(args.targets)
    for path in files:
        diagnostics.extend(validate_file(path))
    diagnostics = sorted(set(diagnostics))
    if diagnostics:
        for item in diagnostics:
            print(item.render())
        print(f"Validation failed with {len(diagnostics)} error(s).")
        return 1
    print(f"Validated {len(files)} Markdown file(s): no errors.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
