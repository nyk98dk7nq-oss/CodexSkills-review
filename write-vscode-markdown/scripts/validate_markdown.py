#!/usr/bin/env python3
"""Markdown文書をwrite-vscode-markdownプロファイルに照らして検証する。"""

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
REFERENCE_IMAGE_RE = re.compile(r"!\[[^\]]*\](?!\s*\()")
REFERENCE_LINK_DEFINITION_RE = re.compile(
    r"^ {0,3}\[[^\]]+\]:[ \t]*(?P<target><[^>]+>|\S+)"
)
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
    r")"
)
MARKDOWN_ESCAPE_RE = re.compile(r"\\([!\"#$%&'()*+,\-./:;<=>?@\[\\\]^_`{|}~])")
HTML_ENTITY_RE = re.compile(
    r"&(?:#[xX][0-9A-Fa-f]{1,8}|#[0-9]{1,8}|[A-Za-z][A-Za-z0-9]{1,31});"
)
TABLE_SEPARATOR_CELL_RE = re.compile(r"^:?-+:?$")


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
    base_slug: str


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
            diagnostic(path, start_line, "FENCE001", "コードブロックのフェンスが閉じられていません")
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
            diagnostic(path, comment_start, "HTML001", "HTMLコメントが閉じられていません")
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


def is_escaped(value: str, index: int) -> bool:
    backslashes = 0
    index -= 1
    while index >= 0 and value[index] == "\\":
        backslashes += 1
        index -= 1
    return backslashes % 2 == 1


def extract_inline_targets(line: str, *, images: bool) -> list[str]:
    """括弧の対応を保ちながらインライン形式のリンク先を抽出する。"""
    targets: list[str] = []
    cursor = 0
    while cursor < len(line):
        if images:
            start = line.find("![", cursor)
            if start < 0:
                break
            bracket = start + 1
            if is_escaped(line, start):
                cursor = bracket + 1
                continue
        else:
            bracket = line.find("[", cursor)
            if bracket < 0:
                break
            if (
                bracket > 0
                and line[bracket - 1] == "!"
                and not is_escaped(line, bracket - 1)
            ) or is_escaped(line, bracket):
                cursor = bracket + 1
                continue
            start = bracket

        index = bracket + 1
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
            cursor = max(index, start + 1)
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
        cursor = max(index, start + 1)
    return [target for target in targets if target]


def extract_inline_link_targets(line: str) -> list[str]:
    return extract_inline_targets(line, images=False)


def decode_url_path(value: str) -> str:
    if re.search(r"%(?![0-9A-Fa-f]{2})", value):
        raise ValueError("パーセントエンコードが不正です")
    try:
        return unquote(value, encoding="utf-8", errors="strict")
    except UnicodeDecodeError as exc:
        raise ValueError("URLをUTF-8としてデコードできません") from exc


def heading_slug(text: str) -> str:
    value = HTML_ENTITY_RE.sub(lambda match: html.unescape(match.group(0)), text.strip())
    value = MARKDOWN_ESCAPE_RE.sub(r"\1", value)
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


class HeadingSlugger:
    """VS Code/GitHub形式で文書内の見出しslugを一意にする。"""

    def __init__(self) -> None:
        self.used: set[str] = set()
        self.next_suffix: dict[str, int] = {}

    def slug(self, text: str) -> tuple[str, str]:
        base = heading_slug(text)
        if base not in self.used:
            self.used.add(base)
            self.next_suffix[base] = 1
            return base, base

        suffix = self.next_suffix.get(base, 1)
        candidate = f"{base}-{suffix}"
        while candidate in self.used:
            suffix += 1
            candidate = f"{base}-{suffix}"
        self.next_suffix[base] = suffix + 1
        self.used.add(candidate)
        return candidate, base


def collect_headings(
    path: Path, scan: ScanResult
) -> tuple[list[Heading], list[Diagnostic]]:
    headings: list[Heading] = []
    diagnostics: list[Diagnostic] = []
    lines = scan.visible_lines
    slugger = HeadingSlugger()

    for index, line in enumerate(lines):
        line_number = index + 1
        if line_number in scan.code_lines:
            continue
        match = ATX_HEADING_RE.match(line)
        if match:
            text = match.group("text").strip()
            slug, base_slug = slugger.slug(text)
            headings.append(
                Heading(
                    level=len(match.group("marks")),
                    text=text,
                    line=line_number,
                    slug=slug,
                    base_slug=base_slug,
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
                    "Setext形式の見出しは使用できません。#で始まるATX形式の見出しを使用してください",
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
            diagnostic(path, line, "HEAD002", "文書にはH1タイトルを1つだけ含めてください")
        )
    elif NUMBER_PREFIX_RE.match(h1_headings[0].text):
        diagnostics.append(
            diagnostic(path, h1_headings[0].line, "HEAD003", "H1タイトルには番号を付けないでください")
        )

    if headings and headings[0].level != 1:
        diagnostics.append(
            diagnostic(path, headings[0].line, "HEAD004", "最初の見出しはH1タイトルにしてください")
        )

    toc_headings = [
        heading
        for heading in headings
        if heading.level == 2 and heading.text.strip() == "目次"
    ]
    if len(toc_headings) != 1:
        line = toc_headings[1].line if len(toc_headings) > 1 else 1
        diagnostics.append(
            diagnostic(path, line, "HEAD005", "文書には番号なしの「## 目次」を1つだけ含めてください")
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
                    f"同じ見出し文字列が{seen_labels[label].line}行目にもあります",
                )
            )
        else:
            seen_labels[label] = heading

        if not heading.base_slug:
            diagnostics.append(
                diagnostic(path, heading.line, "HEAD007", "見出しからアンカーを生成できません")
            )
        elif heading.base_slug in seen_slugs:
            diagnostics.append(
                diagnostic(
                    path,
                    heading.line,
                    "HEAD008",
                    f"見出しアンカーが{seen_slugs[heading.base_slug].line}行目と重複しています",
                )
            )
        else:
            seen_slugs[heading.base_slug] = heading

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
                diagnostic(path, heading.line, "HEAD009", "H4より深い見出しは使用しないでください")
            )
            continue
        if heading.level == 2 and heading.text.strip() == "目次":
            continue

        if heading.level == 2:
            match = re.match(r"^(\d+)\.\s+(.+)$", heading.text)
            if not match:
                diagnostics.append(
                    diagnostic(path, heading.line, "NUM001", "H2は「N. 見出し」の形式で採番してください")
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
                        f"H2の番号は{expected_h2}を期待しましたが、{value}でした",
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
                    diagnostic(path, heading.line, "NUM003", "H3は「N.N. 見出し」の形式で採番してください")
                )
                current_h3 = None
                continue
            parent, value = int(match.group(1)), int(match.group(2))
            if current_h2 is None:
                diagnostics.append(
                    diagnostic(path, heading.line, "NUM004", "H3の前には番号付きH2が必要です")
                )
            elif parent != current_h2:
                diagnostics.append(
                    diagnostic(
                        path,
                        heading.line,
                        "NUM005",
                        f"H3の親番号は{current_h2}にしてください。現在は{parent}です",
                    )
                )
            if value != expected_h3:
                diagnostics.append(
                    diagnostic(
                        path,
                        heading.line,
                        "NUM006",
                        f"H3の子番号は{expected_h3}を期待しましたが、{value}でした",
                    )
                )
            expected_h3 = value + 1
            current_h3 = value
            expected_h4 = 1
            continue

        match = re.match(r"^(\d+)\.(\d+)\.(\d+)\.\s+(.+)$", heading.text)
        if not match:
            diagnostics.append(
                diagnostic(path, heading.line, "NUM007", "H4は「N.N.N. 見出し」の形式で採番してください")
            )
            continue
        parent, child, value = map(int, match.group(1, 2, 3))
        if current_h2 is None or current_h3 is None:
            diagnostics.append(
                diagnostic(path, heading.line, "NUM008", "H4の前には番号付きH3が必要です")
            )
        else:
            if parent != current_h2 or child != current_h3:
                diagnostics.append(
                    diagnostic(
                        path,
                        heading.line,
                        "NUM009",
                        f"H4の親番号は{current_h2}.{current_h3}にしてください",
                    )
                )
        if value != expected_h4:
            diagnostics.append(
                diagnostic(
                    path,
                    heading.line,
                    "NUM010",
                    f"H4の子番号は{expected_h4}を期待しましたが、{value}でした",
                )
            )
        expected_h4 = value + 1

    if toc_headings and first_body_line is not None and toc_headings[0].line > first_body_line:
        diagnostics.append(
            diagnostic(path, toc_headings[0].line, "HEAD010", "目次は本文セクションより前に配置してください")
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
        for target in extract_inline_link_targets(line):
            if not target.startswith("#"):
                continue
            try:
                slug = decode_url_path(target[1:])
            except ValueError as exc:
                diagnostics.append(
                    diagnostic(path, line_number, "TOC004", f"目次リンクのURLが不正です: {exc}")
                )
                continue
            if slug not in anchors:
                diagnostics.append(
                    diagnostic(path, line_number, "TOC001", f"目次リンクの対象「#{slug}」が存在しません")
                )
            if slug in seen_targets:
                diagnostics.append(
                    diagnostic(
                        path,
                        line_number,
                        "TOC002",
                        f"目次リンクの対象「#{slug}」は{seen_targets[slug]}行目と重複しています",
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
                    f"見出しが目次にありません:「#{heading.slug}」",
                )
            )
    return diagnostics


def toc_content_lines(headings: list[Heading], total_lines: int) -> set[int]:
    toc_headings = [
        heading
        for heading in headings
        if heading.level == 2 and heading.text.strip() == "目次"
    ]
    if len(toc_headings) != 1:
        return set()
    toc = toc_headings[0]
    end_line = total_lines + 1
    for heading in headings:
        if heading.line > toc.line and heading.level <= 2:
            end_line = heading.line
            break
    return set(range(toc.line + 1, end_line))


def anchors_for_markdown(
    target: Path, cache: dict[Path, set[str] | None]
) -> set[str] | None:
    resolved = target.resolve()
    if resolved in cache:
        return cache[resolved]
    try:
        text = resolved.read_text(encoding="utf-8")
    except (OSError, UnicodeError):
        cache[resolved] = None
        return None
    lines = text.splitlines()
    scan = scan_blocks(resolved, lines)
    headings, _ = collect_headings(resolved, scan)
    anchors = {heading.slug for heading in headings if heading.slug}
    cache[resolved] = anchors
    return anchors


def validate_link_target(
    path: Path,
    line_number: int,
    target: str,
    current_anchors: set[str],
    markdown_anchor_cache: dict[Path, set[str] | None],
    *,
    skip_current_anchor: bool = False,
) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    target = target.replace(r"\(", "(").replace(r"\)", ")")
    try:
        parsed = urlsplit(target)
    except ValueError:
        return [diagnostic(path, line_number, "LINK004", f"リンクURLが不正です:「{target}」")]

    if parsed.scheme or parsed.netloc or target.startswith("//"):
        return diagnostics
    if parsed.path.startswith(("/", "\\")) or re.match(r"^[A-Za-z]:[\\/]", parsed.path):
        return diagnostics

    try:
        relative_path = decode_url_path(parsed.path)
        fragment = decode_url_path(parsed.fragment) if parsed.fragment else ""
    except ValueError as exc:
        return [diagnostic(path, line_number, "LINK004", f"リンクURLが不正です:「{target}」 ({exc})")]

    if not relative_path:
        if fragment and not skip_current_anchor and fragment not in current_anchors:
            diagnostics.append(
                diagnostic(path, line_number, "LINK001", f"文書内リンクの対象「#{fragment}」が存在しません")
            )
        return diagnostics

    linked_path = (path.parent / relative_path).resolve()
    if not linked_path.exists():
        diagnostics.append(
            diagnostic(path, line_number, "LINK002", f"相対リンク先が存在しません:「{relative_path}」")
        )
        return diagnostics

    if fragment and linked_path.suffix.casefold() in {".md", ".markdown"}:
        anchors = anchors_for_markdown(linked_path, markdown_anchor_cache)
        if anchors is None or fragment not in anchors:
            diagnostics.append(
                diagnostic(
                    path,
                    line_number,
                    "LINK003",
                    f"リンク先Markdownにアンカー「#{fragment}」が存在しません:"
                    f"「{relative_path}」",
                )
            )
    return diagnostics


def validate_links(
    path: Path, scan: ScanResult, headings: list[Heading]
) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    current_anchors = {heading.slug for heading in headings if heading.slug}
    anchor_cache: dict[Path, set[str] | None] = {path.resolve(): current_anchors}
    toc_lines = toc_content_lines(headings, len(scan.visible_lines))

    for line_number, original in enumerate(scan.visible_lines, start=1):
        if line_number in scan.code_lines:
            continue
        line = mask_inline_code(original)
        targets = extract_inline_link_targets(line)
        definition = REFERENCE_LINK_DEFINITION_RE.match(line)
        if definition:
            targets.append(definition.group("target").strip("<>"))
        for target in targets:
            diagnostics.extend(
                validate_link_target(
                    path,
                    line_number,
                    target,
                    current_anchors,
                    anchor_cache,
                    skip_current_anchor=line_number in toc_lines and target.startswith("#"),
                )
            )
    return diagnostics


def split_pipe_row(line: str) -> list[str] | None:
    stripped = line.strip()
    if not stripped:
        return None
    cells: list[str] = []
    current: list[str] = []
    has_separator = False
    index = 0
    while index < len(stripped):
        character = stripped[index]
        if character == "\\" and index + 1 < len(stripped):
            current.extend((character, stripped[index + 1]))
            index += 2
            continue
        if character == "|":
            cells.append("".join(current).strip())
            current = []
            has_separator = True
        else:
            current.append(character)
        index += 1
    cells.append("".join(current).strip())
    if not has_separator:
        return None
    if stripped.startswith("|"):
        cells.pop(0)
    if stripped.endswith("|"):
        cells.pop()
    return cells


def is_table_separator(cells: list[str] | None) -> bool:
    return bool(cells) and all(TABLE_SEPARATOR_CELL_RE.fullmatch(cell) for cell in cells)


def validate_pipe_tables(path: Path, scan: ScanResult) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    lines = scan.visible_lines
    consumed: set[int] = set()
    for index in range(1, len(lines)):
        separator_line = index + 1
        if separator_line in scan.code_lines or separator_line in consumed:
            continue
        separator_cells = split_pipe_row(lines[index])
        if not is_table_separator(separator_cells):
            continue

        header_line = index
        if header_line in scan.code_lines:
            continue
        header_cells = split_pipe_row(lines[index - 1])
        if header_cells is None:
            continue
        consumed.update({header_line, separator_line})
        assert separator_cells is not None
        expected = len(header_cells)
        if len(separator_cells) != expected:
            diagnostics.append(
                diagnostic(
                    path,
                    separator_line,
                    "TABLE001",
                    f"表ヘッダーは{expected}列ですが、区切り行は{len(separator_cells)}列です",
                )
            )

        body_index = index + 1
        while body_index < len(lines):
            body_line = body_index + 1
            if body_line in scan.code_lines or not lines[body_index].strip():
                break
            body_cells = split_pipe_row(lines[body_index])
            if body_cells is None:
                break
            consumed.add(body_line)
            if len(body_cells) != expected:
                diagnostics.append(
                    diagnostic(
                        path,
                        body_line,
                        "TABLE002",
                        f"表ヘッダーは{expected}列ですが、この行は{len(body_cells)}列です",
                    )
                )
            body_index += 1
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
        message = "インラインSVGは使用できません。ソースをコードフェンスで囲むか、外部SVGファイルを参照してください"
        if code == "HTML002":
            message = "コードフェンス外では生HTMLまたはXMLマークアップを使用できません"
        diagnostics.append(diagnostic(path, line_number, code, message))
    return diagnostics


def validate_mermaid(path: Path, scan: ScanResult) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    for block in scan.fences:
        language = block.info.split(maxsplit=1)[0] if block.info else ""
        if language.casefold() == "mermaid" and language != "mermaid":
            diagnostics.append(
                diagnostic(
                    path,
                    block.start_line,
                    "MER010",
                    "Mermaidの言語識別子は小文字のmermaidにしてください",
                )
            )
            continue
        if language != "mermaid" or block.end_line is None:
            continue

        meaningful = [(line_no, text) for line_no, text in block.content if text.strip()]
        if not meaningful:
            diagnostics.append(
                diagnostic(path, block.start_line, "MER001", "Mermaidコードブロックが空です")
            )
            continue
        first_line_number, first_line = meaningful[0]
        if not MERMAID_HEADER_RE.match(first_line.strip()):
            diagnostics.append(
                diagnostic(
                    path,
                    first_line_number,
                    "MER002",
                    "Mermaidコードブロックの先頭には対応済みのダイアグラム宣言が必要です",
                )
            )

        has_acc_title = False
        has_acc_descr = False
        subgraphs = 0
        subgraph_ends = 0
        is_flowchart = bool(re.match(r"^(?:flowchart|graph)\b", first_line.strip()))
        for line_number, text in block.content:
            stripped = text.strip()
            if re.match(r"^accTitle\s*:", stripped, re.IGNORECASE):
                has_acc_title = True
            if re.match(r"^accDescr\s*:", stripped, re.IGNORECASE):
                has_acc_descr = True
            if re.match(r"^click\b", stripped, re.IGNORECASE):
                diagnostics.append(
                    diagnostic(path, line_number, "MER003", "Mermaidのclickディレクティブは使用できません")
                )
            if re.match(r"^%%\s*\{", stripped):
                diagnostics.append(
                    diagnostic(path, line_number, "MER004", "Mermaidのinit・configディレクティブは使用できません")
                )
            if re.search(r"javascript\s*:", stripped, re.IGNORECASE):
                diagnostics.append(
                    diagnostic(path, line_number, "MER005", "Mermaidではjavascript URLを使用できません")
                )
            if RAW_TAG_START_RE.search(stripped) or re.search(r"<![A-Za-z]", stripped):
                diagnostics.append(
                    diagnostic(path, line_number, "MER006", "MermaidではHTML・XMLマークアップを使用できません")
                )
            if is_flowchart and re.match(r"^subgraph\b", stripped, re.IGNORECASE):
                subgraphs += 1
            if is_flowchart and re.match(r"^end\s*$", stripped, re.IGNORECASE):
                subgraph_ends += 1

        if not has_acc_title:
            diagnostics.append(
                diagnostic(path, block.start_line, "MER007", "MermaidコードブロックにはaccTitleが必要です")
            )
        if not has_acc_descr:
            diagnostics.append(
                diagnostic(path, block.start_line, "MER008", "MermaidコードブロックにはaccDescrが必要です")
            )
        if is_flowchart and subgraphs != subgraph_ends:
            diagnostics.append(
                diagnostic(
                    path,
                    block.start_line,
                    "MER009",
                    f"flowchartのsubgraph宣言は{subgraphs}件ですが、対応するend行は{subgraph_ends}件です",
                )
            )
    return diagnostics


SVG_UNSAFE_PATTERNS = (
    (re.compile(r"<\s*script\b", re.IGNORECASE), "script要素"),
    (re.compile(r"<\s*foreignObject\b", re.IGNORECASE), "foreignObject要素"),
    (re.compile(r"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE), "DTDまたはエンティティ宣言"),
    (re.compile(r"\son[a-z]+\s*=", re.IGNORECASE), "イベントハンドラー属性"),
    (re.compile(r"@import\b", re.IGNORECASE), "CSSの@import参照"),
)
SVG_RESOURCE_ATTRIBUTE_RE = re.compile(
    r"\b(?:href|xlink:href|src)\s*=\s*(['\"])(.*?)\1", re.IGNORECASE | re.DOTALL
)
SVG_CSS_URL_RE = re.compile(
    r"url\(\s*(['\"]?)(.*?)\1\s*\)", re.IGNORECASE | re.DOTALL
)


def unsafe_svg_reason(path: Path, cache: dict[Path, str | None]) -> str | None:
    if path in cache:
        return cache[path]
    try:
        source = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        cache[path] = f"SVGを安全に読み取れません: {exc.__class__.__name__}"
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
    for match in SVG_RESOURCE_ATTRIBUTE_RE.finditer(decoded_source):
        reference = match.group(2).strip()
        if reference and not reference.startswith("#"):
            cache[path] = "外部リソース参照"
            return cache[path]
    for match in SVG_CSS_URL_RE.finditer(decoded_source):
        reference = match.group(2).strip()
        if not reference.startswith("#"):
            cache[path] = "外部CSSリソース"
            return cache[path]
    cache[path] = None
    return None


def extract_inline_image_targets(line: str) -> list[str]:
    return extract_inline_targets(line, images=True)


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
            diagnostic(path, line_number, "IMG001", f"画像パスは相対パスにしてください:「{target}」")
        )
        return diagnostics
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or target.startswith("//"):
        diagnostics.append(
            diagnostic(path, line_number, "IMG002", f"リモート画像またはURI画像は使用できません:「{target}」")
        )
        return diagnostics
    try:
        relative_path = decode_url_path(parsed.path)
    except ValueError as exc:
        diagnostics.append(
            diagnostic(path, line_number, "IMG006", f"画像URLが不正です:「{target}」 ({exc})")
        )
        return diagnostics
    document_directory = path.parent.resolve()
    asset = (document_directory / relative_path).resolve()
    if not asset.is_relative_to(document_directory):
        diagnostics.append(
            diagnostic(
                path,
                line_number,
                "IMG005",
                f"画像パスがMarkdown文書のディレクトリ外を指しています:「{relative_path}」",
            )
        )
        return diagnostics
    if not asset.is_file():
        diagnostics.append(
            diagnostic(path, line_number, "IMG003", f"参照先の画像が存在しません:「{relative_path}」")
        )
        return diagnostics
    if asset.suffix.casefold() == ".svg":
        reason = unsafe_svg_reason(asset, svg_cache)
        if reason:
            diagnostics.append(
                diagnostic(path, line_number, "SVG001", f"参照先のSVGは安全ではありません: {reason}")
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
                    "参照形式の画像には対応していません。インライン形式の相対パスを使用してください",
                )
            )
        for target in extract_inline_image_targets(line):
            diagnostics.extend(validate_image_target(path, line_number, target, svg_cache))
    return diagnostics


def validate_file(path: Path) -> list[Diagnostic]:
    if path.is_symlink():
        return [diagnostic(path, 1, "INPUT001", "シンボリックリンクのMarkdownファイルには対応していません")]
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return [diagnostic(path, 1, "INPUT002", "ファイルは有効なUTF-8ではありません")]
    except OSError as exc:
        return [diagnostic(path, 1, "INPUT003", f"ファイルを読み取れません: {exc.__class__.__name__}")]

    lines = text.splitlines()
    scan = scan_blocks(path, lines)
    headings, heading_diagnostics = collect_headings(path, scan)
    diagnostics = list(scan.diagnostics)
    diagnostics.extend(heading_diagnostics)
    diagnostics.extend(validate_heading_structure(path, headings))
    diagnostics.extend(validate_toc(path, scan, headings))
    diagnostics.extend(validate_links(path, scan, headings))
    diagnostics.extend(validate_pipe_tables(path, scan))
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
            diagnostics.append(diagnostic(target, 1, "INPUT004", "検証対象が存在しません"))
        elif target.is_dir():
            files.update(
                candidate
                for candidate in target.rglob("*.md")
                if candidate.is_file() or candidate.is_symlink()
            )
        elif target.suffix.casefold() != ".md":
            diagnostics.append(diagnostic(target, 1, "INPUT005", "検証対象にはMarkdown（.md）ファイルまたはディレクトリを指定してください"))
        else:
            files.add(target)
    if not files and not diagnostics:
        diagnostics.append(diagnostic(Path("."), 1, "INPUT006", "Markdownファイルが見つかりませんでした"))
    return sorted(files, key=lambda item: item.as_posix()), diagnostics


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="VS Code向けMarkdownの構造、リンク、Mermaid、HTML、画像パスを検証します。"
    )
    parser.add_argument("targets", nargs="+", help="検証するMarkdownファイルまたはディレクトリ")
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
        print(f"検証に失敗しました。エラーは{len(diagnostics)}件です。")
        return 1
    print(f"Markdownファイル{len(files)}件を検証しました。エラーはありません。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
