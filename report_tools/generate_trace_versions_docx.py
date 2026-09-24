#!/usr/bin/env python3
"""Generate a client-ready DOCX with clickable Before/After GitHub code snapshots.

This script creates a Word document where each comparison pair has:
- Before snapshot (captured directly from GitHub code page)
- After snapshot (captured directly from GitHub code page)
- Each image is clickable and opens the exact GitHub URL used for capture

Usage:
  python generate_trace_versions_docx.py --config trace_versions_docx_config.sample.json

Requirements:
  pip install python-docx playwright
  python -m playwright install chromium
"""

from __future__ import annotations

import argparse
import asyncio
import json
import re
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from docx import Document
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt, RGBColor


@dataclass
class SnapshotSpec:
    label: str
    github_url: str | None
    description: str
    highlight: str
    source: str
    repo_path: str | None
    file_path: str | None
    git_ref: str | None
    start_line: int | None
    end_line: int | None
    github_repo_base: str | None


@dataclass
class PairSpec:
    pair_title: str
    before: SnapshotSpec
    after: SnapshotSpec


@dataclass
class SectionAnalysis:
    additions: int
    deletions: int
    hunks: int
    rows: list[tuple[str, str, str]]


def _parse_config(path: Path) -> tuple[str, Path, list[PairSpec], Path]:
    payload = json.loads(path.read_text(encoding="utf-8"))

    document_title = str(payload.get("document_title") or "Trace Versions - What Changed and Why")
    output_docx = Path(str(payload.get("output_docx") or "Trace Versions - What Changed and Why.docx"))
    artifacts_dir = Path(str(payload.get("artifacts_dir") or "output/trace_docx_artifacts"))
    global_repo_base = str(payload.get("github_repo_base") or "").strip() or None

    pairs: list[PairSpec] = []
    for row in payload.get("pairs", []):
        before_raw = row.get("before", {})
        after_raw = row.get("after", {})

        def _snap(raw: dict[str, Any], default_label: str) -> SnapshotSpec:
            source = str(raw.get("source") or "github").strip().lower()
            return SnapshotSpec(
                label=str(raw.get("label") or default_label),
                github_url=(str(raw.get("github_url")) if raw.get("github_url") else None),
                description=str(raw.get("description") or ""),
                highlight=str(raw.get("highlight") or "Highlight the changes"),
                source=source,
                repo_path=(str(raw.get("repo_path")) if raw.get("repo_path") else None),
                file_path=(str(raw.get("file_path")) if raw.get("file_path") else None),
                git_ref=(str(raw.get("git_ref")) if raw.get("git_ref") else None),
                start_line=(int(raw.get("start_line")) if raw.get("start_line") is not None else None),
                end_line=(int(raw.get("end_line")) if raw.get("end_line") is not None else None),
                github_repo_base=(str(raw.get("github_repo_base") or global_repo_base or "").strip() or None),
            )

        pairs.append(
            PairSpec(
                pair_title=str(row.get("pair_title") or "Code Comparison"),
                before=_snap(before_raw, "Before"),
                after=_snap(after_raw, "After"),
            )
        )

    if not pairs:
        raise ValueError("Config must contain at least one item in 'pairs'.")

    return document_title, output_docx, pairs, artifacts_dir


def _extract_line_range(github_url: str) -> tuple[int | None, int | None]:
    frag = urlparse(github_url).fragment or ""
    # Supports #L120 or #L120-L170
    m = re.match(r"^L(\d+)(?:-L?(\d+))?$", frag)
    if not m:
        return None, None
    start = int(m.group(1))
    end = int(m.group(2) or start)
    if end < start:
        start, end = end, start
    return start, end


async def _capture_github_snapshot(url: str, out_path: Path, timeout_ms: int = 45000) -> None:
    try:
        from playwright.async_api import async_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency 'playwright'. Install with: pip install playwright ; "
            "python -m playwright install chromium"
        ) from exc

    out_path.parent.mkdir(parents=True, exist_ok=True)

    start_line, end_line = _extract_line_range(url)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={"width": 1780, "height": 1280})
        await page.goto(url, wait_until="networkidle", timeout=timeout_ms)

        # Primary mode: capture line-range region when line anchors are present.
        if start_line is not None and end_line is not None:
            start_loc = page.locator(f"#L{start_line}")
            end_loc = page.locator(f"#L{end_line}")

            if await start_loc.count() > 0 and await end_loc.count() > 0:
                sb = await start_loc.first.bounding_box()
                eb = await end_loc.first.bounding_box()
                if sb and eb:
                    x = max(0, min(sb["x"], eb["x"]) - 28)
                    y = max(0, min(sb["y"], eb["y"]) - 70)
                    right = max(sb["x"] + sb["width"], eb["x"] + eb["width"]) + 980
                    bottom = max(sb["y"] + sb["height"], eb["y"] + eb["height"]) + 55

                    clip = {
                        "x": x,
                        "y": y,
                        "width": min(1700, max(760, right - x)),
                        "height": min(900, max(220, bottom - y)),
                    }
                    await page.screenshot(path=str(out_path), clip=clip)
                    await browser.close()
                    return

        # Fallback: capture main code container.
        candidates = [
            "table.js-file-line-container",
            "div.react-code-lines",
            "div[data-testid='code-viewer']",
            "main",
        ]
        for sel in candidates:
            loc = page.locator(sel)
            if await loc.count() > 0:
                box = await loc.first.bounding_box()
                if box:
                    clip = {
                        "x": max(0, box["x"] - 16),
                        "y": max(0, box["y"] - 16),
                        "width": min(1700, max(760, box["width"] + 32)),
                        "height": min(940, max(260, box["height"] + 32)),
                    }
                    await page.screenshot(path=str(out_path), clip=clip)
                    await browser.close()
                    return

        await page.screenshot(path=str(out_path), full_page=True)
        await browser.close()


def _git_show_file(repo_path: Path, git_ref: str, file_path: str) -> str:
    proc = subprocess.run(
        ["git", "show", f"{git_ref}:{file_path}"],
        cwd=str(repo_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode != 0:
        err = (proc.stderr or "").strip()
        raise RuntimeError(f"git show failed for {git_ref}:{file_path} :: {err}")
    return proc.stdout


def _resolve_git_ref(repo_path: Path, git_ref: str) -> str:
    proc = subprocess.run(
        ["git", "rev-parse", git_ref],
        cwd=str(repo_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode != 0:
        err = (proc.stderr or "").strip()
        raise RuntimeError(f"git rev-parse failed for {git_ref} :: {err}")
    return (proc.stdout or "").strip()


def _repo_web_base(repo_path: Path, override_base: str | None = None) -> str | None:
    if override_base:
        m_override = re.match(r"^https://github\.com/[^/]+/[^/]+/?$", override_base.strip())
        if m_override:
            return override_base.strip().rstrip("/")

    proc = subprocess.run(
        ["git", "config", "--get", "remote.origin.url"],
        cwd=str(repo_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode != 0:
        return None
    origin = (proc.stdout or "").strip()
    m = re.search(r"github\.com[:/](?P<owner>[^/]+)/(?P<repo>[^/.]+)(?:\.git)?$", origin)
    if not m:
        return None
    return f"https://github.com/{m.group('owner')}/{m.group('repo')}"


def _primary_hunk_ranges(
    repo_path: Path,
    file_path: str,
    before_ref: str,
    after_ref: str,
    context: int = 6,
    max_span: int = 120,
) -> tuple[tuple[int, int], tuple[int, int]] | None:
    proc = subprocess.run(
        ["git", "diff", "--unified=0", before_ref, after_ref, "--", file_path],
        cwd=str(repo_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode not in (0, 1):
        return None

    hunk_re = re.compile(r"^@@\s+-(\d+)(?:,(\d+))?\s+\+(\d+)(?:,(\d+))?\s+@@")
    best = None
    for line in (proc.stdout or "").splitlines():
        m = hunk_re.match(line)
        if not m:
            continue

        b_start = int(m.group(1))
        b_count = int(m.group(2) or "1")
        a_start = int(m.group(3))
        a_count = int(m.group(4) or "1")

        # Zero-count appears for pure additions/deletions; keep a visible anchor line.
        b_count_eff = max(1, b_count)
        a_count_eff = max(1, a_count)
        weight = max(b_count_eff, a_count_eff)

        if best is None or weight > best[0]:
            best = (
                weight,
                b_start,
                b_count_eff,
                a_start,
                a_count_eff,
            )

    if best is None:
        return None

    _, b_start, b_count_eff, a_start, a_count_eff = best
    b_lo = max(1, b_start - context)
    b_hi = max(b_lo, b_start + b_count_eff - 1 + context)
    a_lo = max(1, a_start - context)
    a_hi = max(a_lo, a_start + a_count_eff - 1 + context)

    # Keep windows compact for document readability.
    if (b_hi - b_lo + 1) > max_span:
        b_hi = b_lo + max_span - 1
    if (a_hi - a_lo + 1) > max_span:
        a_hi = a_lo + max_span - 1
    return (b_lo, b_hi), (a_lo, a_hi)


def _default_github_url_from_local(spec: SnapshotSpec) -> str | None:
    if spec.github_url:
        return spec.github_url
    return _generated_github_url_from_local(spec)


def _generated_github_url_from_local(spec: SnapshotSpec) -> str | None:
    if not spec.repo_path or not spec.file_path or not spec.git_ref:
        return None

    repo = Path(spec.repo_path)
    base = _repo_web_base(repo, spec.github_repo_base)
    if not base:
        return None
    try:
        full_ref = _resolve_git_ref(repo, spec.git_ref)
    except Exception:
        return None

    base = f"{base}/blob/{full_ref}/{spec.file_path}"
    if spec.start_line is not None and spec.end_line is not None:
        return f"{base}#L{spec.start_line}-L{spec.end_line}"
    if spec.start_line is not None:
        return f"{base}#L{spec.start_line}"
    return base


async def _capture_local_git_snapshot(spec: SnapshotSpec, out_path: Path, timeout_ms: int = 45000) -> None:
    if not spec.repo_path or not spec.file_path or not spec.git_ref:
        raise ValueError("local_git snapshot requires repo_path, file_path, and git_ref")

    repo_path = Path(spec.repo_path)
    text = _git_show_file(repo_path, spec.git_ref, spec.file_path)
    lines = text.splitlines()

    s = spec.start_line or 1
    e = spec.end_line or min(len(lines), s + 120)
    s = max(1, min(s, max(1, len(lines))))
    e = max(s, min(e, max(1, len(lines))))

    # Include light context for readability.
    ctx_before = 2
    ctx_after = 2
    lo = max(1, s - ctx_before)
    hi = min(len(lines), e + ctx_after)

    esc = lambda x: x.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    rows: list[str] = []
    for idx in range(lo, hi + 1):
        ln = lines[idx - 1]
        cls = "hl" if s <= idx <= e else ""
        rows.append(
            f"<tr class='{cls}'><td class='ln'>{idx}</td><td class='code'><pre>{esc(ln)}</pre></td></tr>"
        )

    html = f"""<!doctype html>
<html>
<head>
  <meta charset='utf-8' />
  <style>
    body {{ margin: 0; font-family: 'Segoe UI', Arial, sans-serif; background: #f6f8fa; }}
    .frame {{ margin: 18px; border: 1px solid #d0d7de; border-radius: 10px; overflow: hidden; background: #fff; }}
    .head {{ background: #f6f8fa; border-bottom: 1px solid #d0d7de; padding: 10px 14px; color: #1f2328; font-size: 13px; }}
    table {{ border-collapse: collapse; width: 100%; table-layout: fixed; }}
    td.ln {{ width: 62px; text-align: right; padding: 0 10px; color: #656d76; border-right: 1px solid #d8dee4; vertical-align: top; background: #f6f8fa; font-family: Consolas, monospace; font-size: 12px; }}
    td.code {{ padding: 0 12px; vertical-align: top; }}
    pre {{ margin: 0; white-space: pre-wrap; word-break: break-word; font-family: Consolas, 'Courier New', monospace; font-size: 12px; line-height: 1.45; color: #1f2328; }}
    tr.hl td {{ background: #fff8c5; }}
  </style>
</head>
<body>
  <div class='frame'>
    <div class='head'>{esc(spec.file_path)} @ {esc(spec.git_ref)} lines {s}-{e}</div>
    <table>{''.join(rows)}</table>
  </div>
</body>
</html>"""

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="local_git_snapshot_") as tdir:
        html_path = Path(tdir) / "snapshot.html"
        html_path.write_text(html, encoding="utf-8")

        from playwright.async_api import async_playwright

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page(viewport={"width": 1780, "height": 1260})
            await page.goto(html_path.resolve().as_uri(), wait_until="networkidle", timeout=timeout_ms)
            frame = page.locator(".frame")
            if await frame.count() > 0:
                await frame.first.screenshot(path=str(out_path))
            else:
                await page.screenshot(path=str(out_path), full_page=True)
            await browser.close()


def _add_image_hyperlink(paragraph: Any, image_path: Path, target_url: str, width: Inches) -> None:
    # Add picture in a run first.
    run = paragraph.add_run()
    run.add_picture(str(image_path), width=width)

    # Wrap the image drawing element inside a hyperlink relationship.
    rel_id = paragraph.part.relate_to(target_url, RT.HYPERLINK, is_external=True)
    drawing_nodes = run._r.xpath("./w:drawing")
    if not drawing_nodes:
        return

    drawing = drawing_nodes[0]
    run._r.remove(drawing)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)

    h_run = OxmlElement("w:r")
    h_run.append(drawing)
    hyperlink.append(h_run)

    run._r.addnext(hyperlink)


def _add_source_link(paragraph: Any, url: str, text: str) -> None:
    rel_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)

    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "1155CC")
    rpr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)

    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)

    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def _set_cell_shading(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _set_cell_border(cell: Any, color: str = "8FA8C5", size: str = "8") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        edge = OxmlElement(f"w:{side}")
        edge.set(qn("w:val"), "single")
        edge.set(qn("w:sz"), size)
        edge.set(qn("w:space"), "0")
        edge.set(qn("w:color"), color)
        borders.append(edge)
    tc_pr.append(borders)


def _short_ref(ref: str | None) -> str:
    return (ref or "unknown")[:7]


def _git_diff_patch(repo_path: Path, before_ref: str, after_ref: str, file_path: str) -> str:
    proc = subprocess.run(
        ["git", "diff", "--unified=0", before_ref, after_ref, "--", file_path],
        cwd=str(repo_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode not in (0, 1):
        return ""
    return proc.stdout or ""


def _git_numstat(repo_path: Path, before_ref: str, after_ref: str, file_path: str) -> tuple[int, int]:
    proc = subprocess.run(
        ["git", "diff", "--numstat", before_ref, after_ref, "--", file_path],
        cwd=str(repo_path),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode not in (0, 1):
        return 0, 0
    line = (proc.stdout or "").strip().splitlines()
    if not line:
        return 0, 0
    cols = line[0].split("\t")
    if len(cols) < 2:
        return 0, 0
    try:
        adds = int(cols[0]) if cols[0].isdigit() else 0
        dels = int(cols[1]) if cols[1].isdigit() else 0
    except Exception:
        return 0, 0
    return adds, dels


def _extract_setting_and_value(content: str) -> tuple[str | None, str]:
    raw = content.strip().rstrip(",")
    if not raw:
        return None, ""

    m = re.match(r"^[\"']?([A-Za-z_][A-Za-z0-9_\.-]{2,})[\"']?\s*[:=]\s*(.+)$", raw)
    if m:
        setting = m.group(1)
        value = m.group(2).strip()
        return setting, value[:120]

    return None, raw[:120]


def _build_section_analysis(pair: PairSpec, max_rows: int = 8) -> SectionAnalysis:
    if not (
        pair.before.source == "local_git"
        and pair.after.source == "local_git"
        and pair.before.repo_path
        and pair.before.file_path
        and pair.before.git_ref
        and pair.after.git_ref
    ):
        return SectionAnalysis(additions=0, deletions=0, hunks=0, rows=[])

    repo_path = Path(pair.before.repo_path)
    file_path = pair.before.file_path
    before_ref = pair.before.git_ref
    after_ref = pair.after.git_ref

    patch = _git_diff_patch(repo_path, before_ref, after_ref, file_path)
    additions, deletions = _git_numstat(repo_path, before_ref, after_ref, file_path)

    hunk_re = re.compile(r"^@@\s+-\d+(?:,\d+)?\s+\+\d+(?:,\d+)?\s+@@")
    hunks = 0
    rows_by_setting: dict[str, dict[str, str]] = {}

    for line in patch.splitlines():
        if hunk_re.match(line):
            hunks += 1
            continue
        if line.startswith("+++") or line.startswith("---"):
            continue
        if not line or line[0] not in {"+", "-"}:
            continue

        sign = line[0]
        setting, value = _extract_setting_and_value(line[1:])
        if not setting:
            continue

        row = rows_by_setting.setdefault(setting, {"before": "(empty)", "after": "(empty)"})
        if sign == "-":
            row["before"] = value or "(removed)"
        elif sign == "+":
            row["after"] = value or "(added)"

    rows: list[tuple[str, str, str]] = []
    for setting, vals in rows_by_setting.items():
        if len(rows) >= max_rows:
            break
        rows.append((setting, vals["before"], vals["after"]))

    if not rows:
        rows = [
            ("file_path", file_path, file_path),
            ("commit_ref", _short_ref(before_ref), _short_ref(after_ref)),
            ("line_delta", f"+{additions}", f"-{deletions}"),
        ]

    return SectionAnalysis(additions=additions, deletions=deletions, hunks=hunks, rows=rows)


def _add_section_summary(doc: Document, pair: PairSpec, analysis: SectionAnalysis) -> None:
    file_name = pair.before.file_path or pair.after.file_path or "N/A"
    b_ref = _short_ref(pair.before.git_ref)
    a_ref = _short_ref(pair.after.git_ref)

    p1 = doc.add_paragraph()
    p1.add_run("The problem. ").bold = True
    p1.add_run(
        f"At baseline ({b_ref}), this section reflects earlier logic in {file_name} before the current control refinements."
    )

    p2 = doc.add_paragraph()
    p2.add_run("The change. ").bold = True
    p2.add_run(
        f"The latest commit ({a_ref}) applies {analysis.hunks} focused patch hunk(s), with +{analysis.additions} and -{analysis.deletions} line changes in this file."
    )

    p3 = doc.add_paragraph()
    p3.add_run("The effect. ").bold = True
    p3.add_run(
        f"Comparison now has commit-pinned before/after evidence and a structured delta table with {len(analysis.rows)} key setting/value changes."
    )


def _add_change_table(doc: Document, analysis: SectionAnalysis) -> None:
    t_head = doc.add_paragraph("Code-change table")
    t_head.runs[0].bold = True

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"

    hdr = table.rows[0].cells
    hdr[0].text = "Setting"
    hdr[1].text = "Before (baseline)"
    hdr[2].text = "After (latest)"

    for cell in hdr:
        _set_cell_shading(cell, "1F4E79")
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)

    for setting, before_val, after_val in analysis.rows:
        row = table.add_row().cells
        row[0].text = setting
        row[1].text = before_val
        row[2].text = after_val

    doc.add_paragraph()


def _render_snapshot_cell(cell: Any, spec: SnapshotSpec, image_path: Path) -> None:
    cell.text = ""
    _set_cell_border(cell, color="9DB4CC", size="10")
    _set_cell_shading(cell, "F7FBFF")

    title = cell.paragraphs[0]
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run(spec.label)
    tr.bold = True
    tr.font.size = Pt(16)

    p_img = cell.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    target_url = _generated_github_url_from_local(spec) if spec.source == "local_git" else spec.github_url
    if not target_url:
        target_url = _default_github_url_from_local(spec)
    if target_url:
        _add_image_hyperlink(p_img, image_path, target_url, width=Inches(4.9))
    else:
        run = p_img.add_run()
        run.add_picture(str(image_path), width=Inches(4.9))

    p_desc = cell.add_paragraph(spec.description)
    p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in p_desc.runs:
        r.font.size = Pt(10)

    p_h = cell.add_paragraph(spec.highlight)
    p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = p_h.runs[0]
    hr.bold = True
    hr.italic = True
    hr.font.size = Pt(11)

    # Simulated highlighter effect with yellow background.
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "FFF200")
    hr._r.get_or_add_rPr().append(shd)

    p_link = cell.add_paragraph()
    p_link.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if target_url:
        _add_source_link(p_link, target_url, "Open source on GitHub")
    else:
        p_link.add_run("GitHub source URL not configured")


def _build_docx(document_title: str, output_docx: Path, pairs: list[PairSpec], snapshots: list[tuple[Path, Path]]) -> None:
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(document_title)
    run.bold = True
    run.font.size = Pt(25)

    subtitle = doc.add_paragraph("Before/After code snapshots captured from your repository and mapped to commit-precise GitHub source links")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in subtitle.runs:
        r.font.size = Pt(10)

    stamp = doc.add_paragraph("Validation rule: clicking each image opens the exact source commit, file, and line-range used for that snapshot")
    stamp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = stamp.runs[0]
    sr.font.size = Pt(9)
    sr.italic = True

    doc.add_paragraph()

    for i, pair in enumerate(pairs):
        if i > 0:
            doc.add_page_break()

        ph = doc.add_paragraph(pair.pair_title)
        ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        phr = ph.runs[0]
        phr.bold = True
        phr.font.size = Pt(15)

        analysis = _build_section_analysis(pair)
        _add_section_summary(doc, pair, analysis)
        _add_change_table(doc, analysis)

        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"

        # Set preferred widths for side-by-side snapshot cards.
        table.columns[0].width = Inches(6.2)
        table.columns[1].width = Inches(6.2)

        before_img, after_img = snapshots[i]
        _render_snapshot_cell(table.rows[0].cells[0], pair.before, before_img)
        _render_snapshot_cell(table.rows[0].cells[1], pair.after, after_img)

        doc.add_paragraph()

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))


async def _capture_all_pairs(pairs: list[PairSpec], artifacts_dir: Path) -> list[tuple[Path, Path]]:
    artifacts_dir.mkdir(parents=True, exist_ok=True)
    out: list[tuple[Path, Path]] = []

    for i, pair in enumerate(pairs, start=1):
        before_path = artifacts_dir / f"pair_{i:02d}_before.png"
        after_path = artifacts_dir / f"pair_{i:02d}_after.png"

        # For local git mode, auto-focus each side on the primary changed hunk.
        if (
            pair.before.source == "local_git"
            and pair.after.source == "local_git"
            and pair.before.repo_path
            and pair.after.repo_path
            and pair.before.file_path
            and pair.after.file_path
            and pair.before.file_path == pair.after.file_path
            and pair.before.repo_path == pair.after.repo_path
            and pair.before.git_ref
            and pair.after.git_ref
        ):
            repo = Path(pair.before.repo_path)
            # If config already pins line windows, keep them. Otherwise auto-focus on the primary hunk.
            before_pinned = pair.before.start_line is not None or pair.before.end_line is not None
            after_pinned = pair.after.start_line is not None or pair.after.end_line is not None
            if not before_pinned and not after_pinned:
                ranges = _primary_hunk_ranges(repo, pair.before.file_path, pair.before.git_ref, pair.after.git_ref)
                if ranges is not None:
                    (b_lo, b_hi), (a_lo, a_hi) = ranges
                    pair.before.start_line = b_lo
                    pair.before.end_line = b_hi
                    pair.after.start_line = a_lo
                    pair.after.end_line = a_hi

            # Always force precise permalink generation from local refs to avoid stale URLs.
            before_url = _generated_github_url_from_local(pair.before)
            after_url = _generated_github_url_from_local(pair.after)
            if before_url:
                pair.before.github_url = before_url
            if after_url:
                pair.after.github_url = after_url

        if pair.before.source == "local_git":
            await _capture_local_git_snapshot(pair.before, before_path)
        else:
            if not pair.before.github_url:
                raise ValueError(f"Missing before.github_url for pair {i}")
            await _capture_github_snapshot(pair.before.github_url, before_path)

        if pair.after.source == "local_git":
            await _capture_local_git_snapshot(pair.after, after_path)
        else:
            if not pair.after.github_url:
                raise ValueError(f"Missing after.github_url for pair {i}")
            await _capture_github_snapshot(pair.after.github_url, after_path)
        out.append((before_path, after_path))

    return out


def _validate_pairs(pairs: list[PairSpec]) -> None:
    for idx, pair in enumerate(pairs, start=1):
        for side_name, side in (("before", pair.before), ("after", pair.after)):
            if side.source == "local_git":
                if not side.repo_path or not side.file_path or not side.git_ref:
                    raise ValueError(
                        f"pairs[{idx}].{side_name} in local_git mode requires repo_path, file_path, and git_ref"
                    )
                if side.start_line is not None and side.end_line is not None and side.end_line < side.start_line:
                    raise ValueError(f"pairs[{idx}].{side_name} has end_line < start_line")
                continue

            if not side.github_url or not side.github_url.startswith("https://github.com/"):
                raise ValueError(f"pairs[{idx}].{side_name}.github_url must start with https://github.com/")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate Before/After trace versions DOCX from GitHub snapshots")
    parser.add_argument("--config", required=True, help="Path to config JSON")
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.exists():
        raise FileNotFoundError(f"Config not found: {config_path}")

    document_title, output_docx, pairs, artifacts_dir = _parse_config(config_path)
    _validate_pairs(pairs)

    snapshots = asyncio.run(_capture_all_pairs(pairs, artifacts_dir))
    _build_docx(document_title, output_docx, pairs, snapshots)

    print(json.dumps(
        {
            "status": "PASS",
            "output_docx": str(output_docx),
            "artifacts_dir": str(artifacts_dir),
            "pair_count": len(pairs),
        },
        indent=2,
        ensure_ascii=False,
    ))


if __name__ == "__main__":
    main()
