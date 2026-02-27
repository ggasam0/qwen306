```python
#!/usr/bin/env python3
"""
Linux FREE: .msg (Outlook) -> HTML (from htmlBody) -> PDF (LibreOffice) -> PNGs (pdftoppm)

Goals:
- ONLY use the first message body (this .msg itself).
- Preserve inline CID images (extract attachments, rewrite src="cid:...").
- Do NOT lose table borders: force visible borders on table/th/td as inline CSS.
- Work with extract-msg 0.60.x:
  - NO msg.process()
  - htmlBody may be bytes -> decode
- Export:
  out/<stem>/body.html
  out/<stem>/body.pdf
  out/<stem>/assets/*          (extracted attachments/images)
  out/<stem>-1.png, -2.png...  (in out root for easy consumption)

Install:
  sudo apt-get update
  sudo apt-get install -y libreoffice poppler-utils
  pip install extract-msg beautifulsoup4 lxml

Usage:
  python3 msg_html_to_pdf_png.py mail.msg ./out 200
"""

from __future__ import annotations

import re
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import extract_msg
from bs4 import BeautifulSoup


# ----------------------------
# Utils
# ----------------------------
def die(msg: str, code: int = 2) -> None:
    print(msg, file=sys.stderr)
    raise SystemExit(code)


def run(cmd: List[str]) -> None:
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        die(f"Command failed: {' '.join(cmd)}\nSTDOUT:\n{p.stdout}\nSTDERR:\n{p.stderr}")


def ensure_cmd(cmd: str, apt_pkg_hint: str) -> None:
    if shutil.which(cmd) is None:
        die(f"Missing command: {cmd}. Install: sudo apt-get install -y {apt_pkg_hint}")


def safe_name(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    return s or "mail"


def safe_filename(name: str) -> str:
    name = (name or "").strip().replace("\x00", "")
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    return name or "asset"


def to_text(x) -> str:
    if x is None:
        return ""
    if isinstance(x, bytes):
        # htmlBody in extract-msg 0.60.x can be bytes
        return x.decode("utf-8", errors="ignore")
    if isinstance(x, str):
        return x
    return str(x)


# ----------------------------
# CID image extraction & rewrite
# ----------------------------
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff"}


def looks_like_image_filename(fn: str) -> bool:
    ext = Path(fn).suffix.lower()
    return ext in IMAGE_EXTS


def extract_attachments_and_cids(msg: extract_msg.Message, assets_dir: Path) -> Tuple[Dict[str, str], List[str]]:
    """
    Extract all attachments to assets_dir.
    Build:
      cid_map: cid(lower) -> file:///abs/path
      all_files: list of file:///abs/path for extracted image files (best-effort)
    """
    assets_dir.mkdir(parents=True, exist_ok=True)
    cid_map: Dict[str, str] = {}
    all_image_urls: List[str] = []

    atts = getattr(msg, "attachments", None) or []
    for i, att in enumerate(atts):
        data = getattr(att, "data", None)
        if not data:
            continue

        fn = (
            getattr(att, "longFilename", None)
            or getattr(att, "shortFilename", None)
            or getattr(att, "filename", None)
            or f"attachment_{i}.bin"
        )
        fn = safe_filename(fn)
        out_path = (assets_dir / fn).resolve()
        if out_path.exists():
            out_path = (assets_dir / f"{out_path.stem}_{i}{out_path.suffix}").resolve()

        try:
            out_path.write_bytes(data)
        except Exception:
            continue

        # CID (inline images)
        cid = getattr(att, "cid", None) or getattr(att, "contentId", None) or getattr(att, "content_id", None)
        if isinstance(cid, bytes):
            cid = cid.decode("utf-8", errors="ignore")
        if isinstance(cid, str) and cid.strip():
            cid_clean = cid.strip().strip("<>").lower()
            cid_map[cid_clean] = out_path.as_uri()

        # Collect image files (by filename)
        if looks_like_image_filename(out_path.name):
            all_image_urls.append(out_path.as_uri())

    return cid_map, all_image_urls


def rewrite_cid_images(html: str, cid_map: Dict[str, str]) -> str:
    """
    Replace src="cid:xxxx" or src='cid:xxxx' with file:///... from cid_map.
    """
    def repl_double(m):
        raw = m.group(1)
        cid = raw.strip().strip("<>").lower()
        if cid in cid_map:
            return f'src="{cid_map[cid]}"'
        return m.group(0)

    def repl_single(m):
        raw = m.group(1)
        cid = raw.strip().strip("<>").lower()
        if cid in cid_map:
            return f"src='{cid_map[cid]}'"
        return m.group(0)

    html = re.sub(r'src\s*=\s*"(?:cid:)([^"]+)"', repl_double, html, flags=re.IGNORECASE)
    html = re.sub(r"src\s*=\s*'(?:cid:)([^']+)'", repl_single, html, flags=re.IGNORECASE)
    return html


# ----------------------------
# Border hardening for tables (HTML)
# ----------------------------
PRINT_CSS = """
<style>
@page { margin: 12mm; }
body { font-family: Arial, "PingFang SC", "Microsoft YaHei", sans-serif; font-size: 12pt; }
img { max-width: 100%; height: auto; }
pre { white-space: pre-wrap; word-break: break-word; }
</style>
"""

FORCE_TABLE_STYLE = "border:1px solid #000; border-collapse:collapse;"
FORCE_CELL_STYLE = "border:1px solid #000; padding:4px 6px; vertical-align:top;"


def convert_mso_border_to_standard(style: str) -> str:
    if not style:
        return style
    s = style
    s = re.sub(r"windowtext|auto", "#000", s, flags=re.IGNORECASE)
    if "mso-border-alt" in s.lower() and "border" not in s.lower():
        s += " border:1px solid #000;"
    s = re.sub(r"(\bborder[^;:]*:\s*[^;]*?)\b(\d+(\.\d+)?)pt\b", r"\1 1px", s, flags=re.IGNORECASE)
    return s


def force_table_borders(html: str) -> str:
    soup = BeautifulSoup(html, "lxml")

    # Ensure head exists & inject print css
    if soup.html is None:
        wrapper = BeautifulSoup("<html><head></head><body></body></html>", "lxml")
        # move all nodes into body
        for node in list(soup.contents):
            wrapper.body.append(node)
        soup = wrapper

    if soup.head is None:
        head = soup.new_tag("head")
        soup.html.insert(0, head)

    soup.head.append(BeautifulSoup(PRINT_CSS, "lxml"))

    for table in soup.find_all("table"):
        st = convert_mso_border_to_standard(table.get("style", ""))
        if "border" not in st.lower():
            st = (st + (" " if st else "") + FORCE_TABLE_STYLE).strip()
        else:
            if "border-collapse" not in st.lower():
                st = (st.rstrip(";") + "; border-collapse:collapse;").strip()
        table["style"] = st

    for cell in soup.find_all(["td", "th"]):
        st = convert_mso_border_to_standard(cell.get("style", ""))
        if "border" not in st.lower():
            st = (st + (" " if st else "") + FORCE_CELL_STYLE).strip()
        else:
            if "padding" not in st.lower():
                st = (st.rstrip(";") + "; padding:4px 6px;").strip()
            if "vertical-align" not in st.lower():
                st = (st.rstrip(";") + "; vertical-align:top;").strip()
        cell["style"] = st

    return str(soup)


def ensure_full_html_doc(html: str, title: str) -> str:
    html = html or ""
    if "<html" in html.lower():
        return html
    return f"<!doctype html><html><head><meta charset='utf-8'><title>{title}</title></head><body>{html}</body></html>"


# ----------------------------
# Ensure all images appear (append missing ones)
# ----------------------------
def append_missing_images(html: str, image_urls: List[str]) -> str:
    if not image_urls:
        return html

    soup = BeautifulSoup(html, "lxml")
    if soup.body is None:
        if soup.html is None:
            wrapper = BeautifulSoup("<html><head></head><body></body></html>", "lxml")
            for node in list(soup.contents):
                wrapper.body.append(node)
            soup = wrapper
        else:
            body = soup.new_tag("body")
            soup.html.append(body)

    referenced = set()
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        if src:
            referenced.add(src)

    missing = [u for u in image_urls if u not in referenced]
    if not missing:
        return str(soup)

    soup.body.append(soup.new_tag("hr"))
    h2 = soup.new_tag("h2")
    h2.string = "Images (All extracted)"
    soup.body.append(h2)

    for u in missing:
        block = soup.new_tag("div")
        block["style"] = "margin:12px 0;"
        cap = soup.new_tag("div")
        cap.string = u
        img = soup.new_tag("img")
        img["src"] = u
        block.append(cap)
        block.append(img)
        soup.body.append(block)

    return str(soup)


# ----------------------------
# LibreOffice: HTML -> PDF
# ----------------------------
def html_to_pdf_with_libreoffice(html_path: Path, out_dir: Path) -> Path:
    ensure_cmd("soffice", "libreoffice")

    run([
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(html_path),
    ])

    pdf_path = out_dir / (html_path.stem + ".pdf")
    if not pdf_path.exists() or pdf_path.stat().st_size == 0:
        die(f"LibreOffice produced no PDF: {pdf_path}")
    return pdf_path


# ----------------------------
# PDF -> PNGs
# ----------------------------
def pdf_to_pngs(pdf_path: Path, out_prefix: Path, dpi: int = 200) -> List[Path]:
    ensure_cmd("pdftoppm", "poppler-utils")

    run(["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(out_prefix)])
    return sorted(out_prefix.parent.glob(out_prefix.name + "-*.png"))


# ----------------------------
# Main convert
# ----------------------------
def convert(msg_path: str, out_dir: str, dpi: int = 200) -> None:
    ensure_cmd("soffice", "libreoffice")
    ensure_cmd("pdftoppm", "poppler-utils")

    msg_path_p = Path(msg_path).resolve()
    out_dir_p = Path(out_dir).resolve()
    out_dir_p.mkdir(parents=True, exist_ok=True)

    base = safe_name(msg_path_p.stem)
    work = out_dir_p / base
    assets = work / "assets"
    work.mkdir(parents=True, exist_ok=True)

    msg = extract_msg.Message(str(msg_path_p))  # extract-msg 0.60.x: no process()

    # htmlBody may be bytes
    html = to_text(getattr(msg, "htmlBody", None))
    if not html.strip():
        die("No htmlBody found in this .msg (htmlBody is empty).")

    subject = to_text(getattr(msg, "subject", None)) or base

    # Extract attachments, build CID map
    cid_map, image_urls = extract_attachments_and_cids(msg, assets)

    # Rewrite cid: images
    if cid_map:
        html = rewrite_cid_images(html, cid_map)

    # Ensure full doc
    html = ensure_full_html_doc(html, title=subject)

    # Append any extracted images that were not referenced (ensures signature images show up even if not linked)
    html = append_missing_images(html, image_urls)

    # Force visible table borders
    html = force_table_borders(html)

    html_path = work / "body.html"
    html_path.write_text(html, encoding="utf-8", errors="ignore")

    pdf_path = html_to_pdf_with_libreoffice(html_path, work)
    print(f"PDF: {pdf_path}")

    # PNGs saved to out_dir root: out/<stem>-1.png ...
    out_prefix = out_dir_p / base
    pngs = pdf_to_pngs(pdf_path, out_prefix=out_prefix, dpi=dpi)
    for p in pngs:
        print(f"PNG: {p}")

    # Debug hints
    print(f"HTML: {html_path}")
    print(f"Assets dir: {assets}")
    print(f"Attachments extracted: {len(list(assets.glob('*')))}")
    print(f"CID mapped: {len(cid_map)}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        die("Usage: python3 msg_html_to_pdf_png.py <mail.msg> <out_dir> [dpi]")

    dpi = int(sys.argv[3]) if len(sys.argv) >= 4 else 200
    convert(sys.argv[1], sys.argv[2], dpi)
```
