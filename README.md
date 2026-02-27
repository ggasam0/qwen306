#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Dict, List, Tuple, Set

from bs4 import BeautifulSoup
from email import policy
from email.parser import BytesParser
from email.message import Message


# -------- utils --------
def die(msg: str, code: int = 2) -> None:
    print(msg, file=sys.stderr)
    raise SystemExit(code)


def run(cmd: List[str]) -> None:
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        die(f"Command failed: {' '.join(cmd)}\nSTDOUT:\n{p.stdout}\nSTDERR:\n{p.stderr}")


def safe_name(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    return s or "mail"


def safe_filename(name: str) -> str:
    name = (name or "").strip().replace("\x00", "")
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    return name or "asset"


# -------- msg -> eml --------
def msg_to_eml(msg_path: Path, eml_path: Path) -> None:
    if shutil.which("msgconvert") is None:
        die("msgconvert not found. Install: sudo apt-get install libemail-outlook-message-perl")

    tmp = eml_path.with_suffix(".tmp.eml")
    if tmp.exists():
        tmp.unlink()

    p = subprocess.run(["msgconvert", str(msg_path), "-o", str(tmp)],
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode == 0 and tmp.exists() and tmp.stat().st_size > 0:
        tmp.replace(eml_path)
        return

    p2 = subprocess.run(["msgconvert", str(msg_path)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if p2.returncode != 0:
        die(f"msgconvert failed.\nSTDERR:\n{p2.stderr.decode(errors='ignore')}")
    eml_path.write_bytes(p2.stdout)


def parse_eml(eml_path: Path) -> Message:
    return BytesParser(policy=policy.default).parsebytes(eml_path.read_bytes())


def get_best_body(eml: Message) -> Tuple[str, str]:
    if eml.is_multipart():
        html_part = eml.get_body(preferencelist=("html",))
        if html_part:
            return "text/html", html_part.get_content()
        text_part = eml.get_body(preferencelist=("plain",))
        if text_part:
            return "text/plain", text_part.get_content()
    else:
        ctype = eml.get_content_type()
        if ctype in ("text/html", "text/plain"):
            return ctype, eml.get_content()
    return "text/plain", ""


def iter_leaf_parts(m: Message):
    if m.is_multipart():
        for p in m.iter_parts():
            yield from iter_leaf_parts(p)
    else:
        yield m


# -------- extract images (CID + octet-stream by filename) --------
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff"}


def guess_ext(mime: str) -> str:
    mime = (mime or "").lower()
    return {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/jpg": ".jpg",
        "image/gif": ".gif",
        "image/webp": ".webp",
        "image/bmp": ".bmp",
        "image/tiff": ".tif",
    }.get(mime, "")


def looks_like_image(part: Message) -> bool:
    ctype = (part.get_content_type() or "").lower()
    if ctype.startswith("image/"):
        return True
    fn = part.get_filename() or ""
    return Path(fn).suffix.lower() in IMAGE_EXTS


def extract_all_images(eml: Message, assets_dir: Path) -> Tuple[Dict[str, str], List[str]]:
    assets_dir.mkdir(parents=True, exist_ok=True)
    cid_map: Dict[str, str] = {}
    all_urls: List[str] = []
    idx = 0

    for part in iter_leaf_parts(eml):
        if not looks_like_image(part):
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        cid = part.get("Content-ID")
        fname = part.get_filename()

        ext = ""
        if fname:
            ext = Path(fname).suffix
        if not ext:
            ext = guess_ext(part.get_content_type()) or ".bin"
        if not ext.startswith("."):
            ext = "." + ext

        base = "img"
        if fname:
            base = Path(safe_filename(fname)).stem or "img"

        out_name = f"{base}_{idx}{ext}"
        idx += 1

        out_path = (assets_dir / out_name).resolve()
        out_path.write_bytes(payload)

        file_url = out_path.as_uri()
        all_urls.append(file_url)

        if cid:
            cid_clean = cid.strip().strip("<>").lower()
            if cid_clean:
                cid_map[cid_clean] = file_url

    return cid_map, all_urls


def rewrite_cid_src(html: str, cid_map: Dict[str, str]) -> str:
    def repl_d(m):
        cid = m.group(1).strip().strip("<>").lower()
        return f'src="{cid_map.get(cid, "cid:"+m.group(1))}"'
    def repl_s(m):
        cid = m.group(1).strip().strip("<>").lower()
        return f"src='{cid_map.get(cid, 'cid:'+m.group(1))}'"

    html = re.sub(r'src\s*=\s*"(?:cid:)([^"]+)"', repl_d, html, flags=re.IGNORECASE)
    html = re.sub(r"src\s*=\s*'(?:cid:)([^']+)'", repl_s, html, flags=re.IGNORECASE)
    return html


# -------- table border hardening for wkhtmltopdf --------
PRINT_HEAD_CSS = """
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
    s = re.sub(r"windowtext|auto", "#000", style, flags=re.IGNORECASE)
    if "mso-border-alt" in s.lower() and "border" not in s.lower():
        s += " border:1px solid #000;"
    s = re.sub(r"(\bborder[^;:]*:\s*[^;]*?)\b(\d+(\.\d+)?)pt\b", r"\1 1px", s, flags=re.IGNORECASE)
    return s


def harden_tables_for_wk(html_doc: str) -> str:
    soup = BeautifulSoup(html_doc, "lxml")
    if soup.html is None:
        wrapper = BeautifulSoup("<html><head></head><body></body></html>", "lxml")
        for node in list(soup.contents):
            wrapper.body.append(node)
        soup = wrapper
    if soup.head is None:
        head = soup.new_tag("head")
        soup.html.insert(0, head)

    soup.head.append(BeautifulSoup(PRINT_HEAD_CSS, "lxml"))

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


# -------- ensure all images appear --------
def append_all_images(html_doc: str, all_urls: List[str]) -> str:
    soup = BeautifulSoup(html_doc, "lxml")
    if soup.body is None:
        if soup.html is None:
            wrapper = BeautifulSoup("<html><head></head><body></body></html>", "lxml")
            for node in list(soup.contents):
                wrapper.body.append(node)
            soup = wrapper
        else:
            body = soup.new_tag("body")
            soup.html.append(body)

    referenced: Set[str] = set()
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        if src:
            referenced.add(src)

    missing = [u for u in all_urls if u not in referenced]
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


# -------- HTML doc wrapper --------
def ensure_full_html_doc(mime: str, body: str, title: str) -> str:
    if mime == "text/plain":
        esc = (
            (body or "")
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        return f"<!doctype html><html><head><meta charset='utf-8'/><title>{title}</title></head><body><pre>{esc}</pre></body></html>"

    if "<html" in (body or "").lower():
        return body
    return f"<!doctype html><html><head><meta charset='utf-8'/><title>{title}</title></head><body>{body or ''}</body></html>"


# -------- render: html -> pdf -> png --------
def html_to_pdf_wk(html_path: Path, pdf_path: Path) -> None:
    if shutil.which("wkhtmltopdf") is None:
        die("wkhtmltopdf not found. Install: sudo apt-get install wkhtmltopdf")

    html_path = html_path.resolve()
    work_dir = html_path.parent.resolve()

    cmd = [
        "wkhtmltopdf",
        "--enable-local-file-access",
        "--allow", str(work_dir),
        "--allow", str(work_dir / "assets"),
        "--print-media-type",
        "--disable-smart-shrinking",
        "--load-error-handling", "ignore",
        "--dpi", "200",
        str(html_path),
        str(pdf_path),
    ]
    run(cmd)


def pdf_to_pngs(pdf_path: Path, out_prefix: Path, dpi: int = 200) -> List[Path]:
    """
    Use pdftoppm to generate PNGs:
      out_prefix-1.png, out_prefix-2.png, ...
    """
    if shutil.which("pdftoppm") is None:
        die("pdftoppm not found. Install: sudo apt-get install poppler-utils")

    cmd = [
        "pdftoppm",
        "-png",
        "-r", str(dpi),
        str(pdf_path),
        str(out_prefix),
    ]
    run(cmd)

    # Collect output files
    out_files = sorted(out_prefix.parent.glob(out_prefix.name + "-*.png"))
    return out_files


# -------- pipeline: msg -> images --------
def convert_msg_to_images(msg_path: str, out_dir: str, dpi: int = 200) -> List[Path]:
    msg_path_p = Path(msg_path).resolve()
    out_dir_p = Path(out_dir).resolve()
    out_dir_p.mkdir(parents=True, exist_ok=True)

    base = safe_name(msg_path_p.stem)
    work_dir = out_dir_p / base
    assets_dir = work_dir / "assets"
    work_dir.mkdir(parents=True, exist_ok=True)

    eml_path = work_dir / f"{base}.eml"
    html_path = work_dir / "render.html"
    pdf_path = work_dir / f"{base}.pdf"

    # 1) MSG -> EML
    msg_to_eml(msg_path_p, eml_path)

    # 2) Parse EML
    eml = parse_eml(eml_path)
    subject = str(eml.get("Subject") or base)
    mime, body = get_best_body(eml)

    # 3) Build HTML
    html_doc = ensure_full_html_doc(mime, body, title=subject)

    # 4) Extract images + CID rewrite + append all images
    cid_map, all_urls = extract_all_images(eml, assets_dir)
    if cid_map:
        html_doc = rewrite_cid_src(html_doc, cid_map)
    if all_urls:
        html_doc = append_all_images(html_doc, all_urls)

    # 5) Harden table borders for wkhtmltopdf
    html_doc = harden_tables_for_wk(html_doc)

    html_path.write_text(html_doc, encoding="utf-8")

    # 6) HTML -> PDF
    html_to_pdf_wk(html_path, pdf_path)

    # 7) PDF -> PNG(s) to out_dir root (not work_dir), easier to consume
    out_prefix = out_dir_p / f"{base}"
    pngs = pdf_to_pngs(pdf_path, out_prefix=out_prefix, dpi=dpi)

    return pngs


if __name__ == "__main__":
    if len(sys.argv) < 3:
        die("Usage: python3 msg_to_images.py <input.msg> <output_dir> [dpi]")

    dpi = int(sys.argv[3]) if len(sys.argv) >= 4 else 200
    pngs = convert_msg_to_images(sys.argv[1], sys.argv[2], dpi=dpi)
    for p in pngs:
        print(str(p))