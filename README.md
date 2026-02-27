#!/usr/bin/env python3

import re
import sys
import shutil
import subprocess
from pathlib import Path
from bs4 import BeautifulSoup
import extract_msg


def run(cmd):
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"Command failed:\n{p.stderr}")


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]+', "_", s)


def extract_cid_images(msg, assets_dir):
    assets_dir.mkdir(parents=True, exist_ok=True)
    cid_map = {}

    for i, att in enumerate(msg.attachments or []):
        data = getattr(att, "data", None)
        cid = getattr(att, "cid", None) or getattr(att, "contentId", None)
        name = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None)

        if not data:
            continue

        if not name:
            name = f"img_{i}.bin"

        path = assets_dir / safe_name(name)
        path.write_bytes(data)

        if cid:
            cid_clean = cid.strip("<>").lower()
            cid_map[cid_clean] = path.resolve().as_uri()

    return cid_map


def rewrite_cid(html, cid_map):
    def repl(m):
        cid = m.group(1).strip("<>").lower()
        if cid in cid_map:
            return f'src="{cid_map[cid]}"'
        return m.group(0)

    html = re.sub(r'src="cid:([^"]+)"', repl, html, flags=re.IGNORECASE)
    return html


def force_table_borders(html):
    soup = BeautifulSoup(html, "lxml")

    for table in soup.find_all("table"):
        style = table.get("style", "")
        if "border" not in style.lower():
            style += " border:1px solid #000; border-collapse:collapse;"
        table["style"] = style

    for cell in soup.find_all(["td", "th"]):
        style = cell.get("style", "")
        if "border" not in style.lower():
            style += " border:1px solid #000; padding:4px;"
        cell["style"] = style

    return str(soup)


def convert_with_libreoffice(input_path, out_dir):
    run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(input_path)
    ])

    pdf = out_dir / (input_path.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError("PDF not generated")
    return pdf


def pdf_to_png(pdf_path, out_prefix, dpi=200):
    run([
        "pdftoppm",
        "-png",
        "-r", str(dpi),
        str(pdf_path),
        str(out_prefix)
    ])


def convert(msg_path, out_dir, dpi=200):
    msg = extract_msg.Message(msg_path)

    html = msg.htmlBody
    if not html:
        raise RuntimeError("No htmlBody found")

    base = safe_name(Path(msg_path).stem)
    work = Path(out_dir) / base
    assets = work / "assets"
    work.mkdir(parents=True, exist_ok=True)

    # 提取CID图片
    cid_map = extract_cid_images(msg, assets)

    # 替换cid
    html = rewrite_cid(html, cid_map)

    # 强制表格边框
    html = force_table_borders(html)

    # 确保完整HTML结构
    if "<html" not in html.lower():
        html = f"<html><head><meta charset='utf-8'></head><body>{html}</body></html>"

    html_path = work / "body.html"
    html_path.write_text(html, encoding="utf-8")

    # HTML -> PDF
    pdf = convert_with_libreoffice(html_path, work)
    print("PDF:", pdf)

    # PDF -> PNG
    out_prefix = Path(out_dir) / base
    pdf_to_png(pdf, out_prefix, dpi)
    print("PNG generated")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python3 msg_html_to_pdf_png.py mail.msg ./out")
        sys.exit(1)

    dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 200
    convert(sys.argv[1], sys.argv[2], dpi)