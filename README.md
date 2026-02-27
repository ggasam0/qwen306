#!/usr/bin/env python3
from __future__ import annotations

import os
import re
import shutil
import subprocess
from pathlib import Path
from typing import Optional, List

import extract_msg


def die(msg: str, code: int = 2) -> None:
    raise SystemExit(f"{msg}")


def run(cmd: List[str]) -> None:
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        die(f"Command failed: {' '.join(cmd)}\nSTDOUT:\n{p.stdout}\nSTDERR:\n{p.stderr}")


def safe_name(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    return s or "mail"


def ensure_cmd(name: str, apt_pkg_hint: str) -> None:
    if shutil.which(name) is None:
        die(f"Missing command: {name}. Install: sudo apt-get install -y {apt_pkg_hint}")


def extract_rtf_from_msg(msg_path: Path) -> Optional[str]:
    """
    Extract RTF body (preferred) from .msg.
    """
    msg = extract_msg.Message(str(msg_path))
    msg.extract_attachments = False
    msg.process()

    # extract-msg versions differ; try common attribute names
    rtf = getattr(msg, "rtfBody", None) or getattr(msg, "rtf", None)
    if isinstance(rtf, bytes):
        try:
            rtf = rtf.decode("utf-8", errors="ignore")
        except Exception:
            rtf = None
    if isinstance(rtf, str):
        rtf = rtf.strip()
        return rtf if rtf else None
    return None


def extract_html_or_text_from_msg(msg_path: Path) -> str:
    """
    Fallback: if no RTF, we still keep the first body content as plain text/HTML.
    """
    msg = extract_msg.Message(str(msg_path))
    msg.extract_attachments = False
    msg.process()

    html = getattr(msg, "htmlBody", None) or getattr(msg, "html", None)
    if isinstance(html, str) and html.strip():
        # Wrap as minimal HTML file
        return html.strip()

    body = getattr(msg, "body", None)
    if isinstance(body, str) and body.strip():
        # Wrap as RTF? No. We'll create simple text-in-RTF to allow LO to render.
        text = body.strip()
        # minimal RTF wrapper (keeps plaintext)
        return r"{\rtf1\ansi " + text.replace("\\", r"\\").replace("{", r"\{").replace("}", r"\}") + "}"
    return r"{\rtf1\ansi (No body content found.)}"


def write_attachments(msg_path: Path, out_dir: Path) -> None:
    """
    Extract attachments to out_dir/attachments for troubleshooting/verification.
    (Not required for rendering but useful to inspect what MSG contains.)
    """
    att_dir = out_dir / "attachments"
    att_dir.mkdir(parents=True, exist_ok=True)

    msg = extract_msg.Message(str(msg_path))
    msg.extract_attachments = True
    msg.attachments_dir = str(att_dir)
    msg.process()
    msg.save_attachments()


def rtf_to_pdf_via_libreoffice(rtf_path: Path, out_dir: Path) -> Path:
    ensure_cmd("soffice", "libreoffice")
    out_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice writes PDF with same base name into out_dir
    run([
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(rtf_path),
    ])

    pdf_path = out_dir / (rtf_path.stem + ".pdf")
    if not pdf_path.exists() or pdf_path.stat().st_size == 0:
        die(f"LibreOffice conversion produced no PDF: {pdf_path}")
    return pdf_path


def pdf_to_pngs(pdf_path: Path, out_prefix: Path, dpi: int = 200) -> List[Path]:
    ensure_cmd("pdftoppm", "poppler-utils")
    # output: <prefix>-1.png, <prefix>-2.png, ...
    run(["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(out_prefix)])
    return sorted(out_prefix.parent.glob(out_prefix.name + "-*.png"))


def convert_msg(msg_path: str, out_dir: str, dpi: int = 200) -> None:
    msg_path_p = Path(msg_path).resolve()
    out_dir_p = Path(out_dir).resolve()
    out_dir_p.mkdir(parents=True, exist_ok=True)

    base = safe_name(msg_path_p.stem)
    work = out_dir_p / base
    work.mkdir(parents=True, exist_ok=True)

    # Optional: dump attachments for debugging
    write_attachments(msg_path_p, work)

    # Prefer RTF
    rtf = extract_rtf_from_msg(msg_path_p)
    if rtf is None:
        # fallback: wrap html/plain into something LO can consume
        rtf = extract_html_or_text_from_msg(msg_path_p)

    rtf_path = work / "body.rtf"
    rtf_path.write_text(rtf, encoding="utf-8", errors="ignore")

    pdf_path = rtf_to_pdf_via_libreoffice(rtf_path, work)
    print(f"PDF: {pdf_path}")

    # Convert to images if you want
    out_prefix = out_dir_p / f"{base}"
    pngs = pdf_to_pngs(pdf_path, out_prefix=out_prefix, dpi=dpi)
    for p in pngs:
        print(f"PNG: {p}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        die("Usage: python3 msg_rtf_to_pdf_png.py <input.msg> <output_dir> [dpi]")
    dpi = int(sys.argv[3]) if len(sys.argv) >= 4 else 200
    convert_msg(sys.argv[1], sys.argv[2], dpi=dpi)