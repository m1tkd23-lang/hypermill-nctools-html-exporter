# src/hypermill_nctools_html_exporter/core.py
from __future__ import annotations

from pathlib import Path
from typing import Optional, Callable, Dict, Any, Tuple, List, Literal

from openpyxl import load_workbook

from .parse_html import parse_nctools_html
from .images import resolve_image_path, make_temp_resized_png
from .export_xlsx import write_xlsx
from .util import sanitize_filename
from .export_xlsx_blocks import export_blocks_f2_xlsx


ProgressCb = Callable[[int, int, str], None]  # (done, total, message)
OutLang = Literal["ja", "en"]


def export_from_html(
    html_path: Path,
    out_dir: Path,
    embed_images: bool = True,
    max_px: int = 320,
    progress: Optional[ProgressCb] = None,
) -> Tuple[Path, Dict[str, Any]]:
    """
    HTML 1つ -> XLSX 1つ
    - embed_images: True=埋め込み(推奨), False=非埋め込み（画像処理しない）
    - max_px: 埋め込み画像の最大辺(px)
    """
    html_path = html_path.expanduser().resolve()
    out_dir = out_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not html_path.exists():
        raise FileNotFoundError(str(html_path))

    if progress:
        progress(0, 4, "HTMLを解析中...")

    records, parse_errors = parse_nctools_html(html_path)

    if progress:
        progress(1, 4, "画像パスを解決中...")

    base_name = sanitize_filename(html_path.stem)
    out_folder = out_dir / base_name
    out_folder.mkdir(parents=True, exist_ok=True)

    errors_for_sheet: List[tuple[int, str, str]] = []
    temp_files: list[Path] = []

    # 画像解決 & temp縮小（出力先にimagesは作らない）
    for i, rec in enumerate(records, start=2):  # Excel row index (header=1)
        abs_img = resolve_image_path(html_path, rec.image_rel_src)
        rec.image_abs_path = abs_img

        if embed_images:
            if not abs_img:
                errors_for_sheet.append((i, rec.nctool_name, f"画像が見つかりません: {rec.image_rel_src}"))
            else:
                key = f"{rec.nctool_no or 'NA'}_{rec.nctool_name}".strip()
                tmp_png, err = make_temp_resized_png(abs_img, key_name=key, max_px=max_px)
                rec.image_cached_path = tmp_png
                if tmp_png:
                    temp_files.append(tmp_png)
                if err:
                    errors_for_sheet.append((i, rec.nctool_name, err))

        for w in rec.warnings:
            errors_for_sheet.append((i, rec.nctool_name, w))

    for e in parse_errors:
        errors_for_sheet.append((0, "", e))

    if progress:
        progress(2, 4, "XLSXを書き込み中...")

    out_xlsx = out_folder / f"nctools_list__{base_name}.xlsx"
    try:
        write_xlsx(records, out_xlsx, embed_images=embed_images)
    finally:
        for p in temp_files:
            try:
                p.unlink()
            except Exception:
                pass

    # errorsシートに書き込む
    wb = load_workbook(out_xlsx)
    ws_err = wb["errors"]
    for row_index, name, msg in errors_for_sheet:
        ws_err.append([row_index, name, msg])
    wb.save(out_xlsx)

    if progress:
        progress(4, 4, "完了")

    summary = {
        "html": str(html_path),
        "out_xlsx": str(out_xlsx),
        "records": len(records),
        "embed_images": embed_images,
        "max_px": max_px,
        "errors": len(errors_for_sheet),
    }
    return out_xlsx, summary


def export_report_f2_from_html(
    html_path: Path,
    out_dir: Path,
    embed_images: bool = True,
    max_px: int = 320,
    progress: Optional[ProgressCb] = None,
    out_lang: OutLang = "ja",
) -> Tuple[Path, dict]:
    """
    HTML1つ → F2帳票（3行ブロック）XLSX
    出力先に images フォルダは作らない（縮小はテンポラリ）。
    """
    html_path = html_path.expanduser().resolve()
    out_dir = out_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if progress:
        progress(0, 3, "HTMLを解析中...")

    records, parse_errors = parse_nctools_html(html_path)

    if progress:
        progress(1, 3, "画像を準備中...")

    base_name = sanitize_filename(html_path.stem)
    out_folder = out_dir / base_name
    out_folder.mkdir(parents=True, exist_ok=True)

    errors_for_sheet: list[tuple[int, str, str]] = []
    temp_files: list[Path] = []

    for i, rec in enumerate(records, start=1):
        abs_img = resolve_image_path(html_path, rec.image_rel_src)
        rec.image_abs_path = abs_img

        if embed_images:
            if abs_img:
                key = f"{rec.nctool_no or 'NA'}_{rec.nctool_name}"
                tmp_png, err = make_temp_resized_png(abs_img, key_name=key, max_px=max_px)
                rec.image_cached_path = tmp_png
                if tmp_png:
                    temp_files.append(tmp_png)
                if err:
                    errors_for_sheet.append((i, rec.nctool_name, err))
            else:
                errors_for_sheet.append((i, rec.nctool_name, f"画像が見つかりません: {rec.image_rel_src}"))

        for w in rec.warnings:
            errors_for_sheet.append((i, rec.nctool_name, w))

    for e in parse_errors:
        errors_for_sheet.append((0, "", e))

    if progress:
        progress(2, 3, "XLSXを書き込み中...")

    out_xlsx = out_folder / f"nctools_report__{base_name}.xlsx"
    try:
        written, img_count = export_blocks_f2_xlsx(
            records,
            out_xlsx,
            embed_images=embed_images,
            lang=out_lang,
        )
    finally:
        for p in temp_files:
            try:
                p.unlink()
            except Exception:
                pass

    if progress:
        progress(3, 3, "完了")

    summary = {
        "html": str(html_path),
        "out_xlsx": str(out_xlsx),
        "records": written,
        "embedded_images": img_count,
        "errors": len(errors_for_sheet),
        "out_lang": out_lang,
    }
    return out_xlsx, summary