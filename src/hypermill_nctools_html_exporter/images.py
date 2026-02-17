# src\hypermill_nctools_html_exporter\images.py
from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple
import tempfile

from PIL import Image


def resolve_image_path(html_path: Path, image_rel_src: str) -> Optional[Path]:
    """
    HTML内の img src（例: img\\xxxx.png）を、実ファイルに解決する。
    """
    if not image_rel_src:
        return None
    rel = image_rel_src.replace("\\", "/")
    p = (html_path.parent / rel).resolve()
    return p if p.exists() and p.is_file() else None


def make_temp_resized_png(
    src_img: Path,
    *,
    key_name: str = "",
    max_px: int = 320,
) -> Tuple[Optional[Path], Optional[str]]:
    """
    画像をPNGとして「OSテンポラリ」に縮小保存する（出力先フォルダには一切作らない）。
    戻り: (temp_png_path, error_message)
    """
    try:
        src_img = Path(src_img)
        if not src_img.exists():
            return None, f"画像が見つかりません: {src_img}"

        with Image.open(src_img) as im:
            im = im.convert("RGBA")
            w, h = im.size
            m = max(w, h)
            if m > max_px and m > 0:
                scale = max_px / m
                new_w = max(1, int(w * scale))
                new_h = max(1, int(h * scale))
                im = im.resize((new_w, new_h), Image.Resampling.LANCZOS)

            fd, tmp_name = tempfile.mkstemp(prefix="hmimg_", suffix=".png")
            # fdは使わない（Windowsでロック回避のため閉じる）
            try:
                import os
                os.close(fd)
            except Exception:
                pass

            tmp_path = Path(tmp_name)
            im.save(tmp_path, format="PNG", optimize=True)

        return tmp_path, None

    except Exception as e:
        return None, f"画像縮小(temp)失敗: {src_img} ({e})"
