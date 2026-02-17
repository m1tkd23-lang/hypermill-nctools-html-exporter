#src\hypermill_nctools_html_exporter\model.py
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class NcToolRecord:
    # identity
    nctool_no: Optional[int] = None
    nctool_name: str = ""
    nctool_comment: str = ""

    # assembly
    holder_name: str = ""
    holder_length: str = ""
    tool_name: str = ""
    tool_length: str = ""
    extensions_str: str = ""  # extensionを連結した表示用（例: "EXT1(L=50) / EXT2(L=80)"）


    # computed lengths (mm as string for Excel)
    ext_overhang_mm: str = ""      # extension突き出し（合計）
    tool_overhang_mm: str = ""     # 工具突き出し（tool_length相当を複写）
    overhang_mm: str = ""          # 突き出し長さ（ext + tool）
    total_length_mm: str = ""      # 全長（holder + ext + tool）


    # image
    image_rel_src: str = ""
    image_abs_path: Optional[Path] = None
    image_cached_path: Optional[Path] = None  # resized/cached png

    # tool page
    tool_type: str = ""
    tool_page_name: str = ""

    tool_diameter_mm: str = ""
    tool_corner_radius_mm: str = ""
    tool_flutes: str = ""
    tool_cut_length_ap_mm: str = ""
    tool_shank_d_mm: str = ""
    tool_chamfer_len_mm: str = ""
    tool_tip_len_mm: str = ""
    tool_taper_angle_deg: str = ""
    spindle_rotation: str = ""

    # conditions (first row)
    cond_S_n: str = ""
    cond_FX: str = ""
    cond_FZ: str = ""
    cond_Fr: str = ""
    cond_ap: str = ""
    cond_ae: str = ""

    # holder page
    holder_page_name: str = ""
    holder_comment: str = ""

    # provenance
    source_html_path: str = ""

    # parsing warnings
    warnings: list[str] = field(default_factory=list)
