# src/hypermill_nctools_html_exporter/parse_html.py
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional

from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import warnings

from .model import NcToolRecord
from .util import clean_text

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# ----------------------------
# h3 patterns (JA/EN)
# ----------------------------
_RE_NCTOOL_H3_JA = re.compile(r"NCツール\(N\):(.+?)\s*\((\d+)\)\s*$")
_RE_NCTOOL_H3_EN = re.compile(r"NC-Tool:(.+?)\s*\((\d+)\)\s*$")

_RE_TOOL_H3_JA = re.compile(r"工具:\s*(.+?)\s*\((.+?)\)\s*$")
_RE_TOOL_H3_EN = re.compile(r"Tool:\s*(.+?)\s*\((.+?)\)\s*$")

_RE_HOLDER_H3_JA = re.compile(r"ホルダー:\s*(.+)\s*$")
_RE_HOLDER_H3_EN = re.compile(r"Holder:\s*(.+)\s*$")

_RE_SUBHOLDER_H3_JA = re.compile(r"サブホルダー:\s*(.+)\s*$")
# 英語HTMLでは extension が独立ページとして出る
_RE_EXTENSION_H3_EN = re.compile(r"Extension:\s*(.+)\s*$")


# ----------------------------
# Key normalization
# ----------------------------
_GRID_HEADER_MAP = {
    # coupling table
    "カップリング種類": "coupling_type",
    "Coupling type": "coupling_type",

    "名称": "name",
    "Name": "name",

    # 日本語は「全長」、英語は「Reach」
    "全長": "reach",
    "Reach": "reach",
}

_KV_KEY_MAP = {
    # NC Tool comment
    "NCツール コメント": "nctool_comment",
    "NC-Tool comment": "nctool_comment",

    # Holder comment
    "ホルダー コメント": "holder_comment",
    "Holder comment": "holder_comment",

    # Tool dims (JA/EN)
    "直径": "diameter",
    "Diameter": "diameter",

    "コーナー半径": "corner_radius",
    "Corner radius": "corner_radius",

    "刃数": "cutting_edges",
    "Cutting edges": "cutting_edges",

    # 日本語版は "切削長さ (ap)"、英語版は "Cutting length"
    "切削長さ (ap)": "cutting_length",
    "Cutting length": "cutting_length",

    "シャンク直径": "shank_diameter",
    "Shank diameter": "shank_diameter",

    "面取り長さ": "chamfer_length",
    "Chamfer length": "chamfer_length",

    "先端長さ": "tip_length",
    "Tip length": "tip_length",

    "テーパー角度": "cone_angle",
    "Cone angle": "cone_angle",

    "スピンドル回転方向": "spindle_orientation",
    "Spindle orientation": "spindle_orientation",
}


def _norm_grid_header(h: str) -> str:
    h = clean_text(h)
    return _GRID_HEADER_MAP.get(h, h)


def _norm_kv_key(k: str) -> str:
    k = clean_text(k)
    return _KV_KEY_MAP.get(k, k)


def _parse_kv_table(table) -> Dict[str, str]:
    """
    2列/4列のKV表を dict で返す（キーは“正規化”して返す）
    """
    d: Dict[str, str] = {}
    for tr in table.find_all("tr"):
        tds = [clean_text(td.get_text(" ", strip=True)) for td in tr.find_all("td")]
        if len(tds) == 2 and tds[0]:
            d[_norm_kv_key(tds[0])] = tds[1]
        elif len(tds) == 4:
            if tds[0]:
                d[_norm_kv_key(tds[0])] = tds[1]
            if tds[2]:
                d[_norm_kv_key(tds[2])] = tds[3]
    return d


def _parse_grid_table(table) -> List[Dict[str, str]]:
    """
    1行目をヘッダとして扱うグリッド表を row dict のlistへ。
    ヘッダは“正規化”して返す。
    """
    trs = table.find_all("tr")
    if not trs:
        return []
    header_raw = [clean_text(td.get_text(" ", strip=True)) for td in trs[0].find_all("td")]
    header = [_norm_grid_header(h) for h in header_raw]

    out: List[Dict[str, str]] = []
    for tr in trs[1:]:
        vals = [clean_text(td.get_text(" ", strip=True)) for td in tr.find_all("td")]
        if not any(vals):
            continue
        row = {header[i]: (vals[i] if i < len(vals) else "") for i in range(len(header))}
        out.append(row)
    return out


def _to_float_mm(s: str) -> float | None:
    """
    '130', '130.000', '130,000', '１３０．０００ mm' などから数値をfloat化。
    取れなければ None。
    """
    if not s:
        return None
    s = clean_text(s)

    # 全角数字・全角記号を半角へ
    trans = str.maketrans("０１２３４５６７８９．－＋", "0123456789.-+")
    s = s.translate(trans)

    # 桁区切りカンマ除去（130,000 -> 130000）
    s = s.replace(",", "")

    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


def _fmt_mm(x: float | None) -> str:
    """
    40.0 -> '40', 40.123 -> '40.123'
    """
    if x is None:
        return ""
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.3f}".rstrip("0").rstrip(".")


def _build_extensions_str_from_coupling_rows(rows: List[Dict[str, str]]) -> str:
    """
    NCツールページの「構成部品テーブル」から extension を連結文字列にする。
    kind表記が extension / subholder 等揺れても拾えるようにする。

    出力例: "EXT_A(L=25) / EXT_B(L=50)"
    """
    exts: List[str] = []
    for r in rows:
        kind = (r.get("coupling_type", "") or "").strip().lower()
        if kind in ("extension", "subholder", "ext"):
            name = (r.get("name", "") or "").strip()
            ln = (r.get("reach", "") or "").strip()
            if not name and not ln:
                continue
            if ln:
                exts.append(f"{name}(L={ln})" if name else f"(L={ln})")
            else:
                exts.append(name)
    return " / ".join([e for e in exts if e])


def _match_any(patterns: list[re.Pattern[str]], text: str) -> Optional[re.Match[str]]:
    for pat in patterns:
        m = pat.search(text)
        if m:
            return m
    return None


def parse_nctools_html(html_path: Path) -> Tuple[List[NcToolRecord], List[str]]:
    """
    hyperMILLのNCツールHTMLを解析して NcToolRecord のリストを返す。

    重要:
      - div.page を「3枚固定」とみなさない（サブホルダー等で崩れるため）
      - h3見出しで状態遷移して、NCツール開始→同一レコードに tool/holder/subholder を紐づける

    追加要件:
      - holder / extension / tool を別扱いにする（extensionはNCツールページの構成部品表から集計）
      - 全長 / extension突き出し / 工具突き出し / 突き出し長さ を算出
    """
    errors: List[str] = []
    html_text = html_path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(html_text, "lxml")

    pages = soup.select("div.page")
    if not pages:
        raise RuntimeError("div.page が見つかりません。HTML形式が想定と違います。")

    records: List[NcToolRecord] = []
    current: NcToolRecord | None = None

    def finalize_current():
        nonlocal current
        if current is None:
            return
        if not current.nctool_name.strip():
            current.nctool_name = "(UNKNOWN_NCTOOL)"
            current.warnings.append("nctool_name が空でした（解析失敗の可能性）")
        records.append(current)
        current = None

    for page_idx, p in enumerate(pages, start=1):
        h3_el = p.find("h3")
        h3 = clean_text(h3_el.get_text(" ", strip=True)) if h3_el else ""

        # -----------------------
        # NCツール開始 (JA/EN)
        # -----------------------
        m_nct = _match_any([_RE_NCTOOL_H3_JA, _RE_NCTOOL_H3_EN], h3)
        if m_nct:
            finalize_current()

            current = NcToolRecord(source_html_path=str(html_path))
            current.nctool_name = clean_text(m_nct.group(1))
            current.nctool_no = int(m_nct.group(2))

            # コメント（2番目table想定）
            nct_tables = p.find_all("table")
            if len(nct_tables) >= 2:
                kv = _parse_kv_table(nct_tables[1])
                current.nctool_comment = kv.get("nctool_comment", "")
            else:
                current.warnings.append("NCツールページのtableが不足しています")

            # 構成部品（border=1）: holder / extension / tool を拾う
            coupling_table = None
            for t in p.find_all("table"):
                if t.get("border") == "1":
                    coupling_table = t
                    break

            if coupling_table:
                rows = _parse_grid_table(coupling_table)

                holder_len = None
                tool_len = None
                ext_sum = 0.0
                ext_found = False

                for r in rows:
                    kind = (r.get("coupling_type", "") or "").strip().lower()
                    name = r.get("name", "") or ""
                    ln_str = r.get("reach", "") or ""
                    ln = _to_float_mm(ln_str)

                    if kind == "holder":
                        current.holder_name = name
                        current.holder_length = ln_str
                        holder_len = ln

                    elif kind == "tool":
                        current.tool_name = name
                        current.tool_length = ln_str
                        tool_len = ln

                    elif kind in ("extension", "subholder", "ext"):
                        if ln is not None:
                            ext_sum += ln
                        ext_found = True

                # extension表示文字列
                current.extensions_str = _build_extensions_str_from_coupling_rows(rows)

                # --- 計算値 ---
                current.ext_overhang_mm = _fmt_mm(ext_sum) if ext_found else "0"
                current.tool_overhang_mm = _fmt_mm(tool_len) if tool_len is not None else ""
                overhang = (ext_sum + tool_len) if tool_len is not None else None
                current.overhang_mm = _fmt_mm(overhang) if overhang is not None else ""

            else:
                current.warnings.append("NCツールページの構成部品テーブル（border=1）が見つかりません")

            # 画像（このページ内の img を拾う）
            img = p.find("img")
            if img and img.get("src"):
                current.image_rel_src = img["src"]
            else:
                current.warnings.append("NCツールページの画像srcが見つかりません")

            continue

        # NCツール開始前のページは無視
        if current is None:
            continue

        # -----------------------
        # Tool page (JA/EN)
        # -----------------------
        m_tool = _match_any([_RE_TOOL_H3_JA, _RE_TOOL_H3_EN], h3)
        if m_tool:
            current.tool_page_name = clean_text(m_tool.group(1))
            current.tool_type = clean_text(m_tool.group(2))

            tool_tables = p.find_all("table")
            if tool_tables:
                kv = _parse_kv_table(tool_tables[0])

                current.tool_diameter_mm = kv.get("diameter", "")
                current.tool_corner_radius_mm = kv.get("corner_radius", "")
                current.tool_flutes = kv.get("cutting_edges", "")
                current.tool_cut_length_ap_mm = kv.get("cutting_length", "")
                current.tool_shank_d_mm = kv.get("shank_diameter", "")
                current.tool_chamfer_len_mm = kv.get("chamfer_length", "")
                current.tool_tip_len_mm = kv.get("tip_length", "")
                current.tool_taper_angle_deg = kv.get("cone_angle", "")
                current.spindle_rotation = kv.get("spindle_orientation", "")
            else:
                current.warnings.append("工具ページの寸法tableが見つかりません")

            # 条件（F2では不要だが先頭行だけ保持）
            cond_table = None
            for t in tool_tables[1:]:
                if t.get("border") == "1":
                    cond_table = t
                    break
            if cond_table:
                cond_rows = _parse_grid_table(cond_table)
                if cond_rows:
                    c0 = cond_rows[0]
                    # ここは英語HTMLでも列名が "S (n)" 等のままなのでそのままでOK
                    current.cond_S_n = c0.get("S (n)", "")
                    current.cond_FX = c0.get("FX", "")
                    current.cond_FZ = c0.get("FZ", "")
                    current.cond_Fr = c0.get("Fr", "")
                    current.cond_ap = c0.get("ap", "")
                    current.cond_ae = c0.get("ae", "")
            else:
                current.warnings.append("工具ページの条件テーブル（border=1）が見つかりません")

            continue

        # -----------------------
        # Holder page (JA/EN)
        # -----------------------
        m_holder = _match_any([_RE_HOLDER_H3_JA, _RE_HOLDER_H3_EN], h3)
        if m_holder:
            current.holder_page_name = clean_text(m_holder.group(1))
            holder_tables = p.find_all("table")
            if holder_tables:
                kvh = _parse_kv_table(holder_tables[0])
                current.holder_comment = kvh.get("holder_comment", "")
            else:
                current.warnings.append("ホルダーページのtableが見つかりません")
            continue

        # -----------------------
        # Subholder/Extension page（ズレ原因ページ）
        # extension自体はNCツールページで拾っているので、ここでは検出ログ程度
        # -----------------------
        m_sub = _RE_SUBHOLDER_H3_JA.search(h3)
        if m_sub:
            name = clean_text(m_sub.group(1))
            current.warnings.append(f"サブホルダーページ検出: {name}")
            continue

        m_ext = _RE_EXTENSION_H3_EN.search(h3)
        if m_ext:
            name = clean_text(m_ext.group(1))
            current.warnings.append(f"Extension page detected: {name}")
            continue

        continue

    finalize_current()
    return records, errors