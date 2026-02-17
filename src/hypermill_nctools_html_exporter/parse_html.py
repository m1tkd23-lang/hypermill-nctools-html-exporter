# src/hypermill_nctools_html_exporter/parse_html.py
from __future__ import annotations

import re
from pathlib import Path
from typing import List, Dict, Tuple

from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import warnings

from .model import NcToolRecord
from .util import clean_text

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

_RE_NCTOOL_H3 = re.compile(r"NCツール\(N\):(.+?)\s*\((\d+)\)\s*$")
_RE_TOOL_H3 = re.compile(r"工具:\s*(.+?)\s*\((.+?)\)\s*$")
_RE_HOLDER_H3 = re.compile(r"ホルダー:\s*(.+)\s*$")
_RE_SUBHOLDER_H3 = re.compile(r"サブホルダー:\s*(.+)\s*$")


def _parse_kv_table(table) -> Dict[str, str]:
    d: Dict[str, str] = {}
    for tr in table.find_all("tr"):
        tds = [clean_text(td.get_text(" ", strip=True)) for td in tr.find_all("td")]
        if len(tds) == 2 and tds[0]:
            d[tds[0]] = tds[1]
        elif len(tds) == 4:
            if tds[0]:
                d[tds[0]] = tds[1]
            if tds[2]:
                d[tds[2]] = tds[3]
    return d


def _parse_grid_table(table) -> List[Dict[str, str]]:
    trs = table.find_all("tr")
    if not trs:
        return []
    header = [clean_text(td.get_text(" ", strip=True)) for td in trs[0].find_all("td")]
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
        kind = (r.get("カップリング種類", "") or "").strip().lower()
        if kind in ("extension", "subholder", "ext"):
            name = (r.get("名称", "") or "").strip()
            ln = (r.get("全長", "") or "").strip()
            if not name and not ln:
                continue
            if ln:
                exts.append(f"{name}(L={ln})" if name else f"(L={ln})")
            else:
                exts.append(name)
    return " / ".join([e for e in exts if e])


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
        # NCツール開始
        # -----------------------
        m_nct = _RE_NCTOOL_H3.search(h3)
        if m_nct:
            finalize_current()

            current = NcToolRecord(source_html_path=str(html_path))
            current.nctool_name = clean_text(m_nct.group(1))
            current.nctool_no = int(m_nct.group(2))

            # コメント（2番目table想定）
            nct_tables = p.find_all("table")
            if len(nct_tables) >= 2:
                kv = _parse_kv_table(nct_tables[1])
                current.nctool_comment = kv.get("NCツール コメント", "")
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
                    kind = (r.get("カップリング種類", "") or "").strip().lower()
                    name = r.get("名称", "") or ""
                    ln_str = r.get("全長", "") or ""
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
                # extension突き出し（無いなら0）
                current.ext_overhang_mm = _fmt_mm(ext_sum) if ext_found else "0"

                # 工具突き出し（tool_lengthの数値版）
                current.tool_overhang_mm = _fmt_mm(tool_len) if tool_len is not None else ""

                # 突き出し長さ = ext + tool（toolが無いときは空）
                overhang = (ext_sum + tool_len) if tool_len is not None else None
                current.overhang_mm = _fmt_mm(overhang) if overhang is not None else ""


            else:
                current.warnings.append("NCツールページの構成部品テーブル（border=1）が見つかりません")

            # 画像
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
        # Tool page
        # -----------------------
        m_tool = _RE_TOOL_H3.search(h3)
        if m_tool:
            current.tool_page_name = clean_text(m_tool.group(1))
            current.tool_type = clean_text(m_tool.group(2))

            tool_tables = p.find_all("table")
            if tool_tables:
                kv = _parse_kv_table(tool_tables[0])
                current.tool_diameter_mm = kv.get("直径", "")
                current.tool_corner_radius_mm = kv.get("コーナー半径", "")
                current.tool_flutes = kv.get("刃数", "")
                current.tool_cut_length_ap_mm = kv.get("切削長さ (ap)", "")
                current.tool_shank_d_mm = kv.get("シャンク直径", "")
                current.tool_chamfer_len_mm = kv.get("面取り長さ", "")
                current.tool_tip_len_mm = kv.get("先端長さ", "")
                current.tool_taper_angle_deg = kv.get("テーパー角度", "")
                current.spindle_rotation = kv.get("スピンドル回転方向", "")
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
        # Holder page
        # -----------------------
        m_holder = _RE_HOLDER_H3.search(h3)
        if m_holder:
            current.holder_page_name = clean_text(m_holder.group(1))
            holder_tables = p.find_all("table")
            if holder_tables:
                kvh = _parse_kv_table(holder_tables[0])
                current.holder_comment = kvh.get("ホルダー コメント", "")
            else:
                current.warnings.append("ホルダーページのtableが見つかりません")
            continue

        # -----------------------
        # Subholder page（ズレ原因ページ）
        # extension自体はNCツールページで拾っているので、ここでは検出ログ程度
        # -----------------------
        m_sub = _RE_SUBHOLDER_H3.search(h3)
        if m_sub:
            name = clean_text(m_sub.group(1))
            current.warnings.append(f"サブホルダーページ検出: {name}")
            continue

        # その他ページは無視（安全）
        continue

    finalize_current()
    return records, errors
