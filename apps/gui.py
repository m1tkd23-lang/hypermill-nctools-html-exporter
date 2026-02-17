from __future__ import annotations

import sys
import threading
import queue
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# src を import path に追加
sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from hypermill_nctools_html_exporter.core import export_report_f2_from_html


def main() -> int:
    root = tk.Tk()
    root.title("hypermill-nctools-html-exporter")
    root.geometry("860x480")

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    # HTML path
    ttk.Label(frm, text="入力HTML").grid(row=0, column=0, sticky="w")
    html_var = tk.StringVar()
    ttk.Entry(frm, textvariable=html_var, width=80).grid(row=0, column=1, padx=8, sticky="we")

    def choose_html():
        f = filedialog.askopenfilename(
            title="NCツールHTMLを選択",
            filetypes=[("HTML", "*.html;*.htm"), ("All", "*.*")],
        )
        if f:
            html_var.set(f)

    ttk.Button(frm, text="参照", command=choose_html).grid(row=0, column=2)

    # Out dir
    ttk.Label(frm, text="出力先フォルダ").grid(row=1, column=0, sticky="w")
    out_var = tk.StringVar(value=str(Path.cwd() / "out"))
    ttk.Entry(frm, textvariable=out_var, width=80).grid(row=1, column=1, padx=8, sticky="we")

    def choose_out():
        d = filedialog.askdirectory(title="出力先フォルダを選択")
        if d:
            out_var.set(d)

    ttk.Button(frm, text="参照", command=choose_out).grid(row=1, column=2)

    # max_px
    ttk.Label(frm, text="画像最大辺(px)").grid(row=2, column=0, sticky="w")
    maxpx_var = tk.StringVar(value="320")
    ttk.Entry(frm, textvariable=maxpx_var, width=10).grid(row=2, column=1, sticky="w")

    # progress
    status_var = tk.StringVar(value="待機中")
    prog = ttk.Progressbar(frm, orient="horizontal", mode="determinate")
    lbl = ttk.Label(frm, textvariable=status_var)

    prog.grid(row=4, column=0, columnspan=3, sticky="we", pady=(16, 0))
    lbl.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))

    frm.columnconfigure(1, weight=1)

    # worker queue
    q: queue.Queue[tuple[str, int, int, str]] = queue.Queue()
    busy = {"flag": False}

    def pump_queue():
        try:
            while True:
                kind, done, total, msg = q.get_nowait()
                if kind == "progress":
                    status_var.set(msg)
                    prog["maximum"] = max(1, total)
                    prog["value"] = done
                elif kind == "done":
                    busy["flag"] = False
                    status_var.set("完了")
                    messagebox.showinfo("完了", msg)
                elif kind == "error":
                    busy["flag"] = False
                    status_var.set("エラー")
                    messagebox.showerror("エラー", msg)
        except queue.Empty:
            pass
        root.after(100, pump_queue)

    root.after(100, pump_queue)

    def run_export():
        if busy["flag"]:
            messagebox.showwarning("実行中", "処理が実行中です。完了後に再実行してください。")
            return

        html_path = Path(html_var.get()).expanduser()
        out_dir = Path(out_var.get()).expanduser()

        if not html_path.exists():
            messagebox.showerror("入力エラー", f"HTMLが見つかりません:\n{html_path}")
            return

        try:
            max_px = int(maxpx_var.get())
            if max_px <= 0:
                raise ValueError()
        except Exception:
            messagebox.showerror("入力エラー", "画像最大辺(px) は正の整数で指定してください")
            return

        busy["flag"] = True
        prog["value"] = 0
        status_var.set("開始...")

        def worker():
            try:
                def progress(done, total, msg):
                    q.put(("progress", int(done), int(total), str(msg)))

                out_xlsx, _summary = export_report_f2_from_html(
                    html_path=html_path,
                    out_dir=out_dir,
                    embed_images=True,   # GUIでは埋め込み固定
                    max_px=max_px,       # GUI入力を反映
                    progress=progress,
                )
                q.put(("done", 1, 1, f"F2帳票を出力しました:\n{out_xlsx}"))
            except Exception as e:
                q.put(("error", 0, 1, str(e)))

        threading.Thread(target=worker, daemon=True).start()

    ttk.Button(frm, text="実行（HTML→XLSX）", command=run_export).grid(
        row=6, column=0, columnspan=3, sticky="we", pady=(14, 0)
    )

    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
