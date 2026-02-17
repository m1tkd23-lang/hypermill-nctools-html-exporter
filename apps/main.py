from __future__ import annotations

import argparse
from pathlib import Path

from hypermill_nctools_html_exporter import export_from_html


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--html", required=True, help="input HTML path")
    ap.add_argument("--out", required=True, help="output directory")
    ap.add_argument("--no-embed", action="store_true", help="do not embed images (light mode)")
    ap.add_argument("--max-px", type=int, default=320, help="max image size (px) for cache/embed")
    args = ap.parse_args()

    html_path = Path(args.html)
    out_dir = Path(args.out)

    out_xlsx, summary = export_from_html(
        html_path=html_path,
        out_dir=out_dir,
        embed_images=(not args.no_embed),
        max_px=args.max_px,
    )
    print("OK:", out_xlsx)
    print(summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
