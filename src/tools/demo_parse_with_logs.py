
#!/usr/bin/env python
from __future__ import annotations
import argparse
import glob
import os
import argparse
import glob
import os
import logging
from core.logs.report import ParseReport
from core.dxl.parser import parse_dxl_file


def main():
    ap = argparse.ArgumentParser(
        description="DXL parse demo with logging & summary")
    ap.add_argument("inputs", nargs="+",
                    help="DXL files or globs (e.g., tests/data/*.xml)")
    ap.add_argument("--json", dest="json_mode",
                    action="store_true", help="enable JSON console logs (not implemented with standard logging)")
    ap.add_argument("--level", default="INFO",
                    help="log level (DEBUG/INFO/WARN/ERROR)")
    ap.add_argument("--report", default="dist/parse_report.json",
                    help="summary json path")
    args = ap.parse_args()

    logging.basicConfig(level=args.level.upper(), format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    log = logging.getLogger("cli")

    rep = ParseReport()
    files = []
    for spec in args.inputs:
        files.extend(glob.glob(spec))
    for path in files:
        try:
            doc = parse_dxl_file(path)
            meta = (doc.get("meta") or {})
            links = (doc.get("links") or {})
            rep.add(path=path, ok=True, unid=meta.get("unid"),
                    subject=((doc.get("fields") or {}).get(
                        "Subject") or {}).get("value", ""),
                    warnings=0, notes_links=len(links.get("notes") or []),
                    http_links=len(links.get("http") or []))
        except Exception as e:
            log.error(f"file.failed: file={path}, error={str(e)}")
            rep.add(path=path, ok=False, error=str(e))

    os.makedirs(os.path.dirname(args.report), exist_ok=True)
    rep.write_json(args.report)
    # also CSV next to JSON
    rep.write_csv(os.path.splitext(args.report)[0] + ".csv")
    log.info(f"report.written: path={args.report}")


if __name__ == "__main__":
    main()
