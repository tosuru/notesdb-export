
from __future__ import annotations
import json, os, threading
from pathlib import Path
from datetime import datetime, timezone
from typing import Dict, Iterable, Iterator, Optional

class JsonlProgress:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()

    @staticmethod
    def _now() -> str:
        return datetime.now(timezone.utc).isoformat(timespec="seconds")

    def append(self, db: str, unid: str, status: str, try_count: int = 0, err: Optional[str] = None, **extra) -> None:
        rec = {"ts": self._now(), "db": db, "unid": unid, "status": status, "try": int(try_count)}
        if err: rec["err"] = str(err)[:400]
        if extra: rec.update(extra)
        line = json.dumps(rec, ensure_ascii=False)
        with self._lock:
            with open(self.path, "a", encoding="utf-8", newline="\n") as f:
                f.write(line + "\n")
                f.flush()
                os.fsync(f.fileno())

    def snapshot(self) -> Dict[tuple[str, str], dict]:
        state: Dict[tuple[str, str], dict] = {}
        if not self.path.exists():
            return state
        with open(self.path, "r", encoding="utf-8") as f:
            for ln in f:
                s = ln.strip()
                if not s: continue
                try:
                    rec = json.loads(s)
                except json.JSONDecodeError:
                    continue
                key = (rec.get("db",""), rec.get("unid",""))
                state[key] = rec
        return state

    def remaining_unids(self, all_pairs: Iterable[tuple[str, str]], retry_max: int = 3) -> Iterator[tuple[str, str, int]]:
        st = self.snapshot()
        for db, unid in all_pairs:
            rec = st.get((db, unid))
            if not rec:
                yield (db, unid, 0); continue
            status = rec.get("status"); tried = int(rec.get("try", 0))
            if status in ("done","skipped"): continue
            if status == "error" and tried >= retry_max: continue
            yield (db, unid, tried)
