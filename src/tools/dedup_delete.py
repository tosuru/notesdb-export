#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重複ファイル削除ツール
=================================
指定したルートディレクトリ以下の全ファイルを走査し、
「同一内容のファイル」をグルーピングして、1つだけ残して残りを削除します。
デフォルトは dry-run で削除は行いません。--apply を付けると実際に削除します。

■ “同一”の定義（安全＆高速）
 1) ファイルサイズが同一
 2) 先頭1MBのハッシュが同一（高速フィルタ）
 3) 全体ハッシュ（BLAKE2b）が同一  ※実質的な同一性判定
   → ここまで一致したものを「重複」とみなします。
   → オプションで「更新日時（mtime）も同一」を厳密条件に加えることができます（--strict-same-mtime）。

■ 残すファイルの選び方（優先度）
   A) ファイル名の末尾に "_2", "(1)", " - Copy", " のコピー", "コピー" 等の「複製っぽい印」が *ない* ものを優先
   B) 文字数が短いものを優先
   C) 上記が同点なら、パスの辞書順で最小のもの

■ 使用例
  ドライラン（削除しない）:
    python dedup_delete.py "C:\data"
  実際に削除（※十分ご注意ください）:
    python dedup_delete.py "/home/user/downloads" --apply

  更新日時も同一のものだけを重複とみなす厳密判定 + 実削除：
    python dedup_delete.py "/data" --strict-same-mtime --apply

■ 出力
  ・処理ログ（標準出力）
  ・CSVレポート（dedup_report_YYYYmmdd_HHMMSS.csv）: keep/削除対象・サイズ・mtime・ハッシュ 等

※注意
  ・ゴミ箱へは送りません（本当に削除します）。戻せない可能性があるため、まずは dry-run で結果を確認してください。
  ・大容量のファイルが多数ある場合は時間がかかります。
"""

import os
import sys
import re
import csv
import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Iterable, Optional
from datetime import datetime, timezone
import hashlib

# ============== ユーティリティ ==============


def human_size(num: int) -> str:
    """バイト数を人間に読みやすい単位へ変換"""
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if num < 1024 or unit == "TB":
            return f"{num:.0f} {unit}"
        num /= 1024
    return f"{num:.0f} TB"  # フォールバック


def file_mtime(path: Path) -> float:
    """ファイルの最終更新時刻（エポック秒）"""
    try:
        return path.stat().st_mtime
    except Exception:
        return 0.0


def mtime_iso(path: Path) -> str:
    """ISO形式の更新日時文字列"""
    try:
        ts = path.stat().st_mtime
        return datetime.fromtimestamp(ts).astimezone().isoformat(timespec="seconds")
    except Exception:
        return ""


def chunk_reader(f, chunk_size: int = 1024 * 1024) -> Iterable[bytes]:
    """大きなファイルでも少ないメモリで読み込むためのチャンクリーダ"""
    while True:
        data = f.read(chunk_size)
        if not data:
            break
        yield data


def quick_hash(path: Path, first_bytes: int = 1024 * 1024) -> Optional[str]:
    """先頭 first_bytes の高速ハッシュ（BLAKE2b）。アクセス不能時は None。"""
    try:
        h = hashlib.blake2b(digest_size=20)
        with path.open("rb") as f:
            data = f.read(first_bytes)
            h.update(data)
        return h.hexdigest()
    except Exception as e:
        print(f"[WARN] quick_hash失敗: {path} ({e})")
        return None


def full_hash(path: Path, chunk_size: int = 1024 * 1024) -> Optional[str]:
    """全体ハッシュ（BLAKE2b）。アクセス不能時は None。"""
    try:
        h = hashlib.blake2b(digest_size=32)
        with path.open("rb") as f:
            for chunk in chunk_reader(f, chunk_size):
                h.update(chunk)
        return h.hexdigest()
    except Exception as e:
        print(f"[WARN] full_hash失敗: {path} ({e})")
        return None


# “複製っぽい”名前の判定用パターン（拡張子を除いたベース名に対して使用）
COPYLIKE_PATTERNS = [
    r".*[_\-\s]\d+$",            # 末尾が _2 / -2 /  2 など
    r".*\(\d+\)$",               # 末尾が (1) (2)
    r".*copy$",
    r".*コピー$",
    r".*のコピー$",
    r".* - copy$",               # " - Copy"（OSにより表記差あり）
    r".* - コピー$",
]


def is_copylike_name(stem: str) -> bool:
    s = stem.lower()
    for pat in COPYLIKE_PATTERNS:
        if re.fullmatch(pat, s, flags=re.IGNORECASE):
            return True
    return False


def keep_priority_key(p: Path) -> Tuple[int, int, str]:
    """
    ファイルを残す優先度を決めるキー（tuple の最小が勝ち）
      1) 複製っぽい名前の有無（False=0 < True=1）
      2) ファイル名の長さ（短い方が優先）
      3) パスの辞書順（安定化用）
    """
    stem = p.stem  # 拡張子除いたファイル名
    copylike = is_copylike_name(stem)
    return (1 if copylike else 0, len(p.name), str(p).lower())

# ============== メインロジック ==============


def collect_files(root: Path) -> List[Path]:
    """ルート以下の通常ファイルをすべて取得"""
    files: List[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        for name in filenames:
            p = Path(dirpath) / name
            # シンボリックリンクはスキップ（必要に応じて変更）
            try:
                if not p.is_file() or p.is_symlink():
                    continue
            except Exception:
                continue
            files.append(p)
    return files


def group_by_size(paths: List[Path]) -> Dict[int, List[Path]]:
    """サイズでグループ化"""
    groups: Dict[int, List[Path]] = {}
    for p in paths:
        try:
            sz = p.stat().st_size
        except Exception:
            continue
        groups.setdefault(sz, []).append(p)
    return groups


def group_duplicates(
    same_size_files: List[Path],
    strict_same_mtime: bool = False
) -> Dict[str, List[Path]]:
    """
    同一内容の重複グループを返す。
    1) 先頭1MBハッシュで粗くグループ化
    2) 全体ハッシュで確定グループ化
    3) strict_same_mtime=True の場合は mtime も同一のものだけ同一グループにする
    """
    # 1) quick hash
    qh_groups: Dict[str, List[Path]] = {}
    for p in same_size_files:
        qh = quick_hash(p)
        if qh is None:
            continue
        qh_groups.setdefault(qh, []).append(p)

    # 2) full hash
    full_groups: Dict[str, List[Path]] = {}
    for _, members in qh_groups.items():
        if len(members) < 2:
            continue
        for p in members:
            fh = full_hash(p)
            if fh is None:
                continue
            full_groups.setdefault(fh, []).append(p)

    # 3) mtime でさらに分割（必要なら）
    if strict_same_mtime:
        split_groups: Dict[str, List[Path]] = {}
        for h, members in full_groups.items():
            if len(members) < 2:
                continue
            # mtime（秒）ごとにキーを分ける
            bucket: Dict[int, List[Path]] = {}
            for p in members:
                try:
                    mt = int(p.stat().st_mtime)
                except Exception:
                    mt = -1
                bucket.setdefault(mt, []).append(p)
            for mt, lst in bucket.items():
                if len(lst) >= 2:
                    split_groups[f"{h}|mt={mt}"] = lst
        return split_groups
    else:
        # 2個以上のものだけを返す
        return {h: lst for h, lst in full_groups.items() if len(lst) >= 2}


def decide_keep_and_delete(members: List[Path]) -> Tuple[Path, List[Path]]:
    """残す1つと、削除対象を返す"""
    keep = sorted(members, key=keep_priority_key)[0]
    delete = [p for p in members if p != keep]
    return keep, delete


def write_report(report_path: Path, rows: List[List[str]]) -> None:
    """CSVレポートを書き出し"""
    with report_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "group_id", "action", "path", "size_bytes", "size_human", "mtime", "hash"
        ])
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="重複ファイルを検出し、1つ残して削除します（デフォルトはdry-run）。")
    parser.add_argument("root", type=Path, help="ルートディレクトリ")
    parser.add_argument("--apply", action="store_true",
                        help="実際に削除を行う（指定がなければdry-run）")
    parser.add_argument("--strict-same-mtime",
                        action="store_true", help="更新日時（秒）が同一のものだけを重複とみなす")
    parser.add_argument("--report", type=Path, default=None,
                        help="CSVレポートの出力先（省略時はカレントに自動生成）")
    args = parser.parse_args()

    root: Path = args.root
    if not root.exists() or not root.is_dir():
        print(f"[ERROR] ルートディレクトリが存在しません: {root}")
        return 2

    # レポートパス
    if args.report is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = Path.cwd() / f"dedup_report_{ts}.csv"
    else:
        report_path = args.report

    print(f"[INFO] 走査開始: {root}")
    files = collect_files(root)
    print(f"[INFO] 発見ファイル数: {len(files)}")

    by_size = group_by_size(files)
    candidate_sizes = [sz for sz, lst in by_size.items() if len(lst) >= 2]
    print(f"[INFO] サイズ一致で候補グループ数: {len(candidate_sizes)}")

    all_rows: List[List[str]] = []
    total_dupe_files = 0
    total_delete = 0
    group_idx = 0

    for sz in sorted(candidate_sizes):
        same_size_files = by_size[sz]
        # 重複グルーピング
        groups = group_duplicates(
            same_size_files, strict_same_mtime=args.strict_same_mtime)
        for hkey, members in groups.items():
            group_idx += 1
            print(
                f"\n[GROUP #{group_idx}] サイズ={human_size(sz)} / 候補数={len(members)}")
            for p in members:
                print(f"  - {p} (mtime={mtime_iso(p)})")
            keep, delete_list = decide_keep_and_delete(members)
            print(f"  -> 残す: {keep.name}")
            for delp in delete_list:
                print(f"     削除対象: {delp.name}")

            # レポート行
            for p in members:
                action = "keep" if p == keep else (
                    "delete" if args.apply else "would_delete")
                all_rows.append([
                    str(group_idx), action, str(p), str(sz), human_size(
                        sz), mtime_iso(p), hkey.split("|")[0]
                ])

            total_dupe_files += len(members)
            total_delete += len(delete_list)

            # 削除実行
            if args.apply:
                for delp in delete_list:
                    try:
                        delp.unlink()
                        print(f"[DEL] {delp}")
                    except Exception as e:
                        print(f"[WARN] 削除失敗: {delp} ({e})")

    # レポート出力
    write_report(report_path, all_rows)
    print("\n[SUMMARY]")
    print(f"  重複グループ数: {group_idx}")
    print(
        f"  重複ファイル総数: {total_dupe_files}（うち削除{total_delete if args.apply else f'想定{total_delete}'}）")
    print(f"  レポート: {report_path}")

    if not args.apply:
        print("\n[NOTE] 現在は dry-run です。削除を実行するには --apply を付けて再実行してください。")
    return 0


if __name__ == "__main__":
    # テスト用コード
    root = Path("C://Users//A512292//Box//ISZJ_Group_TF1000//TF1生産企画DB")
    if not root.exists() or not root.is_dir():
        print(f"[ERROR] ルートディレクトリが存在しません: {root}")
        sys.exit(2)
    # --strict-same-mtime --apply　を付けて実行する場合の例
    sys.argv.extend([str(root), "--strict-same-mtime", "--apply"])
    
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n[INFO] 中断されました。")
        sys.exit(130)
