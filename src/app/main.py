from __future__ import annotations
import argparse
import sys
import os
import logging
import json
from dotenv import load_dotenv
from typing import List, Dict, Any
from pathlib import Path
from xml.etree import ElementTree as ET  # normalize のエラー処理用にインポート
from tqdm import tqdm  # normalize/render/extract の進捗表示用にインポート

load_dotenv(override=True)  # .env ファイルの環境変数をロード

# 相対パス→絶対パスに変換
PYTHONPATH = os.getenv("PYTHONPATH")
if PYTHONPATH is not None:
    abs_path = os.path.abspath(PYTHONPATH)
    if abs_path not in sys.path:
        sys.path.append(abs_path)

# プロジェクトルート/src を import パスに追加
ROOT = Path(__file__).resolve().parent
if (ROOT / "src").exists() and str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))
elif str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# このモジュール用のロガーをセットアップ
logger = logging.getLogger(__name__)


def setup_logging():
    """Setup root logger with console and file handlers."""
    log_level_name = os.environ.get("LOG_LEVEL", "INFO").upper()
    log_level = getattr(logging, log_level_name, logging.INFO)

    # ルートロガーを取得
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)  # ルートは最も低いレベルに設定

    # 既存のハンドラをクリア
    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    # コンソールハンドラ
    console_handler = logging.StreamHandler(sys.stdout)
    # コンソールはシンプルに
    console_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)
    console_handler.setLevel(log_level)
    root_logger.addHandler(console_handler)

    # ファイルハンドラ
    log_file_path = Path("logs/app.log")
    log_file_path.parent.mkdir(exist_ok=True)
    # ファイルには詳細な情報を
    file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s(%(lineno)d) - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    file_handler.setLevel(logging.DEBUG)  # ファイルにはDEBUG以上を記録
    root_logger.addHandler(file_handler)

    logger.info(f"Log level set to {log_level_name} for console.")


# Notes パイプライン（ライブラリ化されたモジュールをインポート）
try:
    import pipelines.flows
    from pipelines.flows import (
        run_unified,
        run_from_manifest,
        run_decoupled_export,
        # run_decoupled_process は main.py のサブコマンドに分割されたため不要
    )
    NOTES_FLOWS_AVAILABLE = True
except ImportError as e:
    logger.warning(
        f"Failed to import 'pipelines.flows'. Notes-related commands may fail. Error: {e}")
    run_unified = None
    run_from_manifest = None
    run_decoupled_export = None
    # run_decoupled_process = None # 削除
    NOTES_FLOWS_AVAILABLE = False

try:
    from core.dxl.parser import parse_dxl_document_from_string
except ImportError as e:
    logger.error(
        f"Failed to import 'core.dxl.parser'. The 'normalize' command will fail. Error: {e}")
    parse_dxl_document_from_string = None

try:
    from core.render.engine import render as core_render
except ImportError as e:
    logger.error(
        f"Failed to import 'core.render.engine'. The 'render' command will fail. Error: {e}")
    core_render = None

try:
    from core.attachments import extract_and_save_json_paths
except ImportError as e:
    logger.error(
        f"Failed to import 'core.attachments'. The 'extract' command will fail. Error: {e}")
    extract_and_save_json_paths = None


def _split_inputs(paths: List[str]) -> List[str]:
    # 既存挙動と同じ：パス配列をそのまま受ける（globはシェルに任せる）
    return paths


def _load_profiles(path: Path) -> Dict[str, Any]:
    import json
    obj = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(obj, dict) and "profiles" in obj:
        profiles = obj["profiles"]
    elif isinstance(obj, list):
        profiles = obj
    else:
        raise ValueError(
            "profiles JSON must be an array or an object with 'profiles' array")
    indexed: Dict[str, Any] = {}
    for p in profiles:
        name = p.get("name") or f"profile_{len(indexed)+1}"
        indexed[name] = p
    return indexed


def _exec_profile(p: Dict[str, Any]):
    """
    プロファイル辞書を実行します。
    """
    ptype = p.get("type", "")
    if ptype in ("normalize", "render"):
        # --- 'normalize' を新ハンドラにディスパッチ ---
        if ptype == "normalize":
            if not parse_dxl_document_from_string:  # 互換性チェック
                raise RuntimeError(
                    "core.dxl.parser.parse_dxl_document_from_string is not available")
            logging.info("Running 'normalize' profile via new handler...")
            # 'inputs' は [0] がディレクトリだと仮定
            # (より堅牢にするにはプロファイル仕様の見直しが必要)
            if not p.get("inputs") or not isinstance(p.get("inputs"), list) or len(p.get("inputs")) == 0:
                raise ValueError(
                    "Profile 'normalize' requires 'inputs' (list) with at least one directory path.")

            # main.py の 'normalize' サブコマンドと同じ args オブジェクトを構築
            norm_args = argparse.Namespace(
                dxl_dir=Path(p["inputs"][0]),
                json_dir=Path(p["out"]),
                # 'normalize' コンドのデフォルトと合わせる
                db_title=p.get("db_title", "DXL_Import")
            )
            _handle_normalize(norm_args)  # 新しいハンドラを呼び出す
        else:
            # ---'render' を新ハンドラにディスパッチ ---
            if not core_render:  # 互換性チェック
                raise RuntimeError(
                    "core.render.engine.render is not available")
            logging.info("Running 'render' profile via new handler...")
            if not p.get("inputs") or not isinstance(p.get("inputs"), list) or len(p.get("inputs")) == 0:
                raise ValueError(
                    "Profile 'render' requires 'inputs' (list) with at least one directory path.")

            # 古いプロファイル(fmt: "html"|"md"|"both")を新しいI/F(formats: "html,md")に簡易変換
            legacy_fmt = p.get("fmt", "html,md,docx,pdf")
            if legacy_fmt == "both":
                formats_str = "html,md"
            elif legacy_fmt in ("html", "md", "pdf", "docx"):
                formats_str = legacy_fmt
            else:
                formats_str = legacy_fmt  # カンマ区切りと仮定

            # main.py の 'render' サブコマンドと同じ args オブジェクトを構築
            render_args = argparse.Namespace(
                json_dir=Path(p["inputs"][0]),
                out=Path(p["out"]) if "out" in p and not p.get(
                    "inplace") else None,
                formats=formats_str,
                inplace=bool(p.get("inplace", False)),
                recursive=bool(p.get("recursive", True))  # 既定: 再帰ONでもOK
            )
            _handle_render(render_args)  # 新しいハンドラを呼び出す

        return

    # --- notes.* プロファイルを新ハンドラにディスパッチ ---
    if ptype in ("notes.unified", "notes.manifest", "notes.export", "notes.process"):
        if not NOTES_FLOWS_AVAILABLE:
            logging.error(
                f"Cannot execute profile type '{ptype}': 'pipelines.flows' module is not available.")
            sys.exit(1)

        if ptype == "notes.unified":
            logging.info("Running 'notes.unified' profile via new handler...")
            if "config" not in p:
                raise ValueError(
                    "Profile 'notes.unified' requires 'config' object.")
            db_config = p["config"]
            db_title = db_config.get("title", Path(db_config["db_file"]).stem)

            profile_args = argparse.Namespace(
                server=db_config.get("server", ""),
                db=db_config["db_file"],
                out=Path(p["out"]),
                state=Path(p["state"]),  # 'progress' ではなく 'state' を期待
                db_title=db_title,
                formats=",".join(
                    p.get("formats", ["html", "md", "docx", "pdf"])),
                no_keep_dxl=not bool(p.get("keep_dxl", True)),
                limit=p.get("limit"),
                retry_errors_only=bool(
                    p.get("retry_errors_only", False))  # ★ 追加
            )
            _handle_run_single_db(profile_args)

        elif ptype == "notes.manifest":
            logging.info("Running 'notes.manifest' profile via new handler...")
            profile_args = argparse.Namespace(
                manifest=Path(p["manifest"]),
                out=Path(p["out"]),
                state=Path(p["state"]),
                formats=",".join(
                    p.get("formats", ["html", "md", "docx", "pdf"])),
                no_keep_dxl=not bool(p.get("keep_dxl", True)),
                limit=p.get("limit"),
                retry_errors_only=bool(
                    p.get("retry_errors_only", False))  # ★ 追加
            )
            _handle_run_manifest(profile_args)

        elif ptype == "notes.export":
            logging.info("Running 'notes.export' profile via new handler...")
            if "config" not in p:
                raise ValueError(
                    "Profile 'notes.export' requires 'config' object.")
            db_config = p["config"]
            db_title = db_config.get("title", Path(db_config["db_file"]).stem)

            profile_args = argparse.Namespace(
                server=db_config.get("server", ""),
                db=db_config["db_file"],
                dxl_out=Path(p["dxl_out"]),
                db_title=db_title
            )
            _handle_export(profile_args)

        else:  # "notes.process"
            logging.info(
                "Running 'notes.process' profile via decoupled handlers...")
            # Decoupled 3ステップを実行
            # 1. Normalize
            norm_args = argparse.Namespace(
                dxl_dir=Path(p["dxl_in"]),
                json_dir=Path(p["json_out"]),  # プロファイルに json_out が必要
                db_title=p["db_title"]
            )
            _handle_normalize(norm_args)
            # 2. Extract
            extract_args = argparse.Namespace(
                dxl_dir=Path(p["dxl_in"]),
                json_dir=Path(p["json_out"]),  # normalize の出力先
                attach_dir=Path(p["attach_out"])  # プロファイルに attach_out が必要
            )
            _handle_extract(extract_args)
            # 3. Render
            render_args = argparse.Namespace(
                json_dir=Path(p["inputs"][0]),
                out=Path(p["out"]) if "out" in p and not p.get(
                    "inplace") else None,
                formats=formats_str,
                inplace=bool(p.get("inplace", False)),
                recursive=bool(p.get("recursive", True))  # 既定: 再帰ONでもOK
            )
            _handle_render(render_args)
        return

    raise ValueError(f"Unknown profile type: {ptype}")


def _handle_run_manifest(args):
    if not run_from_manifest:
        logging.error(
            "'run_from_manifest' from 'pipelines.flows' is not available. Cannot execute.")
        sys.exit(1)

    logging.info(f"Executing run-manifest for: {args.manifest}")
    logging.info(
        f"Output: {args.out}, State: {args.state}, Formats: {args.formats}")

    formats_list = [f.strip() for f in args.formats.split(',') if f.strip()]
    keep_dxl = not args.no_keep_dxl

    try:
        run_from_manifest(
            manifest_path=args.manifest,
            out_base=args.out,
            state_base=args.state,
            formats=formats_list,
            keep_dxl=keep_dxl,
            limit=args.limit,
            retry_errors_only=getattr(args, "retry_errors_only", False)  # ★ 追加
        )
        logging.info("run-manifest completed.")
    except Exception as e:
        logging.error(f"run-manifest failed: {e}", exc_info=True)
        sys.exit(1)


def _handle_run_single_db(args):
    if not run_unified:
        logging.error(
            "'run_unified' from 'pipelines.flows' is not available. Cannot execute.")
        sys.exit(1)

    logging.info(f"Executing run-single-db for: {args.db}")
    logging.info(
        f"Output: {args.out}, State: {args.state}, Formats: {args.formats}")

    db_title_safe = args.db_title or Path(args.db).stem
    db_config = {
        "server": args.server,
        "db_file": args.db,
        "title": db_title_safe
    }
    # 規約ベースで progress_path を決定
    progress_path = args.state / f"progress_{db_title_safe}.jsonl"
    formats_list = [f.strip() for f in args.formats.split(',') if f.strip()]
    keep_dxl = not args.no_keep_dxl

    try:
        progress_path.parent.mkdir(
            parents=True, exist_ok=True)  # state ディレクトリ作成
        run_unified(
            db_config=db_config,
            out_base=args.out,
            progress_path=progress_path,
            formats=formats_list,
            keep_dxl=keep_dxl,
            limit=args.limit
        )
        logging.info("run-single-db completed.")
    except Exception as e:
        logging.error(f"run-single-db failed: {e}", exc_info=True)
        sys.exit(1)


def _handle_export(args):
    if not run_decoupled_export:
        logging.error(
            "'run_decoupled_export' from 'pipelines.flows' is not available. Cannot execute.")
        sys.exit(1)

    logging.info(f"Executing export for: {args.db}")
    logging.info(f"DXL Output: {args.dxl_out}")

    db_title_safe = args.db_title or Path(args.db).stem
    db_config = {
        "server": args.server,
        "db_file": args.db,
        "title": db_title_safe
    }

    try:
        run_decoupled_export(
            db_config=db_config,
            dxl_out=args.dxl_out  # dxl_out_dir -> dxl_out に修正
        )
        logging.info("export completed.")
    except Exception as e:
        logging.error(f"export failed: {e}", exc_info=True)
        sys.exit(1)


def _handle_normalize(args):
    """
    DXL -> 初期 JSON の変換を実行します。
    (core.dxl.parser を呼び出します)
    """
    if not parse_dxl_document_from_string:
        logging.error(
            "DXL Parser (core.dxl.parser) is not available. Cannot normalize.")
        sys.exit(1)

    logging.info(f"Executing normalize from: {args.dxl_dir}")
    logging.info(f"JSON Output: {args.json_dir}, DB Title: {args.db_title}")

    dxl_dir = Path(args.dxl_dir)
    json_dir = Path(args.json_dir)
    db_title = args.db_title

    if not dxl_dir.is_dir():
        logging.error(f"DXL input directory not found: {dxl_dir}")
        sys.exit(1)

    try:
        json_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        logging.error(f"Failed to create output directory {json_dir}: {e}")
        sys.exit(1)

    dxl_files = sorted(list(dxl_dir.glob("*.xml")))
    if not dxl_files:
        logging.warning(f"No *.xml files found in {dxl_dir}")
        return

    logging.info(f"Found {len(dxl_files)} DXL files to normalize...")

    success_count = 0
    fail_count = 0

    for dxl_path in tqdm(dxl_files, desc="Normalizing DXL"):
        try:
            # DXL ファイルの読み込み
            try:
                dxl_text = dxl_path.read_text(encoding='utf-8')
            except UnicodeDecodeError:
                dxl_text = dxl_path.read_text(
                    encoding='utf-8-sig', errors='ignore')

            # コアパーサーの呼び出し
            doc = parse_dxl_document_from_string(dxl_text, db_title=db_title)

            # JSON ファイルの書き出し
            out_path = json_dir / f"{dxl_path.stem}.json"
            with out_path.open("w", encoding="utf-8") as f:
                json.dump(doc, f, ensure_ascii=False, indent=2)

            success_count += 1

        except ET.ParseError as e:
            logging.error(f"Failed to parse XML: {dxl_path.name}. Error: {e}")
            fail_count += 1
        except Exception as e:
            logging.error(
                f"Failed to process file {dxl_path.name}: {e}", exc_info=False)
            logging.debug("Traceback:", exc_info=True)  # DEBUG レベルでのみスタックトレース
            fail_count += 1

    logging.info(
        f"Normalization complete. Success: {success_count}, Failed: {fail_count}")
# --- ▲▲▲ 修正 ▲▲▲ ---

# --- ▼▼▼ 修正: _handle_extract にロジックを実装 ---


def _handle_extract(args):
    """
    添付ファイルを抽出し、JSON ファイルのパス情報を更新します。
    (core.attachments を呼び出します)
    """
    if not extract_and_save_json_paths:
        logging.error(
            "Attachment Extractor (core.attachments) is not available. Cannot extract.")
        sys.exit(1)

    logging.info(f"Executing extract using JSONs from: {args.json_dir}")
    logging.info(
        f"DXL Input: {args.dxl_dir}, Attachment Output: {args.attach_dir}")

    json_dir = Path(args.json_dir)
    dxl_dir = Path(args.dxl_dir)
    attach_dir = Path(args.attach_dir)

    if not json_dir.is_dir():
        logging.error(f"JSON input directory not found: {json_dir}")
        sys.exit(1)
    if not dxl_dir.is_dir():
        logging.error(f"DXL input directory not found: {dxl_dir}")
        sys.exit(1)

    try:
        attach_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        logging.error(
            f"Failed to create attachment output directory {attach_dir}: {e}")
        sys.exit(1)

    json_files = sorted(list(json_dir.glob("*.json")))
    if not json_files:
        logging.warning(f"No *.json files found in {json_dir}")
        return

    logging.info(
        f"Found {len(json_files)} JSON files to process for attachments...")

    success_count = 0
    fail_count = 0

    for json_path in tqdm(json_files, desc="Extracting Attachments"):
        try:
            # 対応する DXL ファイルを見つける
            dxl_path = dxl_dir / f"{json_path.stem}.xml"
            if not dxl_path.is_file():
                logging.warning(
                    f"Skipping {json_path.name}: Corresponding DXL file not found at {dxl_path}")
                fail_count += 1
                continue

            # core.attachments の関数を呼び出す
            # notes_to_outputs.py [cite: notes_to_outputs.py] のI/Fに合わせる
            # initial_json_path は、json ファイルの「名前」だけを渡す（このI/Fは将来改善の余地あり）
            # attachment_output_dir は、この *ドキュメント専用* のサブディレクトリパスを渡す

            # (ドキュメント専用の添付ファイル出力ディレクトリを作成)
            # JSON ファイル名からUNIDやタイトルを推測するのは困難なため、
            # JSON ファイルのステム名 (stem) をサブディレクトリとして使用する
            doc_attach_dir = attach_dir / json_path.stem
            doc_attach_dir.mkdir(parents=True, exist_ok=True)

            updated_doc = extract_and_save_json_paths(
                dxl_path=dxl_path,
                # I/F に合わせてファイル名 (str) を渡す
                initial_json_path=str(json_path.name),
                attachment_output_dir=doc_attach_dir
            )

            if updated_doc is None:
                raise RuntimeError("Attachment extraction returned None.")

            # 更新された辞書で JSON ファイルを上書き
            with json_path.open("w", encoding="utf-8") as f:
                json.dump(updated_doc, f, ensure_ascii=False, indent=2)

            success_count += 1

        except Exception as e:
            logging.error(
                f"Failed to extract attachments for {json_path.name}: {e}", exc_info=False)
            logging.debug("Traceback:", exc_info=True)  # DEBUG レベルでのみスタックトレース
            fail_count += 1

    logging.info(
        f"Attachment extraction complete. Success: {success_count}, Failed: {fail_count}")
# --- ▲▲▲ 修正 ▲▲▲ ---

# --- ▼▼▼ 修正: _handle_render にロジックを実装 ---


def _handle_render(args):
    """
    最終 JSON -> 各フォーマット へのレンダリングを実行します。
    (core.render.engine を呼び出します)
    """
    if not core_render:
        logging.error(
            "Render Engine (core.render.engine) is not available. Cannot render.")
        sys.exit(1)

    logging.info(f"Executing render from: {args.json_dir}")
    logging.info(
        f"Output: {getattr(args, 'out', None)}, Formats: {args.formats}")

    # json_dir = Path(args.json_dir)
    # out_dir = Path(args.out) if getattr(args, "out", None) else None

    # if not args.inplace:
    #     if out_dir is None:
    #         logging.error("--out is required unless --inplace is specified.")
    #         sys.exit(1)
    #     try:
    #         out_dir.mkdir(parents=True, exist_ok=True)
    #     except OSError as e:
    #         logging.error(f"Failed to create output directory {out_dir}: {e}")
    #         sys.exit(1)

    # if not json_dir.is_dir():
    #     logging.error(f"JSON input directory not found: {json_dir}")
    #     sys.exit(1)

    # try:
    #     out_dir.mkdir(parents=True, exist_ok=True)
    # except OSError as e:
    #     logging.error(f"Failed to create output directory {out_dir}: {e}")
    #     sys.exit(1)
    json_dir = Path(args.json_dir)
    json_files = sorted(list(json_dir.rglob("*.json"))) if args.recursive \
        else sorted(list(json_dir.glob("*.json")))
    if not json_files:
        logging.warning(f"No *.json files found in {json_dir}")
        return

    # logging.info(
    #     f"Found {len(json_files)} JSON files to render into {len(formats_list)} format(s)...")

    out_dir = Path(args.out) if getattr(args, "out", None) else None

    # --- 出力先の準備 ---
    if not args.inplace:
        # --out が必須
        if out_dir is None:
            logging.error("--out is required unless --inplace is specified.")
            sys.exit(1)
        # out_dir が None でない場合のみ mkdir
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            logging.error(f"Failed to create output directory {out_dir}: {e}")
            sys.exit(1)
    else:
        # inplace の場合は out_dir 不要
        logging.info(
            "Inplace mode: outputs will be written next to each JSON file.")

    formats_list = [f.strip() for f in args.formats.split(',') if f.strip()]

    for json_path in tqdm(json_files, desc="Rendering JSON"):
        for fmt in formats_list:
            try:
                rendered_data = core_render(doc_path=json_path, fmt=fmt)
                ext, mode, encoding = fmt, "w", "utf-8"
                if fmt == "pdf":
                    ext, mode, encoding = "pdf", "wb", None
                elif fmt == "docx":
                    ext, mode, encoding = "docx", "wb", None
                elif fmt == "html":
                    ext = "html"
                elif fmt == "md":
                    ext = "md"

                # ✅ inplace 指定時は JSON の隣に出力
                if args.inplace:
                    # out_path = json_path.parent / f"{json_path.stem}.{ext}"
                    # ".normalized" を削除して命名
                    out_path = json_path.parent / \
                        json_path.stem.replace(".normalized", f".{ext}")
                else:
                    # out_path = out_dir / f"{json_path.stem}.{ext}"
                    # .normalized を含む場合は削除
                    out_path = out_path.parent / \
                        json_path.stem.replace(".normalized", f".{ext}")

                # 書き込み（バイナリ or テキスト）
                if mode == "wb":
                    with open(out_path, mode) as f:
                        f.write(rendered_data)
                else:
                    with open(out_path, mode, encoding=encoding) as f:
                        f.write(rendered_data)

            except Exception as e:
                logging.exception(f"Failed to render {json_path} ({fmt}): {e}")


def _handle_run_profile(args, parser):  # parser を受け取る
    logging.debug(f"Handling run-profile command...")
    prof_file = args.profiles
    try:
        profs = _load_profiles(Path(prof_file))
    except FileNotFoundError:
        logging.error(f"Profile file not found: {prof_file}")
        sys.exit(1)
    except Exception as e:
        logging.error(
            f"Failed to load profiles from {prof_file}: {e}", exc_info=True)
        sys.exit(1)

    name = args.name or os.environ.get("PROFILE")
    if not name:
        if sys.stdin.isatty():
            names = list(profs.keys())
            if not names:
                logging.error(f"No profiles found in {prof_file}")
                sys.exit(1)
            print("Available profiles:")
            for i, nm in enumerate(names, 1):
                print(f"  [{i}] {nm}")
            try:
                sel = input("Select profile number or name: ").strip()
                idx = int(sel) - 1
                if 0 <= idx < len(names):
                    name = names[idx]
                else:  # 名前で選択されたか？
                    if sel in names:
                        name = sel
                    else:
                        raise ValueError("Invalid selection")
            except Exception:
                print("Invalid selection.")
                sys.exit(2)
        else:
            logging.error(
                "Profile name not specified. Use --name or set PROFILE env.")
            sys.exit(2)

    if name not in profs:
        logging.error(f"Profile '{name}' not found in {prof_file}")
        sys.exit(1)

    try:
        logging.info(f"--- Running profile: {name} ---")
        _exec_profile(profs[name])
        logging.info(f"--- Profile finished: {name} ---")
    except Exception as e:
        logging.error(f"Profile execution failed: {e}", exc_info=True)
        sys.exit(1)

# --- ▲▲▲ 新規追加 ▲▲▲ ---


def main():
    parser = argparse.ArgumentParser(
        prog="notes-doc", description="DXL -> JSON -> render + Notes pipeline profiles")

    # --- グローバルロガー設定 (main() の先頭に移動) ---
    setup_logging()

    # --- サブパーサーのセットアップ ---
    sub = parser.add_subparsers(
        dest="cmd", required=False, help="Sub-commands")

    # A-1. run-manifest
    p_manifest = sub.add_parser("run-manifest", help="Run full pipeline for multiple DBs from a manifest file",
                                formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p_manifest.add_argument("--manifest", required=True,
                            type=Path, help="Path to the manifest JSON file")
    p_manifest.add_argument("--out", required=True,
                            type=Path, help="Output base directory")
    p_manifest.add_argument("--state", required=True, type=Path,
                            help="State base directory (for progress.jsonl)")
    p_manifest.add_argument("--formats", default="html,md,docx,pdf",
                            help="Comma-separated list of output formats (e.g., 'html,pdf')")
    p_manifest.add_argument("--no-keep-dxl", action="store_true",
                            help="Do not save intermediate DXL files (default: save)")
    p_manifest.add_argument("--limit", type=int, default=None,
                            help="Limit processing to N documents per DB (default: no limit)")
    p_manifest.set_defaults(func=_handle_run_manifest)

    # A-2. run-single-db
    p_single = sub.add_parser("run-single-db", help="Run full pipeline for a single DB",
                              formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p_single.add_argument("--server", default="",
                          help="Notes server name (default: local)")
    p_single.add_argument(
        "--db", required=True, help="Notes database file path (e.g., 'path/to/db.nsf')")
    p_single.add_argument("--out", required=True,
                          type=Path, help="Output base directory")
    p_single.add_argument("--state", required=True, type=Path,
                          help="State directory (for progress.jsonl)")
    p_single.add_argument("--db-title", default=None,
                          help="Override DB title (default: derived from db file)")
    p_single.add_argument("--formats", default="html,md,docx,pdf",
                          help="Comma-separated list of output formats")
    p_single.add_argument("--no-keep-dxl", action="store_true",
                          help="Do not save intermediate DXL files (default: save)")
    p_single.add_argument("--limit", type=int, default=None,
                          help="Limit processing to N documents (default: no limit)")
    p_single.set_defaults(func=_handle_run_single_db)

    # B-1. export
    p_export = sub.add_parser("export", help="Step 1 (Decoupled): Export DXL files from a single DB",
                              formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p_export.add_argument("--server", default="",
                          help="Notes server name (default: local)")
    p_export.add_argument("--db", required=True,
                          help="Notes database file path")
    p_export.add_argument("--dxl-out", required=True,
                          type=Path, help="Output directory for DXL files")
    p_export.add_argument("--db-title", default=None,
                          help="Override DB title (default: derived from db file)")
    p_export.set_defaults(func=_handle_export)

    # B-2. normalize (引数変更)
    p_norm = sub.add_parser("normalize", help="Step 2 (Decoupled): Convert DXL files to initial JSON (v1.4.4+)",
                            formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p_norm.add_argument("--dxl-dir", required=True, type=Path,
                        help="Input directory containing DXL files (*.xml)")
    p_norm.add_argument("--json-dir", required=True, type=Path,
                        help="Output directory for initial JSON files (*.json)")
    p_norm.add_argument("--db-title", default="DXL_Import",
                        help="Database title to embed in JSON meta")
    p_norm.set_defaults(func=_handle_normalize)

    # B-3. extract (新規追加)
    p_extract = sub.add_parser("extract", help="Step 3 (Decoupled): Extract attachments and update JSON paths",
                               formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p_extract.add_argument("--dxl-dir", required=True, type=Path,
                           help="Input directory containing DXL files (*.xml)")
    p_extract.add_argument("--json-dir", required=True, type=Path,
                           help="Directory containing initial JSON files (*.json) (files will be updated)")
    p_extract.add_argument("--attach-dir", required=True, type=Path,
                           help="Output base directory for extracted attachments")
    p_extract.set_defaults(func=_handle_extract)

    # B-4. render (引数変更)
    p_render = sub.add_parser("render", help="Step 4 (Decoupled): Render final JSON files to output formats",
                              formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p_render.add_argument("--json-dir", required=True, type=Path,
                          help="Input directory containing final (updated) JSON files (*.json)")
    p_render.add_argument("--out", required=False, type=Path,
                          help="Output directory for rendered files (HTML, PDF, etc.). "
                               "Omit when using --inplace.")
    p_render.add_argument("--formats", default="html,md,docx,pdf",
                          help="Comma-separated list of output formats")
    p_render.add_argument("--inplace", action="store_true",
                          help="Write outputs next to each JSON file (ignore --out)")
    p_render.add_argument("--recursive", action="store_true",
                          help="Recursively search for *.json under --json-dir")
    p_render.set_defaults(func=_handle_render)

    # C. run-profile (既存維持)
    p_prof = sub.add_parser("run-profile", help="Run a profile by name from a JSON file (e.g., launch.json)",
                            formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p_prof.add_argument("--profiles", default="launch.json",
                        help="Path to profiles JSON file")
    p_prof.add_argument("--name", "-n", default=None,
                        help="Profile name to run (if omitted, list & prompt when TTY)")
    p_prof.set_defaults(func=lambda args: _handle_run_profile(args, parser))

    # --- 便宜ショートカット (既存維持) ---
    parser.add_argument(
        "--profiles", help="Run profile(s) from a JSON file. If --profile is also given, runs that one. If no command, this acts as 'run-profile' command.")
    parser.add_argument(
        "--profile", help="Profile name to run (requires --profiles)")

    args = parser.parse_args()

    # --- コマンド実行ディスパッチ ---

    # ショートカット動作 (run-profile)
    if args.profiles and args.cmd is None:
        logging.debug("Using profile shortcut...")
        # 擬似的な args オブジェクトを作成して _handle_run_profile を呼ぶ
        profile_args = argparse.Namespace(
            profiles=args.profiles,
            name=args.profile,  # --profile の値 (Noneかもしれない)
            func=lambda args: _handle_run_profile(args, parser)
        )
        profile_args.func(profile_args)
        return

    # サブコマンド実行
    if hasattr(args, 'func'):
        args.func(args)
    else:
        # サブコマンドもショートカットも指定されなかった場合
        logging.warning("No command specified.")
        parser.print_help()


if __name__ == "__main__":
    main()
