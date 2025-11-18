import os

# ======== 初期設定 ========

# 空フォルダ削除対象ディレクトリ（.env で指定）
# 例: TARGET_DIR=/Users/username/Desktop/test_dir
TARGET_DIR = r"C:\Users\A512292\Box\ext_ISZJ_生産部門会議_資料\【過去資料】生産部門_掲示板"

# 除外ディレクトリ（削除対象外）
EXCLUDE_DIRS = {'.git', '.venv', '__pycache__'}


def confirm_deletion():
    """削除実行前にユーザーに確認する"""
    while True:
        ans = input("⚠️ 本当に空フォルダを削除しますか？ (y/n): ").strip().lower()
        if ans in {"y", "yes"}:
            print("🧹 空フォルダ削除を開始します...\n")
            return True
        elif ans in {"n", "no"}:
            print("キャンセルしました。処理を終了します。")
            return False
        else:
            print("無効な入力です。y または n を入力してください。")


def remove_empty_dirs(target_dir):
    """指定ディレクトリ以下の空フォルダを削除し、結果を出力"""
    removed_count = 0
    skipped_count = 0

    # 下層フォルダから順に確認（topdown=False）
    for root, dirs, files in os.walk(target_dir, topdown=False):
        # 除外フォルダをスキップ
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]

        # ファイルもサブフォルダもないフォルダなら削除
        if not files and not dirs:
            try:
                os.rmdir(root)
                print(f"[✅ 削除] {root}")
                removed_count += 1
            except Exception as e:
                print(f"[⚠️ 失敗] {root} -> {e}")
                skipped_count += 1
        else:
            print(f"[⏭ 残す] {root}（ファイルまたはサブフォルダあり）")

    print("\n=== 削除結果 ===")
    print(f"削除したフォルダ数 : {removed_count}")
    print(f"削除できなかったフォルダ数 : {skipped_count}")
    print("================\n")


def main():
    """メイン処理"""
    if not TARGET_DIR:
        print("エラー: .env に TARGET_DIR が設定されていません。")
        return
    if not os.path.exists(TARGET_DIR):
        print(f"エラー: 指定されたディレクトリが存在しません: {TARGET_DIR}")
        return

    print("=== 空フォルダ削除スクリプト ===")
    print(f"対象ディレクトリ: {TARGET_DIR}\n")

    # 削除確認
    if confirm_deletion():
        remove_empty_dirs(TARGET_DIR)

    print("=== 処理完了 ===")


if __name__ == "__main__":
    main()
