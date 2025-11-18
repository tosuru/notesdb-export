import os
import datetime


# 環境変数から設定を取得
TARGET_DIR = r"C:\Users\A512292\Box\ISZJ_Group_TF1000\TF1生産企画DB"
FILE_EXTENSION = ".html"
DELETE_AFTER_DATE = "20251105"  # YYYYMMDD形式 以降に作成されたファイルを削除

# 日付文字列をdatetimeに変換（例: 20251105 → 2025-11-05）
target_date = datetime.datetime.strptime(DELETE_AFTER_DATE, "%Y%m%d")


def should_delete(file_path):
    """削除対象かどうかを判定"""
    # 拡張子チェック
    if not file_path.lower().endswith(FILE_EXTENSION.lower()):
        return False

    # ファイル作成日を取得
    created_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path))

    # 指定日以降なら削除対象
    return created_time >= target_date


def delete_files(target_dir):
    """フォルダを再帰的に探索し、条件に一致するファイルを削除"""
    total = 0
    deleted = 0

    print(f"検索開始: {target_dir}")
    for root, _, files in os.walk(target_dir):
        for file in files:
            file_path = os.path.join(root, file)
            total += 1

            if should_delete(file_path):
                try:
                    os.remove(file_path)
                    deleted += 1
                    print(f"[削除] {file_path}")
                except Exception as e:
                    print(f"[エラー] {file_path} の削除に失敗しました: {e}")

    print(f"\n完了 ✅")
    print(f"総ファイル数: {total}")
    print(f"削除対象ファイル数: {deleted}")


if __name__ == "__main__":
    if not TARGET_DIR or not os.path.exists(TARGET_DIR):
        print("❌ 有効なフォルダパスを指定してください。")
    else:
        delete_files(TARGET_DIR)
