import json
from pathlib import Path

# ===== 設定 =====
INPUT_FILE = Path(r"C:\Users\A512292\Box\ISZJ_Group_TF1000\TF1生産企画DB\progress1.jsonl")
OUTPUT_FILE = Path(r"C:\Users\A512292\Box\ISZJ_Group_TF1000\TF1生産企画DB\progress.jsonl")
DELETE_KEYWORD = "TF1生産企画DB\\1\\週報"  # "out"にこの文字列が含まれていたら削除対象

# ===== メイン処理 =====
def main():
    print("JSONLフィルタリングを開始します...")

    lines = []
    with INPUT_FILE.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                lines.append(json.loads(line))

    print(f"読み込み完了: {len(lines)} 行")

    # 削除対象unidを収集
    delete_unids = set()
    for record in lines:
        out_value = record.get("out", "")
        if DELETE_KEYWORD in out_value:
            delete_unids.add(record["unid"])

    print(f"削除対象のUNID数: {len(delete_unids)}")

    # フィルタリング
    filtered = [rec for rec in lines if rec["unid"] not in delete_unids]

    print(f"出力行数: {len(filtered)} (削除: {len(lines) - len(filtered)})")

    # 結果を書き出し
    with OUTPUT_FILE.open("w", encoding="utf-8") as f:
        for rec in filtered:
            json.dump(rec, f, ensure_ascii=False)
            f.write("\n")

    print(f"完了: {OUTPUT_FILE} に保存しました。")
if __name__ == "__main__":
    main()
