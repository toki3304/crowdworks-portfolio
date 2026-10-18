"""
動作確認用のダミー売上データ（sample_data/）を作るスクリプト

店舗ごと・月ごとに1ファイルずつ、架空の売上明細を作ります。
乱数のシード（種）を固定しているので、何度実行しても同じ内容になります。
"""

import random
from pathlib import Path

import pandas as pd

STORES = {"shibuya": "渋谷店", "shinjuku": "新宿店", "ikebukuro": "池袋店"}
MONTHS = ["2026-07", "2026-08", "2026-09"]
PRODUCTS = {"コーヒー": 450, "カフェラテ": 520, "紅茶": 400, "サンドイッチ": 680, "ケーキ": 550}


def main():
    random.seed(42)
    out_dir = Path(__file__).parent / "sample_data"
    out_dir.mkdir(exist_ok=True)

    for month in MONTHS:
        for store_key, store_name in STORES.items():
            rows = []
            for day in range(1, 29, 3):  # 3日おきに売上を記録
                for product, price in PRODUCTS.items():
                    rows.append({
                        "日付": f"{month}-{day:02d}",
                        "店舗": store_name,
                        "商品": product,
                        "数量": random.randint(5, 40),
                        "単価": price,
                    })
            df = pd.DataFrame(rows)
            df.to_excel(out_dir / f"{month}_{store_key}.xlsx", index=False)

    print(f"サンプルデータを作成しました: {out_dir}")


if __name__ == "__main__":
    main()
