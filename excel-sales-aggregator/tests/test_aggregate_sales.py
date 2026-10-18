"""aggregate_sales.py のテスト（実行: python -m pytest）"""

import sys
from pathlib import Path

import pandas as pd
import pytest

# tests フォルダの1つ上（ツール本体のある場所）を読み込めるようにする
sys.path.insert(0, str(Path(__file__).parent.parent))
import aggregate_sales as agg  # noqa: E402


def make_excel(path, rows):
    pd.DataFrame(rows).to_excel(path, index=False)


def test_集計結果が正しい(tmp_path):
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    make_excel(input_dir / "a.xlsx", [
        {"日付": "2026-07-01", "店舗": "渋谷店", "商品": "コーヒー", "数量": 2, "単価": 450},
        {"日付": "2026-08-01", "店舗": "渋谷店", "商品": "紅茶", "数量": 1, "単価": 400},
    ])
    make_excel(input_dir / "b.xlsx", [
        {"日付": "2026-07-15", "店舗": "新宿店", "商品": "コーヒー", "数量": 3, "単価": 450},
    ])
    output = tmp_path / "out.xlsx"

    agg.run(input_dir, output)

    table = pd.read_excel(output, sheet_name="月別×店舗別", index_col=0)
    assert table.loc["2026-07", "渋谷店"] == 900
    assert table.loc["2026-07", "新宿店"] == 1350
    assert table.loc["2026-08", "渋谷店"] == 400
    assert table.loc["合計", "合計"] == 2650

    products = pd.read_excel(output, sheet_name="商品別")
    assert products.iloc[0]["商品"] == "コーヒー"  # 売上の多い順
    assert products.iloc[0]["数量"] == 5


def test_全角数字やカンマも数値として読める_読めない行は除外():
    raw = pd.DataFrame([
        {"日付": "2026-07-01", "店舗": " 渋谷店 ", "商品": "ケーキ", "数量": "２", "単価": "1,200"},
        {"日付": "2026-07-02", "店舗": "渋谷店", "商品": "ケーキ", "数量": "", "単価": 550},
        {"日付": "不明", "店舗": "渋谷店", "商品": "ケーキ", "数量": 1, "単価": 550},
    ])
    df, skipped = agg.clean_data(raw)

    assert skipped == 2
    assert len(df) == 1
    assert df.iloc[0]["売上金額"] == 2400
    assert df.iloc[0]["店舗"] == "渋谷店"


def test_必要な列がないとわかりやすいエラーになる(tmp_path):
    make_excel(tmp_path / "bad.xlsx", [{"日付": "2026-07-01", "店舗": "渋谷店"}])
    with pytest.raises(agg.InputDataError, match="商品"):
        agg.load_all_files(tmp_path)


def test_ファイルがないフォルダはエラーになる(tmp_path):
    with pytest.raises(agg.InputDataError, match="ファイルがありません"):
        agg.load_all_files(tmp_path)


def test_サンプルデータで最後まで動く(tmp_path):
    sample_dir = Path(__file__).parent.parent / "sample_data"
    df, skipped = agg.run(sample_dir, tmp_path / "out.xlsx")
    assert len(df) > 0
    assert skipped == 0
