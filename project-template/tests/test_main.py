"""main.py のテスト（実行: python -m pytest）

案件ごとに、ご依頼の内容に合わせてテストを足していきます。
"""

import sys
from pathlib import Path

import pandas as pd
import pytest

# tests フォルダの1つ上（ツール本体のある場所）を読み込めるようにする
sys.path.insert(0, str(Path(__file__).parent.parent))
import main  # noqa: E402


def test_サンプルデータを処理してExcelに出力できる(tmp_path):
    output = tmp_path / "結果.xlsx"
    main.run(main.DEFAULT_INPUT, output)

    result = pd.read_excel(output, sheet_name="結果")
    totals = dict(zip(result["区分"], result["金額"]))
    assert totals == {"食費": 2180, "交通費": 840, "日用品": 1500}


def test_Shift_JISのCSVも読める(tmp_path):
    path = tmp_path / "sjis.csv"
    path.write_text("日付,区分,金額\n2026-09-01,食費,100\n", encoding="cp932")
    df = main.read_table(path)
    assert df.loc[0, "区分"] == "食費"


def test_ファイルがないときは日本語のエラー():
    with pytest.raises(main.InputDataError, match="ファイルが見つかりません"):
        main.read_table("ない.csv")


def test_必要な列がないときは日本語のエラー(tmp_path):
    path = tmp_path / "bad.csv"
    path.write_text("日付,金額\n2026-09-01,100\n", encoding="utf-8")
    with pytest.raises(main.InputDataError, match="必要な列がありません: 区分"):
        main.read_table(path)


def test_空のファイルは日本語のエラー(tmp_path):
    path = tmp_path / "empty.csv"
    path.write_text("日付,区分,金額\n", encoding="utf-8")
    with pytest.raises(main.InputDataError, match="データが入っていません"):
        main.read_table(path)


def test_金額が数字でない行を知らせる():
    df = pd.DataFrame({"日付": ["2026-09-01", "2026-09-02"], "区分": ["食費", "食費"], "金額": [100, "1,200円"]})
    with pytest.raises(main.InputDataError, match="3 行目"):
        main.process(df)
