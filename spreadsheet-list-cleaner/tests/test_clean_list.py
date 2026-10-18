"""clean_list.py のテスト（実行: python -m pytest）

Googleスプレッドシートのテストは、本物のGoogleには接続せず、
同じ動きをする「にせもののスプレッドシート」を使って確認します。
"""

import sys
from pathlib import Path

import gspread
import pandas as pd
import pytest

# tests フォルダの1つ上（ツール本体のある場所）を読み込めるようにする
sys.path.insert(0, str(Path(__file__).parent.parent))
import clean_list as cl  # noqa: E402

SAMPLE = Path(__file__).parent.parent / "sample_data" / "申込一覧.csv"


def make_table(rows):
    """テスト用の小さな申込表を作る"""
    return pd.DataFrame(rows, columns=cl.REQUIRED_COLUMNS)


def test_全角や空白の表記ゆれをそろえる():
    assert cl.normalize_text("　佐藤　　花子 ") == "佐藤 花子"
    assert cl.normalize_text("ｱﾍﾞ") == "アベ"
    assert cl.normalize_email(" Taro.Yamada@EXAMPLE.com ") == "taro.yamada@example.com"
    assert cl.normalize_number("２名") == 2
    assert cl.normalize_number("") is None


def test_電話番号をハイフン区切りにそろえる():
    assert cl.normalize_phone("09011110002") == "090-1111-0002"
    assert cl.normalize_phone("０９０－１１１１－０００３") == "090-1111-0003"
    assert cl.normalize_phone("080 1111 0004") == "080-1111-0004"
    # 固定電話は元のハイフンの位置を残す
    assert cl.normalize_phone("045-111-0024") == "045-111-0024"


def test_いろいろな書き方の日付を読み取る():
    expected = pd.Timestamp("2026-08-03")
    assert cl.normalize_date("2026/8/3") == expected
    assert cl.normalize_date("2026-08-03") == expected
    assert cl.normalize_date("2026年8月3日") == expected
    assert cl.normalize_date("あとで") is None
    assert cl.normalize_date("") is None


def test_同じメールアドレスの申込は最初の1件だけ残す():
    df = make_table([
        ["2026/08/01", "山田 太郎", "taro@example.com", "090-1111-0001", "ライト", "1"],
        ["2026/08/05", "山田 太郎", "TARO@example.com", "090-1111-0001", "ライト", "1"],
        ["2026/08/06", "佐藤 花子", "", "090-1111-0002", "ライト", "1"],
        ["2026/08/07", "鈴木 一郎", "", "090-1111-0003", "ライト", "1"],
    ])
    kept, duplicates = cl.split_duplicates(cl.clean_table(df))

    assert list(kept["元の行"]) == [2, 4, 5]  # 空欄のメールアドレスは重複扱いしない
    assert list(duplicates["元の行"]) == [3]


def test_入力ミスの可能性がある行に理由をつける():
    df = make_table([
        ["2026/08/01", "山田 太郎", "taro@example.com", "090-1111-0001", "ライト", "1"],
        ["", "", "taro@example", "090-1111", "", "なし"],
    ])
    cleaned = cl.clean_table(df)

    assert cleaned.loc[0, "要確認の理由"] == ""
    assert cleaned.loc[1, "要確認の理由"] == (
        "申込日が読み取れない、氏名が空欄、メールアドレスの形式が正しくない、"
        "電話番号の桁数が正しくない、プランが空欄、人数が読み取れない"
    )


def test_必要な列がなければ日本語のエラーを出す():
    with pytest.raises(cl.InputDataError, match="必要な列がありません: 人数"):
        cl.clean_table(pd.DataFrame(columns=cl.REQUIRED_COLUMNS[:-1]))


def test_サンプルデータを整理してExcelに出力できる(tmp_path):
    output = tmp_path / "結果.xlsx"
    cl.run_local(SAMPLE, output)

    sheets = pd.read_excel(output, sheet_name=None)
    assert list(sheets) == ["整形済み", "重複", "要確認", "プラン別集計", "月別集計"]
    assert len(sheets["整形済み"]) == 27
    assert len(sheets["重複"]) == 3
    assert len(sheets["要確認"]) == 4

    by_plan = sheets["プラン別集計"].set_index("プラン")
    assert by_plan.loc["合計", "申込件数"] == 23
    assert by_plan.loc["合計", "人数合計"] == 46
    assert sheets["月別集計"]["申込月"].tolist() == ["2026-08", "2026-09", "合計"]


def test_入力ファイルがなければ日本語のエラーを出す(tmp_path):
    with pytest.raises(cl.InputDataError, match="入力ファイルが見つかりません"):
        cl.run_local(tmp_path / "ない.csv", tmp_path / "結果.xlsx")


# ---------------------------------------------------------------
# Googleスプレッドシート モードのテスト
# ---------------------------------------------------------------

class FakeWorksheet:
    def __init__(self, values):
        self.values = values

    def get_all_values(self):
        return self.values

    def clear(self):
        self.values = []

    def update(self, values, range_name):
        assert range_name == "A1"
        self.values = values

    def freeze(self, rows):
        pass


class FakeSpreadsheet:
    def __init__(self, sheets):
        self.sheets = sheets

    def worksheet(self, name):
        if name not in self.sheets:
            raise gspread.exceptions.WorksheetNotFound(name)
        return self.sheets[name]

    def add_worksheet(self, title, rows, cols):
        self.sheets[title] = FakeWorksheet([])
        return self.sheets[title]


class FakeClient:
    def __init__(self, spreadsheet):
        self.spreadsheet = spreadsheet

    def open_by_key(self, key):
        if key != "test-sheet-id":
            raise gspread.exceptions.SpreadsheetNotFound(key)
        return self.spreadsheet


def sample_values():
    """サンプルCSVを、Googleスプレッドシートから読んだときと同じ「文字列の2次元リスト」にする"""
    df = pd.read_csv(SAMPLE, encoding="utf-8-sig", dtype=str).fillna("")
    return [list(df.columns)] + df.values.tolist()


def test_Googleスプレッドシートを読み込み_結果を別シートに書き込む():
    original = sample_values()
    spreadsheet = FakeSpreadsheet({"申込一覧": FakeWorksheet(original)})

    cl.run_google("test-sheet-id", "申込一覧", "credentials.json", client=FakeClient(spreadsheet))

    # 元のシートは書き換えない
    assert spreadsheet.sheets["申込一覧"].values == original
    cleaned = spreadsheet.sheets["整形済み"].values
    assert cleaned[0] == ["元の行", *cl.REQUIRED_COLUMNS, "要確認の理由"]
    assert cleaned[1] == [2, "2026-08-01", "山田 太郎", "taro.yamada@example.com", "090-1111-0001", "スタンダード", 2, ""]
    assert len(cleaned) == 1 + 27
    # Googleに送る値は、普通のPythonの数値・文字列になっている
    assert all(type(v) in (int, str) for row in cleaned for v in row)
    assert spreadsheet.sheets["プラン別集計"].values[-1] == ["合計", 23, 46]


def test_Googleのシート名が違えば日本語のエラーを出す():
    spreadsheet = FakeSpreadsheet({"シート1": FakeWorksheet([["a"]])})
    with pytest.raises(cl.InputDataError, match="シート「申込一覧」が見つかりません"):
        cl.run_google("test-sheet-id", "申込一覧", "credentials.json", client=FakeClient(spreadsheet))


def test_鍵ファイルがなければ日本語のエラーを出す(tmp_path, monkeypatch):
    monkeypatch.delenv("GOOGLE_CREDENTIALS", raising=False)
    with pytest.raises(cl.InputDataError, match="鍵ファイルが見つかりません"):
        cl.run_google("test-sheet-id", "申込一覧", tmp_path / "credentials.json")
