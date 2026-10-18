"""
○○ツール（← 案件ごとに名前を変える）

（このツールが何をするかを2〜3行で書く）
例: 入力の表を読み込み、「区分」ごとに「金額」を合計して、結果をExcelに書き出します。

使い方:
    python main.py                               # sample_data のファイルで動作確認
    python main.py 入力ファイル.xlsx 出力ファイル.xlsx
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

# このファイルがあるフォルダ（どこから実行してもパスがずれないようにする）
BASE_DIR = Path(__file__).parent

# 引数を省略したときに使う入力と出力
DEFAULT_INPUT = BASE_DIR / "sample_data" / "sample.csv"
DEFAULT_OUTPUT = BASE_DIR / "output" / "結果.xlsx"

# 入力の表に必ず必要な列（案件ごとに変える）
REQUIRED_COLUMNS = ["日付", "区分", "金額"]


class InputDataError(Exception):
    """入力データに問題があるときに出すエラー（メッセージは日本語）"""


# ---------------------------------------------------------------
# 1. 読み込み
# ---------------------------------------------------------------

def read_table(path):
    """Excel / CSV を読み込んで表（DataFrame）にする"""
    path = Path(path)
    if not path.exists():
        raise InputDataError(f"ファイルが見つかりません: {path}")

    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        df = pd.read_excel(path)
    elif suffix == ".csv":
        # Excelで保存したCSV（Shift_JIS）にも対応する
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(path, encoding="cp932")
    else:
        raise InputDataError(f"対応していない形式です（.xlsx / .csv のみ）: {path.name}")

    if df.empty:
        raise InputDataError(f"ファイルにデータが入っていません: {path.name}")

    # 列名の前後の空白を取り除く（「 金額」のような入力ミス対策）
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise InputDataError(
            f"必要な列がありません: {', '.join(missing)}（ファイル: {path.name}）"
        )
    return df


# ---------------------------------------------------------------
# 2. メインの処理（← 案件ごとにここを書きかえる）
# ---------------------------------------------------------------

def process(df):
    """表を受け取り、結果の表を返す（ここではサンプルとして区分別の合計）"""
    df = df.copy()
    # 数字に直せない金額はエラーにする（「1,200円」のような書き方もここで気づける）
    df["金額"] = pd.to_numeric(df["金額"], errors="coerce")
    bad_rows = df[df["金額"].isna()]
    if not bad_rows.empty:
        # 表の1行目は列名なので、データの行番号は +2 する
        rows = ", ".join(str(i + 2) for i in bad_rows.index)
        raise InputDataError(f"金額が数字になっていない行があります: {rows} 行目")

    result = df.groupby("区分", as_index=False)["金額"].sum()
    return result.sort_values("金額", ascending=False, ignore_index=True)


# ---------------------------------------------------------------
# 3. 書き出し
# ---------------------------------------------------------------

def write_excel(result, output_path):
    """結果をExcelに書き出す"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        result.to_excel(output_path, index=False, sheet_name="結果")
    except PermissionError:
        raise InputDataError(
            f"出力ファイルに書き込めません。Excelで開いたままになっていないか確認してください: {output_path}"
        )


def run(input_path, output_path):
    print(f"読み込み中: {input_path}")
    df = read_table(input_path)
    result = process(df)
    write_excel(result, output_path)
    print(f"\n処理が完了しました（{len(df)} 行を読み込み）")
    print(f"出力しました: {output_path}")


def parse_args(argv):
    parser = argparse.ArgumentParser(description="○○ツール")
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT, help="入力ファイル（.xlsx / .csv）")
    parser.add_argument("output", nargs="?", default=DEFAULT_OUTPUT, help="出力ファイル（.xlsx）")
    return parser.parse_args(argv)


def main():
    args = parse_args(sys.argv[1:])
    try:
        run(args.input, args.output)
    except InputDataError as e:
        print(f"\nエラー: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
