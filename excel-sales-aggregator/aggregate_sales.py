"""
売上Excel集計ツール

フォルダに置いた複数の売上ファイル（Excel / CSV）を1つにまとめ、
「月別×店舗別」「商品別」の集計表を1つのExcelファイルに出力します。

使い方:
    python aggregate_sales.py                      # sample_data を集計して output/ に出力
    python aggregate_sales.py 入力フォルダ 出力ファイル.xlsx
"""

import sys
from pathlib import Path

import pandas as pd

# 入力ファイルに必ず必要な列
REQUIRED_COLUMNS = ["日付", "店舗", "商品", "数量", "単価"]

# 読み込む対象のファイルの拡張子
SUPPORTED_EXTENSIONS = [".xlsx", ".xls", ".csv"]


class InputDataError(Exception):
    """入力データに問題があるときに出すエラー（メッセージは日本語）"""


def read_one_file(path):
    """1つのファイルを読み込み、DataFrame（表データ）として返す"""
    if path.suffix.lower() == ".csv":
        # Excelで保存したCSVは文字コードがShift_JIS(cp932)のことが多いので、
        # まずUTF-8で読み、ダメならcp932で読み直す
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(path, encoding="cp932")
    else:
        df = pd.read_excel(path)

    # 列名の前後にある余計な空白を取り除く（「 店舗」→「店舗」）
    df.columns = [str(c).strip() for c in df.columns]

    # 必要な列がそろっているかチェック
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise InputDataError(
            f"「{path.name}」に必要な列がありません: {', '.join(missing)}\n"
            f"  必要な列: {', '.join(REQUIRED_COLUMNS)}"
        )

    # どのファイルから来たデータか分かるように列を追加
    df["元ファイル"] = path.name
    return df


def load_all_files(input_dir):
    """フォルダ内の対象ファイルをすべて読み込み、1つの表にまとめる"""
    input_dir = Path(input_dir)
    if not input_dir.is_dir():
        raise InputDataError(f"入力フォルダが見つかりません: {input_dir}")

    # 拡張子が対象のファイルだけを名前順に集める
    # （「~$」で始まるのはExcelを開いている間にできる一時ファイルなので除外）
    files = sorted(
        p for p in input_dir.iterdir()
        if p.suffix.lower() in SUPPORTED_EXTENSIONS and not p.name.startswith("~$")
    )
    if not files:
        raise InputDataError(
            f"「{input_dir}」に Excel / CSV ファイルがありません。"
        )

    tables = []
    for path in files:
        print(f"読み込み中: {path.name}")
        tables.append(read_one_file(path))

    return pd.concat(tables, ignore_index=True)


def clean_data(df):
    """データを集計できる形に整える。問題のある行は除外して、その件数を返す"""
    df = df.copy()

    # 文字列の前後の空白を取り除く
    for col in ["店舗", "商品"]:
        df[col] = df[col].astype(str).str.strip()

    # 数量・単価を数値に変換する。全角数字（「１２」）や「1,200」のカンマにも対応。
    # 数値にできないもの（空欄や文字）は NaN（値なし）になる
    for col in ["数量", "単価"]:
        text = df[col].astype(str).str.replace(",", "", regex=False)
        text = text.str.translate(str.maketrans("０１２３４５６７８９．", "0123456789."))
        df[col] = pd.to_numeric(text, errors="coerce")

    # 日付を日付型に変換する（変換できないものは NaT＝値なし）
    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")

    # 日付・数量・単価のどれかが空の行は集計できないので除外する
    bad_rows = df[["日付", "数量", "単価"]].isna().any(axis=1)
    skipped = int(bad_rows.sum())
    df = df[~bad_rows]

    # 売上金額と「年月」の列を作る
    df["売上金額"] = df["数量"] * df["単価"]
    df["年月"] = df["日付"].dt.strftime("%Y-%m")
    return df, skipped


def summarize_by_month_and_store(df):
    """月別×店舗別の売上表を作る（行：年月、列：店舗、右端と最下行に合計）"""
    return pd.pivot_table(
        df,
        index="年月",
        columns="店舗",
        values="売上金額",
        aggfunc="sum",
        fill_value=0,
        margins=True,        # 合計の行と列を追加する
        margins_name="合計",
    )


def summarize_by_product(df):
    """商品別の販売数量と売上金額を、売上の多い順に並べた表を作る"""
    result = (
        df.groupby("商品", as_index=False)[["数量", "売上金額"]]
        .sum()
        .sort_values("売上金額", ascending=False)
    )
    return result


def write_excel(output_path, details, by_month_store, by_product):
    """集計結果を1つのExcelファイル（3シート）に書き出す"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        by_month_store.to_excel(writer, sheet_name="月別×店舗別")
        by_product.to_excel(writer, sheet_name="商品別", index=False)
        details.to_excel(writer, sheet_name="明細", index=False)

        # 見やすいように列幅を広げ、金額にカンマ区切りを付ける
        for sheet in writer.sheets.values():
            for column in sheet.columns:
                sheet.column_dimensions[column[0].column_letter].width = 16
                for cell in column:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"


def run(input_dir, output_path):
    """読み込み → 整形 → 集計 → 書き出し をまとめて行う"""
    raw = load_all_files(input_dir)
    df, skipped = clean_data(raw)
    if df.empty:
        raise InputDataError("集計できるデータが1行もありませんでした。")

    details = df.copy()
    details["日付"] = details["日付"].dt.strftime("%Y-%m-%d")
    details = details[["日付", "店舗", "商品", "数量", "単価", "売上金額", "元ファイル"]]

    write_excel(
        output_path,
        details,
        summarize_by_month_and_store(df),
        summarize_by_product(df),
    )

    print(f"\n集計が完了しました: {output_path}")
    print(f"  集計した行数: {len(df)} 行")
    if skipped:
        print(f"  ※ 日付・数量・単価が読み取れなかった {skipped} 行は除外しました")
    return df, skipped


def main():
    # このファイルがあるフォルダを基準にする（どこから実行しても動くように）
    base_dir = Path(__file__).parent
    input_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else base_dir / "sample_data"
    output_path = Path(sys.argv[2]) if len(sys.argv) > 2 else base_dir / "output" / "売上集計.xlsx"

    try:
        run(input_dir, output_path)
    except InputDataError as e:
        print(f"\nエラー: {e}")
        sys.exit(1)
    except PermissionError:
        print(f"\nエラー: 「{output_path}」に書き込めません。Excelで開いている場合は閉じてから再実行してください。")
        sys.exit(1)


if __name__ == "__main__":
    main()
