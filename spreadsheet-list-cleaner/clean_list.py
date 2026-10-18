"""
申込リスト整理ツール（Excel / CSV / Googleスプレッドシート対応）

イベントやサービスの申込一覧の表を読み込み、
  1. 表記ゆれの整形（全角→半角、空白、メールの大文字、電話番号、日付の書き方）
  2. 重複した申込の削除（同じメールアドレス）
  3. 入力ミスの可能性がある行の洗い出し
  4. プラン別・月別の集計
を行い、結果をExcel（またはGoogleスプレッドシート）に書き出します。

使い方:
    python clean_list.py                              # sample_data を整理して output/ に出力
    python clean_list.py 入力ファイル.xlsx 出力ファイル.xlsx
    python clean_list.py --google --sheet-id スプレッドシートID   # Googleスプレッドシートで実行
"""

import argparse
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd

# 入力の表に必ず必要な列
REQUIRED_COLUMNS = ["申込日", "氏名", "メールアドレス", "電話番号", "プラン", "人数"]

# 重複を判定する列（同じメールアドレスなら同じ人の申込とみなす）
DUPLICATE_KEY = "メールアドレス"

# メールアドレスとしておかしくないかを確認する簡単なルール
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# 結果を書き出すシート名
SHEET_CLEANED = "整形済み"
SHEET_DUPLICATES = "重複"
SHEET_CHECK = "要確認"
SHEET_BY_PLAN = "プラン別集計"
SHEET_BY_MONTH = "月別集計"


class InputDataError(Exception):
    """入力データに問題があるときに出すエラー（メッセージは日本語）"""


# ---------------------------------------------------------------
# 1. 値の整形
# ---------------------------------------------------------------

def normalize_text(value):
    """全角英数字を半角にし、前後の空白を取り除く（空欄は空文字にする）"""
    if pd.isna(value):
        return ""
    # NFKC：「ＡＢＣ１２３」→「ABC123」、半角カナ「ﾔﾏﾀﾞ」→「ヤマダ」のようにそろえる
    text = unicodedata.normalize("NFKC", str(value))
    # 連続した空白（全角スペースも含む）を1つの半角スペースにまとめる
    return re.sub(r"\s+", " ", text).strip()


def normalize_email(value):
    """メールアドレスを半角・小文字にし、途中の空白を取り除く"""
    return normalize_text(value).replace(" ", "").lower()


def normalize_phone(value):
    """電話番号を「090-1234-5678」のようなハイフン区切りにそろえる"""
    text = normalize_text(value)
    # ハイフンに似た記号（ー、―、‐ など）をハイフンにする
    text = re.sub(r"[ー−―‐]", "-", text)
    digits = re.sub(r"\D", "", text)

    # 携帯電話（090 / 080 / 070 で始まる11桁）はハイフンの位置が決まっている
    if re.fullmatch(r"0[789]0\d{8}", digits):
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    # 固定電話は市外局番の長さが地域で違うので、元のハイフンの位置をそのまま使う
    if "-" in text:
        return "-".join(part for part in re.split(r"[^\d]+", text) if part)
    return digits


def normalize_date(value):
    """「2026/10/1」「2026年10月1日」などの日付を日付型にする（読めなければ None）"""
    if isinstance(value, pd.Timestamp):
        return value.normalize()
    text = normalize_text(value)
    if not text:
        return None
    # 「2026年10月1日」→「2026/10/1」
    text = text.replace("年", "/").replace("月", "/").replace("日", "")
    try:
        return pd.to_datetime(text).normalize()
    except (ValueError, TypeError):
        return None


def normalize_number(value):
    """「２名」「3人」のような人数を整数にする（読めなければ None）"""
    digits = re.sub(r"\D", "", normalize_text(value))
    return int(digits) if digits else None


def clean_table(df):
    """
    表全体を整形する。
    戻り値: 整形後の表（「元の行」「要確認の理由」の列つき）
    """
    df = df.copy()
    df.columns = [normalize_text(c) for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise InputDataError(
            f"必要な列がありません: {', '.join(missing)}\n"
            f"  必要な列: {', '.join(REQUIRED_COLUMNS)}"
        )

    # 元の表で何行目だったかを記録する（見出しが1行目なので、データは2行目から）
    df.insert(0, "元の行", range(2, len(df) + 2))

    # すべての列が空欄の行は取り除く
    df = df[~df[REQUIRED_COLUMNS].map(lambda v: normalize_text(v) == "").all(axis=1)]

    df["申込日"] = df["申込日"].map(normalize_date)
    df["氏名"] = df["氏名"].map(normalize_text)
    df["メールアドレス"] = df["メールアドレス"].map(normalize_email)
    df["電話番号"] = df["電話番号"].map(normalize_phone)
    df["プラン"] = df["プラン"].map(normalize_text)
    df["人数"] = df["人数"].map(normalize_number).astype("Int64")

    df["要確認の理由"] = df.apply(find_problems, axis=1)
    return df.reset_index(drop=True)


def find_problems(row):
    """1行分のデータを見て、入力ミスの可能性がある点を「、」区切りで返す"""
    problems = []
    if row["申込日"] is None or pd.isna(row["申込日"]):
        problems.append("申込日が読み取れない")
    if not row["氏名"]:
        problems.append("氏名が空欄")
    if not EMAIL_PATTERN.match(row["メールアドレス"]):
        problems.append("メールアドレスの形式が正しくない")
    digit_count = len(re.sub(r"\D", "", row["電話番号"]))
    if digit_count not in (10, 11):
        problems.append("電話番号の桁数が正しくない")
    if not row["プラン"]:
        problems.append("プランが空欄")
    if pd.isna(row["人数"]) or row["人数"] <= 0:
        problems.append("人数が読み取れない")
    return "、".join(problems)


# ---------------------------------------------------------------
# 2. 重複の削除
# ---------------------------------------------------------------

def split_duplicates(df):
    """
    同じメールアドレスの申込が2回以上あれば、最初の1件だけを残す。
    メールアドレスが空欄の行は、重複の判定をしない。
    戻り値: (残した表, 取り除いた重複の表)
    """
    has_key = df[DUPLICATE_KEY] != ""
    is_duplicate = has_key & df.duplicated(subset=DUPLICATE_KEY, keep="first")
    return df[~is_duplicate].reset_index(drop=True), df[is_duplicate].reset_index(drop=True)


# ---------------------------------------------------------------
# 3. 集計
# ---------------------------------------------------------------

def summarize(df):
    """プラン別・月別に、申込件数と人数の合計を出す（要確認の行は含めない）"""
    ok = df[df["要確認の理由"] == ""].copy()

    by_plan = (
        ok.groupby("プラン")
        .agg(申込件数=("氏名", "size"), 人数合計=("人数", "sum"))
        .sort_values("人数合計", ascending=False)
        .reset_index()
    )

    ok["申込月"] = pd.to_datetime(ok["申込日"]).dt.strftime("%Y-%m")
    by_month = (
        ok.groupby("申込月")
        .agg(申込件数=("氏名", "size"), 人数合計=("人数", "sum"))
        .reset_index()
    )

    # 一番下に合計の行をつける
    for table, label_column in ((by_plan, "プラン"), (by_month, "申込月")):
        table.loc[len(table)] = {
            label_column: "合計",
            "申込件数": table["申込件数"].sum(),
            "人数合計": table["人数合計"].sum(),
        }
    return by_plan, by_month


def process(df):
    """整形 → 重複削除 → 集計 をまとめて行い、シート名と表の辞書を返す"""
    cleaned = clean_table(df)
    kept, duplicates = split_duplicates(cleaned)
    by_plan, by_month = summarize(kept)
    check = kept[kept["要確認の理由"] != ""]

    return {
        SHEET_CLEANED: kept,
        SHEET_DUPLICATES: duplicates,
        SHEET_CHECK: check,
        SHEET_BY_PLAN: by_plan,
        SHEET_BY_MONTH: by_month,
    }


def for_output(table):
    """書き出し用に、日付を「2026-10-01」の文字列に、空の値を空欄にする"""
    table = table.copy()
    if "申込日" in table.columns:
        table["申込日"] = table["申込日"].map(
            lambda d: "" if d is None or pd.isna(d) else pd.Timestamp(d).strftime("%Y-%m-%d")
        )
    return table.astype(object).where(table.notna(), "")


# ---------------------------------------------------------------
# 4-A. ローカルのExcel / CSV で動かす
# ---------------------------------------------------------------

def read_local(path):
    """Excel または CSV ファイルを読み込む"""
    path = Path(path)
    if not path.is_file():
        raise InputDataError(f"入力ファイルが見つかりません: {path}")

    if path.suffix.lower() == ".csv":
        # Excelで保存したCSVはShift_JIS(cp932)のことが多いので、UTF-8がダメなら読み直す
        try:
            return pd.read_csv(path, encoding="utf-8-sig", dtype=str)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="cp932", dtype=str)
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    raise InputDataError(f"Excel（.xlsx）か CSV（.csv）のファイルを指定してください: {path.name}")


def write_local(sheets, output_path):
    """結果を1つのExcelファイル（シートごと）に書き出す"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for name, table in sheets.items():
                for_output(table).to_excel(writer, sheet_name=name, index=False)
                sheet = writer.sheets[name]
                # 列幅を中身の長さに合わせて広げる（日本語は2文字分で数える）
                for column_cells in sheet.columns:
                    longest = max(text_width(cell.value) for cell in column_cells)
                    sheet.column_dimensions[column_cells[0].column_letter].width = min(longest + 2, 60)
                sheet.freeze_panes = "A2"  # 見出し行を固定する
    except PermissionError:
        raise InputDataError(
            f"出力ファイルに書き込めません。Excelで開いている場合は閉じてから実行してください: {output_path}"
        )
    return output_path


def text_width(value):
    """Excelの列幅用に、文字列の見た目の幅を数える（全角は2）"""
    text = "" if value is None else str(value)
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in text)


def run_local(input_path, output_path):
    print(f"読み込み中: {input_path}")
    sheets = process(read_local(input_path))
    write_local(sheets, output_path)
    print_result(sheets)
    print(f"\n出力しました: {output_path}")
    return sheets


# ---------------------------------------------------------------
# 4-B. Googleスプレッドシートで動かす（任意）
# ---------------------------------------------------------------

def run_google(sheet_id, worksheet_name, credentials_path, client=None):
    """
    Googleスプレッドシートの「worksheet_name」シートを読み込み、
    結果を同じスプレッドシートの別シート（整形済み・重複 など）に書き込む。
    元のシートは書き換えない。
    client はテスト用に差し替えられるようにしている。
    """
    import google_sheets  # gspread が必要なので、Googleモードのときだけ読み込む

    try:
        if client is None:
            client = google_sheets.connect(credentials_path)
        spreadsheet = google_sheets.open_spreadsheet(client, sheet_id)

        print(f"読み込み中: シート「{worksheet_name}」")
        df = google_sheets.read_worksheet(spreadsheet, worksheet_name)
        sheets = process(df)
        for name, table in sheets.items():
            print(f"書き込み中: シート「{name}」")
            google_sheets.write_worksheet(spreadsheet, name, for_output(table))
    except google_sheets.GoogleSheetsError as e:
        raise InputDataError(str(e))

    print_result(sheets)
    print("\nGoogleスプレッドシートに書き込みました")
    return sheets


def print_result(sheets):
    print("\n整理が完了しました")
    print(f"  残した申込: {len(sheets[SHEET_CLEANED])} 件")
    print(f"  取り除いた重複: {len(sheets[SHEET_DUPLICATES])} 件")
    print(f"  要確認: {len(sheets[SHEET_CHECK])} 件")


def parse_args(argv):
    base_dir = Path(__file__).parent
    parser = argparse.ArgumentParser(description="申込リストを整形・重複削除・集計します")
    parser.add_argument(
        "input", nargs="?", default=str(base_dir / "sample_data" / "申込一覧.csv"),
        help="入力ファイル（.xlsx / .csv）。初期値: sample_data/申込一覧.csv",
    )
    parser.add_argument(
        "output", nargs="?", default=str(base_dir / "output" / "申込一覧_整理済み.xlsx"),
        help="出力ファイル（.xlsx）。初期値: output/申込一覧_整理済み.xlsx",
    )
    parser.add_argument("--google", action="store_true", help="Googleスプレッドシートで実行する")
    parser.add_argument("--sheet-id", help="GoogleスプレッドシートのID（URLの /d/ と /edit の間の文字列）")
    parser.add_argument("--worksheet", default="申込一覧", help="読み込むシート名。初期値: 申込一覧")
    parser.add_argument(
        "--credentials", default=str(base_dir / "credentials.json"),
        help="サービスアカウントの鍵ファイル。初期値: credentials.json（環境変数 GOOGLE_CREDENTIALS でも指定可）",
    )
    args = parser.parse_args(argv)
    if args.google and not args.sheet_id:
        parser.error("--google を使うときは --sheet-id でスプレッドシートのIDを指定してください")
    return args


def main():
    args = parse_args(sys.argv[1:])
    try:
        if args.google:
            run_google(args.sheet_id, args.worksheet, args.credentials)
        else:
            run_local(args.input, args.output)
    except InputDataError as e:
        print(f"\nエラー: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
