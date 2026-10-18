"""
Googleスプレッドシートとの読み書き（clean_list.py の --google モードで使う）

gspread というライブラリを使います。
鍵ファイル（credentials.json）の作り方は README.md の「Googleスプレッドシートで使う」を見てください。
鍵ファイルはパスワードと同じなので、GitHubなどに絶対に公開しないでください（.gitignore で除外しています）。
"""

import os
from pathlib import Path

import gspread


class GoogleSheetsError(Exception):
    """Googleスプレッドシートの操作で問題があったときに出すエラー（メッセージは日本語）"""


def connect(credentials_path):
    """鍵ファイルを使ってGoogleにログインし、操作用のクライアントを返す"""
    # 環境変数 GOOGLE_CREDENTIALS があれば、そちらの場所を優先する
    path = Path(os.environ.get("GOOGLE_CREDENTIALS") or credentials_path)
    if not path.is_file():
        raise GoogleSheetsError(
            f"鍵ファイルが見つかりません: {path}\n"
            "  README.md の「Googleスプレッドシートで使う」の手順で credentials.json を用意してください"
        )
    try:
        return gspread.service_account(filename=str(path))
    except ValueError:
        raise GoogleSheetsError(f"鍵ファイルの中身が正しくありません。ダウンロードし直してください: {path}")


def open_spreadsheet(client, sheet_id):
    """IDを指定してスプレッドシートを開く"""
    try:
        return client.open_by_key(sheet_id)
    except gspread.exceptions.SpreadsheetNotFound:
        raise GoogleSheetsError(
            f"スプレッドシートが見つかりません: {sheet_id}\n"
            "  IDが正しいか、スプレッドシートをサービスアカウントのメールアドレスに「編集者」で共有しているか確認してください"
        )
    except gspread.exceptions.APIError as e:
        raise GoogleSheetsError(f"Googleスプレッドシートにアクセスできませんでした: {e}")


def read_worksheet(spreadsheet, name):
    """シートの中身を読み込み、1行目を見出しにした表（DataFrame）で返す"""
    import pandas as pd

    try:
        worksheet = spreadsheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        raise GoogleSheetsError(f"シート「{name}」が見つかりません。--worksheet でシート名を指定してください")

    values = worksheet.get_all_values()
    if not values:
        raise GoogleSheetsError(f"シート「{name}」が空です")
    return pd.DataFrame(values[1:], columns=values[0])


def write_worksheet(spreadsheet, name, table):
    """表をシートに書き込む。同じ名前のシートがあれば中身を消してから書き込む"""
    rows = [list(table.columns)] + [[to_cell(v) for v in row] for row in table.itertuples(index=False)]

    try:
        worksheet = spreadsheet.worksheet(name)
        worksheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=name, rows=len(rows) + 10, cols=len(table.columns) + 2)

    worksheet.update(rows, "A1")
    worksheet.freeze(rows=1)  # 見出し行を固定する


def to_cell(value):
    """pandasの数値（numpy型）を、Googleに送れる普通のPythonの値にする"""
    return value.item() if hasattr(value, "item") else value
