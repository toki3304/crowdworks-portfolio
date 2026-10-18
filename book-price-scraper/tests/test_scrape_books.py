"""scrape_books.py のテスト（実行: python -m pytest）

本物のサイトにはアクセスせず、tests/html にある保存済みのHTMLを使ってテストします。
"""

import sys
from pathlib import Path

import pandas as pd
import pytest

# tests フォルダの1つ上（ツール本体のある場所）を読み込めるようにする
sys.path.insert(0, str(Path(__file__).parent.parent))
import scrape_books as sb  # noqa: E402

HTML_DIR = Path(__file__).parent / "html"
BASE = "https://books.toscrape.com/catalogue/"


def fake_fetch(url):
    """URLの代わりに、保存済みのHTMLファイルを返す"""
    name = url.rsplit("/", 1)[-1]
    return (HTML_DIR / name).read_text(encoding="utf-8")


@pytest.fixture(autouse=True)
def no_sleep(monkeypatch):
    """テストでは待ち時間をなくし、何秒待とうとしたかを記録する"""
    waits = []
    monkeypatch.setattr(sb.time, "sleep", lambda seconds: waits.append(seconds))
    return waits


def test_一覧ページから本の情報を取り出せる():
    books = sb.parse_books(fake_fetch(BASE + "page-1.html"), BASE + "page-1.html")

    assert len(books) == 3
    assert books[0] == {
        "タイトル": "A Light in the Attic",
        "価格(ポンド)": 51.77,
        "評価(星)": 3,
        "在庫": "あり",
        "URL": "https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html",
    }
    # 画面上で「...」と省略されたタイトルも、完全な名前で取れる
    assert books[1]["タイトル"] == "Sapiens: A Brief History of Humankind"
    assert books[2]["在庫"] == "なし"


def test_次のページのURLを見つけられる_最後のページではNone():
    assert sb.find_next_page(fake_fetch(BASE + "page-1.html"), BASE + "page-1.html") == BASE + "page-2.html"
    assert sb.find_next_page(fake_fetch(BASE + "page-2.html"), BASE + "page-2.html") is None


def test_全ページをたどり_ページごとに待つ(no_sleep):
    books = sb.scrape(BASE + "page-1.html", max_pages=0, delay=1.5, fetch=fake_fetch)

    assert len(books) == 5
    assert no_sleep == [1.5]  # 2ページ目の前に1回だけ待つ


def test_ページ数を指定するとそこで止まる():
    books = sb.scrape(BASE + "page-1.html", max_pages=1, delay=1, fetch=fake_fetch)
    assert len(books) == 3


def test_CSVとExcelが出力される(tmp_path):
    output = tmp_path / "out" / "書籍一覧.xlsx"
    sb.run(BASE + "page-1.html", 0, 1, output, fetch=fake_fetch, check_robots=False)

    excel = pd.read_excel(output, sheet_name="書籍一覧")
    csv = pd.read_csv(output.with_suffix(".csv"), encoding="utf-8-sig")
    assert list(excel.columns) == sb.COLUMNS
    assert len(excel) == len(csv) == 5
    assert excel["価格(ポンド)"].sum() == pytest.approx(182.93)


def test_本が1冊も取れないとわかりやすいエラーになる(tmp_path):
    with pytest.raises(sb.ScrapeError, match="1件も取得できませんでした"):
        sb.run(BASE + "page-1.html", 0, 1, tmp_path / "out.xlsx",
               fetch=lambda url: "<html></html>", check_robots=False)


def test_待ち時間が1秒未満だと実行できない():
    with pytest.raises(SystemExit):
        sb.parse_args(["--delay", "0.1"])
