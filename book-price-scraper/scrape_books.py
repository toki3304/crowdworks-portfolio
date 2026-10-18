"""
書籍情報スクレイピングツール

スクレイピング練習用サイト「Books to Scrape」(https://books.toscrape.com/) から
本のタイトル・価格・評価・在庫状況を集め、CSVとExcelに出力します。

使い方:
    python scrape_books.py                 # 先頭から3ページ分（60冊）を取得
    python scrape_books.py --pages 10      # 10ページ分を取得
    python scrape_books.py --pages 0       # 全ページ（50ページ・1000冊）を取得
    python scrape_books.py --output 出力ファイル名.xlsx --delay 2
"""

import argparse
import sys
import time
from pathlib import Path
from urllib.parse import urljoin
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup

# 取得を始めるページ
START_URL = "https://books.toscrape.com/catalogue/page-1.html"

# サーバーに「誰がアクセスしているか」を伝える名前
USER_AGENT = "book-price-scraper/1.0 (portfolio sample)"

# 評価は「Three」のように英単語で書かれているので数字に直す
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

# 出力する列の順番
COLUMNS = ["タイトル", "価格(ポンド)", "評価(星)", "在庫", "URL"]


class ScrapeError(Exception):
    """取得中に問題があったときに出すエラー（メッセージは日本語）"""


def fetch_html(url):
    """URLのページを取得し、HTMLの文字列を返す"""
    try:
        response = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=10)
        response.raise_for_status()  # 404などのエラーなら例外にする
    except requests.exceptions.ConnectionError:
        raise ScrapeError(f"サイトに接続できません。インターネット接続を確認してください: {url}")
    except requests.exceptions.Timeout:
        raise ScrapeError(f"サイトからの応答がありません（タイムアウト）: {url}")
    except requests.exceptions.HTTPError as e:
        raise ScrapeError(f"ページを取得できませんでした（{e.response.status_code}）: {url}")

    # 文字化けを防ぐため、文字コードはUTF-8として扱う
    response.encoding = "utf-8"
    return response.text


def is_allowed_by_robots(url):
    """サイトの robots.txt で、このURLの取得が許可されているか確認する"""
    robots_url = urljoin(url, "/robots.txt")
    parser = RobotFileParser()
    try:
        response = requests.get(robots_url, headers={"User-Agent": USER_AGENT}, timeout=10)
    except requests.exceptions.RequestException:
        # robots.txt が取れないときは、本体の取得時にエラーを表示させる
        return True
    if response.status_code != 200:
        # robots.txt がないサイトは、取得の制限がないものとして扱う
        return True
    parser.parse(response.text.splitlines())
    return parser.can_fetch(USER_AGENT, url)


def parse_books(html, page_url):
    """一覧ページのHTMLから本の情報を取り出し、辞書のリストで返す"""
    soup = BeautifulSoup(html, "html.parser")
    books = []

    # 1冊ごとに <article class="product_pod"> で囲まれている
    for article in soup.select("article.product_pod"):
        link = article.select_one("h3 a")
        price_text = article.select_one("p.price_color").get_text(strip=True)
        rating_tag = article.select_one("p.star-rating")
        stock_text = article.select_one("p.availability").get_text(strip=True)

        # 評価のクラス名は「star-rating Three」のようになっている
        rating_word = next((c for c in rating_tag["class"] if c in RATING_WORDS), None)

        books.append({
            # 画面上は長いタイトルが「...」で省略されるので、title属性から完全な名前を取る
            "タイトル": link.get("title") or link.get_text(strip=True),
            "価格(ポンド)": to_number(price_text),
            "評価(星)": RATING_WORDS.get(rating_word),
            "在庫": "あり" if "In stock" in stock_text else "なし",
            # リンクは相対パスなので、ページのURLを基準に完全なURLにする
            "URL": urljoin(page_url, link["href"]),
        })
    return books


def to_number(price_text):
    """「£51.77」のような文字列を 51.77 という数値にする"""
    digits = "".join(ch for ch in price_text if ch.isdigit() or ch == ".")
    return float(digits) if digits else None


def find_next_page(html, page_url):
    """「next」ボタンのリンク先URLを返す。最後のページなら None"""
    soup = BeautifulSoup(html, "html.parser")
    next_link = soup.select_one("li.next a")
    if next_link is None:
        return None
    return urljoin(page_url, next_link["href"])


def scrape(start_url, max_pages, delay, fetch=fetch_html):
    """
    start_url から順にページをたどって本の情報を集める。
    max_pages が 0 のときは最後のページまで取得する。
    サーバーに負担をかけないよう、ページごとに delay 秒待つ。
    """
    books = []
    url = start_url
    page = 0

    while url is not None:
        page += 1
        if page > 1:
            time.sleep(delay)  # 連続アクセスを避けるために待つ

        print(f"取得中 ({page}ページ目): {url}")
        html = fetch(url)
        books.extend(parse_books(html, url))

        if max_pages and page >= max_pages:
            break
        url = find_next_page(html, url)

    return books


def write_outputs(books, output_path):
    """集めた情報をExcelとCSV（同じ名前で拡張子違い）に書き出す"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(books, columns=COLUMNS)

    # CSV：Excelで開いても文字化けしないように「BOM付きUTF-8」で保存する
    csv_path = output_path.with_suffix(".csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    # Excel：見やすいように列幅を調整する
    xlsx_path = output_path.with_suffix(".xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="書籍一覧", index=False)
        sheet = writer.sheets["書籍一覧"]
        widths = {"A": 60, "B": 14, "C": 10, "D": 8, "E": 80}
        for column, width in widths.items():
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A2"  # 見出し行を固定する

    return csv_path, xlsx_path


def run(start_url, max_pages, delay, output_path, fetch=fetch_html, check_robots=True):
    """取得 → 書き出し をまとめて行う"""
    if check_robots and not is_allowed_by_robots(start_url):
        raise ScrapeError(f"このサイトの robots.txt で取得が禁止されています: {start_url}")

    books = scrape(start_url, max_pages, delay, fetch=fetch)
    if not books:
        raise ScrapeError("本の情報が1件も取得できませんでした。サイトの構成が変わった可能性があります。")

    csv_path, xlsx_path = write_outputs(books, output_path)
    print(f"\n取得が完了しました: {len(books)} 冊")
    print(f"  Excel: {xlsx_path}")
    print(f"  CSV  : {csv_path}")
    return books


def parse_args(argv):
    base_dir = Path(__file__).parent
    parser = argparse.ArgumentParser(description="Books to Scrape から書籍情報を取得してCSV/Excelに出力します")
    parser.add_argument("--pages", type=int, default=3, help="取得するページ数（0で全ページ）。初期値: 3")
    parser.add_argument("--delay", type=float, default=1.0, help="ページごとに待つ秒数。初期値: 1")
    parser.add_argument(
        "--output",
        default=str(base_dir / "output" / "書籍一覧.xlsx"),
        help="出力ファイル名（.xlsx と .csv の両方ができます）",
    )
    args = parser.parse_args(argv)
    if args.pages < 0:
        parser.error("--pages には0以上の数を指定してください")
    if args.delay < 1:
        parser.error("サーバーに負担をかけないよう、--delay は1秒以上にしてください")
    return args


def main():
    args = parse_args(sys.argv[1:])
    try:
        run(START_URL, args.pages, args.delay, args.output)
    except ScrapeError as e:
        print(f"\nエラー: {e}")
        sys.exit(1)
    except PermissionError:
        print(f"\nエラー: 「{args.output}」に書き込めません。Excelで開いている場合は閉じてから再実行してください。")
        sys.exit(1)


if __name__ == "__main__":
    main()
