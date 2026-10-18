# 書籍情報スクレイピングツール

Webサイトの一覧ページを自動で1ページずつたどり、
商品名・価格・評価・在庫状況を集めて **CSV と Excel** に出力するPythonツールです。

「競合サイトの商品価格を毎週コピーして表にまとめている」といった作業を自動化する例として、
スクレイピング練習用に公開されているサイト [Books to Scrape](https://books.toscrape.com/) を対象にしています。

## できること

- 一覧ページから次の情報を取得
  - タイトル（画面上で「...」と省略されていても完全な名前を取得）
  - 価格（「£51.77」→ `51.77` の数値に変換）
  - 評価（星の数を 1〜5 の数値に変換）
  - 在庫の有無
  - 商品ページのURL
- 「next」ボタンをたどって複数ページを自動取得（ページ数は指定可能、全50ページ・1000冊まで対応）
- CSV（Excelで開いても文字化けしない形式）と Excel の両方を出力
- サイトに負担をかけないための配慮
  - ページごとに1秒以上の待ち時間を入れる（1秒未満は指定できないようにしています）
  - 取得前に `robots.txt` を確認し、禁止されていれば取得しない
  - アクセス元が分かるように User-Agent を設定
- エラーは日本語で表示（インターネットにつながらない、ページが見つからない、出力ファイルをExcelで開いたまま、など）

## 実行方法

### 1. 準備（最初の1回だけ）

1. [Python公式サイト](https://www.python.org/downloads/) からPython 3.10以上をインストールします
   （Windowsではインストール画面の **「Add python.exe to PATH」にチェック** を入れてください）
2. このフォルダでコマンドプロンプト（Macはターミナル）を開き、必要なライブラリを入れます

```
pip install -r requirements.txt
```

### 2. 実行

```
python scrape_books.py
```

先頭から3ページ分（60冊）を取得し、`output/書籍一覧.xlsx` と `output/書籍一覧.csv` ができます。

オプションで取得するページ数や待ち時間を変えられます。

```
python scrape_books.py --pages 10          # 10ページ分（200冊）
python scrape_books.py --pages 0           # 全ページ（1000冊）
python scrape_books.py --delay 2           # ページごとに2秒待つ
python scrape_books.py --output 結果.xlsx  # 出力ファイル名を変える
```

## 実行結果のサンプル

実行時の画面表示：

```
取得中 (1ページ目): https://books.toscrape.com/catalogue/page-1.html
取得中 (2ページ目): https://books.toscrape.com/catalogue/page-2.html
取得中 (3ページ目): https://books.toscrape.com/catalogue/page-3.html

取得が完了しました: 60 冊
  Excel: output/書籍一覧.xlsx
  CSV  : output/書籍一覧.csv
```

出力Excel（「書籍一覧」シート）の先頭部分：

| タイトル | 価格(ポンド) | 評価(星) | 在庫 | URL |
|---|---:|---:|---|---|
| A Light in the Attic | 51.77 | 3 | あり | https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html |
| Tipping the Velvet | 53.74 | 1 | あり | https://books.toscrape.com/catalogue/tipping-the-velvet_999/index.html |
| Soumission | 50.10 | 1 | あり | https://books.toscrape.com/catalogue/soumission_998/index.html |
| Sharp Objects | 47.82 | 4 | あり | https://books.toscrape.com/catalogue/sharp-objects_997/index.html |
| Sapiens: A Brief History of Humankind | 54.23 | 5 | あり | https://books.toscrape.com/catalogue/sapiens-a-brief-history-of-humankind_996/index.html |

## テスト

```
python -m pytest
```

本物のサイトにはアクセスせず、`tests/html/` に保存したHTMLを使って
情報の取り出し、ページのたどり方、待ち時間、ファイル出力、エラー時の動作を確認します。

## ファイル構成

```
book-price-scraper/
├── scrape_books.py       # スクレイピングツール本体
├── requirements.txt      # 必要なライブラリ
└── tests/
    ├── test_scrape_books.py  # 自動テスト
    └── html/                 # テスト用のHTML
```

## カスタマイズの例

ご依頼に合わせて、次のような変更が可能です。

- 対象サイト・取得項目の変更（商品詳細ページの説明文や画像URLも取得、など）
- 前回との価格比較や、値下がり商品だけを抽出
- 毎日・毎週の自動実行
- Googleスプレッドシートへの書き込み

※ スクレイピングは、対象サイトの利用規約と robots.txt で許可されている範囲でのみ行います。

## 使用技術

Python 3 / requests / BeautifulSoup4 / pandas / openpyxl / pytest
