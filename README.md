# README

このツールは特定のディレクトリ配下のすべてのExcelファイル（.xlsx, .xls）から、指定した文字列を検索し該当箇所を検索するツールです。

## 特徴
- サブディレクトリも含めて再帰的に検索
- 検索結果はファイル名・パス・列名・該当行を表示
- エラーが発生したファイルも一覧表示

## 必要な環境
- Python 3.7以上
- 必要なパッケージは `requirements.txt` を参照

## 使い方

1. 検索対象ディレクトリのパスを環境変数 `SEARCH_DIR` に設定してください。

```sh
export SEARCH_DIR=/path/to/search/dir
```

2. 必要なパッケージをインストールします。

```sh
pip install -r requirements.txt
```

3. スクリプトを実行します。

```sh
python main.py <検索文字列>
```

例：

```sh
python main.py サンプル
```

## 出力例

```
検索文字列: サンプル
検索ディレクトリ: /path/to/search/dir

=== 検索結果 ===

【ファイル】: example.xlsx
【パス】: /path/to/search/dir/example.xlsx
【列名】: コメント
【該当箇所】:
  行 2: サンプルコメント
  行 5: サンプルデータ

【エラーが発生したファイル一覧】:
- /path/to/search/dir/broken.xlsx
```
