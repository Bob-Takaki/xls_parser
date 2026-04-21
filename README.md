# xls_parser 📊

Pandas不要！Pythonの標準ライブラリだけで動く、超軽量な古いExcelファイル（`.xls` / OLE2・BIFFフォーマット）専用のパーサーです。

「`.xls` ファイルからデータを抽出したいだけなのに、Pandasや他の巨大なライブラリを入れると動作が重くなる…」という課題を解決するためにゼロから開発されました。

## ✨ 特徴 (Features)
- **超軽量＆高速:** 外部ライブラリ（Dependencies）は一切不要。Python標準の `struct` と `os` のみで動作します。
- **メモリに優しい:** セルごとの重いオブジェクトを生成せず、シンプルな「2次元配列（リストのリスト）」として一括でデータを返します。
- **SST完全対応:** BIFFフォーマット特有の複雑な文字列辞書（Shared String Table）も内部で自動解決します。

## 📦 インストール (Installation)

GitHubから直接 `pip install` することができます。

```bash
pip install git+[https://github.com/Bob-Takaki/xls_parser.git](https://github.com/Bob-Takaki/xls_parser.git)
