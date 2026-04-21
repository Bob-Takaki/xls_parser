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
```

## 🚀 使い方 (Usage)

使い方は極めてシンプルです。たった1つの関数を呼び出すだけで、シートのデータが2次元配列として取得できます。

```python
from xls_parser import read_xls_as_array

# .xls ファイルを読み込む
file_path = "sample.xls"
data = read_xls_as_array(file_path)

# データの確認
print(f"行数: {len(data)}, 列数: {len(data[0])}")

# 最初の5行を表示
for row in data[:5]:
    print(row)
```

## ⚠️ 注意点 (Notes)
- このライブラリは古いバイナリ形式である **`.xls` (Excel 97-2003)** 専用です。新しい `.xlsx` 形式には対応していません。
- 複数シートがある場合、現在のアプローチでは一番最初に見つかったWorkbookストリームのデータを抽出します。
- 書式設定、文字色、罫線などのメタデータは無視し、「純粋なセルの中身（テキスト・数値）」のみを高速に抽出することに特化しています。

## 📄 ライセンス (License)
MIT License
