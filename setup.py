# setup.py
from setuptools import setup

setup(
    name="xls_parser",           # ← ここをカッコいい名前に変更！
    version="0.1.0",
    py_modules=["xls_parser"],   # ← 読み込むファイル名も変更
    description="Pandasを使わない軽量なxlsパーサー",
)
