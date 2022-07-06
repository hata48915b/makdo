# sample

サンプルデータです。

## 一覧

下記を用意しました。

- sojo.md
- shoko.md
- wakai.md
- kisoku.md
- header.md

## データの特徴

### sojo.md

訴状のサンプルデータです。

もっとも一般的なデータです。

### shoko.md

証拠説明書のサンプルデータです。

表を含んでいます。

### wakai.md

和解契約書のサンプルデータです。

契約書式のデータです。

### kisoku.md

運動会規則のサンプルデータです。

条文書式のデータです。

### header.md

ヘッダーのサンプルデータです。

参考にしてください。

## Word形式ファイルの作成方法

### 簡易実行

例えば"sojo.md"を使って自動で実行する場合には、
次のように実行してください。

"sojo.md"を、"makdo-md2docx"
（Windowsの場合は"makdo-md2docx.bat"、macOSの場合は"makdo-md2docx.app"）に、
Markdown形式のファイル（例えば"test.md"）を、ドラッグ＆ドロップしてください。

MS Word形式のファイル"sojo.docx"が作成されます。

### 手動実行

まず、Python3とpython-docxをインストールしてください。

[Pythonの公式サイト](https://www.python.org/)

[python-docxのページ](https://package.wiki/python-docx)

その後、例えば"sojo.md"を使って手動で作成する場合には、
次のように実行してください。

'''
python3 makdo-md2docx.py sojo.md sojo.docx
'''

MS Word形式のファイル"sojo.docx"が作成されます。
