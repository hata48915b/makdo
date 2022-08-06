<!-- Time-stamp:   <2022.08.06-09:09:14-JST> -->

# MAKDO

〜〜 わずらわしいナンバリングやインデント（字下げ）から、あなたを解放します 〜〜

## これは何？

**一言で言うと、訴状、答弁書、準備書面、判決書のMS Wordファイルが
簡単に作成できるアプリです。**

**MS Wordや一太郎は使いませんので、なくても大丈夫です。**

Markdown形式を使って、
公用文書書式（司法、行政、立法）のMS Word形式のファイルを読み書きします。
Markdown形式のファイルから公用文書書式のMS Word形式のファイルを作成したり、
公用文書書式のMS Word形式のファイルからMarkdown形式のファイルを作成できます。

Markdown形式のファイルから公用文書書式のMS Word形式のファイルを作成することで、
次のような恩恵があります。

- 「第1、第2…」や「ア、イ…」などのナンバリングを自動でやってくれるようになります
- インデント（字下げ）を自動でやってくれるようになり、手作業で設定する必要がなくなります
- `grep`コマンドを使って、過去に作成した書面に検索をかけられるようになります
- `diff`コマンドを使って、原稿間の違いを確認できるようになります
- 定型文書をコンピューターで自動で作成できるようになります
- MS Wordがなくてもe内容証明（電子内容証明）が送れるようになります
- LinuxやFreeBSDで弁護士の仕事ができるようになります

**見出し機能などのMS Wordの特別な機能は、あえて使っていません。
これは、見出し機能などの特別な機能を使わなくても公用文書書式を実現できる一方で、
ほとんどの方が見出し機能などの難しい機能を使っておらず使いこなせないためです。
pandocなどの優れたソフトを使わずに、あえて自作した理由もその点にあります。**

↓↓↓ Markdown形式のファイルのサンプル画像

![Markdown形式のファイルのサンプル画像](image/sojo-md.png "Markdown形式のファイルのサンプル画像")

↓↓↓ 公用文書書式のMS Word形式のファイルのサンプル画像

![公用文書書式のMS Word形式のファイルのサンプル画像](image/sojo-pdf.png "公用文書書式のMS Word形式のファイルのサンプル画像")

なお、Markdown形式と公用文書書式のMS Word形式についてお知りになりたい方は、
後記の「Markdown形式とは」と「公用文書書式のMS Word形式とは」をご覧ください。

## とりあえず使いたい方へ

簡易実行ファイルを用意しましたので、こちらではなく、そちらをお使いください。

下記のZIPファイルをダウンロードして、展開してください。

[Windows用簡易実行ファイル](https://hata-o.jp/kian/index?tools#makdo-windows.zip)

[macOS用簡易実行ファイル](https://hata-o.jp/kian/index?tools#makdo-macos.zip)

展開したフォルダの中に実行ファイルがあります。

"makdo-md2docx"
（Windowsの場合は"makdo-md2docx.bat"、macOSの場合は"makdo-md2docx.app"）に、
Markdown形式のファイルをドラッグ＆ドロップすると、
公用文書書式のMS Word形式のファイルが作成されます。

"makdo-docx2md"
（Windowsの場合は"makdo-docx2md.bat"、macOSの場合は"makdo-docx2md.app"）に、
公用文書書式のMS Word形式のファイルをドラッグ＆ドロップすると、
Markdown形式のファイルが作成されます。

なお、"makdo-docx2pdf"
（Windowsの場合は"makdo-docx2pdf.bat"、macOSの場合は"makdo-docx2pdf.app"）に、
公用文書書式のMS Word形式のファイルをドラッグ＆ドロップすると、
PDF形式のファイルが作成されます。
ただし、LibreOfficeがインストールされていることが必要です。

[LibreOffice](https://ja.libreoffice.org/download/download/)

## 動作環境

Python3とpython-docxとchardetが動作すれば、
Windowsでも、macOSでも、Linuxでも動作します
（Windows 10、macOS Mojave、Ubuntu 20.04では動作確認済みです。）。

[Pythonの公式サイト](https://www.python.org/)

[python-docxのページ](https://package.wiki/python-docx)

### Python3のインストール

Python3をインストールする必要があります。

なお、macOSにはPython2がインストールされていますが、
Python2では動作しませんので、Python3をインストールする必要があります。

下記の公式サイトのダウンロードページから、
インストーラーをダウンロードして実行し、インストールしてください。

[Windows用のダウンロードサイト](https://www.python.org/downloads/windows/)

[macOS用のダウンロードサイト](https://www.python.org/downloads/macos/)

インストーラ−は、
Windowsの場合は"python-(VERSION)-amd64.exe"、
macOSの場合は"python-(VERSION)-macosx(macOSのVERSION).pkg"です。

なお、公式ストアーやパッケージ管理ソフトからも、
Python3をインストールできるはずで、
もちろんそちらを使ってインストールした方が管理が楽かもしれません。

### python-docxのインストール

python-docxをインストールする必要があります。

コマンドプロンプト（Windowsの場合）又はターミナル（macOSの場合）で、
次のコマンドを実行してインストールしてください。

```
python -m pip install python-docx
```

SSL認証のエラーが出る場合は、`--trusted-host`を付けて、
下記のコマンドを実行してください
（見やすくするため改行していますが、実行する際には1行で入力してください。）。

```
python -m pip install python-docx
    --trusted-host pypi.python.org
    --trusted-host files.pythonhosted.org
    --trusted-host pypi.org
```

### chardetのインストール

文字コードはUTF-8を想定しています。

しかし、現時点ではShift_JISも広く使われているため、
chardetを使って入力ファイルの文字コードを判別し、対応するようにしています。

コマンドプロンプト（Windowsの場合）又はターミナル（macOSの場合）で、
次のコマンドを実行してインストールしてください。

```
python -m pip install chardet
```

SSL認証のエラーが出る場合は、`--trusted-host`を付けて、
下記のコマンドを実行してください
（見やすくするため改行していますが、実行する際には1行で入力してください。）。

```
python -m pip install chardet
    --trusted-host pypi.python.org
    --trusted-host files.pythonhosted.org
    --trusted-host pypi.org
```

## インストール

本アプリは、Python3で書かれたプログラムファイルを実行するだけなので、
インストールは必要ありません。

## 実行方法

### 公用文書書式のMS Word形式のファイルを作成する場合

Markdown形式のファイルを用意し
（"OOOO.md"というファイル名だと仮定します。）、
"makdo-md2docx.py"のあるフォルダに移動させ、次のように実行してください。
"XXXX.docx"というファイル名で、
公用文書書式のMS Word形式のファイルが作成されます。

```
python3 makdo-md2docx.py OOOO.md XXXX.docx
```

オプションについては、`--help`でご確認ください。

```
python3 makdo-md2docx.py --help
```

### Markdown形式のファイルを作成する場合

公用文書書式のMS Word形式のファイルを用意し
（"XXXX.docx"というファイル名だと仮定します。）
"makdo-md2docx.py"のあるフォルダに移動させ、次のように実行してください。
"OOOO.md"というファイル名で、
Markdown形式のファイルが作成されます。

```
python3 makdo-docx2md.py XXXX.docx OOOO.md
```

オプションについては、`--help`でご確認ください。

```
python3 makdo-docx2md.py --help
```

## Markdown形式とは

Markdown形式とは、テキストファイル形式で、次のような書式のファイルです。

```
# 訴状（←タイトル）

令和元年２月３日 :（←右寄せ）

: 広島地方裁判所　御中（←左寄せ）

<=4（全体を左に４文字ずらす）
原告訴訟代理人弁護士　秦　誠一郎 :（←右寄せ）

v=1（←1行下げる）
## 請求の趣旨（←第１、第２…）

### 被告は原告に対し１万円支払え。（←１、２…）

### 訴訟費用は被告の負担とする。（←１、２…）

<<=1 <=1（←１行目の字下げをなくす、全体を左に１文字ずらす）
との判決及び仮執行宣言を求める。

v=1（←1行下げる）
## 請求の原因（←第１、第２…）

…

v=1（←１行下げる）
: 添付書類 :（←中央寄せ）

1. 訴状副本　　１通（←箇条書）
1. 資格証明書　１通（←箇条書）
```

詳細は、"sample"の中のファイルを参考にしてください。

なお、このREADME自体も、Markdown形式で書かれています。

## 公用文書書式のMS Word形式とは

公用文書書式のMS Word形式とは、
マイクロソフト社が開発し販売するワープロソフトWordが使用するファイル形式で、
次のような書式のファイルです。

```
　　　　　　　　　　　　　　　　　　訴状

　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　令和元年２月３日
広島地方裁判所　御中
　　　　　　　　　　　　　　　　　　原告訴訟代理人弁護士　秦　誠一郎

第１　請求の趣旨
　１　被告は原告に対し１万円支払え。
　２　訴訟費用は被告の負担とする。
　との判決及び仮執行宣言を求める。

第２　請求の原因
　…

　　　　　　　　　　　　　　　　　添付書類
１．訴状副本　　１通
２．資格証明書　１通

```

## 実装できていない機能

### 段落内の改行幅を広げたときの段落の前後のスペース

python-docxでは、段落内の改行幅を広げた場合、
自動的に段落の前後のスペースが広がってしまいます。

段落内の改行幅を狭めた場合も、自動的に段落の前後のスペースの幅が狭まりますが、
こちらは手動で段落の前後のスペースを広げることにより、調整しています。

しかし、段落の前後のスペースは負の値を持つことができませんので、
段落内の改行幅を広げた場合、手動で段落の前後のスペースを狭めることができません。

その結果、段落内の改行幅を広げた場合、段落の前後のスペースが広がってしまいます。
しかも、前後で広がり方が均一でないため、見た目が悪くなってしまいます。

これは、python-docxの仕様（もしかしたらMS Wordの仕様？）ですので、
現時点では対応不可能です。

### 均等割付け

均等割付け（「広　島　太　郎」のような文字列を一定の幅で均等に広げること）は、
実装できておりません。

私も、当初はかなり気になりました。

python-docxでの実装は不可能ですが、
出来上がったWord形式のファイルを直接編集すれば、
実装できなくはないように思います。

しかし、今では、本当に必要な機能なのか、疑問に思っています。

実際、均等割付けをせずに文書を作成していますが、何のトラブルも生じていません。

### 行番号の文字の大きさ

技術上の問題で、実用上は問題ありませんが、
行番号の文字の大きさを、行番号用のスタイルで変更できていません。

仕方なく、Normalスタイルの文字を小さくし、別にスタイルを作成して、
本文はそのスタイルを適用してごまかしています。

この方法はスマートではないので、今後、なんとかしたいと考えています。

## 著作権

Copyright © 2022  Seiichiro HATA

## ライセンス

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

## 連絡先

[Seiichiro HATA](<mailto:infotech@hata-o.jp>)

## ウェブページ

[makdo](https://hata-o.jp/kian/index?tools#makdo)

## 名前の由来

"MAKe DOcx"です。

## ヒストリー

### 2022.07.21 v01 Hiroshima

