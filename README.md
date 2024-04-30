<!-- Time-stamp:   <2024.04.30-11:29:18-JST> -->

# MAKDO（MS Wordの互換アプリ）

![MAKDOのLOGO](https://raw.githubusercontent.com/hata48915b/makdo/main/image/md8docx.png "MAKDOのLOGO")

〜〜 Markdown形式からMS Word形式へ、MS Word形式からMarkdown形式へ 〜〜

〜〜 MS Wordがなくても、MS Word形式のファイルを読み書きできます 〜〜

〜〜 わずらわしいナンバリング（番号付け）やインデント（字下げ）から、あなたを解放します 〜〜

## 注意事項

### 簡易実行ファイルについて

githubのレポジトリは、開発者向けです。

一般の方は、下記のリンクから簡易実行ファイルをダウンロードして使ってください。

[簡易実行ファイルのダウンロードページ](https://hata-o.jp/kian/index?tools#makdo)

環境の設定やアプリのインストールが不要で使えます。

※ 現在、macOSの開発環境がないため、macOS用の簡易実行ファイルはありません。
macOSの方は`python`の設定が必要です。

### Python3について

簡易実行ファイルは単独で実行できますので不要ですが、
このアプリ（ライブラリ）は、Python3（Pythonのバージョン3）以上を想定しています。

以下の説明では`python`という表記は、システムによっては`python3`等を意味する場合があります。

`python`をうまく実行できない場合には、`python`の実行ファイルをフルパスで指定するか、
環境変数PATHの設定をお願いします。

### MS Wordの機能の実装状況について

MS Wordの機能のうち、
日常使用する機能のかなりの部分は実装できていますが、
実装できていない機能もたくさんあります。

## とりあえず使いたい方へ

※ 現在、macOSの開発環境がないため、macOS用の簡易実行ファイルはありません。
どなたかご協力いただけますと助かります。

こんな感じで、Markdown形式のファイルからMS Word形式のファイルを作ったり、その逆ができたりします。

↓↓↓ 簡易実行ファイルの実行の様子（Windowsの場合）

![簡易実行ファイルの実行の様子](https://raw.githubusercontent.com/hata48915b/makdo/main/image/simple-exec.gif "簡易実行ファイルの実行の様子")

とりあえず使いたい方のために、簡易実行ファイルを用意しました。

簡易実行ファイルであれば、環境の設定やアプリのインストールが不要で、
しかも、コマンドプロンプト（Windowsの場合）やターミナル（macOSの場合）上のコマンドではなく、
マウスを使ってドラッグ＆ドロップで実行できます。

下記にアクセスしてZIPファイルをダウンロードし、展開してください。

[簡易実行ファイルのダウンロードページ](https://hata-o.jp/kian/index?tools#makdo)

展開したフォルダの中に実行ファイル
（Windowsの場合は`makdo_win.exe`、macOSの場合は`makdo_mac`。）があります。

これを起動するとウィンドウが開きますので、そのウィンドウに、
Markdown形式のファイルをドラッグ＆ドロップすると、MS Word形式のファイルが作成され、
MS Word形式のファイルをドラッグ＆ドロップすると、Markdown形式のファイルが作成されます。

## これは何？

**一言で言うと、テキストエディタ（メモ帳）で編集した原稿から、
公用文書書式（訴状、答弁書、準備書面、判決書等）のMS Wordファイルを簡単に作成できるアプリです。**

**MS Wordや一太郎は使いませんので、なくても大丈夫です。**

Markdown形式（テキスト形式）を使って、
公用文書書式（司法、行政、立法）のMS Word形式のファイルを読み書きします。
Markdown形式のファイルから公用文書書式のMS Word形式のファイルを作成したり、
公用文書書式のMS Word形式のファイルからMarkdown形式のファイルを作成できます。

Markdown形式のファイルから公用文書書式のMS Word形式のファイルを作成することで、
次のような恩恵があります。

- 「第1、第2…」や「ア、イ…」などのナンバリングを自動でやってくれるようになります
- インデント（字下げ）を自動でやってくれるようになり、手作業で設定する必要がなくなります
- 草冠が「十十」の花などの異字体（人名漢字）を「花6;」などと書くことで簡単に使えます
- 簡単なスクリプトが使えるので、金額などを自動で計算してくれるようになります（超便利！）
- 簡単なスクリプトが使えるので、1か所を変更したら、他の所にも反映するようにできます
- `grep`コマンドを使って、過去に作成した書面に検索をかけられるようになります
- `diff`コマンドを使って、原稿間の違いを確認できるようになります
- 定型文書をコンピューターで自動で作成できるようになります
- LinuxやFreeBSDで弁護士の仕事ができるようになります

**見出し機能などのMS Wordの特別な機能は、あえて使っていません。
これは、見出し機能などの特別な機能を使わなくても公用文書書式を実現できる一方で、
ほとんどの方が見出し機能などの難しい機能を使っておらず使いこなせないためです。
pandocなどの優れたソフトを使わずに、あえて自作した理由もその点にあります。**

↓↓↓ Markdown形式のファイルのサンプル画像

![Markdown形式のファイルのサンプル画像](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-md.png "Markdown形式のファイルのサンプル画像")

↓↓↓ 公用文書書式のMS Word形式のファイルのサンプル画像

![公用文書書式のMS Word形式のファイルのサンプル画像](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-pdf.png "公用文書書式のMS Word形式のファイルのサンプル画像")

なお、Markdown形式と公用文書書式のMS Word形式についてお知りになりたい方は、
後記の「Markdown形式とは」と「公用文書書式のMS Word形式とは」をご覧ください。

## 動作環境

ここの記載は、開発者向けです。
簡易実行ファイルをご利用の方は、関係がありません。

Python3とpython-docx、chardet、tkinterdnd2が動作すれば、
Windowsでも、macOSでも、Linuxでも動作します。

[Pythonの公式サイト](https://www.python.org/)

[python-docxのページ](https://pypi.org/project/python-docx/)

### Python3のインストール

Python3をインストールする必要があります。

なお、macOSにはPython2がインストールされていますが、
Python2では動作しませんので、Python3をインストールする必要があります。

下記の公式サイトのダウンロードページから、
インストーラーをダウンロードして実行し、インストールしてください。

[Windows用のダウンロードページ](https://www.python.org/downloads/windows/)

[macOS用のダウンロードページ](https://www.python.org/downloads/macos/)

インストーラ−は、
Windowsの場合は`python-(VERSION)-amd64.exe`、
macOSの場合は`python-(VERSION)-macosx(macOSのVERSION).pkg`です。

なお、公式ストアーやパッケージ管理ソフトからも、
Python3をインストールできるはずで、
もちろんそちらを使ってインストールした方が管理が楽かもしれません。

### Makdoのインストール

下記のライブラリをインストールする必要があります。

- python-docx / MS Word形式のファイルを作成するために必要です。
- chardet     / 入力ファイルの文字コードを判別するために必要です。
- tkinterdnd2 / GUIのために必要です。
- makdo       / 本アプリです。

コマンドプロンプト（Windowsの場合）又はターミナル（macOSの場合）で、
次のコマンドを実行してインストールしてください。

```
python -m pip install python-docx chardet tkinterdnd2 makdo
```

SSL認証のエラーが出る場合は、`--trusted-host`を付けて、
下記のコマンドを実行してください
（見やすくするため改行していますが、実行する際には1行で入力してください。）。

```
python -m pip install python-docx chardet tkinterdnd2 makdo
    --trusted-host pypi.python.org
    --trusted-host files.pythonhosted.org
    --trusted-host pypi.org
```

### 実行ファイル`makdo.py`の作成

`makdo.py`という名前で、次の内容のファイルを作成します。

```
#!/usr/local/bin/python3
from makdo import Makdo
Makdo()
```

## 基本的な使い方

### GUIによる場合

簡易実行ファイル（Windowsの場合は`makdo_win.exe`、macOSの場合は`makdo_mac`）又は
実行ファイル`makdo.py`をクリックして起動します。

開いたウィンドウに、
Markdown形式のファイルをドラッグ＆ドロップすると、MS Word形式のファイルが作成され、
MS Word形式のファイルをドラッグ＆ドロップすると、Markdown形式のファイルが作成されます。

### CUIによる場合

#### Markdown形式からMS Word形式へ

`md2docx.py`という名前で、次のような内容のファイル（スクリプト）を作成します。

```
#!/usr/bin/python3
import sys
from makdo import Md2Docx
m2d = makdo.Md2Docx(sys.argv[1])
m2d.save(sys.argv[2])
```

Markdown形式のファイルを用意し、
コマンドプロンプト（Windowsの場合）又はターミナル（macOSの場合）で、
次のように実行します。

```
md2docx.py foo.md bar.docx
```

`foo.md`というMarkdown形式のファイルから、`bar.docx`というMS Word形式のファイルが作成されます。

#### MS Word形式からMarkdown形式へ

`docx2md.py`という名前で、次のような内容のファイル（スクリプト）を作成します。

```
#!/usr/bin/python3
import sys
from makdo import Docx2Md
d2m = Docx2Md(sys.argv[1])
d2m.save(sys.argv[2])
```

MS Word形式のファイルを用意し、
コマンドプロンプト（Windowsの場合）又はターミナル（macOSの場合）で、
次のように実行します。

```
docx2md.py bar.docx foo.md
```

`bar.docx`というMS Word形式のファイルから、`foo.md`というMarkdown形式のファイルが作成されます。

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

サンプル画像を用意しましたので、
"image"フォルダの中の"sojo-md.png"をご覧ください。

↓↓↓ Markdown形式のファイルのサンプル画像

![Markdown形式のファイルのサンプル画像](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-md.png "Markdown形式のファイルのサンプル画像")

"sample"フォルダの中にサンプルファイルを用意しておりますので、
そちらもご覧ください。

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
①　訴状副本　　１通
②　資格証明書　１通

```

サンプル画像を用意しましたので、
"image"フォルダの中の"sojo-pdf.png"をご覧ください。

↓↓↓ 公用文書書式のMS Word形式のファイルのサンプル画像

![公用文書書式のMS Word形式のファイルのサンプル画像](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-pdf.png "公用文書書式のMS Word形式のファイルのサンプル画像")

"sample"フォルダの中にサンプルファイルを用意しておりますので、
そちらもご覧ください。

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

均等割付けとは、「秦　　誠 一 郎」のように、
文字列を一定の幅で均等に広げることです。
日本では、人名などで見た目を整えるために、昔から広く多用されてきました。

均等割付けは、XMLのタグを直接埋め込むことにより、理論上、実装可能です。

ただし、MS Wordの仕様のために、美しい実装は不可能です。

そもそも、均等割付けは、見た目だけの問題にすぎないうえ、
AIが書面を読んだ際に意味の理解の妨げになる可能性がありますので、
原則として使うべきではないと考えております。

そのため、均等割付けは実装しておりません。

どうしても均等割付けが必要な場合には、
`<N>`（`N`は数字で、小数も可）で任意の空白を入れられるようにしておりますので、
こちらを使って割り付けてください。

例えば「秦　　誠 一 郎」は、「秦<2.0>誠<0.5>一<0.5>郎」となります。

### 行番号の文字の大きさ

技術上の問題で、実用上は問題ありませんが、
行番号の文字の大きさを、行番号用のスタイルで変更できていません。

仕方なく、Normalスタイルの文字を小さくし、別にスタイルを作成して、
本文はそのスタイルを適用してごまかしています。

この方法はスマートではないので、今後、なんとかしたいと考えています。

## 著作権

Copyright © 2022-2024  Seiichiro HATA

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

## 自己紹介

### 名前

秦　誠一郎（はた　せいいちろう）

### 本業

広島弁護士会所属の弁護士です。

国政法律事務所（広島市中区上八丁堀5番27号アーバンビュー上八丁堀701号）に勤務しています。

### 経歴

#### 広島県立安古市高校

#### 大阪大学理学部物理学科（高エネルギー物理学）

#### 京都大学大学院理学研究科（数理解析研究所）

### 連絡先

[弁護士 秦 誠一郎の公式ページ](https://hata-o.jp/hata48915b/)

[弁護士 秦 誠一郎の連絡先](<mailto:hata48915b@post.nifty.jp>)

## ウェブページ

[ソースコードのページ（github）](https://github.com/hata48915b/makdo)

[ライブラリのページ（PyPI）](https://pypi.org/project/makdo/)

[簡易実行ファイルのページ](https://hata-o.jp/kian/index?tools#makdo)

## 名前の由来

"MAKDO"は、"MAKe DOcx"と"MArKDOwn"を兼ねています。

"魔苦怒"は、
「このアプリの"魔"法で、
ナンバリング（番号付け）やインデント（字下げ）の"苦"しみや"怒"りから、
皆様を解放したい」
という思いで、名付けました。

## ヒストリー

### 2022.07.21 v01 Hiroshima リリース

最初のリリースです。

### 2022.08.24 v02 Shin-Hakushima リリース

修正を加えたリリースです。

### 2022.12.25 v03 Yokogawa リリース

修正を加えたリリースです。

### 2023.01.07 v04 Mitaki リリース

日本語対応を強化しました。

### 2023.03.16 v05 Aki-Nagatsuka リリース

修正を加えたリリースです。

### 2023.06.07 v06 Shimo-Gion リリース

大幅な修正を加えたリリースです。

### 2024.04.02 v07 Furuichibashi リリース

大幅な修正を加えたリリースです。

