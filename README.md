<!-- Time-stamp:   <2025.02.23-09:47:04-JST> -->

# Makdo（MS Word形式のファイルをMarkdownで作成・編集するエディタアプリ）

[公式ウェブページ](https://makdo.jp/)

[公式レポジトリ](https://github.com/hata48915b/makdo/)

- [［実行ファイル（Windows）のダウンロード］](https://makdo.jp/makdo_win.zip)

- [［実行ファイル（macOS）のダウンロード］](https://makdo.jp/makdo_mac.zip)

- [［実行ファイル（Linux）のダウンロード］](https://makdo.jp/makdo_lin.zip)

![MakdoのLOGO](https://raw.githubusercontent.com/hata48915b/makdo/main/image/makdoL.png "MakdoのLOGO")

〜〜 LaTeXの論理性とEmacsの機能性を、MS Word形式のファイルの作成・編集に！ 〜〜

〜〜 MS Wordがなくても、MS Word形式のファイルを読んだり編集できたりします 〜〜

〜〜 わずらわしいナンバリング（番号付け）やインデント（字下げ）から、あなたを解放します 〜〜

## このアプリの概要

このアプリは、
Markdownという記法（ウェブページのHTMLを簡単にして使いやすくした記法）を使って、
MS Word形式（拡張子"docx"）のファイルを読んだり編集したりしようという野心的なアプリです。

専用エディタが付属しており（下の画像参照）、
① MS Word形式のファイルをMarkdown形式に変換して開いたり、
② Markdown形式のファイルをそのまま開いたり、
③ Markdown形式の原稿を編集したり、
④ 編集した原稿をMS Word形式に変換して保存したり、
⑤ 編集した原稿をそのままMarkdown形式のファイルに保存したり、
簡単にできます。

↓↓↓ Makdoで原稿を編集する様子

![Makdoで原稿を編集する様子](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-makdo.png "Makdoで原稿を編集する様子")

↓↓↓ Makdoで編集した原稿をMS Wordで開いた様子

![Makdoで編集した原稿をMS Wordで開いた様子](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-msword.png "Makdoで編集した原稿をMS Wordで開いた様子")

## このアプリのメリット

Markdown形式を使って、MS Word形式のファイルを読み書きします。
いわゆるワープロソフトの代わりになります。

**MS Wordや一太郎がなくても、このアプリは使えます。**

このアプリには、次のようなメリットがあります。

- 「第1、第2…」や「ア、イ…」などのナンバリングを自動でやってくれるようになります
- インデント（字下げ）を自動でやってくれるようになり、手作業で設定する必要がなくなります
- 草冠が「++」の花などの異字体（人名漢字）を、「花6;」などと書くことで簡単に使えます
- 簡単なスクリプトが使えるので、金額などを自動で計算してくれるようになります（超便利！）
- 簡単なスクリプトが使えるので、1か所を変更したら、他の所にも反映するようにできます
- `grep`コマンドを使って、過去に作成した書面に検索をかけられるようになります
- `diff`コマンドを使って、原稿間の違いを確認できるようになります
- 原稿を比較する機能を使って、原稿間の違いを確認して、反映させることができます
- 定型文書をコンピューターで自動で作成できるようになります
- LinuxやFreeBSDでも、WindowsやmacOSと同じ仕事ができるようになります
- このアプリを間にかませることによって、MS WordとLibreOfficeとの互換性が高まります

## このアプリと相性の良い書面

このアプリは、
Markdown記法を使って文章を論理構造を指示したものを整形し、
MS Word形式のファイルに変換するものです。

そのため、論理的な文章の作成、具体的には
①法律家の書面の起案、
②官公庁の公文書の作成、
③研究者の学術論文の作成
などに向いています。

逆に、
新聞の折込広告などのように、デザインを重視する書面の作成には向いていません。

## このアプリの動作環境

Windows、macOS、Linux等で動作します。

## 実行ファイルを使って起動する方法

### Windowsの場合

一般の方でWindowsのパソコンをご使用の方は、
下記のリンクから実行ファイルのZIPファイルをダウンロードしてください。

[［実行ファイル（Windows）のダウンロード］](https://makdo.jp/makdo_win.zip)

ZIPファイルをダウンロード後、適当なフォルダに展開して、
トップフォルダにある`makdo_win.exe`をダブルクリックして、起動してください。

環境の設定やアプリのインストールが不要で、そのまま使えます。

### macOSの場合

一般の方でmacOSのパソコンをご使用の方は、
下記のリンクから実行ファイルのZIPファイルをダウンロードしてください。

[［実行ファイル（macOS）のダウンロード］](https://makdo.jp/makdo_mac.zip)

ZIPファイルをダウンロード後、適当なフォルダに展開して、
トップフォルダにある`makdo_mac`をダブルクリックして、起動してください。

環境の設定やアプリのインストールが不要で、そのまま使えます。

ただし、
macOSでは、ファイルのドラッグ＆ドロップ等、一部の機能が実現できていません。

## `pip`を使ってインストールして起動する方法

この方法は、Pythonに慣れている方向けです。

一般の方は、上記「実行ファイルを使って起動する方法」をお読みいただき、
実行ファイルを使って、アプリを起動してください。

この起動方法は、それほど難しくない上に、起動が早いというメリットがあります。

### Python3のインストール

まず、Python3をインストールする必要があります。

なお、macOSにはPython2がインストールされていますが、
Python2では動作しませんので、Python3をインストールする必要があります。

下記の公式サイトのダウンロードページから、
インストーラーをダウンロードして実行し、インストールしてください。

[Windows用のダウンロードページ](https://www.python.org/downloads/windows/)

[macOS用のダウンロードページ](https://www.python.org/downloads/macos/)

インストーラ−は、
Windowsの場合は`python-(VERSION)-amd64.exe`、
macOSの場合は`python-(VERSION)-macosx(macOSのVERSION).pkg`です。

なお、
公式ストアーやパッケージ管理ソフトからも、Python3をインストールできるはずです。
そちらを使ってインストールした方が、管理が楽なのでおすすめです。

### モジュールのインストール

コマンドプロンプト（Windowsの場合）又はターミナル（macOSの場合）で、
次のコマンドを実行して、このアプリのモジュールもインストールしてください
（必要なモジュールは自動的に追加でインストールされます。）。

```
python3 -m pip install makdo
```

### 実行ファイル`makdo.py`の作成

`makdo.py`という名前で、次の内容のファイルを作成します。

```
#!/usr/local/bin/python3
from makdo import Makdo
Makdo()
```

### 起動方法

`makdo.py`をダブルクリックするか、
MS Word形式のファイル（拡張子docx）又はMarkdown形式のファイル（拡張子md）を
ドラッグ＆ドロップするかして、起動してください。

### 生成AIを利用する方法

メニューバーの「裏の技」の
「OpenAIに質問（有料）」や「Llamaに質問（無料）」を利用するには、
追加でモジュールをインストールする必要があります。

```
pip install openai
pip install llama_cpp_python
```

Windowsでllama_cpp_pythonをインストールする場合、
事前に、`vs_BuildTools.exe`を下記のURLからダウンロードして、
インストールしておく必要があるかもしれません。

https://aka.ms/vs/17/release/vs_BuildTools.exe

## ソースファイルから起動する方法

この方法、開発者向けです。

一般の方は、上記の
「実行ファイルを使って起動する方法」や
「`pip`を使ってインストールして起動する方法」をお読みいただき、
実行ファイルを使って、アプリを起動してください。

ソースファイルからアプリを起動するためには、
Python3及び必要なmoduleをインストールする必要があります。

### Python3のインストール

上記「ソースファイルから起動する方法」の「Python3のインストール」に従って、
Python3をインストールしてください。

### モジュールのインストール

ソースファイルからアプリを起動するためには、
次のpython3のモジュールをインストールする必要があります。

- python-docx / MS Word形式のファイルを作成するために必要です。
- chardet / 入力ファイルの文字コードを判別するために必要です。
- tkinterdnd2 / GUIのために必要です。
- pywin32 / MS Wordやブラウザを起動するために必要です。
- Levenshtein / 2つの原稿を比較するために必要です。
- openpyxl / MS Excel形式の表を文書に取り込むために必要です。

コマンドプロンプト（Windowsの場合）又はターミナル（macOSの場合）で、
次のコマンドを実行してインストールしてください（python-docxの場合）。

```
python3 -m pip install python-docx
```

`python3`は、パソコンによっては`python`や`py`になっている場合があります。
`python3`で実行できない場合は、`python`や`py`に変えて実行してみてください。

また、
SSL認証のエラーが出る場合は、`--trusted-host`を付けて、
下記のコマンドを実行してください。

```
python3 -m pip install python-docx --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --trusted-host pypi.org
```

うまく実行できない場合には、`python3`の実行ファイルをフルパスで指定するか、
環境変数PATHの設定をお願いします。

### アプリケーションのインストール

Epwing形式の辞書を使うためには、`eblook`と辞書データも必要です。

[eblookのウェブページ](http://openlab.jp/edict/eblook/)

### 起動方法

`makdo.py`をダブルクリックするか、
MS Word形式のファイル（拡張子docx）又はMarkdown形式のファイル（拡張子md）を
ドラッグ＆ドロップするかして、起動してください。

### 生成AIを利用する方法

上記の「生成AIを利用する方法」をご参照ください。

## MS Word形式とMarkdown形式について

### MS Word形式とは

MS Word形式とは、
マイクロソフト社が開発し販売するワープロソフトWordが使用するファイル形式で、
次のような書式のファイルです。

```
　　　　　　　　　　　　　　　　　　訴状

　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　令和元年２月３日
広島地方裁判所　御中
　　　　　　　　　　　　　　　　　　原告訴訟代理人弁護士　秦　誠 一 郎　㊞

第１　請求の趣旨
　１　被告は、原告に対し、１０万円を支払え。
　２　訴訟費用は被告の負担とする。

第２　請求の原因
　１　総論
　　　本件の争点は、①本件交通事故発生についての過失割合、②原告車両のドア
　　パネルについて交換する必要があるか否か、③代車が必要であった期間は１週間
　　か１か月かの３点である。
```

サンプル画像を用意しましたので、
"image"フォルダの中の"sojo-msword.png"をご覧ください。

↓↓↓ MS Word形式のファイルのサンプル画像

![MS Word形式のファイルのサンプル画像](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-msword.png "MS Word形式のファイルのサンプル画像")

"sample"フォルダの中にサンプルファイルを用意しておりますので、
そちらもご覧ください。

### Markdown形式とは

Markdown形式とは、
テキスト形式のファイルで（メモ帳で開くことができます。）、
次のような書式のファイルです。

```
# 訴状（←タイトルを表示（文字を大きくして中央寄せ））

v=+0.5（←前の段落との間を0.5行空ける）
: 広島架空裁判所　御中（←左寄せ）

v=+0.5（←前の段落との間を0.5行空ける）
令和元年2月3日 :（←右寄せ）

v=+0.5 <=-16.0（←前の段落との間を0.5行空け、全体を右に16文字ずらす）
: 原告訴訟代理人弁護士　　秦　　誠<0.5>一<0.5>郎　㊞（←"誠"と"一"の間、"一"と"郎"の間を0.5文字空ける）

v=+1.0（←前の段落との間を1.0行空ける）
## 請求の趣旨（←"第１"とナンバリングして、1文字字下げする）

### （←"１"とナンバリングして、2文字字下げする）
被告は、原告に対し、10万円を支払え。

### （←"２"とナンバリングして、2文字字下げする）
訴訟費用は被告の負担とする。

v=+1.0
## 請求の原因（←"第２"とナンバリングして、1文字字下げする）

### 総論（←"１"とナンバリングして、2文字字下げする）

（↓単なる改行は無視されますので、読みやすい形で編集できます）
本件の争点は、
①本件交通事故発生についての過失割合、
②原告車両のドアパネルについて交換する必要があるか否か、
③代車が必要であった期間は1週間か1か月か
の3点である。
```

サンプル画像を用意しましたので、
"image"フォルダの中の"sojo-makdo.png"をご覧ください。

↓↓↓ Markdown形式のファイルのサンプル画像

![Markdown形式のファイルのサンプル画像](https://raw.githubusercontent.com/hata48915b/makdo/main/image/sojo-makdo.png "Markdown形式のファイルのサンプル画像")

"sample"フォルダの中にサンプルファイルを用意しておりますので、
そちらもご覧ください。

なお、このREADME自体も、Markdown形式で書かれています。

## 実装できていない機能

MS Wordの機能のうち、
日常使用する機能のかなりの部分は実装できていますが、
実装できていない機能もたくさんあります。

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

### 表や画像の上下の空白

MS Wordには、表や画像の上下の空白を調整する機能がありません。

そのため、MS Wordでは、
表や画像の上の空白を調整する代わりに、
表や画像の前の段落の下の空白を調整したり、
表や画像の下の空白を調整する代わりに、
表や画像の次の段落の上の空白を調整したりすることで、
代用しています。

本アプリでも同様の処理をしてすることで、代用しています。

### コメント

python-docxは、ページ外のコメントを実装していないため、
ページ外のコメントは使えません。

Makdoでは、コメントの内容を本文中に書き込むことで、代用しています。

### 脚注

python-docxは、ページ下部の脚注を実装していないため、
ページ下部の脚注は使えません。

Makdoでは、脚注の内容を本文中に書き込むことで、代用しています。

### 水平線の読み取り

MS Wordは、水平線は一般の段落の書式と扱います。

これに対し、
Markdownでは、水平線は独立の段落として扱います。

この違いにより、
MS Word形式のファイルをMarkdown形式のファイルに変換した際、
うまく変換できない場合があります。

### 画像の左右への文章の回り込み

python-docxは、現時点では、
画像の左右に文章を回り込ませることができません。

二段組で似たようなことができますので、
これで代用していただくのが良いと思います。

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

### 見出し機能

見出し機能は、あえて使っていません。

これは、
見出し機能などの特別な機能を使わなくても論理的な文章は書ける一方で、
ほとんどの方が見出し機能を使っておらず使いこなせないためです。

pandocなどの優れたソフトを使わずに、あえて自作した理由もその点にあります。

### テキストボックス

テキストボックスは、実装していません。

しかし、
テキストボックスは、原稿の論理性と整合しませんので、
論理的な文章を書く場合には、使うべきではないと考えています。

## 著作権

Copyright © 2022-2025  Seiichiro HATA

## ライセンス

GNU General Public Licenseバージョン3 (GPLv3)又はその後継バージョン

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

[公式ページ](https://makdo.jp/)

[ソースコードのページ（github）](https://github.com/hata48915b/makdo)

[ライブラリのページ（PyPI）](https://pypi.org/project/makdo/)

## 免責条項

ライセンスに定められているとおり、本プログラムにより損害が発生したとしても、
著作権者は何らの損害賠償責任も負いませんので、ご注意ください。

作成した文書は、必ず内容を確認し、
意図した内容になっていることを確認したうえで、使用してください。

## 名前の由来

"MAKDO"は、"MAKe DOcx"と"MArKDOwn"を兼ねています。

"魔苦怒"は、
「このアプリの"魔"法で、
ナンバリング（番号付け）やインデント（字下げ）の"苦"しみや"怒"りから、
皆様を解放したい」
という思いで、名付けました。

## ヒストリー

### 2022.07.21 v01 Hiroshima リリース

- 最初のリリースです。

### 2022.08.24 v02 Shin-Hakushima リリース

- 修正を加えたリリースです。

### 2022.12.25 v03 Yokogawa リリース

- 修正を加えたリリースです。

### 2023.01.07 v04 Mitaki リリース

- 日本語対応を強化しました。

### 2023.03.16 v05 Aki-Nagatsuka リリース

- 修正を加えたリリースです。

### 2023.06.07 v06 Shimo-Gion リリース

- 大幅な修正を加えたリリースです。

### 2024.04.02 v07 Furuichibashi リリース

- 大幅な修正を加えたリリースです。

### 2025.01.04 v08 Omachi リリース

- 大幅な修正を加えたリリースです。

- Markdownエディタが追加されました。
