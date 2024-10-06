#!/usr/bin/python3
# Name:         makdo_gui.py
# Version:      v07 Furuichibashi
# Time-stamp:   <2024.10.06-12:38:30-JST>

# makdo_gui.py
# Copyright (C) 2022-2024  Seiichiro HATA
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.


# 2022.07.21 v01 Hiroshima
# 2022.08.24 v02 Shin-Hakushima
# 2022.12.25 v03 Yokogawa
# 2023.01.07 v04 Mitaki
# 2023.03.16 v05 Aki-Nagatsuka
# 2023.06.07 v06 Shimo-Gion
# 2024.04.02 v07 Furuichibashi


# USAGE
# from makdo.makdo_gui import Makdo
# Makdo()


######################################################################
# SETTING


import sys
import os
if sys.platform == 'win32':
    import win32com.client  # pip install pywin32
    CONFIG_DIR = os.getenv('APPDATA') + '/makdo'
    CONFIG_FILE = CONFIG_DIR + '/init.md'
elif sys.platform == 'darwin':
    CONFIG_DIR = os.getenv('HOME') + '/Library/makdo'
    CONFIG_FILE = CONFIG_DIR + '/init.md'
elif sys.platform == 'linux':
    import subprocess
    import makdo.eblook  # epwing
    CONFIG_DIR = os.getenv('HOME') + '/.config/makdo'
    CONFIG_FILE = CONFIG_DIR + '/init.md'
else:
    CONFIG_DIR = None
    CONFIG_FILE = None

import shutil
import argparse     # Python Software Foundation License
import re
import unicodedata
import chardet      # GNU Lesser General Public License v2 or later (LGPLv2+)
import datetime     # Zope Public License
import zipfile
import tempfile
import tkinter
import tkinter.filedialog
import tkinter.simpledialog
import tkinter.messagebox
import tkinter.font
import tkinterdnd2  # MIT License
# from tkinterdnd2 import TkinterDnD, DND_FILES
import importlib    # Python Software Foundation License
import makdo.makdo_md2docx
import makdo.makdo_docx2md
import makdo.makdo_mddiff  # MDDIFF
import openpyxl     # MIT License
import webbrowser
import openai       # Apache Software License


__version__ = 'v07 Furuichibashi'


WINDOW_SIZE = '900x600'

GOTHIC_FONT = 'BIZ UDゴシック'         # 現時点で最適
MINCHO_FONT = 'BIZ UD明朝'
# GOTHIC_FONT = 'Noto Sans Mono CJK JP'  # 使えるがLinuxでは上下に間延びする
# MINCHO_FONT = 'Noto Serif CJK JP'
# GOTHIC_FONT = 'ＭＳ ゴシック'          # ボールドがないため幅が合わない
# MINCHO_FONT = 'ＭＳ 明朝'
# GOTHIC_FONT = 'IPAゴシック'            # ボールドがないため幅が合わない
# MINCHO_FONT = 'IPA明朝'

NOT_ESCAPED = '^((?:(?:.|\n)*?[^\\\\])??(?:\\\\\\\\)*?)??'

MD_TEXT_WIDTH = 68

ICON8_IMG = '''
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAABGdBTUEAALGPC/xhBQAAACBjSFJN
AAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAABKVBMVEUAAAAgHx4gHRsfHx8e
Hh4eHh4gKjogSIMgYL0gX7wfHx4gHh0gIB8gUZogYcEgHx8gYL4gIB8gIB8gYL4eX74gXrwfVaQf
YL4fYL4gX74gX74fX70gLkQgYL4fN1sgX74gICAfHx8xMTEvLy8gKTcaGhppaWnFxcVjY2MlJSUg
NVYfYMAgIB85OTmjo6P////19fWrq6tCQUAeKDVUVFTT09NGRkYeWa8gYMC+vr6Ojo4gNlYfISIc
HBt/f3/l5eUvLy4gYL4eJS4fNVYgVqdvmNVplNQgHx8fHyBvbm4fLUEfSogfV6keXbvF1u6rw+cd
X8EuasNCecrl7Pd/o9oaXL0gSoglY8BjkNKOrt7T4PI5csYeVaewx+j1+PxUhc6jveQgXrwxbMS3
5McxAAAAIHRSTlMACEqp5fzmqkoHGZTu8JSysweT7v385eWpS5QZ8PCpCBaEWb8AAAABYktHRC8j
1CARAAAAB3RJTUUH6AkbFzoTBLWr8gAAAPRJREFUGNMlj+lWwjAUhK+lKiBqS1NWl1ytgiYIqEBQ
lEWkgrUsigsoLu//ECbl/pvvnJk7AwCwooX0Vbq2Ho5AcNGNGB4cIjqbW9FAb1M8Os7lT/CUGZKY
cY5YOCuWyucXFjNM0GIcLyslCapOTVhhCFEsVOpXynJNhG2A3ri5bVZbLRXKWDsBnGO5Thsd5657
32NMgIsP/f7g0as9CX9IiA06RRx4IzIeT55fpiwpQ136Onp79z/82Zy0U/Kt2/mcfC2Ki9k3EVZa
FXM8BX7mRIhMRFXngaU3Fb9GdjlupytD/8huJruca2p7CZskU+l9Kf4BUNkolLs+E+EAAAAldEVY
dGRhdGU6Y3JlYXRlADIwMjQtMDktMjdUMjM6NTg6MTcrMDA6MDCQknpYAAAAJXRFWHRkYXRlOm1v
ZGlmeQAyMDI0LTA5LTI3VDIzOjU4OjE3KzAwOjAw4c/C5AAAAABJRU5ErkJggg==
'''

SPLASH_IMG = '''
iVBORw0KGgoAAAANSUhEUgAAAgAAAAEACAMAAADyTj5VAAAABGdBTUEAALGPC/xhBQAAACBjSFJN
AAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAClFBMVEX+/v74+Pfz8/Pr6+vj
4+Lb29vLy8u4uLazs7KkpKObm5uLi4uAgH50dHNjY2NQUE9ISEdLS0szMzMgIB4gICAgHh4gJCod
IiogKDYwN0EzPEkxQVY1R2BHWHFKXntKY4ZSbZNbeKFff6xpirh7mseBpdaHqNaLrNyiveK1yui+
0OvI1+/b5fPf6PTq8Pfs8fj2+Pvw8O/o6Ofg4N/Q0M/Dw8OYmJaJiYd4eHdra2pgX1w7OzowMC8r
KysoKCceHyAdHR0eICMdJjIeKTkfMUwcNVokQGolSHsgS4spV5ouYKg4a7c7c8VDecZNgMZaic1l
kdF3ntiBpdqrw+S6zenj6vW9vb2Tk5NaWlpAQD8dLUUdNVcdPGseRHwfTI8dVKQdWrMfYMAhYMAg
YL4oZb42cMRAdsV8odadueHD1OzV4fHU1NOrq6t7e3pTU1I4ODYgMUsgOmMgRHsdXr0gXrwybcNA
dshThMtlkc+hvN6Dg4JoaGYgJjEgKjogNlkgQXMgW7McXsAfYL8tacFKfcl8odiSsd3y9fogLkMg
OF0gUZsgWKwpZsBgjc54ntSvxujR3e9YWFZDQ0IgNFIgTZJEeciwxuVgYF4/dsjY2NcgV6hqldLI
yMZ4d3UdPnAgVKRzm9Xo7vcgRoAfOmHR3vAgPmpwcG0eTZK5t7QgSIScuN+mwORJfcfK2e6goJ+w
sK6oqKc4Pka4zOdyfIqZtuEtab2QkI6GqNo+QEMgUJd/g4YfWKyOj5BvmNMxbL6wxujm7fhci9DM
2vCUs+BgjtAgMEcdRoNOgMoeUJl4n9jo7vgeVqsfKDYvOUcuPVOAhpBjcYRjdZBRaIdJZIpKaZdJ
baE5YZoyXZwzYaYzZK4wZbL////iOcRPAAAAAWJLR0TbmQQWFAAAAAd0SU1FB+gKBgMlALOtGowA
ADStSURBVHja7X2LXxRXvmc3T0FF6dOn49593H0lIPGBEkwchKqax11214gvwLk7k9BVEAXEIKda
IorShYICToKP8UEcjFdzjYnRzJhxdmYy0eid3Nm7e/f9/Gv2nKruprurTtU5VdVNazgzH6OAbXf9
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
'''

BLACK_SPACE = ('#9191FF', '#000000', '#2323FF')  # (0.6, 240), BK, (0.2, 240)
WHITE_SPACE = ('#C0C000', '#FFFFFF', '#F7F700')  # (0.7,  60), WH, (0.9,  60)

COLOR_SPACE = (
    # Y=   0.3        0.5        0.7        0.9
    ('#FF1C1C', '#FF5D5D', '#FF9E9E', '#FFDFDF'),  # 000 : comment
    ('#DE2900', '#FF603C', '#FFA08A', '#FFDFD8'),  # 010 : fold
    ('#A63A00', '#FF6512', '#FFA271', '#FFE0D0'),  # 020 : del
    ('#864300', '#E07000', '#FFA64D', '#FFE1C4'),  # 030 : sect1, hnumb
    ('#714900', '#BC7A00', '#FFAC10', '#FFE3AF'),  # 040 : sect2
    ('#604E00', '#A08300', '#E0B700', '#FFE882'),  # 050 : sect3
    ('#525200', '#898900', '#C0C000', '#F7F700'),  # 060 : sect4, 判断者
    ('#465600', '#758F00', '#A4C900', '#D5FF1A'),  # 070 : sect5
    ('#3A5A00', '#619500', '#88D100', '#C2FF50'),  # 080 : sect6, paren1
    ('#2F5D00', '#4E9B00', '#6DD900', '#B8FF70'),  # 090 : sect7, paren2
    ('#226100', '#38A200', '#4FE200', '#B0FF86'),  # 100 : sect8, paren3
    ('#136500', '#1FA900', '#2CED00', '#AAFF97'),  # 110 :
    ('#006B00', '#00B200', '#00FA00', '#A5FFA5'),  # 120 : fontdeco
    ('#006913', '#00AF20', '#00F52D', '#A1FFB2'),  # 130 :
    ('#006724', '#00AC3C', '#00F154', '#9DFFBF'),  # 140 :
    ('#006633', '#00AA55', '#00EE77', '#98FFCC'),  # 150 : length reviser
    ('#006441', '#00A76D', '#00EA99', '#94FFDA'),  # 160 :
    ('#006351', '#00A586', '#00E7BC', '#8EFFEA'),  # 170 :
    ('#006161', '#00A2A2', '#00E3E3', '#87FFFF'),  # 180 : algin, 申立人
    ('#005F75', '#009FC3', '#21D6FF', '#B5F1FF'),  # 190 : table
    ('#005D8E', '#009AED', '#59C5FF', '#C8ECFF'),  # 200 : (fsp), ins
    ('#0059B2', '#1F8FFF', '#79BCFF', '#D2E9FF'),  # 210 : chap1
    ('#0053EF', '#4385FF', '#8EB6FF', '#D9E7FF'),  # 220 : chap2, (tab)
    ('#1F48FF', '#5F7CFF', '#9FB1FF', '#DFE5FF'),  # 230 : chap3
    ('#3F3FFF', '#7676FF', '#ADADFF', '#E4E4FF'),  # 240 : chap4, (hsp)
    ('#5B36FF', '#8A70FF', '#B9A9FF', '#E8E2FF'),  # 250 : chap5
    ('#772EFF', '#9E6AFF', '#C5A5FF', '#ECE1FF'),  # 260 :
    ('#9226FF', '#B164FF', '#D0A2FF', '#EFE0FF'),  # 270 : br, pgbr, hline
    ('#B01DFF', '#C75DFF', '#DD9EFF', '#F4DFFF'),  # 280 :
    ('#D312FF', '#E056FF', '#EC9AFF', '#F9DDFF'),  # 290 :
    ('#FF05FF', '#FF4DFF', '#FF94FF', '#FFDBFF'),  # 300 : 相手方
    ('#FF0AD2', '#FF50DF', '#FF96EC', '#FFDCF9'),  # 310 :
    ('#FF0EAB', '#FF53C3', '#FF98DB', '#FFDDF3'),  # 320 :
    ('#FF1188', '#FF55AA', '#FF99CC', '#FFDDEE'),  # 330 : list, fnumb
    ('#FF1566', '#FF5892', '#FF9BBE', '#FFDEE9'),  # 340 :
    ('#FF1843', '#FF5A79', '#FF9CAE', '#FFDEE4'),  # 350 :
)

KEYWORDS = [
    ['(加害者' +
     '|被告|本訴被告|反訴原告|別訴原告|被控訴人|被上告人' +
     '|相手方' +
     '|被疑者|被告人|弁護人|対象弁護士|弁護士' +
     '|反訴' +
     '|弁護士会' +
     '|乙|戊|辛)',
     'magenta'],
    ['(被害者' +
     '|原告|本訴原告|反訴被告|別訴被告|控訴人|上告人' +
     '|申立人' +
     '|検察官|検察事務官|懲戒請求者' +
     '|本訴' +
     '|検察庁' +
     '|甲|丁|庚|癸)',
     'cyan'],
    ['(裁判官|審判官|調停官|調停委員|司法委員|専門委員|書記官|事務官|訴外' +
     '|別訴' +
     '|裁判所' +
     '|丙|己|壬)',
     'yellow']]

CONFIGURATION_SAMPLE = [
    '',
    '書題名: -',
    '文書式: 普通', '文書式: 契約', '文書式: 条文',
    '用紙サ: A3横', '用紙サ: A3縦', '用紙サ: A4横', '用紙サ: A4縦',
    '上余白: 3.5 cm',
    '下余白: 2.2 cm',
    '左余白: 3.0 cm',
    '右余白: 2.0 cm',
    '頭書き: ',
    '頁番号: 無', '頁番号: 有',
    '行番号: 無', '行番号: 有',
    '明朝体: Times New Roman / ＭＳ 明朝',
    'ゴシ体: = / ＭＳ ゴシック',
    '異字体: IPAmj明朝',
    '文字サ: 12 pt',
    '行間隔: 2.14 倍',
    '前余白: 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍',
    '後余白: 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍',
    '字間整: 無',
    '完成稿: 偽',
    '作成時: - USER',
    '更新時: - USER',
    '']

PARAGRAPH_SAMPLE = ['', '\t',
                    '<!-------q1--------q2--------q3------' +
                    '--q4--------q5--------q6--------q7-->',
                    '<!--コメント-->',
                    '# <!--タイトル-->', '## <!--第１-->', '### <!--１-->',
                    '#### <!--(1)-->', '##### <!--ア-->', '###### <!--(ｱ)-->',
                    '####### <!--ａ-->', '######## <!--(a)-->',
                    '1. <!--番号付き箇条書-->',
                    '  1. <!--番号付き箇条書-->',
                    '    1. <!--番号付き箇条書-->',
                    '      1. <!--番号付き箇条書-->',
                    '- <!--番号なし箇条書-->',
                    '  - <!--番号なし箇条書-->',
                    '    - <!--番号なし箇条書-->',
                    '      - <!--番号なし箇条書-->',
                    '        - <!--番号なし箇条書-->',
                    ': <!--左寄せ-->', ': <!--中寄せ--> :', '<!--右寄せ--> :',
                    '|<!--表のセル-->|<!--表のセル-->|',
                    '![<!--画像の名前-->](<!--画像のファイル名-->)',
                    '$ <!--第１編-->', '$$ <!--第１章-->', '$$$ <!--第１節-->',
                    '$$$$ <!--第１款-->', '$$$$$ <!--第１目-->',
                    '\\[<!--数式-->\\]', '<pgbr><!--改ページ-->',
                    '']

FONT_DECORATOR_SAMPLE = ['\t',
                         '*<!--斜体-->*',
                         '*<!--太字-->*',
                         '__<!--下線-->__',
                         '---<!--微-->---',
                         '--<!--小-->--',
                         '++<!--大-->++',
                         '+++<!--巨-->+++',
                         '^R^<!--字赤-->^R^',
                         '^Y^<!--字黄-->^Y^',
                         '^G^<!--字緑-->^G^',
                         '^C^<!--字シ-->^C^',
                         '^B^<!--字青-->^B^',
                         '^M^<!--字マ-->^M^',
                         '_R_<!--地赤-->_R_',
                         '_Y_<!--地黄-->_Y_',
                         '_G_<!--地緑-->_G_',
                         '_C_<!--地シ-->_C_',
                         '_B_<!--地青-->_B_',
                         '_M_<!--地マ-->_M_',
                         '@游明朝@<!--游明朝-->@游明朝@',
                         '', '\t']

# 平成22年内閣告示第2号
JOYOKANJI = (
    ('0001', '亜亞', 'ア', ''),
    ('0002', '哀', 'アイ、あわ-れ、あわ-れむ', ''),
    ('0003', '挨', 'アイ', ''),
    ('0004', '愛', 'アイ', '愛媛'),
    ('0005', '曖', 'アイ', ''),
    ('0006', '悪惡', 'アク、オ、わる-い', ''),
    ('0007', '握', 'アク、にぎ-る', ''),
    ('0008', '圧壓', 'アツ', ''),
    ('0009', '扱', 'あつか-う', ''),
    ('0010', '宛', 'あ-てる', ''),
    ('0011', '嵐', 'あらし', ''),
    ('0012', '安', 'アン、やす-い', ''),
    ('0013', '案', 'アン', ''),
    ('0014', '暗', 'アン、くら-い', ''),
    ('0015', '以', 'イ', ''),
    ('0016', '衣', 'イ、ころも', '浴衣'),
    ('0017', '位', 'イ、くらい', '三位一体、従三位'),
    ('0018', '囲圍', 'イ、かこ-む、かこ-う', ''),
    ('0019', '医醫', 'イ', ''),
    ('0020', '依', 'イ、（エ）', ''),
    ('0021', '委', 'イ、ゆだ-ねる', ''),
    ('0022', '威', 'イ', ''),
    ('0023', '為爲', 'イ', '為替'),
    ('0024', '畏', 'イ、おそ-れる', ''),
    ('0025', '胃', 'イ', ''),
    ('0026', '尉', 'イ', ''),
    ('0027', '異', 'イ、こと', ''),
    ('0028', '移', 'イ、うつ-る、うつ-す', ''),
    ('0029', '萎', 'イ、な-える', ''),
    ('0030', '偉', 'イ、えら-い', ''),
    ('0031', '椅', 'イ', ''),
    ('0032', '彙', 'イ', ''),
    ('0033', '意', 'イ', '意気地'),
    ('0034', '違', 'イ、ちが-う、ちが-える', ''),
    ('0035', '維', 'イ', ''),
    ('0036', '慰', 'イ、なぐさ-める、なぐさ-む', ''),
    ('0037', '遺', 'イ、（ユイ）', ''),
    ('0038', '緯', 'イ', ''),
    ('0039', '域', 'イキ', ''),
    ('0040', '育', 'イク、そだ-つ、そだ-てる、はぐく-む', ''),
    ('0041', '一', 'イチ、イツ、ひと、ひと-つ', '一日、一人'),
    ('0042', '壱壹', 'イチ', ''),
    ('0043', '逸逸', 'イツ', ''),
    ('0044', '茨', '（いばら）', '茨城'),
    ('0045', '芋', 'いも', ''),
    ('0046', '引', 'イン、ひ-く、ひ-ける', ''),
    ('0047', '印', 'イン、しるし', ''),
    ('0048', '因', 'イン、よ-る', ''),
    ('0049', '咽', 'イン', ''),
    ('0050', '姻', 'イン', ''),
    ('0051', '員', 'イン', ''),
    ('0052', '院', 'イン', ''),
    ('0053', '淫', 'イン、みだ-ら', ''),
    ('0054', '陰', 'イン、かげ、かげ-る', ''),
    ('0055', '飲飮', 'イン、の-む', ''),
    ('0056', '隠隱', 'イン、かく-す、かく-れる', ''),
    ('0057', '韻', 'イン', ''),
    ('0058', '右', 'ウ、ユウ、みぎ', ''),
    ('0059', '宇', 'ウ', ''),
    ('0060', '羽羽', 'ウ、は、はね', ''),
    ('0061', '雨', 'ウ、あめ、（あま）', '五月雨、時雨、梅雨、春雨、小雨、霧雨'),
    ('0062', '唄', '（うた）', ''),
    ('0063', '鬱', 'ウツ', ''),
    ('0064', '畝', 'うね', ''),
    ('0065', '浦', 'うら', ''),
    ('0066', '運', 'ウン、はこ-ぶ', ''),
    ('0067', '雲', 'ウン、くも', ''),
    ('0068', '永', 'エイ、なが-い', ''),
    ('0069', '泳', 'エイ、およ-ぐ', ''),
    ('0070', '英', 'エイ', ''),
    ('0071', '映', 'エイ、うつ-る、うつ-す、は-える', ''),
    ('0072', '栄榮', 'エイ、さか-える、は-え、は-える', ''),
    ('0073', '営營', 'エイ、いとな-む', ''),
    ('0074', '詠', 'エイ、よ-む', ''),
    ('0075', '影', 'エイ、かげ', ''),
    ('0076', '鋭銳', 'エイ、するど-い', ''),
    ('0077', '衛衞', 'エイ', ''),
    ('0078', '易', 'エキ、イ、やさ-しい', ''),
    ('0079', '疫', 'エキ、（ヤク）', ''),
    ('0080', '益益', 'エキ、（ヤク）', ''),
    ('0081', '液', 'エキ', ''),
    ('0082', '駅驛', 'エキ', ''),
    ('0083', '悦悅', 'エツ', ''),
    ('0084', '越', 'エツ、こ-す、こ-える', ''),
    ('0085', '謁謁', 'エツ', ''),
    ('0086', '閲閱', 'エツ', ''),
    ('0087', '円圓', 'エン、まる-い', ''),
    ('0088', '延', 'エン、の-びる、の-べる、の-ばす', ''),
    ('0089', '沿', 'エン、そ-う', ''),
    ('0090', '炎', 'エン、ほのお', ''),
    ('0091', '怨', 'エン、オン', ''),
    ('0092', '宴', 'エン', ''),
    ('0093', '媛', 'エン', '愛媛'),
    ('0094', '援', 'エン', ''),
    ('0095', '園', 'エン、その', ''),
    ('0096', '煙', 'エン、けむ-る、けむり、けむ-い', ''),
    ('0097', '猿', 'エン、さる', ''),
    ('0098', '遠', 'エン、（オン）、とお-い', ''),
    ('0099', '鉛', 'エン、なまり', ''),
    ('0100', '塩鹽', 'エン、しお', ''),
    ('0101', '演', 'エン', ''),
    ('0102', '縁緣', 'エン、ふち', '因縁'),
    ('0103', '艶艷', 'エン、つや', ''),
    ('0104', '汚', 'オ、けが-す、けが-れる、けが-らわしい、よご-す、よご-れる、きたな-い', ''),
    ('0105', '王', 'オウ', '親王、勤王'),
    ('0106', '凹', 'オウ', '凸凹'),
    ('0107', '央', 'オウ', ''),
    ('0108', '応應', 'オウ、こた-える', '反応、順応'),
    ('0109', '往', 'オウ', ''),
    ('0110', '押', 'オウ、お-す、お-さえる', ''),
    ('0111', '旺', 'オウ', ''),
    ('0112', '欧歐', 'オウ', ''),
    ('0113', '殴毆', 'オウ、なぐ-る', ''),
    ('0114', '桜櫻', 'オウ、さくら', ''),
    ('0115', '翁', 'オウ', ''),
    ('0116', '奥奧', 'オウ、おく', ''),
    ('0117', '横橫', 'オウ、よこ', ''),
    ('0118', '岡', '（おか）', ''),
    ('0119', '屋', 'オク、や', '母屋、数寄屋、数奇屋、部屋、八百屋、紺屋'),
    ('0120', '億', 'オク', ''),
    ('0121', '憶', 'オク', ''),
    ('0122', '臆', 'オク', ''),
    ('0123', '虞', 'おそれ', ''),
    ('0124', '乙', 'オツ', '乙女、早乙女'),
    ('0125', '俺', 'おれ', ''),
    ('0126', '卸', 'おろ-す、おろし', ''),
    ('0127', '音', 'オン、イン、おと、ね', '観音'),
    ('0128', '恩', 'オン', ''),
    ('0129', '温溫', 'オン、あたた-か、あたた-かい、あたた-まる、あたた-める', ''),
    ('0130', '穏穩', 'オン、おだ-やか', '安穏'),
    ('0131', '下', 'カ、ゲ、した、しも、もと、さ-げる、さ-がる、くだ-る、くだ-す、くだ-さる、お-ろす、お-りる', '下手'),
    ('0132', '化', 'カ、ケ、ば-ける、ば-かす', ''),
    ('0133', '火', 'カ、ひ、（ほ）', ''),
    ('0134', '加', 'カ、くわ-える、くわ-わる', ''),
    ('0135', '可', 'カ', ''),
    ('0136', '仮假', 'カ、（ケ）、かり', '仮名'),
    ('0137', '何', 'カ、なに、（なん）', ''),
    ('0138', '花', 'カ、はな', ''),
    ('0139', '佳', 'カ', ''),
    ('0140', '価價', 'カ、あたい', ''),
    ('0141', '果', 'カ、は-たす、は-てる、は-て', '果物'),
    ('0142', '河', 'カ、かわ', '河岸、河原'),
    ('0143', '苛', 'カ', ''),
    ('0144', '科', 'カ', ''),
    ('0145', '架', 'カ、か-ける、か-かる', ''),
    ('0146', '夏', 'カ、（ゲ）、なつ', ''),
    ('0147', '家', 'カ、ケ、いえ、や', '母家'),
    ('0148', '荷', 'カ、に', ''),
    ('0149', '華', 'カ、（ケ）、はな', ''),
    ('0150', '菓', 'カ', ''),
    ('0151', '貨', 'カ', ''),
    ('0152', '渦', 'カ、うず', ''),
    ('0153', '過', 'カ、す-ぎる、す-ごす、あやま-つ、あやま-ち', ''),
    ('0154', '嫁', 'カ、よめ、とつ-ぐ', ''),
    ('0155', '暇', 'カ、ひま', ''),
    ('0156', '禍禍', 'カ', ''),
    ('0157', '靴', 'カ、くつ', ''),
    ('0158', '寡', 'カ', ''),
    ('0159', '歌', 'カ、うた、うた-う', '詩歌'),
    ('0160', '箇', 'カ', ''),
    ('0161', '稼', 'カ、かせ-ぐ', ''),
    ('0162', '課', 'カ', ''),
    ('0163', '蚊', 'か', '蚊帳'),
    ('0164', '牙', 'ガ、（ゲ）、きば', ''),
    ('0165', '瓦', 'ガ、かわら', ''),
    ('0166', '我', 'ガ、われ、わ', ''),
    ('0167', '画畫', 'ガ、カク', ''),
    ('0168', '芽', 'ガ、め', ''),
    ('0169', '賀', 'ガ', '滋賀'),
    ('0170', '雅', 'ガ', ''),
    ('0171', '餓', 'ガ', ''),
    ('0172', '介', 'カイ', ''),
    ('0173', '回', 'カイ、（エ）、まわ-る、まわ-す', ''),
    ('0174', '灰', 'カイ、はい', ''),
    ('0175', '会會', 'カイ、エ、あ-う', ''),
    ('0176', '快', 'カイ、こころよ-い', ''),
    ('0177', '戒', 'カイ、いまし-める', ''),
    ('0178', '改', 'カイ、あらた-める、あらた-まる', ''),
    ('0179', '怪', 'カイ、あや-しい、あや-しむ', ''),
    ('0180', '拐', 'カイ', ''),
    ('0181', '悔悔', 'カイ、く-いる、く-やむ、くや-しい', ''),
    ('0182', '海海', 'カイ、うみ', '海女、海士、海原'),
    ('0183', '界', 'カイ', ''),
    ('0184', '皆', 'カイ、みな', ''),
    ('0185', '械', 'カイ', ''),
    ('0186', '絵繪', 'カイ、エ', ''),
    ('0187', '開', 'カイ、ひら-く、ひら-ける、あ-く、あ-ける', ''),
    ('0188', '階', 'カイ', ''),
    ('0189', '塊', 'カイ、かたまり', ''),
    ('0190', '楷', 'カイ', ''),
    ('0191', '解', 'カイ、ゲ、と-く、と-かす、と-ける', ''),
    ('0192', '潰', 'カイ、つぶ-す、つぶ-れる', ''),
    ('0193', '壊壞', 'カイ、こわ-す、こわ-れる', ''),
    ('0194', '懐懷', 'カイ、ふところ、なつ-かしい、なつ-かしむ、なつ-く、なつ-ける', ''),
    ('0195', '諧', 'カイ', ''),
    ('0196', '貝', 'かい', ''),
    ('0197', '外', 'ガイ、ゲ、そと、ほか、はず-す、はず-れる', ''),
    ('0198', '劾', 'ガイ', ''),
    ('0199', '害', 'ガイ', ''),
    ('0200', '崖', 'ガイ、がけ', ''),
    ('0201', '涯', 'ガイ', ''),
    ('0202', '街', 'ガイ、（カイ）、まち', ''),
    ('0203', '慨慨', 'ガイ', ''),
    ('0204', '蓋', 'ガイ、ふた', ''),
    ('0205', '該', 'ガイ', ''),
    ('0206', '概槪', 'ガイ', ''),
    ('0207', '骸', 'ガイ', ''),
    ('0208', '垣', 'かき', ''),
    ('0209', '柿', 'かき', ''),
    ('0210', '各', 'カク、おのおの', ''),
    ('0211', '角', 'カク、かど、つの', ''),
    ('0212', '拡擴', 'カク', ''),
    ('0213', '革', 'カク、かわ', ''),
    ('0214', '格', 'カク、（コウ）', ''),
    ('0215', '核', 'カク', ''),
    ('0216', '殻殼', 'カク、から', ''),
    ('0217', '郭', 'カク', ''),
    ('0218', '覚覺', 'カク、おぼ-える、さ-ます、さ-める', ''),
    ('0219', '較', 'カク', ''),
    ('0220', '隔', 'カク、へだ-てる、へだ-たる', ''),
    ('0221', '閣', 'カク', ''),
    ('0222', '確', 'カク、たし-か、たし-かめる', ''),
    ('0223', '獲', 'カク、え-る', ''),
    ('0224', '嚇', 'カク', ''),
    ('0225', '穫', 'カク', ''),
    ('0226', '学學', 'ガク、まな-ぶ', ''),
    ('0227', '岳嶽', 'ガク、たけ', ''),
    ('0228', '楽樂', 'ガク、ラク、たの-しい、たの-しむ', '神楽'),
    ('0229', '額', 'ガク、ひたい', ''),
    ('0230', '顎', 'ガク、あご', ''),
    ('0231', '掛', 'か-ける、か-かる、かかり', ''),
    ('0232', '潟', 'かた', ''),
    ('0233', '括', 'カツ', ''),
    ('0234', '活', 'カツ', ''),
    ('0235', '喝喝', 'カツ', ''),
    ('0236', '渇渴', 'カツ、かわ-く', ''),
    ('0237', '割', 'カツ、わ-る、わり、わ-れる、さ-く', ''),
    ('0238', '葛', 'カツ、くず', ''),
    ('0239', '滑', 'カツ、コツ、すべ-る、なめ-らか', ''),
    ('0240', '褐褐', 'カツ', ''),
    ('0241', '轄', 'カツ', ''),
    ('0242', '且', 'か-つ', ''),
    ('0243', '株', 'かぶ', ''),
    ('0244', '釜', 'かま', ''),
    ('0245', '鎌', 'かま', ''),
    ('0246', '刈', 'か-る', ''),
    ('0247', '干', 'カン、ほ-す、ひ-る', ''),
    ('0248', '刊', 'カン', ''),
    ('0249', '甘', 'カン、あま-い、あま-える、あま-やかす', ''),
    ('0250', '汗', 'カン、あせ', ''),
    ('0251', '缶罐', 'カン', ''),
    ('0252', '完', 'カン', ''),
    ('0253', '肝', 'カン、きも', ''),
    ('0254', '官', 'カン', ''),
    ('0255', '冠', 'カン、かんむり', ''),
    ('0256', '巻卷', 'カン、ま-く、まき', ''),
    ('0257', '看', 'カン', ''),
    ('0258', '陥陷', 'カン、おちい-る、おとしい-れる', ''),
    ('0259', '乾', 'カン、かわ-く、かわ-かす', ''),
    ('0260', '勘', 'カン', ''),
    ('0261', '患', 'カン、わずら-う', ''),
    ('0262', '貫', 'カン、つらぬ-く', ''),
    ('0263', '寒', 'カン、さむ-い', ''),
    ('0264', '喚', 'カン', ''),
    ('0265', '堪', 'カン、た-える', '堪能'),
    ('0266', '換', 'カン、か-える、か-わる', ''),
    ('0267', '敢', 'カン', ''),
    ('0268', '棺', 'カン', ''),
    ('0269', '款', 'カン', ''),
    ('0270', '間', 'カン、ケン、あいだ、ま', ''),
    ('0271', '閑', 'カン', ''),
    ('0272', '勧勸', 'カン、すす-める', ''),
    ('0273', '寛寬', 'カン', ''),
    ('0274', '幹', 'カン、みき', ''),
    ('0275', '感', 'カン', ''),
    ('0276', '漢漢', 'カン', ''),
    ('0277', '慣', 'カン、な-れる、な-らす', ''),
    ('0278', '管', 'カン、くだ', ''),
    ('0279', '関關', 'カン、せき、かか-わる', ''),
    ('0280', '歓歡', 'カン', ''),
    ('0281', '監', 'カン', ''),
    ('0282', '緩', 'カン、ゆる-い、ゆる-やか、ゆる-む、ゆる-める', ''),
    ('0283', '憾', 'カン', ''),
    ('0284', '還', 'カン', ''),
    ('0285', '館館', 'カン、やかた', ''),
    ('0286', '環', 'カン', ''),
    ('0287', '簡', 'カン', ''),
    ('0288', '観觀', 'カン', ''),
    ('0289', '韓', 'カン', ''),
    ('0290', '艦', 'カン', ''),
    ('0291', '鑑', 'カン、かんが-みる', ''),
    ('0292', '丸', 'ガン、まる、まる-い、まる-める', ''),
    ('0293', '含', 'ガン、ふく-む、ふく-める', ''),
    ('0294', '岸', 'ガン、きし', '河岸'),
    ('0295', '岩', 'ガン、いわ', ''),
    ('0296', '玩', 'ガン', ''),
    ('0297', '眼', 'ガン、（ゲン）、まなこ', '眼鏡'),
    ('0298', '頑', 'ガン', ''),
    ('0299', '顔顏', 'ガン、かお', '笑顔'),
    ('0300', '願', 'ガン、ねが-う', ''),
    ('0301', '企', 'キ、くわだ-てる', ''),
    ('0302', '伎', 'キ', ''),
    ('0303', '危', 'キ、あぶ-ない、あや-うい、あや-ぶむ', ''),
    ('0304', '机', 'キ、つくえ', ''),
    ('0305', '気氣', 'キ、ケ', '意気地、浮気'),
    ('0306', '岐', 'キ', '岐阜'),
    ('0307', '希', 'キ', ''),
    ('0308', '忌', 'キ、い-む、い-まわしい', ''),
    ('0309', '汽', 'キ', ''),
    ('0310', '奇', 'キ', '数奇屋'),
    ('0311', '祈祈', 'キ、いの-る', ''),
    ('0312', '季', 'キ', ''),
    ('0313', '紀', 'キ', ''),
    ('0314', '軌', 'キ', ''),
    ('0315', '既旣', 'キ、すで-に', ''),
    ('0316', '記', 'キ、しる-す', ''),
    ('0317', '起', 'キ、お-きる、お-こる、お-こす', ''),
    ('0318', '飢', 'キ、う-える', ''),
    ('0319', '鬼', 'キ、おに', ''),
    ('0320', '帰歸', 'キ、かえ-る、かえ-す', ''),
    ('0321', '基', 'キ、もと、もとい', ''),
    ('0322', '寄', 'キ、よ-る、よ-せる', '数寄屋、最寄り、寄席'),
    ('0323', '規', 'キ', ''),
    ('0324', '亀龜', 'キ、かめ', ''),
    ('0325', '喜', 'キ、よろこ-ぶ', ''),
    ('0326', '幾', 'キ、いく', ''),
    ('0327', '揮', 'キ', ''),
    ('0328', '期', 'キ、（ゴ）', ''),
    ('0329', '棋', 'キ', ''),
    ('0330', '貴', 'キ、たっと-い、とうと-い、たっと-ぶ、とうと-ぶ', '富貴'),
    ('0331', '棄', 'キ', ''),
    ('0332', '毀', 'キ', ''),
    ('0333', '旗', 'キ、はた', ''),
    ('0334', '器器', 'キ、うつわ', ''),
    ('0335', '畿', 'キ', ''),
    ('0336', '輝', 'キ、かがや-く', ''),
    ('0337', '機', 'キ、はた', ''),
    ('0338', '騎', 'キ', ''),
    ('0339', '技', 'ギ、わざ', ''),
    ('0340', '宜', 'ギ', ''),
    ('0341', '偽僞', 'ギ、いつわ-る、にせ', ''),
    ('0342', '欺', 'ギ、あざむ-く', ''),
    ('0343', '義', 'ギ', ''),
    ('0344', '疑', 'ギ、うたが-う', ''),
    ('0345', '儀', 'ギ', ''),
    ('0346', '戯戲', 'ギ、たわむ-れる', ''),
    ('0347', '擬', 'ギ', ''),
    ('0348', '犠犧', 'ギ', ''),
    ('0349', '議', 'ギ', ''),
    ('0350', '菊', 'キク', ''),
    ('0351', '吉', 'キチ、キツ', ''),
    ('0352', '喫', 'キツ', ''),
    ('0353', '詰', 'キツ、つ-める、つ-まる、つ-む', ''),
    ('0354', '却', 'キャク', ''),
    ('0355', '客', 'キャク、カク', ''),
    ('0356', '脚', 'キャク、（キャ）、あし', ''),
    ('0357', '逆', 'ギャク、さか、さか-らう', ''),
    ('0358', '虐', 'ギャク、しいた-げる', ''),
    ('0359', '九', 'キュウ、ク、ここの、ここの-つ', ''),
    ('0360', '久', 'キュウ、（ク）、ひさ-しい', ''),
    ('0361', '及', 'キュウ、およ-ぶ、およ-び、およ-ぼす', ''),
    ('0362', '弓', 'キュウ、ゆみ', ''),
    ('0363', '丘', 'キュウ、おか', ''),
    ('0364', '旧舊', 'キュウ', ''),
    ('0365', '休', 'キュウ、やす-む、やす-まる、やす-める', ''),
    ('0366', '吸', 'キュウ、す-う', ''),
    ('0367', '朽', 'キュウ、く-ちる', ''),
    ('0368', '臼', 'キュウ、うす', ''),
    ('0369', '求', 'キュウ、もと-める', ''),
    ('0370', '究', 'キュウ、きわ-める', ''),
    ('0371', '泣', 'キュウ、な-く', ''),
    ('0372', '急', 'キュウ、いそ-ぐ', ''),
    ('0373', '級', 'キュウ', ''),
    ('0374', '糾', 'キュウ', ''),
    ('0375', '宮', 'キュウ、グウ、（ク）、みや', '宮城、宮内庁'),
    ('0376', '救', 'キュウ、すく-う', ''),
    ('0377', '球', 'キュウ、たま', ''),
    ('0378', '給', 'キュウ', ''),
    ('0379', '嗅', 'キュウ、か-ぐ', ''),
    ('0380', '窮', 'キュウ、きわ-める、きわ-まる', ''),
    ('0381', '牛', 'ギュウ、うし', ''),
    ('0382', '去', 'キョ、コ、さ-る', ''),
    ('0383', '巨', 'キョ', ''),
    ('0384', '居', 'キョ、い-る', '居士'),
    ('0385', '拒', 'キョ、こば-む', ''),
    ('0386', '拠據', 'キョ、コ', ''),
    ('0387', '挙擧', 'キョ、あ-げる、あ-がる', ''),
    ('0388', '虚虛', 'キョ、（コ）', ''),
    ('0389', '許', 'キョ、ゆる-す', ''),
    ('0390', '距', 'キョ', ''),
    ('0391', '魚', 'ギョ、うお、さかな', '雑魚'),
    ('0392', '御', 'ギョ、ゴ、おん', ''),
    ('0393', '漁', 'ギョ、リョウ', ''),
    ('0394', '凶', 'キョウ', ''),
    ('0395', '共', 'キョウ、とも', ''),
    ('0396', '叫', 'キョウ、さけ-ぶ', ''),
    ('0397', '狂', 'キョウ、くる-う、くる-おしい', ''),
    ('0398', '京', 'キョウ、ケイ', '京浜、京阪'),
    ('0399', '享', 'キョウ', ''),
    ('0400', '供', 'キョウ、（ク）、そな-える、とも', ''),
    ('0401', '協', 'キョウ', ''),
    ('0402', '況', 'キョウ', ''),
    ('0403', '峡峽', 'キョウ', ''),
    ('0404', '挟挾', 'キョウ、はさ-む、はさ-まる', ''),
    ('0405', '狭狹', 'キョウ、せま-い、せば-める、せば-まる', ''),
    ('0406', '恐', 'キョウ、おそ-れる、おそ-ろしい', ''),
    ('0407', '恭', 'キョウ、うやうや-しい', ''),
    ('0408', '胸', 'キョウ、むね、（むな）', ''),
    ('0409', '脅', 'キョウ、おびや-かす、おど-す、おど-かす', ''),
    ('0410', '強', 'キョウ、ゴウ、つよ-い、つよ-まる、つよ-める、し-いる', ''),
    ('0411', '教敎', 'キョウ、おし-える、おそ-わる', ''),
    ('0412', '郷鄕', 'キョウ、ゴウ', ''),
    ('0413', '境', 'キョウ、（ケイ）、さかい', ''),
    ('0414', '橋', 'キョウ、はし', ''),
    ('0415', '矯', 'キョウ、た-める', ''),
    ('0416', '鏡', 'キョウ、かがみ', '眼鏡'),
    ('0417', '競', 'キョウ、ケイ、きそ-う、せ-る', ''),
    ('0418', '響響', 'キョウ、ひび-く', ''),
    ('0419', '驚', 'キョウ、おどろ-く、おどろ-かす', ''),
    ('0420', '仰', 'ギョウ、（コウ）、あお-ぐ、おお-せ', ''),
    ('0421', '暁曉', 'ギョウ、あかつき', ''),
    ('0422', '業', 'ギョウ、ゴウ、わざ', ''),
    ('0423', '凝', 'ギョウ、こ-る、こ-らす', ''),
    ('0424', '曲', 'キョク、ま-がる、ま-げる', ''),
    ('0425', '局', 'キョク', ''),
    ('0426', '極', 'キョク、ゴク、きわ-める、きわ-まる、きわ-み', ''),
    ('0427', '玉', 'ギョク、たま', ''),
    ('0428', '巾', 'キン', ''),
    ('0429', '斤', 'キン', ''),
    ('0430', '均', 'キン', ''),
    ('0431', '近', 'キン、ちか-い', ''),
    ('0432', '金', 'キン、コン、かね、（かな）', ''),
    ('0433', '菌', 'キン', ''),
    ('0434', '勤勤', 'キン、（ゴン）、つと-める、つと-まる', ''),
    ('0435', '琴', 'キン、こと', ''),
    ('0436', '筋', 'キン、すじ', ''),
    ('0437', '僅', 'キン、わず-か', ''),
    ('0438', '禁', 'キン', ''),
    ('0439', '緊', 'キン', ''),
    ('0440', '錦', 'キン、にしき', ''),
    ('0441', '謹謹', 'キン、つつし-む', ''),
    ('0442', '襟', 'キン、えり', ''),
    ('0443', '吟', 'ギン', ''),
    ('0444', '銀', 'ギン', ''),
    ('0445', '区區', 'ク', ''),
    ('0446', '句', 'ク', ''),
    ('0447', '苦', 'ク、くる-しい、くる-しむ、くる-しめる、にが-い、にが-る', ''),
    ('0448', '駆驅', 'ク、か-ける、か-る', ''),
    ('0449', '具', 'グ', ''),
    ('0450', '惧', 'グ', ''),
    ('0451', '愚', 'グ、おろ-か', ''),
    ('0452', '空', 'クウ、そら、あ-く、あ-ける、から', ''),
    ('0453', '偶', 'グウ', ''),
    ('0454', '遇', 'グウ', ''),
    ('0455', '隅', 'グウ、すみ', ''),
    ('0456', '串', 'くし', ''),
    ('0457', '屈', 'クツ', ''),
    ('0458', '掘', 'クツ、ほ-る', ''),
    ('0459', '窟', 'クツ', ''),
    ('0460', '熊', 'くま', ''),
    ('0461', '繰', 'く-る', ''),
    ('0462', '君', 'クン、きみ', ''),
    ('0463', '訓', 'クン', ''),
    ('0464', '勲勳', 'クン', ''),
    ('0465', '薫薰', 'クン、かお-る', ''),
    ('0466', '軍', 'グン', ''),
    ('0467', '郡', 'グン', ''),
    ('0468', '群', 'グン、む-れる、む-れ、（むら）', ''),
    ('0469', '兄', 'ケイ、（キョウ）、あに', '兄さん'),
    ('0470', '刑', 'ケイ', ''),
    ('0471', '形', 'ケイ、ギョウ、かた、かたち', ''),
    ('0472', '系', 'ケイ', ''),
    ('0473', '径徑', 'ケイ', ''),
    ('0474', '茎莖', 'ケイ、くき', ''),
    ('0475', '係', 'ケイ、かか-る、かかり', ''),
    ('0476', '型', 'ケイ、かた', ''),
    ('0477', '契', 'ケイ、ちぎ-る', ''),
    ('0478', '計', 'ケイ、はか-る、はか-らう', '時計'),
    ('0479', '恵惠', 'ケイ、エ、めぐ-む', ''),
    ('0480', '啓', 'ケイ', ''),
    ('0481', '掲揭', 'ケイ、かか-げる', ''),
    ('0482', '渓溪', 'ケイ', ''),
    ('0483', '経經', 'ケイ、キョウ、へ-る', '読経'),
    ('0484', '蛍螢', 'ケイ、ほたる', ''),
    ('0485', '敬', 'ケイ、うやま-う', ''),
    ('0486', '景', 'ケイ', '景色'),
    ('0487', '軽輕', 'ケイ、かる-い、かろ-やか', ''),
    ('0488', '傾', 'ケイ、かたむ-く、かたむ-ける', ''),
    ('0489', '携', 'ケイ、たずさ-える、たずさ-わる', ''),
    ('0490', '継繼', 'ケイ、つ-ぐ', ''),
    ('0491', '詣', 'ケイ、もう-でる', ''),
    ('0492', '慶', 'ケイ', ''),
    ('0493', '憬', 'ケイ', '憧憬'),
    ('0494', '稽', 'ケイ', ''),
    ('0495', '憩', 'ケイ、いこ-い、いこ-う', ''),
    ('0496', '警', 'ケイ', ''),
    ('0497', '鶏鷄', 'ケイ、にわとり', ''),
    ('0498', '芸藝', 'ゲイ', ''),
    ('0499', '迎', 'ゲイ、むか-える', ''),
    ('0500', '鯨', 'ゲイ、くじら', ''),
    ('0501', '隙', 'ゲキ、すき', ''),
    ('0502', '劇', 'ゲキ', ''),
    ('0503', '撃擊', 'ゲキ、う-つ', ''),
    ('0504', '激', 'ゲキ、はげ-しい', ''),
    ('0505', '桁', 'けた', ''),
    ('0506', '欠缺', 'ケツ、か-ける、か-く', ''),
    ('0507', '穴', 'ケツ、あな', ''),
    ('0508', '血', 'ケツ、ち', ''),
    ('0509', '決', 'ケツ、き-める、き-まる', ''),
    ('0510', '結', 'ケツ、むす-ぶ、ゆ-う、ゆ-わえる', ''),
    ('0511', '傑', 'ケツ', ''),
    ('0512', '潔', 'ケツ、いさぎよ-い', ''),
    ('0513', '月', 'ゲツ、ガツ、つき', '五月、五月雨'),
    ('0514', '犬', 'ケン、いぬ', ''),
    ('0515', '件', 'ケン', ''),
    ('0516', '見', 'ケン、み-る、み-える、み-せる', ''),
    ('0517', '券', 'ケン', ''),
    ('0518', '肩', 'ケン、かた', ''),
    ('0519', '建', 'ケン、（コン）、た-てる、た-つ', ''),
    ('0520', '研硏', 'ケン、と-ぐ', ''),
    ('0521', '県縣', 'ケン', ''),
    ('0522', '倹儉', 'ケン', ''),
    ('0523', '兼', 'ケン、か-ねる', ''),
    ('0524', '剣劍', 'ケン、つるぎ', ''),
    ('0525', '拳', 'ケン、こぶし', ''),
    ('0526', '軒', 'ケン、のき', ''),
    ('0527', '健', 'ケン、すこ-やか', ''),
    ('0528', '険險', 'ケン、けわ-しい', ''),
    ('0529', '圏圈', 'ケン', ''),
    ('0530', '堅', 'ケン、かた-い', ''),
    ('0531', '検檢', 'ケン', ''),
    ('0532', '嫌', 'ケン、（ゲン）、きら-う、いや', ''),
    ('0533', '献獻', 'ケン、（コン）', ''),
    ('0534', '絹', 'ケン、きぬ', ''),
    ('0535', '遣', 'ケン、つか-う、つか-わす', ''),
    ('0536', '権權', 'ケン、（ゴン）', ''),
    ('0537', '憲', 'ケン', ''),
    ('0538', '賢', 'ケン、かしこ-い', ''),
    ('0539', '謙', 'ケン', ''),
    ('0540', '鍵', 'ケン、かぎ', ''),
    ('0541', '繭', 'ケン、まゆ', ''),
    ('0542', '顕顯', 'ケン', ''),
    ('0543', '験驗', 'ケン、（ゲン）', ''),
    ('0544', '懸', 'ケン、（ケ）、か-ける、か-かる', ''),
    ('0545', '元', 'ゲン、ガン、もと', ''),
    ('0546', '幻', 'ゲン、まぼろし', ''),
    ('0547', '玄', 'ゲン', '玄人'),
    ('0548', '言', 'ゲン、ゴン、い-う、こと', ''),
    ('0549', '弦', 'ゲン、つる', ''),
    ('0550', '限', 'ゲン、かぎ-る', ''),
    ('0551', '原', 'ゲン、はら', '海原、河原、川原'),
    ('0552', '現', 'ゲン、あらわ-れる、あらわ-す', ''),
    ('0553', '舷', 'ゲン', ''),
    ('0554', '減', 'ゲン、へ-る、へ-らす', ''),
    ('0555', '源', 'ゲン、みなもと', ''),
    ('0556', '厳嚴', 'ゲン、（ゴン）、おごそ-か、きび-しい', ''),
    ('0557', '己', 'コ、キ、おのれ', ''),
    ('0558', '戸戶', 'コ、と', ''),
    ('0559', '古', 'コ、ふる-い、ふる-す', ''),
    ('0560', '呼', 'コ、よ-ぶ', ''),
    ('0561', '固', 'コ、かた-める、かた-まる、かた-い', '固唾'),
    ('0562', '股', 'コ、また', ''),
    ('0563', '虎', 'コ、とら', ''),
    ('0564', '孤', 'コ', ''),
    ('0565', '弧', 'コ', ''),
    ('0566', '故', 'コ、ゆえ', ''),
    ('0567', '枯', 'コ、か-れる、か-らす', ''),
    ('0568', '個', 'コ', ''),
    ('0569', '庫', 'コ、（ク）', ''),
    ('0570', '湖', 'コ、みずうみ', ''),
    ('0571', '雇', 'コ、やと-う', ''),
    ('0572', '誇', 'コ、ほこ-る', ''),
    ('0573', '鼓', 'コ、つづみ', ''),
    ('0574', '錮', 'コ', ''),
    ('0575', '顧', 'コ、かえり-みる', ''),
    ('0576', '五', 'ゴ、いつ、いつ-つ', '五月、五月雨'),
    ('0577', '互', 'ゴ、たが-い', ''),
    ('0578', '午', 'ゴ', ''),
    ('0579', '呉吳', 'ゴ', ''),
    ('0580', '後', 'ゴ、コウ、のち、うし-ろ、あと、おく-れる', ''),
    ('0581', '娯娛', 'ゴ', ''),
    ('0582', '悟', 'ゴ、さと-る', ''),
    ('0583', '碁', 'ゴ', ''),
    ('0584', '語', 'ゴ、かた-る、かた-らう', ''),
    ('0585', '誤', 'ゴ、あやま-る', ''),
    ('0586', '護', 'ゴ', ''),
    ('0587', '口', 'コウ、ク、くち', ''),
    ('0588', '工', 'コウ、ク', ''),
    ('0589', '公', 'コウ、おおやけ', ''),
    ('0590', '勾', 'コウ', ''),
    ('0591', '孔', 'コウ', ''),
    ('0592', '功', 'コウ、（ク）', ''),
    ('0593', '巧', 'コウ、たく-み', ''),
    ('0594', '広廣', 'コウ、ひろ-い、ひろ-まる、ひろ-める、ひろ-がる、ひろ-げる', ''),
    ('0595', '甲', 'コウ、カン', ''),
    ('0596', '交', 'コウ、まじ-わる、まじ-える、ま-じる、ま-ざる、ま-ぜる、か-う、か-わす', ''),
    ('0597', '光', 'コウ、ひか-る、ひかり', ''),
    ('0598', '向', 'コウ、む-く、む-ける、む-かう、む-こう', ''),
    ('0599', '后', 'コウ', ''),
    ('0600', '好', 'コウ、この-む、す-く', ''),
    ('0601', '江', 'コウ、え', ''),
    ('0602', '考', 'コウ、かんが-える', ''),
    ('0603', '行', 'コウ、ギョウ、（アン）、い-く、ゆ-く、おこな-う', '行方'),
    ('0604', '坑', 'コウ', ''),
    ('0605', '孝', 'コウ', ''),
    ('0606', '抗', 'コウ', ''),
    ('0607', '攻', 'コウ、せ-める', ''),
    ('0608', '更', 'コウ、さら、ふ-ける、ふ-かす', ''),
    ('0609', '効效', 'コウ、き-く', ''),
    ('0610', '幸', 'コウ、さいわ-い、さち、しあわ-せ', ''),
    ('0611', '拘', 'コウ', ''),
    ('0612', '肯', 'コウ', ''),
    ('0613', '侯', 'コウ', ''),
    ('0614', '厚', 'コウ、あつ-い', ''),
    ('0615', '恒恆', 'コウ', ''),
    ('0616', '洪', 'コウ', ''),
    ('0617', '皇', 'コウ、オウ', '天皇'),
    ('0618', '紅', 'コウ、（ク）、べに、くれない', '紅葉'),
    ('0619', '荒', 'コウ、あら-い、あ-れる、あ-らす', ''),
    ('0620', '郊', 'コウ', ''),
    ('0621', '香', 'コウ、（キョウ）、か、かお-り、かお-る', ''),
    ('0622', '候', 'コウ、そうろう', ''),
    ('0623', '校', 'コウ', ''),
    ('0624', '耕', 'コウ、たがや-す', ''),
    ('0625', '航', 'コウ', ''),
    ('0626', '貢', 'コウ、（ク）、みつ-ぐ', ''),
    ('0627', '降', 'コウ、お-りる、お-ろす、ふ-る', ''),
    ('0628', '高', 'コウ、たか-い、たか、たか-まる、たか-める', ''),
    ('0629', '康', 'コウ', ''),
    ('0630', '控', 'コウ、ひか-える', ''),
    ('0631', '梗', 'コウ', ''),
    ('0632', '黄黃', 'コウ、オウ、き、（こ）', '硫黄'),
    ('0633', '喉', 'コウ、のど', ''),
    ('0634', '慌', 'コウ、あわ-てる、あわ-ただしい', ''),
    ('0635', '港', 'コウ、みなと', ''),
    ('0636', '硬', 'コウ、かた-い', ''),
    ('0637', '絞', 'コウ、しぼ-る、し-める、し-まる', ''),
    ('0638', '項', 'コウ', ''),
    ('0639', '溝', 'コウ、みぞ', ''),
    ('0640', '鉱鑛', 'コウ', ''),
    ('0641', '構', 'コウ、かま-える、かま-う', ''),
    ('0642', '綱', 'コウ、つな', ''),
    ('0643', '酵', 'コウ', ''),
    ('0644', '稿', 'コウ', ''),
    ('0645', '興', 'コウ、キョウ、おこ-る、おこ-す', ''),
    ('0646', '衡', 'コウ', ''),
    ('0647', '鋼', 'コウ、はがね', ''),
    ('0648', '講', 'コウ', ''),
    ('0649', '購', 'コウ', ''),
    ('0650', '乞', 'こ-う', ''),
    ('0651', '号號', 'ゴウ', ''),
    ('0652', '合', 'ゴウ、ガッ、（カッ）、あ-う、あ-わす、あ-わせる', '合点'),
    ('0653', '拷', 'ゴウ', ''),
    ('0654', '剛', 'ゴウ', ''),
    ('0655', '傲', 'ゴウ', ''),
    ('0656', '豪', 'ゴウ', ''),
    ('0657', '克', 'コク', ''),
    ('0658', '告吿', 'コク、つ-げる', ''),
    ('0659', '谷', 'コク、たに', ''),
    ('0660', '刻', 'コク、きざ-む', ''),
    ('0661', '国國', 'コク、くに', ''),
    ('0662', '黒黑', 'コク、くろ、くろ-い', ''),
    ('0663', '穀穀', 'コク', ''),
    ('0664', '酷', 'コク', ''),
    ('0665', '獄', 'ゴク', ''),
    ('0666', '骨', 'コツ、ほね', ''),
    ('0667', '駒', 'こま', ''),
    ('0668', '込', 'こ-む、こ-める', ''),
    ('0669', '頃', 'ころ', ''),
    ('0670', '今', 'コン、キン、いま', '今日、今朝、今年'),
    ('0671', '困', 'コン、こま-る', ''),
    ('0672', '昆', 'コン', '昆布'),
    ('0673', '恨', 'コン、うら-む、うら-めしい', ''),
    ('0674', '根', 'コン、ね', ''),
    ('0675', '婚', 'コン', ''),
    ('0676', '混', 'コン、ま-じる、ま-ざる、ま-ぜる、こ-む', ''),
    ('0677', '痕', 'コン、あと', ''),
    ('0678', '紺', 'コン', '紺屋'),
    ('0679', '魂', 'コン、たましい', ''),
    ('0680', '墾', 'コン', ''),
    ('0681', '懇', 'コン、ねんご-ろ', ''),
    ('0682', '左', 'サ、ひだり', ''),
    ('0683', '佐', 'サ', ''),
    ('0684', '沙', 'サ', ''),
    ('0685', '査', 'サ', ''),
    ('0686', '砂', 'サ、シャ、すな', '砂利'),
    ('0687', '唆', 'サ、そそのか-す', ''),
    ('0688', '差', 'サ、さ-す', '差し支える'),
    ('0689', '詐', 'サ', ''),
    ('0690', '鎖', 'サ、くさり', ''),
    ('0691', '座', 'ザ、すわ-る', ''),
    ('0692', '挫', 'ザ', ''),
    ('0693', '才', 'サイ', ''),
    ('0694', '再', 'サイ、（サ）、ふたた-び', ''),
    ('0695', '災', 'サイ、わざわ-い', ''),
    ('0696', '妻', 'サイ、つま', ''),
    ('0697', '采', 'サイ', ''),
    ('0698', '砕碎', 'サイ、くだ-く、くだ-ける', ''),
    ('0699', '宰', 'サイ', ''),
    ('0700', '栽', 'サイ', ''),
    ('0701', '彩', 'サイ、いろど-る', ''),
    ('0702', '採', 'サイ、と-る', ''),
    ('0703', '済濟', 'サイ、す-む、す-ます', ''),
    ('0704', '祭', 'サイ、まつ-る、まつ-り', ''),
    ('0705', '斎齋', 'サイ', ''),
    ('0706', '細', 'サイ、ほそ-い、ほそ-る、こま-か、こま-かい', ''),
    ('0707', '菜', 'サイ、な', ''),
    ('0708', '最', 'サイ、もっと-も', '最寄り'),
    ('0709', '裁', 'サイ、た-つ、さば-く', ''),
    ('0710', '債', 'サイ', ''),
    ('0711', '催', 'サイ、もよお-す', ''),
    ('0712', '塞', 'サイ、ソク、ふさ-ぐ、ふさ-がる', ''),
    ('0713', '歳歲', 'サイ、（セイ）', '二十歳'),
    ('0714', '載', 'サイ、の-せる、の-る', ''),
    ('0715', '際', 'サイ、きわ', ''),
    ('0716', '埼', '（さい）', '埼玉'),
    ('0717', '在', 'ザイ、あ-る', ''),
    ('0718', '材', 'ザイ', ''),
    ('0719', '剤劑', 'ザイ', ''),
    ('0720', '財', 'ザイ、（サイ）', ''),
    ('0721', '罪', 'ザイ、つみ', ''),
    ('0722', '崎', 'さき', ''),
    ('0723', '作', 'サク、サ、つく-る', ''),
    ('0724', '削', 'サク、けず-る', ''),
    ('0725', '昨', 'サク', '昨日'),
    ('0726', '柵', 'サク', ''),
    ('0727', '索', 'サク', ''),
    ('0728', '策', 'サク', ''),
    ('0729', '酢', 'サク、す', ''),
    ('0730', '搾', 'サク、しぼ-る', ''),
    ('0731', '錯', 'サク', ''),
    ('0732', '咲', 'さ-く', ''),
    ('0733', '冊册', 'サツ、サク', ''),
    ('0734', '札', 'サツ、ふだ', ''),
    ('0735', '刷', 'サツ、す-る', ''),
    ('0736', '刹', 'サツ、セツ', ''),
    ('0737', '拶', 'サツ', ''),
    ('0738', '殺殺', 'サツ、（サイ）、（セツ）、ころ-す', ''),
    ('0739', '察', 'サツ', ''),
    ('0740', '撮', 'サツ、と-る', ''),
    ('0741', '擦', 'サツ、す-る、す-れる', ''),
    ('0742', '雑雜', 'ザツ、ゾウ', '雑魚'),
    ('0743', '皿', 'さら', ''),
    ('0744', '三', 'サン、み、み-つ、みっ-つ', '三味線'),
    ('0745', '山', 'サン、やま', '山車、築山、富山'),
    ('0746', '参參', 'サン、まい-る', ''),
    ('0747', '桟棧', 'サン', '桟敷'),
    ('0748', '蚕蠶', 'サン、かいこ', ''),
    ('0749', '惨慘', 'サン、ザン、みじ-め', ''),
    ('0750', '産產', 'サン、う-む、う-まれる、うぶ', '土産'),
    ('0751', '傘', 'サン、かさ', ''),
    ('0752', '散', 'サン、ち-る、ち-らす、ち-らかす、ち-らかる', ''),
    ('0753', '算', 'サン', ''),
    ('0754', '酸', 'サン、す-い', ''),
    ('0755', '賛贊', 'サン', ''),
    ('0756', '残殘', 'ザン、のこ-る、のこ-す', '名残'),
    ('0757', '斬', 'ザン、き-る', ''),
    ('0758', '暫', 'ザン', ''),
    ('0759', '士', 'シ', '海士、居士、博士'),
    ('0760', '子', 'シ、ス、こ', '迷子、息子'),
    ('0761', '支', 'シ、ささ-える', '差し支える'),
    ('0762', '止', 'シ、と-まる、と-める', '波止場'),
    ('0763', '氏', 'シ、うじ', ''),
    ('0764', '仕', 'シ、（ジ）、つか-える', ''),
    ('0765', '史', 'シ', ''),
    ('0766', '司', 'シ', ''),
    ('0767', '四', 'シ、よ、よ-つ、よっ-つ、よん', ''),
    ('0768', '市', 'シ、いち', ''),
    ('0769', '矢', 'シ、や', ''),
    ('0770', '旨', 'シ、むね', ''),
    ('0771', '死', 'シ、し-ぬ', ''),
    ('0772', '糸絲', 'シ、いと', ''),
    ('0773', '至', 'シ、いた-る', ''),
    ('0774', '伺', 'シ、うかが-う', ''),
    ('0775', '志', 'シ、こころざ-す、こころざし', ''),
    ('0776', '私', 'シ、わたくし、わたし', ''),
    ('0777', '使', 'シ、つか-う', ''),
    ('0778', '刺', 'シ、さ-す、さ-さる', ''),
    ('0779', '始', 'シ、はじ-める、はじ-まる', ''),
    ('0780', '姉', 'シ、あね', '姉さん'),
    ('0781', '枝', 'シ、えだ', ''),
    ('0782', '祉祉', 'シ', ''),
    ('0783', '肢', 'シ', ''),
    ('0784', '姿', 'シ、すがた', ''),
    ('0785', '思', 'シ、おも-う', ''),
    ('0786', '指', 'シ、ゆび、さ-す', ''),
    ('0787', '施', 'シ、セ、ほどこ-す', ''),
    ('0788', '師', 'シ', '師走'),
    ('0789', '恣', 'シ', ''),
    ('0790', '紙', 'シ、かみ', ''),
    ('0791', '脂', 'シ、あぶら', ''),
    ('0792', '視視', 'シ', ''),
    ('0793', '紫', 'シ、むらさき', ''),
    ('0794', '詞', 'シ', '祝詞'),
    ('0795', '歯齒', 'シ、は', ''),
    ('0796', '嗣', 'シ', ''),
    ('0797', '試', 'シ、こころ-みる、ため-す', ''),
    ('0798', '詩', 'シ', '詩歌'),
    ('0799', '資', 'シ', ''),
    ('0800', '飼飼', 'シ、か-う', ''),
    ('0801', '誌', 'シ', ''),
    ('0802', '雌', 'シ、め、めす', ''),
    ('0803', '摯', 'シ', ''),
    ('0804', '賜', 'シ、たまわ-る', ''),
    ('0805', '諮', 'シ、はか-る', ''),
    ('0806', '示', 'ジ、シ、しめ-す', ''),
    ('0807', '字', 'ジ、あざ', '文字'),
    ('0808', '寺', 'ジ、てら', ''),
    ('0809', '次', 'ジ、シ、つ-ぐ、つぎ', ''),
    ('0810', '耳', 'ジ、みみ', ''),
    ('0811', '自', 'ジ、シ、みずか-ら', ''),
    ('0812', '似', 'ジ、に-る', ''),
    ('0813', '児兒', 'ジ、（ニ）', '稚児、鹿児島'),
    ('0814', '事', 'ジ、（ズ）、こと', ''),
    ('0815', '侍', 'ジ、さむらい', ''),
    ('0816', '治', 'ジ、チ、おさ-める、おさ-まる、なお-る、なお-す', ''),
    ('0817', '持', 'ジ、も-つ', ''),
    ('0818', '時', 'ジ、とき', '時雨、時計'),
    ('0819', '滋', 'ジ', '滋賀'),
    ('0820', '慈', 'ジ、いつく-しむ', ''),
    ('0821', '辞辭', 'ジ、や-める', ''),
    ('0822', '磁', 'ジ', ''),
    ('0823', '餌', 'ジ、えさ、え', ''),
    ('0824', '璽', 'ジ', ''),
    ('0825', '鹿', 'しか、（か）', '鹿児島'),
    ('0826', '式', 'シキ', ''),
    ('0827', '識', 'シキ', ''),
    ('0828', '軸', 'ジク', ''),
    ('0829', '七', 'シチ、なな、なな-つ、（なの）', '七夕、七日'),
    ('0830', '𠮟(叱)', 'シツ、しか-る', ''),
    ('0831', '失', 'シツ、うしな-う', ''),
    ('0832', '室', 'シツ、むろ', ''),
    ('0833', '疾', 'シツ', ''),
    ('0834', '執', 'シツ、シュウ、と-る', ''),
    ('0835', '湿濕', 'シツ、しめ-る、しめ-す', ''),
    ('0836', '嫉', 'シツ', ''),
    ('0837', '漆', 'シツ、うるし', ''),
    ('0838', '質', 'シツ、シチ、（チ）', ''),
    ('0839', '実實', 'ジツ、み、みの-る', ''),
    ('0840', '芝', 'しば', '芝生'),
    ('0841', '写寫', 'シャ、うつ-す、うつ-る', ''),
    ('0842', '社社', 'シャ、やしろ', ''),
    ('0843', '車', 'シャ、くるま', '山車'),
    ('0844', '舎舍', 'シャ', '田舎'),
    ('0845', '者者', 'シャ、もの', '猛者'),
    ('0846', '射', 'シャ、い-る', ''),
    ('0847', '捨', 'シャ、す-てる', ''),
    ('0848', '赦', 'シャ', ''),
    ('0849', '斜', 'シャ、なな-め', ''),
    ('0850', '煮煮', 'シャ、に-る、に-える、に-やす', ''),
    ('0851', '遮', 'シャ、さえぎ-る', ''),
    ('0852', '謝', 'シャ、あやま-る', ''),
    ('0853', '邪', 'ジャ', '風邪'),
    ('0854', '蛇', 'ジャ、ダ、へび', ''),
    ('0855', '尺', 'シャク', ''),
    ('0856', '借', 'シャク、か-りる', ''),
    ('0857', '酌', 'シャク、く-む', ''),
    ('0858', '釈釋', 'シャク', ''),
    ('0859', '爵', 'シャク', ''),
    ('0860', '若', 'ジャク、（ニャク）、わか-い、も-しくは', '若人'),
    ('0861', '弱', 'ジャク、よわ-い、よわ-る、よわ-まる、よわ-める', ''),
    ('0862', '寂', 'ジャク、（セキ）、さび、さび-しい、さび-れる', ''),
    ('0863', '手', 'シュ、て、（た）', '上手、手伝う、下手'),
    ('0864', '主', 'シュ、（ス）、ぬし、おも', ''),
    ('0865', '守', 'シュ、（ス）、まも-る、も-り', ''),
    ('0866', '朱', 'シュ', ''),
    ('0867', '取', 'シュ、と-る', '鳥取'),
    ('0868', '狩', 'シュ、か-る、か-り', ''),
    ('0869', '首', 'シュ、くび', ''),
    ('0870', '殊', 'シュ、こと', ''),
    ('0871', '珠', 'シュ', '数珠'),
    ('0872', '酒', 'シュ、さけ、（さか）', 'お神酒'),
    ('0873', '腫', 'シュ、は-れる、は-らす', ''),
    ('0874', '種', 'シュ、たね', ''),
    ('0875', '趣', 'シュ、おもむき', ''),
    ('0876', '寿壽', 'ジュ、ことぶき', ''),
    ('0877', '受', 'ジュ、う-ける、う-かる', ''),
    ('0878', '呪', 'ジュ、のろ-う', ''),
    ('0879', '授', 'ジュ、さず-ける、さず-かる', ''),
    ('0880', '需', 'ジュ', ''),
    ('0881', '儒', 'ジュ', ''),
    ('0882', '樹', 'ジュ', ''),
    ('0883', '収收', 'シュウ、おさ-める、おさ-まる', ''),
    ('0884', '囚', 'シュウ', ''),
    ('0885', '州', 'シュウ、す', ''),
    ('0886', '舟', 'シュウ、ふね、（ふな）', ''),
    ('0887', '秀', 'シュウ、ひい-でる', ''),
    ('0888', '周', 'シュウ、まわ-り', ''),
    ('0889', '宗', 'シュウ、ソウ', ''),
    ('0890', '拾', 'シュウ、ジュウ、ひろ-う', ''),
    ('0891', '秋', 'シュウ、あき', ''),
    ('0892', '臭臭', 'シュウ、くさ-い、にお-う', ''),
    ('0893', '修', 'シュウ、（シュ）、おさ-める、おさ-まる', ''),
    ('0894', '袖', 'シュウ、そで', ''),
    ('0895', '終', 'シュウ、お-わる、お-える', ''),
    ('0896', '羞', 'シュウ', ''),
    ('0897', '習', 'シュウ、なら-う', ''),
    ('0898', '週', 'シュウ', ''),
    ('0899', '就', 'シュウ、（ジュ）、つ-く、つ-ける', ''),
    ('0900', '衆', 'シュウ、（シュ）', ''),
    ('0901', '集', 'シュウ、あつ-まる、あつ-める、つど-う', ''),
    ('0902', '愁', 'シュウ、うれ-える、うれ-い', ''),
    ('0903', '酬', 'シュウ', ''),
    ('0904', '醜', 'シュウ、みにく-い', ''),
    ('0905', '蹴', 'シュウ、け-る', ''),
    ('0906', '襲', 'シュウ、おそ-う', ''),
    ('0907', '十', 'ジュウ、ジッ、とお、と', '十重二十重、二十、二十歳、二十日、十'),
    ('0908', '汁', 'ジュウ、しる', ''),
    ('0909', '充', 'ジュウ、あ-てる', ''),
    ('0910', '住', 'ジュウ、す-む、す-まう', ''),
    ('0911', '柔', 'ジュウ、ニュウ、やわ-らか、やわ-らかい', ''),
    ('0912', '重', 'ジュウ、チョウ、え、おも-い、かさ-ねる、かさ-なる', '十重二十重'),
    ('0913', '従從', 'ジュウ、（ショウ）、（ジュ）、したが-う、したが-える', ''),
    ('0914', '渋澁', 'ジュウ、しぶ、しぶ-い、しぶ-る', ''),
    ('0915', '銃', 'ジュウ', ''),
    ('0916', '獣獸', 'ジュウ、けもの', ''),
    ('0917', '縦縱', 'ジュウ、たて', ''),
    ('0918', '叔', 'シュク', '叔父、叔母'),
    ('0919', '祝祝', 'シュク、（シュウ）、いわ-う', '祝詞'),
    ('0920', '宿', 'シュク、やど、やど-る、やど-す', ''),
    ('0921', '淑', 'シュク', ''),
    ('0922', '粛肅', 'シュク', ''),
    ('0923', '縮', 'シュク、ちぢ-む、ちぢ-まる、ちぢ-める、ちぢ-れる、ちぢ-らす', ''),
    ('0924', '塾', 'ジュク', ''),
    ('0925', '熟', 'ジュク、う-れる', ''),
    ('0926', '出', 'シュツ、（スイ）、で-る、だ-す', ''),
    ('0927', '述', 'ジュツ、の-べる', ''),
    ('0928', '術', 'ジュツ', ''),
    ('0929', '俊', 'シュン', ''),
    ('0930', '春', 'シュン、はる', ''),
    ('0931', '瞬', 'シュン、またた-く', ''),
    ('0932', '旬', 'ジュン、（シュン）', ''),
    ('0933', '巡', 'ジュン、めぐ-る', 'お巡りさん'),
    ('0934', '盾', 'ジュン、たて', ''),
    ('0935', '准', 'ジュン', ''),
    ('0936', '殉', 'ジュン', ''),
    ('0937', '純', 'ジュン', ''),
    ('0938', '循', 'ジュン', ''),
    ('0939', '順', 'ジュン', ''),
    ('0940', '準', 'ジュン', ''),
    ('0941', '潤', 'ジュン、うるお-う、うるお-す、うる-む', ''),
    ('0942', '遵', 'ジュン', ''),
    ('0943', '処處', 'ショ', ''),
    ('0944', '初', 'ショ、はじ-め、はじ-めて、はつ、うい、そ-める', ''),
    ('0945', '所', 'ショ、ところ', ''),
    ('0946', '書', 'ショ、か-く', ''),
    ('0947', '庶', 'ショ', ''),
    ('0948', '暑暑', 'ショ、あつ-い', ''),
    ('0949', '署署', 'ショ', ''),
    ('0950', '緒緖', 'ショ、（チョ）、お', ''),
    ('0951', '諸諸', 'ショ', ''),
    ('0952', '女', 'ジョ、ニョ、（ニョウ）、おんな、め', '海女、乙女、早乙女'),
    ('0953', '如', 'ジョ、ニョ', ''),
    ('0954', '助', 'ジョ、たす-ける、たす-かる、すけ', ''),
    ('0955', '序', 'ジョ', ''),
    ('0956', '叙敍', 'ジョ', ''),
    ('0957', '徐', 'ジョ', ''),
    ('0958', '除', 'ジョ、（ジ）、のぞ-く', ''),
    ('0959', '小', 'ショウ、ちい-さい、こ、お', '小豆'),
    ('0960', '升', 'ショウ、ます', ''),
    ('0961', '少', 'ショウ、すく-ない、すこ-し', ''),
    ('0962', '召', 'ショウ、め-す', ''),
    ('0963', '匠', 'ショウ', ''),
    ('0964', '床', 'ショウ、とこ、ゆか', ''),
    ('0965', '抄', 'ショウ', ''),
    ('0966', '肖', 'ショウ', ''),
    ('0967', '尚尙', 'ショウ', ''),
    ('0968', '招', 'ショウ、まね-く', ''),
    ('0969', '承', 'ショウ、うけたまわ-る', ''),
    ('0970', '昇', 'ショウ、のぼ-る', ''),
    ('0971', '松', 'ショウ、まつ', ''),
    ('0972', '沼', 'ショウ、ぬま', ''),
    ('0973', '昭', 'ショウ', ''),
    ('0974', '宵', 'ショウ、よい', ''),
    ('0975', '将將', 'ショウ', ''),
    ('0976', '消', 'ショウ、き-える、け-す', ''),
    ('0977', '症', 'ショウ', ''),
    ('0978', '祥祥', 'ショウ', ''),
    ('0979', '称稱', 'ショウ', ''),
    ('0980', '笑', 'ショウ、わら-う、え-む', '笑顔'),
    ('0981', '唱', 'ショウ、とな-える', ''),
    ('0982', '商', 'ショウ、あきな-う', ''),
    ('0983', '渉涉', 'ショウ', ''),
    ('0984', '章', 'ショウ', ''),
    ('0985', '紹', 'ショウ', ''),
    ('0986', '訟', 'ショウ', ''),
    ('0987', '勝', 'ショウ、か-つ、まさ-る', ''),
    ('0988', '掌', 'ショウ', ''),
    ('0989', '晶', 'ショウ', ''),
    ('0990', '焼燒', 'ショウ、や-く、や-ける', ''),
    ('0991', '焦', 'ショウ、こ-げる、こ-がす、こ-がれる、あせ-る', ''),
    ('0992', '硝', 'ショウ', ''),
    ('0993', '粧', 'ショウ', ''),
    ('0994', '詔', 'ショウ、みことのり', ''),
    ('0995', '証證', 'ショウ', ''),
    ('0996', '象', 'ショウ、ゾウ', ''),
    ('0997', '傷', 'ショウ、きず、いた-む、いた-める', ''),
    ('0998', '奨奬', 'ショウ', ''),
    ('0999', '照', 'ショウ、て-る、て-らす、て-れる', ''),
    ('1000', '詳', 'ショウ、くわ-しい', ''),
    ('1001', '彰', 'ショウ', ''),
    ('1002', '障', 'ショウ、さわ-る', ''),
    ('1003', '憧', 'ショウ、あこが-れる', '憧憬'),
    ('1004', '衝', 'ショウ', ''),
    ('1005', '賞', 'ショウ', ''),
    ('1006', '償', 'ショウ、つぐな-う', ''),
    ('1007', '礁', 'ショウ', ''),
    ('1008', '鐘', 'ショウ、かね', ''),
    ('1009', '上', 'ジョウ、（ショウ）、うえ、（うわ）、かみ、あ-げる、あ-がる、のぼ-る、のぼ-せる、のぼ-す', '上手'),
    ('1010', '丈', 'ジョウ、たけ', ''),
    ('1011', '冗', 'ジョウ', ''),
    ('1012', '条條', 'ジョウ', ''),
    ('1013', '状狀', 'ジョウ', ''),
    ('1014', '乗乘', 'ジョウ、の-る、の-せる', ''),
    ('1015', '城', 'ジョウ、しろ', '茨城、宮城'),
    ('1016', '浄淨', 'ジョウ', ''),
    ('1017', '剰剩', 'ジョウ', ''),
    ('1018', '常', 'ジョウ、つね、とこ', ''),
    ('1019', '情', 'ジョウ、（セイ）、なさ-け', ''),
    ('1020', '場', 'ジョウ、ば', '波止場'),
    ('1021', '畳疊', 'ジョウ、たた-む、たたみ', ''),
    ('1022', '蒸', 'ジョウ、む-す、む-れる、む-らす', ''),
    ('1023', '縄繩', 'ジョウ、なわ', ''),
    ('1024', '壌壤', 'ジョウ', ''),
    ('1025', '嬢孃', 'ジョウ', ''),
    ('1026', '錠', 'ジョウ', ''),
    ('1027', '譲讓', 'ジョウ、ゆず-る', ''),
    ('1028', '醸釀', 'ジョウ、かも-す', ''),
    ('1029', '色', 'ショク、シキ、いろ', '景色'),
    ('1030', '拭', 'ショク、ふ-く、ぬぐ-う', ''),
    ('1031', '食', 'ショク、（ジキ）、く-う、く-らう、た-べる', ''),
    ('1032', '植', 'ショク、う-える、う-わる', ''),
    ('1033', '殖', 'ショク、ふ-える、ふ-やす', ''),
    ('1034', '飾', 'ショク、かざ-る', ''),
    ('1035', '触觸', 'ショク、ふ-れる、さわ-る', ''),
    ('1036', '嘱囑', 'ショク', ''),
    ('1037', '織', 'ショク、シキ、お-る', ''),
    ('1038', '職', 'ショク', ''),
    ('1039', '辱', 'ジョク、はずかし-める', ''),
    ('1040', '尻', 'しり', '尻尾'),
    ('1041', '心', 'シン、こころ', '心地'),
    ('1042', '申', 'シン、もう-す', ''),
    ('1043', '伸', 'シン、の-びる、の-ばす、の-べる', ''),
    ('1044', '臣', 'シン、ジン', ''),
    ('1045', '芯', 'シン', ''),
    ('1046', '身', 'シン、み', ''),
    ('1047', '辛', 'シン、から-い', ''),
    ('1048', '侵', 'シン、おか-す', ''),
    ('1049', '信', 'シン', ''),
    ('1050', '津', 'シン、つ', ''),
    ('1051', '神神', 'シン、ジン、かみ、（かん）、（こう）', 'お神酒、神楽、神奈川'),
    ('1052', '唇', 'シン、くちびる', ''),
    ('1053', '娠', 'シン', ''),
    ('1054', '振', 'シン、ふ-る、ふ-るう、ふ-れる', ''),
    ('1055', '浸', 'シン、ひた-す、ひた-る', ''),
    ('1056', '真眞', 'シン、ま', '真面目、真っ赤、真っ青'),
    ('1057', '針', 'シン、はり', ''),
    ('1058', '深', 'シン、ふか-い、ふか-まる、ふか-める', ''),
    ('1059', '紳', 'シン', ''),
    ('1060', '進', 'シン、すす-む、すす-める', ''),
    ('1061', '森', 'シン、もり', ''),
    ('1062', '診', 'シン、み-る', ''),
    ('1063', '寝寢', 'シン、ね-る、ね-かす', ''),
    ('1064', '慎愼', 'シン、つつし-む', ''),
    ('1065', '新', 'シン、あたら-しい、あら-た、にい', ''),
    ('1066', '審', 'シン', ''),
    ('1067', '震', 'シン、ふる-う、ふる-える', ''),
    ('1068', '薪', 'シン、たきぎ', ''),
    ('1069', '親', 'シン、おや、した-しい、した-しむ', ''),
    ('1070', '人', 'ジン、ニン、ひと', '大人、玄人、素人、仲人、一人、二人、若人'),
    ('1071', '刃', 'ジン、は', ''),
    ('1072', '仁', 'ジン、（ニ）', ''),
    ('1073', '尽盡', 'ジン、つ-くす、つ-きる、つ-かす', ''),
    ('1074', '迅', 'ジン', ''),
    ('1075', '甚', 'ジン、はなは-だ、はなは-だしい', ''),
    ('1076', '陣', 'ジン', ''),
    ('1077', '尋', 'ジン、たず-ねる', ''),
    ('1078', '腎', 'ジン', ''),
    ('1079', '須', 'ス', ''),
    ('1080', '図圖', 'ズ、ト、はか-る', ''),
    ('1081', '水', 'スイ、みず', '清水'),
    ('1082', '吹', 'スイ、ふ-く', '息吹、吹雪'),
    ('1083', '垂', 'スイ、た-れる、た-らす', ''),
    ('1084', '炊', 'スイ、た-く', ''),
    ('1085', '帥', 'スイ', ''),
    ('1086', '粋粹', 'スイ、いき', ''),
    ('1087', '衰', 'スイ、おとろ-える', ''),
    ('1088', '推', 'スイ、お-す', ''),
    ('1089', '酔醉', 'スイ、よ-う', ''),
    ('1090', '遂', 'スイ、と-げる', ''),
    ('1091', '睡', 'スイ', ''),
    ('1092', '穂穗', 'スイ、ほ', ''),
    ('1093', '随隨', 'ズイ', ''),
    ('1094', '髄髓', 'ズイ', ''),
    ('1095', '枢樞', 'スウ', ''),
    ('1096', '崇', 'スウ', ''),
    ('1097', '数數', 'スウ、（ス）、かず、かぞ-える', '数珠、数寄屋、数奇屋'),
    ('1098', '据', 'す-える、す-わる', ''),
    ('1099', '杉', 'すぎ', ''),
    ('1100', '裾', 'すそ', ''),
    ('1101', '寸', 'スン', ''),
    ('1102', '瀬瀨', 'せ', ''),
    ('1103', '是', 'ゼ', ''),
    ('1104', '井', 'セイ、（ショウ）、い', ''),
    ('1105', '世', 'セイ、セ、よ', ''),
    ('1106', '正', 'セイ、ショウ、ただ-しい、ただ-す、まさ', ''),
    ('1107', '生',
     'セイ、ショウ、い-きる、い-かす、い-ける、う-まれる、う-む、お-う、は-える、は-やす、き、なま',
     '芝生、弥生'),
    ('1108', '成', 'セイ、（ジョウ）、な-る、な-す', ''),
    ('1109', '西', 'セイ、サイ、にし', ''),
    ('1110', '声聲', 'セイ、（ショウ）、こえ、（こわ）', ''),
    ('1111', '制', 'セイ', ''),
    ('1112', '姓', 'セイ、ショウ', ''),
    ('1113', '征', 'セイ', ''),
    ('1114', '性', 'セイ、ショウ', ''),
    ('1115', '青靑', 'セイ、（ショウ）、あお、あお-い', '真っ青'),
    ('1116', '斉齊', 'セイ', ''),
    ('1117', '政', 'セイ、（ショウ）、まつりごと', ''),
    ('1118', '星', 'セイ、（ショウ）、ほし', ''),
    ('1119', '牲', 'セイ', ''),
    ('1120', '省', 'セイ、ショウ、かえり-みる、はぶ-く', ''),
    ('1121', '凄', 'セイ', ''),
    ('1122', '逝', 'セイ、ゆ-く、い-く', ''),
    ('1123', '清淸', 'セイ、（ショウ）、きよ-い、きよ-まる、きよ-める', '清水'),
    ('1124', '盛', 'セイ、（ジョウ）、も-る、さか-る、さか-ん', ''),
    ('1125', '婿', 'セイ、むこ', ''),
    ('1126', '晴晴', 'セイ、は-れる、は-らす', ''),
    ('1127', '勢', 'セイ、いきお-い', ''),
    ('1128', '聖', 'セイ', ''),
    ('1129', '誠', 'セイ、まこと', ''),
    ('1130', '精精', 'セイ、（ショウ）', ''),
    ('1131', '製', 'セイ', ''),
    ('1132', '誓', 'セイ、ちか-う', ''),
    ('1133', '静靜', 'セイ、（ジョウ）、しず、しず-か、しず-まる、しず-める', ''),
    ('1134', '請', 'セイ、（シン）、こ-う、う-ける', ''),
    ('1135', '整', 'セイ、ととの-える、ととの-う', ''),
    ('1136', '醒', 'セイ', ''),
    ('1137', '税稅', 'ゼイ', ''),
    ('1138', '夕', 'セキ、ゆう', '七夕'),
    ('1139', '斥', 'セキ', ''),
    ('1140', '石', 'セキ、（シャク）、（コク）、いし', ''),
    ('1141', '赤', 'セキ、（シャク）、あか、あか-い、あか-らむ、あか-らめる', '真っ赤'),
    ('1142', '昔', 'セキ、（シャク）、むかし', ''),
    ('1143', '析', 'セキ', ''),
    ('1144', '席', 'セキ', '寄席'),
    ('1145', '脊', 'セキ', ''),
    ('1146', '隻', 'セキ', ''),
    ('1147', '惜', 'セキ、お-しい、お-しむ', ''),
    ('1148', '戚', 'セキ', ''),
    ('1149', '責', 'セキ、せ-める', ''),
    ('1150', '跡', 'セキ、あと', ''),
    ('1151', '積', 'セキ、つ-む、つ-もる', ''),
    ('1152', '績', 'セキ', ''),
    ('1153', '籍', 'セキ', ''),
    ('1154', '切', 'セツ、（サイ）、き-る、き-れる', ''),
    ('1155', '折', 'セツ、お-る、おり、お-れる', ''),
    ('1156', '拙', 'セツ、つたな-い', ''),
    ('1157', '窃竊', 'セツ', ''),
    ('1158', '接', 'セツ、つ-ぐ', ''),
    ('1159', '設', 'セツ、もう-ける', ''),
    ('1160', '雪', 'セツ、ゆき', '雪崩、吹雪'),
    ('1161', '摂攝', 'セツ', ''),
    ('1162', '節節', 'セツ、（セチ）、ふし', ''),
    ('1163', '説說', 'セツ、（ゼイ）、と-く', ''),
    ('1164', '舌', 'ゼツ、した', ''),
    ('1165', '絶絕', 'ゼツ、た-える、た-やす、た-つ', ''),
    ('1166', '千', 'セン、ち', ''),
    ('1167', '川', 'セン、かわ', '川原、神奈川'),
    ('1168', '仙', 'セン', ''),
    ('1169', '占', 'セン、し-める、うらな-う', ''),
    ('1170', '先', 'セン、さき', ''),
    ('1171', '宣', 'セン', ''),
    ('1172', '専專', 'セン、もっぱ-ら', ''),
    ('1173', '泉', 'セン、いずみ', ''),
    ('1174', '浅淺', 'セン、あさ-い', ''),
    ('1175', '洗', 'セン、あら-う', ''),
    ('1176', '染', 'セン、そ-める、そ-まる、し-みる、し-み', ''),
    ('1177', '扇', 'セン、おうぎ', ''),
    ('1178', '栓', 'セン', ''),
    ('1179', '旋', 'セン', ''),
    ('1180', '船', 'セン、ふね、（ふな）', '伝馬船'),
    ('1181', '戦戰', 'セン、いくさ、たたか-う', ''),
    ('1182', '煎', 'セン、い-る', ''),
    ('1183', '羨', 'セン、うらや-む、うらや-ましい', ''),
    ('1184', '腺', 'セン', ''),
    ('1185', '詮', 'セン', ''),
    ('1186', '践踐', 'セン', ''),
    ('1187', '箋', 'セン', ''),
    ('1188', '銭錢', 'セン、ぜに', ''),
    ('1189', '潜潛', 'セン、ひそ-む、もぐ-る', ''),
    ('1190', '線', 'セン', '三味線'),
    ('1191', '遷', 'セン', ''),
    ('1192', '選', 'セン、えら-ぶ', ''),
    ('1193', '薦', 'セン、すす-める', ''),
    ('1194', '繊纖', 'セン', ''),
    ('1195', '鮮', 'セン、あざ-やか', ''),
    ('1196', '全', 'ゼン、まった-く、すべ-て', ''),
    ('1197', '前', 'ゼン、まえ', ''),
    ('1198', '善', 'ゼン、よ-い', ''),
    ('1199', '然', 'ゼン、ネン', ''),
    ('1200', '禅禪', 'ゼン', ''),
    ('1201', '漸', 'ゼン', ''),
    ('1202', '膳', 'ゼン', ''),
    ('1203', '繕', 'ゼン、つくろ-う', ''),
    ('1204', '狙', 'ソ、ねら-う', ''),
    ('1205', '阻', 'ソ、はば-む', ''),
    ('1206', '祖祖', 'ソ', ''),
    ('1207', '租', 'ソ', ''),
    ('1208', '素', 'ソ、ス', '素人'),
    ('1209', '措', 'ソ', ''),
    ('1210', '粗', 'ソ、あら-い', ''),
    ('1211', '組', 'ソ、く-む、くみ', ''),
    ('1212', '疎', 'ソ、うと-い、うと-む', ''),
    ('1213', '訴', 'ソ、うった-える', ''),
    ('1214', '塑', 'ソ', ''),
    ('1215', '遡', 'ソ、さかのぼ-る', ''),
    ('1216', '礎', 'ソ、いしずえ', ''),
    ('1217', '双雙', 'ソウ、ふた', ''),
    ('1218', '壮壯', 'ソウ', ''),
    ('1219', '早', 'ソウ、（サッ）、はや-い、はや-まる、はや-める', '早乙女、早苗'),
    ('1220', '争爭', 'ソウ、あらそ-う', ''),
    ('1221', '走', 'ソウ、はし-る', '師走'),
    ('1222', '奏', 'ソウ、かな-でる', ''),
    ('1223', '相', 'ソウ、ショウ、あい', '相撲'),
    ('1224', '荘莊', 'ソウ', ''),
    ('1225', '草', 'ソウ、くさ', '草履'),
    ('1226', '送', 'ソウ、おく-る', ''),
    ('1227', '倉', 'ソウ、くら', ''),
    ('1228', '捜搜', 'ソウ、さが-す', ''),
    ('1229', '挿揷', 'ソウ、さ-す', ''),
    ('1230', '桑', 'ソウ、くわ', ''),
    ('1231', '巣巢', 'ソウ、す', ''),
    ('1232', '掃', 'ソウ、は-く', ''),
    ('1233', '曹', 'ソウ', ''),
    ('1234', '曽曾', 'ソウ、（ゾ）', ''),
    ('1235', '爽', 'ソウ、さわ-やか', ''),
    ('1236', '窓', 'ソウ、まど', ''),
    ('1237', '創', 'ソウ、つく-る', ''),
    ('1238', '喪', 'ソウ、も', ''),
    ('1239', '痩瘦', 'ソウ、や-せる', ''),
    ('1240', '葬', 'ソウ、ほうむ-る', ''),
    ('1241', '装裝', 'ソウ、ショウ、よそお-う', ''),
    ('1242', '僧僧', 'ソウ', ''),
    ('1243', '想', 'ソウ、（ソ）', ''),
    ('1244', '層層', 'ソウ', ''),
    ('1245', '総總', 'ソウ', ''),
    ('1246', '遭', 'ソウ、あ-う', ''),
    ('1247', '槽', 'ソウ', ''),
    ('1248', '踪', 'ソウ', ''),
    ('1249', '操', 'ソウ、みさお、あやつ-る', ''),
    ('1250', '燥', 'ソウ', ''),
    ('1251', '霜', 'ソウ、しも', ''),
    ('1252', '騒騷', 'ソウ、さわ-ぐ', ''),
    ('1253', '藻', 'ソウ、も', ''),
    ('1254', '造', 'ゾウ、つく-る', ''),
    ('1255', '像', 'ゾウ', ''),
    ('1256', '増增', 'ゾウ、ま-す、ふ-える、ふ-やす', ''),
    ('1257', '憎憎', 'ゾウ、にく-む、にく-い、にく-らしい、にく-しみ', ''),
    ('1258', '蔵藏', 'ゾウ、くら', ''),
    ('1259', '贈贈', 'ゾウ、（ソウ）、おく-る', ''),
    ('1260', '臓臟', 'ゾウ', ''),
    ('1261', '即卽', 'ソク', ''),
    ('1262', '束', 'ソク、たば', ''),
    ('1263', '足', 'ソク、あし、た-りる、た-る、た-す', '足袋'),
    ('1264', '促', 'ソク、うなが-す', ''),
    ('1265', '則', 'ソク', ''),
    ('1266', '息', 'ソク、いき', '息吹、息子'),
    ('1267', '捉', 'ソク、とら-える', ''),
    ('1268', '速', 'ソク、はや-い、はや-める、はや-まる、すみ-やか', ''),
    ('1269', '側', 'ソク、がわ', '側'),
    ('1270', '測', 'ソク、はか-る', ''),
    ('1271', '俗', 'ゾク', ''),
    ('1272', '族', 'ゾク', ''),
    ('1273', '属屬', 'ゾク', ''),
    ('1274', '賊', 'ゾク', ''),
    ('1275', '続續', 'ゾク、つづ-く、つづ-ける', ''),
    ('1276', '卒', 'ソツ', ''),
    ('1277', '率', 'ソツ、リツ、ひき-いる', ''),
    ('1278', '存', 'ソン、ゾン', ''),
    ('1279', '村', 'ソン、むら', ''),
    ('1280', '孫', 'ソン、まご', ''),
    ('1281', '尊', 'ソン、たっと-い、とうと-い、たっと-ぶ、とうと-ぶ', ''),
    ('1282', '損', 'ソン、そこ-なう、そこ-ねる', ''),
    ('1283', '遜', 'ソン', ''),
    ('1284', '他', 'タ、ほか', ''),
    ('1285', '多', 'タ、おお-い', ''),
    ('1286', '汰', 'タ', ''),
    ('1287', '打', 'ダ、う-つ', ''),
    ('1288', '妥', 'ダ', ''),
    ('1289', '唾', 'ダ、つば', '固唾、唾'),
    ('1290', '堕墮', 'ダ', ''),
    ('1291', '惰', 'ダ', ''),
    ('1292', '駄', 'ダ', ''),
    ('1293', '太', 'タイ、タ、ふと-い、ふと-る', '太刀'),
    ('1294', '対對', 'タイ、ツイ', ''),
    ('1295', '体體', 'タイ、テイ、からだ', ''),
    ('1296', '耐', 'タイ、た-える', ''),
    ('1297', '待', 'タイ、ま-つ', ''),
    ('1298', '怠', 'タイ、おこた-る、なま-ける', ''),
    ('1299', '胎', 'タイ', ''),
    ('1300', '退', 'タイ、しりぞ-く、しりぞ-ける', '立ち退く'),
    ('1301', '帯帶', 'タイ、お-びる、おび', ''),
    ('1302', '泰', 'タイ', ''),
    ('1303', '堆', 'タイ', ''),
    ('1304', '袋', 'タイ、ふくろ', '足袋'),
    ('1305', '逮', 'タイ', ''),
    ('1306', '替', 'タイ、か-える、か-わる', '為替'),
    ('1307', '貸', 'タイ、か-す', ''),
    ('1308', '隊', 'タイ', ''),
    ('1309', '滞滯', 'タイ、とどこお-る', ''),
    ('1310', '態', 'タイ', ''),
    ('1311', '戴', 'タイ', ''),
    ('1312', '大', 'ダイ、タイ、おお、おお-きい、おお-いに', '大人、大和、大阪、大分'),
    ('1313', '代', 'ダイ、タイ、か-わる、か-える、よ、しろ', ''),
    ('1314', '台臺', 'ダイ、タイ', ''),
    ('1315', '第', 'ダイ', ''),
    ('1316', '題', 'ダイ', ''),
    ('1317', '滝瀧', 'たき', ''),
    ('1318', '宅', 'タク', ''),
    ('1319', '択擇', 'タク', ''),
    ('1320', '沢澤', 'タク、さわ', ''),
    ('1321', '卓', 'タク', ''),
    ('1322', '拓', 'タク', ''),
    ('1323', '託', 'タク', ''),
    ('1324', '濯', 'タク', ''),
    ('1325', '諾', 'ダク', ''),
    ('1326', '濁', 'ダク、にご-る、にご-す', ''),
    ('1327', '但', 'ただ-し', ''),
    ('1328', '達', 'タツ', '友達'),
    ('1329', '脱脫', 'ダツ、ぬ-ぐ、ぬ-げる', ''),
    ('1330', '奪', 'ダツ、うば-う', ''),
    ('1331', '棚', 'たな', ''),
    ('1332', '誰', 'だれ', ''),
    ('1333', '丹', 'タン', ''),
    ('1334', '旦', 'タン、ダン', ''),
    ('1335', '担擔', 'タン、かつ-ぐ、にな-う', ''),
    ('1336', '単單', 'タン', ''),
    ('1337', '炭', 'タン、すみ', ''),
    ('1338', '胆膽', 'タン', ''),
    ('1339', '探', 'タン、さぐ-る、さが-す', ''),
    ('1340', '淡', 'タン、あわ-い', ''),
    ('1341', '短', 'タン、みじか-い', ''),
    ('1342', '嘆嘆', 'タン、なげ-く、なげ-かわしい', ''),
    ('1343', '端', 'タン、はし、は、はた', ''),
    ('1344', '綻', 'タン、ほころ-びる', ''),
    ('1345', '誕', 'タン', ''),
    ('1346', '鍛', 'タン、きた-える', '鍛冶'),
    ('1347', '団團', 'ダン、（トン）', ''),
    ('1348', '男', 'ダン、ナン、おとこ', ''),
    ('1349', '段', 'ダン', ''),
    ('1350', '断斷', 'ダン、た-つ、ことわ-る', ''),
    ('1351', '弾彈', 'ダン、ひ-く、はず-む、たま', ''),
    ('1352', '暖', 'ダン、あたた-か、あたた-かい、あたた-まる、あたた-める', ''),
    ('1353', '談', 'ダン', ''),
    ('1354', '壇', 'ダン、（タン）', ''),
    ('1355', '地', 'チ、ジ', '意気地、心地'),
    ('1356', '池', 'チ、いけ', ''),
    ('1357', '知', 'チ、し-る', ''),
    ('1358', '値', 'チ、ね、あたい', ''),
    ('1359', '恥', 'チ、は-じる、はじ、は-じらう、は-ずかしい', ''),
    ('1360', '致', 'チ、いた-す', ''),
    ('1361', '遅遲', 'チ、おく-れる、おく-らす、おそ-い', ''),
    ('1362', '痴癡', 'チ', ''),
    ('1363', '稚', 'チ', '稚児'),
    ('1364', '置', 'チ、お-く', ''),
    ('1365', '緻', 'チ', ''),
    ('1366', '竹', 'チク、たけ', '竹刀'),
    ('1367', '畜', 'チク', ''),
    ('1368', '逐', 'チク', ''),
    ('1369', '蓄', 'チク、たくわ-える', ''),
    ('1370', '築', 'チク、きず-く', '築山'),
    ('1371', '秩', 'チツ', ''),
    ('1372', '窒', 'チツ', ''),
    ('1373', '茶', 'チャ、サ', ''),
    ('1374', '着', 'チャク、（ジャク）、き-る、き-せる、つ-く、つ-ける', ''),
    ('1375', '嫡', 'チャク', ''),
    ('1376', '中', 'チュウ、（ジュウ）、なか', ''),
    ('1377', '仲', 'チュウ、なか', '仲人'),
    ('1378', '虫蟲', 'チュウ、むし', ''),
    ('1379', '沖', 'チュウ、おき', ''),
    ('1380', '宙', 'チュウ', ''),
    ('1381', '忠', 'チュウ', ''),
    ('1382', '抽', 'チュウ', ''),
    ('1383', '注', 'チュウ、そそ-ぐ', ''),
    ('1384', '昼晝', 'チュウ、ひる', ''),
    ('1385', '柱', 'チュウ、はしら', ''),
    ('1386', '衷', 'チュウ', ''),
    ('1387', '酎', 'チュウ', ''),
    ('1388', '鋳鑄', 'チュウ、い-る', ''),
    ('1389', '駐', 'チュウ', ''),
    ('1390', '著著', 'チョ、あらわ-す、いちじる-しい', ''),
    ('1391', '貯', 'チョ', ''),
    ('1392', '丁', 'チョウ、テイ', ''),
    ('1393', '弔', 'チョウ、とむら-う', ''),
    ('1394', '庁廳', 'チョウ', ''),
    ('1395', '兆', 'チョウ、きざ-す、きざ-し', ''),
    ('1396', '町', 'チョウ、まち', ''),
    ('1397', '長', 'チョウ、なが-い', '八百長'),
    ('1398', '挑', 'チョウ、いど-む', ''),
    ('1399', '帳', 'チョウ', '蚊帳'),
    ('1400', '張', 'チョウ、は-る', ''),
    ('1401', '彫', 'チョウ、ほ-る', ''),
    ('1402', '眺', 'チョウ、なが-める', ''),
    ('1403', '釣', 'チョウ、つ-る', ''),
    ('1404', '頂', 'チョウ、いただ-く、いただき', ''),
    ('1405', '鳥', 'チョウ、とり', '鳥取'),
    ('1406', '朝', 'チョウ、あさ', '今朝'),
    ('1407', '貼', 'チョウ、は-る', '貼付'),
    ('1408', '超', 'チョウ、こ-える、こ-す', ''),
    ('1409', '腸', 'チョウ', ''),
    ('1410', '跳', 'チョウ、は-ねる、と-ぶ', ''),
    ('1411', '徴徵', 'チョウ', ''),
    ('1412', '嘲', 'チョウ、あざけ-る', ''),
    ('1413', '潮', 'チョウ、しお', ''),
    ('1414', '澄', 'チョウ、す-む、す-ます', ''),
    ('1415', '調', 'チョウ、しら-べる、ととの-う、ととの-える', ''),
    ('1416', '聴聽', 'チョウ、き-く', ''),
    ('1417', '懲懲', 'チョウ、こ-りる、こ-らす、こ-らしめる', ''),
    ('1418', '直', 'チョク、ジキ、ただ-ちに、なお-す、なお-る', ''),
    ('1419', '勅', 'チョク', ''),
    ('1420', '捗', 'チョク', ''),
    ('1421', '沈', 'チン、しず-む、しず-める', ''),
    ('1422', '珍', 'チン、めずら-しい', ''),
    ('1423', '朕', 'チン', ''),
    ('1424', '陳', 'チン', ''),
    ('1425', '賃', 'チン', ''),
    ('1426', '鎮鎭', 'チン、しず-める、しず-まる', ''),
    ('1427', '追', 'ツイ、お-う', ''),
    ('1428', '椎', 'ツイ', ''),
    ('1429', '墜', 'ツイ', ''),
    ('1430', '通', 'ツウ、（ツ）、とお-る、とお-す、かよ-う', ''),
    ('1431', '痛', 'ツウ、いた-い、いた-む、いた-める', ''),
    ('1432', '塚塚', 'つか', ''),
    ('1433', '漬', 'つ-ける、つ-かる', ''),
    ('1434', '坪', 'つぼ', ''),
    ('1435', '爪', 'つめ、（つま）', ''),
    ('1436', '鶴', 'つる', ''),
    ('1437', '低', 'テイ、ひく-い、ひく-める、ひく-まる', ''),
    ('1438', '呈', 'テイ', ''),
    ('1439', '廷', 'テイ', ''),
    ('1440', '弟', 'テイ、（ダイ）、（デ）、おとうと', ''),
    ('1441', '定', 'テイ、ジョウ、さだ-める、さだ-まる、さだ-か', ''),
    ('1442', '底', 'テイ、そこ', ''),
    ('1443', '抵', 'テイ', ''),
    ('1444', '邸', 'テイ', ''),
    ('1445', '亭', 'テイ', ''),
    ('1446', '貞', 'テイ', ''),
    ('1447', '帝', 'テイ', ''),
    ('1448', '訂', 'テイ', ''),
    ('1449', '庭', 'テイ、にわ', ''),
    ('1450', '逓遞', 'テイ', ''),
    ('1451', '停', 'テイ', ''),
    ('1452', '偵', 'テイ', ''),
    ('1453', '堤', 'テイ、つつみ', ''),
    ('1454', '提', 'テイ、さ-げる', ''),
    ('1455', '程', 'テイ、ほど', ''),
    ('1456', '艇', 'テイ', ''),
    ('1457', '締', 'テイ、し-まる、し-める', ''),
    ('1458', '諦', 'テイ、あきら-める', ''),
    ('1459', '泥', 'デイ、どろ', ''),
    ('1460', '的', 'テキ、まと', ''),
    ('1461', '笛', 'テキ、ふえ', ''),
    ('1462', '摘', 'テキ、つ-む', ''),
    ('1463', '滴', 'テキ、しずく、したた-る', ''),
    ('1464', '適', 'テキ', ''),
    ('1465', '敵', 'テキ、かたき', ''),
    ('1466', '溺', 'デキ、おぼ-れる', ''),
    ('1467', '迭', 'テツ', ''),
    ('1468', '哲', 'テツ', ''),
    ('1469', '鉄鐵', 'テツ', ''),
    ('1470', '徹', 'テツ', ''),
    ('1471', '撤', 'テツ', ''),
    ('1472', '天', 'テン、あめ、（あま）', ''),
    ('1473', '典', 'テン', ''),
    ('1474', '店', 'テン、みせ', ''),
    ('1475', '点點', 'テン', '合点'),
    ('1476', '展', 'テン', ''),
    ('1477', '添', 'テン、そ-える、そ-う', ''),
    ('1478', '転轉', 'テン、ころ-がる、ころ-げる、ころ-がす、ころ-ぶ', ''),
    ('1479', '塡(填)', 'テン', ''),
    ('1480', '田', 'デン、た', '田舎'),
    ('1481', '伝傳', 'デン、つた-わる、つた-える、つた-う', '手伝う、伝馬船'),
    ('1482', '殿', 'デン、テン、との、どの', ''),
    ('1483', '電', 'デン', ''),
    ('1484', '斗', 'ト', ''),
    ('1485', '吐', 'ト、は-く', ''),
    ('1486', '妬', 'ト、ねた-む', ''),
    ('1487', '徒', 'ト', ''),
    ('1488', '途', 'ト', ''),
    ('1489', '都都', 'ト、ツ、みやこ', ''),
    ('1490', '渡', 'ト、わた-る、わた-す', ''),
    ('1491', '塗', 'ト、ぬ-る', ''),
    ('1492', '賭', 'ト、か-ける', ''),
    ('1493', '土', 'ド、ト、つち', '土産'),
    ('1494', '奴', 'ド', ''),
    ('1495', '努', 'ド、つと-める', ''),
    ('1496', '度', 'ド、（ト）、（タク）、たび', ''),
    ('1497', '怒', 'ド、いか-る、おこ-る', ''),
    ('1498', '刀', 'トウ、かたな', '竹刀、太刀'),
    ('1499', '冬', 'トウ、ふゆ', ''),
    ('1500', '灯燈', 'トウ、ひ', ''),
    ('1501', '当當', 'トウ、あ-たる、あ-てる', ''),
    ('1502', '投', 'トウ、な-げる', '投網'),
    ('1503', '豆', 'トウ、（ズ）、まめ', '小豆'),
    ('1504', '東', 'トウ、ひがし', ''),
    ('1505', '到', 'トウ', ''),
    ('1506', '逃', 'トウ、に-げる、に-がす、のが-す、のが-れる', ''),
    ('1507', '倒', 'トウ、たお-れる、たお-す', ''),
    ('1508', '凍', 'トウ、こお-る、こご-える', ''),
    ('1509', '唐', 'トウ、から', ''),
    ('1510', '島', 'トウ、しま', '鹿児島'),
    ('1511', '桃', 'トウ、もも', ''),
    ('1512', '討', 'トウ、う-つ', ''),
    ('1513', '透', 'トウ、す-く、す-かす、す-ける', ''),
    ('1514', '党黨', 'トウ', ''),
    ('1515', '悼', 'トウ、いた-む', ''),
    ('1516', '盗盜', 'トウ、ぬす-む', ''),
    ('1517', '陶', 'トウ', ''),
    ('1518', '塔', 'トウ', ''),
    ('1519', '搭', 'トウ', ''),
    ('1520', '棟', 'トウ、むね、（むな）', ''),
    ('1521', '湯', 'トウ、ゆ', ''),
    ('1522', '痘', 'トウ', ''),
    ('1523', '登', 'トウ、ト、のぼ-る', ''),
    ('1524', '答', 'トウ、こた-える、こた-え', ''),
    ('1525', '等', 'トウ、ひと-しい', ''),
    ('1526', '筒', 'トウ、つつ', ''),
    ('1527', '統', 'トウ、す-べる', ''),
    ('1528', '稲稻', 'トウ、いね、（いな）', ''),
    ('1529', '踏', 'トウ、ふ-む、ふ-まえる', ''),
    ('1530', '糖', 'トウ', ''),
    ('1531', '頭', 'トウ、ズ、（ト）、あたま、かしら', ''),
    ('1532', '謄', 'トウ', ''),
    ('1533', '藤', 'トウ、ふじ', ''),
    ('1534', '闘鬭鬪', 'トウ、たたか-う', ''),
    ('1535', '騰', 'トウ', ''),
    ('1536', '同', 'ドウ、おな-じ', ''),
    ('1537', '洞', 'ドウ、ほら', ''),
    ('1538', '胴', 'ドウ', ''),
    ('1539', '動', 'ドウ、うご-く、うご-かす', ''),
    ('1540', '堂', 'ドウ', ''),
    ('1541', '童', 'ドウ、わらべ', ''),
    ('1542', '道', 'ドウ、（トウ）、みち', ''),
    ('1543', '働', 'ドウ、はたら-く', ''),
    ('1544', '銅', 'ドウ', ''),
    ('1545', '導', 'ドウ、みちび-く', ''),
    ('1546', '瞳', 'ドウ、ひとみ', ''),
    ('1547', '峠', 'とうげ', ''),
    ('1548', '匿', 'トク', ''),
    ('1549', '特', 'トク', ''),
    ('1550', '得', 'トク、え-る、う-る', ''),
    ('1551', '督', 'トク', ''),
    ('1552', '徳德', 'トク', ''),
    ('1553', '篤', 'トク', ''),
    ('1554', '毒', 'ドク', ''),
    ('1555', '独獨', 'ドク、ひと-り', ''),
    ('1556', '読讀', 'ドク、トク、（トウ）、よ-む', '読経'),
    ('1557', '栃', '（とち）', '栃木'),
    ('1558', '凸', 'トツ', '凸凹'),
    ('1559', '突突', 'トツ、つ-く', ''),
    ('1560', '届屆', 'とど-ける、とど-く', ''),
    ('1561', '屯', 'トン', ''),
    ('1562', '豚', 'トン、ぶた', ''),
    ('1563', '頓', 'トン', ''),
    ('1564', '貪', 'ドン、むさぼ-る', ''),
    ('1565', '鈍', 'ドン、にぶ-い、にぶ-る', ''),
    ('1566', '曇', 'ドン、くも-る', ''),
    ('1567', '丼', 'どんぶり、（どん）', ''),
    ('1568', '那', 'ナ', ''),
    ('1569', '奈', 'ナ', '神奈川、奈良'),
    ('1570', '内內', 'ナイ、（ダイ）、うち', ''),
    ('1571', '梨', 'なし', ''),
    ('1572', '謎', 'なぞ', ''),
    ('1573', '鍋', 'なべ', ''),
    ('1574', '南', 'ナン、（ナ）、みなみ', ''),
    ('1575', '軟', 'ナン、やわ-らか、やわ-らかい', ''),
    ('1576', '難難', 'ナン、かた-い、むずか-しい', '難しい'),
    ('1577', '二', 'ニ、ふた、ふた-つ', '十重二十重、二十、二十歳、二十日、二人、二日'),
    ('1578', '尼', 'ニ、あま', ''),
    ('1579', '弐貳', 'ニ', ''),
    ('1580', '匂', 'にお-う', ''),
    ('1581', '肉', 'ニク', ''),
    ('1582', '虹', 'にじ', ''),
    ('1583', '日', 'ニチ、ジツ、ひ、か', '明日、昨日、今日、一日、二十日、日和、二日、七日'),
    ('1584', '入', 'ニュウ、い-る、い-れる、はい-る', ''),
    ('1585', '乳', 'ニュウ、ちち、ち', '乳母'),
    ('1586', '尿', 'ニョウ', ''),
    ('1587', '任', 'ニン、まか-せる、まか-す', ''),
    ('1588', '妊', 'ニン', ''),
    ('1589', '忍', 'ニン、しの-ぶ、しの-ばせる', ''),
    ('1590', '認', 'ニン、みと-める', ''),
    ('1591', '寧', 'ネイ', ''),
    ('1592', '熱', 'ネツ、あつ-い', ''),
    ('1593', '年', 'ネン、とし', '今年'),
    ('1594', '念', 'ネン', ''),
    ('1595', '捻', 'ネン', ''),
    ('1596', '粘', 'ネン、ねば-る', ''),
    ('1597', '燃', 'ネン、も-える、も-やす、も-す', ''),
    ('1598', '悩惱', 'ノウ、なや-む、なや-ます', ''),
    ('1599', '納', 'ノウ、（ナッ）、（ナ）、（ナン）、（トウ）、おさ-める、おさ-まる', ''),
    ('1600', '能', 'ノウ', '堪能'),
    ('1601', '脳腦', 'ノウ', ''),
    ('1602', '農', 'ノウ', ''),
    ('1603', '濃', 'ノウ、こ-い', ''),
    ('1604', '把', 'ハ', ''),
    ('1605', '波', 'ハ、なみ', '波止場'),
    ('1606', '派', 'ハ', ''),
    ('1607', '破', 'ハ、やぶ-る、やぶ-れる', ''),
    ('1608', '覇霸', 'ハ', ''),
    ('1609', '馬', 'バ、うま、（ま）', '伝馬船'),
    ('1610', '婆', 'バ', ''),
    ('1611', '罵', 'バ、ののし-る', ''),
    ('1612', '拝拜', 'ハイ、おが-む', ''),
    ('1613', '杯', 'ハイ、さかずき', ''),
    ('1614', '背', 'ハイ、せ、せい、そむ-く、そむ-ける', ''),
    ('1615', '肺', 'ハイ', ''),
    ('1616', '俳', 'ハイ', ''),
    ('1617', '配', 'ハイ、くば-る', ''),
    ('1618', '排', 'ハイ', ''),
    ('1619', '敗', 'ハイ、やぶ-れる', ''),
    ('1620', '廃廢', 'ハイ、すた-れる、すた-る', ''),
    ('1621', '輩', 'ハイ', ''),
    ('1622', '売賣', 'バイ、う-る、う-れる', ''),
    ('1623', '倍', 'バイ', ''),
    ('1624', '梅梅', 'バイ、うめ', '梅雨'),
    ('1625', '培', 'バイ、つちか-う', ''),
    ('1626', '陪', 'バイ', ''),
    ('1627', '媒', 'バイ', ''),
    ('1628', '買', 'バイ、か-う', ''),
    ('1629', '賠', 'バイ', ''),
    ('1630', '白', 'ハク、ビャク、しろ、（しら）、しろ-い', '白髪'),
    ('1631', '伯', 'ハク', '伯父、伯母'),
    ('1632', '拍', 'ハク、（ヒョウ）', ''),
    ('1633', '泊', 'ハク、と-まる、と-める', ''),
    ('1634', '迫', 'ハク、せま-る', ''),
    ('1635', '剝(剥)', 'ハク、は-がす、は-ぐ、は-がれる、は-げる', ''),
    ('1636', '舶', 'ハク', ''),
    ('1637', '博', 'ハク、（バク）', '博士'),
    ('1638', '薄', 'ハク、うす-い、うす-める、うす-まる、うす-らぐ、うす-れる', ''),
    ('1639', '麦麥', 'バク、むぎ', ''),
    ('1640', '漠', 'バク', ''),
    ('1641', '縛', 'バク、しば-る', ''),
    ('1642', '爆', 'バク', ''),
    ('1643', '箱', 'はこ', ''),
    ('1644', '箸', 'はし', ''),
    ('1645', '畑', 'はた、はたけ', ''),
    ('1646', '肌', 'はだ', ''),
    ('1647', '八', 'ハチ、や、や-つ、やっ-つ、（よう）', '八百長、八百屋'),
    ('1648', '鉢', 'ハチ、（ハツ）', ''),
    ('1649', '発發', 'ハツ、ホツ', ''),
    ('1650', '髪髮', 'ハツ、かみ', '白髪'),
    ('1651', '伐', 'バツ', ''),
    ('1652', '抜拔', 'バツ、ぬ-く、ぬ-ける、ぬ-かす、ぬ-かる', ''),
    ('1653', '罰', 'バツ、バチ', ''),
    ('1654', '閥', 'バツ', ''),
    ('1655', '反', 'ハン、（ホン）、（タン）、そ-る、そ-らす', ''),
    ('1656', '半', 'ハン、なか-ば', ''),
    ('1657', '氾', 'ハン', ''),
    ('1658', '犯', 'ハン、おか-す', ''),
    ('1659', '帆', 'ハン、ほ', ''),
    ('1660', '汎', 'ハン', ''),
    ('1661', '伴', 'ハン、バン、ともな-う', ''),
    ('1662', '判', 'ハン、バン', ''),
    ('1663', '坂', 'ハン、さか', ''),
    ('1664', '阪', 'ハン', '大阪'),
    ('1665', '板', 'ハン、バン、いた', ''),
    ('1666', '版', 'ハン', ''),
    ('1667', '班', 'ハン', ''),
    ('1668', '畔', 'ハン', ''),
    ('1669', '般', 'ハン', ''),
    ('1670', '販', 'ハン', ''),
    ('1671', '斑', 'ハン', ''),
    ('1672', '飯飯', 'ハン、めし', ''),
    ('1673', '搬', 'ハン', ''),
    ('1674', '煩', 'ハン、（ボン）、わずら-う、わずら-わす', ''),
    ('1675', '頒', 'ハン', ''),
    ('1676', '範', 'ハン', ''),
    ('1677', '繁繁', 'ハン', ''),
    ('1678', '藩', 'ハン', ''),
    ('1679', '晩晚', 'バン', ''),
    ('1680', '番', 'バン', ''),
    ('1681', '蛮蠻', 'バン', ''),
    ('1682', '盤', 'バン', ''),
    ('1683', '比', 'ヒ、くら-べる', ''),
    ('1684', '皮', 'ヒ、かわ', ''),
    ('1685', '妃', 'ヒ', ''),
    ('1686', '否', 'ヒ、いな', ''),
    ('1687', '批', 'ヒ', ''),
    ('1688', '彼', 'ヒ、かれ、（かの）', ''),
    ('1689', '披', 'ヒ', ''),
    ('1690', '肥', 'ヒ、こ-える、こえ、こ-やす、こ-やし', ''),
    ('1691', '非', 'ヒ', ''),
    ('1692', '卑卑', 'ヒ、いや-しい、いや-しむ、いや-しめる', ''),
    ('1693', '飛', 'ヒ、と-ぶ、と-ばす', ''),
    ('1694', '疲', 'ヒ、つか-れる', ''),
    ('1695', '秘祕', 'ヒ、ひ-める', ''),
    ('1696', '被', 'ヒ、こうむ-る', ''),
    ('1697', '悲', 'ヒ、かな-しい、かな-しむ', ''),
    ('1698', '扉', 'ヒ、とびら', ''),
    ('1699', '費', 'ヒ、つい-やす、つい-える', ''),
    ('1700', '碑碑', 'ヒ', ''),
    ('1701', '罷', 'ヒ', ''),
    ('1702', '避', 'ヒ、さ-ける', ''),
    ('1703', '尾', 'ビ、お', '尻尾'),
    ('1704', '眉', 'ビ、（ミ）、まゆ', ''),
    ('1705', '美', 'ビ、うつく-しい', ''),
    ('1706', '備', 'ビ、そな-える、そな-わる', ''),
    ('1707', '微', 'ビ', ''),
    ('1708', '鼻', 'ビ、はな', ''),
    ('1709', '膝', 'ひざ', ''),
    ('1710', '肘', 'ひじ', ''),
    ('1711', '匹', 'ヒツ、ひき', ''),
    ('1712', '必', 'ヒツ、かなら-ず', ''),
    ('1713', '泌', 'ヒツ、ヒ', '分泌'),
    ('1714', '筆', 'ヒツ、ふで', ''),
    ('1715', '姫姬', 'ひめ', ''),
    ('1716', '百', 'ヒャク', '八百長、八百屋'),
    ('1717', '氷', 'ヒョウ、こおり、ひ', ''),
    ('1718', '表', 'ヒョウ、おもて、あらわ-す、あらわ-れる', ''),
    ('1719', '俵', 'ヒョウ、たわら', ''),
    ('1720', '票', 'ヒョウ', ''),
    ('1721', '評', 'ヒョウ', ''),
    ('1722', '漂', 'ヒョウ、ただよ-う', ''),
    ('1723', '標', 'ヒョウ', ''),
    ('1724', '苗', 'ビョウ、なえ、（なわ）', '早苗'),
    ('1725', '秒', 'ビョウ', ''),
    ('1726', '病', 'ビョウ、（ヘイ）、や-む、やまい', ''),
    ('1727', '描', 'ビョウ、えが-く、か-く', ''),
    ('1728', '猫', 'ビョウ、ねこ', ''),
    ('1729', '品', 'ヒン、しな', ''),
    ('1730', '浜濱', 'ヒン、はま', ''),
    ('1731', '貧', 'ヒン、ビン、まず-しい', ''),
    ('1732', '賓賓', 'ヒン', ''),
    ('1733', '頻頻', 'ヒン', ''),
    ('1734', '敏敏', 'ビン', ''),
    ('1735', '瓶甁', 'ビン', ''),
    ('1736', '不', 'フ、ブ', ''),
    ('1737', '夫', 'フ、（フウ）、おっと', ''),
    ('1738', '父', 'フ、ちち', '叔父、伯父、父さん'),
    ('1739', '付', 'フ、つ-ける、つ-く', '貼付'),
    ('1740', '布', 'フ、ぬの', '昆布'),
    ('1741', '扶', 'フ', ''),
    ('1742', '府', 'フ', ''),
    ('1743', '怖', 'フ、こわ-い', ''),
    ('1744', '阜', '（フ）', '岐阜'),
    ('1745', '附', 'フ', ''),
    ('1746', '訃', 'フ', ''),
    ('1747', '負', 'フ、ま-ける、ま-かす、お-う', ''),
    ('1748', '赴', 'フ、おもむ-く', ''),
    ('1749', '浮', 'フ、う-く、う-かれる、う-かぶ、う-かべる', '浮気、浮つく'),
    ('1750', '婦', 'フ', ''),
    ('1751', '符', 'フ', ''),
    ('1752', '富', 'フ、（フウ）、と-む、とみ', '富山、富貴'),
    ('1753', '普', 'フ', ''),
    ('1754', '腐', 'フ、くさ-る、くさ-れる、くさ-らす', ''),
    ('1755', '敷', 'フ、し-く', '桟敷'),
    ('1756', '膚', 'フ', ''),
    ('1757', '賦', 'フ', ''),
    ('1758', '譜', 'フ', ''),
    ('1759', '侮侮', 'ブ、あなど-る', ''),
    ('1760', '武', 'ブ、ム', ''),
    ('1761', '部', 'ブ', '部屋'),
    ('1762', '舞', 'ブ、ま-う、まい', ''),
    ('1763', '封', 'フウ、ホウ', ''),
    ('1764', '風', 'フウ、（フ）、かぜ、（かざ）', '風邪'),
    ('1765', '伏', 'フク、ふ-せる、ふ-す', ''),
    ('1766', '服', 'フク', ''),
    ('1767', '副', 'フク', ''),
    ('1768', '幅', 'フク、はば', ''),
    ('1769', '復', 'フク', ''),
    ('1770', '福福', 'フク', ''),
    ('1771', '腹', 'フク、はら', ''),
    ('1772', '複', 'フク', ''),
    ('1773', '覆', 'フク、おお-う、くつがえ-す、くつがえ-る', ''),
    ('1774', '払拂', 'フツ、はら-う', ''),
    ('1775', '沸', 'フツ、わ-く、わ-かす', ''),
    ('1776', '仏佛', 'ブツ、ほとけ', ''),
    ('1777', '物', 'ブツ、モツ、もの', '果物'),
    ('1778', '粉', 'フン、こ、こな', ''),
    ('1779', '紛', 'フン、まぎ-れる、まぎ-らす、まぎ-らわす、まぎ-らわしい', ''),
    ('1780', '雰', 'フン', ''),
    ('1781', '噴', 'フン、ふ-く', ''),
    ('1782', '墳', 'フン', ''),
    ('1783', '憤', 'フン、いきどお-る', ''),
    ('1784', '奮', 'フン、ふる-う', ''),
    ('1785', '分', 'ブン、フン、ブ、わ-ける、わ-かれる、わ-かる、わ-かつ', '大分'),
    ('1786', '文', 'ブン、モン、ふみ', '文字'),
    ('1787', '聞', 'ブン、モン、き-く、き-こえる', ''),
    ('1788', '丙', 'ヘイ', ''),
    ('1789', '平', 'ヘイ、ビョウ、たい-ら、ひら', ''),
    ('1790', '兵', 'ヘイ、ヒョウ', ''),
    ('1791', '併倂', 'ヘイ、あわ-せる', ''),
    ('1792', '並竝', 'ヘイ、なみ、なら-べる、なら-ぶ、なら-びに', ''),
    ('1793', '柄', 'ヘイ、がら、え', ''),
    ('1794', '陛', 'ヘイ', ''),
    ('1795', '閉', 'ヘイ、と-じる、と-ざす、し-める、し-まる', ''),
    ('1796', '塀塀', 'ヘイ', ''),
    ('1797', '幣', 'ヘイ', ''),
    ('1798', '弊', 'ヘイ', ''),
    ('1799', '蔽', 'ヘイ', ''),
    ('1800', '餅餠', 'ヘイ、もち', ''),
    ('1801', '米', 'ベイ、マイ、こめ', ''),
    ('1802', '壁', 'ヘキ、かべ', ''),
    ('1803', '璧', 'ヘキ', ''),
    ('1804', '癖', 'ヘキ、くせ', ''),
    ('1805', '別', 'ベツ、わか-れる', ''),
    ('1806', '蔑', 'ベツ、さげす-む', ''),
    ('1807', '片', 'ヘン、かた', ''),
    ('1808', '辺邊', 'ヘン、あた-り、べ', ''),
    ('1809', '返', 'ヘン、かえ-す、かえ-る', ''),
    ('1810', '変變', 'ヘン、か-わる、か-える', ''),
    ('1811', '偏', 'ヘン、かたよ-る', ''),
    ('1812', '遍', 'ヘン', ''),
    ('1813', '編', 'ヘン、あ-む', ''),
    ('1814', '弁辨瓣辯', 'ベン', ''),
    ('1815', '便', 'ベン、ビン、たよ-り', ''),
    ('1816', '勉勉', 'ベン', ''),
    ('1817', '歩步', 'ホ、ブ、（フ）、ある-く、あゆ-む', ''),
    ('1818', '保', 'ホ、たも-つ', ''),
    ('1819', '哺', 'ホ', ''),
    ('1820', '捕', 'ホ、と-らえる、と-らわれる、と-る、つか-まえる、つか-まる', ''),
    ('1821', '補', 'ホ、おぎな-う', ''),
    ('1822', '舗舖', 'ホ', '老舗'),
    ('1823', '母', 'ボ、はは', '乳母、叔母、伯母、母屋、母家、母さん'),
    ('1824', '募', 'ボ、つの-る', ''),
    ('1825', '墓', 'ボ、はか', ''),
    ('1826', '慕', 'ボ、した-う', ''),
    ('1827', '暮', 'ボ、く-れる、く-らす', ''),
    ('1828', '簿', 'ボ', ''),
    ('1829', '方', 'ホウ、かた', '行方'),
    ('1830', '包', 'ホウ、つつ-む', ''),
    ('1831', '芳', 'ホウ、かんば-しい', ''),
    ('1832', '邦', 'ホウ', ''),
    ('1833', '奉', 'ホウ、（ブ）、たてまつ-る', ''),
    ('1834', '宝寶', 'ホウ、たから', ''),
    ('1835', '抱', 'ホウ、だ-く、いだ-く、かか-える', ''),
    ('1836', '放', 'ホウ、はな-す、はな-つ、はな-れる、ほう-る', ''),
    ('1837', '法', 'ホウ、（ハッ）、（ホッ）', ''),
    ('1838', '泡', 'ホウ、あわ', ''),
    ('1839', '胞', 'ホウ', ''),
    ('1840', '俸', 'ホウ', ''),
    ('1841', '倣', 'ホウ、なら-う', ''),
    ('1842', '峰', 'ホウ、みね', ''),
    ('1843', '砲', 'ホウ', ''),
    ('1844', '崩', 'ホウ、くず-れる、くず-す', '雪崩'),
    ('1845', '訪', 'ホウ、おとず-れる、たず-ねる', ''),
    ('1846', '報', 'ホウ、むく-いる', ''),
    ('1847', '蜂', 'ホウ、はち', ''),
    ('1848', '豊豐', 'ホウ、ゆた-か', ''),
    ('1849', '飽', 'ホウ、あ-きる、あ-かす', ''),
    ('1850', '褒襃', 'ホウ、ほ-める', ''),
    ('1851', '縫', 'ホウ、ぬ-う', ''),
    ('1852', '亡', 'ボウ、（モウ）、な-い', ''),
    ('1853', '乏', 'ボウ、とぼ-しい', ''),
    ('1854', '忙', 'ボウ、いそが-しい', ''),
    ('1855', '坊', 'ボウ、（ボッ）', ''),
    ('1856', '妨', 'ボウ、さまた-げる', ''),
    ('1857', '忘', 'ボウ、わす-れる', ''),
    ('1858', '防', 'ボウ、ふせ-ぐ', ''),
    ('1859', '房', 'ボウ、ふさ', ''),
    ('1860', '肪', 'ボウ', ''),
    ('1861', '某', 'ボウ', ''),
    ('1862', '冒', 'ボウ、おか-す', ''),
    ('1863', '剖', 'ボウ', ''),
    ('1864', '紡', 'ボウ、つむ-ぐ', ''),
    ('1865', '望', 'ボウ、モウ、のぞ-む', ''),
    ('1866', '傍', 'ボウ、かたわ-ら', ''),
    ('1867', '帽', 'ボウ', ''),
    ('1868', '棒', 'ボウ', ''),
    ('1869', '貿', 'ボウ', ''),
    ('1870', '貌', 'ボウ', ''),
    ('1871', '暴', 'ボウ、（バク）、あば-く、あば-れる', ''),
    ('1872', '膨', 'ボウ、ふく-らむ、ふく-れる', ''),
    ('1873', '謀', 'ボウ、（ム）、はか-る', ''),
    ('1874', '頰(頬)', 'ほお', ''),
    ('1875', '北', 'ホク、きた', ''),
    ('1876', '木', 'ボク、モク、き、（こ）', '木綿'),
    ('1877', '朴', 'ボク', ''),
    ('1878', '牧', 'ボク、まき', ''),
    ('1879', '睦', 'ボク', ''),
    ('1880', '僕', 'ボク', ''),
    ('1881', '墨墨', 'ボク、すみ', ''),
    ('1882', '撲', 'ボク', '相撲'),
    ('1883', '没沒', 'ボツ', ''),
    ('1884', '勃', 'ボツ', ''),
    ('1885', '堀', 'ほり', ''),
    ('1886', '本', 'ホン、もと', ''),
    ('1887', '奔', 'ホン', ''),
    ('1888', '翻飜', 'ホン、ひるがえ-る、ひるがえ-す', ''),
    ('1889', '凡', 'ボン、（ハン）', ''),
    ('1890', '盆', 'ボン', ''),
    ('1891', '麻', 'マ、あさ', ''),
    ('1892', '摩', 'マ', ''),
    ('1893', '磨', 'マ、みが-く', ''),
    ('1894', '魔', 'マ', ''),
    ('1895', '毎每', 'マイ', ''),
    ('1896', '妹', 'マイ、いもうと', ''),
    ('1897', '枚', 'マイ', ''),
    ('1898', '昧', 'マイ', ''),
    ('1899', '埋', 'マイ、う-める、う-まる、う-もれる', ''),
    ('1900', '幕', 'マク、バク', ''),
    ('1901', '膜', 'マク', ''),
    ('1902', '枕', 'まくら', ''),
    ('1903', '又', 'また', ''),
    ('1904', '末', 'マツ、バツ、すえ', ''),
    ('1905', '抹', 'マツ', ''),
    ('1906', '万萬', 'マン、バン', ''),
    ('1907', '満滿', 'マン、み-ちる、み-たす', ''),
    ('1908', '慢', 'マン', ''),
    ('1909', '漫', 'マン', ''),
    ('1910', '未', 'ミ', ''),
    ('1911', '味', 'ミ、あじ、あじ-わう', '三味線'),
    ('1912', '魅', 'ミ', ''),
    ('1913', '岬', 'みさき', ''),
    ('1914', '密', 'ミツ', ''),
    ('1915', '蜜', 'ミツ', ''),
    ('1916', '脈', 'ミャク', ''),
    ('1917', '妙', 'ミョウ', ''),
    ('1918', '民', 'ミン、たみ', ''),
    ('1919', '眠', 'ミン、ねむ-る、ねむ-い', ''),
    ('1920', '矛', 'ム、ほこ', ''),
    ('1921', '務', 'ム、つと-める、つと-まる', ''),
    ('1922', '無', 'ム、ブ、な-い', ''),
    ('1923', '夢', 'ム、ゆめ', ''),
    ('1924', '霧', 'ム、きり', ''),
    ('1925', '娘', 'むすめ', ''),
    ('1926', '名', 'メイ、ミョウ、な', '仮名、名残'),
    ('1927', '命', 'メイ、ミョウ、いのち', ''),
    ('1928', '明',
     'メイ、ミョウ、あ-かり、あか-るい、あか-るむ、あか-らむ、あき-らか、あ-ける、あ-く、あ-くる、あ-かす',
     '明日'),
    ('1929', '迷', 'メイ、まよ-う', '迷子'),
    ('1930', '冥', 'メイ、ミョウ', ''),
    ('1931', '盟', 'メイ', ''),
    ('1932', '銘', 'メイ', ''),
    ('1933', '鳴', 'メイ、な-く、な-る、な-らす', ''),
    ('1934', '滅', 'メツ、ほろ-びる、ほろ-ぼす', ''),
    ('1935', '免免', 'メン、まぬか-れる', '免れる'),
    ('1936', '面', 'メン、おも、おもて、つら', '真面目'),
    ('1937', '綿', 'メン、わた', '木綿'),
    ('1938', '麺麵', 'メン', ''),
    ('1939', '茂', 'モ、しげ-る', ''),
    ('1940', '模', 'モ、ボ', ''),
    ('1941', '毛', 'モウ、け', ''),
    ('1942', '妄', 'モウ、ボウ', ''),
    ('1943', '盲', 'モウ', ''),
    ('1944', '耗', 'モウ、（コウ）', ''),
    ('1945', '猛', 'モウ', '猛者'),
    ('1946', '網', 'モウ、あみ', '投網'),
    ('1947', '目', 'モク、（ボク）、め、（ま）', '真面目'),
    ('1948', '黙默', 'モク、だま-る', ''),
    ('1949', '門', 'モン、かど', ''),
    ('1950', '紋', 'モン', ''),
    ('1951', '問', 'モン、と-う、と-い、（とん）', '問屋'),
    ('1952', '冶', 'ヤ', '鍛冶'),
    ('1953', '夜', 'ヤ、よ、よる', ''),
    ('1954', '野', 'ヤ、の', '野良'),
    ('1955', '弥彌', 'や', '弥生'),
    ('1956', '厄', 'ヤク', ''),
    ('1957', '役', 'ヤク、エキ', ''),
    ('1958', '約', 'ヤク', ''),
    ('1959', '訳譯', 'ヤク、わけ', ''),
    ('1960', '薬藥', 'ヤク、くすり', ''),
    ('1961', '躍', 'ヤク、おど-る', ''),
    ('1962', '闇', 'やみ', ''),
    ('1963', '由', 'ユ、ユウ、（ユイ）、よし', ''),
    ('1964', '油', 'ユ、あぶら', ''),
    ('1965', '喩', 'ユ', ''),
    ('1966', '愉', 'ユ', ''),
    ('1967', '諭', 'ユ、さと-す', ''),
    ('1968', '輸', 'ユ', ''),
    ('1969', '癒', 'ユ、い-える、い-やす', ''),
    ('1970', '唯', 'ユイ、（イ）', ''),
    ('1971', '友', 'ユウ、とも', '友達'),
    ('1972', '有', 'ユウ、ウ、あ-る', ''),
    ('1973', '勇', 'ユウ、いさ-む', ''),
    ('1974', '幽', 'ユウ', ''),
    ('1975', '悠', 'ユウ', ''),
    ('1976', '郵', 'ユウ', ''),
    ('1977', '湧', 'ユウ、わ-く', ''),
    ('1978', '猶', 'ユウ', ''),
    ('1979', '裕', 'ユウ', ''),
    ('1980', '遊', 'ユウ、（ユ）、あそ-ぶ', ''),
    ('1981', '雄', 'ユウ、お、おす', ''),
    ('1982', '誘', 'ユウ、さそ-う', ''),
    ('1983', '憂', 'ユウ、うれ-える、うれ-い、う-い', ''),
    ('1984', '融', 'ユウ', ''),
    ('1985', '優', 'ユウ、やさ-しい、すぐ-れる', ''),
    ('1986', '与與', 'ヨ、あた-える', ''),
    ('1987', '予豫', 'ヨ', ''),
    ('1988', '余餘', 'ヨ、あま-る、あま-す', ''),
    ('1989', '誉譽', 'ヨ、ほま-れ', ''),
    ('1990', '預', 'ヨ、あず-ける、あず-かる', ''),
    ('1991', '幼', 'ヨウ、おさな-い', ''),
    ('1992', '用', 'ヨウ、もち-いる', ''),
    ('1993', '羊', 'ヨウ、ひつじ', ''),
    ('1994', '妖', 'ヨウ、あや-しい', ''),
    ('1995', '洋', 'ヨウ', ''),
    ('1996', '要', 'ヨウ、かなめ、い-る', ''),
    ('1997', '容', 'ヨウ', ''),
    ('1998', '庸', 'ヨウ', ''),
    ('1999', '揚', 'ヨウ、あ-げる、あ-がる', ''),
    ('2000', '揺搖', 'ヨウ、ゆ-れる、ゆ-る、ゆ-らぐ、ゆ-るぐ、ゆ-する、ゆ-さぶる、ゆ-すぶる', ''),
    ('2001', '葉', 'ヨウ、は', '紅葉'),
    ('2002', '陽', 'ヨウ', ''),
    ('2003', '溶', 'ヨウ、と-ける、と-かす、と-く', ''),
    ('2004', '腰', 'ヨウ、こし', ''),
    ('2005', '様樣', 'ヨウ、さま', ''),
    ('2006', '瘍', 'ヨウ', ''),
    ('2007', '踊', 'ヨウ、おど-る、おど-り', ''),
    ('2008', '窯', 'ヨウ、かま', ''),
    ('2009', '養', 'ヨウ、やしな-う', ''),
    ('2010', '擁', 'ヨウ', ''),
    ('2011', '謡謠', 'ヨウ、うたい、うた-う', ''),
    ('2012', '曜', 'ヨウ', ''),
    ('2013', '抑', 'ヨク、おさ-える', ''),
    ('2014', '沃', 'ヨク', ''),
    ('2015', '浴', 'ヨク、あ-びる、あ-びせる', '浴衣'),
    ('2016', '欲', 'ヨク、ほっ-する、ほ-しい', ''),
    ('2017', '翌', 'ヨク', ''),
    ('2018', '翼', 'ヨク、つばさ', ''),
    ('2019', '拉', 'ラ', ''),
    ('2020', '裸', 'ラ、はだか', ''),
    ('2021', '羅', 'ラ', ''),
    ('2022', '来來', 'ライ、く-る、きた-る、きた-す', ''),
    ('2023', '雷', 'ライ、かみなり', ''),
    ('2024', '頼賴', 'ライ、たの-む、たの-もしい、たよ-る', ''),
    ('2025', '絡', 'ラク、から-む、から-まる、から-める', ''),
    ('2026', '落', 'ラク、お-ちる、お-とす', ''),
    ('2027', '酪', 'ラク', ''),
    ('2028', '辣', 'ラツ', ''),
    ('2029', '乱亂', 'ラン、みだ-れる、みだ-す', ''),
    ('2030', '卵', 'ラン、たまご', ''),
    ('2031', '覧覽', 'ラン', ''),
    ('2032', '濫', 'ラン', ''),
    ('2033', '藍', 'ラン、あい', ''),
    ('2034', '欄欄', 'ラン', ''),
    ('2035', '吏', 'リ', ''),
    ('2036', '利', 'リ、き-く', '砂利'),
    ('2037', '里', 'リ、さと', ''),
    ('2038', '理', 'リ', ''),
    ('2039', '痢', 'リ', ''),
    ('2040', '裏', 'リ、うら', ''),
    ('2041', '履', 'リ、は-く', '草履'),
    ('2042', '璃', 'リ', ''),
    ('2043', '離', 'リ、はな-れる、はな-す', ''),
    ('2044', '陸', 'リク', ''),
    ('2045', '立', 'リツ、（リュウ）、た-つ、た-てる', '立ち退く'),
    ('2046', '律', 'リツ、（リチ）', ''),
    ('2047', '慄', 'リツ', ''),
    ('2048', '略', 'リャク', ''),
    ('2049', '柳', 'リュウ、やなぎ', ''),
    ('2050', '流', 'リュウ、（ル）、なが-れる、なが-す', ''),
    ('2051', '留', 'リュウ、（ル）、と-める、と-まる', ''),
    ('2052', '竜龍', 'リュウ、たつ', ''),
    ('2053', '粒', 'リュウ、つぶ', ''),
    ('2054', '隆隆', 'リュウ', ''),
    ('2055', '硫', 'リュウ', '硫黄'),
    ('2056', '侶', 'リョ', ''),
    ('2057', '旅', 'リョ、たび', ''),
    ('2058', '虜虜', 'リョ', ''),
    ('2059', '慮', 'リョ', ''),
    ('2060', '了', 'リョウ', ''),
    ('2061', '両兩', 'リョウ', ''),
    ('2062', '良', 'リョウ、よ-い', '野良、奈良'),
    ('2063', '料', 'リョウ', ''),
    ('2064', '涼', 'リョウ、すず-しい、すず-む', ''),
    ('2065', '猟獵', 'リョウ', ''),
    ('2066', '陵', 'リョウ、みささぎ', ''),
    ('2067', '量', 'リョウ、はか-る', ''),
    ('2068', '僚', 'リョウ', ''),
    ('2069', '領', 'リョウ', ''),
    ('2070', '寮', 'リョウ', ''),
    ('2071', '療', 'リョウ', ''),
    ('2072', '瞭', 'リョウ', ''),
    ('2073', '糧', 'リョウ、（ロウ）、かて', ''),
    ('2074', '力', 'リョク、リキ、ちから', ''),
    ('2075', '緑綠', 'リョク、（ロク）、みどり', ''),
    ('2076', '林', 'リン、はやし', ''),
    ('2077', '厘', 'リン', ''),
    ('2078', '倫', 'リン', ''),
    ('2079', '輪', 'リン、わ', ''),
    ('2080', '隣', 'リン、とな-る、となり', ''),
    ('2081', '臨', 'リン、のぞ-む', ''),
    ('2082', '瑠', 'ル', ''),
    ('2083', '涙淚', 'ルイ、なみだ', ''),
    ('2084', '累', 'ルイ', ''),
    ('2085', '塁壘', 'ルイ', ''),
    ('2086', '類類', 'ルイ、たぐ-い', ''),
    ('2087', '令', 'レイ', ''),
    ('2088', '礼禮', 'レイ、ライ', ''),
    ('2089', '冷', 'レイ、つめ-たい、ひ-える、ひ-や、ひ-やす、ひ-やかす、さ-める、さ-ます', ''),
    ('2090', '励勵', 'レイ、はげ-む、はげ-ます', ''),
    ('2091', '戻戾', 'レイ、もど-す、もど-る', ''),
    ('2092', '例', 'レイ、たと-える', ''),
    ('2093', '鈴', 'レイ、リン、すず', ''),
    ('2094', '零', 'レイ', ''),
    ('2095', '霊靈', 'レイ、リョウ、たま', ''),
    ('2096', '隷', 'レイ', ''),
    ('2097', '齢齡', 'レイ', ''),
    ('2098', '麗', 'レイ、うるわ-しい', ''),
    ('2099', '暦曆', 'レキ、こよみ', ''),
    ('2100', '歴歷', 'レキ', ''),
    ('2101', '列', 'レツ', ''),
    ('2102', '劣', 'レツ、おと-る', ''),
    ('2103', '烈', 'レツ', ''),
    ('2104', '裂', 'レツ、さ-く、さ-ける', ''),
    ('2105', '恋戀', 'レン、こ-う、こい、こい-しい', ''),
    ('2106', '連', 'レン、つら-なる、つら-ねる、つ-れる', ''),
    ('2107', '廉', 'レン', ''),
    ('2108', '練練', 'レン、ね-る', ''),
    ('2109', '錬鍊', 'レン', ''),
    ('2110', '呂', 'ロ', ''),
    ('2111', '炉爐', 'ロ', ''),
    ('2112', '賂', 'ロ', ''),
    ('2113', '路', 'ロ、じ', ''),
    ('2114', '露', 'ロ、（ロウ）、つゆ', ''),
    ('2115', '老', 'ロウ、お-いる、ふ-ける', '老舗'),
    ('2116', '労勞', 'ロウ', ''),
    ('2117', '弄', 'ロウ、もてあそ-ぶ', ''),
    ('2118', '郎郞', 'ロウ', ''),
    ('2119', '朗朗', 'ロウ、ほが-らか', ''),
    ('2120', '浪', 'ロウ', ''),
    ('2121', '廊廊', 'ロウ', ''),
    ('2122', '楼樓', 'ロウ', ''),
    ('2123', '漏', 'ロウ、も-る、も-れる、も-らす', ''),
    ('2124', '籠', 'ロウ、かご、こ-もる', ''),
    ('2125', '六', 'ロク、む、む-つ、むっ-つ、（むい）', ''),
    ('2126', '録錄', 'ロク', ''),
    ('2127', '麓', 'ロク、ふもと', ''),
    ('2128', '論', 'ロン', ''),
    ('2129', '和', 'ワ、（オ）、やわ-らぐ、やわ-らげる、なご-む、なご-やか', '日和、大和'),
    ('2130', '話', 'ワ、はな-す、はなし', ''),
    ('2131', '賄', 'ワイ、まかな-う', ''),
    ('2132', '脇', 'わき', ''),
    ('2133', '惑', 'ワク、まど-う', ''),
    ('2134', '枠', 'わく', ''),
    ('2135', '湾灣', 'ワン', ''),
    ('2136', '腕', 'ワン、うで', ''),
)

TYPEFACES = (
    '1⑴①', '2⑵②', '3⑶③', '4⑷④', '5⑸⑤', '6⑹⑥', '7⑺⑦', '8⑻⑧', '9⑼⑨',
    '印㊞', '有㈲', '株㈱', '社㈳', '財㈶', '学㈻',
    '吉𠮷', '崎﨑嵜', '高髙',
    '頬頰', '侠俠', '巌巖', '桑桒', '桧檜', '槙槇', '祐祐', '祷禱', '禄祿',
    '秦䅈', '穣穰', '第㐧', '蝉蟬', '鴎鷗', '鴬鶯', '脇𦚰',
    # 常用漢字
    '亜亞', '悪惡', '圧壓', '囲圍', '医醫', '為爲', '壱壹', '逸逸', '飲飮',
    '隠隱', '羽羽', '栄榮', '営營', '鋭銳', '衛衞', '益益', '駅驛', '悦悅',
    '謁謁', '閲閱', '円圓', '塩鹽', '縁緣', '艶艷', '応應', '欧歐', '殴毆',
    '桜櫻', '奥奧', '横橫', '温溫', '穏穩', '仮假', '価價', '禍禍', '画畫',
    '会會', '悔悔', '海海', '絵繪', '壊壞', '懐懷', '慨慨', '概槪', '拡擴',
    '殻殼', '覚覺', '学學', '岳嶽', '楽樂', '喝喝', '渇渴', '褐褐', '缶罐',
    '巻卷', '陥陷', '勧勸', '寛寬', '漢漢', '関關', '歓歡', '館館', '観觀',
    '顔顏', '気氣', '祈祈', '既旣', '帰歸', '亀龜', '器器', '偽僞', '戯戲',
    '犠犧', '旧舊', '拠據', '挙擧', '虚虛', '峡峽', '挟挾', '狭狹', '教敎',
    '郷鄕', '響響', '暁曉', '勤勤', '謹謹', '区區', '駆驅駈',  # "駈"を追加
    '勲勳', '薫薰',
    '径徑', '茎莖', '恵惠', '掲揭', '渓溪', '経經', '蛍螢', '軽輕', '継繼',
    '鶏鷄', '芸藝', '撃擊', '欠缺', '研硏', '県縣', '倹儉', '剣劍', '険險',
    '圏圈', '検檢', '献獻', '権權', '顕顯', '験驗', '厳嚴', '戸戶', '呉吳',
    '娯娛', '広廣', '効效', '恒恆', '黄黃', '鉱鑛', '号號', '告吿', '国國',
    '黒黑', '穀穀', '砕碎', '済濟', '斎齋', '歳歲', '剤劑', '冊册', '殺殺',
    '雑雜', '参參', '桟棧', '蚕蠶', '惨慘', '産產', '賛贊', '残殘', '糸絲',
    '祉祉', '視視', '歯齒', '飼飼', '児兒', '辞辭', '𠮟叱', '湿濕', '実實',
    '写寫', '社社', '舎舍', '者者', '煮煮', '釈釋', '寿壽', '収收', '臭臭',
    '従從', '渋澁', '獣獸', '縦縱', '祝祝', '粛肅', '処處', '暑暑', '署署',
    '緒緖', '諸諸', '叙敍敘',  # "敘"を追加
    '尚尙', '将將', '祥祥', '称稱', '渉涉', '焼燒',
    '証證', '奨奬', '条條', '状狀', '乗乘', '浄淨', '剰剩', '畳疊', '縄繩',
    '壌壤', '嬢孃', '譲讓', '醸釀', '触觸', '嘱囑', '神神', '真眞', '寝寢',
    '慎愼', '尽盡', '図圖', '粋粹', '酔醉', '穂穗', '随隨', '髄髓', '枢樞',
    '数數', '瀬瀨', '声聲', '青靑', '斉齊', '清淸', '晴晴', '精精', '静靜',
    '税稅', '窃竊', '摂攝', '節節', '説說', '絶絕', '専專', '浅淺', '戦戰',
    '践踐', '銭錢', '潜潛', '繊纖', '禅禪', '祖祖', '双雙', '壮壯', '争爭',
    '荘莊', '捜搜', '挿揷', '巣巢', '曽曾', '痩瘦', '装裝', '僧僧', '層層',
    '総總', '騒騷', '増增', '憎憎', '蔵藏', '贈贈', '臓臟', '即卽', '属屬',
    '続續', '堕墮', '対對', '体體', '帯帶', '滞滯', '台臺', '滝瀧', '択擇',
    '沢澤', '脱脫', '担擔', '単單', '胆膽', '嘆嘆', '団團', '断斷', '弾彈',
    '遅遲', '痴癡', '虫蟲', '昼晝', '鋳鑄', '著著', '庁廳', '徴徵', '聴聽',
    '懲懲', '鎮鎭', '塚塚', '逓遞', '鉄鐵', '点點', '転轉', '塡填', '伝傳',
    '都都', '灯燈', '当當', '党黨', '盗盜', '稲稻', '闘鬭鬪', '徳德', '独獨',
    '読讀', '突突', '届屆', '内內', '難難', '弐貳', '悩惱', '脳腦', '覇霸',
    '拝拜', '廃廢', '売賣', '梅梅', '剝剥', '麦麥', '発發', '髪髮', '抜拔',
    '飯飯', '繁繁', '晩晚', '蛮蠻', '卑卑', '秘祕', '碑碑', '姫姬',
    '浜濱濵',  # "濵"を追加
    '賓賓', '頻頻', '敏敏', '瓶甁', '侮侮', '福福', '払拂', '仏佛', '併倂',
    '並竝', '塀塀', '餅餠', '辺邊邉',  # "邉"を追加
    '変變', '弁辨瓣辯', '勉勉', '歩步', '舗舖', '宝寶', '豊豐', '褒襃', '頰頬',
    '墨墨', '没沒', '翻飜', '毎每', '万萬', '満滿', '免免', '麺麵', '黙默',
    '弥彌', '訳譯', '薬藥', '与與', '予豫', '余餘', '誉譽', '揺搖', '様樣',
    '謡謠', '来來', '頼賴', '乱亂', '覧覽', '欄欄', '竜龍', '隆隆', '虜虜',
    '両兩', '猟獵', '緑綠', '涙淚', '塁壘', '類類', '礼禮礼',  # "礼"を追加
    '励勵', '戻戾', '霊靈', '齢齡', '暦曆', '歴歷', '恋戀', '練練', '錬鍊',
    '炉爐', '労勞', '郎郞', '朗朗', '廊廊', '楼樓', '録錄', '湾灣',
)

SAMPLE_BASIS = '''
# ★（タイトル）

v=+1.0
### ★（第1項）

★

### ★（第2項）

★
'''

SAMPLE_LAW = '''
v=+0.5 V=+0.5
$ 総則

v=+0.5 V=+0.5
$$ 通則

v=+0.5
: （基本原則）

##
私権は、公共の福祉に適合しなければならない。

###
権利の行使及び義務の履行は、信義に従い誠実に行わなければならない。

###
権利の濫用は、これを許さない。

v=+0.5
: （解釈の基準）

##
この法律は、個人の尊厳と両性の本質的平等を旨として、解釈しなければならない。

v=+0.5 V=+0.5
$$ 人

v=+0.5 V=+0.5
$$$ 権利能力

v=+0.5
##
私権の享有は、出生に始まる。

###
外国人は、法令又は条約の規定により禁止される場合を除き、私権を享有する。

v=+0.5 V=+0.5
$$$ 意思能力

v=+0.5
##-#
法律行為の当事者が意思表示をした時に意思能力を有しなかったときは、
その法律行為は、無効とする。

v=+0.5 V=+0.5
$$$ 行為能力

v=+0.5
: （成年）

##
年齢十八歳をもって、成年とする。

v=+0.5
: （未成年者の法律行為）

##
未成年者が法律行為をするには、その法定代理人の同意を得なければならない。
ただし、単に権利を得、又は義務を免れる法律行為については、この限りでない。

###
前項の規定に反する法律行為は、取り消すことができる。

###
第一項の規定にかかわらず、法定代理人が目的を定めて処分を許した財産は、
その目的の範囲内において、未成年者が自由に処分することができる。
目的を定めないで処分を許した財産を処分するときも、同様とする。

v=+0.5
: （未成年者の営業の許可）

##
一種又は数種の営業を許された未成年者は、その営業に関しては、
成年者と同一の行為能力を有する。

###
前項の場合において、未成年者がその営業に堪えることができない事由があるときは、
その法定代理人は、第四編（親族）の規定に従い、その許可を取り消し、
又はこれを制限することができる。

v=+0.5
: （後見開始の審判）

##
精神上の障害により事理を弁識する能力を欠く常況にある者については、
家庭裁判所は、本人、配偶者、四親等内の親族、未成年後見人、未成年後見監督人、
保佐人、保佐監督人、補助人、補助監督人又は検察官の請求により、
後見開始の審判をすることができる。

v=+0.5
: （成年被後見人及び成年後見人）

##
後見開始の審判を受けた者は、成年被後見人とし、これに成年後見人を付する。

v=+0.5
: （成年被後見人の法律行為）

##
成年被後見人の法律行為は、取り消すことができる。
ただし、日用品の購入その他日常生活に関する行為については、この限りでない。

v=+0.5
: （後見開始の審判の取消し）

##
第七条に規定する原因が消滅したときは、家庭裁判所は、本人、配偶者、
四親等内の親族、後見人（未成年後見人及び成年後見人をいう。以下同じ。）、
後見監督人（未成年後見監督人及び成年後見監督人をいう。以下同じ。）又は
検察官の請求により、後見開始の審判を取り消さなければならない。

v=+0.5
: （保佐開始の審判）

##
精神上の障害により事理を弁識する能力が著しく不十分である者については、
家庭裁判所は、本人、配偶者、四親等内の親族、後見人、後見監督人、補助人、
補助監督人又は検察官の請求により、保佐開始の審判をすることができる。
ただし、第七条に規定する原因がある者については、この限りでない。

v=+0.5
: （被保佐人及び保佐人）

##
保佐開始の審判を受けた者は、被保佐人とし、これに保佐人を付する。

v=+0.5
: （保佐人の同意を要する行為等）

##
被保佐人が次に掲げる行為をするには、その保佐人の同意を得なければならない。
ただし、第九条ただし書に規定する行為については、この限りでない。

####
元本を領収し、又は利用すること。

####
借財又は保証をすること。

####
不動産その他重要な財産に関する権利の得喪を目的とする行為をすること。

####
訴訟行為をすること。

####
贈与、和解又は仲裁合意
（仲裁法（平成十五年法律第百三十八号）第二条第一項に規定する仲裁合意をいう。）
をすること。

####
相続の承認若しくは放棄又は遺産の分割をすること。

####
贈与の申込みを拒絶し、遺贈を放棄し、負担付贈与の申込みを承諾し、
又は負担付遺贈を承認すること。

####
新築、改築、増築又は大修繕をすること。

####
第六百二条に定める期間を超える賃貸借をすること。

####
前各号に掲げる行為を制限行為能力者
（未成年者、成年被後見人、被保佐人及び第十七条第一項の審判を受けた被補助人を
いう。以下同じ。）の法定代理人としてすること。

###
家庭裁判所は、
第十一条本文に規定する者又は保佐人若しくは保佐監督人の請求により、
被保佐人が前項各号に掲げる行為以外の行為をする場合であっても
その保佐人の同意を得なければならない旨の審判をすることができる。
ただし、第九条ただし書に規定する行為については、この限りでない。

###
保佐人の同意を得なければならない行為について、
保佐人が被保佐人の利益を害するおそれがないにもかかわらず同意をしないときは、
家庭裁判所は、被保佐人の請求により、保佐人の同意に代わる許可を与えることができる。

###
保佐人の同意を得なければならない行為であって、
その同意又はこれに代わる許可を得ないでしたものは、取り消すことができる。
'''

SAMPLE_PETITION = '''
# 訴状

v=+0.5
令和★年★月★日 :

v=+0.5
: ★裁判所　御中

v=+0.5
<!--------------------------vv----------------------------vv------------->
: \\　　　　　　　　　　　　　原告★訴訟代理人弁護士　　　　★
: \\　　　　　　　　　　　　　　　同　　　　　　　　　　　　★
: \\　　　　　　　　　　　　　　　同（担当）　　　　　　　　★

v=+0.5
: 〒★－★　広島市★
<!--------------------------vv----------------------------vv------------->
: \\　　　　　　　　　　　　　原告　　　　　　　　　　　　　★
: \\　　　　　　　　　　　　　上記代表者代表取締役　　　　　★

: 〒★－★　広島市★
<!--------------------------vv----------------------------vv------------->
: \\　　　　　　　　　　　　　原告　　　　　　　　　　　　　★
: \\　　　　　　　　　　　　　上記代表者代表取締役　　　　　★

: 〒★－★　広島市★
: \\　　　　　　　★法律事務所（送達場所）
<!--------------------------vv----------------------------vv------------->
: \\　　　　　　　　　　　　　原告★訴訟代理人弁護士　　　　★
: \\　　　　　　　　　　　　　　　同　　　　　　　　　　　　★
: \\　　　　　　　　　　　　　　　同（担当）　　　　　　　　★
: \\　　　　　　　TEL ★－★－★　　FAX ★－★－★

: 〒★－★　広島市★
<!--------------------------vv----------------------------vv------------->
: \\　　　　　　　　　　　　　被告　　　　　　　　　　　　　★
: \\　　　　　　　　　　　　　上記代表者代表取締役　　　　　★

: 〒★－★　広島市★
<!--------------------------vv----------------------------vv------------->
: \\　　　　　　　　　　　　　被告　　　　　　　　　　　　　★
: \\　　　　　　　　　　　　　上記代表者代表取締役　　　　　★

v=+1.0
: ★請求事件
: 訴訟物の価額　　★★★★万★★★★円
: 貼用印紙額　　　　　　★万★★★★円

## 請求の趣旨

###
被告★は、原告に対し、★連帯して、
★円及びこれに対する令和★年★月★日から支払済みまで年３分
の割合による金員を支払え。

###
訴訟費用は被告★の負担とする。

<<=1 <=1
との判決並びに仮執行の宣言を求める。

# 請求の原因

### ★について

★

### ★について

★

### まとめ

よって、原告は、被告★に対し、不法行為に基づき、
損害金★円及びこれに対する本件事故日である
令和★年★月★日から支払済みまで民法所定年３分
の割合による遅延損害金の支払を求める。

v=+1.0
# ##=1 ###=1

: 証拠方法 :

### 甲第１号証　　　　　★
### 甲第２号証　　　　　★
### 甲第３号証　　　　　★
### 甲第４号証　　　　　★
### 甲第５号証　　　　　★
### 甲第６号証　　　　　★
### 甲第７号証　　　　　★
### 甲第８号証　　　　　★
### 甲第９号証　　　　　★
### 甲第１０号証　　　　★
### 甲第１１号証　　　　★

v=+1.0
# ##=1 ###=1

: 附属書類 :

### 訴状副本　　　　　　　　　　　　　　★通<!--[被告の数]-->
### 資格証明書　　　　　　　　　　　　　★通<!--[法人当事者の数]-->
### 訴訟委任状　　　　　　　　　　　　　★通<!--[原告の数]-->
### 甲号証の写し　　　　　　　　　　　各★通<!--[被告の数＋1]-->
'''

SAMPLE_EVIDENCE = '''
: 令和★年（★）第★号　★請求事件
: 原告　★
: 被告　★

v=+0.5
# 証拠説明書

v=+0.5
令和★年★月★日 :

v=+0.5
: ★裁判所　御中

v=+0.5
<!--------------------------vv----------------------------vv------------->
: \\　　　　　　　　　　　　　★★★訴訟代理人弁護士　　　　★
: \\　　　　　　　　　　　　　　　同　　　　　　　　　　　　★
: \\　　　　　　　　　　　　　　　同（担当）　　　　　　　　★

v=+1.0
--
|号証 |標目|原写|作成日|作成者|立証趣旨|備考|
=
|:----|:---------|:--:|:-------|:-----------|:-----------------------|:---------|
|★1|★書|原本|R★.★.★|★|①★であったこと<br>②★であったこと||
|★2|★書|原本|R★.★.★|★|①★であったこと<br>②★であったこと||
|★3|★書|原本|R★.★.★|★|①★であったこと<br>②★であったこと||
--
'''

SAMPLE_SETTLEMENT = '''
# 和解契約書

v=+1.0
★（以下「甲」という。）と
★（以下「乙」という。）は、
★に関し、次のとおり和解した。

## （★）

★

###
★

###
★

## （債務）

乙は、甲に対し、
★万★円の債務を負っていることを認める。

## （支払）

乙は、甲に対し、
令和★年★月★日限り、
前条の★万★円を下記の口座に振り込んで支払う。
ただし、振込手数料は乙の負担とする。

<=-1.0 v=+0.5
金融機関__　　　　　　　　　　　__　本支店名__　　　　　　　　　　　__

<=-1.0 v=+0.5
普通・当座等__　　　　　　　__　口座番号__　　　　　　　　　　　　　__

<=-1.0 v=+0.5
<名義/フリガナ>__　　　　　　　　　　　　　　　　　　　　　　　　　　　　　__

## （清算条項）

甲と乙は、甲と乙の間には、
★本件に関し、
上記各条項に定めるほか、何らの債権債務のないことを相互に確認する。

v=+1.0
本和解の成立を証するため、本書を★通作成し、各自1通を所持するものとする。

# ##=1 ###=1

v=+1.0
: 令和★年★月★日

v=+1.0
: 甲　　　　★
: \\　　　　　　　　　　　★　　　　　　　　　　　　　　　　　　　　㊞

: ★代理人　★
: \\　　　　　　　弁護士　★　　　　　　　　　　　　　　　　　　　　㊞

v=+0.5
: ★　住所　^DDD^__　　　　　　　　　　　　　　　　　　　　　　　　　　　　__^DDD^

v=+0.5
: \\　　氏名　__　　　　　　　　　　　　　　　　　　　　　　　　　　　㊞__
'''

DONT_EDIT_MESSAGE = '<!--【以下は必要なデータですので編集しないでください】-->'

TAB_WIDTH = 4


######################################################################
# FUNCTION


def get_ideal_width(s):
    wid = 0
    for c in s:
        if (c == '\t'):
            wid += (int(wid / TAB_WIDTH) + 1) * TAB_WIDTH
            continue
        w = unicodedata.east_asian_width(c)
        if (w == 'F'):    # Full alphabet ...
            wid += 2
        elif(w == 'H'):   # Half katakana ...
            wid += 1
        elif(w == 'W'):   # Chinese character ...
            wid += 2
        elif(w == 'Na'):  # Half alphabet ...
            wid += 1
        elif(w == 'A'):   # Greek character ...
            wid += 1
        elif(w == 'N'):   # Arabic character ...
            wid += 1
    return wid


def c2n_n_arab(s):
    n = 0
    for c in s:
        n *= 10
        if re.match('^[0-9]$', c):
            n += int(c)
        elif re.match('^[０-９]$', c):
            n += ord(c) - 65296
        else:
            return -1
    return n


def c2n_n_kata(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 65392
    if i >= n + 1 and i <= n + 44:
        # ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜ
        return i - n
    n = 65337
    if i == n + 45:
        # ｦ
        return i - n
    n = 65391
    if i == n + 46:
        # ﾝ
        return i - n
    n = 12448
    if i >= n + 2 * 1 and i <= n + 2 * 5:
        # アイウエオ
        return int((i - n) / 2)
    n = 12447
    if i >= n + 2 * 6 and i <= n + 2 * 17:
        # カキクケコサシスセソタチ
        return int((i - n) / 2)
    n = 12448
    if i >= n + 2 * 18 and i <= n + 2 * 20:
        # ツテト
        return int((i - n) / 2)
    n = 12469
    if i >= n + 1 * 21 and i <= n + 1 * 25:
        # ナニヌネノ
        return int((i - n) / 1)
    n = 12417
    if i >= n + 3 * 26 and i <= n + 3 * 30:
        # ハヒフヘホ
        return int((i - n) / 3)
    n = 12479
    if i >= n + 1 * 31 and i <= n + 1 * 35:
        # マミムメモ
        return int((i - n) / 1)
    n = 12444
    if i >= n + 2 * 36 and i <= n + 2 * 38:
        # ヤユヨ
        return int((i - n) / 2)
    n = 12482
    if i >= n + 1 * 39 and i <= n + 1 * 43:
        # ラリルレロ
        return int((i - n) / 1)
    n = 12483
    if i >= n + 1 * 44 and i <= n + 1 * 49:
        # ワヰヱヲン
        return int((i - n) / 1)
    return -1


def c2n_n_alph(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 96
    if i >= n + 1 and i <= n + 26:
        # a...z
        return i - n
    n = 65344
    if i >= n + 1 and i <= n + 26:
        # ａ...ｚ
        return i - n
    return -1


def c2n_n_kanj(s):
    i = s
    i = re.sub('[０〇零]', '0', i)
    i = re.sub('[１一壱]', '1', i)
    i = re.sub('[２二弐]', '2', i)
    i = re.sub('[３三参]', '3', i)
    i = re.sub('[４四]', '4', i)
    i = re.sub('[５五伍]', '5', i)
    i = re.sub('[６六]', '6', i)
    i = re.sub('[７七]', '7', i)
    i = re.sub('[８八]', '8', i)
    i = re.sub('[９九]', '9', i)
    #
    i = re.sub('[拾]', '十', i)
    i = re.sub('[佰陌]', '百', i)
    i = re.sub('[仟阡]', '千', i)
    i = re.sub('[萬]', '万', i)
    #
    i = re.sub('^([千百十])', '1\\1', i)
    i = re.sub('([^0-9])([千百十])', '\\1 1\\2', i)
    #
    i = re.sub('(万)([^千]*)$', '\\1 0千\\2', i)
    i = re.sub('(千)([^百]*)$', '\\1 0百\\2', i)
    i = re.sub('(百)([^十]*)$', '\\1 0十\\2', i)
    i = re.sub('(十)$', '\\1 0', i)
    #
    i = re.sub('[万千百十 ]', '', i)
    #
    if re.match('^[0-9]+$', i):
        return int(i)
    return -1


def adjust_line(document):
    old = document
    old = re.sub('。', '。\n', old)
    old = re.sub('\n\n+', '\n\n', old)
    old = re.sub('^\n+', '', old)
    old = re.sub('\n+$', '', old)
    new = ''
    tmp = ''
    for sen in old.split('\n'):
        t = sen
        t = re.sub('、', '、\n', t)
        # t = re.sub('を', 'を\n', t)
        t = re.sub('「', '\n「', t)
        t = re.sub('」', '」\n', t)
        t = re.sub('（', '\n（', t)
        t = re.sub('）', '）\n', t)
        for phr in t.split('\n'):
            if get_ideal_width(tmp + phr) > makdo.makdo_docx2md.MD_TEXT_WIDTH:
                new += tmp + '\n'
                tmp = ''
            tmp += phr
        if tmp != '':
            new += tmp
            tmp = ''
        new += '\n'
    new = re.sub('\n+$', '', new)
    return new


######################################################################
# CLASS

############################################################
# SIMPLE DAILOG

class OneWordDialog(tkinter.simpledialog.Dialog):

    def __init__(self, pane, mother, title, prompt, init=''):
        self.pane = pane
        self.mother = mother
        self.prompt = prompt
        self.init = init
        self.value = None
        super().__init__(pane, title=title)

    def body(self, pane):
        font_size = self.mother.font_size.get()
        font = (GOTHIC_FONT, font_size)
        prompt = tkinter.Label(pane, text=self.prompt)
        prompt.pack(side='top', anchor='w')
        self.entry = tkinter.Entry(pane, width=25, font=font)
        self.entry.pack(side='top')
        self.entry.insert(0, self.init)
        super().body(pane)
        return self.entry

    def apply(self):
        self.value = self.entry.get()

    def get_value(self):
        return self.value


class TwoWordsDialog(tkinter.simpledialog.Dialog):

    def __init__(self, pane, mother, title, prompt, init1='', init2=''):
        self.pane = pane
        self.mother = mother
        self.prompt = prompt
        self.init1 = init1
        self.init2 = init2
        self.value1 = None
        self.value2 = None
        super().__init__(pane, title=title)

    def body(self, pane):
        font_size = self.mother.font_size.get()
        font = (GOTHIC_FONT, font_size)
        prompt = tkinter.Label(pane, text=self.prompt)
        prompt.pack(side='top', anchor='w')
        self.entry1 = tkinter.Entry(pane, width=25, font=font)
        self.entry1.pack(side='top')
        self.entry1.insert(0, self.init1)
        self.entry2 = tkinter.Entry(pane, width=25, font=font)
        self.entry2.pack(side='top')
        self.entry2.insert(0, self.init2)
        super().body(pane)
        return self.entry1

    def apply(self):
        self.value1 = self.entry1.get()
        self.value2 = self.entry2.get()

    def get_value(self):
        return self.value1, self.value2


############################################################
# CHARS STATE

class CharsState:

    def __init__(self):
        self.del_or_ins = ''
        self.is_in_comment = False
        self.parentheses = []
        self.has_underline = False
        self.has_specific_font = False
        self.has_frame = False
        self.is_resized = ''
        self.is_stretched = ''
        self.is_length_reviser = False
        self.chapter_depth = 0
        self.section_depth = 0

    def __eq__(self, other):
        if self.del_or_ins != other.del_or_ins:
            return False
        if self.is_in_comment != other.is_in_comment:
            return False
        if self.parentheses != other.parentheses:
            return False
        if self.has_underline != other.has_underline:
            return False
        if self.has_specific_font != other.has_specific_font:
            return False
        if self.has_frame != other.has_frame:
            return False
        if self.is_resized != other.is_resized:
            return False
        if self.is_stretched != other.is_stretched:
            return False
        return True

    def copy(self):
        copy = CharsState()
        copy.del_or_ins = self.del_or_ins
        copy.is_in_comment = self.is_in_comment
        for p in self.parentheses:
            copy.parentheses.append(p)
        copy.has_underline = self.has_underline
        copy.has_specific_font = self.has_specific_font
        copy.has_frame = self.has_frame
        copy.is_resized = self.is_resized
        copy.is_stretched = self.is_stretched
        copy.is_length_reviser = self.is_length_reviser
        copy.chapter_depth = self.chapter_depth
        copy.section_depth = self.section_depth
        # for v in vars(copy):
        #     vars(copy)[v] = vars(self)[v]
        return copy

    def reset_partially(self):
        self.is_length_reviser = False
        self.chapter_depth = 0
        self.section_depth = 0

    def set_is_in_comment(self):
        self.is_in_comment = not self.is_in_comment

    def set_del_or_ins(self, del_or_ins):
        if del_or_ins == 'del':
            if self.del_or_ins == 'del':
                self.del_or_ins = ''
            else:
                self.del_or_ins = 'del'
        if del_or_ins == 'ins':
            if self.del_or_ins == 'ins':
                self.del_or_ins = ''
            else:
                self.del_or_ins = 'ins'

    def toggle_has_underline(self):
        self.has_underline = not self.has_underline

    def toggle_has_specific_font(self):
        self.has_specific_font = not self.has_specific_font

    def attach_or_remove_frame(self, fd):
        if fd == '[|':
            self.has_frame = True
        elif fd == '|]':
            self.has_frame = False

    def set_is_resized(self, fd):
        if fd == '---':
            if self.is_resized == '---':
                self.is_resized = ''
            else:
                self.is_resized = '---'
        elif fd == '--':
            if self.is_resized == '--':
                self.is_resized = ''
            else:
                self.is_resized = '--'
        elif fd == '++':
            if self.is_resized == '++':
                self.is_resized = ''
            else:
                self.is_resized = '++'
        elif fd == '+++':
            if self.is_resized == '+++':
                self.is_resized = ''
            else:
                self.is_resized = '+++'

    def set_is_stretched(self, fd):
        if fd == '>>>':
            if self.is_stretched == '<<<':
                self.is_stretched = ''
            else:
                self.is_stretched = '>>>'
        elif fd == '>>':
            if self.is_stretched == '<<':
                self.is_stretched = ''
            else:
                self.is_stretched = '>>'
        elif fd == '<<':
            if self.is_stretched == '>>':
                self.is_stretched = ''
            else:
                self.is_stretched = '<<'
        elif fd == '<<<':
            if self.is_stretched == '>>>':
                self.is_stretched = ''
            else:
                self.is_stretched = '<<<'

    def apply_parenthesis(self, parenthesis):
        ps = self.parentheses
        p = parenthesis
        if p == '「' or p == '『' or p == '[' or p == '（' or p == '(':
            ps.append(p)
        if p == ')' or p == '）' or p == ']' or p == '』' or p == '」':
            if len(ps) > 0:
                if ps[-1] == '(' and p == ')' or \
                   ps[-1] == '（' and p == '）' or \
                   ps[-1] == '[' and p == ']' or \
                   ps[-1] == '『' and p == '』' or \
                   ps[-1] == '「' and p == '」':
                    ps.pop(-1)

    def set_chapter_depth(self, depth):
        self.chapter_depth = depth

    def set_section_depth(self, depth):
        self.section_depth = depth

    def get_key(self, chars):
        key = 'c'
        # ANGLE
        if False:
            pass
        elif chars == ' ':
            return 'hsp_tag'
        elif chars == '\u3000':
            return 'fsp_tag'
        elif chars == '\t':
            return 'tab_tag'
        elif self.is_in_comment:
            key += '-0'
        elif chars == 'font decorator':
            key += '-120'
        elif chars == 'table':
            key += '-190'
        elif chars == 'half number':
            key += '-30'
        elif chars == 'full number':
            key += '-330'
        elif chars == 'list':
            key += '-330'
        elif chars == 'alignment':
            key += '-180'
        elif chars == 'image':
            if len(self.parentheses) == 0:
                key += '-80'
            elif len(self.parentheses) == 1:
                key += '-120'
            elif len(self.parentheses) >= 2:
                key += '-160'
        elif len(self.parentheses) == 1:
            key += '-80'
        elif len(self.parentheses) == 2:
            key += '-120'
        elif len(self.parentheses) >= 3:
            key += '-160'
        elif chars == '<br>' or chars == '<pgbr>' or chars == 'hline':
            key += '-270'
        elif chars == 'R' or chars == 'red':
            key += '-0'
        elif chars == 'Y' or chars == 'yellow':
            key += '-60'
        elif chars == 'G' or chars == 'green':
            key += '-120'
        elif chars == 'C' or chars == 'cyan':
            key += '-180'
        elif chars == 'B' or chars == 'blue':
            key += '-240'
        elif chars == 'M' or chars == 'magenta':
            key += '-300'
        elif chars == 'fold':
            key += '-10'
        elif self.is_length_reviser:
            key += '-150'
        elif self.chapter_depth > 0:
            key += '-' + str(210 + ((self.chapter_depth - 1) * 10))
        elif self.section_depth > 0:
            key += '-' + str(30 + ((self.section_depth - 1) * 10))
        else:
            key += '-XXX'
        # LIGHTNESS
        if self.del_or_ins == 'del':
            key += '-0'
        elif self.del_or_ins == 'ins':
            key += '-2'
        else:
            key += '-1'
        # FONT
        if chars == 'mincho':
            key += '-m'  # mincho
        else:
            key += '-g'  # gothic
        # UNDERLINE
        if chars == 'font decorator':
            key += '-x'  # no underline
        elif chars == 'table':
            key += '-x'  # no underline
        elif chars == ' ' or chars == '\t' or chars == '\u3000':
            # if not self.is_in_comment:
            key += '-u'  # underline
        elif not self.is_in_comment and self.has_underline:
            key += '-u'  # underline
        elif not self.is_in_comment and self.has_specific_font:
            key += '-u'  # specific font
        elif not self.is_in_comment and self.has_frame:
            key += '-u'  # frame
        elif not self.is_in_comment and self.is_resized != '':
            key += '-u'  # resized
        elif not self.is_in_comment and self.is_stretched != '':
            key += '-u'  # stretched
        else:
            key += '-x'  # no underline
        # RETURN
        return key

############################################################
# LINE DATUM


class LineDatum:

    def __init__(self):
        self.line_number = 0
        self.line_text = ''
        self.beg_chars_state = CharsState()
        self.end_chars_state = CharsState()
        self.paint_keywords = False

    def paint_line(self, txt, paint_keywords=False):
        # PREPARE
        i = self.line_number
        line_text = self.line_text
        chars_state = self.beg_chars_state.copy()
        self.paint_keywords = paint_keywords
        # RESET TAG
        for tag in txt.tag_names():
            if tag != 'search_tag':
                txt.tag_remove(tag, str(i + 1) + '.0', str(i + 1) + '.end')
        if line_text == '':
            self.end_chars_state = chars_state.copy()
            return
        if not chars_state.is_in_comment:
            # PAGE BREAK
            if line_text == '<pgbr>\n':
                beg, end = str(i + 1) + '.0', str(i + 1) + '.end'
                key = chars_state.get_key('<pgbr>')                     # 1.key
                #                                                       # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                #                                                       # 5.tmp
                #                                                       # 6.beg
                self.end_chars_state = chars_state.copy()
                return
            # HORIZONTAL LINE
            if re.match('^-{5,}\n$', line_text):
                beg, end = str(i + 1) + '.0', str(i + 1) + '.end'
                key = chars_state.get_key('hline')                      # 1.key
                #                                                       # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                #                                                       # 5.tmp
                #                                                       # 6.beg
                self.end_chars_state = chars_state.copy()
                return
            # LENGTH REVISERS
            res = '^((<<|<|>|v|V|X)=(\\+|\\-)?[\\.0-9]+\\s+)+$'
            if re.match(res, line_text):
                chars_state.is_length_reviser = True
            # CHAPTER
            if line_text[0] == '$':
                res = '^(\\${,5})(?:-\\$+)*(=[\\.0-9]+)?(?:\\s.*)?\n?$'
                if re.match(res, line_text):
                    dep = len(re.sub(res, '\\1', line_text))
                    chars_state.set_chapter_depth(dep)
            # SECTION
            if line_text[0] == '#':
                res = '^(#{,8})(?:-#+)*(=[\\.0-9]+)?(?:\\s.*)?\n?$'
                if re.match(res, line_text):
                    dep = len(re.sub(res, '\\1', line_text))
                    chars_state.set_section_depth(dep)
        # LOOP
        beg, tmp = str(i + 1) + '.0', ''
        for j, c in enumerate(line_text):
            tmp += c
            s1 = line_text[j - 0:j + 1] if True else ''
            s2 = line_text[j - 1:j + 1] if j > 0 else ''
            s3 = line_text[j - 2:j + 1] if j > 1 else ''
            s4 = line_text[j - 3:j + 1] if j > 2 else ''
            c0 = line_text[j + 1] if j < len(line_text) - 1 else ''
            c1 = c
            c2 = line_text[j - 1] if j > 0 else ''
            c3 = line_text[j - 2] if j > 1 else ''
            c4 = line_text[j - 3] if j > 2 else ''
            c5 = line_text[j - 4] if j > 3 else ''
            s_lft = line_text[:j + 1]
            s_rgt = line_text[j + 1:]
            # BEGINNING OF COMMENT "<!--"
            if (not chars_state.is_in_comment and s4 == '<!--') and \
               (c5 != '\\' or re.match(NOT_ESCAPED + '<!--$', tmp)):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 3)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.set_is_in_comment()                         # 4.set
                tmp = '<!--'                                            # 5.tmp
                beg = end                                               # 6.beg
                continue
            # END OF COMMENT "-->"
            if (chars_state.is_in_comment and s3 == '-->') and \
               (c4 != '\\' or re.match(NOT_ESCAPED + '-->$', tmp)):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.set_is_in_comment()                         # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # LIST
            if not chars_state.is_in_comment and j == 0 and \
               c == '-' and re.match('\\s', c0):
                key = chars_state.get_key('list')                       # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            if not chars_state.is_in_comment and j == 1 and \
               re.match('^[0-9]$', c2) and c == '.' and re.match('\\s', c0):
                key = chars_state.get_key('half number')
                txt.tag_remove(key, str(i + 1) + '.0', str(i + 1) + '.1')
                beg, end = str(i + 1) + '.0', str(i + 1) + '.' + str(j + 1)
                key = chars_state.get_key('list')                       # 1.key
                #                                                       # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # ALIGNMENT
            if not chars_state.is_in_comment and j == 0 and \
               c == ':' and re.match('\\s', c0):
                key = chars_state.get_key('alignment')                  # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            if not chars_state.is_in_comment and j >= 2 and \
               re.match('\\s', c3) and c2 == ':' and c == '\n':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 2)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = ' :\n'                                          # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('alignment')                  # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # DEL ("->", "<-")
            if ((chars_state.del_or_ins == '' and s2 == '->' and
                 (c3 != '\\' or re.match(NOT_ESCAPED + '\\->$', tmp))) or
                (chars_state.del_or_ins == 'del' and s2 == '<-' and
                 (c3 != '\\' or re.match(NOT_ESCAPED + '<\\-$', tmp)))):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.set_del_or_ins('del')                       # 4.set
                # tmp = '->' or '<-'                                    # 5.tmp
                beg = end                                               # 6.beg
                key = 'c-20-1-g-x'                                      # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # INS ("+>", "<+")
            if ((chars_state.del_or_ins == '' and s2 == '+>' and
                 (c3 != '\\' or re.match(NOT_ESCAPED + '\\+>$', tmp))) or
                (chars_state.del_or_ins == 'ins' and s2 == '<+' and
                 (c3 != '\\' or re.match(NOT_ESCAPED + '<\\+$', tmp)))):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.set_del_or_ins('ins')                       # 4.set
                # tmp = '+>' or '<+'                                    # 5.tmp
                beg = end                                               # 6.beg
                key = 'c-200-1-g-x'                                     # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # LINE BREAK
            if (not chars_state.is_in_comment) and re.match('^.*<br>$', tmp):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 3)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = <br>                                            # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('<br>')                       # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # RELAX
            if (not chars_state.is_in_comment) and \
               re.match('^.*<>$', tmp):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '<>'                                            # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # SPACE
            if ((re.match('^.*<\\s*[0-9]+$', s_lft) and
                 re.match('^[0-9]*\\s*>.*$', s_rgt)) or
                (re.match('^.*<\\s*[0-9]+$', s_lft) and
                 re.match('^[0-9]*\\.[0-9]+\\s*>.*$', s_rgt)) or
                (re.match('^.*<\\s*[0-9]*\\.$', s_lft) and
                 re.match('^[0-9]+\\s*>.*$', s_rgt)) or
                (re.match('^.*<\\s*[0-9]*\\.[0-9]+$', s_lft) and
                 re.match('^[0-9]*\\s*>.*$', s_rgt))):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '[0-9]'                                         # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            if re.match('^.*<$', s_lft) and \
               re.match('^\\s*[\\.0-9]+\\s*>.*$', s_rgt):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '<'                                             # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            if re.match('^.*<\\s*[\\.0-9]+\\s*>$', s_lft):
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # COLOR
            res_color = '(R|red|Y|yellow|G|green|C|cyan|B|blue|M|magenta)'
            if ((not chars_state.is_in_comment) and
                (re.match('^.*_' + res_color + '_$', tmp) or
                 re.match('^.*\\^' + res_color + '\\^$', tmp))):
                res = '^(.*)[_\\^]' + res_color + '[_\\^]$'
                mdt = re.sub(res, '\\1', tmp)
                col = re.sub(res, '\\2', tmp)
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - len(col) - 1)          # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '_.+_' or '^.+^'                                # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key(col)                          # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # TABLE CONFIGURE
            if (c == ':' and (c2 == '|' or c2 == '-' or c2 == ':')) or \
               (c == ':' and c0 == '|') or \
               ((c == '^' or c == '=') and c2 == '-'):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = ':'                                             # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # FONT DECORATOR ("---", "+++", ">>>", "<<<")
            if (not chars_state.is_in_comment) and \
               (s3 == '---' or s3 == '+++' or s3 == '>>>' or s3 == '<<<') and \
               (c4 != '\\' or re.match(NOT_ESCAPED + '...$', tmp)):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 2)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = re.sub('^(.*)(...)$', '\\2', tmp)                 # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                res1, res2 = '^.*:-+$', '^-*:.*$'
                if not re.match(res1, s_lft) and not re.match(res2, s_rgt):
                    if tmp == '---' or tmp == '+++':
                        chars_state.set_is_resized(tmp)                 # 4.set
                    else:
                        chars_state.set_is_stretched(tmp)               # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # FONT DECORATOR ("--", "++", ">>", "<<")
            if (not chars_state.is_in_comment) and \
               (s2 == '--' or s2 == '++' or s2 == '>>' or s2 == '<<') and \
               (c0 != c1) and \
               (c3 != '\\' or re.match(NOT_ESCAPED + '..$', tmp)):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = re.sub('^(.*)(..)$', '\\2', tmp)                  # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                res1, res2 = '^.*:-+$', '^-*:.*$'
                if not re.match(res1, s_lft) and not re.match(res2, s_rgt):
                    res = '^=[-\\+]?[0-9]*(\\.?[0-9]+)(\\s.*)?$'
                    if s2 != '<<' or not re.match(res, s_rgt):
                        if tmp == '--' or tmp == '++':
                            chars_state.set_is_resized(tmp)             # 4.set
                        else:
                            chars_state.set_is_stretched(tmp)           # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # FONT DECORATOR ("@.+@", "^.*^", "_.*_")
            if ((re.match('^.*@[0-9]+$', s_lft) and
                 re.match('^[0-9]*@.*$', s_rgt)) or
                (re.match('^.*@[0-9]+$', s_lft) and
                 re.match('^[0-9]*\\.[0-9]+@.*$', s_rgt)) or
                (re.match('^.*@[0-9]*\\.$', s_lft) and
                 re.match('^[0-9]+@.*$', s_rgt)) or
                (re.match('^.*@[0-9]*\\.[0-9]+$', s_lft) and
                 re.match('^[0-9]*@.*$', s_rgt))):
                continue  # @n@
            res = NOT_ESCAPED + '(@[^@]{1,66}@|\\^.*\\^|_.*_)$'
            if re.match(res, tmp) and not chars_state.is_in_comment:
                mdt = re.sub(res, '\\2', tmp)
                hul = chars_state.has_underline
                hsf = chars_state.has_specific_font
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - len(mdt) + 1)          # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                if re.match('_.*_', mdt) and hul:
                    chars_state.toggle_has_underline()                  # 4.set
                elif re.match('@.*@', mdt) and hsf:
                    chars_state.toggle_has_specific_font()              # 4.set
                tmp = mdt                                               # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                if re.match('_.*_', mdt) and not hul:
                    chars_state.toggle_has_underline()                  # 4.set
                elif re.match('@.*@', mdt) and not hsf:
                    chars_state.toggle_has_specific_font()              # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # FRAME
            if (c == '[' and c0 == '|') or (c == '|' and c0 == ']'):
                continue
            if (c2 == '[' and c == '|') or (c2 == '|' and c == ']'):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '[|' or '|]'                                    # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.attach_or_remove_frame(c2 + c)              # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # TABLE
            if c == '|':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '|'                                             # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('table')                      # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # IMAGE
            if c == '!' and re.match('^\\[.*\\]\\(.*\\)', line_text[j+1:]):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '!'                                             # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('image')                      # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
            # FOLDING
            if re.match('^#+(-#+)*(\\s.*)?\\.\\.\\.\\[$', s_lft) and \
               re.match(NOT_ESCAPED + '\\.\\.\\.\\[$', s_lft) and \
               re.match('^[0-9]+\\]$', s_rgt):
                continue  # # xxx...[ / n]
            if re.match('^\\.\\.\\.\\[$', s_lft) and \
               re.match('^[0-9]+\\]#+(-#+)*(\\s.*)?$', s_rgt):
                continue  # ...[ / n]# xxx
            if re.match('^#+(-#+)*(\\s.*)?\\.\\.\\.\\[[0-9]+$', s_lft) and \
               re.match(NOT_ESCAPED + '\\.\\.\\.\\[[0-9]+$', s_lft) and \
               re.match('^[0-9]*\\]$', s_rgt):
                continue  # # xxx...[n / ]
            if re.match('^\\.\\.\\.\\[[0-9]+$', s_lft) and \
               re.match('^[0-9]*\\]#+(-#+)*(\\s.*)?$', s_rgt):
                continue  # ...[n / ]xxx
            res = '^(#+(?:-#+)*(?:\\s.*)?)(\\.\\.\\.\\[[0-9]+\\])$'
            if re.match(res, s_lft) and \
               re.match(NOT_ESCAPED + '\\.\\.\\.\\[[0-9]+\\]$', s_lft) and \
               re.match('^\n$', s_rgt):
                fld = re.sub(res, '\\2', s_lft)
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j + 1 - len(fld))          # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '...[n]'                                        # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('fold')                       # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue  # xxx...[n] /
            if re.match('^\\.\\.\\.\\[[0-9]+\\]$', s_lft) and \
               re.match('^#+(-#+)*(\\s.*)?\n$', s_rgt):
                key = chars_state.get_key('fold')                       # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue  # ...[n]# xxx /
            # PARENTHESES
            if c == '「' or c == '『' or c == '[' or c == '（' or c == '(':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.apply_parenthesis(c)                        # 4.set
                tmp = c                                                 # 5.tmp
                beg = end                                               # 6.beg
                continue
            if c == ')' or c == '）' or c == ']' or c == '』' or c == '」':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.apply_parenthesis(c)                        # 4.set
                # tmp = ''                                              # 5.tmp
                beg = end                                               # 6.beg
                continue
            # NUMBER
            if re.match('[0-9]', c):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '[0-9]'                                         # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('half number')                # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            if re.match('[' +
                        '０-９' +
                        '零一二三四五六七八九十' +
                        '⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇' +
                        '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳' +
                        ']', c):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '[０-９...]'                                    # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('full number')                # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # ERROR ("★")
            if c == '★':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '★'                                            # 5.tmp
                beg = end                                               # 6.beg
                key = 'error_tag'                                       # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # MINCHO
            # 002D "-" HYPHEN-MINUS
            # 2010 "‐" HYPHEN
            # 2014 "—" EM DASH
            # 2015 "―" HORIZONTAL BAR
            # 2212 "−" MINUS SIGN
            # 30FC "ー" KATAKANA-HIRAGANA PROLONGED SOUND MARK
            # FF0D "－" FULLWIDTH HYPHEN-MINUS
            if c == '\u2010' or c == '\u2014' or \
               c == '\u2212' or c == '\u30FC':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = c                                               # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('mincho')                     # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # SPACE (" ", "\t", "\u3000")
            if c == ' ' or c == '\t' or c == '\u3000':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = ' ' or '\t' or '\u3000'                         # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key(c)                            # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # SEARCH WORD
            wrd = Makdo.search_word
            if wrd != '' and re.match('^.*' + wrd + '$', tmp):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - len(wrd) + 1)          # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = wrd                                             # 5.tmp
                beg = end                                               # 6.beg
                key = 'rev-gx'                                          # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # KEYWORD
            if self.paint_keywords:
                for kw in KEYWORDS:
                    if re.match('^(.*?)' + kw[0] + '$', tmp):
                        t1 = re.sub('^(.*?)' + kw[0] + '$', '\\1', tmp)
                        t2 = re.sub('^(.*?)' + kw[0] + '$', '\\2', tmp)
                        if t2 == '本訴' or t2 == '反訴' or t2 == '別訴':
                            if re.match('^(原|被)告', s_rgt):
                                continue  # 本訴/原告
                        if t2 == '原告' or t2 == '被告':
                            if re.match('^(?:.|\n)*(本|反|別)訴$', t1):
                                continue  # 本訴原告/
                        if t2 == '被告' and c0 == '人':
                            continue  # 被告/人
                        key = chars_state.get_key('')                   # 1.key
                        end = str(i + 1) + '.' + str(j - len(t2) + 1)   # 2.end
                        txt.tag_add(key, beg, end)                      # 3.tag
                        #                                               # 4.set
                        # tmp = t2                                      # 5.tmp
                        beg = end                                       # 6.beg
                        key = chars_state.get_key(kw[1])                # 1.key
                        end = str(i + 1) + '.' + str(j + 1)             # 2.end
                        txt.tag_add(key, beg, end)                      # 3.tag
                        #                                               # 4.set
                        tmp = ''                                        # 5.tmp
                        beg = end                                       # 6.beg
                if tmp == '':
                    continue
            # END OF THE LINE "\n"
            if c1 == '\n':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                #                                                       # 5.tmp
                #                                                       # 6.beg
                break
        self.end_chars_state = chars_state.copy()
        return


############################################################
# MAKDO

class Makdo:

    args_dont_show_help = None     # True|+False
    file_dont_show_help = None
    args_background_color = None   # +W|B|G
    file_background_color = None
    args_font_size = None          # 3|6|9|12|15|+18|21|24|27|30|33|36|...
    file_font_size = None
    args_paint_keywords = None     # True|+False
    file_paint_keywords = None
    args_digit_separator = None    # +0|3|4
    file_digit_separator = None
    args_read_only = None          # True|+False
    # file_read_only = None
    args_make_backup_file = False  # True|+False
    file_make_backup_file = False

    args_input_file = None

    search_word = ''

    ##############################################
    # INIT

    def __init__(self):
        self.temp_dir = ''
        self.file_path = self.args_input_file
        self.init_text = ''
        self.file_lines = []
        self.has_made_backup_file = False
        self.line_data = []
        self.standard_line = 0
        self.global_line_to_paint = 0
        self.local_line_to_paint = 0
        self.key_history = ['', '', '', '', '', '', '', '', '', '',
                            '', '', '', '', '', '', '', '', '', '', '']
        self.last_position = ''
        self.current_pane = 'txt'
        self.openai_model = 'gpt-3.5-turbo'
        self.openai_key = None
        self.dict_directory = None
        self.formula_number = -1
        self.memo_pad_memory = None
        self.rectangle_text_list = []
        #
        self.must_show_folding_help_message = True
        self.must_show_keyboard_macro_help_message = True
        self.must_show_config_help_message = True
        # GET CONFIGURATION
        self.get_and_set_configurations()
        # WINDOW
        self.win = tkinterdnd2.TkinterDnD.Tk()  # drag and drop
        # self.win = tkinter.Tk()
        self.win.title('MAKDO')
        self.win.geometry(WINDOW_SIZE)
        self.win.protocol("WM_DELETE_WINDOW", self.quit_makdo)
        icon8_img = tkinter.PhotoImage(data=ICON8_IMG, master=self.win)
        self.win.iconphoto(False, icon8_img)
        # FRAME
        # self.frm = tkinter.Frame()
        # self.frm.pack(expand=True, fill=tkinter.BOTH)
        # MENU BAR
        self.mnb = tkinter.Menu(self.win)
        self._make_menu()
        # PANED WINDOW
        self.pnd = tkinter.PanedWindow(self.win, bd=0, sashwidth=3,
                                       orient='vertical')
        self.pnd.pack(expand=True, fill=tkinter.BOTH)
        self.pnd1 = tkinter.PanedWindow(self.pnd, bd=0, bg='#FF5D5D')  # 000
        self.pnd2 = tkinter.PanedWindow(self.pnd, bd=0, bg='#BC7A00')  # 040
        self.pnd3 = tkinter.PanedWindow(self.pnd, bd=0, bg='#758F00')  # 070
        self.pnd4 = tkinter.PanedWindow(self.pnd, bd=0, bg='#00A586')  # 170
        self.pnd5 = tkinter.PanedWindow(self.pnd, bd=0, bg='#7676FF')  # 240
        self.pnd6 = tkinter.PanedWindow(self.pnd, bd=0, bg='#C75DFF')  # 280
        self.pnd.add(self.pnd1)
        # MAIN TEXT
        self.txt = tkinter.Text(self.pnd1, undo=True)
        self.txt.pack(expand=True, fill=tkinter.BOTH)
        self.txt.config(insertbackground='#FF7777', blockcursor=True)  # cursor
        self._make_txt_key_configuration()
        scb = tkinter.Scrollbar(self.txt, orient=tkinter.VERTICAL,
                                command=self.txt.yview)
        scb.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.txt['yscrollcommand'] = scb.set
        self.txt.drop_target_register(tkinterdnd2.DND_FILES)   # drag and drop
        self.txt.dnd_bind('<<Drop>>', self.open_dropped_file)  # drag and drop
        # SUB TEXT
        self.sub = tkinter.Text(self.pnd2, undo=True)
        # self.sub.pack(expand=True, fill=tkinter.BOTH)
        self.sub.config(insertbackground='#FF7777', blockcursor=True)  # cursor
        self._make_sub_key_configuration()
        scb = tkinter.Scrollbar(self.sub, orient=tkinter.VERTICAL,
                                command=self.sub.yview)
        scb.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.sub['yscrollcommand'] = scb.set
        self.sub_btn = tkinter.Button(self.pnd2, text='終了',
                                      command=self._unify_window)
        # STATUS BAR
        self.stbr = tkinter.Frame(self.win)
        self.stbr.pack(side='right', anchor='e')
        self.stb = tkinter.Frame(self.win)
        self.stb.pack(side='left', anchor='w')
        self._make_status_bar()
        # FONT
        self.set_font()
        # OPEN FILE
        if self.args_input_file is not None:
            self.just_open_file(self.args_input_file)
        self.show_first_help_message()
        self.txt.focus_set()
        # RUN PERIODICALLY
        self.run_periodically()
        # LOOP
        self.win.mainloop()

    ####################################
    # TOOLS

    def _get_v_position_of_insert(self):
        insert_position = self.txt.index('insert')
        insert_v_number = int(re.sub('\\.[0-9]+$', '', insert_position))
        return insert_v_number

    def _get_h_position_of_insert(self):
        insert_position = self.txt.index('insert')
        insert_h_number = int(re.sub('^[0-9]+\\.', '', insert_position))
        return insert_h_number

    def _get_max_v_position(self):
        max_position = self.txt.index('end-1c')
        max_v_number = int(re.sub('\\.[0-9]+$', '', max_position))
        return max_v_number

    def _get_max_h_position(self):
        line_end_position = self.txt.index('insert lineend')
        max_h_number = int(re.sub('^[0-9]+\\.', '', line_end_position))
        return max_h_number

    def _get_ideal_h_position_of_insert(self, pane):
        s = pane.get('insert linestart', 'insert')
        return get_ideal_width(s)

    def _execute_when_delete_is_pressed(self, pane):
        if pane.tag_ranges('sel'):
            if self._is_read_only_pane(pane):
                self.copy_region()
            else:
                self.cut_region()
        elif 'akauni' in pane.mark_names():
            akn = pane.index('akauni')
            ins = pane.index('insert')
            beg = re.sub('\\..*$', '.0', ins)
            if akn == ins and akn != beg:
                c = pane.get(beg, akn)
                self.win.clipboard_clear()
                self.win.clipboard_append(c)
                if not self._is_read_only_pane(pane):
                    pane.delete(beg, akn)
                self._cancel_region(pane)
            else:
                if self._is_read_only_pane(pane):
                    self.copy_region()
                else:
                    self.cut_region()
        else:
            ins = pane.index('insert')
            end = re.sub('\\..*$', '.end', ins)
            c = pane.get(ins, end)
            if self._is_read_only_pane(pane):
                self.win.clipboard_clear()
                self.win.clipboard_append(c)
            else:
                if c == '':
                    self.win.clipboard_append('\n')
                    pane.delete(ins, end + '+1c')
                else:
                    self.win.clipboard_append(c)
                    pane.delete(ins, end)

    def paint_out_line(self, line_number):
        ln = line_number
        # REGION IS SET
        if self.txt.tag_ranges('sel'):
            return
        if 'akauni' in self.txt.mark_names():
            return
        # UPDATE TEXT
        file_text = self.txt.get('1.0', 'end-1c')
        self.file_lines = file_text.split('\n')
        m = len(self.file_lines) - 1
        while len(self.line_data) < m + 1:
            self.line_data.append(LineDatum())
            self.line_data[-1].line_number = len(self.line_data) - 1
        while len(self.line_data) > m + 1:
            self.line_data.pop(-1)
        if m < 0:
            return
        # BAD LINE ID
        if ln < 0 or ln >= len(self.line_data):
            return
        # PREPARE
        line_text = self.file_lines[ln] + '\n'
        if ln == 0:
            chars_state = CharsState()
        else:
            chars_state \
                = self.line_data[ln - 1].end_chars_state.copy()
            chars_state.reset_partially()
        paint_keywords = self.paint_keywords.get()
        # EXCLUSION
        # if self.line_data[ln].line_text == line_text and \
        #    self.line_data[ln].beg_chars_state == chars_state and \
        #    self.line_data[ln].paint_keywords == paint_keywords:
        #     return
        # PAINT
        # self.line_data[ln].line_number = ln
        self.line_data[ln].line_text = line_text
        self.line_data[ln].beg_chars_state = chars_state
        self.line_data[ln].end_chars_state = CharsState()
        self.line_data[ln].paint_line(self.txt, paint_keywords)

    @staticmethod
    def _get_now():
        now = datetime.datetime.utcnow() + datetime.timedelta(hours=+9)
        jst = datetime.timezone(datetime.timedelta(hours=+9))
        now = now.replace(tzinfo=jst)
        return now

    @staticmethod
    def _convert_half_to_full(half):
        full = half
        full = re.sub('0', '０', full)
        full = re.sub('1', '１', full)
        full = re.sub('2', '２', full)
        full = re.sub('3', '３', full)
        full = re.sub('4', '４', full)
        full = re.sub('5', '５', full)
        full = re.sub('6', '６', full)
        full = re.sub('7', '７', full)
        full = re.sub('8', '８', full)
        full = re.sub('9', '９', full)
        return full

    @staticmethod
    def _get_encoding(raw_data):
        encoding = 'SHIFT_JIS'
        if raw_data != '':
            encoding = chardet.detect(raw_data)['encoding']
        if encoding is None:
            encoding = 'SHIFT_JIS'
        elif (re.match('^utf[-_]?.*$', encoding, re.I)) or \
             (re.match('^shift[-_]?jis.*$', encoding, re.I)) or \
             (re.match('^cp932.*$', encoding, re.I)) or \
             (re.match('^euc[-_]?(jp|jis).*$', encoding, re.I)) or \
             (re.match('^iso[-_]?2022[-_]?jp.*$', encoding, re.I)) or \
             (re.match('^ascii.*$', encoding, re.I)):
            pass
        else:
            # Windows-1252 (Western Europe)
            # MacCyrillic (Macintosh Cyrillic)
            # ...
            encoding = 'SHIFT_JIS'
            n = '警告'
            m = '文字コードを「SHIFT_JIS」に修正しました'
            tkinter.messagebox.showwarning(n, m)
        return encoding

    @staticmethod
    def _decode_data(encoding, raw_data):
        try:
            decoded_data = raw_data.decode(encoding)
        except BaseException:
            n = 'エラー'
            m = 'データを読みません（テキストでないかも？）'
            tkinter.messagebox.showwarning(n, m)
            raise BaseException('failed to read data')
            return None
        return decoded_data

    def _get_tmp_md(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        md_path = self.temp_dir.name + '/doc.md'
        file_text = self.txt.get('1.0', 'end-1c')
        file_text = self.get_fully_unfolded_document(file_text)
        with open(md_path, 'w') as f:
            f.write(file_text)
        return md_path

    def _get_tmp_docx(self):
        md_path = self._get_tmp_md()
        docx_path = re.sub('md$', 'docx', md_path)
        stderr = sys.stderr
        sys.stderr = tempfile.TemporaryFile(mode='w+')
        importlib.reload(makdo.makdo_md2docx)
        try:
            m2d = makdo.makdo_md2docx.Md2Docx(md_path)
            m2d.save(docx_path)
        except BaseException:
            pass
        sys.stderr.seek(0)
        msg = sys.stderr.read()
        sys.stderr = stderr
        if msg != '':
            n = 'エラー'
            tkinter.messagebox.showerror(n, msg)
            return
        return docx_path

    def _insert_line_break_as_necessary(self):
        t = self.txt.get('1.0', 'insert')
        if len(t) == 0:
            pass
        elif len(t) == 1:
            if t[-1] == '\n':
                pass
            else:
                self.txt.insert('insert', '\n\n')
        elif len(t) >= 2:
            if t[-2] == '\n' and t[-1] == '\n':
                pass
            elif t[-1] == '\n':
                self.txt.insert('insert', '\n')
            else:
                self.txt.insert('insert', '\n\n')
        p = self.txt.index('insert')
        t = self.txt.get('insert', 'end-1c')
        if len(t) == 0:
            self.txt.insert('insert', '\n')
        elif len(t) == 1:
            if t[0] == '\n':
                pass
            else:
                self.txt.insert('insert', '\n\n')
        elif len(t) >= 2:
            if t[0] == '\n' and t[1] == '\n':
                pass
            elif t[0] == '\n':
                self.txt.insert('insert', '\n')
            else:
                self.txt.insert('insert', '\n\n')
        self.txt.mark_set('insert', p)

    def _is_read_only_pane(self, pane):
        if pane == self.sub:
            if self.formula_number <= 0 and self.memo_pad_memory is None:
                return True
            else:
                return False
        else:
            if self.is_read_only.get():
                return True
            else:
                return False

    def _is_region_specified(self, pane):
        if pane.tag_ranges('sel'):
            return True
        elif 'akauni' in pane.mark_names():
            return True
        return False

    def _get_region(self, pane):
        if pane.tag_ranges('sel'):
            beg, end = pane.index('sel.first'), pane.index('sel.last')
        elif 'akauni' in pane.mark_names():
            beg, end = self._get_indices_in_order(pane, 'insert', 'akauni')
        else:
            beg, end = '', ''
        return beg, end

    def _cancel_region(self, pane):
        if pane.tag_ranges('sel'):
            pane.tag_remove('sel', "1.0", "end")
        if 'akauni' in pane.mark_names():
            pane.tag_remove('akauni_tag', '1.0', 'end')
            pane.mark_unset('akauni')

    def _show_no_region_error(self):
        n = 'エラー'
        m = '範囲が指定されていません．'
        tkinter.messagebox.showerror(n, m)

    def _get_indices_in_order(self, pane, index1, index2):
        position1 = pane.index(index1)
        position2 = pane.index(index2)
        p1_v = int(re.sub('\\..+$', '', position1))
        p1_h = int(re.sub('^.+\\.', '', position1))
        p2_v = int(re.sub('\\..+$', '', position2))
        p2_h = int(re.sub('^.+\\.', '', position2))
        if (p1_v < p2_v) or (p1_v == p2_v and p1_h < p2_h):
            return position1, position2
        if (p2_v < p1_v) or (p2_v == p1_v and p2_h < p1_h):
            return position2, position1
        return position1, position2

    ####################################
    # MENU

    def _make_menu(self):
        self._make_menu_file()
        self._make_menu_edit()
        self._make_menu_insert()
        self._make_menu_paragraph()
        self._make_menu_move()
        self._make_menu_tool()
        self._make_menu_configuration()
        self._make_menu_internet()
        self._make_menu_help()
        self.win['menu'] = self.mnb

    ##########################
    # MENU FILE

    def _make_menu_file(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='ファイル(F)', menu=menu, underline=5)
        #
        menu.add_command(label='ファイルを開く(O)', underline=8,
                         command=self.open_file)
        menu.add_command(label='ファイルを閉じる(C)', underline=9,
                         command=self.close_file)
        menu.add_separator()
        #
        menu.add_command(label='ファイルを保存(S)', underline=8,
                         command=self.save_file, accelerator='Ctrl+S')
        menu.add_command(label='Markdown形式で名前を付けて保存(M)', underline=20,
                         command=self.name_and_save_by_md)
        menu.add_command(label='MS Word形式で名前を付けて保存(D)', underline=19,
                         command=self.name_and_save_by_docx)
        # menu.add_command(label='名前を付けて保存(A)', underline=9,
        #                  command=self.name_and_save)
        menu.add_separator()
        #
        menu.add_command(label='ファイル形式を相互に直接変換',
                         command=self.convert_directly)
        menu.add_separator()
        #
        menu.add_command(label='PDFに返還',
                         command=self.convert_to_pdf)
        menu.add_command(label='MS Wordを起動して確認・印刷(P)', underline=18,
                         command=self.start_writer, accelerator='Ctrl+P')
        menu.add_separator()
        #
        menu.add_command(label='終了(Q)', underline=3,
                         command=self.quit_makdo, accelerator='Ctrl+Q')
        # menu.add_separator()

    ################
    # COMMAND

    # OPEN FILE

    def open_file(self):
        ans = self.close_file()
        if ans is None:
            return False
        typ = [('可能な形式', '.md .docx'),
               ('Markdown', '.md'), ('MS Word', '.docx'),
               ('全てのファイル', '*')]
        file_path = tkinter.filedialog.askopenfilename(filetypes=typ)
        if file_path == () or file_path == '':
            return False
        self.just_open_file(file_path)
        return True

    def just_open_file(self, file_path):
        if self.exists_auto_file(file_path):
            self.file_path = ''
            self.init_text = ''
            self.file_lines = []
            return
        # DOCX OR MD
        if re.match('^(?:.|\n)+.docx$', file_path):
            self.temp_dir = tempfile.TemporaryDirectory()
            md_path = self.temp_dir.name + '/doc.md'
        else:
            md_path = file_path
        # OPEN DOCX FILE
        if re.match('^(?:.|\n)+.docx$', file_path):
            stderr = sys.stderr
            sys.stderr = tempfile.TemporaryFile(mode='w+')
            importlib.reload(makdo.makdo_docx2md)
            try:
                d2m = makdo.makdo_docx2md.Docx2Md(file_path)
                d2m.save(md_path)
            except BaseException:
                pass
            sys.stderr.seek(0)
            msg = sys.stderr.read()
            sys.stderr = stderr
            if msg != '':
                n = 'エラー'
                tkinter.messagebox.showerror(n, msg)
                return
        # OPEN MD FILE
        try:
            with open(md_path, 'rb') as f:
                raw_data = f.read()
        except BaseException:
            return
        encoding = self._get_encoding(raw_data)
        try:
            decoded_data = self._decode_data(encoding, raw_data)
        except BaseException:
            self.file_path = None
            return
        init_text = self.get_fully_unfolded_document(decoded_data)
        self.file_path = file_path
        self.init_text = init_text
        self.file_lines = init_text.split('\n')
        # self.txt.delete('1.0', 'end')
        self.txt.insert('1.0', init_text)
        self.txt.focus_set()
        self.txt.mark_set('insert', '1.0')
        file_name = re.sub('^.*[/\\\\]', '', file_path)
        self.win.title(file_name + ' - MAKDO')
        self.set_file_name_on_status_bar(file_name)
        # PAINT
        paint_keywords = self.paint_keywords.get()
        self.line_data = [LineDatum() for line in self.file_lines]
        for i, line in enumerate(self.file_lines):
            self.line_data[i].line_number = i
            self.line_data[i].line_text = line + '\n'
            if i > 0:
                self.line_data[i].beg_chars_state \
                    = self.line_data[i - 1].end_chars_state.copy()
                self.line_data[i].beg_chars_state.reset_partially()
            self.line_data[i].paint_line(self.txt, paint_keywords)
        # CLEAR THE UNDO STACK
        self.txt.edit_reset()

    def open_dropped_file(self, event):                         # drag and drop
        res_doc = '^(.|\n)+\\.(md|docx)$'                       # drag and drop
        res_xls = '^(.|\n)+\\.(xlsx)$'                          # drag and drop
        res_img = '^(.|\n)+\\.(jpg|jpeg|png|gif|tif|tiff|bmp)$'
        file_path = event.data                                  # drag and drop
        file_path = re.sub('^{(.*)}$', '\\1', file_path)        # drag and drop
        if re.match(res_doc, file_path, re.I):                  # drag and drop
            ans = self.close_file()                             # drag and drop
            if ans is None:                                     # drag and drop
                return None                                     # drag and drop
            self.just_open_file(file_path)                      # drag and drop
        elif re.match(res_xls, file_path, re.I):                # drag and drop
            self.insert_table_from_excel(file_path)             # drag and drop
        elif re.match(res_img, file_path, re.I):                # drag and drop
            image_md_text = '![代替テキスト:縦x横](' + file_path + ' "説明")'
            self.txt.insert('insert', image_md_text)            # drag and drop

    # CLOSE FILE

    def close_file(self):
        # SAVE FILE
        if self._has_edited():
            ans = self._ask_to_save('保存しますか？')
            if ans is None:
                return None
            elif ans is True:
                if not self.save_file():
                    return None
        if self._has_edited():
            ans = self._ask_to_save('データが消えますが、保存しますか？')
            if ans is None:
                return None
            elif ans is True:
                if not self.save_file():
                    return None
        # REMOVE AUTO SAVE FILE
        self.remove_auto_file(self.file_path)
        self.file_path = None
        self.init_text = ''
        self.txt.delete('1.0', 'end')
        self.win.title('MAKDO')
        self.set_file_name_on_status_bar('')
        return True

    # SAVE FILE

    def _has_edited(self):
        file_text = self.txt.get('1.0', 'end-1c')
        file_text = self.get_fully_unfolded_document(file_text)
        if file_text != '':
            if self.init_text != file_text:
                return True
        return False

    def _ask_to_save(self, message):
        tkinter.Tk().withdraw()
        n, m, d = '確認', message, 'yes'
        return tkinter.messagebox.askyesnocancel(n, m, default=d)

    def save_file(self):
        if self._has_edited():
            file_text = self.txt.get('1.0', 'end-1c')
            self._stamp_time(file_text)
            if file_text == '' or file_text[-1] != '\n':
                self.txt.insert('end', '\n')
            file_text = self.txt.get('1.0', 'end-1c')
            file_text = self.get_fully_unfolded_document(file_text)
            if (self.file_path is None) or (self.file_path == ''):
                typ = [('Markdown', '*.md')]
                file_path = tkinter.filedialog.asksaveasfilename(filetypes=typ)
                if file_path == ():
                    return False
                self.file_path = file_path
            if self.make_backup_file.get() and not self.has_made_backup_file:
                if os.path.exists(self.file_path) and \
                   not os.path.islink(self.file_path):
                    try:
                        os.rename(self.file_path, self.file_path + '~')
                        self.has_made_backup_file = True
                    except BaseException:
                        n, m = 'エラー', 'バックアップに失敗しました．'
                        tkinter.messagebox.showerror(n, m)
                        return False
            # DOCX OR MD
            if re.match('^(?:.|\n)+.docx$', self.file_path):
                md_path = self.temp_dir.name + '/doc.md'
            else:
                md_path = self.file_path
            # SAVE MD FILE
            try:
                with open(md_path, 'w') as f:
                    f.write(file_text)
            except BaseException:
                n, m = 'エラー', 'ファイルの保存に失敗しました．'
                tkinter.messagebox.showerror(n, m)
                return False
            # SAVE DOCX FILE
            if re.match('^(?:.|\n)+\\.docx$', self.file_path):
                stderr = sys.stderr
                sys.stderr = tempfile.TemporaryFile(mode='w+')
                importlib.reload(makdo.makdo_md2docx)
                try:
                    m2d = makdo.makdo_md2docx.Md2Docx(md_path)
                    m2d.save(self.file_path)
                except BaseException:
                    pass
                sys.stderr.seek(0)
                msg = sys.stderr.read()
                sys.stderr = stderr
                if msg != '':
                    n = '警告'
                    tkinter.messagebox.showwarning(n, msg)
                    # return
            self.set_message_on_status_bar('保存しました')
            self.init_text = self.get_fully_unfolded_document(file_text)
            #
            return True

    def _stamp_time(self, file_text):
        if not re.match('^\\s*<!--', file_text):
            return
        file_text = re.sub('-->(.|\n)*$', '', file_text)
        now = datetime.datetime.utcnow() + datetime.timedelta(hours=+9)
        jst = datetime.timezone(datetime.timedelta(hours=+9))
        now = now.replace(tzinfo=jst)
        res = '^(\\S+:\\s*)(\\S+)(\\s.*)?$'
        for i, line in enumerate(file_text.split('\n')):
            # CREATED TIME
            if re.match('^作成時:', line) or re.match('^created_time:', line):
                cfg = re.sub(res, '\\1', line)
                tim = re.sub(res, '\\2', line)
                usr = re.sub(res, '\\3', line)
                j, k = len(cfg),  len(tim)
                beg = str(i + 1) + '.' + str(j)
                end = str(i + 1) + '.' + str(j + k)
                res_jst = '^' + '[0-9]{4}-[0-9]{2}-[0-9]{2}' + \
                    'T[0-9]{2}:[0-9]{2}:[0-9]{2}\\+09:00' + '(\\s.*)?$'
                if not re.match(res_jst, tim):
                    tim = ''
                try:
                    dt = datetime.datetime.fromisoformat(tim)
                except BaseException:
                    self.txt.delete(beg, end)
                    self.txt.insert(beg, now.isoformat(timespec='seconds'))
            if re.match('^更新時:', line) or re.match('^modified_time:', line):
                cfg = re.sub(res, '\\1', line)
                tim = re.sub(res, '\\2', line)
                usr = re.sub(res, '\\3', line)
                j, k = len(cfg),  len(tim)
                beg = str(i + 1) + '.' + str(j)
                end = str(i + 1) + '.' + str(j + k)
                self.txt.delete(beg, end)
                self.txt.insert(beg, now.isoformat(timespec='seconds'))

    # NAME AND SAVE

    def name_and_save_by_md(self):
        typ = [('Markdown', '.md')]
        file_path = tkinter.filedialog.asksaveasfilename(filetypes=typ)
        if file_path == () or file_path == '':
            return False
        self.remove_auto_file(self.file_path)
        self.file_path = file_path
        self.init_text = ''
        file_name = re.sub('^.*[/\\\\]', '', file_path)
        self.win.title(file_name + ' - MAKDO')
        self.set_file_name_on_status_bar(file_name)
        self.save_file()
        return True

    def name_and_save_by_docx(self):
        typ = [('MS Word', '.docx')]
        file_path = tkinter.filedialog.asksaveasfilename(filetypes=typ)
        if file_path == () or file_path == '':
            return False
        self.remove_auto_file(self.file_path)
        self.file_path = file_path
        self.init_text = ''
        file_name = re.sub('^.*[/\\\\]', '', file_path)
        self.win.title(file_name + ' - MAKDO')
        self.set_file_name_on_status_bar(file_name)
        self.temp_dir = tempfile.TemporaryDirectory()
        self.save_file()
        return True

    # def name_and_save(self):
    #     typ = [('可能な形式', '.md .docx'),
    #            ('Markdown', '.md'), ('MS Word', '.docx')]
    #     file_path = tkinter.filedialog.asksaveasfilename(filetypes=typ)
    #     if file_path == () or file_path == '':
    #         return False
    #     self.remove_auto_file(self.file_path)
    #     self.file_path = file_path
    #     self.init_text = ''
    #     file_name = re.sub('^.*[/\\\\]', '', file_path)
    #     self.win.title(file_name + ' - MAKDO')
    #     self.set_file_name_on_status_bar(file_name)
    #     if re.match('^(.|\n)+\\.docx$', self.file_path):
    #         self.temp_dir = tempfile.TemporaryDirectory()
    #     self.save_file()
    #     return True

    # AUTO SAVE

    def get_auto_path(self, file_path):
        if file_path is None or file_path == '':
            return None
        if '/' in file_path or '\\' in file_path:
            d = re.sub('^((?:.|\n)*[/\\\\])(.*)$', '\\1', file_path)
            f = re.sub('^((?:.|\n)*[/\\\\])(.*)$', '\\2', file_path)
        else:
            d = ''
            f = file_path
        if '.' in f:
            n = re.sub('^((?:.|\n)*)(\\..*)$', '\\1', f)
            e = re.sub('^((?:.|\n)*)(\\..*)$', '\\2', f)
        else:
            n = f
            e = ''
        n = re.sub('^((?:.|\n){,240})(.*)$', '\\1', n)
        return d + '~$' + n + e + '.zip'

    def exists_auto_file(self, file_path):
        auto_path = self.get_auto_path(file_path)
        if os.path.exists(auto_path):
            # auto_file = re.sub('^(.|\n)*[/\\\\]', '', auto_path)
            n = 'エラー'
            m = '自動保存ファイルが存在します．\n' + \
                '"' + auto_path + '"\n\n' + \
                '①現在、ファイルを編集中\n' + \
                '②過去の編集中のファイルが残存\n' + \
                'の2つの可能性が考えられます．\n\n' + \
                '①現在、ファイルを編集中\n' + \
                'の場合は、「No」を選択してください．\n\n' + \
                '②過去の編集中のファイルが残存\n' + \
                'の場合、異常終了したものと思われます．\n' + \
                '「No」を選択して、' + \
                '自動保存ファイルの中身を確認してから、' + \
                '削除することをおすすめします．\n\n' + \
                '自動保存ファイルを削除しますか？'
            ans = tkinter.messagebox.askyesno(n, m, default='no')
            if ans:
                try:
                    self.remove_auto_file(file_path)
                except BaseException:
                    n, m = 'エラー', '自動保存ファイルの削除に失敗しました．'
                    tkinter.messagebox.showerror(n, m)
        if os.path.exists(auto_path):
            return True
        else:
            return False

    def save_auto_file(self, file_path):
        if file_path is not None and file_path != '':
            new_text = self.txt.get('1.0', 'end-1c')
            auto_path = self.get_auto_path(file_path)
            if os.path.exists(auto_path):
                with zipfile.ZipFile(auto_path, 'r') as old_zip:
                    with old_zip.open('doc.md', 'r') as f:
                        old_text = f.read()
                        if new_text == old_text.decode():
                            return
            with zipfile.ZipFile(auto_path, 'w',
                                 compression=zipfile.ZIP_DEFLATED,
                                 compresslevel=9) as new_zip:
                new_zip.writestr('doc.md', new_text)

    def remove_auto_file(self, file_path):
        if file_path is not None and file_path != '':
            auto_path = self.get_auto_path(file_path)
            if re.match('(^|(.|\n)*[/\\\\])~\\$(.|\n)+\\.zip$', auto_path):
                if os.path.exists(auto_path):
                    os.remove(auto_path)

    # CONVERT DIRECTLY

    def convert_directly(self):
        self.quit_editing_formula()
        self.close_memo_pad()
        self.pnd.update()
        half_height = int(self.pnd.winfo_height() / 2) - 5
        self.pnd.remove(self.pnd1)
        self.pnd.remove(self.pnd2)
        self.pnd.remove(self.pnd3)
        self.pnd.remove(self.pnd4)
        self.pnd.remove(self.pnd5)
        self.pnd.remove(self.pnd6)
        self.pnd.add(self.pnd1, height=half_height, minsize=100)
        self.pnd.add(self.pnd4, height=half_height)
        self.pnd.update()
        #
        btn = tkinter.Button(self.pnd4, text='キャンセル',
                             command=self._quit_converting_directly)
        btn.pack(side='bottom')
        #
        self.pool = tkinter.Text(self.pnd4)
        self.pool.drop_target_register(tkinterdnd2.DND_FILES)
        self.pool.insert('end', 'ここにmdファイル又はdocxファイルをドロップしてください\n')
        self.pool.dnd_bind('<<Drop>>', self._convert_dropped_file)
        self.pool.pack(expand=True, side='top', fill='both')
        self.pool.config(bg='#00A586', fg='white')
        size = self.font_size.get()
        self.pool['font'] = (GOTHIC_FONT, size)

    def _convert_dropped_file(self, event):
        filename = event.data
        filename = re.sub('^{(.*)}$', '\\1', filename)
        basename = os.path.basename(filename)
        self.pool.delete('1.0', 'end')
        self.pool.insert('end', '"' + basename + '"を受け取りました\n')
        stderr = sys.stderr
        sys.stderr = tempfile.TemporaryFile(mode='w+')
        if re.match('^.*\\.(m|M)(d|D)$', filename):
            self.pool.insert('end', 'docxファイルを作成します\n')
            try:
                importlib.reload(makdo.makdo_md2docx)
                m2d = makdo.makdo_md2docx.Md2Docx(filename)
                m2d.save('')
                self.pool.insert('end', 'docxファイルを作成しました\n')
            except BaseException:
                sys.stderr.seek(0)
                self.pool.insert('end', sys.stderr.read())
                self.pool.insert('end', 'docxファイルを作成できませんでした\n')
        elif re.match('^.*\\.(d|D)(o|O)(c|C)(x|X)$', filename):
            self.pool.insert('end', 'mdファイルを作成します\n')
            try:
                importlib.reload(makdo.makdo_docx2md)
                d2m = makdo.makdo_docx2md.Docx2Md(filename)
                d2m.save('')
                self.pool.insert('end', 'mdファイルを作成しました\n')
            except BaseException:
                sys.stderr.seek(0)
                self.pool.insert('end', sys.stderr.read())
                self.pool.insert('end', 'mdファイルを作成できませんでした\n')
        else:
            self.pool.insert('end', '不適切なファイルです\n')
        sys.stderr = stderr
        self.pool.insert('end', '\nここにmdファイル又はdocxファイルをドロップしてください\n')

    def _quit_converting_directly(self):
        self.pnd.remove(self.pnd4)
        self.txt.focus_set()

    # CONVERT TO PDF

    def convert_to_pdf(self):
        typ = [('PDF', '.pdf')]
        pdf_path = tkinter.filedialog.asksaveasfilename(filetypes=typ)
        tmp_docx = self._get_tmp_docx()
        if sys.platform == 'win32':
            Application = win32com.client.Dispatch("Word.Application")
            Application.Visible = False
            doc = Application.Documents.Open(FileName=tmp_docx,
                                             ConfirmConversions=False,
                                             ReadOnly=True)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17=PDF
        elif sys.platform == 'darwin':
            n, m = 'お詫び', '準備中です．\n（macの開発環境が手元にない…）'
            tkinter.messagebox.showinfo(n, m)
        elif sys.platform == 'linux':
            dir_path = re.sub('((?:.|\n)*)/(?:.|\n)+$', '\\1', tmp_docx)
            com = '/usr/bin/libreoffice --headless --convert-to pdf --outdir '
            doc = subprocess.run(com + dir_path + ' ' + tmp_docx,
                                 check=True,
                                 shell=True,
                                 stdout=subprocess.PIPE,
                                 encoding="utf-8")
            tmp_pdf = re.sub('docx$', 'pdf', tmp_docx)
            shutil.move(tmp_pdf, pdf_path)

    # START WRITER

    def start_writer(self):
        docx_path = self._get_tmp_docx()
        if sys.platform == 'win32':
            Application = win32com.client.Dispatch("Word.Application")
            Application.Visible = True
            doc = Application.Documents.Open(FileName=docx_path,
                                             ConfirmConversions=False,
                                             ReadOnly=True)
        elif sys.platform == 'darwin':
            n, m = 'お詫び', '準備中です．\n（macの開発環境が手元にない…）'
            tkinter.messagebox.showinfo(n, m)
        elif sys.platform == 'linux':
            doc = subprocess.run('/usr/bin/libreoffice ' + docx_path,
                                 check=True,
                                 shell=True,
                                 stdout=subprocess.PIPE,
                                 encoding="utf-8")

    # QUIT

    def quit_makdo(self):
        ans = self.close_file()
        if ans is None:
            return None
        self.win.quit()
        self.win.destroy()
        sys.exit(0)

    ##########################
    # MENU EDIT

    def _make_menu_edit(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='編集(E)', menu=menu, underline=3)
        #
        menu.add_command(label='元に戻す(U)', underline=5,
                         command=self.edit_modified_undo, accelerator='Ctrl+Z')
        menu.add_command(label='やり直す(R)', underline=5,
                         command=self.edit_modified_redo, accelerator='Ctrl+Y')
        menu.add_separator()
        #
        menu.add_command(label='切り取り(C)', underline=5,
                         command=self.cut_region, accelerator='Ctrl+X')
        menu.add_command(label='コピー(Y)', underline=4,
                         command=self.copy_region, accelerator='Ctrl+C')
        menu.add_command(label='貼り付け(P)', underline=5,
                         command=self.paste_region, accelerator='Ctrl+V')
        menu.add_separator()
        #
        menu.add_command(label='矩形（四角形）を切り取り',
                         command=self.cut_rectangle)
        menu.add_command(label='矩形（四角形）をコピー',
                         command=self.copy_rectangle)
        menu.add_command(label='矩形（四角形）を貼り付け',
                         command=self.paste_rectangle)
        menu.add_separator()
        #
        menu.add_command(label='全て選択(A)', underline=5,
                         command=self.select_all, accelerator='Ctrl+A')
        menu.add_separator()
        #
        menu.add_command(label='全て置換',
                         command=self.replace_all)
        menu.add_separator()
        #
        menu.add_command(label='選択範囲の行を正順にソート（並替え）',
                         command=self.sort_lines)
        menu.add_command(label='選択範囲の行を逆順にソート（並替え）',
                         command=self.sort_lines_in_reverse_order)
        menu.add_separator()
        #
        menu.add_command(label='数式を計算',
                         command=self.calculate)
        menu.add_separator()
        #
        menu.add_command(label='字体を変える',
                         command=self.change_typeface)
        menu.add_separator()
        #
        menu.add_command(label='コメントアウトにする',
                         command=self.comment_out_region)
        menu.add_command(label='コメントアウトを取り消す',
                         command=self.uncomment_in_region)
        # menu.add_separator()

    ######
    # COMMAND

    def edit_modified_undo(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        try:
            pane.edit_undo()
        except BaseException:
            pass

    def edit_modified_redo(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        try:
            pane.edit_redo()
        except BaseException:
            pass

    def cut_region(self):
        self._cut_or_copy_region(True)

    def copy_region(self):
        self._cut_or_copy_region(False)

    def _cut_or_copy_region(self, must_cut=False):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if must_cut:
            if self._is_read_only_pane(pane):
                return False
        beg, end = self._get_region(pane)
        if beg == '' or end == '':
            self._show_no_region_error()
            return False
        c = pane.get(beg, end)
        self.win.clipboard_clear()
        self.win.clipboard_append(c)
        if must_cut:
            pane.delete(beg, end)
        self._cancel_region(pane)
        return True

    def paste_region(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if self._is_read_only_pane(pane):
            return False
        if self.current_pane == 'txt':
            beg_v = self._get_v_position_of_insert()
        try:
            cb = self.win.clipboard_get()
        except BaseException:
            cb = ''
        if cb == '':
            return True
        pane.insert('insert', cb)
        # pane.yview('insert -20 line')
        if self.current_pane == 'txt':
            end_v = self._get_v_position_of_insert()
        if self.current_pane == 'txt':
            for i in range(beg_v - 1, end_v - 1):
                self.paint_out_line(i)
        return True

    def cut_rectangle(self):
        self._cut_or_copy_rectangle(True)

    def copy_rectangle(self):
        self._cut_or_copy_rectangle(False)

    def _cut_or_copy_rectangle(self, must_cut=False):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if must_cut:
            if self._is_read_only_pane(pane):
                return False
        beg, end = self._get_region(pane)
        if beg == '' or end == '':
            self._show_no_region_error()
            return False
        beg_v = int(re.sub('\\.[0-9]+$', '', beg))
        s = pane.get(beg + ' linestart', beg)
        beg_ih = get_ideal_width(s)
        end_v = int(re.sub('\\.[0-9]+$', '', end))
        s = pane.get(end + ' linestart', end)
        end_ih = get_ideal_width(s)
        min_ih = min(beg_ih, end_ih)
        max_ih = max(beg_ih, end_ih)
        self.rectangle_text_list = []
        for i in range(beg_v, end_v + 1):
            line = pane.get(str(i) + '.0', str(i) + '.end')
            line_pre, line_mid, line_pos = '', '', ''
            for c in line:
                if get_ideal_width(line_pre) < min_ih:
                    line_pre += c
                elif get_ideal_width(line_pre + line_mid) < max_ih:
                    line_mid += c
                else:
                    line_pos += c
            self.rectangle_text_list.append(line_mid)
            if must_cut:
                pane.delete(str(i) + '.' + str(len(line_pre)),
                            str(i) + '.' + str(len(line_pre + line_mid)))
                self.paint_out_line(i)
        self._cancel_region(pane)
        return True

    def paste_rectangle(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if self._is_read_only_pane(pane):
            return False
        if self.rectangle_text_list == []:
            return True
        ins_v = self._get_v_position_of_insert()
        max_v = self._get_max_v_position()
        s = pane.get(str(ins_v) + '.0', 'insert')
        ins_ih = get_ideal_width(s)
        for j, line_md in enumerate(self.rectangle_text_list):
            i = ins_v + j
            if i < max_v:
                line = pane.get(str(i) + '.0', str(i) + '.end')
                line_pre, line_pos = '', ''
                for c in line:
                    if get_ideal_width(line_pre) < ins_ih:
                        line_pre += c
                    else:
                        break
                ins_h = str(i) + '.' + str(len(line_pre))
            else:
                ins_h = 'end'
                line_md += '\n'
            pane.insert(ins_h, line_md)
            pane.mark_set('insert', ins_h)
            self.paint_out_line(i)
        return True

    def select_all(self):
        self.txt.tag_add('sel', '1.0', 'end-1c')

    def replace_all(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if self._is_read_only_pane(pane):
            return
        word1 = self.stb_sor1.get()
        word2 = self.stb_sor2.get()
        if word1 == '':
            t = '全置換'
            m = '検索する言葉と置換する言葉を入力してください．'
            sd = TwoWordsDialog(pane, self, t, m, word1, word2)
            word1, word2 = sd.get_value()
        if word1 == '':
            return
        if Makdo.search_word != word1:
            Makdo.search_word = word1
        if pane.tag_ranges('sel'):
            beg, end = pane.index('sel.first'), pane.index('sel.last')
        elif 'akauni' in pane.mark_names():
            beg, end = self._get_indices_in_order(pane, 'insert', 'akauni')
        else:
            beg, end = '1.0', 'end-1c'
        m = pane.get(beg, end).count(word1)
        while True:
            tex = pane.get(beg, end)
            if word1 not in tex:
                break
            res = '^((?:.|\n)*?)' + word1 + '(?:.|\n)*$'
            sub = re.sub(res, '\\1', tex)
            pane.delete(beg + '+' + str(len(sub)) + 'c',
                        beg + '+' + str(len(sub + word1)) + 'c')
            pane.insert(beg + '+' + str(len(sub)) + 'c', word2)
        if pane.tag_ranges('sel'):
            pane.tag_remove('sel', "1.0", "end")
        elif 'akauni' in pane.mark_names():
            pane.tag_remove('akauni_tag', '1.0', 'end')
            pane.mark_unset('akauni')
        pane.focus_set()
        # MESSAGE
        self.set_message_on_status_bar(str(m) + '個を置換しました')

    def sort_lines(self):
        self._sort_lines(True)

    def sort_lines_in_reverse_order(self):
        self._sort_lines(False)

    def _sort_lines(self, is_ascending_order=True):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if self._is_read_only_pane(pane):
            return
        if pane.tag_ranges('sel'):
            beg, end = pane.index('sel.first'), pane.index('sel.last')
        elif 'akauni' in pane.mark_names():
            beg, end = self._get_indices_in_order(pane, 'insert', 'akauni')
        else:
            return
        beg_line = int(re.sub('\\.[0-9]+', '', beg))
        end_line = int(re.sub('\\.[0-9]+', '', end))
        if not re.match('^[0-9]+\\.0$', beg):
            beg_line += 1
        end_line -= 1
        lines_str = pane.get(str(beg_line) + '.0', str(end_line) + '.end')
        lines_lst = lines_str.split('\n')
        pane.delete(str(beg_line) + '.0', str(end_line) + '.end')
        if pane.tag_ranges('sel'):
            pane.tag_remove('sel', "1.0", "end")
        elif 'akauni' in pane.mark_names():
            pane.tag_remove('akauni_tag', '1.0', 'end')
            pane.mark_unset('akauni')
        sorted_lst = sorted(lines_lst)
        if not is_ascending_order:
            sorted_lst.reverse()
        sorted_str = '\n'.join(sorted_lst)
        pane.insert(str(beg_line) + '.0', sorted_str)
        for j, line in enumerate(sorted_lst):
            i = beg_line - 1 + j
            self.paint_out_line(i)

    def calculate(self):
        line = self.txt.get('insert linestart', 'insert lineend')
        line_head = ''
        line_math = line
        line_rslt = ''
        line_tail = ''
        res = '^(.*(?:<!--|@))(.*)$'
        if re.match(res, line_math):
            line_head = re.sub(res, '\\1', line_math)
            line_math = re.sub(res, '\\2', line_math)
        res = '^(.*)((?:-->|#).*)$'
        if re.match(res, line_math):
            line_tail = re.sub(res, '\\2', line_math)
            line_math = re.sub(res, '\\1', line_math)
        res = '^(.*)(=.*)$'
        if re.match(res, line_math):
            line_rslt = re.sub(res, '\\2', line_math)
            line_math = re.sub(res, '\\1', line_math)
        if line_math == '':
            return
        math = line_math
        math = math.replace('\t', ' ').replace('\u3000', ' ')
        math = math.replace('，', ',').replace('．', '.')
        math = math.replace('０', '0').replace('１', '1').replace('２', '2')
        math = math.replace('３', '3').replace('４', '4').replace('５', '5')
        math = math.replace('６', '6').replace('７', '7').replace('８', '8')
        math = math.replace('９', '9')
        math = math.replace('〇', '0').replace('一', '1').replace('二', '2')
        math = math.replace('三', '3').replace('四', '4').replace('五', '5')
        math = math.replace('六', '6').replace('七', '7').replace('八', '8')
        math = math.replace('九', '9')
        math = math.replace('（', '(').replace('）', ')')
        math = math.replace('｛', '{').replace('｝', '}')
        math = math.replace('［', '[').replace('］', ']')
        math = math.replace('｜', '|').replace('！', '!').replace('＾', '^')
        math = math.replace('＊', '*').replace('／', '/').replace('％', '%')
        math = math.replace('＋', '+').replace('−', '-')
        math = math.replace('×', '*').replace('÷', '/').replace('ー', '-')
        math = math.replace('△', '-').replace('▲', '-')
        math = math.replace('パ-セント', '%')
        # ' ', ','
        math = math.replace(' ', '').replace(',', '')
        # {, }, [, ]
        math = math.replace('{', '(').replace('}', ')')
        math = math.replace('[', '(').replace(']', ')')
        # 千, 百, 十
        temp = ''
        unit = ['千', '百', '十']
        for i in range(len(unit)):
            res = '^([^' + unit[i] + ']*' + unit[i] + ')(.*)$'
            while re.match(res, math):
                t1 = re.sub(res, '\\1', math)  # [^千]*千
                t2 = re.sub(res, '\\2', math)  # .*
                if not re.match('^.*[0-9]' + unit[i] + '$', t1):
                    t1 = re.sub(unit[i] + '$', '1' + unit[i], t1)  # 千 -> 1千
                temp += t1
                math = t2
        math = temp + math
        temp = ''
        unit = ['千', '百', '十', '']
        for i in range(len(unit) - 1):
            res = '^([^' + unit[i] + ']*' + unit[i] + ')(.*)$'
            while re.match(res, math):
                t1 = re.sub(res, '\\1', math)  # [^千]*千
                t2 = re.sub(res, '\\2', math)  # .*
                temp += t1
                if not re.match('^[0-9]' + unit[i + 1], t2):
                    t2 = '0' + unit[i + 1] + t2
                math = t2
        math = temp + math
        math = math.replace('千', '').replace('百', '').replace('十', '')
        # 京, 兆, 億, 万
        temp = ''
        unit = ['京', '兆', '億', '万', '']
        for i in range(len(unit) - 1):
            res = '^([^' + unit[i] + ']*' + unit[i] + ')(.*)$'
            while re.match(res, math):
                t1 = re.sub(res, '\\1', math)  # [^京]*京
                t2 = re.sub(res, '\\2', math)  # .*
                temp += t1
                if re.match('[0-9]{,4}' + unit[i + 1], t2):
                    t2 = '0000' + t2
                    math = re.sub('^[0-9]*([0-9]{4})', '\\1', t2)
                else:
                    math = '0000' + unit[i + 1] + t2  # 0000兆
        math = temp + math
        math = math.replace('京', '').replace('兆', '')
        math = math.replace('億', '').replace('万', '')
        # %, 割, 分, 厘
        math = re.sub('([0-9\\.]+)%', '(\\1/100)', math)
        math = re.sub('([0-9\\.]+)割', '(\\1/10)', math)
        math = re.sub('([0-9\\.]+)分', '(\\1/100)', math)
        math = re.sub('([0-9\\.]+)厘', '(\\1/1000)', math)
        # FRACTION
        res = '^(.*?)' \
            + '([0-9]+|\\([^\\(\\)]+\\))分の([0-9]+|\\([^\\(\\)]+\\))' \
            + '(.*?)$'
        while re.match(res, math):
            math = re.sub(res, '\\1(\\3/\\2)\\4', math)
        # POWER
        math = re.sub('\\^', '**', math)
        # REMOVE
        math = re.sub('pi', '3.141592653589793', math)
        math = re.sub('e', '2.718281828459045', math)
        math = re.sub('[^\\(\\)\\|\\*/%\\-\\+0-9\\.]', '', math)
        # EVAL
        r = str(round(eval(math), 10))
        r = re.sub('\\.0$', '', r)
        if not re.match('^-?([0-9]*\\.)?[0-9]+', r):
            return False
        # REPLACE
        digit_separator = self.digit_separator.get()
        if '.' in r:
            i = re.sub('^(.*)(\\..*)$', '\\1', r)
            f = re.sub('^(.*)(\\..*)$', '\\2', r)
        else:
            i = r
            f = ''
        if digit_separator == '3':
            if re.match('^.*[0-9]{19}$', i):
                i = re.sub('([0-9]{18})$', ',\\1', i)
            if re.match('^.*[0-9]{16}$', i):
                i = re.sub('([0-9]{15})$', ',\\1', i)
            if re.match('^.*[0-9]{13}$', i):
                i = re.sub('([0-9]{12})$', ',\\1', i)
            if re.match('^.*[0-9]{10}$', i):
                i = re.sub('([0-9]{9})$', ',\\1', i)
            if re.match('^.*[0-9]{7}$', i):
                i = re.sub('([0-9]{6})$', ',\\1', i)
            if re.match('^.*[0-9]{4}$', i):
                i = re.sub('([0-9]{3})$', ',\\1', i)
        elif digit_separator == '4':
            if re.match('^.*[0-9]{17}$', i):
                i = re.sub('([0-9]{16})$', '京\\1', i)
            if re.match('^.*[0-9]{13}$', i):
                i = re.sub('([0-9]{12})$', '兆\\1', i)
            if re.match('^.*[0-9]{9}$', i):
                i = re.sub('([0-9]{8})$', '億\\1', i)
            if re.match('^.*[0-9]{5}$', i):
                i = re.sub('([0-9]{4})$', '万\\1', i)
        r = i + f
        v_number = self._get_v_position_of_insert()
        beg = str(v_number) + '.' + str(len(line_head + line_math))
        end = str(v_number) + '.' + str(len(line_head + line_math + line_rslt))
        self.txt.delete(beg, end)
        self.txt.insert(beg, '=' + r)
        self.win.clipboard_clear()
        self.win.clipboard_append(r)

    def change_typeface(self):
        c = self.txt.get('insert', 'insert+1c')
        for tf in TYPEFACES:
            if c in tf:
                self.TypefaceDialog(self.txt, self, c, list(tf))
                break
        else:
            n = '警告'
            m = '"' + c + '"に異字体は登録されていません．'
            tkinter.messagebox.showwarning(n, m)

    class TypefaceDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, old_typeface, candidates):
            self.pane = pane
            self.mother = mother
            self.old_typeface = old_typeface
            self.candidates = candidates
            super().__init__(pane, title='字体を変える')

        def body(self, pane):
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            self.typeface = tkinter.StringVar()
            for cnd in self.candidates:
                rd = tkinter.Radiobutton(pane, text=cnd, font=font,
                                         variable=self.typeface, value=cnd)
                rd.pack(side=tkinter.LEFT, padx=3, pady=3)
                if cnd == self.old_typeface:
                    rd.select()
            # self.bind('<Key-Return>', self.ok)
            # self.bind('<Key-Escape>', self.cancel)
            super().body(pane)

        # def buttonbox(self):
        #     btn = tkinter.Frame(self)
        #     self.btn1 = tkinter.Button(btn, text='OK', width=6,
        #                                command=self.ok)
        #     self.btn1.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     self.btn2 = tkinter.Button(btn, text='Cancel', width=6,
        #                                command=self.cancel)
        #     self.btn2.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     btn.pack()

        def apply(self):
            new_typeface = self.typeface.get()
            self.pane.delete('insert', 'insert+1c')
            self.pane.insert('insert', new_typeface)
            self.pane.mark_set('insert', 'insert-1c')
            self.pane.focus_set()

    def comment_out_region(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if self._is_read_only_pane(pane):
            return
        if pane.tag_ranges('sel'):
            beg, end = pane.index('sel.first'), pane.index('sel.last')
        elif 'akauni' in pane.mark_names():
            beg, end = self._get_indices_in_order(pane, 'insert', 'akauni')
        else:
            n = 'エラー'
            m = 'コメントアウトする範囲が指定されていません．'
            tkinter.messagebox.showerror(n, m)
            return
        tex = pane.get(beg, end)
        for i in ['8', '7', '6', '5', '4', '3', '2', '1', '-']:
            if i == '-':
                j = '1'
            else:
                j = str(int(i) + 1)
            for t in (('<!' + i + '-', '<!' + j + '-'),
                      ('-' + i + '>', '-' + j + '>')):
                res = '^((?:.|\n)*?)' + t[0] + '((?:.|\n)*)$'
                while re.match(res, tex):
                    sub = re.sub(res, '\\1', tex)
                    tex = re.sub(res, '\\1' + t[1] + '\\2', tex)
                    pane.delete(beg + '+' + str(len(sub)) + 'c',
                                beg + '+' + str(len(sub + t[0])) + 'c')
                    pane.insert(beg + '+' + str(len(sub)) + 'c', t[1])
        pane.insert(end, '-->')
        pane.insert(beg, '<!--')
        if pane.tag_ranges('sel'):
            pane.tag_remove('sel', "1.0", "end")
        elif 'akauni' in pane.mark_names():
            pane.tag_remove('akauni_tag', '1.0', 'end')
            pane.mark_unset('akauni')

    def uncomment_in_region(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        if self._is_read_only_pane(pane):
            return
        #
        if pane.tag_ranges('sel'):
            beg, end = pane.index('sel.first'), pane.index('sel.last')
        elif 'akauni' in pane.mark_names():
            beg, end = self._get_indices_in_order(pane, 'insert', 'akauni')
        else:
            n = 'エラー'
            m = 'コメントアウトを解除する範囲が指定されていません．'
            tkinter.messagebox.showerror(n, m)
            return
        tex = pane.get(beg, end)
        is_in_comment = False
        tmp = ''
        for c in tex:
            tmp += c
            if re.match('^((?:.|\n)*)<!--$', tmp) and not is_in_comment:
                tmp = re.sub('<!--$', '', tmp)
                pane.delete(beg + '+' + str(len(tmp)) + 'c',
                            beg + '+' + str(len(tmp) + 4) + 'c')
                is_in_comment = True
            if re.match('^((?:.|\n)*)-->$', tmp) and is_in_comment:
                tmp = re.sub('-->$', '', tmp)
                pane.delete(beg + '+' + str(len(tmp)) + 'c',
                            beg + '+' + str(len(tmp) + 3) + 'c')
                is_in_comment = False
        tex = tmp
        for i in ['-', '1', '2', '3', '4', '5', '6', '7', '8']:
            if i == '-':
                j = '1'
            else:
                j = str(int(i) + 1)
            for t in (('<!' + i + '-', '<!' + j + '-'),
                      ('-' + i + '>', '-' + j + '>')):
                res = '^((?:.|\n)*?)' + t[1] + '((?:.|\n)*)$'
                while re.match(res, tex):
                    sub = re.sub(res, '\\1', tex)
                    tex = re.sub(res, '\\1' + t[0] + '\\2', tex)
                    pane.delete(beg + '+' + str(len(sub)) + 'c',
                                beg + '+' + str(len(sub + t[1])) + 'c')
                    pane.insert(beg + '+' + str(len(sub)) + 'c', t[0])
        if pane.tag_ranges('sel'):
            pane.tag_remove('sel', "1.0", "end")
        elif 'akauni' in pane.mark_names():
            pane.tag_remove('akauni_tag', '1.0', 'end')
            pane.mark_unset('akauni')

    ##########################
    # MENU INSERT

    def _make_menu_insert(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='挿入(I)', menu=menu, underline=3)
        #
        menu.add_command(label='空白を挿入',
                         command=self.insert_space)
        menu.add_command(label='改行を挿入',
                         command=self.insert_line_break)
        menu.add_command(label='画像を挿入',
                         command=self.insert_images)
        self._make_submenu_insert_font_change(menu)
        self._make_submenu_insert_font_size_change(menu)
        self._make_submenu_insert_font_width_change(menu)
        self._make_submenu_insert_underline(menu)
        self._make_submenu_insert_font_color_change(menu)
        self._make_submenu_insert_highlight_color_change(menu)
        menu.add_command(label='コード番号から文字を挿入',
                         command=self.insert_character_by_code)
        self._make_submenu_insert_ivs_character(menu)
        menu.add_separator()
        #
        self._make_submenu_insert_time(menu)
        self._make_submenu_insert_file_name(menu)
        menu.add_command(label='テキストファイルの内容を挿入',
                         command=self.insert_file)
        menu.add_separator()
        #
        menu.add_command(label='記号を挿入',
                         command=self.insert_symbol)
        self._make_submenu_insert_horizontal_line(menu)
        menu.add_separator()
        #
        self._make_submenu_insert_sample(menu)
        # menu.add_separator()

    ################
    # COMMAND

    def insert_space(self):
        t = '空白の幅'
        p = '空白の幅を文字数（整数又は小数）で入力してください．'
        f = ''
        while not re.match('^([0-9]*\\.)?[0-9]+$', f):
            f = OneWordDialog(self.txt, self, t, p, f).get_value()
            if f is None:
                return
        self.txt.insert('insert', '< ' + f + ' >')

    def insert_line_break(self):
        self.txt.insert('insert', '<br>')

    def insert_images(self):
        typ = [('画像', '.jpg .jpeg .png .gif .tif .tiff .bmp'),
               ('全てのファイル', '*')]
        image_paths = tkinter.filedialog.askopenfilenames(filetypes=typ)
        for i in image_paths:
            image_md_text = '![代替テキスト:縦x横](' + i + ' "説明")'
            self.txt.insert('insert', image_md_text)

    ################
    # SUBMENU INSERT FONT CHANGE

    def _make_submenu_insert_font_change(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='フォントの変更を挿入', menu=submenu)
        #
        submenu.add_command(label='明朝体を変える',
                            command=self.insert_selected_font)
        submenu.add_separator()
        submenu.add_command(label='ゴシック体に変える',
                            command=self.insert_gothic_font)
        submenu.add_separator()
        submenu.add_command(label='手動入力',
                            command=self.insert_font_manually)

    ######
    # COMMAND

    def insert_selected_font(self):
        mincho_list = []
        for f in tkinter.font.families():
            if '明朝' in f:
                mincho_list.append(f)
        self.MinchoDialog(self.txt, self, mincho_list)

    class MinchoDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, candidates):
            self.pane = pane
            self.mother = mother
            self.candidates = candidates
            super().__init__(pane, title='明朝体を変える')

        def body(self, pane):
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            self.mincho = tkinter.StringVar()
            for cnd in self.candidates:
                rd = tkinter.Radiobutton(pane, text=cnd, font=font,
                                         variable=self.mincho, value=cnd)
                rd.pack(side=tkinter.TOP, padx=3, pady=3)
            # self.bind('<Key-Return>', self.ok)
            # self.bind('<Key-Escape>', self.cancel)
            # super().body(pane)

        # def buttonbox(self):
        #     btn = tkinter.Frame(self)
        #     self.btn1 = tkinter.Button(btn, text='OK', width=6,
        #                                command=self.ok)
        #     self.btn1.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     self.btn2 = tkinter.Button(btn, text='Cancel', width=6,
        #                                command=self.cancel)
        #     self.btn2.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     btn.pack()

        def apply(self):
            m = self.mincho.get()
            d = '@' + m + '@（ここはフォントが変わる）@' + m + '@'
            self.pane.insert('insert', d)
            self.pane.mark_set('insert', 'insert-' + str(len(m) + 2) + 'c')
            self.pane.focus_set()

    def insert_gothic_font(self):
        self.txt.insert('insert', '`（ここはゴシック体）`')
        self.txt.mark_set('insert', 'insert-1c')

    def insert_font_manually(self):
        t = 'フォント'
        p = 'フォント名を入力してください．'
        s = OneWordDialog(self.txt, self, t, p).get_value()
        if s is None:
            return
        d = '@' + s + '@（ここはフォントが変わる）@' + s + '@'
        self.txt.insert('insert', d)
        self.txt.mark_set('insert', 'insert-' + str(len(s) + 2) + 'c')

    ################
    # SUBMENU INSERT FONT SIZE CHANGE

    def _make_submenu_insert_font_size_change(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='文字の大きさの変更を挿入', menu=submenu)
        #
        submenu.add_command(label='特小サイズ',
                            command=self.insert_ss_font_size)
        submenu.add_command(label='小サイズ',
                            command=self.insert_s_font_size)
        submenu.add_command(label='大サイズ',
                            command=self.insert_l_font_size)
        submenu.add_command(label='特大サイズ',
                            command=self.insert_ll_font_size)
        submenu.add_separator()
        submenu.add_command(label='手動入力',
                            command=self.insert_font_size_manually)

    ######
    # COMMAND

    def insert_ss_font_size(self):
        self.txt.insert('insert', '---（ここは文字が特に小さい）---')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_s_font_size(self):
        self.txt.insert('insert', '--（ここは文字が小さい）--')
        self.txt.mark_set('insert', 'insert-2c')

    def insert_l_font_size(self):
        self.txt.insert('insert', '++（ここは文字が大きい）++')
        self.txt.mark_set('insert', 'insert-2c')

    def insert_ll_font_size(self):
        self.txt.insert('insert', '+++（ここは文字が特に大きい）+++')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_font_size_manually(self):
        t = '文字の大きさ'
        p = '文字の大きさを1から100までの数字を入力してください．'
        f = ''
        while not re.match('^([0-9]*\\.)?[0-9]+$', f):
            f = OneWordDialog(self.txt, self, t, p, f).get_value()
            if f is None:
                return
        f = re.sub('\\.0+$', '', f)
        d = '@' + f + '@（ここは文字の大きさが変わる）@' + f + '@'
        self.txt.insert('insert', d)
        self.txt.mark_set('insert', 'insert-' + str(len(f) + 2) + 'c')

    ################
    # SUBMENU INSERT FONT WIDTH CHANGE

    def _make_submenu_insert_font_width_change(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='文字の幅の変更を挿入', menu=submenu)
        #
        submenu.add_command(label='特細サイズ',
                            command=self.insert_ss_font_width)
        submenu.add_command(label='細サイズ',
                            command=self.insert_s_font_width)
        submenu.add_command(label='太サイズ',
                            command=self.insert_l_font_width)
        submenu.add_command(label='特太サイズ',
                            command=self.insert_ll_font_width)

    ######
    # COMMAND

    def insert_ss_font_width(self):
        self.txt.insert('insert', '>>>（ここは文字が特に細い）<<<')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_s_font_width(self):
        self.txt.insert('insert', '>>（ここは文字が細い）<<')
        self.txt.mark_set('insert', 'insert-2c')

    def insert_l_font_width(self):
        self.txt.insert('insert', '<<（ここは文字が太い）>>')
        self.txt.mark_set('insert', 'insert-2c')

    def insert_ll_font_width(self):
        self.txt.insert('insert', '<<<（ここは文字が特に太い）>>>')
        self.txt.mark_set('insert', 'insert-3c')

    ################
    # SUBMENU INSERT UNDERLINE

    def _make_submenu_insert_underline(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='文字に下線をを引く', menu=submenu)
        #
        submenu.add_command(label='単線',
                            command=self.insert_single_underline)
        submenu.add_command(label='二重線',
                            command=self.insert_double_underline)
        submenu.add_command(label='波線',
                            command=self.insert_wave_underline)
        submenu.add_command(label='破線',
                            command=self.insert_dash_underline)
        submenu.add_command(label='点線',
                            command=self.insert_dot_underline)

    ######
    # COMMAND

    def insert_single_underline(self):
        self.txt.insert('insert', '__（ここは下線が引かれる）__')
        self.txt.mark_set('insert', 'insert-2c')

    def insert_double_underline(self):
        self.txt.insert('insert', '_=_（ここは下線が引かれる）_=_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_wave_underline(self):
        self.txt.insert('insert', '_~_（ここは下線が引かれる）_~_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_dash_underline(self):
        self.txt.insert('insert', '_-_（ここは下線が引かれる）_-_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_dot_underline(self):
        self.txt.insert('insert', '_._（ここは下線が引かれる）_._')
        self.txt.mark_set('insert', 'insert-3c')

    ################
    # SUBMENU INSERT FONT COLOR CHANGE

    def _make_submenu_insert_font_color_change(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='文字色を変える', menu=submenu)
        #
        submenu.add_command(label='赤色',
                            command=self.insert_r_font_color)
        submenu.add_command(label='黄色',
                            command=self.insert_y_font_color)
        submenu.add_command(label='緑色',
                            command=self.insert_g_font_color)
        submenu.add_command(label='シアン',
                            command=self.insert_c_font_color)
        submenu.add_command(label='青色',
                            command=self.insert_b_font_color)
        submenu.add_command(label='マゼンタ',
                            command=self.insert_m_font_color)
        submenu.add_command(label='白色',
                            command=self.insert_w_font_color)

    ######
    # COMMAND

    def insert_r_font_color(self):
        self.txt.insert('insert', '^R^（ここは文字が赤色）^R^')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_y_font_color(self):
        self.txt.insert('insert', '^Y^（ここは文字が黄色）^Y^')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_g_font_color(self):
        self.txt.insert('insert', '^G^（ここは文字が緑色）^G^')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_c_font_color(self):
        self.txt.insert('insert', '^C^（ここは文字がシアン）^C^')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_b_font_color(self):
        self.txt.insert('insert', '^B^（ここは文字が青色）^B^')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_m_font_color(self):
        self.txt.insert('insert', '^M^（ここは文字がマゼンタ）^M^')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_w_font_color(self):
        self.txt.insert('insert', '^^（ここは文字が白色）^^')
        self.txt.mark_set('insert', 'insert-2c')

    ################
    # SUBMENU INSERT HIGHLIGHT COLOR CHANGE

    def _make_submenu_insert_highlight_color_change(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='下地色を変える', menu=submenu)
        #
        submenu.add_command(label='赤色',
                            command=self.insert_r_highlight_color)
        submenu.add_command(label='黄色',
                            command=self.insert_y_highlight_color)
        submenu.add_command(label='緑色',
                            command=self.insert_g_highlight_color)
        submenu.add_command(label='シアン',
                            command=self.insert_c_highlight_color)
        submenu.add_command(label='青色',
                            command=self.insert_b_highlight_color)
        submenu.add_command(label='マゼンタ',
                            command=self.insert_m_highlight_color)

    ######
    # COMMAND

    def insert_r_highlight_color(self):
        self.txt.insert('insert', '_R_（ここは下地が赤色）_R_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_y_highlight_color(self):
        self.txt.insert('insert', '_Y_（ここは下地が黄色）_Y_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_g_highlight_color(self):
        self.txt.insert('insert', '_G_（ここは下地が緑色）_G_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_c_highlight_color(self):
        self.txt.insert('insert', '_C_（ここは下地がシアン）_C_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_b_highlight_color(self):
        self.txt.insert('insert', '_B_（ここは下地が青色）_B_')
        self.txt.mark_set('insert', 'insert-3c')

    def insert_m_highlight_color(self):
        self.txt.insert('insert', '_M_（ここは下地がマゼンタ）_M_')
        self.txt.mark_set('insert', 'insert-3c')

    ################
    # COMMAND

    def insert_character_by_code(self):
        t = 'コード番号'
        p = 'コード番号を入力してください．'
        s = ''
        while not re.match('^[0-9a-fA-F]{4}$', s):
            s = OneWordDialog(self.txt, self, t, p, s).get_value()
            if s is None:
                return
        self.txt.insert('insert', chr(int(s, 16)))

    ################
    # SUBMENU INSERT IVS CHARACTER

    def _make_submenu_insert_ivs_character(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='人名・地名の字体を挿入', menu=submenu)
        #
        submenu.add_command(label='文字コードから人名・地名の字体を挿入',
                            command=self.insert_ivs)
        submenu.add_separator()
        submenu.add_command(label='"祇"の人名・地名の字体の候補を全て挿入',
                            command=self.insert_ivs_of_7947)
        submenu.add_command(label='"花"の人名・地名の字体の候補を全て挿入',
                            command=self.insert_ivs_of_82b1)
        submenu.add_command(label='"葛"の人名・地名の字体の候補を全て挿入',
                            command=self.insert_ivs_of_845b)
        submenu.add_command(label='"邉"の人名・地名の字体の候補を全て挿入',
                            command=self.insert_ivs_of_9089)
        submenu.add_command(label='"邊"の人名・地名の字体の候補を全て挿入',
                            command=self.insert_ivs_of_908a)

    ######
    # COMMAND

    def insert_ivs(self):
        c = ''
        if self.txt.tag_ranges('sel'):
            c = self.txt.get('sel.first', 'sel.last')
        elif 'akauni' in self.txt.mark_names():
            c = ''
            c += self.txt.get('akauni', 'insert')
            c += self.txt.get('insert', 'akauni')
        if len(c) == 1:
            i = self.IvsDialog(self.txt, self, c)
        else:
            i = self.IvsDialog(self.txt, self)
        if len(c) == 1 and i.has_inserted:
            if self.txt.tag_ranges('sel'):
                self.txt.delete('sel.first', 'sel.first+1c')
            elif 'akauni' in self.txt.mark_names():
                if self.txt.get('akauni', 'insert') != '':
                    self.txt.delete('akauni', 'akauni+1c')
                elif self.txt.get('insert', 'akauni') != '':
                    self.txt.delete('akauni-1c', 'akauni')

    class IvsDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, char=None):
            self.pane = pane
            self.mother = mother
            self.char = None
            self.code = None
            if char is not None:
                self.char = char
                self.code = re.sub('^0x', '', hex(ord(char))).upper()
            self.has_inserted = False
            super().__init__(pane, title='文字コードから人名・地名漢字を挿入')

        def body(self, pane):
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            t = '下記のURLで漢字を検索してください．\n' + \
                'https://moji.or.jp/mojikibansearch/basic\n\n' + \
                '「対応するUCS」の下の段を下に入力してください．\n' + \
                '例：花の場合→<82B1,E0102>\n'
            frm = tkinter.Frame(pane)
            frm.pack(side='top')
            txt = tkinter.Label(frm, text=t, justify='left')
            txt.pack(side='left')
            frm = tkinter.Frame(pane)
            frm.pack(side=tkinter.TOP)
            txt = tkinter.Label(frm, text='<')
            txt.pack(side=tkinter.LEFT)
            self.entry1 = tkinter.Entry(frm, width=7, font=font)
            self.entry1.pack(side=tkinter.LEFT)
            if self.code is not None:
                self.entry1.insert(0, self.code)
            txt = tkinter.Label(frm, text=',')
            txt.pack(side=tkinter.LEFT)
            self.entry2 = tkinter.Entry(frm, width=7, font=font)
            self.entry2.pack(side=tkinter.LEFT)
            txt = tkinter.Label(frm, text='>')
            txt.pack(side=tkinter.LEFT)
            # self.bind('<Key-Return>', self.ok)
            # self.bind('<Key-Escape>', self.cancel)
            # super().body(pane)
            if self.code is None:
                return self.entry1
            else:
                return self.entry2

        # def buttonbox(self):
        #     btn = tkinter.Frame(self)
        #     self.btn1 = tkinter.Button(btn, text='OK', width=6,
        #                                command=self.ok)
        #     self.btn1.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     self.btn2 = tkinter.Button(btn, text='Cancel', width=6,
        #                                command=self.cancel)
        #     self.btn2.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     btn.pack()

        def apply(self):
            ucs = self.entry1.get()
            ivs = self.entry2.get()
            if re.match('^[0-9a-fA-F]{4}$', ucs):
                self.pane.insert('insert', chr(int(ucs, 16)))
                if re.match('^E01[0-9a-eA-E][0-9a-fA-F]$', ivs):
                    i = int(ivs, 16) - 917760
                    self.pane.insert('insert', str(i) + ';')
                    self.has_inserted = True

    def insert_ivs_of_7947(self):
        self.txt.insert('insert',
                        'A祇2;' +  # E0102
                        'B祇3;')   # E0103

    def insert_ivs_of_82b1(self):
        self.txt.insert('insert',
                        'A花2;' +  # E0102
                        'B花3;' +  # E0103
                        'C花4;' +  # E0104
                        'D花6;')   # E0106

    def insert_ivs_of_845b(self):
        self.txt.insert('insert',
                        'A葛2;' +  # E0102
                        'B葛3;' +  # E0103
                        'C葛4;' +  # E0104
                        'D葛5;' +  # E0105
                        'E葛6;' +  # E0106
                        'F葛7;' +  # E0107
                        'G葛8;')   # E0108

    def insert_ivs_of_9089(self):
        self.txt.insert('insert',
                        'A邉15;' +  # E010F
                        'B邉16;' +  # E0110
                        'C邉17;' +  # E0111
                        'D邉18;' +  # E0112
                        'E邉19;' +  # E0113
                        'F邉20;' +  # E0114
                        'G邉21;' +  # E0115
                        'H邉22;' +  # E0116
                        'I邉23;' +  # E0117
                        'J邉24;' +  # E0118
                        'K邉25;' +  # E0119
                        'L邉26;' +  # E011A
                        'M邉27;' +  # E011B
                        'N邉28;' +  # E011C
                        'O邉29;' +  # E011D
                        'P邉31;')   # E011F

    def insert_ivs_of_908a(self):
        self.txt.insert('insert',
                        'A邊8;' +   # E0108
                        'B邊9;' +   # E0109
                        'C邊10;' +  # E010A
                        'D邊11;' +  # E010B
                        'E邊12;' +  # E010C
                        'F邊13;' +  # E010D
                        'G邊14;' +  # E010E
                        'H邊15;' +  # E010F
                        'I邊16;' +  # E0110
                        'J邊17;' +  # E0111
                        'K邊18;')   # E0112

    ################
    # SUBMENU INSERT TIME

    def _make_submenu_insert_time(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='日時を挿入', menu=submenu)
        #
        submenu.add_command(label='YY年M月D日',
                            command=self.insert_date_YYMD)
        submenu.add_command(label='令和Y年M月D日',
                            command=self.insert_date_GYMD)
        submenu.add_command(label='yy年m月d日',
                            command=self.insert_date_yymd)
        submenu.add_command(label='令和y年m月d日',
                            command=self.insert_date_Gymd)
        submenu.add_command(label='yyyy-mm-dd',
                            command=self.insert_date_iso)
        submenu.add_command(label='gyy-mm-dd',
                            command=self.insert_date_giso)
        submenu.add_separator()
        #
        submenu.add_command(label='H時M分S秒',
                            command=self.insert_time_HHMS)
        submenu.add_command(label='午前H時M分S秒',
                            command=self.insert_time_GHMS)
        submenu.add_command(label='h時m分s秒',
                            command=self.insert_time_hhms)
        submenu.add_command(label='午前h時m分s秒',
                            command=self.insert_time_Ghms)
        submenu.add_command(label='hh:mm:ss',
                            command=self.insert_time_iso)
        submenu.add_command(label='AMhh:mm:ss',
                            command=self.insert_time_giso)
        submenu.add_separator()
        #
        submenu.add_command(label='yyyy-mm-ddThh:mm:ss+09:00',
                            command=self.insert_datetime)
        submenu.add_command(label='yy-mm-ddThh:mm:ss',
                            command=self.insert_datetime_simple)

    ######
    # COMMAND

    def insert_date_YYMD(self):
        now = self._get_now()
        date = now.strftime('%Y年%m月%d日')
        date = self._remove_zero(date)
        date = self._convert_half_to_full(date)
        self.txt.insert('insert', date)

    def insert_date_GYMD(self):
        now = self._get_now()
        year = int(now.strftime('%Y')) - 2018
        date = '令和' + str(year) + '年' + now.strftime('%m月%d日')
        date = self._remove_zero(date)
        date = self._convert_half_to_full(date)
        self.txt.insert('insert', date)

    def insert_date_yymd(self):
        now = self._get_now()
        date = now.strftime('%Y年%m月%d日')
        date = self._remove_zero(date)
        self.txt.insert('insert', date)

    def insert_date_Gymd(self):
        now = self._get_now()
        year = int(now.strftime('%Y')) - 2018
        date = '令和' + str(year) + '年' + now.strftime('%m月%d日')
        date = self._remove_zero(date)
        self.txt.insert('insert', date)

    def insert_date_iso(self):
        now = self._get_now()
        date = now.strftime('%Y-%m-%d')
        self.txt.insert('insert', date)

    def insert_date_giso(self):
        now = self._get_now()
        year = int(now.strftime('%Y')) - 2018
        if year < 10:
            date = 'R0' + str(year) + '-' + now.strftime('%m-%d')
        else:
            date = 'R' + str(year) + '-' + now.strftime('%m-%d')
        self.txt.insert('insert', date)

    def insert_time_HHMS(self):
        now = self._get_now()
        time = now.strftime('%H時%M分%S秒')
        time = self._remove_zero(time)
        time = self._convert_half_to_full(time)
        self.txt.insert('insert', time)

    def insert_time_GHMS(self):
        now = self._get_now()
        hour = int(now.strftime('%H'))
        if hour < 12:
            time = '午前' + str(hour) + '時' + now.strftime('%M分%S秒')
        else:
            time = '午後' + str(hour - 12) + '時' + now.strftime('%M分%S秒')
        time = self._remove_zero(time)
        time = self._convert_half_to_full(time)
        self.txt.insert('insert', time)

    def insert_time_hhms(self):
        now = self._get_now()
        time = now.strftime('%H時%M分%S秒')
        time = self._remove_zero(time)
        self.txt.insert('insert', time)

    def insert_time_Ghms(self):
        now = self._get_now()
        hour = int(now.strftime('%H'))
        if hour < 12:
            time = '午前' + str(hour) + '時' + now.strftime('%M分%S秒')
        else:
            time = '午後' + str(hour - 12) + '時' + now.strftime('%M分%S秒')
        time = self._remove_zero(time)
        self.txt.insert('insert', time)

    def insert_time_iso(self):
        now = self._get_now()
        time = now.strftime('%H:%M:%S')
        self.txt.insert('insert', time)

    def insert_time_giso(self):
        now = self._get_now()
        hour = int(now.strftime('%H'))
        if hour < 12:
            time = 'AM' + str(hour) + ':' + now.strftime('%M:%S')
        else:
            time = 'PM' + str(hour - 12) + ':' + now.strftime('%M:%S')
        self.txt.insert('insert', time)

    def insert_datetime(self):
        now = self._get_now()
        self.txt.insert('insert', now.isoformat(timespec='seconds'))

    def insert_datetime_simple(self):
        now = self._get_now()
        self.txt.insert('insert', now.strftime('%y-%m-%dT%H:%M:%S'))

    @staticmethod
    def _remove_zero(text):
        text = re.sub('^0', '', text)
        text = re.sub('年0', '年', text)
        text = re.sub('月0', '月', text)
        text = re.sub('時0', '時', text)
        text = re.sub('分0', '分', text)
        return text

    ################
    # SUBMENU INSERT FILE

    def _make_submenu_insert_file_name(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='ファイル名を挿入', menu=submenu)
        #
        submenu.add_command(label='フルパスで挿入',
                            command=self.insert_file_paths)
        submenu.add_command(label='ファイル名のみを挿入',
                            command=self.insert_file_names)
        submenu.add_command(label='編集中のファイルと同じフォルダにある全ファイルのファイル名のみを挿入',
                            command=self.insert_file_names_in_same_folder)

    ######
    # COMMAND

    def insert_file_paths(self):
        file_paths = tkinter.filedialog.askopenfilenames()
        for f in file_paths:
            self.txt.insert('insert', f + '\n')

    def insert_file_names(self):
        file_paths = tkinter.filedialog.askopenfilenames()
        for f in file_paths:
            f = re.sub('^(.|\n)*/', '', f)
            self.txt.insert('insert', f + '\n')

    def insert_file_names_in_same_folder(self):
        file_path = self.file_path
        if file_path is None:
            return
        elif re.match('^(.*)[/\\\\](.*)$', file_path):
            dir_path = re.sub('^(.*)[/\\\\](.*)$', '\\1', file_path)
        else:
            dir_path = os.getcwd()
        files = os.listdir(dir_path)
        for f in sorted(files):
            if not re.match('^\\.', f) and os.path.isfile(f):
                if not re.match('^~\\$.*\\.zip$', f):
                    self.txt.insert('insert', f + '\n')

    ################
    # COMMAND

    def insert_file(self):
        file_path = tkinter.filedialog.askopenfilename()
        if file_path != () and file_path != '':
            with open(file_path, 'rb') as f:
                raw_data = f.read()
            encoding = self._get_encoding(raw_data)
            decoded_data = self._decode_data(encoding, raw_data)
            self.txt.insert('insert', decoded_data)

    def insert_symbol(self):
        candidates = ['⑴', '⑵', '⑶', '⑷', '⑸', '⑹', '⑺', '⑻', '⑼', '⑽',
                      '⑾', '⑿', '⒀', '⒁', '⒂', '⒃', '⒄', '⒅', '⒆', '⒇',
                      '⓪',
                      '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩',
                      '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰', '⑱', '⑲', '⑳',
                      '²', '³',
                      '㊞',
                      ]
        self.SymbolDialog(self.txt, self, candidates)

    class SymbolDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, candidates):
            self.pane = pane
            self.mother = mother
            self.candidates = candidates
            super().__init__(pane, title='記号を挿入')

        def body(self, pane):
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            self.symbol = tkinter.StringVar()
            for i, cnd in enumerate(self.candidates):
                rd = tkinter.Radiobutton(pane, text=cnd, font=font,
                                         variable=self.symbol, value=cnd)
                y, x = int(i / 10), (i % 10)
                rd.grid(row=y, column=x, columnspan=1, padx=3, pady=3)
            # self.bind('<Key-Return>', self.ok)
            # self.bind('<Key-Escape>', self.cancel)
            # super().body(pane)

        # def buttonbox(self):
        #     btn = tkinter.Frame(self)
        #     self.btn1 = tkinter.Button(btn, text='OK', width=6,
        #                                command=self.ok)
        #     self.btn1.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     self.btn2 = tkinter.Button(btn, text='Cancel', width=6,
        #                                command=self.cancel)
        #     self.btn2.pack(side=tkinter.LEFT, padx=3, pady=3)
        #     btn.pack()

        def apply(self):
            symbol = self.symbol.get()
            self.pane.insert('insert', symbol)
            # self.pane.mark_set('insert', 'insert-1c')
            self.pane.focus_set()

    ################
    # SUBMENU INSERT HORIZONTAL LINE

    def _make_submenu_insert_horizontal_line(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='横棒を挿入', menu=submenu)
        #
        submenu.add_command(label='"-"（002D）半角ハイフンマイナス',
                            command=self.insert_hline_002d)
        submenu.add_command(label='"‐"（2010）全角ハイフン',
                            command=self.insert_hline_2010)
        submenu.add_command(label='"—"（2014）全角Ｍダッシュ',
                            command=self.insert_hline_2014)
        submenu.add_command(label='"―"（2015）全角水平線',
                            command=self.insert_hline_2015)
        submenu.add_command(label='"−"（2212）全角マイナスサイン',
                            command=self.insert_hline_2212)
        submenu.add_command(label='"－"（FF0D）全角ハイフンマイナス',
                            command=self.insert_hline_ff0d)

    ######
    # COMMAND

    # "-"（002D）半角ハイフンマイナス
    # "‐"（2010）全角ハイフン
    # "—"（2014）全角Ｍダッシュ
    # "―"（2015）全角水平線
    # "−"（2212）全角マイナスサイン
    # "－"（FF0D）ハイフンマイナス

    def insert_hline_002d(self):
        self.txt.insert('insert', '\u002D')  # 半角ハイフンマイナス

    def insert_hline_2010(self):
        self.txt.insert('insert', '\u2010')  # 全角ハイフン

    def insert_hline_2014(self):
        self.txt.insert('insert', '\u2014')  # 全角Ｍダッシュ

    def insert_hline_2015(self):
        self.txt.insert('insert', '\u2015')  # 全角水平線

    def insert_hline_2212(self):
        self.txt.insert('insert', '\u2212')  # 全角マイナスサイン

    def insert_hline_ff0d(self):
        self.txt.insert('insert', '\uFF0D')  # 全角ハイフンマイナス

    ################
    # SUBMENU INSERT SAMPLE

    def _make_submenu_insert_sample(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='サンプルを挿入', menu=submenu)
        #
        submenu.add_command(label='基本',
                            command=self.insert_basis_sample)
        submenu.add_command(label='民法',
                            command=self.insert_law_sample)
        submenu.add_command(label='訴状',
                            command=self.insert_petition_sample)
        submenu.add_command(label='証拠説明書',
                            command=self.insert_evidence_sample)
        submenu.add_command(label='和解契約書',
                            command=self.insert_settlement_sample)

    ######
    # COMMAND

    def insert_basis_sample(self):
        document = self.insert_configuration_sample('普通', '0.0') + \
            SAMPLE_BASIS
        self.insert_sample(document)

    def insert_law_sample(self):
        document = self.insert_configuration_sample('条文', '0.0') + \
            SAMPLE_LAW
        self.insert_sample(document)

    def insert_petition_sample(self):
        document = self.insert_configuration_sample('普通', '1.0') + \
            SAMPLE_PETITION
        self.insert_sample(document)

    def insert_evidence_sample(self):
        document = self.insert_configuration_sample('普通', '0.0') + \
            SAMPLE_EVIDENCE
        self.insert_sample(document)

    def insert_settlement_sample(self):
        document = self.insert_configuration_sample('契約', '1.0') + \
            SAMPLE_SETTLEMENT
        self.insert_sample(document)

    def insert_configuration_sample(self, document_style, space_before):
        document = '''\
<!--------------------------【設定】-----------------------------

# プロパティに表示される文書のタイトルを指定できます。
書題名: -

# 3つの書式（普通、契約、条文）を指定できます。
文書式: ''' + document_style + '''

# 用紙のサイズ（A3横、A3縦、A4横、A4縦）を指定できます。
用紙サ: A4縦

# 用紙の上下左右の余白をセンチメートル単位で指定できます。
上余白: 3.5 cm
下余白: 2.2 cm
左余白: 3.0 cm
右余白: 2.0 cm

# ページのヘッダーに表示する文字列（別紙 :等）を指定できます。
頭書き:

# ページ番号の書式（無、有、n :、-n-、n/N等）を指定できます。
頁番号: 有

# 行番号の記載（無、有）を指定できます。
行番号: 無

# 明朝体とゴシック体と異字体（IVS）のフォントを指定できます。
明朝体: Times New Roman / ＭＳ 明朝
ゴシ体: = / ＭＳ ゴシック
異字体: IPAmj明朝

# 基本の文字の大きさをポイント単位で指定できます。
文字サ: 12 pt

# 行間隔を基本の文字の高さの何倍にするかを指定できます。
行間隔: 2.14 倍

# セクションタイトル前後の余白を行間隔の倍数で指定できます。
前余白: 0.0 倍, ''' + space_before + ''' 倍, 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍
後余白: 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍, 0.0 倍

# 半角文字と全角文字の間の間隔調整（無、有）を指定できます。
字間整: 無

# 備考書（コメント）などを消して完成させます。
完成稿: 偽

# 原稿の作成日時と更新日時が自動で記録されます。
作成時: - USER
更新時: - USER

---------------------------------------------------------------->
'''
        return document

    def insert_sample(self, sample_document):
        txt_text = self.txt.get('1.0', 'end-1c')
        if txt_text != '':
            n, m = 'エラー', 'テキストが空ではありません．'
            tkinter.messagebox.showerror(n, m)
            return
        self.file_lines = sample_document.split('\n')
        self.txt.insert('1.0', sample_document)
        self.txt.focus_set()
        self.txt.mark_set('insert', '1.0')
        # PAINT
        paint_keywords = self.paint_keywords.get()
        self.line_data = [LineDatum() for line in self.file_lines]
        for i, line in enumerate(self.file_lines):
            self.line_data[i].line_number = i
            self.line_data[i].line_text = line + '\n'
            if i > 0:
                self.line_data[i].beg_chars_state \
                    = self.line_data[i - 1].end_chars_state.copy()
                self.line_data[i].beg_chars_state.reset_partially()
            self.line_data[i].paint_line(self.txt, paint_keywords)
        # CLEAR THE UNDO STACK
        self.txt.edit_reset()

    ##########################
    # MENU PARAGRAPH

    def _make_menu_paragraph(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='段落(P)', menu=menu, underline=3)
        #
        menu.add_command(label='段落の余白の長さを設定',
                         command=self.set_paragraph_length)
        menu.add_separator()
        #
        self._make_submenu_insert_chapter(menu)
        self._make_submenu_insert_section(menu)
        self._make_submenu_insert_list(menu)
        menu.add_command(label='画像を挿入',
                         command=self.insert_image_paragraph)
        self._make_submenu_insert_table(menu)
        menu.add_command(label='改ページを挿入',
                         command=self.insert_page_break)
        menu.add_separator()
        #
        menu.add_command(label='チャプターの番号を変更',
                         command=self.set_chapter_number)
        menu.add_command(label='セクションの番号を変更',
                         command=self.set_section_number)
        menu.add_command(label='箇条書きの番号を変更',
                         command=self.set_list_number)
        # menu.add_separator()

    ################
    # COMMAND

    def set_paragraph_length(self):
        self.LengthRevisersDialog(self.txt, self)

    class LengthRevisersDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, length=None):
            self.pane = pane
            self.mother = mother
            bef_text = self.pane.get('1.0', 'insert')
            aft_text = self.pane.get('insert', 'end-1c')
            self.head_text \
                = re.sub('^((?:.|\n)*\n\n)((?:.|\n)*)?', '\\1', bef_text)
            bef_para = re.sub('^(.|\n)*\n\n', '', bef_text)
            aft_para = re.sub('\n\n(.|\n)*$', '', aft_text)
            paragraph = bef_para + aft_para
            res_length_reviser = '(?:v|V|X|<<|<|>)=[-\\+]?(?:[0-9]*\\.)?[0-9]+'
            res = '^((?:' + res_length_reviser + '(?:\\s|\n)*)*)((?:.|\n)*)$'
            self.length_revisers = re.sub(res, '\\1', paragraph)
            if length is not None:
                self.length = length
            else:
                self.length = {'space before': '0.0', 'space after': '0.0',
                               'line spacing': '0.0', 'first indent': '0.0',
                               'left indent': '0.0', 'right indent': '0.0'}
                res_bef = '(?:.|\n)*'
                res_aft = '=([-\\+]?(?:[0-9]*\\.)?[0-9]+)' + '(?:.|\n)*'
                res = res_bef + 'v' + res_aft
                if re.match(res, self.length_revisers):
                    self.length['space before'] \
                        = str(float(re.sub(res, '\\1', self.length_revisers)))
                res = res_bef + 'V' + res_aft
                if re.match(res, self.length_revisers):
                    self.length['space after'] \
                        = str(float(re.sub(res, '\\1', self.length_revisers)))
                res = res_bef + 'X' + res_aft
                if re.match(res, self.length_revisers):
                    self.length['line spacing'] \
                        = str(float(re.sub(res, '\\1', self.length_revisers)))
                res = res_bef + '<<' + res_aft
                if re.match(res, self.length_revisers):
                    self.length['first indent'] \
                        = str(float(re.sub(res, '\\1', self.length_revisers))
                              * -1)
                res = res_bef + '<' + res_aft
                if re.match(res, self.length_revisers):
                    self.length['left indent'] \
                        = str(float(re.sub(res, '\\1', self.length_revisers))
                              * -1)
                res = res_bef + '>' + res_aft
                if re.match(res, self.length_revisers):
                    self.length['right indent'] \
                        = str(float(re.sub(res, '\\1', self.length_revisers))
                              * -1)
            super().__init__(pane, title='段落の長さを設定')

        def body(self, pane):
            font_size = self.mother.font_size.get()
            f = (GOTHIC_FONT, font_size)
            self.title1 = tkinter.Label(pane, text='前の段落との間の幅')
            self.title1.grid(row=0, column=0)
            self.entry1 = tkinter.Entry(pane, width=7, font=f, justify='right')
            self.entry1.insert(0, self.length['space before'])
            self.entry1.grid(row=0, column=1)
            self.unit1 = tkinter.Label(pane, text='行間')
            self.unit1.grid(row=0, column=2)
            self.title2 = tkinter.Label(pane, text='次の段落との間の幅')
            self.title2.grid(row=1, column=0)
            self.entry2 = tkinter.Entry(pane, width=7, font=f, justify='right')
            self.entry2.insert(0, self.length['space after'])
            self.entry2.grid(row=1, column=1)
            self.unit2 = tkinter.Label(pane, text='行間')
            self.unit2.grid(row=1, column=2)
            self.title3 = tkinter.Label(pane, text='段落内の改行の幅　')
            self.title3.grid(row=2, column=0)
            self.entry3 = tkinter.Entry(pane, width=7, font=f, justify='right')
            self.entry3.insert(0, self.length['line spacing'])
            self.entry3.grid(row=2, column=1)
            self.unit3 = tkinter.Label(pane, text='行間')
            self.unit3.grid(row=2, column=2)
            self.title4 = tkinter.Label(pane, text='一行目の字下げの幅')
            self.title4.grid(row=3, column=0)
            self.entry4 = tkinter.Entry(pane, width=7, font=f, justify='right')
            self.entry4.insert(0, self.length['first indent'])
            self.entry4.grid(row=3, column=1)
            self.unit4 = tkinter.Label(pane, text='文字')
            self.unit4.grid(row=3, column=2)
            self.title5 = tkinter.Label(pane, text='左の字下げの幅　　')
            self.title5.grid(row=4, column=0)
            self.entry5 = tkinter.Entry(pane, width=7, font=f, justify='right')
            self.entry5.insert(0, self.length['left indent'])
            self.entry5.grid(row=4, column=1)
            self.unit5 = tkinter.Label(pane, text='文字')
            self.unit5.grid(row=4, column=2)
            self.title6 = tkinter.Label(pane, text='右の字下げの幅　　')
            self.title6.grid(row=5, column=0)
            self.entry6 = tkinter.Entry(pane, width=7, font=f, justify='right')
            self.entry6.insert(0, self.length['right indent'])
            self.entry6.grid(row=5, column=1)
            self.unit6 = tkinter.Label(pane, text='文字')
            self.unit6.grid(row=5, column=2)
            return self.entry1

        def apply(self):
            has_error = False
            res = '^[-\\+]?(?:[0-9]*\\.)?[0-9]+$'
            space_before = re.sub('\\s', '', self.entry1.get())
            if re.match(res, space_before):
                self.length['space before'] = space_before
            else:
                has_error = True
            space_after = re.sub('\\s', '', self.entry2.get())
            if re.match(res, space_after):
                self.length['space after'] = space_after
            else:
                has_error = True
            line_spacing = re.sub('\\s', '', self.entry3.get())
            if re.match(res, line_spacing):
                self.length['line spacing'] = line_spacing
            else:
                has_error = True
            first_indent = re.sub('\\s', '', self.entry4.get())
            if re.match(res, first_indent):
                self.length['first indent'] = first_indent
            else:
                has_error = True
            left_indent = re.sub('\\s', '', self.entry5.get())
            if re.match(res, left_indent):
                self.length['left indent'] = left_indent
            else:
                has_error = True
            right_indent = re.sub('\\s', '', self.entry6.get())
            if re.match(res, right_indent):
                self.length['right indent'] = right_indent
            else:
                has_error = True
            if has_error:
                n = 'エラー'
                m = '値に正負の小数以外が含まれています．'
                tkinter.messagebox.showerror(n, m)
                Makdo.LengthRevisersDialog(self.pane, self.length)
            else:
                len_beg = len(self.head_text)
                len_end = len(self.head_text + self.length_revisers)
                beg = '1.0+' + str(len_beg) + 'c'
                end = '1.0+' + str(len_end) + 'c'
                self.pane.delete(beg, end)
                leng_revs = ''
                leng = float(self.length['space before'])
                if leng > 0:
                    leng_revs += 'v=+' + re.sub('\\.0+$', '', str(leng)) + ' '
                else:
                    leng_revs += 'v=' + re.sub('\\.0+$', '', str(leng)) + ' '
                leng = float(self.length['space after'])
                if leng > 0:
                    leng_revs += 'V=+' + re.sub('\\.0+$', '', str(leng)) + ' '
                else:
                    leng_revs += 'V=' + re.sub('\\.0+$', '', str(leng)) + ' '
                leng = float(self.length['line spacing'])
                if leng > 0:
                    leng_revs += 'X=+' + re.sub('\\.0+$', '', str(leng)) + ' '
                elif leng < 0:
                    leng_revs += 'X=' + re.sub('\\.0+$', '', str(leng)) + ' '
                leng = float(self.length['first indent']) * -1
                if leng > 0:
                    leng_revs += '<<=+' + re.sub('\\.0+$', '', str(leng)) + ' '
                elif leng < 0:
                    leng_revs += '<<=' + re.sub('\\.0+$', '', str(leng)) + ' '
                leng = float(self.length['left indent']) * -1
                if leng > 0:
                    leng_revs += '<=+' + re.sub('\\.0+$', '', str(leng)) + ' '
                elif leng < 0:
                    leng_revs += '<=' + re.sub('\\.0+$', '', str(leng)) + ' '
                leng = float(self.length['right indent']) * -1
                if leng > 0:
                    leng_revs += '>=+' + re.sub('\\.0+$', '', str(leng)) + ' '
                elif leng < 0:
                    leng_revs += '>=' + re.sub('\\.0+$', '', str(leng)) + ' '
                leng_revs = re.sub(' $', '', leng_revs)
                self.pane.insert(beg, leng_revs + '\n')

    ################
    # SUBMENU INSERT CHAPTER

    def _make_submenu_insert_chapter(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='チャプターを挿入', menu=submenu)
        #
        submenu.add_command(label='第１編　…',
                            command=self.insert_chap_1)
        submenu.add_command(label='　第１章　…',
                            command=self.insert_chap_2)
        submenu.add_command(label='　　第１節　…',
                            command=self.insert_chap_3)
        submenu.add_command(label='　　　第１款　…',
                            command=self.insert_chap_4)
        submenu.add_command(label='　　　　第１目　…',
                            command=self.insert_chap_5)

    ######
    # COMMAND

    def insert_chap_1(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '$ ')  # 第1編

    def insert_chap_2(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '$$ ')  # 第1章

    def insert_chap_3(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '$$$ ')  # 第1節

    def insert_chap_4(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '$$$$ ')  # 第1款

    def insert_chap_5(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '$$$$$ ')  # 第1目

    ################
    # SUBMENU INSERT SECTION

    def _make_submenu_insert_section(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='セクションを挿入', menu=submenu)
        #
        submenu.add_command(label='（書面のタイトル）',
                            command=self.insert_sect_1)
        submenu.add_command(label='第１　…',
                            command=self.insert_sect_2)
        submenu.add_command(label='　１　…',
                            command=self.insert_sect_3)
        submenu.add_command(label='　　(1) …',
                            command=self.insert_sect_4)
        submenu.add_command(label='　　　ア　…',
                            command=self.insert_sect_5)
        submenu.add_command(label='　　　　(ｱ) …',
                            command=self.insert_sect_6)
        submenu.add_command(label='　　　　　ａ　…',
                            command=self.insert_sect_7)
        submenu.add_command(label='　　　　　　(a) …',
                            command=self.insert_sect_8)

    ######
    # COMMAND

    def insert_sect_1(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '# ')  # タイトル

    def insert_sect_2(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '## ')  # 第1

    def insert_sect_3(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '### ')  # 1

    def insert_sect_4(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '#### ')  # (1)

    def insert_sect_5(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '##### ')  # ア

    def insert_sect_6(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '###### ')  # (ｱ)

    def insert_sect_7(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '####### ')  # ａ

    def insert_sect_8(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '######## ')  # (a)

    ################
    # SUBMENU INSERT LIST

    def _make_submenu_insert_list(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='箇条書きを挿入', menu=submenu)
        #
        submenu.add_command(label='①　…',
                            command=self.insert_nlist_1)
        submenu.add_command(label='　㋐　…',
                            command=self.insert_nlist_2)
        submenu.add_command(label='　　ⓐ　…',
                            command=self.insert_nlist_3)
        submenu.add_command(label='　　　㊀　…',
                            command=self.insert_nlist_4)
        submenu.add_separator()
        #
        submenu.add_command(label='・　…',
                            command=self.insert_blist_1)
        submenu.add_command(label='　○　…',
                            command=self.insert_blist_2)
        submenu.add_command(label='　　△　…',
                            command=self.insert_blist_3)
        submenu.add_command(label='　　　◇　…',
                            command=self.insert_blist_4)

    ######
    # COMMAND

    def insert_nlist_1(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '1. ')

    def insert_nlist_2(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '  1. ')

    def insert_nlist_3(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '    1. ')

    def insert_nlist_4(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '      1. ')

    def insert_blist_1(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '- ')

    def insert_blist_2(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '  - ')

    def insert_blist_3(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '    - ')

    def insert_blist_4(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '      - ')

    ################
    # COMMAND

    def insert_image_paragraph(self):
        typ = [('画像', '.jpg .jpeg .png .gif .tif .tiff .bmp'),
               ('全てのファイル', '*')]
        image_path = tkinter.filedialog.askopenfilename(filetypes=typ)
        if image_path != () and image_path != '':
            self._insert_line_break_as_necessary()
            image_md_text = '![代替テキスト:縦x横](' + image_path + ' "説明")'
            self.txt.insert('insert', image_md_text)

    def insert_page_break(self):
        self._insert_line_break_as_necessary()
        self.txt.insert('insert', '<pgbr>')

    ################
    # SUBMENU INSERT TABLE

    def _make_submenu_insert_table(self, menu):
        submenu = tkinter.Menu(menu, tearoff=False)
        menu.add_cascade(label='表を挿入', menu=submenu)
        submenu.add_command(label='エクセルから挿入',
                            command=self.insert_table_from_excel)
        submenu.add_command(label='書式を挿入',
                            command=self.insert_table_format)

    ######
    # COMMAND

    def insert_table_from_excel(self, file_path=None):
        if file_path is None:
            typ = [('エクセル', '.xlsx')]
            file_path = tkinter.filedialog.askopenfilename(filetypes=typ)
        wb = openpyxl.load_workbook(file_path)
        for sheet_name in wb.sheetnames:
            self.txt.insert('insert', '<!-- ' + sheet_name + ' -->\n')
            ws = wb[sheet_name]
            table = ''
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                    min_col=1, max_col=ws.max_column):
                for cell in row:
                    table += '|' + str(cell.value)
                table += '|\n'
            self.txt.insert('insert', table + '\n')

    def insert_table_format(self):
        self._insert_line_break_as_necessary()
        table_md_text = ''
        table_md_text += '|タイトル  |タイトル  |タイトル  |=\n'
        table_md_text += '|:---------|:--------:|---------:|\n'
        table_md_text += '|左寄せセル|中寄せセル|右寄せセル|\n'
        table_md_text += '|左寄せセル|中寄せセル|右寄せセル|'
        self.txt.insert('insert', table_md_text)

    def set_chapter_number(self):
        self.ChapterNumberDialog(self.txt, self)

    class ChapterNumberDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, cnd=[-1, -1, -1, -1, -1]):
            self.pane = pane
            self.mother = mother
            self.cnd = cnd
            super().__init__(pane, title='チャプターの番号を変更')

        def body(self, pane):
            self.entry1 = self._body(pane, 0, '編', self.cnd[0])
            self.entry2 = self._body(pane, 1, '章', self.cnd[1])
            self.entry3 = self._body(pane, 2, '節', self.cnd[2])
            self.entry4 = self._body(pane, 3, '款', self.cnd[3])
            self.entry5 = self._body(pane, 4, '目', self.cnd[4])
            return self.entry1

        def _body(self, pane, row, unit, cnd):
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            head = tkinter.Label(pane, text='第１' + unit + '　→　第')
            head.grid(row=row, column=0)
            entry = tkinter.Entry(pane, width=4, justify='center', font=font)
            entry.grid(row=row, column=1)
            if cnd >= 0:
                entry.insert(0, str(cnd))
            tail = tkinter.Label(pane, text=unit)
            tail.grid(row=row, column=2)
            return entry

        def apply(self):
            str1 = self.entry1.get()
            int1, err1 = self._apply(str1)
            str2 = self.entry2.get()
            int2, err2 = self._apply(str2)
            str3 = self.entry3.get()
            int3, err3 = self._apply(str3)
            str4 = self.entry4.get()
            int4, err4 = self._apply(str4)
            str5 = self.entry5.get()
            int5, err5 = self._apply(str5)
            if err1 or err2 or err3 or err4 or err5:
                Makdo.ChapterNumberDialog(self.pane,
                                          [int1, int2, int3, int4, int5])
            else:
                doc = self.pane.get('1.0', 'insert')
                res = '^(' \
                    + '((.|\n)*\n\n)?' \
                    + '(((v|V|X|<<|<|>)=[-\\+]?[0-9]+\\s*)*\n)?' \
                    + ')(.|\n)*$'
                doc = re.sub(res, '\\1', doc)
                ins = ''
                if int1 >= 0:
                    ins += '$=' + str(int1) + ' '
                if int2 >= 0:
                    ins += '$$=' + str(int2) + ' '
                if int3 >= 0:
                    ins += '$$$=' + str(int3) + ' '
                if int4 >= 0:
                    ins += '$$$$=' + str(int4) + ' '
                if int5 >= 0:
                    ins += '$$$$$=' + str(int5) + ' '
                if ins != '':
                    ins = re.sub('\\s+$', '\n', ins)
                    self.pane.insert('1.0+' + str(len(doc)) + 'c', ins)

        def _apply(self, strn):
            if strn == '':
                return -1, False
            intn = c2n_n_arab(strn)
            if intn == -1:
                return -1, True
            return intn, False

    def set_section_number(self):
        self.SectionNumberDialog(self.txt, self)

    class SectionNumberDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, cnd=['', '', '', '', '', '', '']):
            self.pane = pane
            self.mother = mother
            self.cnd = cnd
            super().__init__(pane, title='セクションの番号を変更')

        def body(self, pane):
            self.entry1 = self._body(pane, 0, '第', '１', '', self.cnd[0])
            self.entry2 = self._body(pane, 1, '', '１', '', self.cnd[1])
            self.entry3 = self._body(pane, 2, '（', '1', '）', self.cnd[2])
            self.entry4 = self._body(pane, 3, '', 'ア', '', self.cnd[3])
            self.entry5 = self._body(pane, 4, '（', 'ｱ', '）', self.cnd[4])
            self.entry6 = self._body(pane, 5, '', 'ａ', '', self.cnd[5])
            self.entry7 = self._body(pane, 6, '（', 'a', '）', self.cnd[6])
            return self.entry1

        def _body(self, pane, row, pre, num, pos, cnd):
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            txt = tkinter.Label(pane, text=pre + num + pos)
            txt.grid(row=row, column=0)
            txt = tkinter.Label(pane, text='　→　')
            txt.grid(row=row, column=1)
            txt = tkinter.Label(pane, text=pre)
            txt.grid(row=row, column=2)
            entry = tkinter.Entry(pane, width=4, justify='center', font=font)
            entry.grid(row=row, column=3)
            if cnd is not None:
                entry.insert(0, str(cnd))
            txt = tkinter.Label(pane, text=pos)
            txt.grid(row=row, column=4)
            return entry

        def apply(self):
            str1 = self.entry1.get()
            str1, int1, err1 = self._apply(str1, 'arab')
            str2 = self.entry2.get()
            str2, int2, err2 = self._apply(str2, 'arab')
            str3 = self.entry3.get()
            str3, int3, err3 = self._apply(str3, 'arab')
            str4 = self.entry4.get()
            str4, int4, err4 = self._apply(str4, 'kata')
            str5 = self.entry5.get()
            str5, int5, err5 = self._apply(str5, 'kata')
            str6 = self.entry6.get()
            str6, int6, err6 = self._apply(str6, 'alph')
            str7 = self.entry7.get()
            str7, int7, err7 = self._apply(str7, 'alph')
            if err1 or err2 or err3 or err4 or err5 or err6 or err7:
                lst = [str1, str2, str3, str4, str5, str6, str7]
                Makdo.SectionNumberDialog(self.pane, lst)
            else:
                doc = self.pane.get('1.0', 'insert')
                res = '^(' \
                    + '((.|\n)*\n\n)?' \
                    + '(((v|V|X|<<|<|>)=[-\\+]?[0-9]+\\s*)*\n)?' \
                    + ')(.|\n)*$'
                doc = re.sub(res, '\\1', doc)
                ins = ''
                if int1 >= 0:
                    ins += '#=' + str(int1) + ' '
                if int2 >= 0:
                    ins += '##=' + str(int2) + ' '
                if int3 >= 0:
                    ins += '###=' + str(int3) + ' '
                if int4 >= 0:
                    ins += '####=' + str(int4) + ' '
                if int5 >= 0:
                    ins += '#####=' + str(int5) + ' '
                if int6 >= 0:
                    ins += '######=' + str(int6) + ' '
                if int7 >= 0:
                    ins += '#######=' + str(int7) + ' '
                if ins != '':
                    ins = re.sub('\\s+$', '\n', ins)
                    self.pane.insert('1.0+' + str(len(doc)) + 'c', ins)

        def _apply(self, strn, kind):
            if strn == '':
                return '', -1, False
            if kind == 'arab':
                intn = c2n_n_arab(strn)
            elif kind == 'kata':
                intn = c2n_n_kata(strn)
            elif kind == 'alph':
                intn = c2n_n_alph(strn)
            if intn == -1:
                return '', -1, True
            return strn, intn, False

    def set_list_number(self):
        self.ListNumberDialog(self.txt, self)

    class ListNumberDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, cnd=['', '', '', '']):
            self.pane = pane
            self.mother = mother
            self.cnd = cnd
            super().__init__(pane, title='箇条書きの番号を変更')

        def body(self, pane):
            self.entry1 = self._body(pane, 0, '①', self.cnd[0])
            self.entry2 = self._body(pane, 1, '㋐', self.cnd[1])
            self.entry3 = self._body(pane, 2, 'ⓐ', self.cnd[2])
            self.entry4 = self._body(pane, 3, '㊀', self.cnd[3])
            return self.entry1

        def _body(self, pane, row, num, cnd):
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            txt = tkinter.Label(pane, text=num)
            txt.grid(row=row, column=0)
            txt = tkinter.Label(pane, text='　→　')
            txt.grid(row=row, column=1)
            txt = tkinter.Label(pane, text='（')
            txt.grid(row=row, column=2)
            entry = tkinter.Entry(pane, width=4, justify='center', font=font)
            entry.grid(row=row, column=3)
            if cnd is not None:
                entry.insert(0, str(cnd))
            txt = tkinter.Label(pane, text='）')
            txt.grid(row=row, column=4)
            return entry

        def apply(self):
            str1 = self.entry1.get()
            str1, int1, err1 = self._apply(str1, 'arab')
            str2 = self.entry2.get()
            str2, int2, err2 = self._apply(str2, 'kata')
            str3 = self.entry3.get()
            str3, int3, err3 = self._apply(str3, 'alph')
            str4 = self.entry4.get()
            str4, int4, err4 = self._apply(str4, 'kanj')
            if err1 or err2 or err3 or err4:
                Makdo.ListNumberDialog(self.pane, [str1, str2, str3, str4])
            else:
                doc = self.pane.get('1.0', 'insert')
                res = '^(' \
                    + '((.|\n)*\n\n)?' \
                    + '(((v|V|X|<<|<|>)=[-\\+]?[0-9]+\\s*)*\n)?' \
                    + ')(.|\n)*$'
                doc = re.sub(res, '\\1', doc)
                ins = ''
                if int1 >= 0:
                    ins += '1.=' + str(int1) + '\n'
                if int2 >= 0:
                    ins += '  1.=' + str(int2) + '\n'
                if int3 >= 0:
                    ins += '    1.=' + str(int3) + '\n'
                if int4 >= 0:
                    ins += '      1.=' + str(int4) + '\n'
                if ins != '':
                    self.pane.insert('1.0+' + str(len(doc)) + 'c', ins)

        def _apply(self, strn, kind):
            if strn == '':
                return '', -1, False
            if kind == 'arab':
                intn = c2n_n_arab(strn)
            elif kind == 'kata':
                intn = c2n_n_kata(strn)
            elif kind == 'alph':
                intn = c2n_n_alph(strn)
            elif kind == 'kanj':
                intn = c2n_n_kanj(strn)
            if intn == -1:
                return '', -1, True
            return strn, intn, False

    ##########################
    # MENU MOVE

    def _make_menu_move(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='移動(M)', menu=menu, underline=3)
        #
        menu.add_command(label='文頭に移動',
                         command=self.goto_beg_of_doc)
        menu.add_command(label='文末に移動',
                         command=self.goto_end_of_doc)
        menu.add_command(label='行頭に移動',
                         command=self.goto_beg_of_line)
        menu.add_command(label='行末に移動',
                         command=self.goto_end_of_line)
        menu.add_separator()
        #
        self._make_submenu_place_flag(menu)
        self._make_submenu_goto_flag(menu)
        menu.add_separator()
        #
        menu.add_command(label='行数・文字数を指定して移動',
                         command=self.goto_by_position)
        # menu.add_separator()

    ################
    # COMMAND

    def goto_beg_of_doc(self):
        self.txt.mark_set('insert', '1.0')

    def goto_end_of_doc(self):
        self.txt.mark_set('insert', 'end-1c')

    def goto_beg_of_line(self):
        self.txt.mark_set('insert', 'insert linestart')

    def goto_end_of_line(self):
        self.txt.mark_set('insert', 'insert lineend')

    ################
    # SUBMENU PLACE FLAG

    def _make_submenu_place_flag(self, menu):
        submenu = tkinter.Menu(self.mnb, tearoff=False)
        menu.add_cascade(label='フラグを設置', menu=submenu)
        #
        submenu.add_command(label='フラグ１を設置',
                            command=self.place_flag1)
        submenu.add_command(label='フラグ２を設置',
                            command=self.place_flag2)
        submenu.add_command(label='フラグ３を設置',
                            command=self.place_flag3)
        submenu.add_command(label='フラグ４を設置',
                            command=self.place_flag4)
        submenu.add_command(label='フラグ５を設置',
                            command=self.place_flag5)

    #######
    # COMMAND

    def place_flag1(self):
        if 'flag1' in self.txt.mark_names():
            self.txt.mark_unset('flag1')
        self.txt.mark_set('flag1', 'insert')

    def place_flag2(self):
        if 'flag2' in self.txt.mark_names():
            self.txt.mark_unset('flag2')
        self.txt.mark_set('flag2', 'insert')

    def place_flag3(self):
        if 'flag3' in self.txt.mark_names():
            self.txt.mark_unset('flag3')
        self.txt.mark_set('flag3', 'insert')

    def place_flag4(self):
        if 'flag4' in self.txt.mark_names():
            self.txt.mark_unset('flag4')
        self.txt.mark_set('flag4', 'insert')

    def place_flag5(self):
        if 'flag5' in self.txt.mark_names():
            self.txt.mark_unset('flag5')
        self.txt.mark_set('flag5', 'insert')

    ################
    # SUBMENU GOTO FLAG

    def _make_submenu_goto_flag(self, menu):
        submenu = tkinter.Menu(self.mnb, tearoff=False)
        menu.add_cascade(label='フラグに移動', menu=submenu)
        #
        submenu.add_command(label='フラグ１に移動',
                            command=self.goto_flag1)
        submenu.add_command(label='フラグ２に移動',
                            command=self.goto_flag2)
        submenu.add_command(label='フラグ３に移動',
                            command=self.goto_flag3)
        submenu.add_command(label='フラグ４に移動',
                            command=self.goto_flag4)
        submenu.add_command(label='フラグ５に移動',
                            command=self.goto_flag5)

    #######
    # COMMAND

    def goto_flag1(self):
        if 'flag1' not in self.txt.mark_names():
            n, m = 'エラー', 'フラグ１は設定されていません．'
            tkinter.messagebox.showerror(n, m)
            return
        self.txt.mark_set('insert', 'flag1')

    def goto_flag2(self):
        if 'flag2' not in self.txt.mark_names():
            n, m = 'エラー', 'フラグ２は設定されていません．'
            tkinter.messagebox.showerror(n, m)
            return
        self.txt.mark_set('insert', 'flag2')

    def goto_flag3(self):
        if 'flag3' not in self.txt.mark_names():
            n, m = 'エラー', 'フラグ３は設定されていません．'
            tkinter.messagebox.showerror(n, m)
            return
        self.txt.mark_set('insert', 'flag3')

    def goto_flag4(self):
        if 'flag4' not in self.txt.mark_names():
            n, m = 'エラー', 'フラグ４は設定されていません．'
            tkinter.messagebox.showerror(n, m)
            return
        self.txt.mark_set('insert', 'flag4')

    def goto_flag5(self):
        if 'flag5' not in self.txt.mark_names():
            n, m = 'エラー', 'フラグ５は設定されていません．'
            tkinter.messagebox.showerror(n, m)
            return
        self.txt.mark_set('insert', 'flag5')

    def goto_by_position(self):
        self.PositionDialog(self.txt, self)

    class PositionDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother):
            self.pane = pane
            self.mother = mother
            super().__init__(pane, title='行数・文字数を指定して移動')

        def body(self, pane):
            t = '行数・文字数を入力してください．\n'
            self.text1 = tkinter.Label(pane, text=t)
            self.text1.pack(side=tkinter.TOP, anchor=tkinter.W)
            self.frame = tkinter.Frame(pane)
            self.frame.pack(side=tkinter.TOP)
            font_size = self.mother.font_size.get()
            font = (GOTHIC_FONT, font_size)
            self.entry1 = tkinter.Entry(self.frame, width=7, font=font)
            self.entry1.pack(side='left')
            tkinter.Label(self.frame, text='行目').pack(side='left')
            self.entry2 = tkinter.Entry(self.frame, width=7, font=font)
            self.entry2.pack(side='left')
            tkinter.Label(self.frame, text='文字目').pack(side='left')
            # self.bind('<Key-Return>', self.ok)
            # self.bind('<Key-Escape>', self.cancel)
            # super().body(pane)
            return self.entry1

        def apply(self):
            line = self.entry1.get()
            char = self.entry2.get()
            if re.match('^[0-9]+$', line) and re.match('^[0-9]+$', char):
                self.pane.mark_set('insert', line + '.' + char)

    ##########################
    # MENU TOOL

    def _make_menu_tool(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='ツール(T)', menu=menu, underline=4)
        #
        menu.add_command(label='定型句を挿入',
                         command=self.insert_formula)
        menu.add_command(label='定型句を編集',
                         command=self.edit_formula)
        menu.add_separator()
        #
        menu.add_command(label='メモ帳を開く',
                         command=self.open_memo_pad)
        menu.add_separator()
        #
        menu.add_command(label='画面を二つに分割・統合',
                         command=self.split_or_unify_window)
        menu.add_separator()
        #
        menu.add_command(label='編集前の原稿と比較して元に戻す',
                         command=self.compare_with_previous_draft)
        menu.add_command(label='別ファイルと内容を比較して反映',
                         command=self.compare_files)
        menu.add_separator()
        #
        menu.add_command(label='セクションを折り畳む・展開',
                         command=self.fold_or_unfold_section)
        menu.add_command(label='セクションを全て展開',
                         command=self.unfold_section_fully)
        menu.add_separator()
        #
        menu.add_command(label='OpenAIに質問',
                         command=self.ask_openai)
        menu.add_command(label='OpenAIの履歴を消去',
                         command=self.reset_openai)
        menu.add_separator()
        #
        menu.add_command(label='キーボードマクロを実行',
                         command=self.execute_keyboard_macro,
                         accelerator='Ctrl+E')
        menu.add_separator()
        #
        menu.add_command(label='コマンドを入力して実行',
                         command=self.start_minibuffer,
                         accelerator='Esc+X')
        # menu.add_separator()

    ################
    # COMMAND

    # INSERT AND EDIT FORMULA

    def insert_formula(self):
        t = '定型句を挿入'
        m = '挿入する定型句を選んでください．'
        fd = self.FormulaDialog(self.txt, self, t, m)
        self.formula_number = fd.get_value()
        self._insert_formula()

    def _insert_formula(self):
        n = self.formula_number
        formula_path = CONFIG_DIR + '/formula' + str(n) + '.md'
        try:
            with open(formula_path, 'r') as f:
                a = f.read()
        except BaseException:
            return
        self.txt.insert('insert', a)

    def insert_formula1(self):
        self.formula_number = 1
        self._insert_formula()

    def insert_formula2(self):
        self.formula_number = 2
        self._insert_formula()

    def insert_formula3(self):
        self.formula_number = 3
        self._insert_formula()

    def insert_formula4(self):
        self.formula_number = 4
        self._insert_formula()

    def insert_formula5(self):
        self.formula_number = 5
        self._insert_formula()

    def edit_formula(self):
        self.quit_editing_formula()
        t = '定型句を編集'
        m = '編集する定型句を選んでください．'
        fd = self.FormulaDialog(self.txt, self, t, m)
        self.formula_number = fd.get_value()
        self._edit_formula()

    def _edit_formula(self):
        if self.formula_number <= 0:
            return False
        n = self.formula_number
        formula_path = CONFIG_DIR + '/formula' + str(n) + '.md'
        if not os.path.exists(formula_path):
            open(formula_path, 'w').close()
        try:
            with open(formula_path, 'r') as f:
                formula = f.read()
        except BaseException:
            return
        #
        self.close_memo_pad()
        self.pnd.update()
        half_height = int(self.pnd.winfo_height() / 2) - 5
        self.pnd.remove(self.pnd1)
        self.pnd.remove(self.pnd2)
        self.pnd.remove(self.pnd3)
        self.pnd.remove(self.pnd4)
        self.pnd.remove(self.pnd5)
        self.pnd.remove(self.pnd6)
        self.pnd.add(self.pnd1, height=half_height, minsize=100)
        self.pnd.add(self.pnd2, height=half_height)
        self.pnd.update()
        #
        self.sub.pack(expand=True, fill=tkinter.BOTH)
        for key in self.txt.configure():
            self.sub.configure({key: self.txt.cget(key)})
        self.sub_btn.pack()
        #
        self.sub.delete('1.0', 'end')
        self.sub.insert('1.0', formula)
        self.sub.mark_set('insert', '1.0')
        self.txt.focus_force()
        # self.sub.focus_set()

    def edit_formula1(self):
        self.quit_editing_formula()
        self.formula_number = 1
        self._edit_formula()

    def edit_formula2(self):
        self.quit_editing_formula()
        self.formula_number = 2
        self._edit_formula()

    def edit_formula3(self):
        self.quit_editing_formula()
        self.formula_number = 3
        self._edit_formula()

    def edit_formula4(self):
        self.quit_editing_formula()
        self.formula_number = 4
        self._edit_formula()

    def edit_formula5(self):
        self.quit_editing_formula()
        self.formula_number = 5
        self._edit_formula()

    def quit_editing_formula(self):
        if self.formula_number <= 0:
            return False
        n = self.formula_number
        self.formula_number = -1
        formula_path = CONFIG_DIR + '/formula' + str(n) + '.md'
        try:
            os.rename(formula_path, formula_path + '~')
        except BaseException:
            pass
        try:
            with open(formula_path, 'w') as f:
                a = self.sub.get('1.0', 'end-1c')
                f.write(a)
        except BaseException:
            return False
        return True

    class FormulaDialog(tkinter.simpledialog.Dialog):

        def __init__(self, pane, mother, title, prompt):
            self.pane = pane
            self.mother = mother
            self.prompt = prompt
            self.value = None
            super().__init__(pane, title=title)

        def body(self, pane):
            prompt = tkinter.Label(pane, text=self.prompt)
            prompt.pack(side='top', anchor='w')
            self.value = tkinter.IntVar()
            self.value.set(1)
            rb1 = tkinter.Radiobutton(pane, text=self.get_head(1),
                                      variable=self.value, value=1)
            rb1.pack(side='top', anchor='w')
            rb2 = tkinter.Radiobutton(pane, text=self.get_head(2),
                                      variable=self.value, value=2)
            rb2.pack(side='top', anchor='w')
            rb3 = tkinter.Radiobutton(pane, text=self.get_head(3),
                                      variable=self.value, value=3)
            rb3.pack(side='top', anchor='w')
            rb4 = tkinter.Radiobutton(pane, text=self.get_head(4),
                                      variable=self.value, value=4)
            rb4.pack(side='top', anchor='w')
            rb5 = tkinter.Radiobutton(pane, text=self.get_head(5),
                                      variable=self.value, value=5)
            rb5.pack(side='top', anchor='w')
            super().body(pane)
            return rb1

        def get_head(self, n):
            try:
                with open(CONFIG_DIR + '/formula' + str(n) + '.md', 'r') as f:
                    a = f.read()
                    h = re.sub('\n', ' ', a)
                    if len(h) > 15:
                        h = h[:14] + '…'
                    return h
            except BaseException:
                return '（空）'

        def apply(self):
            pass

        def get_value(self):
            return self.value.get()

    # OPEN MEMO PAD

    def open_memo_pad(self):
        if CONFIG_DIR is None:
            return False
        memo_pad_path = CONFIG_DIR + '/memo.md'
        if not os.path.exists(memo_pad_path):
            try:
                open(memo_pad_path, 'w').close()
            except BaseException:
                return False
        if not os.path.exists(memo_pad_path):
            return False
        try:
            with open(memo_pad_path, 'r') as f:
                self.memo_pad_memory = f.read()
        except BaseException:
            return False
        #
        self.quit_editing_formula()
        self.pnd.update()
        half_height = int(self.pnd.winfo_height() / 2) - 5
        self.pnd.remove(self.pnd1)
        self.pnd.remove(self.pnd2)
        self.pnd.remove(self.pnd3)
        self.pnd.remove(self.pnd4)
        self.pnd.remove(self.pnd5)
        self.pnd.remove(self.pnd6)
        self.pnd.add(self.pnd1, height=half_height, minsize=100)
        self.pnd.add(self.pnd2, height=half_height)
        self.pnd.update()
        #
        self.sub.pack(expand=True, fill=tkinter.BOTH)
        for key in self.txt.configure():
            self.sub.configure({key: self.txt.cget(key)})
        self.sub_btn.pack()
        #
        self.sub.delete('1.0', 'end')
        self.sub.insert('1.0', self.memo_pad_memory)
        self.sub.mark_set('insert', '1.0')
        self.txt.focus_force()

    def update_memo_pad(self):
        memo_pad_memory = self.memo_pad_memory
        if self.memo_pad_memory is None:
            return False
        memo_pad_path = CONFIG_DIR + '/memo.md'
        # DISPLAY
        memo_pad_display = self.sub.get('1.0', 'end-1c')
        if memo_pad_display != memo_pad_memory:
            # MEMORY
            self.memo_pad_memory = memo_pad_display
            # FILE
            try:
                os.rename(memo_pad_path, memo_pad_path + '~')
            except BaseException:
                pass
            try:
                with open(memo_pad_path, 'w') as f:
                    f.write(memo_pad_display)
            except BaseException:
                return False
            return True
        # FILE
        if not os.path.exists(memo_pad_path):
            return False
        try:
            with open(memo_pad_path, 'r') as f:
                memo_pad_file = f.read()
        except BaseException:
            return False
        if memo_pad_file != memo_pad_memory:
            # MEMORY
            self.memo_pad_memory = memo_pad_file
            # DISPLAY
            self.sub.delete('1.0', 'end')
            self.sub.insert('1.0', memo_pad_file)

    def close_memo_pad(self):
        if self.memo_pad_memory is not None:
            self.update_memo_pad()
            self.memo_pad_memory = None

    # SPLIT OR UNIFY WINDOW

    def split_or_unify_window(self):
        if len(self.pnd.panes()) == 1:
            self._split_window()
        else:
            self._unify_window()

    def _split_window(self):
        self.quit_editing_formula()
        self.close_memo_pad()
        self.pnd.update()
        half_height = int(self.pnd.winfo_height() / 2) - 5
        self.pnd.remove(self.pnd1)
        self.pnd.remove(self.pnd2)
        self.pnd.remove(self.pnd3)
        self.pnd.remove(self.pnd4)
        self.pnd.remove(self.pnd5)
        self.pnd.remove(self.pnd6)
        self.pnd.add(self.pnd1, height=half_height, minsize=100)
        self.pnd.add(self.pnd2, height=half_height)
        self.pnd.update()
        #
        self.sub.pack(expand=True, fill=tkinter.BOTH)
        for key in self.txt.configure():
            self.sub.configure({key: self.txt.cget(key)})
        self.sub_btn.pack()
        #
        doc = self.txt.get('1.0', 'end-1c')
        self.sub.delete('1.0', 'end')
        self.sub.insert('1.0', doc)
        self.sub.mark_set('insert', '1.0')
        # self.sub.configure(state='disabled')
        #
        self.txt.focus_force()

    def _unify_window(self):
        self.quit_editing_formula()
        self.update_memo_pad()
        self.memo_pad_memory = None
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        self.pnd.remove(self.pnd2)
        self.txt.focus_set()

    # COMPARE

    # MDDIFF>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    def compare_with_previous_draft(self):
        importlib.reload(makdo.makdo_mddiff)
        text2 = self.init_text
        file2 = makdo.makdo_mddiff.File()
        file2.set_up_from_text(text2)
        file2.cmp_paragraphs \
            = makdo.makdo_mddiff.File.reset_configs(file2.cmp_paragraphs)
        para2 = file2.cmp_paragraphs
        self._compare_files_loop(para2)

    def compare_files(self):
        importlib.reload(makdo.makdo_mddiff)
        text2 = self._get_text2()
        file2 = makdo.makdo_mddiff.File()
        file2.set_up_from_text(text2)
        file2.cmp_paragraphs \
            = makdo.makdo_mddiff.File.reset_configs(file2.cmp_paragraphs)
        para2 = file2.cmp_paragraphs
        self._compare_files_loop(para2)

    def _get_text2(self):
        typ = [('可能な形式', '.md .docx'),
               ('Markdown', '.md'), ('MS Word', '.docx'),
               ('全てのファイル', '*')]
        file_path = tkinter.filedialog.askopenfilename(filetypes=typ)
        if file_path == () or file_path == '':
            return False
        # DOCX OR MD
        if re.match('^(?:.|\n)+.docx$', file_path):
            self.temp_dir = tempfile.TemporaryDirectory()
            md_path = self.temp_dir.name + '/doc.md'
        else:
            md_path = file_path
        # OPEN DOCX FILE
        if re.match('^(?:.|\n)+.docx$', file_path):
            stderr = sys.stderr
            sys.stderr = tempfile.TemporaryFile(mode='w+')
            importlib.reload(makdo.makdo_docx2md)
            try:
                d2m = makdo.makdo_docx2md.Docx2Md(file_path)
                d2m.save(md_path)
            except BaseException:
                pass
            sys.stderr.seek(0)
            msg = sys.stderr.read()
            sys.stderr = stderr
            if msg != '':
                n = 'エラー'
                tkinter.messagebox.showerror(n, msg)
                return
        # OPEN MD FILE
        try:
            with open(md_path, 'rb') as f:
                raw_data = f.read()
        except BaseException:
            return
        encoding = self._get_encoding(raw_data)
        decoded_data = self._decode_data(encoding, raw_data)
        return decoded_data

    def _compare_files_loop(self, para2):
        text1 = self.txt.get('1.0', 'end-1c')
        file1 = makdo.makdo_mddiff.File()
        file1.set_up_from_text(text1)
        #
        configs = makdo.makdo_mddiff.File.get_configs(file1.raw_paragraphs)
        file1.cmp_paragraphs \
            = makdo.makdo_mddiff.File.reset_configs(file1.cmp_paragraphs)
        #
        para1 = file1.cmp_paragraphs
        comp = makdo.makdo_mddiff.Comparison(para1, para2)
        #
        p = [comp.paragraphs[0].main_paragraph]
        comp.paragraphs[0].main_paragraph \
            = makdo.makdo_mddiff.File.set_configs(p, configs)[0]
        p = [comp.paragraphs[0].sub_paragraph]
        comp.paragraphs[0].sub_paragraph \
            = makdo.makdo_mddiff.File.set_configs(p, configs)[0]
        #
        self.quit_editing_formula()
        self.close_memo_pad()
        self.pnd.update()
        half_height = int(self.pnd.winfo_height() / 2) - 5
        self.pnd.remove(self.pnd1)
        self.pnd.remove(self.pnd2)
        self.pnd.remove(self.pnd3)
        self.pnd.remove(self.pnd4)
        self.pnd.remove(self.pnd5)
        self.pnd.remove(self.pnd6)
        self.pnd.forget(self.pnd3)
        self.pnd3 = tkinter.PanedWindow(self.pnd, bd=0, bg='#758F00')  # 070
        self.pnd.add(self.pnd1, height=half_height, minsize=100)
        self.pnd.add(self.pnd3, height=half_height)
        self.pnd.update()
        #
        background_color = self.background_color.get()
        if background_color == 'W':
            cvs = tkinter.Canvas(self.pnd3, bg='white')
            cvs_frm = tkinter.Frame(cvs, bg='white')
        elif background_color == 'B':
            cvs = tkinter.Canvas(self.pnd3, bg='black')
            cvs_frm = tkinter.Frame(cvs, bg='black')
        elif background_color == 'G':
            cvs = tkinter.Canvas(self.pnd3, bg='darkgreen')
            cvs_frm = tkinter.Frame(cvs, bg='darkgreen')
        cvs.pack(expand=True, fill=tkinter.BOTH, anchor=tkinter.W)
        scb = tkinter.Scrollbar(cvs, orient=tkinter.VERTICAL,
                                command=cvs.yview)
        scb.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        cvs['yscrollcommand'] = scb.set
        cvs.create_window((0, 0), window=cvs_frm, anchor='nw', )
        cvs_frm.bind(
            '<Configure>',
            lambda e: cvs.configure(scrollregion=cvs.bbox('all')))
        # cvs_frm.bind('<Up>', lambda e: cvs.yview_scroll(-1, 'units'))
        # cvs_frm.bind('<Down>', lambda e: cvs.yview_scroll(1, 'units'))
        self.btns = []
        size = self.font_size.get()
        for p in comp.paragraphs:
            if p.ses_symbol == '.':
                continue
            frm0 = tkinter.Frame(cvs_frm)
            frm1 = tkinter.Frame(frm0)
            frm2 = tkinter.Frame(frm0)
            btn1 = tkinter.Button(frm1, text='適用',
                                  command=self._apply_diff(frm0,
                                                           p.diff_id, comp))
            self.btns.append(btn1)
            btn2 = tkinter.Button(frm1, text='除外',
                                  command=self._exclude_diff(frm0))
            self.btns.append(btn2)
            btn3 = tkinter.Button(frm1, text='移動',
                                  command=self._goto_diff(p.diff_id, comp))
            self.btns.append(btn3)
            lbl = tkinter.Label(frm2, text=p.diff_text,
                                font=(GOTHIC_FONT, size), justify='left')
            if background_color == 'W':
                frm0.configure(bg='white')
                frm1.configure(bg='white')
                frm2.configure(bg='white')
                lbl.configure(bg='white', fg='black')
            elif background_color == 'B':
                frm0.configure(bg='black')
                frm1.configure(bg='black')
                frm2.configure(bg='black')
                lbl.configure(bg='black', fg='white')
            elif background_color == 'G':
                frm0.configure(bg='darkgreen')
                frm1.configure(bg='darkgreen')
                frm2.configure(bg='darkgreen')
                lbl.configure(bg='darkgreen', fg='lightyellow')
            frm0.pack(expand=True,
                      side=tkinter.TOP, anchor=tkinter.W, fill=tkinter.X)
            frm1.pack(expand=True,
                      side=tkinter.TOP, anchor=tkinter.W, fill=tkinter.X)
            btn1.pack(side=tkinter.LEFT)
            btn2.pack(side=tkinter.LEFT)
            btn3.pack(side=tkinter.LEFT)
            frm2.pack(expand=True,
                      side=tkinter.TOP, anchor=tkinter.W, fill=tkinter.X)
            lbl.pack(expand=True, side=tkinter.LEFT, anchor=tkinter.W)
        btn = tkinter.Button(self.pnd3, text='キャンセル', command=self._quit_diff)
        btn.pack()
        # cvs_frm.focus_set()

    def _apply_diff(self, frame, diff_id, comp):
        def x():
            txt = self.txt.get('1.0', 'end-1c')
            beg, end = self._get_diff_position(diff_id, comp, txt)
            if beg < 0 or end < 0:
                return False
            for cp in comp.paragraphs:
                if cp.diff_id != diff_id:
                    continue
                if cp.ses_symbol == '&':
                    self.txt.delete('1.0+' + str(beg) + 'c',
                                    '1.0+' + str(end) + 'c')
                    if cp.sub_paragraph != '':  # for empty configuration
                        insert_text = cp.sub_paragraph + '\n\n'
                        self.txt.insert('1.0+' + str(beg) + 'c', insert_text)
                        t = self.txt.get('1.0', '1.0+' + str(beg) + 'c')
                        beg_line = t.count('\n')
                        end_line = beg_line + insert_text.count('\n')
                        for i in range(beg_line - 1, end_line):
                            self.paint_out_line(i)
                elif cp.ses_symbol == '-':
                    self.txt.delete('1.0+' + str(beg) + 'c',
                                    '1.0+' + str(end) + 'c')
                elif cp.ses_symbol == '+':
                    if cp.sub_paragraph != '':  # for empty configuration
                        if beg >= len(txt) and \
                           not re.match('^(.|\n)*\n$', txt):
                            insert_text = '\n\n' + cp.sub_paragraph + '\n'
                        else:
                            insert_text = '\n' + cp.sub_paragraph + '\n'
                        self.txt.insert('1.0+' + str(beg) + 'c', insert_text)
                        t = self.txt.get('1.0', '1.0+' + str(beg) + 'c')
                        beg_line = t.count('\n')
                        end_line = beg_line + insert_text.count('\n')
                        for i in range(beg_line, end_line + 1):
                            self.paint_out_line(i)
                cp.has_applied = True
                frame.destroy()
                return True
        return x

    def _exclude_diff(self, frame):
        def x():
            frame.destroy()
            return True
        return x

    def _goto_diff(self, diff_id, comp):
        def x():
            txt = self.txt.get('1.0', 'end-1c')
            beg, end = self._get_diff_position(diff_id, comp, txt)
            if beg < 0 or end < 0:
                return False
            self.txt.mark_set('insert', '1.0+' + str(beg) + 'c')
            return True
        return x

    @staticmethod
    def _get_diff_position(diff_id, comp, txt):
        pars = makdo.makdo_mddiff.File.get_raw_paragraphs(txt)
        if pars[0] == '':
            pars.pop(0)  # for empty configuration
        p = ''
        n = 0
        s = ''
        for cp in comp.paragraphs:
            if cp.get_current_paragraph() != '':
                p = cp.get_current_paragraph()
                n += 1
            if cp.diff_id == diff_id:
                s = cp.ses_symbol
                break
        par = pars[n - 1]
        if p != re.sub('\n+$', '', par):
            n = 'エラー'
            m = '編集場所が見当たりません．'
            tkinter.messagebox.showerror(n, m)
            return -1, -1
        if s != '+':
            pre = ''.join(pars[:n - 1])
        else:
            pre = ''.join(pars[:n])
        beg = len(pre)
        if re.match('^(.|\n)*\n\n$', pre):
            if s == '+':
                beg -= 1
        if s == '+':
            end = beg
        else:
            end = beg + len(par)
        return beg, end

    def _quit_diff(self):
        self.pnd.remove(self.pnd3)
        self.txt.focus_set()

    # MDDIFF<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    # FOLD

    def fold_section(self):
        sub_document = self.txt.get('insert linestart', 'end-1c')
        # CHECK THAT THE LINE IS SECITION
        res = '^#+(?:-#+)*(?:\\s.*)?\n'
        if not re.match(res, sub_document):
            n = 'エラー'
            m = '行がセクションの見出し（"#"から始まる行）ではありません．'
            tkinter.messagebox.showerror(n, m)
            return
        # CHECK THAT HEADING IS NOT EMPTY
        res = '^#+(?:-#+)*\\s*\n\n'
        if re.match(res, sub_document):
            n = 'エラー'
            m = 'セクションの見出しがありません（字下げの調整です）．'
            tkinter.messagebox.showerror(n, m)
            return
        # CHECK THAT THE END OF LINE IS NOT ESCAPED
        fln = sub_document.split('\n')[0]
        if not re.match(NOT_ESCAPED + 'X$', fln + 'X'):
            n = 'エラー'
            m = 'セクションの見出しがエスケープされています' + \
                '（バックスラッシュで終わっています）．'
            tkinter.messagebox.showerror(n, m)
            return
        # CHECK THAT SECITION IS NOT FOLDED
        res = '^#+(?:-#+)*(?:\\s.*)?\\.\\.\\.\\[([0-9]+)\\]\n(?:.|\n)*$'
        if re.match(res, sub_document):
            n, m = 'エラー', 'セクションは折り畳まれています．'
            tkinter.messagebox.showerror(n, m)
            return
        # SHOW MESSAGE
        self.show_folding_help_message()
        # GET FOLDING NUMBER
        folding_number = 1
        all_document = self.txt.get('1.0', 'end-1c')
        res = '^\\.\\.\\.\\[([0-9]+)\\].*$'
        for line in all_document.split('\n'):
            if re.match(res, line):
                n = int(re.sub(res, '\\1', line))
                if folding_number <= n:
                    folding_number = n + 1
        # GET SECTION LINE
        sub_lines = sub_document.split('\n')
        section_line = sub_lines[0]
        # GET SECTION LEVEL
        res = '^(#+).*$'
        section_level = len(re.sub(res, '\\1', section_line))
        # GET TEXT TO FOLD
        text_to_fold = ''
        is_end_of_document = False
        m = len(sub_lines) - 1
        for i in range(1, m + 1):
            line = sub_lines[i]
            if re.match('^(#+)(?:-#+)*(?:\\s.*)?$', line):
                # SECTION
                level = len(re.sub(res, '\\1', line))
                if level <= section_level:
                    if not re.match('^#+(?:-#+)*\\s*$', line) or \
                       not (i < m and sub_lines[i + 1] == ''):
                        tmp = re.sub('<!--(.|\n)*?-->', '', text_to_fold)
                        if re.match('^(.|\n)*\n<!--(.|\n)*$', tmp):
                            # "\n<!--\n## xxx"
                            text_to_fold \
                                = re.sub('<!--(.|\n)*$', '', text_to_fold)
                            break
                        if not re.match('^(.|\n)*<!--(.|\n)*$', tmp):
                            # not "yyy<!--\n## xxx"
                            break
            if re.match('^\\.\\.\\.\\[[0-9]+\\]#+(?:-#+)*(?:\\s.*)?$', line):
                # FOLDED SECTION
                text_to_fold \
                    = re.sub(DONT_EDIT_MESSAGE + '\n\n$', '', text_to_fold)
                break
            text_to_fold += line + '\n'
        else:
            is_end_of_document = True
        # INSERT FOLDING MARK
        self.txt.insert('insert lineend', '...[' + str(folding_number) + ']')
        # INSERT TEXT TO FOLD
        if is_end_of_document:
            self.txt.insert('end', '\n')
        else:
            if not re.match('^(.|\n)*\n$', sub_document):
                self.txt.insert('end', '\n\n')
            elif not re.match('^(.|\n)*\n\n$', sub_document):
                self.txt.insert('end', '\n')
        self.txt.insert('end', DONT_EDIT_MESSAGE + '\n\n')
        self.txt.insert('end', '...[' + str(folding_number) + ']')
        self.txt.insert('end', section_line + '\n')
        self.txt.insert('end', text_to_fold)
        if re.match('^(.|\n)*\n\n\n', text_to_fold):
            self.txt.delete('end-1c', 'end')
        # DELETE FOLDING TEXT
        beg = 'insert lineend + 1c'
        end = 'insert lineend +' + str(len(text_to_fold)) + 'c'
        self.txt.delete(beg, end)
        # MOVE
        # self.txt.mark_set('insert', 'insert linestart')

    def unfold_section(self):
        sub_document = self.txt.get('insert linestart', 'end-1c')
        # CHECK THAT THE LINE IS SECITION
        res = '^#+(?:-#+)*(?:\\s.*)?\n'
        if not re.match(res, sub_document):
            n = 'エラー'
            m = '行がセクションの見出し（"#"から始まる行）ではありません．'
            tkinter.messagebox.showerror(n, m)
            return
        # CHECK THAT SECITION IS FOLDED
        res = '^#+(?:-#+)*(?:\\s.*)?\\.\\.\\.\\[([0-9]+)\\]\n(?:.|\n)*$'
        if not re.match(res, sub_document):
            n, m = 'エラー', 'セクションが折り畳まれていません．'
            tkinter.messagebox.showerror(n, m)
            return
        # CHECK THAT TEXT TO UNFOLD EXISTS
        folding_number = re.sub(res, '\\1', sub_document)
        res_mark = '\\.\\.\\.\\[' + folding_number + '\\]'
        res = '^' + '((?:.|\n)*?\n)' \
            + '((?:' + DONT_EDIT_MESSAGE + '\n+)?)' \
            + '(' + res_mark + '#+(?:-#+)*(?:\\s.*)?\n)' \
            + '((?:.|\n)*)$'
        if not re.match(res, sub_document):
            n, m = 'エラー', '折り畳み先が見付かりません．'
            tkinter.messagebox.showerror(n, m)
            return
        # DISPLAY MESSAGE
        self.show_folding_help_message()
        # GET TEXT
        text_a = re.sub(res, '\\1', sub_document)  # unconcerned
        text_b = re.sub(res, '\\2', sub_document)  # dont edit message
        text_c = re.sub(res, '\\3', sub_document)  # folding mark line
        text_d = re.sub(res, '\\4', sub_document)  # text to unfold
        res = '^' + '((?:.|\n)*?\n)' \
            + '((?:' + DONT_EDIT_MESSAGE + '\n+)?)' \
            + '(\\.\\.\\.\\[[0-9]+\\]#+(?:-#+)*(?:\\s.*)?\n)' \
            + '((?:.|\n)*)$'
        if re.match(res, text_d):
            text_d = re.sub(res, '\\1', text_d)
        # ADJUST LINE BREAK
        number_of_line_break_to_insert = 0
        if self.txt.get('insert lineend +1c', 'insert lineend +2c') == '\n':
            number_of_line_break_to_insert -= 1
        if not re.match('^(.|\n)*\n$', text_d):
            number_of_line_break_to_insert += 2
        elif not re.match('^(.|\n)*\n\n$', text_d):
            number_of_line_break_to_insert += 1
        # INSERT TEXT TO UNFOLD
        self.txt.insert('insert lineend +1c', text_d)
        # PAINT
        beg = self._get_v_position_of_insert()
        end = beg + text_d.count('\n')
        # REMOVE TEXT TO UNFOLD
        text_e = text_a + text_b + text_c + text_d
        beg = 'insert linestart +' + str(len(text_d + text_a)) + 'c'
        end = 'insert linestart +' + str(len(text_d + text_e)) + 'c'
        self.txt.delete(beg, end)
        # ADJUST LINE BREAK
        if number_of_line_break_to_insert == -1:
            beg = 'insert lineend +' + str(len(text_d)) + 'c'
            end = 'insert lineend +' + str(len(text_d) + 1) + 'c'
            self.txt.delete(beg, end)
        elif number_of_line_break_to_insert > 0:
            ins = 'insert lineend +' + str(len(text_d) + 1) + 'c'
            self.txt.insert(ins, '\n' * number_of_line_break_to_insert)
        # REMOVE FOLDING MARK
        text_f = '...[' + folding_number + ']'
        beg = 'insert lineend -' + str(len(text_f)) + 'c'
        end = 'insert lineend'
        self.txt.delete(beg, end)
        # MOVE
        # self.txt.mark_set('insert', 'insert linestart')

    def unfold_section_fully(self):
        old_document = self.txt.get('1.0', 'end-1c')
        if old_document == '':
            return
        new_document = self.get_fully_unfolded_document(old_document)
        self.file_lines = new_document.split('\n')
        self.txt.insert('1.0', new_document)
        self.txt.delete('1.0+' + str(len(new_document)) + 'c', 'end')
        self.txt.focus_set()
        self.txt.mark_set('insert', '1.0')
        # PAINT
        paint_keywords = self.paint_keywords.get()
        self.line_data = [LineDatum() for line in self.file_lines]
        for i, line in enumerate(self.file_lines):
            self.line_data[i].line_number = i
            self.line_data[i].line_text = line + '\n'
            if i > 0:
                self.line_data[i].beg_chars_state \
                    = self.line_data[i - 1].end_chars_state.copy()
                self.line_data[i].beg_chars_state.reset_partially()
            self.line_data[i].paint_line(self.txt, paint_keywords)

    def get_fully_unfolded_document(self, old_document):
        # |                ->  |
        # |## www...[3]    ->  |## www
        # |                ->  |
        # |...[1]#### yyy  ->  |### xxx
        # |                ->  |
        # |zzz             ->  |#### yyy
        # |                ->  |
        # |...[2]### xxx   ->  |zzz
        # |                ->  |
        # |#### yyy...[1]  ->  |
        # |                ->  |
        # |...[3]## www    ->  |
        # |                ->  |
        # |### xxx...[2]   ->  |
        # |                ->  |
        if old_document == '':
            return ''
        old_lines = old_document.split('\n')
        new_lines = []
        remain_lines = [True for i in old_lines]
        m = len(old_lines) - 1
        line_numbers = [0]
        res_mark = '\\.\\.\\.\\[([0-9])+\\]'
        res_from = '^(#+(?:-#+)*(?:\\s.*)?)' + res_mark + '$'
        res_to = '^' + res_mark + '#+(-#+)*(\\s|$)'
        while line_numbers != []:
            i = line_numbers[-1]
            if i > m:
                line_numbers.pop(-1)
                continue
            if not remain_lines[i]:
                line_numbers.pop(-1)
                continue
            if re.match(res_to, old_lines[i]):
                line_numbers.pop(-1)
                if new_lines[-2] == DONT_EDIT_MESSAGE and \
                   new_lines[-1] == '':
                    new_lines.pop(-1)
                    new_lines.pop(-1)
                continue
            if re.match(res_from, old_lines[i]) and \
               re.match(NOT_ESCAPED + res_mark + '$', old_lines[i]):
                folding_number \
                    = re.sub(res_from, '\\2', old_lines[i])
                old_lines[i] \
                    = re.sub(res_from, '\\1', old_lines[i])
                # APPEND "FROM LINE"
                new_lines.append(old_lines[i])
                remain_lines[i] = False
                line_numbers[-1] += 1
                if i < m and old_lines[i + 1] == '':
                    # SKIP "NEXT EMPTY LINE"
                    # new_lines.append(old_lines[i])
                    remain_lines[i + 1] = False
                    line_numbers[-1] += 1
                for j, line in enumerate(old_lines):
                    if not remain_lines[j]:
                        continue
                    res = '^\\.\\.\\.\\[' + folding_number + '\\]'
                    if re.match(res, line):
                        if j >= 2:
                            if old_lines[j - 2] == DONT_EDIT_MESSAGE and \
                               old_lines[j - 1] == '':
                                # SKIP "DONT EDIT MESSAGE"
                                remain_lines[j - 2] = False
                                remain_lines[j - 1] = False
                        line_numbers.append(j)
                        # JUMP TO "TO LINE"
                        # new_lines.append(old_lines[j])
                        remain_lines[j] = False
                        line_numbers[-1] += 1
            else:
                # APPEND "USUAL LINE"
                new_lines.append(old_lines[i])
                remain_lines[i] = False
                line_numbers[-1] += 1
        must_warn = True
        for i, ml in enumerate(old_lines):
            if remain_lines[i]:
                if must_warn:
                    n, m = 'エラー', '折り畳まれたセクションが残っています．'
                    tkinter.messagebox.showerror(n, m)
                new_lines.append(old_lines[i])
        new_document = '\n'.join(new_lines) + '\n\n'
        new_document = re.sub('\n\n+', '\n\n', new_document)
        new_document = re.sub('\n+$', '\n', new_document)
        return new_document

    def fold_or_unfold_section(self):
        sub_document = self.txt.get('insert linestart', 'end-1c')
        # CHECK THAT THE LINE IS SECITION
        res = '^#+(?:-#+)*(?:\\s.*)?\n'
        if not re.match(res, sub_document):
            n = 'エラー'
            m = '行がセクションの見出し（"#"から始まる行）ではありません．'
            tkinter.messagebox.showerror(n, m)
            return
        # CHECK THAT SECITION IS FOLDED
        res = '^#+(?:-#+)*(?:\\s.*)?\\.\\.\\.\\[([0-9]+)\\]\n(?:.|\n)*$'
        if not re.match(res, sub_document):
            self.fold_section()
        else:
            self.unfold_section()

    # LOOK IN DICTIONARY
    def look_in_dictionary(self, pane):
        if sys.platform != 'linux':  # epwing
            return
        w = ''
        if self.txt.tag_ranges('sel'):
            w = self.txt.get('sel.first', 'sel.last')
        if 'akauni' in self.txt.mark_names():
            w = ''
            w += self.txt.get('akauni', 'insert')
            w += self.txt.get('insert', 'akauni')
        #
        t = '辞書で調べる'
        p = '調べる言葉を入力してください．'
        s = OneWordDialog(pane, self, t, p, w).get_value()
        if s is None:
            return
        eb = makdo.eblook.Eblook()
        if self.dict_directory is None:
            return
        eb.set_dictionary_directory(self.dict_directory)
        eb.set_search_word(s)
        dic = ''
        for ei in eb.items:
            dic += '====================================='
            dic += '=====================================\n'
            dic += '●\u3000' + ei.dictionary.k_name \
                + '\u3000' + ei.title + '\n'
            dic += ei.content + '\n\n'
        self.quit_editing_formula()
        self.close_memo_pad()
        self.pnd.update()
        half_height = int(self.pnd.winfo_height() / 2) - 5
        self.pnd.remove(self.pnd1)
        self.pnd.remove(self.pnd2)
        self.pnd.remove(self.pnd3)
        self.pnd.remove(self.pnd4)
        self.pnd.remove(self.pnd5)
        self.pnd.remove(self.pnd6)
        self.pnd.add(self.pnd1, height=half_height, minsize=100)
        self.pnd.add(self.pnd2, height=half_height)
        self.pnd.update()
        #
        self.sub.pack(expand=True, fill=tkinter.BOTH)
        for key in self.txt.configure():
            self.sub.configure({key: self.txt.cget(key)})
        self.sub_btn.pack()
        #
        self.sub.delete('1.0', 'end')
        self.sub.insert('1.0', dic)
        self.sub.mark_set('insert', '1.0')
        n = 0
        pos = dic
        res = '^((?:.|\n)*?)(<gaiji=[^<>]+>)((?:.|\n)*)$'
        while re.match(res, pos):
            pre = re.sub(res, '\\1', pos)
            key = re.sub(res, '\\2', pos)
            pos = re.sub(res, '\\3', pos)
            beg = '1.0+' + str(n + len(pre)) + 'c'
            end = '1.0+' + str(n + len(pre) + len(key)) + 'c'
            n += len(pre) + len(key)
            self.sub.tag_add('error_tag', beg, end)
        #
        self.txt.focus_set()

    # OPENAI>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    openai_pepper = \
        'oBj5zRwt68RD2sgeLgtQKT39CoVPi7gesQPOJn0yGatizlokBCheNry4KsUrlO0J3' + \
        'Z5LXChEYjAaIzsNpFw27YDY65ADQPFgiCtOSpyKkvxhngvqPDXs5NUhcb8ZhsBBtX' + \
        'qiyf24IwPwRe6HJucepPA1CRpTIfkAEZsxReTvL4GSDzLiwd22HE12VhAcZ4p28Pv' + \
        'WrgiIv7JucXdOa5l1wBJj1zGRPDXmERLGevOsZ3nLlFinxhbb30TiHtVPcgIlDXlw' + \
        'i6jVojxMXUVKiRuOAaTx60c68K3NAYOwQxUwJawwz8JdCAQyq8HbdiPU6yat24Ado'

    def _encode(self, decoded):
        encoded = ''
        for n in range(len(decoded)):
            if re.match('^[0-9A-Za-z]$', decoded[n]):
                x = self._c2i(decoded[n])
                y = self._c2i(self.openai_pepper[n])
                z = (x + y) % 62
                encoded += self._i2c(z)
            else:
                encoded += decoded[n]
        return encoded

    def _decode(self, encoded):
        decoded = ''
        for n in range(len(encoded)):
            if re.match('^[0-9A-Za-z]$', encoded[n]):
                z = self._c2i(encoded[n])
                y = self._c2i(self.openai_pepper[n])
                x = (z - y + 62) % 62
                decoded += self._i2c(x)
            else:
                decoded += encoded[n]
        return decoded

    @staticmethod
    def _c2i(c):
        if re.match('^[0-9]$', c):
            return ord(c) - 48 + 0
        elif re.match('^[A-Z]$', c):
            return ord(c) - 65 + 10
        else:
            return ord(c) - 97 + 36

    @staticmethod
    def _i2c(i):
        if i < 10:
            return chr(i + 48 - 0)
        elif i < 36:
            return chr(i + 65 - 10)
        else:
            return chr(i + 97 - 36)

    openai_file = CONFIG_DIR + '/openai.md'

    def ask_openai(self, pane=None):
        if self.openai_key is None:
            self.input_openai_key()
        if self.openai_key is None:
            return
        k = self._decode(self.openai_key)
        if self.txt.tag_ranges('sel'):
            beg, end = self.txt.index('sel.first'), self.txt.index('sel.last')
            q = self.txt.get(beg, end)
            self.txt.tag_remove('sel', "1.0", "end")
        elif 'akauni' in self.txt.mark_names():
            beg, end = self._get_indices_in_order(self.txt, 'insert', 'akauni')
            q = self.txt.get(beg, end)
            self.txt.tag_remove('akauni_tag', '1.0', 'end')
            self.txt.mark_unset('akauni')
        else:
            t = 'OpenAIに質問'
            p = 'OpenAIへの質問を入力してください．'
            if pane is None:
                q = OneWordDialog(self.txt, self, t, p).get_value()
            else:
                q = OneWordDialog(pane, self, t, p).get_value()
        if q is None:
            return
        q = re.sub('\n$', '', q)
        pos = self.txt.index('insert')
        self.set_message_on_status_bar('質問しています', True)
        res = openai.OpenAI(api_key=k).chat.completions.create(
            model=self.openai_model,
            n=1, max_tokens=1000,
            messages=[{'role': 'user', 'content': q}],
        )
        self.set_message_on_status_bar('')
        #
        self.quit_editing_formula()
        self.close_memo_pad()
        self.pnd.update()
        half_height = int(self.pnd.winfo_height() / 2) - 5
        self.pnd.remove(self.pnd1)
        self.pnd.remove(self.pnd2)
        self.pnd.remove(self.pnd3)
        self.pnd.remove(self.pnd4)
        self.pnd.remove(self.pnd5)
        self.pnd.remove(self.pnd6)
        self.pnd.add(self.pnd1, height=half_height, minsize=100)
        self.pnd.add(self.pnd2, height=half_height)
        self.pnd.update()
        #
        self.sub.pack(expand=True, fill=tkinter.BOTH)
        for key in self.txt.configure():
            self.sub.configure({key: self.txt.cget(key)})
        self.sub_btn.pack()
        #
        md = ''
        for c in res.choices:
            n = adjust_line(c.message.content)
            md = '# ' + q + '\n\n' + n + '\n\n' + '-' * MD_TEXT_WIDTH + '\n\n'
        if os.path.exists(self.openai_file):
            try:
                with open(self.openai_file, 'r') as of:
                    md += of.read()
                os.rename(self.openai_file, self.openai_file + '~')
            except BaseException:
                pass
        try:
            with open(self.openai_file, 'w') as of:
                of.write(md)
        except BaseException:
            pass
        #
        self.sub.delete('1.0', 'end')
        self.sub.insert('1.0', md)
        self.sub.mark_set('insert', '1.0')
        # self.sub.configure(state='disabled')
        #
        self.txt.focus_force()

    def reset_openai(self):
        t = 'OpenAIの履歴の削除'
        m = 'OpenAIの履歴を削除しますか'
        if tkinter.messagebox.askyesno(t, m, default='no'):
            if os.path.exists(self.openai_file):
                os.remove(self.openai_file)
            if os.path.exists(self.openai_file + '~'):
                os.remove(self.openai_file + '~')
            return True
        return False

    def input_openai_model(self):
        t = 'OpenAIのモデル'
        m = 'OpenAIのモデルを入力してください．'
        om = OneWordDialog(self.txt, self, t, m, self.openai_model)
        if om is None:
            return
        self.openai_model = om
        self.show_config_help_message()

    def input_openai_key(self):
        t = 'OpenAIのキー'
        m = 'OpenAIのキーを入力してください．'
        ok = tkinter.simpledialog.askstring(t, m, show='*')
        if ok is None:
            return
        self.openai_key = self._encode(ok)
        self.show_config_help_message()

    # OPENAI<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    def execute_keyboard_macro(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        self.show_keyboard_macro_help_message()
        reversed_history = list(reversed(self.key_history))
        if reversed_history[1] != 'Ctrl+e':
            if reversed_history[0] == 'Ctrl+e':
                reversed_history.pop(0)
            for i in range(10, -1, -1):
                kh1 = []
                for j in range(i):
                    kh1.append(reversed_history[j])
                kh2 = []
                for j in range(i, i * 2):
                    kh2.append(reversed_history[j])
                if kh1 == kh2:
                    break
            if kh1 == kh2:
                self.keyboard_macro = list(reversed(kh1))
                self.keyborad_macro_h_position \
                    = self._get_ideal_h_position_of_insert(pane)
            else:
                self.keyboard_macro = []
        for i, key in enumerate(self.keyboard_macro):
            if key == 'BackSpace':
                pane.delete('insert-1c', 'insert')
            elif key == 'Delete':
                if i > 0 and self.keyboard_macro[i - 1] != 'Delete':
                    self.win.clipboard_clear()
                self._execute_when_delete_is_pressed(pane)
            elif key == 'Return':
                pane.insert('insert', '\n')
            elif key == 'Ctrl+p' or key == 'F15':
                self.paste_region()
            elif key == 'Home':
                pane.mark_set('insert', 'insetr linestart')
            elif key == 'End':
                pane.mark_set('insert', 'insetr lineend')
            elif key == 'Up':
                i = self._get_v_position_of_insert() - 1
                line = pane.get(str(i) + '.0', str(i) + '.end')
                line_pre, line_pos = '', ''
                for c in line:
                    ih = get_ideal_width(line_pre + c)
                    if ih > self.keyborad_macro_h_position:
                        break
                    line_pre += c
                j = len(line_pre)
                pane.mark_set('insert', str(i) + '.' + str(j))
            elif key == 'Down':
                i = self._get_v_position_of_insert() + 1
                line = pane.get(str(i) + '.0', str(i) + '.end')
                line_pre, line_pos = '', ''
                for c in line:
                    ih = get_ideal_width(line_pre + c)
                    if ih > self.keyborad_macro_h_position:
                        break
                    line_pre += c
                j = len(line_pre)
                pane.mark_set('insert', str(i) + '.' + str(j))
            elif key == 'Left':
                pane.mark_set('insert', 'insert-1c')
            elif key == 'Right':
                pane.mark_set('insert', 'insert+1c')
            elif key == 'F22':            # f (mark, save)
                if 'akauni' in pane.mark_names():
                    pane.mark_unset('akauni')
                pane.mark_set('akauni', 'insert')
            else:
                pane.insert('insert', key)
            if key != 'Up' and key != 'Down':
                self.keyborad_macro_h_position \
                    = self._get_ideal_h_position_of_insert(pane)

    # MINIBUFFER

    def start_minibuffer(self):
        self.MiniBuffer(self.txt, self)

    class MiniBuffer(tkinter.simpledialog.Dialog):

        commands = ['help',
                    'ask-openai',
                    'change_typeface',
                    'comment-out-region',
                    'compare-with-previous-draft',
                    'edit-formula1',
                    'edit-formula2',
                    'edit-formula3',
                    'edit-formula4',
                    'edit-formula5',
                    'fold-or-unfold-section',
                    'goto-flag1',
                    'goto-flag2',
                    'goto-flag3',
                    'goto-flag4',
                    'goto-flag5',
                    'insert-current-date',
                    'insert-current-time',
                    'insert-file',
                    'insert-file-names-in-same-folder',
                    'insert-formula1',
                    'insert-formula2',
                    'insert-formula3',
                    'insert-formula4',
                    'insert-formula5',
                    'insert-symbol',
                    'open-memo-pad',
                    'place-flag1',
                    'place-flag2',
                    'place-flag3',
                    'place-flag4',
                    'place-flag5',
                    'replace-all',
                    'save-file',
                    'search-or-replace-backward',
                    'search-or-replace-forward',
                    'split-or-unify-window',
                    'toggle-read-only',
                    'uncomment-in-region',
                    'quit-makdo',
                    'show-character-information']

        if sys.platform == 'linux':  # epwing
            commands.append('look-in-dictionary')

        help_message = \
            'help\n' + \
            '　このメッセージを表示\n' + \
            'ask-openai\n' + \
            '　OpenAIに質問\n' + \
            'change_typeface\n' + \
            '　字体を変える\n' + \
            'comment-out-region\n' + \
            '　指定範囲をコメントアウト\n' + \
            'compare-with-previous-draft\n' + \
            '　編集前の原稿と比較\n' + \
            'edit-formulaX(X=1..5)\n' + \
            '　定型句Xを編集\n' + \
            'insert-formulaX(X=1..5)\n' + \
            '　定型句Xを挿入\n' + \
            'uncomment-in-region\n' + \
            '　指定範囲のコメントアウトを解除\n' + \
            'fold-or-unfold-section\n' + \
            '　セクションの折畳又は展開\n' + \
            'place-flagX(X=1..5)\n' + \
            '　フラグXを設置\n' + \
            'goto-flagX(X=1..5)\n' + \
            '　フラグXに移動\n' + \
            'insert-current-date\n' + \
            '　今日の日付を挿入\n' + \
            'insert-current-time\n' + \
            '　現在の日時を挿入\n' + \
            'insert-file\n' + \
            '　テキストファイルの内容を挿入\n' + \
            'insert-file-names-in-same-folder\n' + \
            '　ファイル名のみを一括挿入\n' + \
            'insert-symbol\n' + \
            '　記号を挿入\n' + \
            'open-memo-pad\n' + \
            '　メモ帳を開く\n' + \
            'replace-all\n' + \
            '　文章全体又は指定範囲を全置換\n' + \
            'save-file\n' + \
            '　ファイルを保存\n' + \
            'search-or-replace-backward\n' + \
            '　前を検索又は置換\n' + \
            'search-or-replace-forward\n' + \
            '　次を検索又は置換\n' + \
            'split-or-unify-window\n' + \
            '　画面を分割又は統合\n' + \
            'toggle-read-only\n' + \
            '　読取専用を指定又は解除\n' + \
            'quit-makdo\n' + \
            '　Makdoを終了\n' + \
            'show-character-information\n' + \
            '　文字情報を表示'

        if sys.platform == 'linux':  # epwing
            help_message += \
                '\n' + \
                'look-in-dictionary\n' + \
                '　epwing形式の辞書で意味を調べる'

        history = []

        def __init__(self, pane, mother, init=''):
            self.pane = pane
            self.mother = mother
            self.init = init
            self.history_number = 0
            if len(self.history) == 0:
                Makdo.MiniBuffer.history.append('')
            elif self.history[-1] in self.commands:
                Makdo.MiniBuffer.history.append('')
            else:
                Makdo.MiniBuffer.history[-1] = ''
            super().__init__(pane, title='ミニバッファ')

        def body(self, pane):
            t = 'コマンドを入力してください．\n' \
                + '分からなければ"help"と入力してください．'
            lbl = tkinter.Label(pane, text=t, justify='left')
            lbl.pack(side='top', anchor='w')
            size = self.mother.font_size.get()
            self.etr = tkinter.Entry(pane, font=(GOTHIC_FONT, size), width=50)
            self.etr.pack(side='top')
            self.etr.insert(0, self.init)
            self.bind('<Key-Tab>', self.key_tab)
            self.bind('<Key-Up>', self.key_up)
            self.bind('<Key-Down>', self.key_down)
            super().body(pane)
            return self.etr

        def apply(self):
            com = self.etr.get()
            Makdo.MiniBuffer.history[-1] = com
            if len(self.history) > 1:
                if Makdo.MiniBuffer.history[-2] == com:
                    Makdo.MiniBuffer.history.pop(-1)
            if com == '':
                return
            elif com == 'help':
                tkinter.messagebox.showinfo('ヘルプ', self.help_message)
                Makdo.MiniBuffer(self, self.mother)
            elif com == 'ask-openai':
                self.mother.ask_openai(self)
            elif com == 'change_typeface':
                self.mother.change_typeface()
            elif com == 'comment-out-region':
                self.mother.comment_out_region()
            elif com == 'compare-with-previous-draft':
                self.mother.compare_with_previous_draft()
            elif com == 'edit-formula1' or com == 'edit-formula':
                self.mother.edit_formula1()
            elif com == 'edit-formula2':
                self.mother.edit_formula2()
            elif com == 'edit-formula3':
                self.mother.edit_formula3()
            elif com == 'edit-formula4':
                self.mother.edit_formula4()
            elif com == 'edit-formula5':
                self.mother.edit_formula5()
            elif com == 'fold-or-unfold-section':
                self.mother.fold_or_unfold_section()
            elif com == 'goto-flag1' or com == 'goto-flag':
                self.mother.goto_flag1()
            elif com == 'goto-flag2':
                self.mother.goto_flag2()
            elif com == 'goto-flag3':
                self.mother.goto_flag3()
            elif com == 'goto-flag4':
                self.mother.goto_flag4()
            elif com == 'goto-flag5':
                self.mother.goto_flag5()
            elif com == 'insert-current-date':
                self.mother.insert_date_Gymd()
            elif com == 'insert-current-time':
                self.mother.insert_datetime_simple()
            elif com == 'insert-file':
                self.mother.insert_file()
            elif com == 'insert-file-names-in-same-folder':
                self.mother.insert_file_names_in_same_folder()
            elif com == 'insert-formula1' or com == 'insert-formula':
                self.mother.insert_formula1()
            elif com == 'insert-formula2':
                self.mother.insert_formula2()
            elif com == 'insert-formula3':
                self.mother.insert_formula3()
            elif com == 'insert-formula4':
                self.mother.insert_formula4()
            elif com == 'insert-formula5':
                self.mother.insert_formula5()
            elif com == 'insert-symbol':
                self.mother.insert_symbol()
            elif com == 'look-in-dictionary':
                self.mother.look_in_dictionary(self)
            elif com == 'open-memo-pad':
                self.mother.open_memo_pad()
            elif com == 'place-flag1' or com == 'place-flag':
                self.mother.place_flag1()
            elif com == 'place-flag2':
                self.mother.place_flag2()
            elif com == 'place-flag3':
                self.mother.place_flag3()
            elif com == 'place-flag4':
                self.mother.place_flag4()
            elif com == 'place-flag5':
                self.mother.place_flag5()
            elif com == 'replace-all':
                self.mother.replace_all()
            elif com == 'save-file':
                self.mother.save_file()
            elif com == 'search-or-replace-backward':
                self.mother.search_or_replace_backward_from_dialog(self)
            elif com == 'search-or-replace-forward':
                self.mother.search_or_replace_forward_from_dialog(self)
            elif com == 'split-or-unify-window':
                self.mother.split_or_unify_window()
            elif com == 'toggle-read-only':
                is_read_only = self.mother.is_read_only.get()
                if is_read_only:
                    self.mother.is_read_only.set(False)
                else:
                    self.mother.is_read_only.set(True)
                # self.mother.toggle_read_only()
            elif com == 'uncomment-in-region':
                self.mother.uncomment_in_region()
            elif com == 'quit-makdo':
                # 2 ERRORS OCCUR
                self.mother.quit_makdo()
            elif com == 'show-character-information':
                self.mother.show_char_info()
            else:
                Makdo.MiniBuffer(self, self.mother, com)

        def key_tab(self, event):
            com = self.etr.get()
            if com == '':
                return  # empty
            cnd = []
            for c in self.commands:
                if com == c:
                    return  # completed
                if re.match('^' + com, c):
                    cnd.append(c)
            x = ''
            for y in cnd:
                if x == '':
                    x = y
                else:
                    nx, ny = len(x), len(y)
                    nz = min(nx, ny)
                    for n in range(nz):
                        if x[:n+1] != y[:n+1]:
                            if n == 0:
                                x = ''
                            else:
                                x = x[:n]
                            break
            if x != '':
                self.etr.delete(0, 'end')
                self.etr.insert(0, x)
            return 'break'

        def key_up(self, event):
            if self.history_number == 0:
                Makdo.MiniBuffer.history[-1] = self.etr.get()
            if self.history_number < len(self.history) - 1:
                self.history_number += 1
                self.etr.delete(0, 'end')
                self.etr.insert(0, self.history[-self.history_number - 1])

        def key_down(self, event):
            if self.history_number > 0:
                self.history_number -= 1
                self.etr.delete(0, 'end')
                self.etr.insert(0, self.history[-self.history_number - 1])

    ##########################
    # MENU CONFIGURATION

    def _make_menu_configuration(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='設定(S)', menu=menu, underline=3)
        #
        self.is_read_only = tkinter.BooleanVar(value=False)
        if self.args_read_only:
            self.is_read_only.set(True)
        menu.add_checkbutton(label='読取専用',
                             variable=self.is_read_only)
        menu.add_separator()
        #
        self.dont_show_help = tkinter.BooleanVar(value=False)
        if self.args_dont_show_help:
            self.dont_show_help.set(True)
        elif self.file_dont_show_help:
            self.dont_show_help.set(True)
        menu.add_checkbutton(label='ヘルプを表示しない',
                             variable=self.dont_show_help,
                             command=self.show_config_help_message)
        menu.add_separator()
        #
        self._make_submenu_background_color(menu)
        self._make_submenu_character_size(menu)
        menu.add_separator()
        #
        self.paint_keywords = tkinter.BooleanVar(value=False)
        if self.args_paint_keywords:
            self.paint_keywords.set(True)
        elif self.file_paint_keywords:
            self.paint_keywords.set(True)
        menu.add_checkbutton(label='キーワードに色付け',
                             variable=self.paint_keywords,
                             command=self.show_config_help_message)
        menu.add_separator()
        #
        self.make_backup_file = tkinter.BooleanVar(value=False)
        if self.args_make_backup_file:
            self.make_backup_file.set(True)
        elif self.file_make_backup_file:
            self.make_backup_file.set(True)
        menu.add_checkbutton(label='バックアップファイルを残す',
                             variable=self.make_backup_file,
                             command=self.show_config_help_message)
        menu.add_separator()
        #
        self._make_submenu_digit_separator(menu)
        menu.add_separator()
        #
        menu.add_command(label='OpenAIのモデルを入力',
                         command=self.input_openai_model)
        menu.add_command(label='OpenAIのキーを入力',
                         command=self.input_openai_key)
        menu.add_separator()
        #
        menu.add_checkbutton(label='設定を保存',
                             command=self.save_configurations)
        # menu.add_separator()

    ################
    # SUBMENU BACKGROUND COLOR AND CHARACTER SIZE

    def _make_submenu_background_color(self, menu):
        submenu = tkinter.Menu(self.mnb, tearoff=False)
        menu.add_cascade(label='背景色', menu=submenu)
        self.background_color \
            = tkinter.StringVar(value='W')
        if self.args_background_color is not None:
            self.background_color.set(self.args_background_color)
        elif self.file_background_color is not None:
            self.background_color.set(self.file_background_color)
        colors = {'W': '白色', 'B': '黒色', 'G': '緑色'}
        for c in colors:
            submenu.add_radiobutton(label=colors[c],
                                    variable=self.background_color, value=c,
                                    command=self.set_background_color)

    def _make_submenu_character_size(self, menu):
        submenu = tkinter.Menu(self.mnb, tearoff=False)
        menu.add_cascade(label='文字サイズ', menu=submenu)
        sizes = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36,
                 42, 48, 54, 60, 66, 72, 78, 84, 90, 96, 102, 108]
        self.font_size = tkinter.IntVar(value=18)
        if self.args_font_size is not None:
            self.font_size.set(self.args_font_size)
        elif self.file_font_size is not None:
            self.font_size.set(self.file_font_size)
        for s in sizes:
            submenu.add_radiobutton(label=str(s) + 'px',
                                    variable=self.font_size, value=s,
                                    command=self.set_character_size)

    ######
    # COMMAND

    def set_background_color(self):
        self.show_config_help_message()
        self.set_font()

    def set_character_size(self):
        self.show_config_help_message()
        self.set_font()

    def set_font(self):
        background_color = self.background_color.get()
        size = self.font_size.get()
        # BASIC FONT
        self.txt['font'] = (GOTHIC_FONT, size)
        self.stb_sor1['font'] = (GOTHIC_FONT, size)
        self.stb_sor2['font'] = (GOTHIC_FONT, size)
        self.txt.tag_config('error_tag', foreground='#FF0000')
        self.sub.tag_config('error_tag', foreground='#FF0000')
        self.txt.tag_config('search_tag', background='#777777')
        # COLOR FONT
        if background_color == 'W':
            self.txt.config(bg='white', fg='black')
            self.txt.tag_config('line_eof_tag', background='#CCCCCC')
            self.txt.tag_config('line_tag', background='#EEEEEE')
            self.txt.tag_config('eof_tag', background='#EEEEEE')
            self.txt.tag_config('akauni_tag', background='#CCCCCC')
            self.sub.tag_config('akauni_tag', background='#CCCCCC')
            self.txt.tag_config('hsp_tag', foreground='#C8C8FF',
                                underline=True)                   # (0.8, 240)
            self.txt.tag_config('tab_tag', background='#D9E7FF')  # (0.9, 220)
            self.txt.tag_config('fsp_tag', foreground='#90D9FF',
                                underline=True)                   # (0.8, 200)
        elif background_color == 'B':
            self.txt.config(bg='black', fg='white')
            self.txt.tag_config('line_eof_tag', background='#666666')
            self.txt.tag_config('line_tag', background='#333333')
            self.txt.tag_config('eof_tag', background='#333333')
            self.txt.tag_config('akauni_tag', background='#666666')
            self.sub.tag_config('akauni_tag', background='#666666')
            self.txt.tag_config('hsp_tag', foreground='#7676FF',
                                underline=True)                   # (0.5, 240)
            self.txt.tag_config('tab_tag', background='#0053EF')  # (0.3, 220)
            self.txt.tag_config('fsp_tag', foreground='#009AED',
                                underline=True)                   # (0.5, 200)
        elif background_color == 'G':
            self.txt.config(bg='darkgreen', fg='lightyellow')
            self.txt.tag_config('line_eof_tag', background='#339733')
            self.txt.tag_config('line_tag', background='#117511')
            self.txt.tag_config('eof_tag', background='#117511')
            self.txt.tag_config('akauni_tag', background='#888888')
            self.sub.tag_config('akauni_tag', background='#888888')
            self.txt.tag_config('hsp_tag', foreground='#7676FF',
                                underline=True)                   # (0.5, 240)
            self.txt.tag_config('tab_tag', background='#0053EF')  # (0.3, 220)
            self.txt.tag_config('fsp_tag', foreground='#009AED',
                                underline=True)                   # (0.5, 200)
        for u in ['-x', '-u']:
            und = False if u == '-x' else True
            for f in ['-g', '-m']:
                if f == '-g':
                    fon = (GOTHIC_FONT, size)
                else:
                    fon = (MINCHO_FONT, size)
                # WHITE
                for i in range(3):
                    a = '-XXX'
                    y = '-' + str(i)
                    tag = 'c' + a + y + f + u
                    if background_color == 'W':
                        col = BLACK_SPACE[i]
                    else:
                        col = WHITE_SPACE[i]
                    self.txt.tag_config(tag, font=fon,
                                        foreground=col, underline=und)
                if f == '-g':
                    fon = (GOTHIC_FONT, size, 'bold')
                else:
                    fon = (MINCHO_FONT, size, 'bold')
                # COLOR
                for i in range(3):  # lightness
                    y = '-' + str(i)
                    for j, c in enumerate(COLOR_SPACE):  # angle
                        a = '-' + str(j * 10)
                        tag = 'c' + a + y + f + u  # example: c-120-1-g-x
                        if background_color == 'W':
                            col = c[i]
                        elif background_color == 'B':
                            col = c[i + 1]
                        elif background_color == 'G':
                            col = c[i + 1]
                        self.txt.tag_config(tag, font=fon,
                                            foreground=col, underline=und)

    ################
    # SUBMENU DIGIT SEPARATOR

    def _make_submenu_digit_separator(self, menu):
        submenu = tkinter.Menu(self.mnb, tearoff=False)
        menu.add_cascade(label='計算結果', menu=submenu)
        #
        self.digit_separator = tkinter.StringVar(value='4')
        submenu.add_radiobutton(label='桁区切りなし（12345678）',
                                variable=self.digit_separator, value='0',
                                command=self.show_config_help_message)
        submenu.add_radiobutton(label='3桁区切り（12,345,678）',
                                variable=self.digit_separator, value='3',
                                command=self.show_config_help_message)
        submenu.add_radiobutton(label='4桁区切り（1234万5678）',
                                variable=self.digit_separator, value='4',
                                command=self.show_config_help_message)
        # menu.add_separator()

    ################
    # CONFIGURATION FILE

    def get_and_set_configurations(self):
        if not os.path.exists(CONFIG_DIR):
            os.mkdir(CONFIG_DIR)
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                for line in f:
                    line = line.rstrip()
                    item = re.sub('^\\s*(\\S*)\\s*:\\s*(.*)\\s*$', '\\1', line)
                    valu = re.sub('^\\s*(\\S*)\\s*:\\s*(.*)\\s*$', '\\2', line)
                    if item == 'dont_show_help':
                        if valu == 'True':
                            Makdo.file_dont_show_help = True
                        elif valu == 'True':
                            Makdo.file_dont_show_help = False
                    elif item == 'background_color':
                        if valu == 'W' or valu == 'B' or valu == 'G':
                            Makdo.file_background_color = valu
                    elif item == 'font_size':
                        if re.match('^[0-9]+$', valu) and (int(valu) % 3) == 0:
                            Makdo.file_font_size = int(valu)
                    elif item == 'paint_keywords':
                        if valu == 'True':
                            Makdo.file_paint_keywords = True
                        elif valu == 'False':
                            Makdo.file_paint_keywords = False
                    elif item == 'digit_separator':
                        if valu == '3' or valu == '4':
                            Makdo.file_digit_separator = valu
                    elif item == 'make_backup_file':
                        if valu == 'True':
                            Makdo.file_make_backup_file = True
                        elif valu == 'False':
                            Makdo.file_make_backup_file = False
                    elif item == 'openai_model':
                        self.openai_model = valu
                    elif item == 'openai_key':
                        self.openai_key = valu
                    elif item == 'dict_directory':
                        self.dict_directory = valu

    def save_configurations(self):
        if os.path.exists(CONFIG_FILE + '~'):
            os.remove(CONFIG_FILE + '~')
        if os.path.exists(CONFIG_FILE):
            os.rename(CONFIG_FILE, CONFIG_FILE + '~')
        with open(CONFIG_FILE, 'w') as f:
            f.write('dont_show_help:   '
                    + str(self.dont_show_help.get()) + '\n')
            f.write('background_color: '
                    + self.background_color.get() + '\n')
            f.write('font_size:        '
                    + str(self.font_size.get()) + '\n')
            f.write('paint_keywords:   '
                    + str(self.paint_keywords.get()) + '\n')
            f.write('digit_separator:  '
                    + str(self.digit_separator.get()) + '\n')
            f.write('make_backup_file: '
                    + str(self.make_backup_file.get()) + '\n')
            if self.openai_model is not None:
                f.write('openai_model:     '
                        + self.openai_model + '\n')
            if self.openai_key is not None:
                f.write('openai_key:       '
                        + self.openai_key + '\n')
            if sys.platform == 'linux':  # epwing
                if self.dict_directory is not None:
                    f.write('dict_directory:   '
                            + self.dict_directory + '\n')
            self.set_message_on_status_bar('設定を保存しました')

    ##########################
    # MENU INTERNET

    def _make_menu_internet(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='ネット(N)', menu=menu, underline=4)
        #
        menu.add_command(label='最新のMakdoを確認',
                         command=self.browse_makdo_home)
        menu.add_separator()
        #
        menu.add_command(label='人名・地名漢字を探す',
                         command=self.browse_ivs)
        menu.add_separator()
        #
        menu.add_command(label='goo辞書で調べる',
                         command=self.browse_goo_dictionary)
        menu.add_command(label='weblio辞書で調べる',
                         command=self.browse_weblio_dictionary)
        menu.add_command(label='Wikipediaで調べる',
                         command=self.browse_wikipedia)
        menu.add_separator()
        #
        menu.add_command(label='法律を調べる',
                         command=self.browse_law)
        menu.add_command(label='・日本国憲法',
                         command=self.browse_law_constitution_law)
        menu.add_command(label='・民法',
                         command=self.browse_law_civil_law)
        menu.add_command(label='・商法',
                         command=self.browse_law_commercial_law)
        menu.add_command(label='・会社法',
                         command=self.browse_law_corporation_law)
        menu.add_command(label='・民事訴訟法',
                         command=self.browse_law_civil_procedure)
        menu.add_command(label='・刑法',
                         command=self.browse_law_crime_law)
        menu.add_command(label='・刑事訴訟法',
                         command=self.browse_law_crime_procedure)
        menu.add_command(label='裁判所規則を調べる',
                         command=self.browse_rule_of_court)
        menu.add_separator()
        #
        menu.add_command(label='ChatGPTに接続',
                         command=self.browse_chatgpt)
        menu.add_command(label='OpenAIに接続',
                         command=self.browse_openai)
        menu.add_separator()
        #
        menu.add_command(label='Google Driveに接続',
                         command=self.browse_google_drive)
        menu.add_command(label='Microsoft OneDriveに接続',
                         command=self.browse_onedrive)
        menu.add_separator()
        #
        menu.add_command(label='mints（民事裁判書類電子提出システム）に接続',
                         command=self.browse_mints)
        # menu.add_separator()

    ################
    # COMMAND

    def browse_makdo_home(self):
        webbrowser.open('https://github.com/hata48915b/makdo/')

    def browse_ivs(self):
        c = ''
        if self.txt.tag_ranges('sel'):
            c = self.txt.get('sel.first', 'sel.last')
        elif 'akauni' in self.txt.mark_names():
            c = ''
            c += self.txt.get('akauni', 'insert')
            c += self.txt.get('insert', 'akauni')
        if len(c) == 1:
            d = re.sub('^0x', '', hex(ord(c))).upper()
            u = 'https://moji.or.jp/mojikibansearch/result' \
                + '?UCS%E6%BC%A2%E5%AD%97=' + d
            webbrowser.open(u)
            i = self.IvsDialog(self.txt, self, c)
        else:
            d = None
            u = 'https://moji.or.jp/mojikibansearch/basic'
            webbrowser.open(u)
            i = self.IvsDialog(self.txt, self)
        if len(c) == 1 and i.has_inserted:
            if self.txt.tag_ranges('sel'):
                self.txt.delete('sel.first', 'sel.first+1c')
            elif 'akauni' in self.txt.mark_names():
                if self.txt.get('akauni', 'insert') != '':
                    self.txt.delete('akauni', 'akauni+1c')
                elif self.txt.get('insert', 'akauni') != '':
                    self.txt.delete('akauni-1c', 'akauni')

    def browse_goo_dictionary(self):
        if self.txt.tag_ranges('sel'):
            w = self.txt.get('sel.first', 'sel.last')
            u = 'https://dictionary.goo.ne.jp/srch/all/' + w + '/m6u/'
            webbrowser.open(u)
        elif 'akauni' in self.txt.mark_names():
            w = ''
            w += self.txt.get('akauni', 'insert')
            w += self.txt.get('insert', 'akauni')
            u = 'https://dictionary.goo.ne.jp/srch/all/' + w + '/m6u/'
            webbrowser.open(u)
        else:
            webbrowser.open('https://dictionary.goo.ne.jp/')

    def browse_weblio_dictionary(self):
        if self.txt.tag_ranges('sel'):
            w = self.txt.get('sel.first', 'sel.last')
            u = 'https://www.weblio.jp/content/' + w
            webbrowser.open(u)
        elif 'akauni' in self.txt.mark_names():
            w = ''
            w += self.txt.get('akauni', 'insert')
            w += self.txt.get('insert', 'akauni')
            u = 'https://www.weblio.jp/content/' + w
            webbrowser.open(u)
        else:
            webbrowser.open('https://www.weblio.jp/')

    def browse_wikipedia(self):
        if self.txt.tag_ranges('sel'):
            w = self.txt.get('sel.first', 'sel.last')
            webbrowser.open('https://ja.wikipedia.org/wiki/' + w)
        if 'akauni' in self.txt.mark_names():
            w = ''
            w += self.txt.get('akauni', 'insert')
            w += self.txt.get('insert', 'akauni')
            webbrowser.open('https://ja.wikipedia.org/wiki/' + w)

    def browse_law(self):
        webbrowser.open('https://laws.e-gov.go.jp/')

    def browse_law_constitution_law(self):
        webbrowser.open('https://laws.e-gov.go.jp/law/321CONSTITUTION')

    def browse_law_civil_law(self):
        webbrowser.open('https://laws.e-gov.go.jp/law/129AC0000000089')

    def browse_law_commercial_law(self):
        webbrowser.open('https://laws.e-gov.go.jp/law/132AC0000000048')

    def browse_law_corporation_law(self):
        webbrowser.open('https://laws.e-gov.go.jp/law/417AC0000000086')

    def browse_law_civil_procedure(self):
        webbrowser.open('https://laws.e-gov.go.jp/law/408AC0000000109')

    def browse_law_crime_law(self):
        webbrowser.open('https://laws.e-gov.go.jp/law/140AC0000000045')

    def browse_law_crime_procedure(self):
        webbrowser.open('https://laws.e-gov.go.jp/law/323AC0000000131')

    def browse_rule_of_court(self):
        u = 'https://www.courts.go.jp/toukei_siryou/kisokusyu/index.html'
        webbrowser.open(u)

    def browse_chatgpt(self):
        webbrowser.open('https://chatgpt.com/')

    def browse_openai(self):
        webbrowser.open('https://openai.com/')

    def browse_google_drive(self):
        webbrowser.open('https://drive.google.com/drive/my-drive')

    def browse_onedrive(self):
        webbrowser.open('https://onedrive.live.com/')

    def browse_mints(self):
        webbrowser.open('https://www.mints.courts.go.jp/user/')

    ##########################
    # MENU HELP

    def _make_menu_help(self):
        menu = tkinter.Menu(self.mnb, tearoff=False)
        self.mnb.add_cascade(label='ヘルプ(H)', menu=menu, underline=4)
        #
        menu.add_command(label='文字情報',
                         command=self.show_char_info)
        menu.add_separator()
        #
        menu.add_command(label='ヘルプ(H)', underline=4,
                         command=self.show_help)
        menu.add_separator()
        #
        menu.add_command(label='ライセンス情報(F)', underline=8,
                         command=self.show_license_info)
        menu.add_separator()
        #
        menu.add_command(label='Makdoについて(A)', underline=10,
                         command=self.show_about_makdo)
        # menu.add_separator()

    ################
    # COMMAND

    def show_char_info(self):
        n = '文字情報'
        c = self.txt.get('insert', 'insert+1c')
        if c != '' and c != '\n':
            m = ''
            if c == ' ':
                m += '文字：\t（半角スペース）\n'
            elif c == '\t':
                m += '文字：\t（水平タブ）\n'
            elif c == '\u3000':
                m += '文字：\t（全角スペース）\n'
            else:
                m += '文字：\t' + c + '\n'
            m += 'UTF-8：\t' + re.sub('^0x', '', hex(ord(c))).upper() + '\n\n'
            for jk in JOYOKANJI:
                if c in jk[1]:
                    m += '常用漢字です．\n'
                    m += '字体：\u3000' + jk[1] + '\n'
                    m += '読み：\u3000' + jk[2] + '\n'
                    if jk[3] != '':
                        m += '用例：' + jk[3] + '\n'
                    break
            else:
                m += '常用漢字ではありません．\n'
                if re.match('^[ -~]$', c):
                    m += '半角英数記号\n'
                if re.match('^[ｦ-ﾟ]$', c):
                    m += '半角カタカナ\n'
                if re.match('^[ぁ-ゖ]$', c):
                    m += 'ひらがな\n'
                if re.match('^[ァ-ヺ]$', c) or re.match('^[ㇰ-ㇿ]$', c):
                    m += 'カタカナ\n'
                if re.match('^[０-９]$', c):
                    m += '数字\n'
            m = re.sub('\n+$', '', m)
            tkinter.messagebox.showinfo(n, m)

    def show_help(self):
        n = 'ヘルプ'
        m = 'このダイアログを閉じた後、' + \
            'ウィンドウにMS_Wordのファイル（拡張子docx）を' + \
            'ドラッグアンドドロップしてみてください．'
        tkinter.messagebox.showinfo(n, m)

    def show_license_info(self):
        n = 'ライセンス情報'
        m = 'Copyright (C) 2022-2024  Seiichiro HATA\n\n' + \
            'このソフトウェアは、\n' + \
            '"GNU GENERAL PUBLIC LICENSE\n' + \
            'Version 3 (GPLv3)"という\n' + \
            'ライセンスで開発されています．\n\n' + \
            'このソフトウェアは、\n' + \
            '次のモジュールを利用しており、\n' + \
            'それぞれ付記したライセンスで\n' + \
            '配布されています．\n' + \
            '- argparse: PSF License\n' + \
            '- chardet: LGPLv2+\n' + \
            '- python-docx: MIT License\n' + \
            '- lxml: BSD License (3-Clause)\n' + \
            '- typing_extensions: PSF License\n' + \
            '- tkinterdnd2: MIT License\n' + \
            '- openpyxl: MIT License\n' + \
            '- et-xmlfile: MIT License\n' + \
            '- openai: Apache Software License\n' + \
            '- annotated-types: MIT License\n' + \
            '- anyio: MIT License\n' + \
            '- certifi: Mozilla Public License 2.0\n' + \
            '- distro: Apache Software License\n' + \
            '- exceptiongroup: MIT License\n' + \
            '- h11: MIT License\n' + \
            '- httpcore: BSD License\n' + \
            '- httpx: BSD License\n' + \
            '- idna: BSD License\n' + \
            '- jiter: MIT License\n' + \
            '- pydantic: MIT License\n' + \
            '- pydantic_core: MIT License\n' + \
            '- sniffio: Apache Software License;\n' + \
            '　　MIT License\n' + \
            '- tqdm: MIT License;\n' + \
            '　　Mozilla Public License 2.0\n' + \
            '- pywin32: PSF License\n' + \
            '- Levenshtein: GPLv2+\n' + \
            '\n利用、改変、再配布等をする場合には、\n' + \
            'ライセンスに十分ご注意ください．\n' + \
            'スクリプトファイルは、\n' + \
            '外部のモジュールを読み込みますが、\n' + \
            'バイナリファイルは、\n' + \
            '内部にモジュールを含んでいますので、\n' + \
            '特にご注意ください．'

        tkinter.messagebox.showinfo(n, m)

    def show_about_makdo(self):
        n = 'バージョン情報'
        m = 'makdo ' + __version__ + '\n\n' + \
            '秦誠一郎により開発されています．'
        tkinter.messagebox.showinfo(n, m)

    ####################################
    # KEY CONFIGURATION

    def _make_txt_key_configuration(self):
        self.txt.bind('<Key>', self.txt_process_key)
        self.txt.bind('<KeyRelease>', self.txt_process_key_release)
        self.txt.bind('<Button-1>', self.txt_process_button1)
        self.txt.bind('<Button-2>', self.txt_process_button2)
        self.txt.bind('<Button-3>', self.txt_process_button3)
        self.txt.bind('<ButtonRelease-1>', self.txt_process_button1_release)
        self.txt.bind('<ButtonRelease-2>', self.txt_process_button2_release)
        self.txt.bind('<ButtonRelease-3>', self.txt_process_button3_release)

    def _make_sub_key_configuration(self):
        self.sub.bind('<Key>', self.sub_process_key)
        self.sub.bind('<KeyRelease>', self.sub_process_key_release)
        self.sub.bind('<Button-1>', self.sub_process_button1)
        self.sub.bind('<Button-2>', self.sub_process_button2)
        self.sub.bind('<Button-3>', self.sub_process_button3)
        self.sub.bind('<ButtonRelease-1>', self.sub_process_button1_release)
        self.sub.bind('<ButtonRelease-2>', self.sub_process_button2_release)
        self.sub.bind('<ButtonRelease-3>', self.sub_process_button3_release)

    ##########################
    # COMMAND

    def txt_process_key(self, key):
        self.current_pane = 'txt'
        is_read_only = self.is_read_only.get()
        if is_read_only:
            return self.read_only_process_key(self.txt, key)
        else:
            return self.read_and_write_process_key(self.txt, key)

    def sub_process_key(self, key):
        self.current_pane = 'sub'
        if key.keysym == 'Escape':
            self._unify_window()
            return 'break'
        if self.formula_number < 0 and self.memo_pad_memory is None:
            return self.read_only_process_key(self.sub, key)
        else:
            return self.read_and_write_process_key(self.sub, key)

    def read_and_write_process_key(self, pane, key):
        self.set_message_on_status_bar('')
        self.set_position_info_on_status_bar()
        self.paint_out_line(self._get_v_position_of_insert() - 1)
        # HISTORY
        if key.keysym == 'Shift_L' or key.keysym == 'Shift_R':
            return
        if key.keysym == 'Control_L' or key.keysym == 'Control_R':
            return
        if key.keysym == 'Alt_L' or key.keysym == 'Alt_R':
            return
        if key.keysym == 'Mode_switch':
            return
        if key.state == 4:
            self.key_history.append('Ctrl+' + key.keysym)
        else:
            self.key_history.append(key.keysym)
        self.key_history.pop(0)
        #
        if key.keysym == 'F19':              # x (ctrl)
            if self.key_history[-2] == 'F19':
                if pane == self.sub:
                    self.txt.focus_set()
                    self.current_pane = 'txt'
                else:
                    self.sub.focus_set()
                    self.current_pane = 'sub'
                self.key_history[-1] = ''
            return 'break'
        elif key.keysym == 'F16':            # c (search)
            if self.key_history[-2] == 'F13':
                if self.key_history[-3] == 'F16' and \
                   self.key_history[-4] == 'F13' and \
                   Makdo.search_word != '':
                    self.search_or_replace_backward()
                else:
                    self.search_or_replace_backward_from_dialog(pane)
            else:
                if self.key_history[-2] == 'F16' and \
                   self.key_history[-3] != 'F13' and \
                   Makdo.search_word != '':
                    self.search_or_replace_forward()
                else:
                    self.search_or_replace_forward_from_dialog(pane)
            return 'break'
        elif key.keysym == 'Left':
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert-1c')
                pane.tag_add('akauni_tag', 'insert-1c', 'akauni')
        elif key.keysym == 'Right':
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert+1c')
                pane.tag_add('akauni_tag', 'insert+1c', 'akauni')
        elif key.keysym == 'Up':
            if self.key_history[-2] == 'F19':
                if pane == self.sub:
                    self.txt.focus_set()
                    self.current_pane = 'txt'
                else:
                    self.sub.focus_set()
                    self.current_pane = 'sub'
                return 'break'
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert-1l')
                pane.tag_add('akauni_tag', 'insert-1l', 'akauni')
        elif key.keysym == 'Down':
            if self.key_history[-2] == 'F19':
                if pane == self.sub:
                    self.txt.focus_set()
                    self.current_pane = 'txt'
                else:
                    self.sub.focus_set()
                    self.current_pane = 'sub'
                return 'break'
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert+1l')
                pane.tag_add('akauni_tag', 'insert+1l', 'akauni')
        elif key.keysym == 'F17':            # } (, calc)
            if self.key_history[-2] == 'F13':
                self.calculate()
                return 'break'
        elif key.keysym == 'F21':            # w (undo)
            self.edit_modified_undo()
            return 'break'
        elif key.keysym == 'XF86AudioMute':  # W (redo)
            self.edit_modified_redo()
            return 'break'
        elif key.keysym == 'F22':            # f (mark, save)
            if self.key_history[-2] == 'F19':
                self.save_file()
                return 'break'
            else:
                if 'akauni' in pane.mark_names():
                    pane.mark_unset('akauni')
                pane.mark_set('akauni', 'insert')
                return 'break'
        elif key.keysym == 'Delete':         # d (delete, quit)
            if self.key_history[-2] == 'F19':
                self.quit_makdo()
                return 'break'
            if self.key_history[-2] == 'F13':
                self.cut_rectangle()
                return 'break'
            if self.key_history[-2] != 'Delete':
                self.win.clipboard_clear()
            self._execute_when_delete_is_pressed(pane)
            return 'break'
        elif key.keysym == 'F14':            # v (quit)
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.mark_unset('akauni')
                return 'break'
        elif key.keysym == 'F15':            # g (paste)
            if self.key_history[-2] == 'F13':
                self.paste_rectangle()
                return 'break'
            self.paste_region()
            return 'break'
        elif key.keysym == 'F16':            # c (search forward)
            self.search_or_replace_forward()
            return 'break'
        elif key.keysym == 'cent':           # cent (search backward)
            self.search_or_replace_backward()
            return 'break'
        elif key.keysym == 'x':
            if self.key_history[-2] == 'Escape':
                self.MiniBuffer(pane, self)
                return 'break'
        # ctrl+a '\x01' select all          # ctrl+n '\x0e' new document
        # ctrl+b '\x02' bold                # ctrl+o '\x0f' open document
        # ctrl+c '\x03' copy                # ctrl+p '\x10' print
        # ctrl+d '\x04' font                # ctrl+q '\x11' quit
        # ctrl+e '\x05' centered            # ctrl+r '\x12' right
        # ctrl+f '\x06' search              # ctrl+s '\x13' save
        # ctrl+g '\x07' move                # ctrl+t '\x14' hanging indent
        # ctrl+h '\x08' replace             # ctrl+u '\x15' underline
        # ctrl+i '\x09' italic              # ctrl+v '\x16' paste
        # ctrl+j '\x0a' justified           # ctrl+w '\x17' close document
        # ctrl+k '\x0b' hyper link          # ctrl+x '\x18' cut
        # ctrl+l '\x0c' left                # ctrl+y '\x19' redo
        # ctrl+m '\x0d' indent              # ctrl+z '\x1a' undo
        if key.char == '\x01':    # ctrl-a
            self.select_all()
            return 'break'
        elif key.char == '\x03':  # ctrl-c
            self.copy_region()
        elif key.char == '\x05':  # ctrl-e
            self.execute_keyboard_macro()
            return 'break'
        elif key.char == '\x10':  # ctrl-p
            self.start_writer()
        elif key.char == '\x11':  # ctrl-q
            self.quit_makdo()
        elif key.char == '\x13':  # ctrl-s
            self.save_file()
        elif key.char == '\x16':  # ctrl-v
            self.paste_region()
        elif key.char == '\x18':  # ctrl-x
            self.cut_region()
        elif key.char == '\x19':  # ctrl-y
            self.edit_modified_redo()
        elif key.char == '\x1a':  # ctrl-z
            self.edit_modified_undo()
        elif key.keysym == 'Tab':
            text = pane.get('1.0', 'insert')
            line = pane.get('insert linestart', 'insert lineend')
            posi = pane.index('insert')
            # CONFIGURATION
            res_open = '^<!--(?:.|\n)*'
            res_close = '^(?:.|\n)*-->(?:.|\n)*'
            if re.match(res_open, text) and not re.match(res_close, text):
                for i, sample in enumerate(CONFIGURATION_SAMPLE):
                    if line == sample:
                        pane.delete('insert linestart', 'insert lineend')
                        pane.insert('insert', CONFIGURATION_SAMPLE[i + 1])
                        pane.mark_set('insert', 'insert linestart')
                        return 'break'
            # CALCULATE
            res_open = '^((?:.|\n)*)(<!--(?:.|\n)*)'
            res_close = '^((?:.|\n)*)(-->(?:.|\n)*)'
            if re.match(res_open, text):
                text = re.sub(res_open, '\\2', text)
                if not re.match(res_close, text):
                    self.calculate()
                    return 'break'
            # INSERT
            if re.match('^.*\\.0$', posi):
                for i, sample in enumerate(PARAGRAPH_SAMPLE):
                    if line == sample:
                        pane.delete('insert linestart', 'insert lineend')
                        pane.insert('insert', PARAGRAPH_SAMPLE[i + 1])
                        pane.mark_set('insert', 'insert linestart')
                        return 'break'
            else:
                for i, sample in enumerate(FONT_DECORATOR_SAMPLE):
                    sample_esc = sample
                    sample_esc = sample_esc.replace('*', '\\*')
                    sample_esc = sample_esc.replace('+', '\\+')
                    sample_esc = sample_esc.replace('^', '\\^')
                    cur_to_end = pane.get('insert', 'insert lineend')
                    if re.match('^' + sample_esc, cur_to_end):
                        pane.delete(posi, posi + '+' + str(len(sample)) + 'c')
                        pane.insert('insert', FONT_DECORATOR_SAMPLE[i + 1])
                        pane.mark_set('insert', posi)
                        return 'break'
        elif key.keysym == 'Prior':
            if self.key_history[-2] == 'Prior':
                if self.last_position == pane.get('insert'):
                    pane.mark_set('insert', '1.0')
        elif key.keysym == 'Next':
            if self.key_history[-2] == 'Next':
                if self.last_position == pane.get('insert'):
                    pane.mark_set('insert', 'end-1c')

    def read_only_process_key(self, pane, key):
        # HISTORY
        if key.keysym == 'Shift_L' or key.keysym == 'Shift_R':
            return
        if key.keysym == 'Control_L' or key.keysym == 'Control_R':
            return
        if key.keysym == 'Alt_L' or key.keysym == 'Alt_R':
            return
        if key.keysym == 'Mode_switch':
            return
        if key.state == 4:
            self.key_history.append('Ctrl+' + key.keysym)
        else:
            self.key_history.append(key.keysym)
        self.key_history.pop(0)
        #
        if key.keysym == 'F19':              # x (ctrl)
            if self.key_history[-2] == 'F19':
                if pane == self.sub:
                    self.txt.focus_set()
                    self.current_pane = 'txt'
                else:
                    self.sub.focus_set()
                    self.current_pane = 'sub'
                self.key_history[-1] = ''
                return 'break'
        elif key.keysym == 'Left':
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert-1c')
                pane.tag_add('akauni_tag', 'insert-1c', 'akauni')
            return
        elif key.keysym == 'Right':
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert+1c')
                pane.tag_add('akauni_tag', 'insert+1c', 'akauni')
            return
        elif key.keysym == 'Up':
            if self.key_history[-2] == 'F19':
                if pane == self.sub:
                    self.txt.focus_set()
                    self.current_pane = 'txt'
                else:
                    self.sub.focus_set()
                    self.current_pane = 'sub'
                return 'break'
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert-1l')
                pane.tag_add('akauni_tag', 'insert-1l', 'akauni')
            return
        elif key.keysym == 'Down':
            if self.key_history[-2] == 'F19':
                if pane == self.sub:
                    self.txt.focus_set()
                    self.current_pane = 'txt'
                else:
                    self.sub.focus_set()
                    self.current_pane = 'sub'
                return 'break'
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.tag_add('akauni_tag', 'akauni', 'insert+1l')
                pane.tag_add('akauni_tag', 'insert+1l', 'akauni')
            return
        elif key.keysym == 'F22':            # f (mark, save)
            if 'akauni' in pane.mark_names():
                pane.mark_unset('akauni')
            pane.mark_set('akauni', 'insert')
            return 'break'
        elif key.keysym == 'Delete':         # d (delete, quit)
            if self.key_history[-2] == 'F19':
                self.quit_makdo()
                return 'break'
            if self.key_history[-2] == 'F13':
                self.copy_rectangle()
                return 'break'
            if self.key_history[-2] != 'Delete':
                self.win.clipboard_clear()
            self._execute_when_delete_is_pressed(pane)
            return 'break'
        elif key.keysym == 'F14':            # v (quit)
            if 'akauni' in pane.mark_names():
                pane.tag_remove('akauni_tag', '1.0', 'end')
                pane.mark_unset('akauni')
                return 'break'
        elif key.keysym == 'F16':            # c (search forward)
            self.search_or_replace_forward()
            return 'break'
        elif key.keysym == 'cent':           # cent (search backward)
            self.search_or_replace_backward()
            return 'break'
        # ctrl+a '\x01' select all          # ctrl+n '\x0e' new document
        # ctrl+b '\x02' bold                # ctrl+o '\x0f' open document
        # ctrl+c '\x03' copy                # ctrl+p '\x10' print
        # ctrl+d '\x04' font                # ctrl+q '\x11' quit
        # ctrl+e '\x05' centered            # ctrl+r '\x12' right
        # ctrl+f '\x06' search              # ctrl+s '\x13' save
        # ctrl+g '\x07' move                # ctrl+t '\x14' hanging indent
        # ctrl+h '\x08' replace             # ctrl+u '\x15' underline
        # ctrl+i '\x09' italic              # ctrl+v '\x16' paste
        # ctrl+j '\x0a' justified           # ctrl+w '\x17' close document
        # ctrl+k '\x0b' hyper link          # ctrl+x '\x18' cut
        # ctrl+l '\x0c' left                # ctrl+y '\x19' redo
        # ctrl+m '\x0d' indent              # ctrl+z '\x1a' undo
        if key.char == '\x01':    # ctrl-a
            self.select_all()
            return 'break'
        elif key.char == '\x03':  # ctrl-c
            self.copy_region()
        # elif key.char == '\x05':  # ctrl-e
        #     self.execute_keyboard_macro()
        # elif key.char == '\x10':  # ctrl-p
        #     self.start_writer()
        # elif key.char == '\x11':  # ctrl-q
        #     self.quit_makdo()
        # elif key.char == '\x13':  # ctrl-s
        #     self.save_file()
        # elif key.char == '\x16':  # ctrl-v
        #     self.paste_region()
        # elif key.char == '\x18':  # ctrl-x
        #     self.cut_region()
        # elif key.char == '\x19':  # ctrl-y
        #     self.edit_modified_redo()
        # elif key.char == '\x1a':  # ctrl-z
        #     self.edit_modified_undo()
        elif key.keysym == 'Prior':
            if self.key_history[-2] == 'Prior':
                if self.last_position == pane.get('insert'):
                    pane.mark_set('insert', '1.0')
        elif key.keysym == 'Next':
            if self.key_history[-2] == 'Next':
                if self.last_position == pane.get('insert'):
                    pane.mark_set('insert', 'end-1c')
        return 'break'

    def txt_process_key_release(self, key):
        self.set_position_info_on_status_bar()
        # self.paint_out_line(self._get_v_position_of_insert() - 1)
        # FOR AKAUNI
        if 'akauni' in self.txt.mark_names():
            self.txt.tag_remove('akauni_tag', '1.0', 'end')
            self.txt.tag_add('akauni_tag', 'akauni', 'insert')
            self.txt.tag_add('akauni_tag', 'insert', 'akauni')
        self.last_position = self.txt.get('insert')

    def sub_process_key_release(self, key):
        # FOR AKAUNI
        if 'akauni' in self.sub.mark_names():
            self.sub.tag_remove('akauni_tag', '1.0', 'end')
            self.sub.tag_add('akauni_tag', 'akauni', 'insert')
            self.sub.tag_add('akauni_tag', 'insert', 'akauni')
        self.last_position = self.sub.get('insert')
        return 'break'

    # MOUSE BUTTON LEFT

    def txt_process_button1(self, click):
        self.current_pane = 'txt'
        self.txt.focus_set()
        return

    def txt_process_button1_release(self, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        self.set_position_info_on_status_bar()
        return 'break'

    def sub_process_button1(self, click):
        self.current_pane = 'sub'
        self.sub.focus_set()
        return

    def sub_process_button1_release(self, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        return 'break'

    # MOUSE BUTTON CENTER

    def txt_process_button2(self, click):
        return 'break'

    def txt_process_button2_release(self, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        # self.paste_region()
        return 'break'

    def sub_process_button2(self, click):
        return 'break'

    def sub_process_button2_release(self, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        # self.paste_region()
        return 'break'

    # MOUSE BUTTON RIGHT

    def txt_process_button3(self, click):
        self.any_process_button3(self.txt, click)
        return 'break'

    def txt_process_button3_release(self, click):
        return 'break'

    def sub_process_button3(self, click):
        self.any_process_button3(self.sub, click)
        return 'break'

    def sub_process_button3_release(self, click):
        return 'break'

    def any_process_button3(self, pane, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        self.bt3 = tkinter.Menu(self.win, tearoff=False)
        if not self._is_read_only_pane(pane):
            self.bt3.add_command(label='切り取り',
                                 command=self.cut_region)
        self.bt3.add_command(label='コピー',
                             command=self.copy_region)
        if not self._is_read_only_pane(pane):
            try:
                cb = self.win.clipboard_get()
            except BaseException:
                cb = ''
            if cb != '':
                self.bt3.add_command(label='貼り付け',
                                     command=self.paste_region)
        self.bt3.post(click.x_root, click.y_root)

    ####################################
    # STATUS BAR

    def _make_status_bar(self):
        self._make_status_search_or_replace()
        self._make_status_file_name()
        self._make_status_position_information()
        self._make_status_message()

    ##########################
    # STATUS FILE NAME

    def _make_status_file_name(self):
        self.stb_fnm1 = tkinter.Label(self.stb, anchor='w', text='')
        self.stb_fnm1.pack(side='left')
        tkinter.Label(self.stb, text=' ').pack(side='left')

    ################
    # COMMAND

    def set_file_name_on_status_bar(self, file_name, must_update=False):
        fn = file_name
        fn = re.sub('\n', '/', fn)
        res = '^(.*)(\\..{1,4})$'
        if re.match(res, fn):
            nam = re.sub(res, '\\1', fn)
            ext = re.sub(res, '\\2', fn)
        else:
            nam = fn
            ext = ''
        if len(fn) > 15:
            nam = re.sub('^(.{' + str(14 - len(ext)) + '})(.*)', '\\1…', nam)
        self.stb_fnm1['text'] = nam + ext
        if must_update:
            self.stb.update()

    ##########################
    # STATUS POSITION INFORMATION

    def _make_status_position_information(self):
        self.stb_pos1 = tkinter.Label(self.stb, anchor='w', text='1x0/1x0')
        self.stb_pos1.pack(side='left')
        tkinter.Label(self.stb, text=' ').pack(side='left')

    ################
    # COMMAND

    def set_position_info_on_status_bar(self, must_update=False):
        p = self.txt.index('insert')
        cur_v = re.sub('\\.[0-9]+$', '', p)
        s = self.txt.get('insert linestart', 'insert')
        cur_h = str(get_ideal_width(s))
        cur_p = cur_v + 'x' + cur_h
        p = self.txt.index('end-1c')
        max_v = re.sub('\\.[0-9]+$', '', p)
        s = self.txt.get('insert linestart', 'insert lineend')
        max_h = str(get_ideal_width(s))
        max_p = max_v + 'x' + max_h
        self.stb_pos1['text'] = cur_p + '/' + max_p
        if must_update:
            self.stb.update()

    ##########################
    # STATUS MESSAGE

    def _make_status_message(self):
        self.stb_msg1 = tkinter.Label(self.stb, anchor='w', text='')
        self.stb_msg1.pack(side='left')
        # tkinter.Label(self.stb, text=' ').pack(side=tkinter.LEFT)

    ################
    # COMMAND

    def set_message_on_status_bar(self, msg, must_update=False):
        self.stb_msg1['text'] = msg
        if must_update:
            self.stb.update()

    ##########################
    # STATUS SEARCH OR REPLACE

    def _make_status_search_or_replace(self):
        tkinter.Label(self.stbr, text=' ').pack(side='left')
        # tkinter.Label(self.stb, text='探').pack(side='left')
        self.stb_sor1 = tkinter.Entry(self.stbr, width=20)
        self.stb_sor1.pack(side='left')
        # self.stb_sor1.insert(0, '（検索語）')
        # tkinter.Label(self.stb, text='換').pack(side='left')
        self.stb_sor2 = tkinter.Entry(self.stbr, width=20)
        self.stb_sor2.pack(side='left')
        # self.stb_sor2.insert(0, '（置換語）')
        self.stb_sor3 = tkinter.Button(self.stbr, text='前',
                                       command=self.search_or_replace_backward)
        self.stb_sor3.pack(side='left')
        self.stb_sor4 = tkinter.Button(self.stbr, text='次',
                                       command=self.search_or_replace_forward)
        self.stb_sor4.pack(side='left')
        self.stb_sor5 = tkinter.Button(self.stbr, text='消',
                                       command=self.clear_search_or_replace)
        self.stb_sor5.pack(side='left')

    ################
    # COMMAND

    def search_or_replace_backward(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        word1 = self.stb_sor1.get()
        word2 = self.stb_sor2.get()
        if word1 == '':
            return
        if Makdo.search_word != word1:
            Makdo.search_word = word1
            if word1 != '':
                self._highlight_search_word()
        tex = pane.get('1.0', 'insert')
        tex = re.sub(word1 + '$', '', tex)
        res = '^((?:.|\n)*?)(' + word1 + ')((?:.|\n)*)$'
        if re.match(res, tex):
            sub = ''
            while re.match(res, tex):
                s = re.sub(res, '\\1', tex)
                w = re.sub(res, '\\2', tex)
                tex = re.sub(res, '\\3', tex)
                sub += s + w
                wrd = w
                if wrd == '':
                    return
            if wrd == '':
                return
            # SEARCH
            pane.mark_set('insert', '1.0 +' + str(len(sub)) + 'c')
            pane.yview('insert -10line')
            if word2 != '':  # and word2 != '（置換語）'
                if not self._is_read_only_pane(pane):
                    # REPLACE
                    pane.delete('insert-' + str(len(wrd)) + 'c', 'insert')
                    pane.insert('insert', word2)
        pane.focus_set()
        # MESSAGE
        n, m = self._count_word(pane, word1)
        self.set_message_on_status_bar(str(m) + '個が見付かりました' +
                                       '（' + str(n) + '/' + str(m) + '）')
        self.stb.update()

    def search_or_replace_forward(self):
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        word1 = self.stb_sor1.get()
        word2 = self.stb_sor2.get()
        if word1 == '':
            return
        if Makdo.search_word != word1:
            Makdo.search_word = word1
            if word1 != '':
                self._highlight_search_word()
        tex = pane.get('insert', 'end-1c')
        res = '^((?:.|\n)*?)(' + word1 + ')(?:.|\n)*$'
        if re.match(res, tex):
            sub = re.sub(res, '\\1\\2', tex)
            wrd = re.sub(res, '\\2', tex)
            if wrd == '':
                return
            # SEARCH
            pane.mark_set('insert', 'insert +' + str(len(sub)) + 'c')
            pane.yview('insert -10line')
            if word2 != '':  # and word2 != '（置換語）'
                if not self._is_read_only_pane(pane):
                    # REPLACE
                    pane.delete('insert-' + str(len(wrd)) + 'c', 'insert')
                    pane.insert('insert', word2)
        pane.focus_set()
        # MESSAGE
        n, m = self._count_word(pane, word1)
        self.set_message_on_status_bar(str(m) + '個が見付かりました' +
                                       '（' + str(n) + '/' + str(m) + '）')
        self.stb.update()

    def _count_word(self, pane, word):
        res = '^((?:.|\n)*?' + word + ')((?:.|\n)*)$'
        #
        x = 0
        tex = pane.get('1.0', 'insert')
        while re.match(res, tex):
            x += 1
            pre = re.sub(res, '\\1', tex)
            tex = re.sub(res, '\\2', tex)
            if pre == '':
                break
        #
        y = 0
        tex = pane.get('insert', 'end-1c')
        while re.match(res, tex):
            y += 1
            pre = re.sub(res, '\\1', tex)
            tex = re.sub(res, '\\2', tex)
            if pre == '':
                break
        return x, x + y

    def clear_search_or_replace(self):
        self.stb_sor1.delete('0', 'end')
        self.stb_sor2.delete('0', 'end')
        self.txt.tag_remove('search_tag', '1.0', 'end')
        Makdo.search_word = ''

    def _highlight_search_word(self):
        self.txt.tag_remove('search_tag', '1.0', 'end')
        if self.current_pane == 'sub':
            pane = self.sub
        else:
            pane = self.txt
        word = Makdo.search_word
        tex = pane.get('1.0', 'end-1c')
        beg = 0
        res = '^((?:.|\n)*?)(' + word + ')((?:.|\n)*)$'
        while re.match(res, tex):
            pre = re.sub(res, '\\1', tex)
            wrd = re.sub(res, '\\2', tex)
            tex = re.sub(res, '\\3', tex)
            if wrd == '':
                break
            beg += len(pre)
            end = beg + len(wrd)
            pane.tag_add('search_tag',
                         '1.0+' + str(beg) + 'c',
                         '1.0+' + str(end) + 'c',)
            beg = end

    def search_or_replace_backward_from_dialog(self, pane):
        t = '前検索'
        m = '検索する言葉と置換する言葉を入力してください．'
        word1 = self.stb_sor1.get()
        word2 = self.stb_sor2.get()
        sd = TwoWordsDialog(pane, self, t, m, word1, word2)
        word1, word2 = sd.get_value()
        if word1 is not None and word2 is not None:
            if word1 == '':
                self.clear_search_or_replace()
            else:
                Makdo.search_word = word1
                self._highlight_search_word()
                self.stb_sor1.delete(0, 'end')
                self.stb_sor1.insert(0, word1)
                self.stb_sor2.delete(0, 'end')
                self.stb_sor2.insert(0, word2)
                self.search_or_replace_backward()

    def search_or_replace_forward_from_dialog(self, pane):
        t = '後検索'
        m = '検索する言葉と置換する言葉を入力してください．'
        word1 = self.stb_sor1.get()
        word2 = self.stb_sor2.get()
        sd = TwoWordsDialog(pane, self, t, m, word1, word2)
        word1, word2 = sd.get_value()
        if word1 is not None and word2 is not None:
            if word1 == '':
                self.clear_search_or_replace()
            else:
                Makdo.search_word = word1
                self._highlight_search_word()
                self.stb_sor1.delete(0, 'end')
                self.stb_sor1.insert(0, word1)
                self.stb_sor2.delete(0, 'end')
                self.stb_sor2.insert(0, word2)
                self.search_or_replace_forward()

    ####################################
    # SHOW MESSAGE

    def show_first_help_message(self):
        if self.dont_show_help.get():
            return
        n = 'ご説明'
        m = 'MS Word形式（拡張子docx）の\n' + \
            'ファイルを、この画面に\n' + \
            'ドラッグ＆ドロップしてみてください．\n' + \
            'Markdown形式に変換されて、\n' + \
            '画面に表示されます．\n\n' + \
            'その内容を編集して保存することで、\n' + \
            'MS Word形式（拡張子docx）の\n' + \
            'ファイルを編集できます．\n\n' + \
            '編集方法が分からない場合は、\n' + \
            'MS Wordで必要な編集したものを\n' + \
            'このアプリで開いてみて、\n' + \
            '編集前と見比べてください．'
        tkinter.messagebox.showinfo(n, m)

    def show_folding_help_message(self):
        if self.dont_show_help.get():
            return
        if not self.must_show_folding_help_message:
            return
        n = 'ご説明'
        m = 'セクションを折り畳みます．' + \
            '（セクションの中身を一時的に文面の最後に移動させます）．\n\n' + \
            'そうすることで、' + \
            '文面の構造を視覚的に把握しやすくできます．\n\n' + \
            '他方で、' + \
            '一時的に文の順序が入れ替わってしまいますので、' + \
            'コメントや下線などの範囲を正しく把握できず、' + \
            '画面上の見た目が崩れる可能性があります．\n\n' + \
            'ファイルを保存する際には、' + \
            '全て展開した状態で保存されます．\n\n' + \
            '注）"...[n]"という記号は、' + \
            '折り畳んだことを記録したもので展開する際に必要ですので、' + \
            '絶対に書き替えたり消したりしないでください．'
        tkinter.messagebox.showinfo(n, m)
        self.must_show_folding_help_message = False

    def show_keyboard_macro_help_message(self):
        if self.dont_show_help.get():
            return
        if not self.must_show_keyboard_macro_help_message:
            return
        n = 'ご説明'
        m = 'キー入力の中から、繰り返しを探して、\n' + \
            'その繰り返しを実行します．\n\n' + \
            '同じ作業を何度も繰り返すときに、\n' + \
            '便利です．\n\n' + \
            '"Ctrl+E"でも実行できます．'
        tkinter.messagebox.showinfo(n, m)
        self.must_show_keyboard_macro_help_message = False

    def show_config_help_message(self):
        if self.dont_show_help.get():
            return
        if not self.must_show_config_help_message:
            return
        n = 'ご説明'
        m = '設定を次回以降に引き継ぐ場合は、\n' + \
            '「設定」の項目の「設定を保存」を\n' + \
            'クリックして、保存してください．'
        tkinter.messagebox.showinfo(n, m)
        self.must_show_config_help_message = False

    ####################################
    # RUN PERIODICALLY

    def run_periodically(self):
        self.run_periodically_to_paint_line()
        self.run_periodically_to_set_position_info()
        self.run_periodically_to_save_auto_file()
        self.run_periodically_to_update_memo_pad()

    ##########################
    # COMMAND

    def run_periodically_to_paint_line(self):
        # GLOBAL PAINTING
        self.paint_out_line(self.global_line_to_paint)
        self.global_line_to_paint += 1
        if self.global_line_to_paint >= len(self.file_lines) - 1:
            self.global_line_to_paint = 0
        # LOCAL PAINTING
        self.paint_out_line(self.standard_line + self.local_line_to_paint - 10)
        self.local_line_to_paint += 1
        if self.local_line_to_paint >= 150:
            i = self.txt.index('insert')
            self.standard_line = int(re.sub('\\..*$', '', i)) - 1
            self.local_line_to_paint = 0
        # POSITION PAINTING
        # self.paint_out_line(self._get_v_position_of_insert() - 1)
        # LINE AND EOF PAINTING

        ii = self.txt.index('insert lineend +1c')
        ei = self.txt.index('end lineend')
        self.txt.tag_remove('line_eof_tag', '1.0', 'end')
        self.txt.tag_remove('line_tag', '1.0', 'end')
        self.txt.tag_remove('eof_tag', '1.0', 'end')
        if ii == ei:
            # LINE EOF PAINTING
            self.txt.tag_add('line_eof_tag',
                             'insert lineend', 'insert lineend +1c')
        else:
            # LINE PAINTING
            self.txt.tag_add('line_tag',
                             'insert lineend', 'insert lineend +1c')
            # EOF PAINTING
            self.txt.tag_add('eof_tag',
                             'end-1c', 'end')
        # TO NEXT
        interval = 10
        self.win.after(interval, self.run_periodically_to_paint_line)  # NEXT

    def run_periodically_to_set_position_info(self):
        self.set_position_info_on_status_bar()
        interval = 100
        self.win.after(interval, self.run_periodically_to_set_position_info)

    def run_periodically_to_save_auto_file(self):
        self.save_auto_file(self.file_path)
        interval = 60_000
        self.win.after(interval, self.run_periodically_to_save_auto_file)

    def run_periodically_to_update_memo_pad(self):
        self.update_memo_pad()
        interval = 1_000
        self.win.after(interval, self.run_periodically_to_update_memo_pad)


######################################################################
# MAIN


class Splash:

    def __init__(self, image):
        self.splash = tkinter.Tk()
        sw = self.splash.winfo_screenwidth()
        sh = self.splash.winfo_screenheight()
        self.splash_img = tkinter.PhotoImage(data=image, master=self.splash)
        iw = self.splash_img.width()
        ih = self.splash_img.height()
        size = str(iw) + 'x' + str(ih)
        position = str(int((sw - iw) / 2)) + '+' + str(int((sh - ih) / 2))
        self.splash.geometry(size + '+' + position)
        self.splash.overrideredirect(1)  # no title bar
        canvas = tkinter.Canvas(self.splash, bg=None, width=iw, height=ih)
        canvas.place(x=-1, y=-1)
        canvas.create_image(0, 0, image=self.splash_img, anchor='nw')
        self.splash.after(10000, self.destroy_myself)

    def destroy_myself(self):
        self.splash_img = None
        self.splash.destroy()


if __name__ == '__main__':

    # SPLASH
    if getattr(sys, 'frozen', False):
        import pyi_splash
        pyi_splash.close()
    else:
        splash = Splash(SPLASH_IMG)

    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description='Markdownファイルを編集します',
        add_help=False)
    parser.add_argument(
        '-h', '--help',
        action='help',
        help='ヘルプメッセージを表示します')
    parser.add_argument(
        '-v', '--version',
        action='version',
        version=('%(prog)s ' + __version__),
        help='バージョン番号を表示します')
    parser.add_argument(
        '-H', '--dont-show-help',
        action='store_true',
        help='ヘルプを表示します')
    parser.add_argument(
        '-c', '--background-color',
        type=str,
        choices=['W', 'B', 'G'],
        help='背景の色（白、黒、緑）を指定します')
    parser.add_argument(
        '-s', '--font-size',
        type=int,
        choices=[12, 15, 18, 21, 24, 27, 30, 33, 36, 42, 48],
        help='文字の大きさをピクセル単位で指定します')
    parser.add_argument(
        '-p', '--paint-keywords',
        action='store_true',
        help='キーワードに色を付けます')
    parser.add_argument(
        '-d', '--digit-separator',
        type=str,
        choices=['3', '4'],
        help='計算結果の区切りを指定します')
    parser.add_argument(
        '-r', '--read-only',
        action='store_true',
        help='読み取り専用で開きます')
    parser.add_argument(
        '-b', '--make-backup-file',
        action='store_true',
        help='バックアップファイルを残します')
    parser.add_argument(
        'input_file',
        nargs='?',
        help='Markdownファイル or MS Wordファイル')
    args = parser.parse_args()

    Makdo.args_dont_show_help = args.dont_show_help
    Makdo.args_background_color = args.background_color
    Makdo.args_font_size = args.font_size
    Makdo.args_paint_keywords = args.paint_keywords
    Makdo.args_digit_separator = args.digit_separator
    Makdo.args_read_only = args.read_only
    Makdo.args_make_backup_file = args.make_backup_file
    Makdo.args_input_file = args.input_file

    Makdo()
