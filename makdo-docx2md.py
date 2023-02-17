#!/usr/bin/python3
# Name:         docx2md.py
# Version:      v05a Aki-Nagatsuka
# Time-stamp:   <2023.02.18-08:27:10-JST>

# docx2md.py
# Copyright (C) 2022-2023  Seiichiro HATA
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
# 20XX.XX.XX v05 Aki-Nagatsuka


############################################################
# SETTING


import os
import sys
import tempfile
import shutil
import argparse
import re
import unicodedata
import datetime


__version__ = 'v05 Aki-Nagatsuka'


def get_arguments():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description='MS WordファイルからMarkdownファイルを作ります',
        add_help=False,
        epilog=HELP_EPILOG)
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
        '-T', '--document-title',
        type=str,
        metavar='STRING',
        help='文書の標題')
    parser.add_argument(
        '-p', '--paper-size',
        type=str,
        choices=['A3', 'A3L', 'A3P', 'A4', 'A4L', 'A4P'],
        help='用紙設定（A3、A3L、A3P、A4、A4L、A4P）')
    parser.add_argument(
        '-t', '--top-margin',
        type=float,
        metavar='NUMBER',
        help='上余白（単位cm）')
    parser.add_argument(
        '-b', '--bottom-margin',
        type=float,
        metavar='NUMBER',
        help='下余白（単位cm）')
    parser.add_argument(
        '-l', '--left-margin',
        type=float,
        metavar='NUMBER',
        help='左余白（単位cm）')
    parser.add_argument(
        '-r', '--right-margin',
        type=float,
        metavar='NUMBER',
        help='右余白（単位cm）')
    parser.add_argument(
        '-d', '--document-style',
        type=str,
        choices=['k', 'j'],
        help='文書スタイルの指定（契約、条文）')
    parser.add_argument(
        '-H', '--header-string',
        type=str,
        help='ヘッダーの文字列')
    parser.add_argument(
        '-P', '--page-number',
        type=str,
        help='ページ番号の書式')
    parser.add_argument(
        '-L', '--line-number',
        action='store_true',
        help='行番号を出力します')
    parser.add_argument(
        '-m', '--mincho-font',
        type=str,
        metavar='FONT_NAME',
        help='明朝フォント')
    parser.add_argument(
        '-g', '--gothic-font',
        type=str,
        metavar='FONT_NAME',
        help='ゴシックフォント')
    parser.add_argument(
        '-f', '--font-size',
        type=float,
        metavar='NUMBER',
        help='フォントサイズ（単位pt）')
    parser.add_argument(
        '-s', '--line-spacing',
        type=float,
        metavar='NUMBER',
        help='行間の高さ（単位文字）')
    parser.add_argument(
        '-B', '--space-before',
        type=floats6,
        metavar='NUMBER,NUMBER,...',
        help='タイトル前の空白')
    parser.add_argument(
        '-A', '--space-after',
        type=floats6,
        metavar='NUMBER,NUMBER,...',
        help='タイトル後の空白')
    parser.add_argument(
        '-a', '--auto-space',
        action='store_true',
        help='全角文字と半角文字との間の間隔を微調整します')
    parser.add_argument(
        'docx_file',
        help='MS Wordファイル')
    parser.add_argument(
        'md_file',
        default='',
        nargs='?',
        help='Markdownファイル（"-"は標準出力）')
    return parser.parse_args()


def floats6(s):
    if not re.match('^' + RES_NUMBER6 + '$', s):
        raise argparse.ArgumentTypeError
    return s


HELP_EPILOG = '''
'''

DEFAULT_DOCUMENT_TITLE = ''

DEFAULT_DOCUMENT_STYLE = 'n'

DEFAULT_PAPER_SIZE = 'A4'
PAPER_HEIGHT = {'A3': 29.7, 'A3L': 29.7, 'A3P': 42.0,
                'A4': 29.7, 'A4L': 21.0, 'A4P': 29.7}
PAPER_WIDTH = {'A3': 42.0, 'A3L': 42.0, 'A3P': 29.7,
               'A4': 21.0, 'A4L': 29.7, 'A4P': 21.0}

DEFAULT_TOP_MARGIN = 3.5
DEFAULT_BOTTOM_MARGIN = 2.2
DEFAULT_LEFT_MARGIN = 3.0
DEFAULT_RIGHT_MARGIN = 2.0

DEFAULT_HEADER_STRING = ''

DEFAULT_PAGE_NUMBER = 'n'

DEFAULT_LINE_NUMBER = False

DEFAULT_MINCHO_FONT = 'ＭＳ 明朝'
DEFAULT_GOTHIC_FONT = 'ＭＳ ゴシック'
DEFAULT_FONT_SIZE = 12.0

DEFAULT_LINE_SPACING = 2.14  # (2.0980+2.1812)/2=2.1396

DEFAULT_SPACE_BEFORE = ''
DEFAULT_SPACE_AFTER = ''

DEFAULT_AUTO_SPACE = False

NOT_ESCAPED = '^((?:(?:.*\n)*.*[^\\\\])?(?:\\\\\\\\)*)?'

RES_NUMBER = '(?:[-\\+]?(?:(?:[0-9]+(?:\\.[0-9]+)?)|(?:\\.[0-9]+)))'
RES_NUMBER6 = '(?:' + RES_NUMBER + '?,){,5}' + RES_NUMBER + '?,?'

FONT_DECORATORS = [
    '\\*\\*\\*',           # italic and bold
    '\\*\\*',              # bold
    '\\*',                 # italic
    '~~',                  # strikethrough
    '`',                   # preformatted
    '//',                  # italic
    '__',                  # underline
    '\\-\\-',              # small
    '\\+\\+',              # large
    '\\^[0-9A-Za-z]*\\^',  # font color
    '_[0-9A-Za-z]+_',      # higilight color
]
RES_FONT_DECORATORS = '((?:' + '|'.join(FONT_DECORATORS) + ')*)'

MD_TEXT_WIDTH = 68

FONT_COLOR = {
    'FF0000': 'red',
    # 'FF0000': 'R',
    '7F0000': 'darkRed',
    # '7F0000': 'DR',
    'FFFF00': 'yellow',
    # 'FFFF00': 'Y',
    '7F7F00': 'darkYellow',
    # '7F7F00': 'DY',
    '00FF00': 'green',
    # '00FF00': 'G',
    '007F00': 'darkGreen',
    # '007F00': 'DG',
    '00FFFF': 'cyan',
    # '00FFFF': 'C',
    '007F7F': 'darkCyan',
    # '007F7F': 'DC',
    '0000FF': 'blue',
    # '0000FF': 'B',
    '00007F': 'darkBlue',
    # '00007F': 'DB',
    'FF00FF': 'magenta',
    # 'FF00FF': 'M',
    '7F007F': 'darkMagenta',
    # '7F007F': 'DM',
    'BFBFBF': 'lightGray',
    # 'BFBFBF': 'G1',
    '7F7F7F': 'darkGray',
    # '7F7F7F': 'G2',
    '000000': 'black',
    # '000000': 'BK',
    'FF5D5D': 'a000',
    'FF603C': 'a010',
    'FF6512': 'a020',
    'E07000': 'a030',
    'BC7A00': 'a040',
    'A08300': 'a050',
    '898900': 'a060',
    '758F00': 'a070',
    '619500': 'a080',
    '4E9B00': 'a090',
    '38A200': 'a100',
    '1FA900': 'a110',
    '00B200': 'a120',
    '00AF20': 'a130',
    '00AC3C': 'a140',
    '00AA55': 'a150',
    '00A76D': 'a160',
    '00A586': 'a170',
    '00A2A2': 'a180',
    '009FC3': 'a190',
    '009AED': 'a200',
    '1F8FFF': 'a210',
    '4385FF': 'a220',
    '5F7CFF': 'a230',
    '7676FF': 'a240',
    '8A70FF': 'a250',
    '9E6AFF': 'a260',
    'B164FF': 'a270',
    'C75DFF': 'a280',
    'E056FF': 'a290',
    'FF4DFF': 'a300',
    'FF50DF': 'a310',
    'FF53C3': 'a320',
    'FF55AA': 'a330',
    'FF5892': 'a340',
    'FF5A79': 'a350',
}

CONJUNCTIONS = [
    # 複合
    'しかし[，、]だからといって',
    # 単一
    '(?:こ|そ|あ|ど)うなると',
    '(?:こ|そ|あ|ど)うなれば',
    '(?:こ|そ|あ|ど)のうえ', '(?:こ|そ|あ|ど)の上',
    '(?:こ|そ|あ|ど)のうえで', '(?:こ|そ|あ|ど)の上で',
    '(?:こ|そ|あ|ど)のかわり', '(?:こ|そ|あ|ど)の代わり',
    '(?:こ|そ|あ|ど)のくせ',
    '(?:こ|そ|あ|ど)のことから',
    '(?:こ|そ|あ|ど)のことから',
    '(?:こ|そ|あ|ど)のため',
    '(?:こ|そ|あ|ど)のためには',
    '(?:こ|そ|あ|ど)のなかでも', '(?:こ|そ|あ|ど)の中でも',
    '(?:こ|そ|あ|ど)のような中',
    '(?:こ|そ|あ|ど)のように',
    '(?:こ|そ|あ|ど)のようにして',
    '(?:こ|そ|あ|ど)の反面',
    '(?:こ|そ|あ|ど)の場合',
    '(?:こ|そ|あ|ど)の後',
    '(?:こ|そ|あ|ど)の結果',
    '(?:こ|そ|あ|ど)の際',
    '(?:こ|そ|あ|ど)れから',
    '(?:こ|そ|あ|ど)れで',
    '(?:こ|そ|あ|ど)れでこそ',
    '(?:こ|そ|あ|ど)れでは',
    '(?:こ|そ|あ|ど)れでは',
    '(?:こ|そ|あ|ど)れでも',
    '(?:こ|そ|あ|ど)れどころか',
    '(?:こ|そ|あ|ど)れなのに',
    '(?:こ|そ|あ|ど)れなら',
    '(?:こ|そ|あ|ど)れに',
    '(?:こ|そ|あ|ど)れにしても',
    '(?:こ|そ|あ|ど)れには',
    '(?:こ|そ|あ|ど)れにもかかわらず',
    '(?:こ|そ|あ|ど)れによって',
    '(?:こ|そ|あ|ど)れに加えて',
    '(?:こ|そ|あ|ど)れに対して',
    '(?:こ|そ|あ|ど)ればかりか',
    '(?:こ|そ|あ|ど)ればかりでなく',
    '(?:こ|そ|あ|ど)れゆえ', '(?:こ|そ|あ|ど)れ故',
    '(?:こ|そ|あ|ど)れゆえに', '(?:こ|そ|あ|ど)れ故に',
    '(?:こ|そ|あ|ど)れより',
    '(?:こ|そ|あ|ど)れよりは',
    '(?:こ|そ|あ|ど)れよりも',
    '(?:こ|そ|あ|ど)れらのことから',
    '(?:こ|そ|あ|ど)れらを踏まえて',
    '(?:こ|そ|あ|ど)んな中',
    '(?:こ|そ|あそ|ど)こで',
    '(?:こう|そう|ああ|どう)いえば',
    '(?:こう|そう|ああ|どう)したところ',
    '(?:こう|そう|ああ|どう)したら',
    '(?:こう|そう|ああ|どう)して',
    '(?:こう|そう|ああ|どう)してみると',
    '(?:こう|そう|ああ|どう)しなければ',
    '(?:こう|そう|ああ|どう)することで',
    '(?:こう|そう|ああ|どう)すると',
    '(?:こう|そう|ああ|どう)すれば',
    '(?:こう|そう|ああ|どう)だからといって',
    '(?:こう|そう|ああ|どう)だとしても',
    '(?:こう|そう|ああ|どう)だとすると',
    '(?:こう|そう|ああ|どう)だとすれば',
    '(?:こう|そう|ああ|どう)であるにもかかわらず',
    '(?:こう|そう|ああ|どう)でないならば',
    '(?:こう|そう|ああ|どう)ではあるが',
    '(?:こう|そう|ああ|どう)ではなく',
    '(?:こう|そう|ああ|どう)はいうものの',
    '[1-9１-９一二三四五六七八九]つ目は',
    '[1-9１-９一二三四五六七八九]点目は',
    '[1１一]つは', 'もう[1１一]つは', '[2-9２-９二三四五六七八九]つには',
    '[1１一]点は', 'もう[1１一]点は',
    'あと', '後',
    'あるいは',
    'いうならば', '言うならば',
    'いうなれば', '言うなれば',
    'いずれにしても',
    'いずれにしろ',
    'いずれにせよ',
    'いってみれば', '言ってみれば',
    'いわば',
    'いわんや',
    'おまけに',
    'および', '及び',
    'かえって', '却って', '反って',
    'かくして', '斯くして',
    'かつ', '且つ',
    'が',
    'けだし', '蓋し',
    'けど',
    'けれど',
    'けれども',
    'さて',
    'さもないと',
    'さらに', '更に',
    'しかし',
    'しかしながら',
    'しかも',
    'しかるに', '然るに',
    'したがって', '従って',
    'してみると',
    'じつは', '実は',
    'すなわち',
    'すると',
    'そして',
    'そもそも',
    'それとも',
    'それはさておき',
    'それはそうと',
    'たしかに', '確かに',
    'ただ',
    'ただし',
    'たとえば', '例えば',
    'だから',
    'だからこそ',
    'だからといって',
    'だが',
    'だけど',
    'だって',
    'だとしたら',
    'だとしても',
    'だとすると',
    'だとすれば',
    'ちなみに', '因みに',
    'つぎに', '次に',
    'つまり',
    'つまるところ', '詰まる所',
    'ですが',
    'では',
    'でも',
    'というか',
    'というのは',
    'というのも',
    'というより',
    'というよりも',
    'ときに', '時に',
    'ところが',
    'ところで',
    'となると',
    'となれば',
    'とにかく',
    'とにもかくにも',
    'とはいうものの',
    'とはいえ',
    'とはいっても',
    'ともあれ',
    'ともかく',
    'とりわけ', '取分け',
    'どころか',
    'どちらにしても',
    'どちらにせよ',
    'どっちにしても',
    'どっちにせよ',
    'どっち道', 'どっちみち',
    'どのみち', 'どの道',
    'なお', '尚',
    'なおさら', '尚更',
    'なかでも', '中でも',
    'なぜかというと', '何故かというと',
    'なぜかといえば', '何故かといえば',
    'なぜなら', '何故なら',
    'なぜならば', '何故ならば',
    'なにしろ', '何しろ',
    'なにせ', '何せ',
    'なので',
    'なのに',
    'ならば',
    'ならびに', '並びに',
    'なるほど', '成程',
    'にもかかわらず',
    'のに',
    'はじめに', '初めに', '始めに', 'おわりに', '終わりに', '終りに',
    'ひいては', '延いては',
    'まして',
    'ましてや',
    'まず', '先ず',
    'また', '又',
    'または', '又は',
    'むしろ',
    'むろん', '無論',
    'もし',
    'もしかしたら',
    'もしくは', '若しくは',
    'もしも',
    'もちろん', '勿論',
    'もっとも', '尤も',
    'ものの',
    'ゆえに', '故に',
    'よって', '因って',
    '一方', '他方',
    '一方で', '他方で',
    '一方では', '他方では',
    '一般的',
    '一般的に',
    '事実',
    '他には',
    '他にも',
    '以上',
    '以上から',
    '以上のように',
    '以上を踏まえて',
    '仮に',
    '仮にも',
    '具体的には',
    '加えて',
    '反対に',
    '反面',
    '同じく',
    '同じように',
    '同時に',
    '同様に',
    '実のところ',
    '実を言うと',
    '実を言えば',
    '実際',
    '実際に',
    '対して',
    '当たり前ですが',
    '当然ですが',
    '換言すると',
    '普通',
    '最初に', '最後に',
    '次いで',
    '殊に',
    '特に',
    '現に',
    '百歩譲って',
    '百歩譲って仮に',
    '第[1-9１-９一二三四五六七八九]に',
    '結局',
    '結果として',
    '結果的に',
    '続いて',
    '裏を返せば',
    '裏返せば',
    '要するに',
    '要は',
    '言い換えると',
    '言ってみれば',
    '逆に',
    '逆に言えば',
    '通常',
]


############################################################
# FUNCTION


def get_real_width(s):
    p = ''
    wid = 0.0
    for c in s:
        if (c == '\t'):
            wid = (wid + 8) // 8 * 8
            continue
        w = unicodedata.east_asian_width(c)
        if c == '':
            wid += 0.0
        elif re.match('^[−☐☑]$', c):
            wid += 2.0
        elif re.match('^[´¨―‐∥…‥‘’“”±×÷≠≦≧∞∴♂♀°′″℃§]$', c):
            wid += 2.0
        elif re.match('^[☆★○●◎◇◆□■△▲▽▼※→←↑↓]$', c):
            wid += 2.0
        elif re.match('^[∈∋⊆⊇⊂⊃∪∩∧∨⇒⇔∀∃∠⊥⌒∂∇≡≒≪≫√∽∝∵]$', c):
            wid += 2.0
        elif re.match('^[∫∬Å‰♯♭♪†‡¶◯]$', c):
            wid += 2.0
        elif re.match('^[ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ]$', c):
            wid += 2.0
        elif re.match('^[αβγδεζηθικλμνξοπρστυφχψω]$', c):
            wid += 2.0
        elif re.match('^[АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ]$', c):
            wid += 2.0
        elif re.match('^[абвгдеёжзийклмнопрстуфхцчшщъыьэюя]$', c):
            wid += 2.0
        elif re.match('^[─│┌┐┘└├┬┤┴┼━┃┏┓┛┗┣┳┫┻╋┠┯┨┷┿┝┰┥┸╂]$', c):
            wid += 2.0
        elif re.match('^[№℡≒≡∫∮∑√⊥∠∟⊿∵∩∪]$', c):
            wid += 2.0
        elif re.match('^[⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇]$', c):
            wid += 2.0
        elif re.match('^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]$', c):
            wid += 2.0
        elif re.match('^[⒈⒉⒊⒋⒌⒍⒎⒏⒐⒑⒒⒓⒔⒕⒖⒗⒘⒙⒚⒛]$', c):
            wid += 2.0
        elif re.match('^[ⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅺⅻ]$', c):
            wid += 2.0
        elif re.match('^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]$', c):
            wid += 2.0
        elif re.match('^[⒜⒝⒞⒟⒠⒡⒢⒣⒤⒥⒦⒧⒨⒩⒪⒫⒬⒭⒮⒯⒰⒱⒲⒳⒴⒵]$', c):
            wid += 2.0
        elif re.match('^[ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙⓚⓛⓜⓝⓞⓟⓠⓡⓢⓣⓤⓥⓦⓧⓨⓩ]$', c):
            wid += 2.0
        elif re.match('^[🄐🄑🄒🄓🄔🄕🄖🄗🄘🄙🄚🄛🄜🄝🄞🄟🄠🄡🄢🄣🄤🄥🄦🄧🄨🄩]$', c):
            wid += 2.0
        elif re.match('^[ⒶⒷⒸⒹⒺⒻⒼⒽⒾⒿⓀⓁⓂⓃⓄⓅⓆⓇⓈⓉⓊⓋⓌⓍⓎⓏ]$', c):
            wid += 2.0
        elif re.match('^[㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿]$', c):
            wid += 2.0
        elif re.match('^[🄋➀➁➂➃➄➅➆➇➈➉]$', c):
            wid += 2.0
        elif re.match('^[㋐㋑㋒㋓㋔㋕㋖㋗㋘㋙㋚㋛㋜㋝㋞㋟㋠㋡㋢㋣㋤㋥㋦㋧㋨]$', c):
            wid += 2.0
        elif re.match('^[㋩㋪㋫㋬㋭㋮㋯㋰㋱㋲㋳㋴㋵㋶㋷㋸㋹㋺㋻㋼㋽㋾]$', c):
            wid += 2.0
        elif re.match('^㊀㊁㊂㊃㊄㊅㊆㊇㊈㊉$', c):
            wid += 2.0
        elif (w == 'F'):  # Full alphabet ...
            wid += 2.0
        elif(w == 'H'):   # Half katakana ...
            wid += 1.0
        elif(w == 'W'):   # Chinese character ...
            wid += 2.0
        elif(w == 'Na'):  # Half alphabet ...
            wid += 1.0
        elif(w == 'A'):   # Greek character ...
            wid += 1.0
        elif(w == 'N'):   # Arabic character ...
            wid += 1.0
        if p != '' and p != c:
            wid += 0.5
        p = w
    return wid


def get_ideal_width(s):
    wid = 0
    for c in s:
        if (c == '\t'):
            wid = (wid + 8) // 8 * 8
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


def c2i_n_arab(s):
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


def c2i_p_arab(s):
    c = s
    if re.match('^[⑴-⒇]$', c):
        return ord(c) - 9331
    else:
        return c2i_p_arab(s)


def c2i_c_arab(s):
    c = s
    n = ord(c)
    if n == 9450:
        return n - 9450        # 0
    elif n >= 9312 and n <= 9331:
        return n - 9312 + 1    # 1-20
    elif n >= 12881 and n <= 12895:
        return n - 12881 + 21  # 21-35
    elif n >= 12977 and n <= 12991:
        return n - 12977 + 36  # 36-50
    elif n == 127243:
        return n - 127243 + 0  # 0
    elif n >= 10112 and n <= 10121:
        return n - 10112 + 1   # 1-10
    return -1


def c2i_n_kata(s):
    c = s
    if re.match('^[ｱ-ﾜ]$', c):
        return ord(c) - 65392
    elif c == 'ｦ':
        return ord(c) - 65392 + 55
    elif c == 'ﾝ':
        return ord(c) - 65392 + 1
    elif re.match('^[ア-オ]$', c):
        return int((ord(c) - 12448) / 2)
    elif re.match('^[カ-チ]$', c):
        return int((ord(c) - 12448 + 1) / 2)
    elif re.match('^[ツ-ト]$', c):
        return int((ord(c) - 12448) / 2)
    elif re.match('^[ナ-ノ]$', c):
        return int((ord(c) - 12448 - 21) / 1)
    elif re.match('^[ハ-ホ]$', c):
        return int((ord(c) - 12448 + 31) / 3)
    elif re.match('^[マ-モ]$', c):
        return int((ord(c) - 12448 - 31) / 1)
    elif re.match('^[ヤ-ヨ]$', c):
        return int((ord(c) - 12448 + 4) / 2)
    elif re.match('^[ラ-ロ]$', c):
        return int((ord(c) - 12448 - 34) / 1)
    elif re.match('^[ワヲ]$', c):
        return int((ord(c) - 12448 + 53) / 3)
    elif re.match('^[ン]$', c):
        return int((ord(c) - 12448 - 37) / 1)
    return -1


def c2i_c_kata(s):
    c = s
    return ord(c) - 13007


def c2i_n_alph(s):
    c = s
    if re.match('^[a-z]$', c):
        return ord(c) - 96
    elif re.match('^[ａ-ｚ]$', c):
        return ord(c) - 65344
    return -1


def c2i_c_alph(s):
    c = s
    return ord(c) - 9423


def c2i_c_kanj(s):
    c = s
    return ord(c) - 12927


def get_xml_value(tag_name, value_name, init_value, tag):
    if re.match('<' + tag_name + ' .+>', tag):
        res = '^.* ' + value_name + '=[\'"]([^\'"]*)[\'"].*$'
        if re.match(res, tag):
            value = re.sub(res, '\\1', tag)
            if type(init_value) is int:
                return int(value)
            if type(init_value) is float:
                return float(value)
            if type(init_value) is bool:
                if re.match('^true$', value, re.IGNORECASE):
                    return True
                else:
                    return False
            return value
    return init_value


############################################################
# CLASS


class Document:

    """A class to handle document"""

    document_title = DEFAULT_DOCUMENT_TITLE
    document_style = DEFAULT_DOCUMENT_STYLE
    paper_size = DEFAULT_PAPER_SIZE
    top_margin = DEFAULT_TOP_MARGIN
    bottom_margin = DEFAULT_BOTTOM_MARGIN
    left_margin = DEFAULT_LEFT_MARGIN
    right_margin = DEFAULT_RIGHT_MARGIN
    header_string = DEFAULT_HEADER_STRING
    page_number = DEFAULT_PAGE_NUMBER
    line_number = DEFAULT_LINE_NUMBER
    mincho_font = DEFAULT_MINCHO_FONT
    gothic_font = DEFAULT_GOTHIC_FONT
    font_size = DEFAULT_FONT_SIZE
    line_spacing = DEFAULT_LINE_SPACING
    space_before = DEFAULT_SPACE_BEFORE
    space_after = DEFAULT_SPACE_AFTER
    auto_space = DEFAULT_AUTO_SPACE
    original_file = ''

    tmpdir = ''
    media_dir = ''
    styles = []
    rels = {}

    def __init__(self):
        self.media_dir = None
        self.docx_file = None
        self.md_file = None
        self.core_raw_xml_lines = None
        self.footer1_raw_xml_lines = None
        self.footer2_raw_xml_lines = None
        self.styles_raw_xml_lines = None
        self.rels_raw_xml_lines = None
        self.document_raw_xml_lines = None
        self.raw_paragraphs = None
        self.paragraphs = None

    @staticmethod
    def make_tmpdir():
        tmpdir = tempfile.TemporaryDirectory()
        # Document.tmpdir = tmpdir
        return tmpdir

    @staticmethod
    def get_media_dir_name(md_file, docx_file):
        if md_file != '':
            if md_file == '-':
                media_dir = ''
            elif re.match('^.*\\.md$', md_file, re.I):
                media_dir = re.sub('\\.md$', '', md_file, re.I)
            else:
                media_dir = md_file + '.dir'
        else:
            if re.match('^.*\\.docx$', docx_file, re.I):
                media_dir = re.sub('\\.docx$', '', docx_file, re.I)
            else:
                media_dir = docx_file + '.dir'
        # Document.media_dir = media_dir
        return media_dir

    def extract_docx_file(self, docx_file):
        self.docx_file = docx_file
        tmpdir = Document.tmpdir.name
        try:
            shutil.unpack_archive(docx_file, tmpdir, 'zip')
        except BaseException:
            msg = '※ エラー: ' \
                + '入力ファイル「' + docx_file + '」を展開できません'
            # msg = 'error: ' \
            #     + 'not a ms word file "' + docx_file + '"'
            sys.stderr.write(msg + '\n\n')
            sys.exit(1)

    def get_raw_xml_lines(self, xml_file):
        path = self.tmpdir.name + '/' + xml_file
        if not os.path.exists(path):
            return []
        try:
            xf = open(path, 'r', encoding='utf-8')
        except BaseException:
            msg = '※ エラー: ' \
                + 'XMLファイル「' + xml_file + '」を読み込めません'
            # msg = 'error: ' \
            #     + 'can\'t read "' + xml_file + '"'
            sys.stderr.write(msg + '\n\n')
            sys.exit(1)
        tmp = ''
        for ln in xf:
            ln = re.sub('\n', '', ln)
            ln = re.sub('\r', '', ln)
            tmp += ln
        tmp = re.sub('<', '\n<', tmp)
        tmp = re.sub('>', '>\n', tmp)
        tmp = re.sub('\n+', '\n', tmp)
        raw_xml_lines = tmp.split('\n')
        return raw_xml_lines

    def configure(self, args):
        # PAPER SIZE, MARGIN, LINE NUMBER, DOCUMENT STYLE
        self._configure_by_document_xml(self.document_raw_xml_lines)
        # DOCUMENT TITLE, DOCUMENT STYLE, ORIGINAL FILE
        self._configure_by_core_xml(self.core_raw_xml_lines)
        # HEADER STRING
        self._configure_by_headerX_xml(self.header1_raw_xml_lines)
        self._configure_by_headerX_xml(self.header2_raw_xml_lines)
        # PAGE NUMBER
        self._configure_by_footerX_xml(self.footer1_raw_xml_lines)
        self._configure_by_footerX_xml(self.footer2_raw_xml_lines)
        # FONT, LINE SPACING, AUTOSPACE, SAPCE BEFORE AND AFTER
        self._configure_by_styles_xml(self.styles_raw_xml_lines)
        # REVISE
        self._configure_by_args(args)
        # PARAGRAPH
        Paragraph.mincho_font = self.mincho_font
        Paragraph.gothic_font = self.gothic_font
        Paragraph.font_size = self.font_size

    def _configure_by_document_xml(self, raw_xml_lines):
        width_x = -1.0
        height_x = -1.0
        top_x = -1.0
        bottom_x = -1.0
        left_x = -1.0
        right_x = -1.0
        for rxl in raw_xml_lines:
            width_x = get_xml_value('w:pgSz', 'w:w', width_x, rxl)
            height_x = get_xml_value('w:pgSz', 'w:h', height_x, rxl)
            top_x = get_xml_value('w:pgMar', 'w:top', top_x, rxl)
            bottom_x = get_xml_value('w:pgMar', 'w:bottom', bottom_x, rxl)
            left_x = get_xml_value('w:pgMar', 'w:left', left_x, rxl)
            right_x = get_xml_value('w:pgMar', 'w:right', right_x, rxl)
            # LINE NUMBER
            if re.match('^<w:lnNumType( .*)?>$', rxl):
                self.line_number = True
        # PAPER SIZE
        width = width_x / 567
        height = height_x / 567
        if 41.9 <= width and width <= 42.1:
            if 29.6 <= height and height <= 29.8:
                self.paper_size = 'A3'
        if 29.6 <= width and width <= 29.8:
            if 41.9 <= height and height <= 42.1:
                self.paper_size = 'A3P'
        if 20.9 <= width and width <= 21.1:
            if 29.6 <= height and height <= 29.8:
                self.paper_size = 'A4'
        if 29.6 <= width and width <= 29.8:
            if 20.9 <= height and height <= 21.1:
                self.paper_size = 'A4L'
        # MARGIN
        if top_x > 0:
            self.top_margin = round(top_x / 567, 1)
        if bottom_x > 0:
            self.bottom_margin = round(bottom_x / 567, 1)
        if left_x > 0:
            self.left_margin = round(left_x / 567, 1)
        if right_x > 0:
            self.right_margin = round(right_x / 567, 1)
        # DOCUMENT STYLE
        xml_body = self._get_xml_body('w:body', raw_xml_lines)
        xml_blocks = self._get_xml_blocks(xml_body)
        par_text = []
        for xb in xml_blocks:
            plain_text = ''
            for xl in xb:
                if not re.match('^<.*>$', xl):
                    plain_text += xl
            par_text.append(plain_text)
        has_a1 = False
        has_p1 = False
        for t in par_text:
            if re.match('^第(1|１)+条\\s.*$', t):
                has_a1 = True
            if re.match('^(1|１)\\s.*$', t):
                has_p1 = True
        if has_a1:
            if has_p1:
                self.document_style = 'k'
            else:
                self.document_style = 'j'

    def _configure_by_core_xml(self, raw_xml_lines):
        for i, rxl in enumerate(raw_xml_lines):
            # DOCUMUNT TITLE
            resb = '^<dc:title>$'
            rese = '^</dc:title>$'
            if i > 0 and re.match(resb, raw_xml_lines[i - 1], re.I):
                if not re.match(rese, rxl, re.I):
                    self.document_title = rxl
            # DOCUMENT STYLE
            resb = '^<cp:category>$'
            rese = '^</cp:category>$'
            if i > 0 and re.match(resb, raw_xml_lines[i - 1], re.I):
                if not re.match(rese, rxl, re.I):
                    if re.match('^.*（普通）.*$', rxl):
                        self.document_style = 'n'
                    elif re.match('^.*（契約）.*$', rxl):
                        self.document_style = 'k'
                    elif re.match('^.*（条文）.*$', rxl):
                        self.document_style = 'j'
            # ORIGINAL FILE
            resb = '^<dcterms:modified( .*)?>$'
            rese = '^</dcterms:modified>$'
            if i > 0 and re.match(resb, raw_xml_lines[i - 1], re.I):
                if not re.match(rese, rxl, re.I):
                    dt = datetime.datetime.strptime(rxl, '%Y-%m-%dT%H:%M:%S%z')
                    if dt.tzname() == 'UTC':
                        dt += datetime.timedelta(hours=9)
                        jst = datetime.timezone(datetime.timedelta(hours=9))
                        dt = dt.replace(tzinfo=jst)
                    self.original_file = dt.strftime('%Y-%m-%dT%H:%M:%S+09:00')

    def _configure_by_headerX_xml(self, raw_xml_lines):
        # HEADER STRING
        hs = ''
        is_in_paragraph = False
        alg = 'L'
        for rxl in raw_xml_lines:
            if re.match('^<w:p( .*)?>$', rxl):
                is_in_paragraph = True
                continue
            elif re.match('^</w:p>$', rxl):
                is_in_paragraph = False
                continue
            if not is_in_paragraph:
                continue
            if re.match('<w:jc( .*)w:val=[\'"]center[\'"]( .*)?/>', rxl):
                alg = 'C'
            elif re.match('<w:jc( .*)w:val=[\'"]right[\'"]( .*)?/>', rxl):
                alg = 'R'
            elif re.match('^<.*>$', rxl):
                continue
            elif re.match('^PAGE( .*)?', rxl):
                hs += 'n'
            elif re.match('^NUMPAGES( .*)?', rxl):
                hs += 'N'
            else:
                hs += rxl
        if hs != '':
            hs = re.sub('n-\\s[0-9]+\\s-', '- n -', hs)
            hs = re.sub('N-\\s[0-9]+\\s-', '- N -', hs)
            hs = re.sub('n[0-9A-Za-z]+', 'n', hs)
            hs = re.sub('N[0-9A-Za-z]+', 'N', hs)
            if alg == 'L':
                hs = ': ' + hs
            elif alg == 'R':
                hs = hs + ' :'
            self.header_string = hs

    def _configure_by_footerX_xml(self, raw_xml_lines):
        # PAGE NUMBER
        pn = ''
        is_in_paragraph = False
        alg = 'L'
        for rxl in raw_xml_lines:
            if re.match('^<w:p( .*)?>$', rxl):
                is_in_paragraph = True
                continue
            elif re.match('^</w:p>$', rxl):
                is_in_paragraph = False
                continue
            if not is_in_paragraph:
                continue
            if re.match('<w:jc( .*)w:val=[\'"]center[\'"]( .*)?/>', rxl):
                alg = 'C'
            elif re.match('<w:jc( .*)w:val=[\'"]right[\'"]( .*)?/>', rxl):
                alg = 'R'
            elif re.match('^<.*>$', rxl):
                continue
            elif re.match('^PAGE( .*)?', rxl):
                pn += 'n'
            elif re.match('^NUMPAGES( .*)?', rxl):
                pn += 'N'
            else:
                pn += rxl
        if pn != '':
            pn = re.sub('n-\\s[0-9]+\\s-', '- n -', pn)
            pn = re.sub('N-\\s[0-9]+\\s-', '- N -', pn)
            pn = re.sub('n[0-9A-Za-z]+', 'n', pn)
            pn = re.sub('N[0-9A-Za-z]+', 'N', pn)
            if alg == 'L':
                pn = ': ' + pn
            elif alg == 'R':
                pn = pn + ' :'
            self.page_number = pn

    def _configure_by_styles_xml(self, raw_xml_lines):
        xml_body = self._get_xml_body('w:styles', raw_xml_lines)
        xml_blocks = self._get_xml_blocks(xml_body)
        sb = ['0.0', '0.0', '0.0', '0.0', '0.0', '0.0']
        sa = ['0.0', '0.0', '0.0', '0.0', '0.0', '0.0']
        for xb in xml_blocks:
            name = ''
            font = ''
            sz_x = -1.0
            ls_x = -1.0
            ase = -1
            asn = -1
            for xl in xb:
                name = get_xml_value('w:name', 'w:val', name, xl)
                font = get_xml_value('w:rFonts', '*', font, xl)
                sz_x = get_xml_value('w:sz', 'w:val', sz_x, xl)
                ls_x = get_xml_value('w:spacing', 'w:line', ls_x, xl)
                ase = get_xml_value('w:autoSpaceDE', 'w:val', ase, xl)
                asn = get_xml_value('w:autoSpaceDN', 'w:val', asn, xl)
            if name == 'makdo':
                # MINCHO FONT
                self.mincho_font = font
                # FONT SIZE
                if sz_x > 0:
                    self.font_size = round(sz_x / 2, 1)
                # LINE SPACING
                if ls_x > 0:
                    self.line_spacing = round(ls_x / 20 / self.font_size, 2)
                # AUTOSPACE
                if ase == 0 and asn == 0:
                    self.auto_space = False
                else:
                    self.auto_space = True
            elif name == 'makdo-g':
                # GOTHIC FONT
                self.gothic_font = font
            else:
                for i in range(6):
                    if name != 'makdo-' + str(i + 1):
                        continue
                    for xl in xb:
                        sb[i] \
                            = get_xml_value('w:spacing', 'w:before', sb[i], xl)
                        sa[i] \
                            = get_xml_value('w:spacing', 'w:after', sa[i], xl)
                    if sb[i] != '':
                        f = float(sb[i])
                        f = f / 20 / self.font_size / self.line_spacing
                        sb[i] = str(round(f, 2))
                    if sa[i] != '':
                        f = float(sa[i])
                        f = f / 20 / self.font_size / self.line_spacing
                        sa[i] = str(round(f, 2))
        # SPACE BEFORE, SPACE AFTER
        csb = ',' + ', '.join(sb) + ','
        # csb = re.sub(',0\\.0,', ',,', csb)
        # csb = re.sub('\\.0,', ',', csb)
        csb = re.sub('^,', '', csb)
        csb = re.sub(',$', '', csb)
        csa = ',' + ', '.join(sa) + ','
        # csa = re.sub(',0\\.0,', ',,', csa)
        # csa = re.sub('\\.0,', ',', csa)
        csa = re.sub('^,', '', csa)
        csa = re.sub(',$', '', csa)
        if csb != '':
            Document.space_before = csb
        if csa != '':
            Document.space_after = csa

    @staticmethod
    def _configure_by_args(args):
        if args.document_title is not None:
            Document.document_title = args.document_title
        if args.document_style is not None:
            Document.document_style = args.document_style
        if args.paper_size is not None:
            Document.paper_size = args.paper_size
        if args.top_margin is not None:
            Document.top_margin = args.top_margin
        if args.bottom_margin is not None:
            Document.bottom_margin = args.bottom_margin
        if args.left_margin is not None:
            Document.left_margin = args.left_margin
        if args.right_margin is not None:
            Document.right_margin = args.right_margin
        if args.header_string is not None:
            Document.header_string = args.header_string
        if args.page_number is not None:
            Document.page_number = args.page_number
        if args.line_number:
            Document.line_number = True
        if args.mincho_font is not None:
            Document.mincho_font = args.mincho_font
        if args.gothic_font is not None:
            Document.gothic_font = args.gothic_font
        if args.font_size is not None:
            Document.font_size = args.font_size
        if args.line_spacing is not None:
            Document.line_spacing = args.line_spacing
        if args.space_before is not None:
            Document.space_before = args.space_before
        if args.space_after is not None:
            Document.space_after = args.space_after
        if args.auto_space is not None:
            Document.auto_space = args.auto_space

    def get_styles(self, raw_xml_lines):
        styles = []
        xml_body = self._get_xml_body('w:styles', raw_xml_lines)
        xml_blocks = self._get_xml_blocks(xml_body)
        for n, xb in enumerate(xml_blocks):
            s = Style(n + 1, xb)
            styles.append(s)
        # self.styles = styles
        return styles

    def get_rels(self, raw_xml_lines):
        rels = {}
        res = '^<Relationship Id=[\'"](.*)[\'"] .* Target=[\'"](.*)[\'"]/>$'
        for rxl in raw_xml_lines:
            if re.match(res, rxl):
                rel_id = re.sub(res, '\\1', rxl)
                rel_tg = re.sub(res, '\\2', rxl)
                rels[rel_id] = rel_tg
        # self.rels = rels
        return rels

    def get_raw_paragraphs(self, raw_xml_lines):
        raw_paragraphs = []
        xml_body = self._get_xml_body('w:body', raw_xml_lines)
        xml_blocks = self._get_xml_blocks(xml_body)
        for xb in xml_blocks:
            rp = RawParagraph(xb)
            raw_paragraphs.append(rp)
        # self.raw_paragraphs = raw_paragraphs
        return raw_paragraphs

    def get_paragraphs(self, raw_paragraphs):
        paragraphs = []
        for rp in raw_paragraphs:
            if rp.paragraph_class == 'configuration':
                continue
            p = rp.get_paragraph()
            paragraphs.append(p)
        # self.paragraphs = paragraphs
        return paragraphs

    @staticmethod
    def _get_xml_body(tag_name, raw_xml_lines):
        xml_body = []
        is_in_body = False
        for rxl in raw_xml_lines:
            if re.match('^</?' + tag_name + '( .*)?>$', rxl):
                is_in_body = not is_in_body
                continue
            if is_in_body:
                xml_body.append(rxl)
        return xml_body

    @staticmethod
    def _get_xml_blocks(xml_body):
        xml_blocks = []
        xml_class = None
        res_oneline_tag = '<(\\S+)( .*)?/>'
        res_beginning_tag = '<(\\S+)( .*)?>'
        for xl in xml_body:
            if xml_class == '':
                if not re.match(res_beginning_tag, xl):
                    xb.append(xl)
                    continue
                else:
                    xml_blocks.append(xb)
                    xml_class = None
            if xml_class is None:
                xb = []
                xb.append(xl)
                if re.match(res_oneline_tag, xl):
                    xml_blocks.append(xb)
                elif re.match(res_beginning_tag, xl):
                    xml_class = re.sub(res_beginning_tag, '\\1', xl)
                else:
                    xml_class = ''
            else:
                xb.append(xl)
                res_end_tag = '</' + xml_class + '>'
                if re.match(res_end_tag, xl):
                    xml_blocks.append(xb)
                    xml_class = None
        return xml_blocks

    def modify_paragraphs(self):
        # ORDER IS IMPORTANT
        self.paragraphs = self._modpar_left_alignment()
        self.paragraphs = self._modpar_blank_paragraph_to_space_before()
        self.paragraphs = self._modpar_section_space_before_and_after()
        self.paragraphs = self._modpar_one_line_paragraph()
        self.paragraphs = self._modpar_spaced_and_centered()
        return self.paragraphs

    def _modpar_left_alignment(self):
        for i, p in enumerate(self.paragraphs):
            if p.paragraph_class == 'sentence':
                if p.length_docx['first indent'] == 0:
                    if p.length_docx['left indent'] == 0:
                        p.paragraph_class = 'alignment'
                        p.alignment = 'left'
        return self.paragraphs

    def _modpar_blank_paragraph_to_space_before(self):
        m = len(self.paragraphs) - 1
        for i, p in enumerate(self.paragraphs):
            if i < m:
                p_next = self.paragraphs[i + 1]
            if p.paragraph_class == 'blank':
                v_line = p.md_text.count('\n') + 1.0
                p.md_text = ''
                p.length_supp['space before'] += v_line
                # RENEW
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines = p._get_md_lines(p.md_text)
                p.text_to_write = p.get_text_to_write()
                p.paragraph_class = 'empty'
            if p.paragraph_class == 'empty' and i < m:
                if i == m:
                    continue
                lg_sb = p.length_revi['space before']
                lg_sa = p.length_revi['space after']
                lg_nx = p_next.length_revi['space before']
                p.length_supp['space before'] -= lg_sb
                p.length_supp['space after'] -= lg_sa
                if lg_sa < lg_nx:
                    p_next.length_supp['space before'] += lg_sb
                else:
                    p_next.length_supp['space before'] = lg_sa + lg_sb
                # RENEW
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines = p._get_md_lines(p.md_text)
                p.text_to_write = p.get_text_to_write()
                p_next.length_revi = p_next._get_length_revi()
                p_next.length_revisers \
                    = p_next._get_length_revisers(p_next.length_revi)
                # p_next.md_lines = p_next._get_md_lines(p_next.md_text)
                p_next.text_to_write = p_next.get_text_to_write()
        return self.paragraphs

    def _modpar_spaced_and_centered(self):
        # self.paragraphs = self._modpar_blank_paragraph_to_space_before()
        m = len(self.paragraphs) - 1
        for i, p in enumerate(self.paragraphs):
            if p.paragraph_class == 'alignment':
                if p.alignment == 'center':
                    if p.length_revi['space before'] == 1.0:
                        Paragraph.previous_head_section_depth = 1
                        Paragraph.previous_tail_section_depth = 1
                        p.pre_text_to_write = 'v=+1.0\n# \n\n'
                        p.length_supp['space before'] -= 1.0
            p.head_section_depth, p.tail_section_depth \
                = p._get_section_depths(p.raw_text)
            p.length_dept = p._get_length_dept()
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            p.md_lines = p._get_md_lines(p.md_text)
            p.text_to_write = p.get_text_to_write()
        return self.paragraphs

    def _modpar_one_line_paragraph(self):
        paper_size = Document.paper_size
        left_margin = Document.left_margin
        right_margin = Document.right_margin
        font_size = Document.font_size
        for p in self.paragraphs:
            if p.paragraph_class == 'table' or p.paragraph_class == 'image':
                indent = p.length_revi['first indent'] \
                    + p.length_revi['left indent']
                if indent == 0:
                    p.length_supp['first indent'] \
                        -= p.length_revi['first indent']
                    p.length_supp['left indent'] \
                        -= p.length_revi['left indent']
                    # RENEW
                    p.length_revi = p._get_length_revi()
                    p.length_revisers = p._get_length_revisers(p.length_revi)
                    p.md_lines = p._get_md_lines(p.md_text)
                    p.text_to_write = p.get_text_to_write()
                continue
            rt = p.raw_text
            for fd in FONT_DECORATORS:
                res = NOT_ESCAPED + fd
                while re.match(res, rt):
                    rt = re.sub(res, '\\1', rt)
            while re.match(NOT_ESCAPED + '\\\\', rt):
                rt = re.sub(NOT_ESCAPED + '\\\\', '\\1', rt)
            unit = 12 * 2.54 / 72 / 2
            line_width_cm = float(get_real_width(rt)) * unit
            indent = p.length_docx['first indent'] \
                + p.length_docx['left indent'] \
                + p.length_docx['right indent']
            region_width_cm = PAPER_WIDTH[paper_size] \
                - left_margin - right_margin \
                - (indent * unit)
            if line_width_cm > region_width_cm:
                continue
            indent \
                = p.length_revi['first indent'] + p.length_revi['left indent']
            if indent != 0:
                continue
            p.length_supp['first indent'] -= p.length_revi['first indent']
            p.length_supp['left indent'] -= p.length_revi['left indent']
            # RENEW
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            p.md_lines = p._get_md_lines(p.md_text)
            p.text_to_write = p.get_text_to_write()
        return self.paragraphs

    def _modpar_section_space_before_and_after(self):
        m = len(self.paragraphs) - 1
        for i, p in enumerate(self.paragraphs):
            if i > 0:
                p_prev = self.paragraphs[i - 1]
            if i < m:
                p_next = self.paragraphs[i + 1]
            # TITLE
            if p.paragraph_class == 'section' and \
               ParagraphSection._get_section_depths(p.raw_text) == (1, 1):
                # BEFORE
                if i > 0:
                    if p_prev.length_docx['space after'] >= 0.2:
                        p_prev.length_docx['space after'] -= 0.1
                    elif p_prev.length_docx['space after'] >= 0.0:
                        p_prev.length_docx['space after'] /= 2
                if True:
                    if p.length_docx['space before'] >= 0.2:
                        p.length_docx['space before'] -= 0.1
                    elif p.length_docx['space before'] >= 0.0:
                        p.length_docx['space before'] /= 2
                # AFTER
                if True:
                    if p.length_docx['space after'] >= 0.1:
                        p.length_docx['space after'] += 0.1
                    elif p.length_docx['space after'] >= 0.0:
                        p.length_docx['space after'] *= 2
                if i < m:
                    if p_next.length_docx['space before'] >= 0.1:
                        p_next.length_docx['space before'] += 0.1
                    elif p_next.length_docx['space before'] >= 0.0:
                        p_next.length_docx['space before'] *= 2
            # TABLE
            elif p.paragraph_class == 'table':
                if i > 0:
                    p.length_docx['space before'] \
                        = p_prev.length_docx['space after']
                    p_prev.length_docx['space after'] = 0.0
                if i < m:
                    p.length_docx['space after'] \
                        = p_next.length_docx['space before']
                    p_next.length_docx['space before'] = 0.0
            else:
                continue
            # RENEW
            if i > 0:
                p_prev.length_revi = p_prev._get_length_revi()
                p_prev.length_revisers \
                    = p_prev._get_length_revisers(p_prev.length_revi)
                p_prev.md_lines = p_prev._get_md_lines(p_prev.md_text)
                p_prev.text_to_write = p_prev.get_text_to_write()
            if True:
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                p.md_lines = p._get_md_lines(p.md_text)
                p.text_to_write = p.get_text_to_write()
            if i < m:
                p_next.length_revi = p_next._get_length_revi()
                p_next.length_revisers \
                    = p_next._get_length_revisers(p_next.length_revi)
                p_next.md_lines = p_next._get_md_lines(p_next.md_text)
                p_next.text_to_write = p_next.get_text_to_write()
        return self.paragraphs

    def open_md_file(self, md_file, docx_file):
        self.mdi_file = md_file
        if md_file == '-':
            mf = sys.stdout
        else:
            if md_file == '':
                if re.match('^.*\\.docx$', docx_file):
                    md_file = re.sub('\\.docx$', '.md', docx_file)
                else:
                    md_file = docx_file + '.md'
            if os.path.exists(md_file):
                if not os.access(md_file, os.W_OK):
                    msg = '※ エラー: ' \
                        + '出力ファイル「' + md_file + '」に書き込み権限が' \
                        + 'ありません'
                    # msg = 'error: ' \
                    #     + 'overwriting a unwritable file "' + md_file + '"'
                    sys.stderr.write(msg + '\n\n')
                    sys.exit(1)
                if os.path.getmtime(docx_file) < os.path.getmtime(md_file):
                    msg = '※ エラー: ' \
                        + '出力ファイル「' + md_file + '」の方が' \
                        + '入力ファイル「' + docx_file + '」よりも新しいです'
                    # msg = 'error: ' \
                    #     + 'overwriting a newer file "' + md_file + '"'
                    sys.stderr.write(msg + '\n\n')
                    sys.exit(1)
                if os.path.exists(md_file + '~'):
                    os.remove(md_file + '~')
                os.rename(md_file, md_file + '~')
            try:
                mf = open(md_file, 'w', encoding='utf-8', newline='\n')
            except BaseException:
                msg = '※ エラー: ' \
                    + '出力ファイル「' + md_file + '」の書き込みに失敗しました'
                # msg = 'error: ' \
                #     + 'can\'t write "' + md_file + '"'
                sys.stderr.write(msg + '\n\n')
                sys.exit(1)
        return mf

    def write_configurations(self, mf):
        mf.write(
            '<!---------------------------【設定】----------------------------'
            + '\n')
        mf.write('\n')
        # self._write_configurations_in_english(mf)
        self._write_configurations_in_japanese(mf)
        mf.write(
            '---------------------------------------------------------------->'
            + '\n')
        mf.write('\n')
        return

    def _write_configurations_in_english(self, mf):
        mf.write('document_title: '
                 + self.document_title + '\n')
        mf.write('document_style: '
                 + self.document_style + '\n')
        mf.write('paper_size:     '
                 + str(self.paper_size) + '\n')
        mf.write('top_margin:     '
                 + str(round(self.top_margin, 1)) + '\n')
        mf.write('bottom_margin:  '
                 + str(round(self.bottom_margin, 1)) + '\n')
        mf.write('left_margin:    '
                 + str(round(self.left_margin, 1)) + '\n')
        mf.write('right_margin:   '
                 + str(round(self.right_margin, 1)) + '\n')
        mf.write('header_string:  '
                 + str(self.header_string) + '\n')
        mf.write('page_number:    '
                 + str(self.page_number) + '\n')
        mf.write('line_number:    '
                 + str(self.line_number) + '\n')
        mf.write('mincho_font:    '
                 + self.mincho_font + '\n')
        mf.write('gothic_font:    '
                 + self.gothic_font + '\n')
        mf.write('font_size:      '
                 + str(round(self.font_size, 1)) + '\n')
        mf.write('line_spacing:   '
                 + str(round(self.line_spacing, 2)) + '\n')
        mf.write('space_before:   '
                 + self.space_before + '\n')
        mf.write('space_after:    '
                 + self.space_after + '\n')
        mf.write('auto_space:     '
                 + str(self.auto_space) + '\n')
        mf.write('original_file:  '
                 + self.original_file + '\n')

    def _write_configurations_in_japanese(self, mf):

        mf.write(
            '# プロパティに表示される書面のタイトルを指定ください。'
            + '\n')
        if self.document_title != '':
            mf.write('書題名: ' + self.document_title + '\n')
        else:
            mf.write('書題名: -\n')
        mf.write('\n')

        mf.write(
            '# 3つの書式（普通、契約、条文）を指定できます。'
            + '\n')
        if self.document_style == 'k':
            mf.write('文書式: 契約\n')
        elif self.document_style == 'j':
            mf.write('文書式: 条文\n')
        else:
            mf.write('文書式: 普通\n')
        mf.write('\n')

        mf.write(
            '# 用紙のサイズ（A3横、A3縦、A4横、A4縦）を指定できます。'
            + '\n')
        if self.paper_size == 'A3L' or self.paper_size == 'A3':
            mf.write('用紙サ: A3横\n')
        elif self.paper_size == 'A3P':
            mf.write('用紙サ: A3縦\n')
        elif self.paper_size == 'A4L':
            mf.write('用紙サ: A4横\n')
        else:
            mf.write('用紙サ: A4縦\n')
        mf.write('\n')

        mf.write(
            '# 用紙の上下左右の余白をセンチメートル単位で指定できます。'
            + '\n')
        mf.write('上余白: ' + str(round(self.top_margin, 1)) + ' cm\n')
        mf.write('下余白: ' + str(round(self.bottom_margin, 1)) + ' cm\n')
        mf.write('左余白: ' + str(round(self.left_margin, 1)) + ' cm\n')
        mf.write('右余白: ' + str(round(self.right_margin, 1)) + ' cm\n')
        mf.write('\n')

        mf.write(
            '# ページのヘッダーに表示する文字列（別紙 :等）を指定できます。'
            + '\n')
        mf.write('頭書き: ' + self.header_string + '\n')
        mf.write('\n')

        mf.write(
            '# ページ番号の書式（無、有、n :、-n-、n/N等）を指定できます。'
            + '\n')
        if self.page_number == '':
            mf.write('頁番号: 無\n')
        elif self.page_number == DEFAULT_PAGE_NUMBER:
            mf.write('頁番号: 有\n')
        else:
            mf.write('頁番号: ' + self.page_number + '\n')
        mf.write('\n')

        mf.write(
            '# 行番号の記載（無、有）を指定できます。'
            + '\n')
        if self.line_number:
            mf.write('行番号: 有\n')
        else:
            mf.write('行番号: 無\n')
        mf.write('\n')

        mf.write(
            '# 明朝体とゴシック体のフォントを指定できます。'
            + '\n')
        mf.write('明朝体: ' + self.mincho_font + '\n')
        mf.write('ゴシ体: ' + self.gothic_font + '\n')
        mf.write('\n')

        mf.write(
            '# 基本の文字の大きさをポイント単位で指定できます。'
            + '\n')
        mf.write('文字サ: ' + str(round(self.font_size, 1)) + ' pt\n')
        mf.write('\n')

        mf.write(
            '# 行間の高さを基本の文字の高さの何倍にするかを指定できます。'
            + '\n')
        mf.write('行間高: ' + str(round(self.line_spacing, 2)) + ' 倍\n')
        mf.write('\n')

        mf.write(
            '# セクションタイトル前後の余白を行間の高さの倍数で指定できます。'
            + '\n')
        mf.write('前余白: ' + re.sub(',', ' 倍,', self.space_before) + ' 倍\n')
        mf.write('後余白: ' + re.sub(',', ' 倍,', self.space_after) + ' 倍\n')
        mf.write('\n')

        mf.write(
            '# 半角文字と全角文字の間の間隔調整（無、有）を指定できます。'
            + '\n')
        if self.auto_space:
            mf.write('字間整: 有\n')
        else:
            mf.write('字間整: 無\n')
        mf.write('\n')

        mf.write(
            '# 変換元のWordファイルの最終更新日時が自動で指定されます。'
            + '\n')
        mf.write('元原稿: ' + self.original_file + '\n')
        mf.write('\n')

    def write_document(self, mf):
        ps = self.paragraphs
        for i, p in enumerate(ps):
            p.write_paragraph(mf)
            if p.paragraph_class != 'empty':
                mf.write('\n')

    def make_media_dir(self, media_dir):
        paragraphs = self.paragraphs
        if media_dir == '':
            return
        for p in paragraphs:
            if len(p.images) == 0:
                continue
            if os.path.exists(media_dir):
                if not os.path.isdir(media_dir):
                    msg = '※ 警告: ' \
                        + '画像の保存先「' + media_dir + '」' \
                        + 'と同名のファイルが存在します'
                    # msg = 'warning: ' \
                    #     + 'non-directory "' + media_dir + '"'
                    sys.stderr.write(msg + '\n\n')
                    return
            else:
                try:
                    os.mkdir(media_dir)
                except BaseException:
                    msg = '※ 警告: ' \
                        + '画像の保存先「' + media_dir + '」' \
                        + 'を作成できません'
                    # msg = 'warning: ' \
                    #     + 'cannot make "' + media_dir + '"'
                    sys.stderr.write(msg + '\n\n')
                    return
            p.save_images()
        return


class Style:

    """A class to handle styles"""

    def __init__(self, number, raw_xml_lines):
        self.number = number
        self.raw_xml_lines = raw_xml_lines
        self.type = None
        self.styleid = None
        self.name = None
        self.font = None
        self.font_size = None
        self.alignment = None
        self.raw_length \
            = {'space before': None, 'space after': None, 'line spacing': None,
               'first indent': None, 'left indent': None, 'right indent': None}
        self.substitute_values()

    def substitute_values(self):
        type = None
        stid = None
        name = None
        font = None
        fs_x = None
        alig = None
        rl = {'sb': None, 'sa': None, 'ls': None,
              'fi': None, 'hi': None, 'li': None, 'ri': None}
        for rxl in self.raw_xml_lines:
            type = get_xml_value('w:style', 'w:type', type, rxl)
            stid = get_xml_value('w:style', 'w:styleId', stid, rxl)
            name = get_xml_value('w:name', 'w:val', name, rxl)
            font = get_xml_value('w:rFonts', '*', font, rxl)
            fs_x = get_xml_value('w:sz', 'w:val', fs_x, rxl)
            alig = get_xml_value('w:jc', 'w:val', alig, rxl)
            rl['sb'] = get_xml_value('w:spacing', 'w:before', rl['sb'], rxl)
            rl['sa'] = get_xml_value('w:spacing', 'w:after', rl['sa'], rxl)
            rl['ls'] = get_xml_value('w:spacing', 'w:line', rl['ls'], rxl)
            rl['ls'] = get_xml_value('w:spacing', 'w:line', rl['ls'], rxl)
            rl['fi'] = get_xml_value('w:ind', 'w:firstLine', rl['fi'], rxl)
            rl['hi'] = get_xml_value('w:ind', 'w:hanging', rl['hi'], rxl)
            rl['li'] = get_xml_value('w:ind', 'w:left', rl['li'], rxl)
            rl['ri'] = get_xml_value('w:ind', 'w:right', rl['ri'], rxl)
        self.type = type
        self.styleid = stid
        self.name = name
        self.font = font
        if fs_x is not None:
            self.font_size = round(float(fs_x) / 2, 1)
        self.alignment = alig
        if rl['sb'] is not None:
            self.raw_length['space before'] = float(rl['sb'])
        if rl['sa'] is not None:
            self.raw_length['space after'] = float(rl['sa'])
        if rl['ls'] is not None:
            self.raw_length['line spacing'] = float(rl['ls'])
        if (rl['fi'] is not None) or (rl['hi'] is not None):
            self.raw_length['first indent'] = 0.0
            if rl['fi'] is not None:
                self.raw_length['first indent'] += float(rl['fi'])
            if rl['hi'] is not None:
                self.raw_length['first indent'] -= float(rl['hi'])
        if rl['li'] is not None:
            self.raw_length['left indent'] = float(rl['li'])
        if rl['ri'] is not None:
            self.raw_length['right indent'] = float(rl['ri'])


class RawParagraph:

    """A class to handle raw paragraph"""

    raw_paragraph_number = 0

    def __init__(self, raw_xml_lines):
        # DECLARATION
        self.raw_paragraph_number = -1
        self.raw_xml_lines = []
        self.raw_class = ''
        self.xml_lines = []
        self.raw_text = ''
        self.images = {}
        self.beg_space = ''
        self.end_space = ''
        self.style = ''
        self.alignment = ''
        self.paragraph_class = ''
        # SUBSTITUTION
        RawParagraph.raw_paragraph_number += 1
        self.raw_paragraph_number = RawParagraph.raw_paragraph_number
        self.raw_xml_lines = raw_xml_lines
        self.raw_class = self._get_raw_class(self.raw_xml_lines)
        self.xml_lines \
            = self._get_xml_lines(self.raw_class, self.raw_xml_lines)
        self.raw_text, self.images \
            = self._get_raw_text_and_images(self.xml_lines)
        self.beg_space, self.raw_text, self.end_space \
            = self._separate_space(self.raw_text)
        self.style = self._get_style(raw_xml_lines)
        self.alignment = self._get_alignment(self.raw_xml_lines)
        self.paragraph_class = self._get_paragraph_class()

    @staticmethod
    def _get_raw_class(raw_xml_lines):
        res = '^<(\\S+)( .*)?>$'
        rxlz = raw_xml_lines[0]
        if re.match(res, rxlz):
            return re.sub(res, '\\1', rxlz)
        else:
            return None

    @staticmethod
    def _get_xml_lines(raw_class, raw_xml_lines):
        size = Paragraph.font_size
        s_size = 0.9 * size  # not 0.8
        xml_lines = []
        is_large = False
        is_small = False
        is_italic = False
        is_bold = False
        is_gothic = False
        has_strike = False
        has_underline = False
        font_color = ''
        highlight_color = ''
        has_deleted = False   # TRACK CHANGES
        has_inserted = False  # TRACK CHANGES
        is_in_text = False
        res_img_ms \
            = '^<v:imagedata r:id=[\'"](.+)[\'"] o:title=[\'"](.+)[\'"]/>$'
        res_img_py_name \
            = '^<pic:cNvPr id=[\'"](.+)[\'"] name=[\'"](.+)[\'"]/>$'
        res_img_py_id \
            = '^<a:blip r:embed=[\'"](.+)[\'"]/>$'
        img_size = 'medium'
        res_img_size \
            = '<wp:extent cx=[\'"]([0-9]+)[\'"] cy=[\'"]([0-9]+)[\'"]/>'
        for rxl in raw_xml_lines:
            if re.match(res_img_size, rxl):
                # IMAGE SIZE
                sz_w = re.sub(res_img_size, '\\1', rxl)
                sz_h = re.sub(res_img_size, '\\2', rxl)
                cm_w = round(float(sz_w) / 12700, 1)
                cm_h = round(float(sz_h) / 12700, 1)
                if cm_h > size * 1.1:
                    img_size = 'large'
                elif cm_h < size * 0.9:
                    img_size = 'small'
                else:
                    img_size = 'medium'
            if re.match(res_img_ms, rxl):
                # IMAGE MS WORD
                if img_size == 'small':
                    xml_lines.append('--')
                    xml_lines.append(rxl)
                    xml_lines.append('--')
                elif img_size == 'large':
                    xml_lines.append('++')
                    xml_lines.append(rxl)
                    xml_lines.append('++')
                else:
                    xml_lines.append(rxl)
                img_size = 'medium'
                continue
            if re.match(res_img_py_name, rxl) or re.match(res_img_py_id, rxl):
                # IMAGE PYTHON-DOCX
                if img_size == 'small':
                    xml_lines.append('--')
                    xml_lines.append(rxl)
                    xml_lines.append('--')
                elif img_size == 'large':
                    xml_lines.append('++')
                    xml_lines.append(rxl)
                    xml_lines.append('++')
                else:
                    xml_lines.append(rxl)
                img_size = 'medium'
                continue
            if re.match('^<w:r( .*)?>$', rxl):
                text = ''
                xml_lines.append(rxl)
                is_in_text = True
                continue
            # TRACK CHANGES
            if re.match('^<w:ins( .*)?>$', rxl):
                has_inserted = True
                continue
            elif re.match('^</w:ins( .*)?>$', rxl):
                has_inserted = False
                continue
            if re.match('^</w:r>$', rxl):
                # ITALIC
                if is_italic:
                    text = '*' + text + '*'
                    is_italic = False
                # BOLD
                if is_bold:
                    text = '**' + text + '**'
                    is_bold = False
                # STRIKETHROUGH
                if has_strike:
                    text = '~~' + text + '~~'
                    has_strike = False
                # PREFORMATTED
                if is_gothic:
                    text = '`' + text + '`'
                    is_gothic = False
                # SMALL
                if is_small:
                    text = '--' + text + '--'
                    is_small = False
                # LARGE
                if is_large:
                    text = '++' + text + '++'
                    is_large = False
                # FONT COLOR
                if font_color != '':
                    if font_color == 'FFFFFF':
                        text = '^^' + text + '^^'
                    elif font_color in FONT_COLOR:
                        text = '^' + FONT_COLOR[font_color] + '^' \
                            + text \
                            + '^' + FONT_COLOR[font_color] + '^'
                    else:
                        text = '^' + font_color + '^' \
                            + text \
                            + '^' + font_color + '^'
                    font_color = ''
                # UNDERLINE
                if has_underline:
                    text = '__' + text + '__'
                    has_underline = False
                # HIGILIGHT COLOR
                if highlight_color != '':
                    text = '_' + highlight_color + '_' \
                        + text \
                        + '_' + highlight_color + '_'
                    highlight_color = ''
                # TRACK CHANGES (DELETED)
                if has_deleted:
                    text = '&lt;!--' + text + '--&gt;'
                    has_deleted = False
                # TRACK CHANGES (INSERTED)
                elif has_inserted:
                    text = '&lt;!+&gt;' + text + '&lt;+&gt;'
                xml_lines.append(text)
                text = ''
                is_in_text = False
                continue
            if not is_in_text:
                xml_lines.append(rxl)
                continue
            s = get_xml_value('w:sz', 'w:val', -1.0, rxl) / 2
            w = get_xml_value('w:w', 'w:val', -1.0, rxl)
            if s > 0:
                if not RawParagraph._is_table(raw_class, raw_xml_lines):
                    if s > size:
                        is_large = True
                    if s < size:
                        is_small = True
                else:
                    if s > s_size:
                        is_large = True
            elif w > 0:
                if w > 100:
                    is_large = True
                if w < 100:
                    is_small = True
            elif re.match('^<w:i/?>$', rxl):
                is_italic = True
            elif re.match('^<w:b/?>$', rxl):
                is_bold = True
            elif re.match('^<w:rFonts .*((Gothic)|(ゴシック)).*>$', rxl):
                is_gothic = True
            elif re.match('^<w:strike/?>$', rxl):
                has_strike = True
            elif re.match('^<w:u( .*)?>$', rxl):
                has_underline = True
            elif re.match('^<w:color w:val="[0-9A-F]+"( .*)?/?>$', rxl):
                font_color \
                    = re.sub('^<.*w:val="([0-9A-F]+)".*>$', '\\1', rxl, re.I)
                font_color = font_color.upper()
            elif re.match('^<w:highlight w:val="[a-zA-Z]+"( .*)?/?>$', rxl):
                highlight_color \
                    = re.sub('^<.*w:val="([a-zA-Z]+)".*>$', '\\1', rxl)
            elif re.match('^<w:br/?>$', rxl):
                text += '\n'
            # TRACK CHANGES
            elif re.match('^<w:delText( .*)?>$', rxl):
                has_deleted = True
            elif not re.match('^<.*>$', rxl):
                rxl = rxl.replace('\\', '\\\\')
                rxl = rxl.replace('*', '\\*')
                rxl = rxl.replace('`', '\\`')
                rxl = rxl.replace('~~', '\\~\\~')
                rxl = rxl.replace('__', '\\_\\_')
                rxl = rxl.replace('//', '\\/\\/')
                # http https ftp ...
                rxl = re.sub('([a-z]+:)\\\\/\\\\/', '\\1//', rxl)
                rxl = rxl.replace('++', '\\+\\+')
                rxl = rxl.replace('--', '\\-\\-')
                rxl = rxl.replace('%%', '\\%\\%')
                rxl = rxl.replace('&lt;', '\\&lt;')
                rxl = rxl.replace('&gt;', '\\&gt;')
                text += rxl
        # self.xml_lines = xml_lines
        return xml_lines

    @staticmethod
    def _is_table(raw_class, raw_xml_lines):
        if raw_class != 'w:tbl':
            return False
        tbl_type = ''
        col = 0
        for rxl in raw_xml_lines:
            if re.match('<w:tblStyle w:val=[\'"].+[\'"]/>', rxl):
                return True
            if re.match('<w:gridCol w:w=[\'"][0-9]+[\'"]/>', rxl):
                col += 1
        if col != 3:
            return True
        return False

    @staticmethod
    def _get_raw_text_and_images(xml_lines):
        media_dir = Document.media_dir
        img_rels = Document.rels
        raw_text = ''
        images = {}
        res_img_ms \
            = '^<v:imagedata r:id=[\'"](.+)[\'"] o:title=[\'"](.+)[\'"]/>$'
        res_img_py_name \
            = '^<pic:cNvPr id=[\'"](.+)[\'"] name=[\'"](.+)[\'"]/>$'
        res_img_py_id \
            = '^<a:blip r:embed=[\'"](.+)[\'"]/>$'
        for xl in xml_lines:
            if re.match(res_img_ms, xl):
                img_id = re.sub(res_img_ms, '\\1', xl)
                img_name = re.sub(res_img_ms, '\\2', xl)
                img_rel_name = img_rels[img_id]
                img_ext = re.sub('^.*\\.', '', img_rel_name)
                img = img_name + '.' + img_ext
                images[img_rel_name] = img
                raw_text += '![' + img + '](' + media_dir + '/' + img + ')'
            if re.match(res_img_py_name, xl):
                img = re.sub(res_img_py_name, '\\2', xl)
                images[''] = img
                raw_text += '![' + img + '](' + media_dir + '/' + img + ')'
            if re.match(res_img_py_id, xl):
                img_id = re.sub(res_img_py_id, '\\1', xl)
                img_rel_name = img_rels[img_id]
                images[img_rel_name] = images['']
            if re.match('^<.*>$', xl):
                continue
            while True:
                # ITALIC AND BOLD
                if re.match('^.*(\n.*)*[^\\*]\\*\\*\\*$', raw_text) and \
                   re.match('^\\*\\*\\*[^\\*].*$', xl):
                    raw_text = re.sub('\\*\\*\\*$', '', raw_text)
                    xl = re.sub('^\\*\\*\\*', '', xl)
                    continue
                # ITALIC
                if re.match('^.*(\n.*)*[^\\*]\\*\\*$', raw_text) and \
                   re.match('^\\*\\*[^\\*].*$', xl):
                    raw_text = re.sub('\\*\\*$', '', raw_text)
                    xl = re.sub('^\\*\\*', '', xl)
                    continue
                # BOLD
                if re.match('^.*(\n.*)*[^\\*]\\*$', raw_text) and \
                   re.match('^\\*[^\\*].*$', xl):
                    raw_text = re.sub('\\*$', '', raw_text)
                    xl = re.sub('^\\*', '', xl)
                    continue
                # STRIKETHROUGH
                if re.match('^.*(\n.*)*~~$', raw_text) and \
                   re.match('^~~.*$', xl):
                    raw_text = re.sub('~~$', '', raw_text)
                    xl = re.sub('^~~', '', xl)
                    continue
                # PREFORMATTED
                if re.match('^.*(\n.*)*`$', raw_text) and \
                   re.match('^`.*$', xl):
                    raw_text = re.sub('`$', '', raw_text)
                    xl = re.sub('^`', '', xl)
                    continue
                # SMALL
                if re.match('^.*(\n.*)*\\-\\-$', raw_text) and \
                   re.match('^\\-\\-.*$', xl):
                    raw_text = re.sub('\\-\\-$', '', raw_text)
                    xl = re.sub('^\\-\\-', '', xl)
                    continue
                # LARGE
                if re.match('^.*(\n.*)*\\+\\+$', raw_text) and \
                   re.match('^\\+\\+.*$', xl):
                    raw_text = re.sub('\\+\\+$', '', raw_text)
                    xl = re.sub('^\\+\\+', '', xl)
                    continue
                # FOND COLOR
                if re.match('^.*(\n.*)*\\^[0-9A-Za-z]*\\^$', raw_text) and \
                   re.match('^\\^[0-9A-Za-z]*\\^.*$', xl):
                    ce = re.sub('^.*(?:\n.*)*(\\^[0-9A-Za-z]*\\^)$', '\\1',
                                raw_text)
                    cb = re.sub('^(\\^[0-9A-Za-z]*\\^).*$', '\\1', xl)
                    if ce == cb:
                        raw_text = re.sub('\\^[0-9A-Za-z]*\\^$', '', raw_text)
                        xl = re.sub('^\\^[0-9A-Za-z]*\\^', '', xl)
                        continue
                # UNDERLINE
                if re.match('^.*(\n.*)*__$', raw_text) and \
                   re.match('^__.*$', xl):
                    raw_text = re.sub('__$', '', raw_text)
                    xl = re.sub('^__', '', xl)
                    continue
                # HIGILIGHT COLOR
                if re.match('^.*(\n.*)*_[0-9A-Za-z]*_$', raw_text) and \
                   re.match('^_[0-9A-Za-z]*_.*$', xl):
                    ce = re.sub('^.*(?:\n.*)*(_[0-9A-Za-z]*_)$', '\\1',
                                raw_text)
                    cb = re.sub('^(_[0-9A-Za-z]*_).*$', '\\1', xl)
                    if ce == cb:
                        raw_text = re.sub('_[0-9A-Za-z]*_$', '', raw_text)
                        xl = re.sub('^_[0-9A-Za-z]*_', '', xl)
                        continue
                # TRACK CHANGES (DELETED)
                if re.match('^.*(\n.*)*\\-\\-&gt;$', raw_text) and \
                   re.match('^&lt;!\\-\\-.*$', xl):
                    raw_text = re.sub('\\-\\-&gt;$', '', raw_text)
                    xl = re.sub('^&lt;!\\-\\-', '', xl)
                    continue
                # TRACK CHANGES (INSERTED)
                if re.match('^.*(\n.*)*&lt;\\+&gt;$', raw_text) and \
                   re.match('^&lt;!\\+&gt;.*$', xl):
                    raw_text = re.sub('&lt;\\+&gt;$', '', raw_text)
                    xl = re.sub('^&lt;!\\+&gt;', '', xl)
                    continue
                break
            raw_text += xl
        raw_text = raw_text.replace('&lt;', '<')
        raw_text = raw_text.replace('&gt;', '>')
        raw_text = raw_text.replace('&amp;', '&')
        while True:
            for fd in FONT_DECORATORS:
                res = fd + '(\\s+)' + fd
                if re.match('^.*' + res, raw_text):
                    raw_text = re.sub(res, '\\1', raw_text)
                    continue
            break
        if re.match('^\\s*(?:\\$+(?:\\-\\$)*|#+(?:\\-#)*)', raw_text):
            raw_text = '\\' + raw_text
        if re.match('^\\s*(v|V|X|<<|<|>)=\\s*[0-9]+', raw_text):
            raw_text = '\\' + raw_text
        if '' in images:
            images.pop('')
        # self.raw_text = raw_text
        # self.images = images
        return raw_text, images

    @staticmethod
    def _separate_space(raw_text):
        beg_space = ''
        end_space = ''
        res = '^([ \t\u3000]+)(.*)$'
        if re.match(res, raw_text):
            beg_space = re.sub(res, '\\1', raw_text)
            raw_text = re.sub(res, '\\2', raw_text)
        if re.match(res, raw_text[::-1]):
            end_space = re.sub(res, '\\1', raw_text[::-1])[::-1]
            raw_text = re.sub(res, '\\2', raw_text[::-1])[::-1]
        # self.raw_text = raw_text
        # self.beg_space = beg_space
        # self.end_space = end_space
        return beg_space, raw_text, end_space

    @staticmethod
    def _get_style(raw_xml_lines):
        style = None
        for rxl in raw_xml_lines:
            style = get_xml_value('w:pStyle', 'w:val', style, rxl)
        for ds in Document.styles:
            if style != ds.name:
                continue
            # COMMENTOUTED 23.02.18
            # self.alignment = ds.alignment
            # for s in self.length:
            #     if ds.raw_length[s] is not None:
            #         self.length[s] = ds.raw_length[s]
        # self.style = style
        return style

    @staticmethod
    def _get_alignment(raw_xml_lines):
        alignment = ''
        for rxl in raw_xml_lines:
            alignment = get_xml_value('w:jc', 'w:val', alignment, rxl)
            if not re.match('^(left|center|right)$', alignment):
                alignment = ''
        # self.alignment = alignment
        return alignment

    def _get_paragraph_class(self):
        if False:
            pass
        elif ParagraphEmpty.is_this_class(self):
            return 'empty'
        elif ParagraphBlank.is_this_class(self):
            return 'blank'
        elif ParagraphChapter.is_this_class(self):
            return 'chapter'
        elif ParagraphSection.is_this_class(self):
            return 'section'
        elif ParagraphSystemlist.is_this_class(self):
            return 'systemlist'
        elif ParagraphList.is_this_class(self):
            return 'list'
        elif ParagraphTable.is_this_class(self):
            return 'table'
        elif ParagraphImage.is_this_class(self):
            return 'image'
        elif ParagraphAlignment.is_this_class(self):
            return 'alignment'
        elif ParagraphPreformatted.is_this_class(self):
            return 'preformatted'
        elif ParagraphPagebreak.is_this_class(self):
            return 'pagebreak'
        elif ParagraphBreakdown.is_this_class(self):
            return 'breakdown'
        elif ParagraphConfiguration.is_this_class(self):
            return 'configuration'
        else:
            return 'sentence'

    def get_paragraph(self):
        paragraph_class = self.paragraph_class
        if False:
            pass
        elif paragraph_class == 'empty':
            return ParagraphEmpty(self)
        elif paragraph_class == 'blank':
            return ParagraphBlank(self)
        elif paragraph_class == 'chapter':
            return ParagraphChapter(self)
        elif paragraph_class == 'section':
            return ParagraphSection(self)
        elif paragraph_class == 'systemlist':
            return ParagraphSystemlist(self)
        elif paragraph_class == 'list':
            return ParagraphList(self)
        elif paragraph_class == 'table':
            return ParagraphTable(self)
        elif paragraph_class == 'image':
            return ParagraphImage(self)
        elif paragraph_class == 'alignment':
            return ParagraphAlignment(self)
        elif paragraph_class == 'preformatted':
            return ParagraphPreformatted(self)
        elif paragraph_class == 'pagebreak':
            return ParagraphPagebreak(self)
        elif paragraph_class == 'breakdown':
            return ParagraphBreakdown(self)
        else:
            return ParagraphSentence(self)


class Paragraph:

    """A class to handle paragraph"""

    paragraph_number = 0

    paragraph_class = None

    mincho_font = None
    gothic_font = None
    font_size = None

    previous_head_section_depth = 0
    previous_tail_section_depth = 0

    @classmethod
    def is_this_class(cls, raw_paragraph):
        # rp = raw_paragraph
        # rp_rxl = rp.raw_xml_lines
        # rp_rcl = rp.raw_class
        # rp_rtx = rp.raw_text
        # rp_img = rp.images
        # rp_sty = rp.style
        # rp_alg = rp.alignment
        # rp_fsz = Document.font_size
        return False

    def __init__(self, raw_paragraph):
        # RECEIVED
        self.raw_paragraph_number = raw_paragraph.raw_paragraph_number
        self.raw_xml_lines = raw_paragraph.raw_xml_lines
        self.raw_class = raw_paragraph.raw_class
        self.xml_lines = raw_paragraph.xml_lines
        self.raw_text = raw_paragraph.raw_text
        self.images = raw_paragraph.images
        self.beg_space = raw_paragraph.beg_space
        self.end_space = raw_paragraph.end_space
        self.style = raw_paragraph.style
        self.alignment = raw_paragraph.alignment
        self.paragraph_class = raw_paragraph.paragraph_class
        # DECLARATION
        self.paragraph_number = -1
        self.head_section_depth = -1
        self.tail_section_depth = -1
        self.proper_depth = -1
        self.head_font_revisers = []
        self.tail_font_revisers = []
        self.md_text = ''
        self.length_docx = {}
        self.length_dept = {}
        self.length_conf = {}
        self.length_supp = {}
        self.length_revi = {}
        self.length_revisers = []
        self.pre_text_to_write = ''
        self.post_text_to_write = ''
        # SUBSTITUTION
        Paragraph.paragraph_number += 1
        self.paragraph_number = Paragraph.paragraph_number
        self.head_section_depth, self.tail_section_depth \
            = self._get_section_depths(self.raw_text)
        self.proper_depth = self._get_proper_depth(self.raw_text)
        self.numbering_revisers, \
            self.head_font_revisers, \
            self.tail_font_revisers, \
            self.md_text \
            = self._get_revisers_and_md_text(self.raw_text)
        self.length_docx = self._get_length_docx()
        self.length_dept = self._get_length_dept()
        self.length_conf = self._get_length_conf()
        self.length_supp = self._get_length_supp()
        self.length_revi = self._get_length_revi()
        self.length_revisers = self._get_length_revisers(self.length_revi)
        # EXECUTION
        ParagraphList.reset_states(self.paragraph_class)
        self.md_lines = self._get_md_lines(self.md_text)
        self.text_to_write = self.get_text_to_write()
        # self.write_paragraph()

    @classmethod
    def _get_section_depths(cls, raw_text):
        head_section_depth = 0
        tail_section_depth = 0
        # self.head_section_depth = head_section_depth
        # self.tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    @classmethod
    def _get_proper_depth(cls, raw_text):
        proper_depth = 0
        # self.proper_depth = proper_depth
        return proper_depth

    def _get_revisers_and_md_text(self, raw_text):
        numbering_revisers = []
        head_font_revisers, tail_font_revisers, raw_text \
            = Paragraph._get_font_revisers_and_md_text(raw_text)
        md_text = self._get_md_text(raw_text)
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            md_text

    @staticmethod
    def _get_font_revisers_and_md_text(raw_text):
        head_font_revisers = []
        tail_font_revisers = []
        while True:
            for fd in FONT_DECORATORS:
                res = '^(' + fd + ')(.*(?:.*\n)*)$'
                if re.match(res, raw_text):
                    t1 = re.sub(res, '\\1', raw_text)
                    t2 = re.sub(res, '\\2', raw_text)
                    head_font_revisers.append(t1)
                    raw_text = t2
                    break
            else:
                break
        while True:
            for fd in FONT_DECORATORS:
                res = '^(.*(?:.*\n)*)(' + fd + ')$'
                if re.match(res, raw_text):
                    t1 = re.sub(res, '\\1', raw_text)
                    t2 = re.sub(res, '\\2', raw_text)
                    raw_text = t1
                    tail_font_revisers.append(t2)
                    break
            else:
                break
        return head_font_revisers, tail_font_revisers, raw_text

    @classmethod
    def _get_numbering_revisers(cls, xdepth, state):
        paragraph_class = cls.paragraph_class
        numbering_revisers = []
        for ydepth, value in enumerate(state):
            cvalue = cls.states[xdepth][ydepth]
            if value != cvalue:
                if paragraph_class == 'chapter':
                    rev = '$' * (xdepth + 1) + '-$' * ydepth + '=' + str(value)
                elif paragraph_class == 'section':
                    rev = '#' * (xdepth + 1) + '-#' * ydepth + '=' + str(value)
                elif paragraph_class == 'list':
                    rev = '  ' * xdepth + '1.=' + str(value)
                numbering_revisers.append(rev)
                cls.states[xdepth][ydepth] = value
        return numbering_revisers

    def _get_md_text(self, raw_text):
        md_text = raw_text
        return md_text

    @classmethod
    def _set_states(cls, xdepth, ydepth, value, text=''):
        paragraph_class_ja = cls.paragraph_class_ja
        paragraph_class = cls.paragraph_class
        states = cls.states
        if xdepth >= len(states):
            msg = '※ 警告: ' + paragraph_class_ja \
                + 'の深さが上限を超えています'
            # msg = 'warning: ' + paragraph_class \
            #     + ' depth exceeds limit'
            sys.stderr.write(msg + '\n\n')
            md_line.append_warning_message(msg)
        elif ydepth >= len(states[xdepth]):
            msg = '※ 警告: ' + paragraph_class_ja \
                + 'の枝が上限を超えています'
            # msg = 'warning: ' + paragraph_class \
            #     + ' branch exceeds limit'
            sys.stderr.write(msg + '\n\n')
        for x in range(len(states)):
            for y in range(len(states[x])):
                if x < xdepth:
                    continue
                elif x == xdepth:
                    if y < ydepth:
                        if states[x][y] == 0:
                            msg = '※ 警告: ' + paragraph_class_ja \
                                + 'の枝が"0"を含んでいます'
                            # msg = 'warning: ' + paragraph_class \
                            #     + ' branch has "0"'
                            sys.stderr.write(msg + '\n\n')
                    elif y == ydepth:
                        if value is None:
                            states[x][y] += 1
                        else:
                            states[x][y] = value
                    else:
                        states[x][y] = 0
                else:
                    states[x][y] = 0

    @classmethod
    def _step_states(cls, xdepth, ydepth):
        value = cls.states[xdepth][ydepth] + 1
        cls._set_states(xdepth, ydepth, value)

    def _get_length_docx(self):
        size = Document.font_size
        lnsp = Document.line_spacing
        rxls = self.raw_xml_lines
        length_docx \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        sb_xml = 0.0
        sa_xml = 0.0
        ls_xml = 0.0
        fi_xml = 0.0
        hi_xml = 0.0
        li_xml = 0.0
        ri_xml = 0.0
        ti_xml = 0.0
        for rxl in rxls:
            sb_xml = get_xml_value('w:spacing', 'w:before', sb_xml, rxl)
            sa_xml = get_xml_value('w:spacing', 'w:after', sa_xml, rxl)
            ls_xml = get_xml_value('w:spacing', 'w:line', ls_xml, rxl)
            fi_xml = get_xml_value('w:ind', 'w:firstLine', fi_xml, rxl)
            hi_xml = get_xml_value('w:ind', 'w:hanging', hi_xml, rxl)
            li_xml = get_xml_value('w:ind', 'w:left', li_xml, rxl)
            ri_xml = get_xml_value('w:ind', 'w:right', ri_xml, rxl)
            ti_xml = get_xml_value('w:tblInd', 'w:w', ti_xml, rxl)
        length_docx['space before'] = round(sb_xml / 20 / size / lnsp, 2)
        length_docx['space after'] = round(sa_xml / 20 / size / lnsp, 2)
        ls = 0.0
        if ls_xml > 0.0:
            ls = (ls_xml / 20 / size / lnsp) - 1
        length_docx['line spacing'] = round(ls, 2)
        ls75 = round(ls * .75, 2)
        ls25 = round(ls * .25, 2)
        if ls <= 0:
            if length_docx['space before'] >= ls75 * 2:
                length_docx['space before'] += ls75
            elif length_docx['space before'] >= 0:
                length_docx['space before'] /= 2
            if length_docx['space after'] >= ls25 * 2:
                length_docx['space after'] += ls25
            elif length_docx['space after'] >= 0:
                length_docx['space after'] /= 2
        else:
            if length_docx['space before'] >= ls75:
                length_docx['space before'] += ls75
            elif length_docx['space before'] >= 0:
                length_docx['space before'] *= 2
            if length_docx['space after'] >= ls25:
                length_docx['space after'] += ls25
            elif length_docx['space after'] >= 0:
                length_docx['space after'] *= 2
        length_docx['first indent'] = round((fi_xml - hi_xml) / 20 / size, 2)
        length_docx['left indent'] = round((li_xml + ti_xml) / 20 / size, 2)
        length_docx['right indent'] = round(ri_xml / 20 / size, 2)
        # self.length_docx = length_docx
        return length_docx

    def _get_length_dept(self):
        paragraph_class = self.paragraph_class
        head_section_depth = self.head_section_depth
        tail_section_depth = self.tail_section_depth
        proper_depth = self.proper_depth
        length_dept \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        if paragraph_class == 'chapter':
            length_dept['first indent'] = -1.0
            length_dept['left indent'] = proper_depth + 0.0
        elif paragraph_class == 'section':
            if head_section_depth > 1:
                length_dept['first indent'] \
                    = head_section_depth - tail_section_depth - 1.0
            if tail_section_depth > 1:
                length_dept['left indent'] = tail_section_depth - 1.0
        elif paragraph_class == 'list':
            length_dept['first indent'] = -1.0
            length_dept['left indent'] = proper_depth + 0.0
            if tail_section_depth > 0:
                length_dept['left indent'] += tail_section_depth - 1.0
        elif paragraph_class == 'preformatted':
            if tail_section_depth > 0:
                length_dept['first indent'] = 0.0
                length_dept['left indent'] = tail_section_depth - 0.0
        elif paragraph_class == 'sentence':
            if tail_section_depth > 0:
                length_dept['first indent'] = 1.0
                length_dept['left indent'] = tail_section_depth - 1.0
        if paragraph_class == 'section' or \
           paragraph_class == 'list' or \
           paragraph_class == 'preformatted' or \
           paragraph_class == 'sentence':
            if ParagraphSection.states[1][0] == 0 and tail_section_depth > 2:
                length_dept['left indent'] -= 1.0
        # self.length_dept = length_dept
        return length_dept

    def _get_length_conf(self):
        hd = self.head_section_depth
        td = self.tail_section_depth
        length_conf \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        if self.paragraph_class == 'section':
            sb = (Document.space_before + ',,,,,,,').split(',')
            sa = (Document.space_after + ',,,,,,,').split(',')
            if hd <= len(sb) and sb[hd - 1] != '':
                length_conf['space before'] += float(sb[hd - 1])
            if td <= len(sa) and sa[td - 1] != '':
                length_conf['space after'] += float(sa[td - 1])
        # self.length_conf = length_conf
        return length_conf

    def _get_length_supp(self):
        length_supp \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        # self.length_supp = length_supp
        return length_supp

    def _get_length_revi(self):
        length_docx = self.length_docx
        length_conf = self.length_conf
        length_supp = self.length_supp
        length_dept = self.length_dept
        length_revi \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        for ln in length_revi:
            lg = length_docx[ln] - length_dept[ln] \
                - length_conf[ln] + length_supp[ln]
            length_revi[ln] = round(lg, 2)
        # self.length_revi = length_revi
        return length_revi

    @staticmethod
    def _get_length_revisers(length_revi):
        length_revisers = []
        if length_revi['space before'] != 0.0:
            length_revisers.append('v=' + str(length_revi['space before']))
        if length_revi['space after'] != 0.0:
            length_revisers.append('V=' + str(length_revi['space after']))
        if length_revi['line spacing'] != 0.0:
            length_revisers.append('X=' + str(length_revi['line spacing']))
        if length_revi['first indent'] != 0.0:
            length_revisers.append('<<=' + str(-length_revi['first indent']))
        if length_revi['left indent'] != 0.0:
            length_revisers.append('<=' + str(-length_revi['left indent']))
        if length_revi['right indent'] != 0.0:
            length_revisers.append('>=' + str(-length_revi['right indent']))
        # self.length_revisers = length_revisers
        return length_revisers

    def _get_md_lines(self, md_text):
        paragraph_class = self.paragraph_class
        # FOR TRAILING WHITE SPACE
        md_text = re.sub('  \n', '  \\\n', md_text)
        if False:
            pass
        # elif paragraph_class == 'chapter':
        #     md_lines = Paragraph._split_into_lines(md_text)
        elif paragraph_class == 'section':
            md_lines = Paragraph._split_into_lines(md_text)
        # elif paragraph_class == 'list':
        #     md_lines = Paragraph._split_into_lines(md_text)
        elif paragraph_class == 'sentence':
            md_lines = Paragraph._split_into_lines(md_text)
        else:
            md_lines = md_text
        return md_lines

    @classmethod
    def _split_into_lines(cls, md_text):
        md_lines = ''
        for line in md_text.split('\n'):
            phrases = cls._split_into_phrases(line)
            splited = cls._concatenate_phrases(phrases)
            md_lines += splited + '<br>\n'
        md_lines = re.sub('<br>\n$', '', md_lines)
        return md_lines

    @staticmethod
    def _split_into_phrases(line):
        phrases = []
        tmp = ''
        m = len(line) - 1
        for i, c in enumerate(line):
            tmp += c
            if i == m:
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
                break
            c2 = line[i + 1]
            tmp2 = line[i + 1:]
            # + ' '
            if re.match('^ $', c2):
                continue
            # ' ' + '[^ ]'
            if re.match('^ $', c) and (not re.match('^ $', c2)):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            # '[[{(]' + '[^[{(]'
            # res = '^[\\[{\\(]$'
            # if re.match('^ $', c) and (not re.match(res, c2)):
            #     if tmp != '':
            #         phrases.append(tmp)
            #         tmp = ''
            # '[,.)}]]' + '[^,.)}] ]'
            res = '^[,\\.\\)}\\]]$'
            if re.match(res, c) and (not re.match(res, c2)) \
               and (not re.match('^ $', c2)):
                if re.match('^[,\\.]$', c) and \
                   ((i > 0) and re.match('^[0-9０-９]$', line[i - 1])) and \
                   ((i < m) and re.match('^[0-9０-９]$', line[i + 1])):
                    continue
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            # '[^『「｛（＜]' + '[『「｛（＜]'
            res = '^[『「｛（＜]$'
            if (not re.match(res, c)) and re.match(res, c2):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            # '[，、．。＞）｝」』]' + '[^，、．。＞）｝」』]'
            res = '^[，、．。＞）｝」』]$'
            if re.match(res, c) and (not re.match(res, c2)) \
               and (not re.match('^ $', c2)):
                if re.match('^[，．]$', c) and \
                   ((i > 0) and re.match('^[0-9０-９]$', line[i - 1])) and \
                   ((i < m) and re.match('^[0-9０-９]$', line[i + 1])):
                    continue
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            # '<!--' or '<!+>' (TRACK CHANGES)
            if re.match(NOT_ESCAPED + 'x$', tmp + 'x') and \
               re.match('^(<!\\-\\-|<!\\+>).*$', tmp2):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            # '-->' or '<+>' (TRACK CHANGES)
            if re.match('^.*(\\-\\->|<\\+>)$', tmp):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
        return phrases

    @staticmethod
    def _concatenate_phrases(phrases):
        tex = ''
        tmp = ''
        for p in phrases:
            res = '(?:#+(?:\\-#)* )+'
            if tex == '':
                if re.match('^' + res + '$', tmp):
                    if not re.match('^' + res + '.*$', p):
                        if re.match('^.*[.．。]$', phrases[-1]):
                            tex += tmp + '\n'
                            tmp = p
                            continue
            if get_ideal_width(tmp) <= MD_TEXT_WIDTH:
                if re.match('^.*[．。]$', tmp):
                    if tmp != '':
                        tex += tmp + '\n'
                        tmp = ''
                if re.match('^(<!\\-\\-|<!\\+>).*', p):
                    if tmp != '':
                        tex += tmp + '\n'
                        tmp = ''
                if re.match('^.*(\\-\\->|<\\+>)$', tmp):
                    if tmp != '':
                        tex += tmp + '\n'
                        tmp = ''
            if get_ideal_width(tmp + p) > MD_TEXT_WIDTH:
                if tmp != '':
                    tex += tmp + '\n'
                    tmp = ''
            tmp += p
            if get_ideal_width(tmp) <= MD_TEXT_WIDTH:
                if re.match('^.*[，、]$', tmp):
                    for c in CONJUNCTIONS:
                        if re.match('^' + c + '[，、]$', tmp):
                            tex += tmp + '\n'
                            tmp = ''
                            break
                if re.match('^.*[．。]$', tmp):
                    tex += tmp + '\n'
                    tmp = ''
            while get_ideal_width(tmp) > MD_TEXT_WIDTH:
                for i in range(len(tmp), -1, -1):
                    s1 = tmp[:i]
                    s2 = tmp[i:]
                    if get_ideal_width(s1) > MD_TEXT_WIDTH:
                        continue
                    if re.match('^.*[０-９][，．]$', s1) and \
                       re.match('^[０-９].*$', s2):
                        continue
                    if re.match('^.*を$', s1):
                        if s1 != '':
                            tex += s1 + '\n'
                            tmp = s2
                            break
                    if re.match('^.*[ぁ-ん，、．。]$', s1) and \
                       re.match('^[^ぁ-ん，、．。].*$', s2):
                        if s1 != '':
                            tex += s1 + '\n'
                            tmp = s2
                            break
                else:
                    for i in range(len(tmp), -1, -1):
                        s1 = tmp[:i]
                        s2 = tmp[i:]
                        # '\' +
                        if re.match('^.*\\\\$', s1):
                            continue
                        # + '\'
                        # if re.match('^\\\\.*$', s2):
                        #     continue
                        # '*' + '*' (BOLD)
                        if re.match('^.*\\*$', s1) and re.match('^\\*.*$', s2):
                            continue
                        # '~' + '~' (STRIKETHROUGH)
                        if re.match('^.*~$', s1) and re.match('^~.*$', s2):
                            continue
                        # '`' + '`' (PREFORMATTED)
                        if re.match('^.*`$', s1) and re.match('^`.*$', s2):
                            continue
                        # '/' + '/' (ITALIC)
                        if re.match('^.*/$', s1) and re.match('^/.*$', s2):
                            continue
                        # '-' + '-' (SMALL)
                        if re.match('^.*\\-$', s1) and re.match('^\\-.*$', s2):
                            continue
                        # '+' + '+' (LARGE)
                        if re.match('^.*\\+$', s1) and re.match('^\\+.*$', s2):
                            continue
                        # '^.*' + '.*^' (FONT COLOR)
                        if re.match('^.*\\^[0-9A-Za-z]*$', s1) and \
                           re.match('^[0-9A-Za-z]*\\^.*$', s2):
                            continue
                        # '_.*' + '.*_' (UNDERLINE AND HIGHLIGHT COLOR)
                        if re.match('^.*_[0-9A-Za-z]*$', s1) and \
                           re.match('^[0-9A-Za-z]*_.*$', s2):
                            continue
                        # ' ' + ' ' (LINE BREAK)
                        if re.match('^.* $', s1) and re.match('^ .*$', s2):
                            continue
                        # '<!' + '[-+]' or '<' + '![-+]' (TRACK CHANGES)
                        if re.match('^.*<!?$', s1) and \
                           re.match('^!?[\\-\\+].*$', s2):
                            continue
                        # '[-+]' + '>' (TRACK CHANGES)
                        if re.match('^.*[\\-\\+]$', s1) and \
                           re.match('^>.*$', s2):
                            continue
                        # '</?.*' + '.*>'
                        if re.match('^.*</?[0-9a-z]*$', s1) and \
                           re.match('^/?[0-9a-z]*>.*$', s2):
                            continue
                        if get_ideal_width(s1) <= MD_TEXT_WIDTH:
                            if s1 != '':
                                tex += s1 + '\n'
                                tmp = s2
                                break
                    else:
                        tex += tmp + '\n'
                        tmp = ''
        if tmp != '':
            tex += tmp + '\n'
            tmp = ''
        tex = re.sub('\n$', '', tex)
        return tex

    def get_text_to_write(self):
        numbering_revisers = self.numbering_revisers
        length_revisers = self.length_revisers
        head_font_revisers = self.head_font_revisers
        tail_font_revisers = self.tail_font_revisers
        md_lines = self.md_lines
        pre_text_to_write = self.pre_text_to_write
        post_text_to_write = self.post_text_to_write
        res = '^((?:.*\n)*.*) $'
        text_to_write = ''
        if pre_text_to_write != '':
            text_to_write += pre_text_to_write
        for rev in numbering_revisers:
            text_to_write += rev + ' '
        if re.match(res, text_to_write):
            text_to_write = re.sub(res, '\\1\n', text_to_write)
        for rev in length_revisers:
            text_to_write += rev + ' '
        if re.match(res, text_to_write):
            text_to_write = re.sub(res, '\\1\n', text_to_write)
        for rev in head_font_revisers:
            text_to_write += rev
        text_to_write += md_lines
        for rev in tail_font_revisers:
            text_to_write += rev
        if post_text_to_write != '':
            text_to_write += post_text_to_write
        return text_to_write

    def write_paragraph(self, mf):
        paragraph_class = self.paragraph_class
        text_to_write = self.text_to_write
        if paragraph_class != 'empty':
            if text_to_write != '':
                mf.write(text_to_write + '\n')

    def save_images(self):
        tmpdir = Document.tmpdir.name
        media_dir = Document.media_dir
        images = self.images
        for img in images:
            try:
                shutil.copy(tmpdir + '/word/' + img,
                            media_dir + '/' + images[img])
            except BaseException:
                msg = '※ 警告: ' \
                    + '画像「' + images[img] + '」' \
                    + 'を保存できません'
                # msg = 'warning: ' \
                #     + 'cannot make "' + media_dir + '"'
                sys.stderr.write(msg + '\n\n')


class ParagraphEmpty(Paragraph):

    """A class to handle empty paragraph"""

    paragraph_class = 'empty'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        return False


class ParagraphBlank(Paragraph):

    """A class to handle blank paragraph"""

    paragraph_class = 'blank'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rxl = rp.raw_xml_lines
        rp_rcl = rp.raw_class
        rp_rtx = rp.raw_text
        if ParagraphTable.is_this_class(raw_paragraph):
            return False
        if ParagraphImage.is_this_class(raw_paragraph):
            return False
        if ParagraphPagebreak.is_this_class(raw_paragraph):
            return False
        if ParagraphConfiguration.is_this_class(raw_paragraph):
            return False
        if re.match('^\\s*$', rp_rtx):
            return True
        return False


class ParagraphChapter(Paragraph):

    """A class to handle chapter paragraph"""

    paragraph_class = 'chapter'
    paragraph_class_ja = 'チャプター'

    res_branch = '(の[0-9０-９]+)*'
    unit_chars = ['編', '章', '節', '款', '目']
    res_separator = '(?:  ?|\t|\u3000)'
    res_symbols = ['(第([0-9０-９]+)' + unit_chars[0] + ')'
                   + res_branch + res_separator,
                   '(第([0-9０-９]+)' + unit_chars[1] + ')'
                   + res_branch + res_separator,
                   '(第([0-9０-９]+)' + unit_chars[2] + ')'
                   + res_branch + res_separator,
                   '(第([0-9０-９]+)' + unit_chars[3] + ')'
                   + res_branch + res_separator,
                   '(第([0-9０-９]+)' + unit_chars[4] + ')'
                   + res_branch + res_separator]
    res_rest = '(.*\\S(?:.*\n*)*)'
    states = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１編
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１章
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１節
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１款
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]  # 第１目

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rtx = rp.raw_text
        for i in range(len(cls.res_symbols)):
            res = '^' \
                + RES_FONT_DECORATORS \
                + cls.res_symbols[i] \
                + cls.res_rest + '$'
            if re.match(res, rp_rtx):
                return True
        return False

    @classmethod
    def _get_proper_depth(cls, raw_text):
        rss = cls.res_symbols
        xdepth = 0
        for i, rs in enumerate(rss):
            if re.match(rs, raw_text):
                xdepth = i
        proper_depth = xdepth + 1
        # self.proper_depth = proper_depth
        return proper_depth

    def _get_revisers_and_md_text(self, raw_text):
        rss = self.res_symbols
        rre = self.res_rest
        numbering_revisers = []
        head_font_revisers, tail_font_revisers, raw_text \
            = Paragraph._get_font_revisers_and_md_text(raw_text)
        head_symbol = ''
        for xdepth in range(len(rss)):
            res = '^' + rss[xdepth] + rre + '$'
            if re.match(res, raw_text):
                head_string, raw_text, state \
                    = self._decompose_text(res, raw_text, -1, -1)
                ydepth = len(state) - 1
                self._step_states(xdepth, ydepth)
                numbering_revisers \
                    = self._get_numbering_revisers(xdepth, state)
                head_symbol = '$' * (xdepth + 1) + '-$' * ydepth + ' '
                break
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            head_symbol + raw_text

    @staticmethod
    def _decompose_text(res, raw_text, num1, num2):
        hdstr = re.sub(res, '\\1', raw_text)
        nmsym = re.sub(res, '\\2', raw_text)
        branc = re.sub(res, '\\3', raw_text)
        rtext = re.sub(res, '\\4', raw_text)
        state = []
        for b in branc.split('の'):
            state.append(c2i_n_arab(b) - 1)
        if re.match('[0-9０-９]+', nmsym):
            state[0] = c2i_n_arab(nmsym)
        return hdstr, rtext, state


class ParagraphSection(Paragraph):

    """A class to handle section paragraph"""

    paragraph_class = 'section'
    paragraph_class_ja = 'セクション'

    # r0 = '((?:' + '|'.join(FONT_DECORATORS) + ')*)'
    r1 = '\\+\\+(.*)\\+\\+'
    r2 = '(?:(第([0-9０-９]+)条?)((?:の[0-9０-９]+)*))'
    r3 = '(?:(([0-9０-９]+))((?:の[0-9０-９]+)*))'
    r4 = '(?:([⑴-⒇]|[\\(（]([0-9０-９]+)[\\)）])((?:の[0-9０-９]+)*))'
    r5 = '(?:(([ｱ-ﾝア-ン]))((?:の[0-9０-９]+)*))'
    r6 = '(?:([(\\(（]([ｱ-ﾝア-ン])[\\)）])((?:の[0-9０-９]+)*))'
    r7 = '(?:(([a-zａ-ｚ]))((?:の[0-9０-９]+)*))'
    r8 = '(?:([(\\(（]([a-zａ-ｚ])[\\)）])((?:の[0-9０-９]+)*))'
    r9 = '(?:  ?|\t|\u3000|\\. ?|．)'
    res_symbols = [
        r1,
        r2 + '()' + r9,
        r3 + '(' + r4 + '?' + r5 + '?' + r6 + '?' + r7 + '?' + r8 + '?)' + r9,
        r3 + '?' + r4 + '(' + r5 + '?' + r6 + '?' + r7 + '?' + r8 + '?)' + r9,
        r3 + '?' + r4 + '?' + r5 + '(' + r6 + '?' + r7 + '?' + r8 + '?)' + r9,
        r3 + '?' + r4 + '?' + r5 + '?' + r6 + '(' + r7 + '?' + r8 + '?)' + r9,
        r3 + '?' + r4 + '?' + r5 + '?' + r6 + '?' + r7 + '(' + r8 + '?)' + r9,
        r3 + '?' + r4 + '?' + r5 + '?' + r6 + '?' + r7 + '?' + r8 + '()' + r9]
    res_number = '^[0-9０-９]+(?:, ?|\\. ?|，|．)[0-9０-９]+'
    res_rest = '(.*\\S(?:.*\n*)*)'
    states = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # -
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # １
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # (1)
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # ア
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # (ｱ)
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # ａ
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]  # (a)

    @classmethod
    def is_this_class(cls, raw_paragraph):
        raw_text = raw_paragraph.raw_text
        alignment = raw_paragraph.alignment
        head_section_depth, tail_section_depth \
            = cls._get_section_depths(raw_text)
        if ParagraphImage.is_this_class(raw_paragraph):
            return False
        if tail_section_depth == 1 and alignment == 'center':
            return True
        elif tail_section_depth > 1:
            return True
        return False

    @classmethod
    def _get_section_depths(cls, raw_text):
        rss = cls.res_symbols
        rfd = RES_FONT_DECORATORS
        rre = cls.res_rest
        rnm = cls.res_number
        head_section_depth = 0
        tail_section_depth = 0
        for xdepth in range(1, len(rss)):
            res = '^' + rfd + rss[xdepth] + rre + '$'
            if re.match(res, raw_text) and not re.match(rnm, raw_text):
                if head_section_depth == 0:
                    head_section_depth = xdepth + 1
                tail_section_depth = xdepth + 1
            if head_section_depth == 0 and tail_section_depth == 0:
                res = '^' + rfd + rss[0] + rfd + '$'
                if re.match(res, raw_text):
                    head_section_depth = 1
                    tail_section_depth = 1
        Paragraph.previous_head_section_depth = head_section_depth
        Paragraph.previous_tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    def _get_revisers_and_md_text(self, raw_text):
        rss = self.res_symbols
        rre = self.res_rest
        rnm = self.res_number
        numbering_revisers = []
        head_font_revisers, tail_font_revisers, raw_text \
            = Paragraph._get_font_revisers_and_md_text(raw_text)
        head_symbol = ''
        for xdepth in range(1, len(rss)):
            res = '^' + rss[xdepth] + rre + '$'
            if re.match(res, raw_text) and not re.match(rnm, raw_text):
                if xdepth == 1:
                    beg_num = 1
                    end_num = 5
                else:
                    beg_num = (3 * xdepth) - 5
                    end_num = 20
                head_string, raw_text, state \
                    = self._decompose_text(res, raw_text, beg_num, end_num)
                ydepth = len(state) - 1
                self._step_states(xdepth, ydepth)
                numbering_revisers \
                    = self._get_numbering_revisers(xdepth, state)
                head_symbol += '#' * (xdepth + 1) + '-#' * ydepth + ' '
        raw_text = re.sub('^\u3000', '', raw_text)
        if head_symbol == '':
            self._step_states(0, 0)
            head_font_revisers.remove('++')
            tail_font_revisers.remove('++')
            head_symbol = '# '
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            head_symbol + raw_text

    @staticmethod
    def _decompose_text(res, raw_text, beg_num, end_num):
        hdstr_rep = '\\' + str(beg_num) + '\\' + str(beg_num + 2)
        nmsym_rep = '\\' + str(beg_num + 1)
        branc_rep = '\\' + str(beg_num + 2)
        rtext_rep = '\\' + str(beg_num + 3) + '\u3000\\' + str(end_num)
        hdstr = re.sub(res, hdstr_rep, raw_text)
        nmsym = re.sub(res, nmsym_rep, raw_text)
        branc = re.sub(res, branc_rep, raw_text)
        rtext = re.sub(res, rtext_rep, raw_text)
        state = []
        for b in branc.split('の'):
            state.append(c2i_n_arab(b) - 1)
        if nmsym == '':
            nmsym = hdstr
        if re.match('[0-9０-９]+', nmsym):
            state[0] = c2i_n_arab(nmsym)
        elif re.match('[⑴-⒇]+', nmsym):
            state[0] = c2i_p_arab(nmsym)
        elif re.match('[ｱ-ﾝア-ン]+', nmsym):
            state[0] = c2i_n_kata(nmsym)
        elif re.match('[a-zａ-ｚ]+', nmsym):
            state[0] = c2i_n_alph(nmsym)
        return hdstr, rtext, state


class ParagraphSystemlist(Paragraph):

    """A class to handle systemlist paragraph"""

    paragraph_class = 'systemlist'

    res_xml_bullet_ms = '^<w:ilvl w:val=[\'"]([0-9]+)[\'"]/>$'
    res_xml_number_ms = '^<w:numId w:val=[\'"]([0-9]+)[\'"]/>$'
    res_xml_bullet_lo = '^<w:pStyle w:val=[\'"]ListBullet([0-9]?)[\'"]/>$'
    res_xml_number_lo = '^<w:pStyle w:val=[\'"]ListNumber([0-9]?)[\'"]/>$'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        raw_xml_lines = rp.raw_xml_lines
        res_xml_bullet_ms = cls.res_xml_bullet_ms
        res_xml_number_ms = cls.res_xml_number_ms
        res_xml_bullet_lo = cls.res_xml_bullet_lo
        res_xml_number_lo = cls.res_xml_number_lo
        for rxl in raw_xml_lines:
            if re.match(res_xml_bullet_ms, rxl):
                return True
            if re.match(res_xml_number_ms, rxl):
                return True
            if re.match(res_xml_bullet_lo, rxl):
                return True
            if re.match(res_xml_number_lo, rxl):
                return True
        return False

    @classmethod
    def _get_section_depths(cls, raw_text):
        head_section_depth = Paragraph.previous_tail_section_depth
        tail_section_depth = Paragraph.previous_tail_section_depth
        # self.head_section_depth = head_section_depth
        # self.tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    def _get_proper_depth(self, raw_text):
        res_xml_bullet_ms = self.res_xml_bullet_ms
        res_xml_number_ms = self.res_xml_number_ms
        res_xml_bullet_lo = self.res_xml_bullet_lo
        res_xml_number_lo = self.res_xml_number_lo
        raw_xml_lines = self.raw_xml_lines
        raw_text = self.raw_text
        list_type = ''
        depth = 1
        for rxl in raw_xml_lines:
            if re.match(res_xml_bullet_ms, rxl):
                n = re.sub(res_xml_bullet_ms, '\\1', rxl)
                depth = int(n) + 1
            if re.match(res_xml_number_ms, rxl):
                n = re.sub(res_xml_number_ms, '\\1', rxl)
                if n == '10':
                    list_type = 'bullet'
                else:
                    list_type = 'number'
            if re.match(res_xml_bullet_lo, rxl):
                list_type = 'bullet'
                n = re.sub(res_xml_bullet_lo, '\\1', rxl)
                if n != '':
                    depth = int(n)
            if re.match(res_xml_number_lo, rxl):
                list_type = 'number'
                n = re.sub(res_xml_number_lo, '\\1', rxl)
                if n != '':
                    depth = int(n)
        proper_depth = depth
        # self.proper_depth = proper_depth
        return proper_depth

    def _get_md_text(self, raw_text):
        res_xml_bullet_ms = self.res_xml_bullet_ms
        res_xml_number_ms = self.res_xml_number_ms
        res_xml_bullet_lo = self.res_xml_bullet_lo
        res_xml_number_lo = self.res_xml_number_lo
        raw_xml_lines = self.raw_xml_lines
        raw_text = self.raw_text
        list_type = ''
        depth = 1
        for rxl in raw_xml_lines:
            if re.match(res_xml_bullet_ms, rxl):
                n = re.sub(res_xml_bullet_ms, '\\1', rxl)
                depth = int(n) + 1
            if re.match(res_xml_number_ms, rxl):
                n = re.sub(res_xml_number_ms, '\\1', rxl)
                if n == '10':
                    list_type = 'bullet'
                else:
                    list_type = 'number'
            if re.match(res_xml_bullet_lo, rxl):
                list_type = 'bullet'
                n = re.sub(res_xml_bullet_lo, '\\1', rxl)
                if n != '':
                    depth = int(n)
            if re.match(res_xml_number_lo, rxl):
                list_type = 'number'
                n = re.sub(res_xml_number_lo, '\\1', rxl)
                if n != '':
                    depth = int(n)
        if list_type == 'bullet':
            md_text = '  ' * (depth - 1) + '- ' + raw_text
        else:
            md_text = '  ' * (depth - 1) + '1. ' + raw_text
        return md_text


class ParagraphList(Paragraph):

    """A class to handle list paragraph"""

    paragraph_class = 'list'
    paragraph_class_ja = 'リスト'

    res_separator = '(?:  ?|\t|\u3000)'
    res_symbols_b = ['((・))' + '()' + res_separator,
                     '((○))' + '()' + res_separator,
                     '((△))' + '()' + res_separator,
                     '((◇))' + '()' + res_separator]
    # res_symbols_b = ['(•)' + res_separator,  #  U+2022 Bullet
    #                  '(◦)' + res_separator,  #  U+25E6 White Bullet
    #                  '(‣)' + res_separator,  #  U+2023 Triangular Bullet
    #                  '(⁃)' + res_separator]  #  U+2043 Hyphen Bullet
    res_symbols_n = [('((' + chr(9450 + 0) + '|'
                      + '[' + chr(9311 + 1) + '-' + chr(9311 + 20) + ']|'
                      + '[' + chr(12860 + 21) + '-' + chr(12860 + 35) + ']|'
                      + '[' + chr(12941 + 36) + '-' + chr(12941 + 50) + ']|'
                      + chr(127243) + '|'
                      + '[' + chr(10111 + 1) + '-' + chr(10111 + 10) + ']))'
                      + '()' + res_separator),
                     ('(([' + chr(13007 + 1) + '-' + chr(13007 + 47) + ']))'
                      + '()' + res_separator),
                     ('(([' + chr(9423 + 1) + '-' + chr(9423 + 26) + ']))'
                      + '()' + res_separator),
                     ('(([' + chr(12927 + 1) + '-' + chr(12927 + 10) + ']))'
                      + '()' + res_separator)]
    res_rest = '(.*\\S(?:.*\n*)*)'
    states = [[0],  # ①
              [0],  # ㋐
              [0],  # ⓐ
              [0]]  # ㊀

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        raw_text = rp.raw_text
        proper_depth = cls._get_proper_depth(raw_text)
        if proper_depth > 0:
            return True
        return False

    @classmethod
    def _get_section_depths(cls, full_text):
        head_section_depth = Paragraph.previous_tail_section_depth
        tail_section_depth = Paragraph.previous_tail_section_depth
        # self.head_section_depth = head_section_depth
        # self.tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    @classmethod
    def _get_proper_depth(cls, raw_text):
        rsbs = cls.res_symbols_b
        rsns = cls.res_symbols_n
        rss = rsbs + rsns
        rfd = RES_FONT_DECORATORS
        rre = cls.res_rest
        proper_depth = 0
        for i in range(len(rss)):
            res = '^' + rfd + rss[i] + rre + '$'
            if re.match(res, raw_text):
                xdepth = i % 4
                proper_depth = xdepth + 1
                break
        return proper_depth

    def _get_revisers_and_md_text(self, raw_text):
        rsbs = self.res_symbols_b
        rsns = self.res_symbols_n
        rre = self.res_rest
        rss = rsbs + rsns
        numbering_revisers = []
        head_font_revisers, tail_font_revisers, raw_text \
            = Paragraph._get_font_revisers_and_md_text(raw_text)
        head_symbol = ''
        for i in range(len(rss)):
            res = '^' + rss[i] + rre + '$'
            if re.match(res, raw_text):
                xdepth = i % 4
                if i < 4:
                    head_string = re.sub(res, '\\1', raw_text)
                    raw_text = re.sub(res, '\\4', raw_text)
                    head_symbol = '  ' * xdepth + '- '
                else:
                    head_string, raw_text, state \
                        = self._decompose_text(res, raw_text, xdepth, -1)
                    head_symbol = '  ' * xdepth + '1. '
                    self._step_states(xdepth, 0)
                    numbering_revisers \
                        = self._get_numbering_revisers(xdepth, state)
                break
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            head_symbol + raw_text

    @staticmethod
    def _decompose_text(res, raw_text, xdepth, num):
        hdstr = re.sub(res, '\\1', raw_text)
        nmsym = re.sub(res, '\\2', raw_text)
        branc = re.sub(res, '\\3', raw_text)
        rtext = re.sub(res, '\\4', raw_text)
        if xdepth == 0:
            state = [c2i_c_arab(nmsym)]
        elif xdepth == 1:
            state = [c2i_c_kata(nmsym)]
        elif xdepth == 2:
            state = [c2i_c_alph(nmsym)]
        elif xdepth == 3:
            state = [c2i_c_kanj(nmsym)]
        else:
            state = [-1]
        return hdstr, rtext, state

    @classmethod
    def reset_states(cls, paragraph_class):
        if paragraph_class != 'list':
            for s in cls.states:
                s[0] = 0
        return


class ParagraphTable(Paragraph):

    """A class to handle table paragraph"""

    paragraph_class = 'table'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_cls = rp.raw_class
        if rp_cls == 'w:tbl':
            return True
        return False

    def _get_md_text(self, raw_text):
        s_size = 0.8 * Document.font_size
        xml_lines = self.xml_lines
        is_in_row = False
        is_in_cel = False
        tab = []
        wid = []
        for xl in xml_lines:
            res = '^<w:gridCol w:w=[\'"]([0-9]+)[\'"]/>$'
            if re.match(res, xl):
                w = round((float(re.sub(res, '\\1', xl)) / s_size / 10) - 4)
                wid.append(w)
            if is_in_cel:
                cell.append(xl)
            if re.match('<w:tr(.*)?>', xl):
                row = []
                is_in_row = True
            elif xl == '<w:tc>':
                cell = []
                is_in_cel = True
            elif xl == '</w:tc>':
                row.append(cell)
                is_in_cel = True
            elif xl == '</w:tr>':
                tab.append(row)
                is_in_row = False
        ali = []
        for row in tab:
            tmp = []
            for j, cell in enumerate(row):
                for xml in cell:
                    if re.match('<w:jc w:val=[\'"]left[\'"]/>', xml):
                        tmp.append(':' + '-' * (wid[j] - 1))
                    elif re.match('<w:jc w:val=[\'"]center[\'"]/>', xml):
                        tmp.append(':' + '-' * (wid[j] - 2) + ':')
                    elif re.match('<w:jc w:val=[\'"]right[\'"]/>', xml):
                        tmp.append('-' * (wid[j] - 1) + ':')
            ali.append(tmp)
        md_text = ''
        half_row = int(len(tab) / 2)
        is_in_head = True
        for i, row in enumerate(tab):
            if is_in_head:
                if ali[i] == ali[half_row]:
                    for cell in ali[half_row]:
                        md_text += '|' + cell + '|'
                    is_in_head = False
                    md_text += '\n'
            for cell in row:
                tmp = ''
                for lin in cell:
                    if not re.match('<.*>', lin):
                        tmp += lin
                tmp = re.sub('\n', '<br>', tmp)
                md_text += '|' + tmp + '|'
            md_text += '\n'
        md_text = md_text.replace('||', '|')
        md_text = md_text.replace('&lt;', '<')
        md_text = md_text.replace('&gt;', '>')
        md_text = re.sub('\n$', '', md_text)
        for line in md_text.split('\n'):
            if get_ideal_width(line) > MD_TEXT_WIDTH:
                # md_text = re.sub('\\|\n', '|\n\\\n', md_text)
                md_text = re.sub('\\|', '\\  |', md_text)
                md_text = re.sub('(^|\n)\\\\  \\|', '\\1|', md_text)
                md_text = re.sub('\\\\  \\|(\n|$)', '|\\1', md_text)
                md_text = re.sub('\\\\  \\|', '\\\n  |', md_text)
                md_text = re.sub('<br>', '<br>\\\n    ', md_text)
                break
        return md_text


class ParagraphImage(Paragraph):

    """A class to handle image paragraph"""

    paragraph_class = 'image'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rtx = rp.raw_text
        rp_img = rp.images
        rp_rtx = re.sub('!\\[[^\\[\\]]+\\]\\([^\\(\\)]+\\)', '', rp_rtx)
        rp_rtx = re.sub('(\\-\\-|\\+\\+)', '', rp_rtx)
        if rp_rtx == '' and len(rp_img) > 0:
            return True
        return False

    def _get_md_text(self, raw_text):
        raw_xml_lines = self.raw_xml_lines
        md_text = raw_text
        size_w = -1
        size_h = -1
        for rxl in raw_xml_lines:
            size_w = get_xml_value('wp:extent', 'cx', size_w, rxl)
            size_h = get_xml_value('wp:extent', 'cy', size_h, rxl)
        cm_w = float(size_w) / 360000
        cm_h = float(size_h) / 360000
        text_w = PAPER_WIDTH[Document.paper_size] \
            - Document.left_margin - Document.right_margin
        text_h = PAPER_HEIGHT[Document.paper_size] \
            - Document.top_margin - Document.bottom_margin
        if text_w * 0.99 < cm_w and text_w * 1.01 > cm_w:
            if text_w > text_h:
                md_text = '++' + raw_text + '++'
            else:
                md_text = '--' + raw_text + '--'
        if text_h * 0.99 < cm_h and text_h * 1.01 > cm_h:
            if text_w > text_h:
                md_text = '--' + raw_text + '--'
            else:
                md_text = '++' + raw_text + '++'
        # COMMENTOUTED 23.02.16
        # image = ''
        # res = '^<pic:cNvPr id=[\'"].+[\'"] name=[\'"](.*)[\'"]/>$'
        # for rxl in raw_xml_lines:
        #     if re.match(res, rxl):
        #         image = re.sub(res, '\\1', rxl)
        # md_text = '![' + image + '](' + image + ')'
        return md_text


class ParagraphAlignment(Paragraph):

    """A class to handle alignment paragraph"""

    paragraph_class = 'alignment'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        # rp_sty = rp.style
        rp_alg = rp.alignment
        if rp.raw_class == 'w:sectPr':
            return False  # configuration
        # if rp_sty == 'makdo-a':
        #     return True
        if rp.alignment != '':
            return True
        return False

    def _get_md_text(self, raw_text):
        alignment = self.alignment
        md_text = ''
        for ln in raw_text.split('\n'):
            if ln == '':
                continue
            if alignment == 'right':
                md_text += ln + ' :\n'
            elif alignment == 'center':
                md_text += ': ' + ln + ' :\n'
            else:
                if re.match('^.*  +$', ln):
                    md_text += ': ' + ln + '\\\n'
                else:
                    md_text += ': ' + ln + '\n'
        md_text = re.sub('\n$', '', md_text)
        return md_text


class ParagraphPreformatted(Paragraph):

    """A class to handle preformatted paragraph"""

    paragraph_class = 'preformatted'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_sty = rp.style
        if rp_sty == 'makdo-g':
            return True
        return False

    @classmethod
    def _get_section_depths(cls, full_text):
        head_section_depth = Paragraph.previous_tail_section_depth
        tail_section_depth = Paragraph.previous_tail_section_depth
        # self.head_section_depth = head_section_depth
        # self.tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    def _get_md_text(self, raw_text):
        md_text = raw_text
        md_text = re.sub('^`', '', md_text)
        md_text = re.sub('`$', '', md_text)
        res = '^(\\s+)\\[(.*)\\]'
        if re.match(res, md_text):
            md_text = re.sub(res, '\\1\\2', md_text)
        else:
            md_text = '\n' + md_text
        md_text = '``` ' + md_text + '\n```'
        return md_text


class ParagraphPagebreak(Paragraph):

    """A class to handle pagebreak paragraph"""

    paragraph_class = 'pagebreak'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rxl = rp.raw_xml_lines
        for rxl in rp_rxl:
            if re.match('^<w:br w:type=[\'"]page[\'"]/>$', rxl):
                return True
        return False

    def _get_md_text(self, raw_text):
        md_text = '<pgbr>'
        return md_text


class ParagraphBreakdown(Paragraph):

    """A class to handle breakdown paragraph"""

    paragraph_class = 'breakdown'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        return False


class ParagraphSentence(Paragraph):

    """A class to handle sentence paragraph"""

    paragraph_class = 'sentence'

    @classmethod
    def _get_section_depths(cls, full_text):
        head_section_depth = Paragraph.previous_tail_section_depth
        tail_section_depth = Paragraph.previous_tail_section_depth
        # self.head_section_depth = head_section_depth
        # self.tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth


class ParagraphConfiguration(Paragraph):

    """A class to handle configuration paragraph"""

    paragraph_class = 'configuration'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        if rp.raw_class == 'w:sectPr':
            return 'configuration'


############################################################
# MAIN


def main():

    args = get_arguments()

    doc = Document()

    Document.tmpdir = doc.make_tmpdir()

    Document.media_dir = doc.get_media_dir_name(args.md_file, args.docx_file)

    doc.extract_docx_file(args.docx_file)

    doc.core_raw_xml_lines = doc.get_raw_xml_lines('/docProps/core.xml')
    doc.header1_raw_xml_lines = doc.get_raw_xml_lines('/word/header1.xml')
    doc.header2_raw_xml_lines = doc.get_raw_xml_lines('/word/header2.xml')
    doc.footer1_raw_xml_lines = doc.get_raw_xml_lines('/word/footer1.xml')
    doc.footer2_raw_xml_lines = doc.get_raw_xml_lines('/word/footer2.xml')
    doc.styles_raw_xml_lines = doc.get_raw_xml_lines('/word/styles.xml')
    doc.rels_raw_xml_lines \
        = doc.get_raw_xml_lines('/word/_rels/document.xml.rels')
    doc.document_raw_xml_lines = doc.get_raw_xml_lines('/word/document.xml')

    doc.configure(args)

    Document.styles = doc.get_styles(doc.styles_raw_xml_lines)
    Document.rels = doc.get_rels(doc.rels_raw_xml_lines)

    doc.raw_paragraphs = doc.get_raw_paragraphs(doc.document_raw_xml_lines)
    doc.paragraphs = doc.get_paragraphs(doc.raw_paragraphs)
    doc.paragraphs = doc.modify_paragraphs()

    mf = doc.open_md_file(args.md_file, args.docx_file)
    doc.write_configurations(mf)
    doc.write_document(mf)
    mf.close()

    doc.make_media_dir(Document.media_dir)

    # print(Paragraph._split_into_lines(''))

    sys.exit(0)


if __name__ == '__main__':

    main()
