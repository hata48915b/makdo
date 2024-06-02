#!/usr/bin/python3
# Name:         docx2md.py
# Version:      v07 Furuichibashi
# Time-stamp:   <2024.06.02-12:36:16-JST>

# docx2md.py
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
# from makdo_docx2md import Docx2Md
# d2m = Docx2Md('xxx.docx')
# d2m.set_document_title('aaa')
# d2m.set_document_style('bbb')
# d2m.set_paper_size('ccc')
# d2m.set_top_margin('ddd')
# d2m.set_bottom_margin('eee')
# d2m.set_left_margin('fff')
# d2m.set_right_margin('ggg')
# d2m.set_header_string('hhh')
# d2m.set_page_number('hhh')
# d2m.set_line_number('iii')
# d2m.set_mincho_font('jjj')
# d2m.set_gothic_font('kkk')
# d2m.set_ivs_font('lll')
# d2m.set_font_size('mmm')
# d2m.set_line_spacing('nnn')
# d2m.set_space_before('ooo')
# d2m.set_space_after('ppp')
# d2m.set_auto_space('qqq')
# d2m.set_version_number('rrr')
# d2m.set_content_status('sss')
# m2d.set_has_completed('ttt')
# d2m.save('xxx.md')


############################################################
# POLICY

# document -> paragraph -> text -> string -> chars


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


__version__ = 'v07 Furuichibashi'


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
        '-d', '--document-style',
        type=str,
        choices=['k', 'j'],
        help='文書スタイルの指定（契約、条文）')
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
        metavar='FONT_NAME or ASCII_FONT_NAME/KANJI_FONT_NAME',
        help='明朝フォント')
    parser.add_argument(
        '-g', '--gothic-font',
        type=str,
        metavar='FONT_NAME or ASCII_FONT_NAME/KANJI_FONT_NAME',
        help='ゴシックフォント')
    parser.add_argument(
        '-i', '--ivs-font',
        type=str,
        metavar='FONT_NAME',
        help='異字体（IVS）フォント')
    # parser.add_argument(
    #     '--math_font',
    #     type=str,
    #     help=argparse.SUPPRESS)
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
        help='セクションタイトル前の空白')
    parser.add_argument(
        '-A', '--space-after',
        type=floats6,
        metavar='NUMBER,NUMBER,...',
        help='セクションタイトル後の空白')
    parser.add_argument(
        '-a', '--auto-space',
        action='store_true',
        help='全角文字と半角文字との間の間隔を微調整します')
    parser.add_argument(
        '--version-number',
        type=str,
        metavar='VERSION_NUMBER',
        help='バージョン番号')
    parser.add_argument(
        '--content-status',
        type=str,
        metavar='CONTENT_STATUS',
        help='書面の状態')
    parser.add_argument(
        '-c', '--has-completed',
        action='store_true',
        help='備考書（コメント）などを消して完成させます')
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
        msg = 'invalid 6 floats separated by commas value: \'' + s + '\''
        raise argparse.ArgumentTypeError(msg)
    return s


# def positive_integer(s):
#     if not re.match('[1-9][0-9]*', s):
#         msg = 'invalid positive integer value: \'' + s + '\''
#         raise argparse.ArgumentTypeError(msg)
#     return int(s)


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

DEFAULT_PAGE_NUMBER = ': n :'

DEFAULT_LINE_NUMBER = False

DEFAULT_MINCHO_FONT = 'Times New Roman / ＭＳ 明朝'
DEFAULT_GOTHIC_FONT = 'ＭＳ ゴシック'
DEFAULT_IVS_FONT = 'IPAmj明朝'  # IPAmjMincho
DEFAULT_MATH_FONT = 'Cambria Math'
# DEFAULT_MATH_FONT = 'Liberation Serif'
DEFAULT_FONT_SIZE = 12.0
TABLE_FONT_SCALE = 0.8

MS_FONTS = [
    ['ＭＳ 明朝', 'ＭＳ明朝',
     'ＭＳ 明朝;MS Mincho', 'MS Mincho;ＭＳ 明朝',
     'Mincho;MS Mincho'],
    ['ＭＳ ゴシック', 'ＭＳゴシック',
     'ＭＳ ゴシック;MS Gothic', 'MS Gothic;ＭＳ ゴシック',
     'Gothic;MS Gothic'],
    ['ＭＳ Ｐ明朝', 'ＭＳＰ明朝',
     'ＭＳ Ｐ明朝;MS PMincho', 'MS PMincho;ＭＳ Ｐ明朝'
     'PMincho;MS PMincho'],
    ['ＭＳ Ｐゴシック', 'ＭＳＰゴシック',
     'ＭＳ Ｐゴシック;MS PGothic', 'MS PGothic; ＭＳ Ｐゴシック',
     'PGothic;MS PGothic'],
    ['游明朝', 'Yu Mincho'],
    ['游ゴシック', 'Yu Gothic'],
    ['ヒラギノ明朝', 'Hiragino Mincho'],
    ['ヒラギノ角ゴ', 'Hiragino Kaku Gothic'],
    ['ヒラギノ丸ゴ', 'Hiragino Maru Gothic'],
]

DEFAULT_LINE_SPACING = 2.14  # (2.0980+2.1812)/2=2.1396

DEFAULT_SPACE_BEFORE = ''
DEFAULT_SPACE_AFTER = ''
TABLE_SPACE_BEFORE = 0.45
TABLE_SPACE_AFTER = 0.20
IMAGE_SPACE_BEFORE = 0.68
IMAGE_SPACE_AFTER = 0.00

DEFAULT_AUTO_SPACE = False

DEFAULT_VERSION_NUMBER = ''

DEFAULT_CONTENT_STATUS = ''

DEFAULT_HAS_COMPLETED = False

NOT_ESCAPED = '^((?:(?:.|\n)*?[^\\\\])?(?:\\\\\\\\)*?)?'
# NOT_ESCAPED = '^((?:(?:.|\n)*[^\\\\])?(?:\\\\\\\\)*)?'

RES_NUMBER = '(?:[-\\+]?(?:(?:[0-9]+(?:\\.[0-9]+)?)|(?:\\.[0-9]+)))'
RES_NUMBER6 = '(?:' + RES_NUMBER + '?,){,5}' + RES_NUMBER + '?,?'

RES_KATAKANA = '[' + 'ｦｱ-ﾝ' + \
    'アイウエオカキクケコサシスセソタチツテトナニヌネノ' + \
    'ハヒフヘホマミムメモヤユヨラリルレロワヰヱヲン' + ']'

RES_IMAGE = '! *\\[([^\\[\\]]*)\\] *\\(([^\\(\\)]+)\\)'
RES_IMAGE_WITH_SIZE \
    = '!' \
    + ' *' \
    + '\\[([^\\[\\]]+):(' + RES_NUMBER + ')x(' + RES_NUMBER + ')\\]' \
    + ' *' \
    + '\\(([^\\(\\)]+)\\)'

# MS OFFICE
RES_XML_IMG_MS \
    = '^<v:imagedata r:id=[\'"](.+)[\'"] o:title=[\'"](.+)[\'"]/>$'
# PYTHON-DOCX AND LIBREOFFICE
RES_XML_IMG_PY_ID \
    = '^<a:blip r:embed=[\'"](.+)[\'"]/?>$'
RES_XML_IMG_PY_NAME \
    = '^<pic:cNvPr id=[\'"](.+)[\'"] name=[\'"]([^\'"]+)[\'"](?: .*)?/?>$'
RES_XML_IMG_SIZE \
    = '^<wp:extent cx=[\'"]([0-9]+)[\'"] cy=[\'"]([0-9]+)[\'"]/>$'

FONT_DECORATORS_INVISIBLE = [
    '\\*\\*\\*',                # italic and bold
    '\\*\\*',                   # bold
    '\\*',                      # italic
    '//',                       # italic
    '\\^[0-9A-Za-z]{0,11}\\^',  # font color
]
FONT_DECORATORS_VISIBLE = [
    '\\-\\-\\-',                # xsmall
    '\\-\\-',                   # small
    '\\+\\+\\+',                # xlarge
    '\\+\\+',                   # large
    '>>>',                      # xnarrow or reset
    '>>',                       # narrow or reset
    '<<<',                      # xwide or reset
    '<<',                       # wide or reset
    '~~',                       # strikethrough
    '_[\\$=\\.#\\-~\\+]{,4}_',  # underline
    '_[0-9A-Za-z]{1,11}_',      # higilight color
    '`',                        # preformatted
    '@[^@]{1,66}@',             # font
]
FONT_DECORATORS = FONT_DECORATORS_INVISIBLE + FONT_DECORATORS_VISIBLE
RES_FONT_DECORATORS = '((?:' + '|'.join(FONT_DECORATORS) + ')*)'

RELAX_SYMBOL = '<>'

MD_TEXT_WIDTH = 68

UNDERLINE = {
    'single':          '',
    'words':           '$',
    'double':          '=',
    'dotted':          '.',
    'thick':           '#',
    'dash':            '-',
    'dotDash':         '.-',
    'dotDotDash':      '..-',
    'wave':            '~',
    'dottedHeavy':     '.#',
    'dashedHeavy':     '-#',
    'dashDotHeavy':    '.-#',
    'dashDotDotHeavy': '..-#',
    'wavyHeavy':       '~#',
    'dashLong':        '-+',
    'wavyDouble':      '~=',
    'dashLongHeavy':   '-+#',
}

FONT_COLOR = {
    'FF0000': 'red',
    # 'FF0000': 'R',
    '770000': 'darkRed',
    # '770000': 'DR',
    'FFFF00': 'yellow',
    # 'FFFF00': 'Y',
    '777700': 'darkYellow',
    # '777700': 'DY',
    '00FF00': 'green',
    # '00FF00': 'G',
    '007700': 'darkGreen',
    # '007700': 'DG',
    '00FFFF': 'cyan',
    # '00FFFF': 'C',
    '007777': 'darkCyan',
    # '007777': 'DC',
    '0000FF': 'blue',
    # '0000FF': 'B',
    '000077': 'darkBlue',
    # '000077': 'DB',
    'FF00FF': 'magenta',
    # 'FF00FF': 'M',
    '770077': 'darkMagenta',
    # '770077': 'DM',
    'BBBBBB': 'lightGray',
    # 'BBBBBB': 'G1',
    '777777': 'darkGray',
    # '777777': 'G2',
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

HIGHLIGHT_COLOR = {
    'red':         'R',
    'darkRed':     'DR',
    'yellow':      'Y',
    'darkYellow':  'DY',
    'green':       'G',
    'darkGreen':   'DG',
    'cyan':        'C',
    'darkCyan':    'DC',
    'blue':        'B',
    'darkBlue':    'DB',
    'magenta':     'M',
    'darkMagenta': 'DM',
    'lightGray':   'G1',
    'darkGray':    'G2',
    'black':       'BK',
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
    'さらにいえば',
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

UNIX_TIME = datetime.datetime.timestamp(datetime.datetime.now())


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
        elif re.match('^[㊀㊁㊂㊃㊄㊅㊆㊇㊈㊉]$', c):
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


def c2n_p_arab(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 9331
    if i >= n + 1 and i <= n + 20:
        # ⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇
        return i - n
    res = '^[\\(（]([0-9０-９]+)[\\)）]$'
    if re.match(res, s):
        # (0)...
        c = re.sub(res, '\\1', s)
        return c2n_n_arab(c)
    return -1


def c2n_c_arab(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 9450
    if i == n:
        # ⓪
        return i - n
    n = 9311
    if i >= n + 1 and i <= n + 20:
        # ①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳
        return i - n
    n = 12860
    if i >= n + 21 and i <= n + 35:
        # ㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟
        return i - n
    n = 12941
    if i >= n + 36 and i <= n + 50:
        # ㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿
        return i - n
    n = 127243
    if i == n:
        # 🄋
        return i - n
    n = 10111
    if i >= n + 1 and i <= n + 10:
        # ➀➁➂➃➄➅➆➇➈➉
        return i - n
    return -1


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


def c2n_p_kata(s):
    res = '^[\\(（](' + RES_KATAKANA + ')[\\)）]$'
    if re.match(res, s):
        # (ｱ)...(ﾝ)
        c = re.sub(res, '\\1', s)
        return c2n_n_kata(c)
    return -1


def c2n_c_kata(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 13007
    if i >= n + 1 and i <= n + 47:
        # ㋐㋑㋒㋓㋔㋕㋖㋗㋘㋙㋚㋛㋜㋝㋞㋟㋠㋡㋢㋣㋤㋥㋦㋧㋨
        # ㋩㋪㋫㋬㋭㋮㋯㋰㋱㋲㋳㋴㋵㋶㋷㋸㋹㋺㋻㋼㋽㋾
        return i - n
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


def c2n_p_alph(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 9371
    if i >= n + 1 and i <= n + 26:
        # ⒜⒝⒞⒟⒠⒡⒢⒣⒤⒥⒦⒧⒨⒩⒪⒫⒬⒭⒮⒯⒰⒱⒲⒳⒴⒵
        return i - n
    res = '^[\\(（]([a-zａ-ｚ])[\\)）]$'
    if re.match(res, s):
        # (a)...(z)
        c = re.sub(res, '\\1', s)
        return c2n_n_alph(c)
    return -1


def c2n_c_alph(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 9423
    if i >= n + 1 and i <= n + 26:
        # ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙⓚⓛⓜⓝⓞⓟⓠⓡⓢⓣⓤⓥⓦⓧⓨⓩ
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
        return i
    return -1


def c2n_p_kanj(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 12831
    if i >= n + 1 and i <= n + 10:
        # ㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩
        return i - n
    return -1


def c2n_c_kanj(s):
    i = -1
    if len(s) == 1:
        i = ord(s)
    n = 12927
    if i >= n + 1 and i <= n + 10:
        # ㊀㊁㊂㊃㊄㊅㊆㊇㊈㊉
        return i - n
    return -1


############################################################
# CLASS


class IO:

    """A class to handle input and output"""

    media_dir = ''

    def __init__(self):
        # DECLARE
        self.inputed_docx_file = None
        self.inputed_md_file = None
        self.docx_file = None
        self.md_file = None
        self.temp_dir_inst = None
        self.temp_dir = None
        self.docx_input = None
        self.md_file_inst = None
        # SUBSTITUTE
        self.temp_dir_inst = tempfile.TemporaryDirectory()
        self.temp_dir = self.temp_dir_inst.name

    def set_docx_file(self, inputed_docx_file):
        docx_file = inputed_docx_file
        if not self._verify_input_file(docx_file):
            return False
        self.inputed_docx_file = inputed_docx_file
        self.docx_file = docx_file
        return True

    def unpack_docx_file(self):
        self.docx_input = DocxFile(self.docx_file)
        docx_input = self.docx_input
        docx_input.unpack_docx_file(self.temp_dir)

    def read_xml_file(self, xml_file):
        xml_lines = self.docx_input.read_xml_file(xml_file)
        return xml_lines

    def set_md_file(self, inputed_md_file):
        inputed_docx_file = self.inputed_docx_file
        docx_file = self.docx_file
        md_file = inputed_md_file
        if md_file == '':
            if inputed_docx_file == '-':
                msg = '※ エラー: ' \
                    + '出力ファイルの指定がありません'
                # msg = 'error: ' \
                #     + 'no output file name'
                sys.stderr.write(msg + '\n\n')
                if __name__ == '__main__':
                    sys.exit(201)
                return False
            elif re.match('^.*\\.docx$', inputed_docx_file):
                md_file = re.sub('\\.docx$', '.md', inputed_docx_file)
            else:
                md_file = inputed_docx_file + '.md'
        if not self._verify_output_file(md_file):
            return False
        if not self._verify_older(docx_file, md_file):
            return False
        self.inputed_md_file = inputed_md_file
        self.md_file = md_file
        return True

    def open_md_file(self):
        self.md_file_inst = MdFile(self.md_file)
        self.md_file_inst.open()

    def write_md_file(self, article):
        self.md_file_inst.write(article)

    def close_md_file(self):
        self.md_file_inst.close()

    def save_images(self, images):
        media_dir = self.media_dir
        if len(images) == 0:
            return
        if media_dir == '':
            return
        self._make_media_dir(media_dir)
        self._copy_images(images)

    @staticmethod
    def _make_media_dir(media_dir):
        if os.path.exists(media_dir):
            if not os.path.isdir(media_dir):
                msg = '※ 警告: ' \
                    + '画像の保存先「' + media_dir + '」' \
                    + 'と同名のファイルが存在します'
                # msg = 'warning: ' \
                #     + 'non-directory "' + media_dir + '"'
                sys.stderr.write(msg + '\n\n')
                return False
        else:
            try:
                os.mkdir(media_dir)
            except BaseException:
                msg = '※ 警告: ' \
                    + '画像の保存先「' + media_dir + '」' \
                    + 'を作成できません'
                # msg = 'warning: ' \
                #     + 'can\'t make "' + media_dir + '"'
                sys.stderr.write(msg + '\n\n')
                return False

    def _copy_images(self, images):
        temp_dir = self.temp_dir
        media_dir = self.media_dir
        for img in images:
            orig_img = temp_dir + '/word/' + img
            targ_img = media_dir + '/' + images[img]
            bkup_img = targ_img + '~'
            if os.path.exists(targ_img) and os.path.exists(bkup_img):
                os.remove(bkup_img)
            if os.path.exists(targ_img) and os.path.exists(bkup_img):
                msg = '※ 警告: ' \
                    + '画像「' + images[img] + '~」' \
                    + 'を削除できません'
                # msg = 'warning: ' \
                #     + 'can\'t remove "' + images[img] + '~"'
                sys.stderr.write(msg + '\n\n')
                continue
            if os.path.exists(targ_img):
                os.rename(targ_img, bkup_img)
            if os.path.exists(targ_img):
                msg = '※ 警告: ' \
                    + '画像「' + images[img] + '」' \
                    + 'をバックアップできません'
                # msg = 'warning: ' \
                #     + 'can\'t backup "' + images[img] + '"'
                sys.stderr.write(msg + '\n\n')
                continue
            try:
                shutil.copy(orig_img, targ_img)
            except BaseException:
                msg = '※ 警告: ' \
                    + '画像「' + images[img] + '」' \
                    + 'を保存できません'
                # msg = 'warning: ' \
                #     + 'can\'t save "' + images[img] + '"'
                sys.stderr.write(msg + '\n\n')
                continue

    def get_media_dir(self):
        md_file = self.md_file
        if md_file == '':
            media_dir = ''
        else:
            if md_file == '-':
                media_dir = ''
            elif re.match('^.*\\.md$', md_file, re.I):
                media_dir = re.sub('\\.md$', '', md_file, re.I)
            else:
                media_dir = md_file + '.dir'
        # self.media_dir = media_dir
        return media_dir

    @staticmethod
    def _verify_input_file(input_file):
        if input_file == '-':
            return True
        if not os.path.exists(input_file):
            msg = '※ エラー: ' \
                + '入力ファイル「' + input_file + '」がありません'
            # msg = 'error: ' \
            #     + 'no input file "' + input_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(101)
            return False
        if not os.path.isfile(input_file):
            msg = '※ エラー: ' \
                + '入力「' + input_file + '」はファイルではありません'
            # msg = 'error: ' \
            #     + 'not a file "' + input_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(102)
            return False
        if not os.access(input_file, os.R_OK):
            msg = '※ エラー: ' \
                + '入力ファイル「' + input_file + '」に読込権限が' \
                + 'ありません'
            # msg = 'error: ' \
            #     + 'unreadable "' + input_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(103)
            return False
        return True

    @staticmethod
    def _verify_output_file(output_file):
        if output_file == '-':
            return True
        if not os.path.exists(output_file):
            return True
        if not os.path.isfile(output_file):
            msg = '※ エラー: ' \
                + '出力「' + output_file + '」はファイルではありません'
            # msg = 'error: ' \
            #     + 'not a file "' + output_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(202)
            return False
        if not os.access(output_file, os.W_OK):
            msg = '※ エラー: ' \
                + '出力ファイル「' + output_file + '」に書込権限が' \
                + 'ありません'
            # msg = 'error: ' \
            #     + 'unwritable "' + output_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(203)
            return False
        return True

    @staticmethod
    def _verify_older(input_file, output_file):
        if input_file != '-' and os.path.exists(input_file) and \
           output_file != '-' and os.path.exists(output_file):
            if os.path.getmtime(input_file) < os.path.getmtime(output_file):
                msg = '※ エラー: ' \
                    + '出力ファイルの方が入力ファイルよりも新しいです'
                # msg = 'error: ' \
                #     + 'overwriting a newer file'
                sys.stderr.write(msg + '\n\n')
                if __name__ == '__main__':
                    sys.exit(301)
                return False
        return True


class DocxFile:

    """A class to handle docx file"""

    def __init__(self, docx_file):
        # DECLARE
        self.docx_file = None
        self.temp_dir = None
        # SUBSTITUTE
        self.docx_file = docx_file

    def unpack_docx_file(self, temp_dir):
        self.temp_dir = temp_dir
        docx_file = self.docx_file
        if docx_file is None:
            return False
        try:
            shutil.unpack_archive(docx_file, temp_dir, 'zip')
        except BaseException:
            msg = '※ エラー: ' \
                + '入力ファイル「' + docx_file + '」を展開できません'
            # msg = 'error: ' \
            #     + 'can\'t unpack a input file "' + docx_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(104)
            return False
        if not os.path.exists(temp_dir + '/word/document.xml'):
            msg = '※ エラー: ' \
                + '入力ファイル「' + docx_file + '」はMS Wordのファイルでは' \
                + 'ありません'
            # msg = 'error: ' \
            #     + 'not a ms word file "' + docx_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(105)
            return False
        return True

    def read_xml_file(self, xml_file):
        path = self.temp_dir + '/' + xml_file
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
            if __name__ == '__main__':
                sys.exit(106)
            return []
        tmp = ''
        for ln in xf:
            ln = re.sub('\n', '', ln)
            ln = re.sub('\r', '', ln)
            tmp += ln
        # LIBREOFFICE
        res = '<wp:align>[a-z]+</wp:align>'
        if re.match('^.*' + res, tmp):
            tmp = re.sub(res, '', tmp)
        # LIBREOFFICE
        res = '<wp:posOffset>[0-9]+</wp:posOffset>'
        if re.match('^.*' + res, tmp):
            tmp = re.sub(res, '', tmp)
        tmp = re.sub('<', '\n<', tmp)
        tmp = re.sub('>', '>\n', tmp)
        tmp = re.sub('\n+', '\n', tmp)
        xml_lines = tmp.split('\n')
        return xml_lines


class MdFile:

    """A class to handle md file"""

    def __init__(self, md_file):
        # DECLARE
        self.md_file = None
        self.md_output = None
        # SUBSTITUTE
        self.md_file = md_file

    def open(self):
        md_file = self.md_file
        # OPEN
        if md_file == '-':
            md_output = sys.stdout
        else:
            self._save_old_file(md_file)
            try:
                md_output = open(md_file, 'w', encoding='utf-8', newline='\n')
            except BaseException:
                msg = '※ エラー: ' \
                    + '出力ファイル「' + md_file + '」の書き込みに失敗しました'
                # msg = 'error: ' \
                #     + 'can\'t write "' + md_file + '"'
                sys.stderr.write(msg + '\n\n')
                if __name__ == '__main__':
                    sys.exit(204)
                return False
        self.md_output = md_output
        return True

    def write(self, article):
        self.md_output.write(article)

    def close(self):
        self.md_output.close()

    @staticmethod
    def _save_old_file(output_file):
        if output_file == '-':
            return True
        backup_file = output_file + '~'
        if os.path.exists(output_file):
            if os.path.exists(backup_file):
                os.remove(backup_file)
            if os.path.exists(backup_file):
                msg = '※ エラー: ' \
                    + '古いファイル「' + backup_file + '」を削除できません'
                # msg = 'error: ' \
                #     + 'can\'t remove "' + backup_file + '"'
                sys.stderr.write(msg + '\n\n')
                if __name__ == '__main__':
                    sys.exit(205)
                return False
            os.rename(output_file, backup_file)
        if os.path.exists(output_file):
            msg = '※ エラー: ' \
                + '古いファイル「' + output_file + '」を改名できません'
            # msg = 'error: ' \
            #     + 'can\'t rename "' + output_file + '"'
            sys.stderr.write(msg + '\n\n')
            if __name__ == '__main__':
                sys.exit(206)
            return False
        return True


class Form:

    """A class to handle form"""

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
    ivs_font = DEFAULT_IVS_FONT
    font_size = DEFAULT_FONT_SIZE
    line_spacing = DEFAULT_LINE_SPACING
    space_before = DEFAULT_SPACE_BEFORE
    space_after = DEFAULT_SPACE_AFTER
    auto_space = DEFAULT_AUTO_SPACE
    version_number = DEFAULT_VERSION_NUMBER
    content_status = DEFAULT_CONTENT_STATUS
    has_completed = DEFAULT_HAS_COMPLETED
    original_file = ''

    styles = None
    rels = None
    remarks = None

    def __init__(self):
        # DECLARE
        self.document_xml_lines = None
        self.core_xml_lines = None
        self.styles_xml_lines = None
        self.header1_xml_lines = None
        self.header2_xml_lines = None
        self.footer1_xml_lines = None
        self.footer2_xml_lines = None
        self.rels_xml_lines = None
        self.comments_xml_lines = None
        self.args = None

    def configure(self):
        # PAPER SIZE, MARGIN, LINE NUMBER, DOCUMENT STYLE
        self._configure_by_document_xml(self.document_xml_lines)
        # DOCUMENT TITLE, DOCUMENT STYLE, VERSION NUMBER, CONTENT STATUS,
        # ORIGINAL FILE
        self._configure_by_core_xml(self.core_xml_lines)
        # FONT, LINE SPACING, AUTO SPACE, SAPCE BEFORE AND AFTER
        self._configure_by_styles_xml(self.styles_xml_lines)
        # HEADER STRING
        self._configure_by_headerX_xml(self.header1_xml_lines)
        self._configure_by_headerX_xml(self.header2_xml_lines)
        # PAGE NUMBER
        self._configure_by_footerX_xml(self.footer1_xml_lines)
        self._configure_by_footerX_xml(self.footer2_xml_lines)
        if len(self.footer1_xml_lines) == 0 and \
           len(self.footer2_xml_lines) == 0:
            Form.set_page_number('False')
        # REVISE BY ARGUMENTS
        self._configure_by_args(self.args)
        # DOCUMENT TITLE
        if Form.document_title == '':
            Form.document_title = hex(int(UNIX_TIME * 1000000))
        # FOR LIBREOFFICE (NOT SUPPORT "SECTIONPAGES")
        has_two_or_more_sections = False
        is_in_p = False
        for xl in self.document_xml_lines:
            if re.match('<w:p( .*)?>', xl):
                is_in_p = True
            if re.match('</w:p( .*)?>', xl):
                is_in_p = False
            if is_in_p and re.match('<w:sectPr( .*)?>', xl):
                has_two_or_more_sections = True
        if not has_two_or_more_sections:
            while re.match(NOT_ESCAPED + 'M', Form.page_number):
                Form.page_number \
                    = re.sub(NOT_ESCAPED + 'M', '\\1N', Form.page_number)
        elif re.match(NOT_ESCAPED + '(N|M)', Form.page_number):
            msg = '※ 警告: ' \
                + '"<Pgbr>"を含む場合、' \
                + 'Libreofficeでは総ページ番号を適切に表示できません'
            # msg = 'warning: ' \
            #     + 'If "<Pgbr>" is present, ' \
            #     + 'Libreoffice can\'t display total page numbers properly'
            sys.stderr.write(msg + '\n\n')

    def _configure_by_document_xml(self, xml_lines):
        width_x = -1.0
        height_x = -1.0
        top_x = -1.0
        bottom_x = -1.0
        left_x = -1.0
        right_x = -1.0
        for xl in xml_lines:
            width_x = XML.get_value('w:pgSz', 'w:w', width_x, xl)
            height_x = XML.get_value('w:pgSz', 'w:h', height_x, xl)
            top_x = XML.get_value('w:pgMar', 'w:top', top_x, xl)
            bottom_x = XML.get_value('w:pgMar', 'w:bottom', bottom_x, xl)
            left_x = XML.get_value('w:pgMar', 'w:left', left_x, xl)
            right_x = XML.get_value('w:pgMar', 'w:right', right_x, xl)
            # LINE NUMBER
            if re.match('^<w:lnNumType( .*)?>$', xl):
                Form.line_number = True
        # PAPER SIZE
        width = width_x / 567
        height = height_x / 567
        if 41.9 <= width and width <= 42.1:
            if 29.6 <= height and height <= 29.8:
                Form.paper_size = 'A3'
        if 29.6 <= width and width <= 29.8:
            if 41.9 <= height and height <= 42.1:
                Form.paper_size = 'A3P'
        if 20.9 <= width and width <= 21.1:
            if 29.6 <= height and height <= 29.8:
                Form.paper_size = 'A4'
        if 29.6 <= width and width <= 29.8:
            if 20.9 <= height and height <= 21.1:
                Form.paper_size = 'A4L'
        # MARGIN
        if top_x > 0:
            Form.top_margin = round(top_x / 567, 1)
        if bottom_x > 0:
            Form.bottom_margin = round(bottom_x / 567, 1)
        if left_x > 0:
            Form.left_margin = round(left_x / 567, 1)
        if right_x > 0:
            Form.right_margin = round(right_x / 567, 1)
        # DOCUMENT STYLE
        xml_body = XML.get_body('w:body', xml_lines)
        xml_blocks = XML.get_blocks(xml_body)
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
                Form.document_style = 'k'
            else:
                Form.document_style = 'j'

    def _configure_by_core_xml(self, xml_lines):
        for i, xl in enumerate(xml_lines):
            # DOCUMUNT TITLE
            resb = '^<dc:title>$'
            rese = '^</dc:title>$'
            if i > 0 and re.match(resb, xml_lines[i - 1], re.I):
                if not re.match(rese, xl, re.I):
                    Form.document_title = xl
            # DOCUMENT STYLE
            resb = '^<cp:category>$'
            rese = '^</cp:category>$'
            if i > 0 and re.match(resb, xml_lines[i - 1], re.I):
                if not re.match(rese, xl, re.I):
                    if re.match('^.*（普通）.*$', xl):
                        Form.document_style = 'n'
                    elif re.match('^.*（契約）.*$', xl):
                        Form.document_style = 'k'
                    elif re.match('^.*（条文）.*$', xl):
                        Form.document_style = 'j'
            # VERSION NUMBER
            resb = '^<cp:version>$'
            rese = '^</cp:version>$'
            if i > 0 and re.match(resb, xml_lines[i - 1], re.I):
                if not re.match(rese, xl, re.I):
                    Form.version_number = xl
            # CONTENT STATUS
            resb = '^<cp:contentStatus>$'
            rese = '^</cp:contentStatus>$'
            if i > 0 and re.match(resb, xml_lines[i - 1], re.I):
                if not re.match(rese, xl, re.I):
                    Form.content_status = xl
            # ORIGINAL FILE
            resb = '^<dcterms:modified( .*)?>$'
            rese = '^</dcterms:modified>$'
            if i > 0 and re.match(resb, xml_lines[i - 1], re.I):
                if not re.match(rese, xl, re.I):
                    dt = datetime.datetime.strptime(xl, '%Y-%m-%dT%H:%M:%S%z')
                    if dt.tzname() == 'UTC':
                        dt += datetime.timedelta(hours=9)
                        jst = datetime.timezone(datetime.timedelta(hours=9))
                        dt = dt.replace(tzinfo=jst)
                    Form.original_file \
                        = dt.strftime('%Y-%m-%dT%H:%M:%S+09:00')

    def _configure_by_styles_xml(self, xml_lines):
        xml_body = XML.get_body('w:styles', xml_lines)
        xml_blocks = XML.get_blocks(xml_body)
        sb = ['0.0', '0.0', '0.0', '0.0', '0.0', '0.0']
        sa = ['0.0', '0.0', '0.0', '0.0', '0.0', '0.0']
        for xb in xml_blocks:
            name = ''
            afnt = ''
            kfnt = ''
            sz_x = -1.0
            f_it = False
            f_bd = False
            f_sk = False
            f_ul = ''
            f_cl = ''
            f_hc = ''
            alig = ''
            ls_x = -1.0
            ase = -1
            asn = -1
            for xl in xb:
                name = XML.get_value('w:name', 'w:val', name, xl)
                afnt = XML.get_value('w:rFonts', 'w:ascii', afnt, xl)
                kfnt = XML.get_value('w:rFonts', 'w:eastAsia', kfnt, xl)
                sz_x = XML.get_value('w:sz', 'w:val', sz_x, xl)
                f_it = XML.is_this_tag('w:i', f_it, xl)
                f_bd = XML.is_this_tag('w:b', f_bd, xl)
                f_sk = XML.is_this_tag('w:strike', f_sk, xl)
                f_ul = XML.get_value('w:u', 'w:val', f_ul, xl)
                f_cl = XML.get_value('w:color', 'w:val', f_cl, xl)
                f_hc = XML.get_value('w:highlight', 'w:val', f_hc, xl)
                alig = XML.get_value('w:jc', 'w:val', alig, xl)
                ls_x = XML.get_value('w:spacing', 'w:line', ls_x, xl)
                ase = XML.get_value('w:autoSpaceDE', 'w:val', ase, xl)
                asn = XML.get_value('w:autoSpaceDN', 'w:val', asn, xl)
            if name == 'makdo':
                # MINCHO FONT
                if afnt != '' and kfnt != '':
                    if afnt == kfnt:
                        Form.mincho_font = afnt
                    else:
                        Form.mincho_font = afnt + ' / ' + kfnt
                elif afnt != '' and kfnt == '':
                    Form.mincho_font = afnt
                elif afnt == '' and kfnt != '':
                    Form.mincho_font = kfnt
                # FONT SIZE
                if sz_x > 0:
                    Form.font_size = round(sz_x / 2, 1)
                # LINE SPACING
                if ls_x > 0:
                    Form.line_spacing = round(ls_x / 20 / Form.font_size, 2)
                # AUTO SPACE
                if ase == 0 and asn == 0:
                    Form.auto_space = False
                else:
                    Form.auto_space = True
            elif name == 'makdo-g':
                # GOTHIC FONT
                if afnt != '' and kfnt != '':
                    if afnt == kfnt:
                        Form.gothic_font = afnt
                    else:
                        Form.gothic_font = afnt + ' / ' + kfnt
                elif afnt != '' and kfnt == '':
                    Form.gothic_font = afnt
                elif afnt == '' and kfnt != '':
                    Form.gothic_font = kfnt
            elif name == 'makdo-i':
                # IVS FONT
                if kfnt != '':
                    Form.ivs_font = kfnt
                elif afnt != '':
                    Form.ivs_font = afnt
            else:
                for i in range(6):
                    if name != 'makdo-' + str(i + 1):
                        continue
                    for xl in xb:
                        sb[i] \
                            = XML.get_value('w:spacing', 'w:before', sb[i], xl)
                        sa[i] \
                            = XML.get_value('w:spacing', 'w:after', sa[i], xl)
                    if sb[i] != '':
                        f = float(sb[i]) / 20 \
                            / Form.font_size / Form.line_spacing
                        sb[i] = str(round(f, 2))
                    if sa[i] != '':
                        f = float(sa[i]) / 20 \
                            / Form.font_size / Form.line_spacing
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
            Form.space_before = csb
        if csa != '':
            Form.space_after = csa

    @staticmethod
    def _configure_by_headerX_xml(xml_lines):
        # HEADER STRING
        chars_data, images \
            = RawParagraph._get_chars_data_and_images('', xml_lines, 'header')
        raw_text = RawParagraph._get_raw_text(chars_data)
        alignment = RawParagraph._get_alignment(xml_lines)
        if alignment == 'center':
            raw_text = ': ' + raw_text + ' :'
        elif alignment == 'right':
            raw_text = raw_text + ' :'
        if raw_text != '':
            Form.header_string = raw_text

    @staticmethod
    def _configure_by_footerX_xml(xml_lines):
        # PAGE NUMBER
        chars_data, images \
            = RawParagraph._get_chars_data_and_images('', xml_lines, 'footer')
        raw_text = RawParagraph._get_raw_text(chars_data)
        alignment = RawParagraph._get_alignment(xml_lines)
        if alignment == 'center':
            raw_text = ': ' + raw_text + ' :'
        elif alignment == 'right':
            raw_text = raw_text + ' :'
        if raw_text != '':
            Form.page_number = raw_text

    @staticmethod
    def _configure_by_args(args):
        if args is not None:
            if args.document_title is not None:
                Form.set_document_title(args.document_title)
            if args.document_style is not None:
                Form.set_document_style(args.document_style)
            if args.paper_size is not None:
                Form.set_paper_size(args.paper_size)
            if args.top_margin is not None:
                Form.set_top_margin(str(args.top_margin))
            if args.bottom_margin is not None:
                Form.set_bottom_margin(str(args.bottom_margin))
            if args.left_margin is not None:
                Form.set_left_margin(str(args.left_margin))
            if args.right_margin is not None:
                Form.set_right_margin(str(args.right_margin))
            if args.header_string is not None:
                Form.set_header_string(args.header_string)
            if args.page_number is not None:
                Form.set_page_number(args.page_number)
            if args.line_number:
                Form.set_line_number(str(args.line_number))
            if args.mincho_font is not None:
                Form.set_mincho_font(args.mincho_font)
            if args.gothic_font is not None:
                Form.set_gothic_font(args.gothic_font)
            if args.ivs_font is not None:
                Form.set_ivs_font(args.ivs_font)
            if args.font_size is not None:
                Form.set_font_size(str(args.font_size))
            if args.line_spacing is not None:
                Form.set_line_spacing(str(args.line_spacing))
            if args.space_before is not None:
                Form.set_space_before(args.space_before)
            if args.space_after is not None:
                Form.set_space_after(args.space_after)
            if args.auto_space:
                Form.set_auto_space(str(args.auto_space))
            if args.version_number is not None:
                Form.set_version_number(args.version_number)
            if args.content_status is not None:
                Form.set_content_status(args.content_status)
            if args.has_completed:
                Form.set_has_completed(str(args.has_completed))

    @staticmethod
    def set_document_title(value, item='document_title'):
        if value is None:
            return False
        Form.document_title = value
        return True

    @staticmethod
    def set_document_style(value, item='document_style'):
        if value is None:
            return False
        if value == 'n' or value == '普通' or value == '-':
            Form.document_style = 'n'
            return True
        if value == 'k' or value == '契約':
            Form.document_style = 'k'
            return True
        if value == 'j' or value == '条文':
            Form.document_style = 'j'
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '"普通"、"契約"又は"条文"でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + nam + '" must be "n", "k" or "j"'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_paper_size(value, item='paper_size'):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        if value == 'A3':
            Form.paper_size = 'A3'
            return True
        elif value == 'A3L' or value == 'A3横':
            Form.paper_size = 'A3L'
            return True
        elif value == 'A3P' or value == 'A3縦':
            Form.paper_size = 'A3P'
            return True
        elif value == 'A4':
            Form.paper_size = 'A4'
            return True
        elif value == 'A4L' or value == 'A4横':
            Form.paper_size = 'A4L'
            return True
        elif value == 'A4P' or value == 'A4縦':
            Form.paper_size = 'A4P'
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '"A3横"、"A3縦"、"A4横"又は"A4縦"でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be "A3", "A3P", "A4" or "A4L"'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_top_margin(value, item='top_margin'):
        return Form._set_margin(value, item)

    @staticmethod
    def set_bottom_margin(value, item='bottom_margin'):
        return Form._set_margin(value, item)

    @staticmethod
    def set_left_margin(value, item='left_margin'):
        return Form._set_margin(value, item)

    @staticmethod
    def set_right_margin(value, item='right_margin'):
        return Form._set_margin(value, item)

    @staticmethod
    def _set_margin(value, item):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        value = re.sub('\\s*cm$', '', value)
        if re.match('^' + RES_NUMBER + '$', value):
            if item == 'top_margin' or item == '上余白':
                Form.top_margin = float(value)
                return True
            if item == 'bottom_margin' or item == '下余白':
                Form.bottom_margin = float(value)
                return True
            if item == 'left_margin' or item == '左余白':
                Form.left_margin = float(value)
                return True
            if item == 'right_margin' or item == '右余白':
                Form.right_margin = float(value)
                return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '整数又は小数でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be an integer or a decimal'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_header_string(value, item='header_string'):
        if value is None:
            return False
        Form.header_string = value
        return True

    @staticmethod
    def set_page_number(value, item='page_number'):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        if value == 'True' or value == '有':
            Form.page_number = DEFAULT_PAGE_NUMBER
            return True
        elif value == 'False' or value == '無':
            Form.page_number = ''
            return True
        else:
            Form.page_number = value
            return True

    @staticmethod
    def set_line_number(value, item='line_number'):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        if value == 'True' or value == '有':
            Form.line_number = True
            return True
        elif value == 'False' or value == '無':
            Form.line_number = False
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '"有"又は"無"でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be "True" or "False"'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_mincho_font(value, item='mincho_font'):
        if value is None:
            return False
        Form.mincho_font = value
        return True

    @staticmethod
    def set_gothic_font(value, item='gothic_font'):
        if value is None:
            return False
        Form.gothic_font = value
        return True

    @staticmethod
    def set_ivs_font(value, item='ivs_font'):
        if value is None:
            return False
        Form.ivs_font = value
        return True

    @staticmethod
    def set_font_size(value, item='font_size'):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        value = re.sub('\\s*pt$', '', value)
        if re.match('^' + RES_NUMBER + '$', value):
            Form.font_size = float(value)
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '整数又は小数でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be an integer or a decimal'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_line_spacing(value, item='line_spacing'):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        value = re.sub('\\s*倍$', '', value)
        if re.match('^' + RES_NUMBER + '$', value):
            Form.line_spacing = float(value)
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '整数又は小数でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be an integer or a decimal'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_space_before(value, item='space_before'):
        return Form._set_space(value, item)

    @staticmethod
    def set_space_after(value, item='space_after'):
        return Form._set_space(value, item)

    @staticmethod
    def _set_space(value, item):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        value = value.replace('、', ',')
        value = value.replace('倍', '')
        value = value.replace(' ', '')
        if re.match('^' + RES_NUMBER6 + '$', value):
            if item == 'space_before' or item == '前余白':
                Form.space_before = value
                return True
            elif item == 'space_after' or item == '後余白':
                Form.space_after = value
                return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '整数又は小数をカンマで区切って並べたものでなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be 6 integers or decimals'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_auto_space(value, item='auto_space'):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        if value == 'True' or value == '有':
            Form.auto_space = True
            return True
        elif value == 'False' or value == '無':
            Form.auto_space = False
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '"有"又は"無"でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be "True" or "False"'
        sys.stderr.write(msg + '\n\n')

    @staticmethod
    def set_version_number(value, item='version_number'):
        if value is None:
            return False
        Form.version_number = value
        return True

    @staticmethod
    def set_content_status(value, item='content_status'):
        if value is None:
            return False
        Form.content_status = value
        return True

    @staticmethod
    def set_has_completed(value, item='has_completed'):
        if value is None:
            return False
        value = unicodedata.normalize('NFKC', value)
        if value == 'True' or value == '真偽':
            Form.has_completed = True
            return True
        elif value == 'False' or value == '偽':
            Form.has_completed = False
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '"真"又は"偽"でなければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be "True" or "False"'
        sys.stderr.write(msg + '\n\n')

    @classmethod
    def get_configurations(cls):
        return cls.get_configurations_in_japanese()
        # return cls.get_configurations_in_english()

    @classmethod
    def get_configurations_in_english(cls):
        cfgs = ''
        cfgs += \
            '<!-----------------------[CONFIGRATIONS]-------------------------'
        cfgs += '\n'
        cfgs += 'document_title: ' + cls.document_title + '\n'
        cfgs += 'document_style: ' + cls.document_style + '\n'
        cfgs += 'paper_size:     ' + str(cls.paper_size) + '\n'
        cfgs += 'top_margin:     ' + str(round(cls.top_margin, 1)) + '\n'
        cfgs += 'bottom_margin:  ' + str(round(cls.bottom_margin, 1)) + '\n'
        cfgs += 'left_margin:    ' + str(round(cls.left_margin, 1)) + '\n'
        cfgs += 'right_margin:   ' + str(round(cls.right_margin, 1)) + '\n'
        cfgs += 'header_string:  ' + str(cls.header_string) + '\n'
        cfgs += 'page_number:    ' + str(cls.page_number) + '\n'
        cfgs += 'line_number:    ' + str(cls.line_number) + '\n'
        cfgs += 'mincho_font:    ' + cls.mincho_font + '\n'
        cfgs += 'gothic_font:    ' + cls.gothic_font + '\n'
        cfgs += 'ivs_font:       ' + cls.ivs_font + '\n'
        cfgs += 'font_size:      ' + str(round(cls.font_size, 1)) + '\n'
        cfgs += 'line_spacing:   ' + str(round(cls.line_spacing, 2)) + '\n'
        cfgs += 'space_before:   ' + cls.space_before + '\n'
        cfgs += 'space_after:    ' + cls.space_after + '\n'
        cfgs += 'auto_space:     ' + str(cls.auto_space) + '\n'
        cfgs += 'version_number: ' + cls.version_number + '\n'
        cfgs += 'content_status: ' + cls.content_status + '\n'
        cfgs += 'has_completed:  ' + str(cls.has_completed) + '\n'
        cfgs += 'original_file:  ' + cls.original_file + '\n'
        cfgs += \
            '---------------------------------------------------------------->'
        cfgs += '\n'
        cfgs += '\n'
        return cfgs

    @classmethod
    def get_configurations_in_japanese(cls):
        cfgs = ''

        cfgs += \
            '<!--------------------------【設定】-----------------------------'
        cfgs += '\n\n'

        cfgs += \
            '# プロパティに表示される文書のタイトルを指定できます。'
        cfgs += '\n'
        cfgs += '書題名: ' + cls.document_title + '\n'
        cfgs += '\n'

        cfgs += \
            '# 3つの書式（普通、契約、条文）を指定できます。'
        cfgs += '\n'
        if cls.document_style == 'k':
            cfgs += '文書式: 契約\n'
        elif cls.document_style == 'j':
            cfgs += '文書式: 条文\n'
        else:
            cfgs += '文書式: 普通\n'
        cfgs += '\n'

        cfgs += \
            '# 用紙のサイズ（A3横、A3縦、A4横、A4縦）を指定できます。'
        cfgs += '\n'
        if cls.paper_size == 'A3L' or cls.paper_size == 'A3':
            cfgs += '用紙サ: A3横\n'
        elif cls.paper_size == 'A3P':
            cfgs += '用紙サ: A3縦\n'
        elif cls.paper_size == 'A4L':
            cfgs += '用紙サ: A4横\n'
        else:
            cfgs += '用紙サ: A4縦\n'
        cfgs += '\n'

        cfgs += \
            '# 用紙の上下左右の余白をセンチメートル単位で指定できます。'
        cfgs += '\n'
        cfgs += '上余白: ' + str(round(cls.top_margin, 1)) + ' cm\n'
        cfgs += '下余白: ' + str(round(cls.bottom_margin, 1)) + ' cm\n'
        cfgs += '左余白: ' + str(round(cls.left_margin, 1)) + ' cm\n'
        cfgs += '右余白: ' + str(round(cls.right_margin, 1)) + ' cm\n'
        cfgs += '\n'

        cfgs += \
            '# ページのヘッダーに表示する文字列（別紙 :等）を指定できます。'
        cfgs += '\n'
        cfgs += '頭書き: ' + cls.header_string + '\n'
        cfgs += '\n'

        cfgs += \
            '# ページ番号の書式（無、有、n :、-n-、n/N等）を指定できます。'
        cfgs += '\n'
        if cls.page_number == '':
            cfgs += '頁番号: 無\n'
        elif cls.page_number == DEFAULT_PAGE_NUMBER:
            cfgs += '頁番号: 有\n'
        else:
            cfgs += '頁番号: ' + cls.page_number + '\n'
        cfgs += '\n'

        cfgs += \
            '# 行番号の記載（無、有）を指定できます。'
        cfgs += '\n'
        if cls.line_number:
            cfgs += '行番号: 有\n'
        else:
            cfgs += '行番号: 無\n'
        cfgs += '\n'

        cfgs += \
            '# 明朝体とゴシック体と異字体（IVS）のフォントを指定できます。'
        cfgs += '\n'
        if '/' in cls.mincho_font:
            cfgs += '明朝体: ' + cls.mincho_font + '\n'
        else:
            cfgs += '明朝体: = / ' + cls.mincho_font + '\n'
        if '/' in cls.gothic_font:
            cfgs += 'ゴシ体: ' + cls.gothic_font + '\n'
        else:
            cfgs += 'ゴシ体: = / ' + cls.gothic_font + '\n'
        cfgs += '異字体: ' + cls.ivs_font + '\n'
        cfgs += '\n'

        cfgs += \
            '# 基本の文字の大きさをポイント単位で指定できます。'
        cfgs += '\n'
        cfgs += '文字サ: ' + str(round(cls.font_size, 1)) + ' pt\n'
        cfgs += '\n'

        cfgs += \
            '# 行間の高さを基本の文字の高さの何倍にするかを指定できます。'
        cfgs += '\n'
        cfgs += '行間高: ' + str(round(cls.line_spacing, 2)) + ' 倍\n'
        cfgs += '\n'

        cfgs += \
            '# セクションタイトル前後の余白を行間の高さの倍数で指定できます。'
        cfgs += '\n'
        cfgs += '前余白: ' + re.sub(',', ' 倍,', cls.space_before) + ' 倍\n'
        cfgs += '後余白: ' + re.sub(',', ' 倍,', cls.space_after) + ' 倍\n'
        cfgs += '\n'

        cfgs += \
            '# 半角文字と全角文字の間の間隔調整（無、有）を指定できます。'
        cfgs += '\n'
        if cls.auto_space:
            cfgs += '字間整: 有\n'
        else:
            cfgs += '字間整: 無\n'
        cfgs += '\n'

        if cls.version_number != '':
            cfgs += \
                '# 文書のバージョン番号を文字列で指定できます。'
            cfgs += '\n'
            cfgs += '版番号: ' + cls.version_number + '\n'
            cfgs += '\n'

        if cls.content_status != '':
            cfgs += \
                '# 文書の状態を文字列で指定できます。'
            cfgs += '\n'
            cfgs += '書状態: ' + cls.content_status + '\n'
            cfgs += '\n'

        cfgs += \
            '# 備考書（コメント）などを消して完成させます。'
        cfgs += '\n'
        if cls.has_completed:
            cfgs += '完成稿: 真\n'
        else:
            cfgs += '完成稿: 偽\n'
        cfgs += '\n'

        cfgs += \
            '# 変換元のWordファイルの最終更新日時が自動で指定されます。'
        cfgs += '\n'
        cfgs += '元原稿: ' + cls.original_file + '\n'
        cfgs += '\n'

        cfgs += \
            '---------------------------------------------------------------->'
        cfgs += '\n\n'

        return cfgs

    @staticmethod
    def get_styles(xml_lines):
        styles = []
        xml_body = XML.get_body('w:styles', xml_lines)
        xml_blocks = XML.get_blocks(xml_body)
        for n, xb in enumerate(xml_blocks):
            s = Style(n + 1, xb)
            styles.append(s)
        # Form.styles = styles
        return styles

    @staticmethod
    def get_rels(xml_lines):
        rels = {}
        res = '^<Relationship Id=[\'"](.*)[\'"] .* Target=[\'"](.*)[\'"]/>$'
        for xl in xml_lines:
            if re.match(res, xl):
                rel_id = re.sub(res, '\\1', xl)
                rel_tg = re.sub(res, '\\2', xl)
                rels[rel_id] = rel_tg
        # Form.rels = rels
        return rels

    @staticmethod
    def get_remarks(xml_lines):
        remarks = {}
        res_beg = '^<w:comment w:id="([^"]+)"( .*)?>$'
        res_end = '^</w:comment>$'
        remark_id = ''
        remark_str = ''
        is_in_remarks = False
        for xl in xml_lines:
            if re.match(res_beg, xl):
                remark_id = re.sub(res_beg, '\\1', xl)
                remark_str = ''
                is_in_remarks = True
            elif re.match(res_end, xl):
                remarks[remark_id] = remark_str
                is_in_remarks = False
            if re.match('^<.*>$', xl):
                continue
            if is_in_remarks:
                remark_str += xl
        return remarks


class CharsDatum:

    """A class to keep characters data"""

    res_fds = [['->', '<-'], ['\\+>', '<\\+'],
               ['_{', '_}'], ['\\^{', '\\^}'],
               ['`', '`'], ['@[^@]{1,66}@', '@[^@]{1,66}@'],
               ['\\^[0-9A-Za-z]{0,11}\\^', '\\^[0-9A-Za-z]{0,11}\\^'],
               ['_[0-9A-Za-z]{1,11}_', '_[0-9A-Za-z]{1,11}_'],
               ['_[\\$=\\.#\\-~\\+]{,4}_', '_[\\$=\\.#\\-~\\+]{,4}_'],
               ['>>>', '<<<'], ['>>', '<<'],
               ['---', '---'], ['--', '--'],
               ['\\+\\+', '\\+\\+'], ['\\+\\+\\+', '\\+\\+\\+'],
               ['<<', '>>'], ['<<<', '>>>'],
               ['~~', '~~'], ['\\*\\*', '\\*\\*'], ['\\*', '\\*']]

    def __init__(self, pre_fds, chars, pos_fds):
        # PRE FDS
        self.pre_fds = []
        self.raw_pre_fds = []
        for fd in pre_fds:
            self.pre_fds.append(fd)
            self.raw_pre_fds.append(fd)
        # CHARS
        self.chars = chars
        # POS FDS
        self.pos_fds = []
        self.raw_pos_fds = []
        for fd in pos_fds:
            self.pos_fds.append(fd)
            self.raw_pos_fds.append(fd)

    def append_fds(self, pre_fd, pos_fd):
        group_fds = [['---', '--', '++', '+++'], ['>>>', '>>', '<<', '<<<']]
        for gr_fds in group_fds:
            if pre_fd in gr_fds:
                for fd in gr_fds:
                    if fd in self.pre_fds:
                        self.pre_fds.remove(fd)
                    if fd in self.raw_pre_fds:
                        self.raw_pre_fds.remove(fd)
            if pos_fd in gr_fds:
                for fd in gr_fds:
                    if fd in self.pos_fds:
                        self.pos_fds.remove(fd)
                    if fd in self.raw_pos_fds:
                        self.raw_pos_fds.remove(fd)
        # PRE FDS
        if pre_fd not in self.pre_fds:
            self.pre_fds.append(pre_fd)
        if pre_fd not in self.raw_pre_fds:
            self.raw_pre_fds.append(pre_fd)
        # POS FDS
        if pos_fd not in self.pos_fds:
            self.pos_fds.append(pos_fd)
        if pos_fd not in self.raw_pos_fds:
            self.raw_pos_fds.append(pos_fd)

    def get_chars_with_fd(self):
        chars = self.chars
        pre_fds = self.pre_fds
        pos_fds = self.pos_fds
        pre = ''
        pos = ''
        for rf in self.res_fds:
            pre, pre_fds = self._touch_in_pre(pre, pre_fds, rf[0])
            pos, pos_fds = self._touch_in_pos(pos, pos_fds, rf[1])
        pre = pre + ''.join(pre_fds)        # just in case
        pos = ''.join(pos_fds[::-1]) + pos  # just in case
        cwf = chars
        cwf = self._concatenate_chars(pre, cwf)
        cwf = self._concatenate_chars(cwf, pos)
        return cwf

    @staticmethod
    def _concatenate_chars(chars1, chars2):
        # "~" + "~"
        if re.match(NOT_ESCAPED + '~$', chars1) and re.match('^~', chars2):
            return chars1 + '<>' + chars2
        # "/" + "/"
        if re.match(NOT_ESCAPED + '/$', chars1) and re.match('^/', chars2):
            return chars1 + '<>' + chars2
        # "-" + "-"
        if re.match(NOT_ESCAPED + '-$', chars1) and re.match('^-', chars2):
            return chars1 + '<>' + chars2
        # "+" + "+"
        if re.match(NOT_ESCAPED + '\\+$', chars1) and re.match('^\\+', chars2):
            return chars1 + '<>' + chars2
        # ">" + ">"
        if re.match(NOT_ESCAPED + '>$', chars1) and re.match('^>', chars2):
            return chars1 + '<>' + chars2
        # "<" + "<"
        if re.match(NOT_ESCAPED + '<$', chars1) and re.match('^<', chars2):
            return chars1 + '<>' + chars2
        # "<" + ">"
        if re.match(NOT_ESCAPED + '<$', chars1) and re.match('^>', chars2):
            return chars1 + '<>' + chars2
        # "@.*" + ".*@"
        if re.match(NOT_ESCAPED + '@([^@]{0,66})$', chars1) and \
           not re.match(NOT_ESCAPED + '@([^@]{1,66})@[^@]*$', chars1) and \
           re.match('^([^@]{0,66})@(.|\n)*', chars2) and \
           not re.match('^[^@]*@([^@]{1,66})@(.|\n)*', chars2):
            c1 = re.sub(NOT_ESCAPED + '@([^@]{0,66})$', '\\2', chars1)
            c2 = re.sub('^([^@]{0,66})@(.|\n)*', '\\1', chars2)
            if len(c1 + c2) <= 66:
                return chars1 + '<>' + chars2
        # "_.*" + ".*_"
        if re.match(NOT_ESCAPED + '_([\\$=\\.#\\-~\\+]*)$', chars1) and \
           re.match('^([\\$=\\.#\\-~\\+]*)_(.|\n)*', chars2):
            c1 = re.sub(NOT_ESCAPED + '_([\\$=\\.#\\-~\\+]*)$', '\\2', chars1)
            c2 = re.sub('^([\\$=\\.#\\-~\\+]*)_(.|\n)*', '\\1', chars2)
            for ul in UNDERLINE:
                if c1 + c2 == UNDERLINE[ul]:
                    return chars1 + '<>' + chars2
        # "^.*" + ".*^"
        if re.match(NOT_ESCAPED + '\\^([0-9a-zA-Z]*)$', chars1) and \
           re.match('^([0-9a-zA-Z]*)\\^(.|\n)*', chars2):
            c1 = re.sub(NOT_ESCAPED + '\\^([0-9a-zA-Z]*)$', '\\2', chars1)
            c2 = re.sub('^([0-9a-zA-Z]*)\\^(.|\n)*', '\\1', chars2)
            if re.match('^([0-9A-F]{3})([0-9A-F]{3})?$', c1 + c2):
                return chars1 + '<>' + chars2
            for fc in FONT_COLOR:
                if c1 + c2 == FONT_COLOR[fc]:
                    return chars1 + '<>' + chars2
        # "_.*" + ".*_"
        if re.match(NOT_ESCAPED + '_([0-9a-zA-Z]*)$', chars1) and \
           re.match('^([0-9a-zA-Z]*)_(.|\n)*', chars2):
            c1 = re.sub(NOT_ESCAPED + '_([0-9a-zA-Z]*)$', '\\2', chars1)
            c2 = re.sub('^([0-9a-zA-Z]*)_(.|\n)*', '\\1', chars2)
            for hc in HIGHLIGHT_COLOR:
                if (c1 + c2 == hc) or (c1 + c2 == HIGHLIGHT_COLOR[hc]):
                    return chars1 + '<>' + chars2
        # "-|+" + ">"
        # if re.match(NOT_ESCAPED + '(-|\\+)$', chars1) and \
        #    re.match('^>', chars2):
        #     return chars1 + '<>' + chars2
        # "<" + "-|+"
        # if re.match(NOT_ESCAPED + '<$', chars1) and \
        #    re.match('^(-|\\+)', chars2):
        #     return chars1 + '<>' + chars2
        # "\" + "[|]"
        # if re.match(NOT_ESCAPED + '\\\\$', chars1) and \
        #    re.match('^(\\[|\\])', chars2):
        #     return chars1 + '<>' + chars2
        return chars1 + chars2

    @staticmethod
    def _touch_in_pre(pre, fds, res):
        match = []
        for fd in fds:
            if re.match('^' + res + '$', fd):
                match.append(fd)
        if len(match) > 0:
            for m in match:
                fds.remove(m)
            pre = pre + match[-1]
        return pre, fds

    @staticmethod
    def _touch_in_pos(pos, fds, res):
        match = []
        for fd in fds:
            if re.match('^' + res + '$', fd):
                match.append(fd)
        if len(match) > 0:
            for m in match:
                fds.remove(m)
            if re.match('^(_|\\^)}', match[-1]):
                pos = '}' + pos  # '(_|^)}' -> '}'
            else:
                pos = match[-1] + pos
        return pos, fds

    @classmethod
    def are_consecutive(cls, unit_data1, unit_data2):
        pos_fds = []
        for fd in unit_data1.pos_fds:
            pos_fds.append(fd)
        pre_fds = []
        for fd in unit_data2.pre_fds:
            pre_fds.append(fd)
        for rf in cls.res_fds:
            for pos in pos_fds:
                if re.match('^' + rf[1] + '$', pos):
                    for pre in pre_fds:
                        if re.match('^' + rf[0] + '$', pre):
                            if (rf[1] == rf[0]) and (pos != pre):
                                continue
                            pos_fds.remove(pos)
                            pre_fds.remove(pre)
                            break
                    break
        if pos_fds == [] and pre_fds == []:
            return True
        return False

    @classmethod
    def cancel_fds(cls, unit_data1, unit_data2):
        pos_fds = unit_data1.pos_fds
        pre_fds = unit_data2.pre_fds
        for rf in cls.res_fds:
            for pos in pos_fds:
                if re.match('^' + rf[1] + '$', pos):
                    for pre in pre_fds:
                        if re.match('^' + rf[0] + '$', pre):
                            if (rf[1] == rf[0]) and (pos != pre):
                                continue
                            pos_fds.remove(pos)
                            pre_fds.remove(pre)
                            break
                    break
        return unit_data1, unit_data2


class XML:

    """A class to handle xml"""

    @staticmethod
    def get_body(tag_name, xml_lines):
        xml_body = []
        is_in_body = False
        for xl in xml_lines:
            if re.match('^</?' + tag_name + '( .*)?>$', xl):
                is_in_body = not is_in_body
                continue
            if is_in_body:
                xml_body.append(xl)
        return xml_body

    @staticmethod
    def get_blocks(xml_body):
        xml_blocks = []
        res_oneline_tag = '<(\\S+)( .*)?/>'
        res_beginning_tag = '<(\\S+)( .*)?>'
        xb = []
        xml_class = None
        xml_depth = 0
        for xl in xml_body:
            if xml_class == '':
                # ABNORMAL STATE (JUST TO MAKE SURE)
                if not re.match(res_beginning_tag, xl):
                    # ABNORMAL STATE CONTINUES
                    xb.append(xl)
                    continue
                else:
                    # SAVE AND RESET
                    xml_blocks.append(xb)
                    xb = []
                    xml_class = None
                    xml_depth = 0
            # NORMAL STATE
            xb.append(xl)
            if xml_class is None:
                if re.match(res_oneline_tag, xl):
                    # SAVE AND RESET
                    xml_blocks.append(xb)
                    xb = []
                    xml_class = None
                    xml_depth = 0
                elif re.match(res_beginning_tag, xl):
                    xml_class = re.sub(res_beginning_tag, '\\1', xl)
                    xml_depth = 1
                    res_class_tag = '<' + xml_class + '( .*)?>'
                    res_end_tag = '</' + xml_class + '>'
                else:
                    # MOVE TO ABNORMAL STATE
                    xml_class = ''
            elif re.match(res_class_tag, xl):
                xml_depth += 1
            elif re.match(res_end_tag, xl):
                xml_depth -= 1
                if xml_depth == 0:
                    # SAVE AND RESET
                    xml_blocks.append(xb)
                    xb = []
                    xml_class = None
                    xml_depth = 0
            else:
                pass
        if len(xb) > 0:
            # SAVE AND RESET (JUST TO MAKE SURE)
            xml_blocks.append(xb)
            xb = []
            xml_class = None
            xml_depth = 0
        return xml_blocks

    @staticmethod
    def get_value(tag_name, value_name, init_value, tag):
        if re.match('<' + tag_name + ' .+>', tag):
            res = '^.* ' + value_name + '=[\'"]([^\'"]*)[\'"].*$'
            if re.match(res, tag):
                value = re.sub(res, '\\1', tag)
                if type(init_value) is int:
                    # INT
                    if re.match('^[-\\+]?[0-9]+$', value):
                        return int(value)
                    if re.match('^true$', value, re.IGNORECASE):
                        return 1
                    if re.match('^false$', value, re.IGNORECASE):
                        return -1
                    return init_value  # bad value
                if type(init_value) is float:
                    # FLOAT
                    if re.match('^' + RES_NUMBER + '$', value):
                        return float(value)
                    return init_value  # bad value
                if type(init_value) is bool:
                    # BOOL
                    if re.match('^true$', value, re.IGNORECASE):
                        return True
                    if re.match('^false$', value, re.IGNORECASE):
                        return False
                    if value == '1':
                        return True
                    if value == '-1':
                        return False
                    return init_value  # bad value
                # STRING
                return value
        return init_value

    @staticmethod
    def is_this_tag(tag_name, init_value, tag):
        if re.match('<' + tag_name + '( .*)?/?>', tag):
            return True
        else:
            return init_value


class Document:

    """A class to handle document"""

    images = {}

    def __init__(self):
        self.docx_file = None
        self.md_file = None
        self.document_xml_lines = None
        self.raw_paragraphs = None
        self.paragraphs = None

    def get_raw_paragraphs(self, xml_lines):
        raw_paragraphs = []
        xml_body = XML.get_body('w:body', xml_lines)
        xml_blocks = XML.get_blocks(xml_body)
        for xb in xml_blocks:
            rp = RawParagraph(xb)
            raw_paragraphs.append(rp)
        # self.raw_paragraphs = raw_paragraphs
        return raw_paragraphs

    def get_paragraphs(self, raw_paragraphs):
        paragraphs = []
        for rp in raw_paragraphs:
            if rp.paragraph_class == 'configuration':
                if len(paragraphs) > 0:
                    if paragraphs[-1].md_text == '<pgbr>':
                        paragraphs[-1].md_text = '<Pgbr>'
                    if paragraphs[-1].md_lines_text == '<pgbr>':
                        paragraphs[-1].md_lines_text = '<Pgbr>'
                    if paragraphs[-1].text_to_write == '<pgbr>':
                        paragraphs[-1].text_to_write = '<Pgbr>'
                    if paragraphs[-1].text_to_write_with_reviser == '<pgbr>':
                        paragraphs[-1].text_to_write_with_reviser = '<Pgbr>'
                    # ATTACHED PAGE BREAK
                    if paragraphs[-1].attached_pagebreak == 'pgbr':
                        paragraphs[-1].attached_pagebreak = 'Pgbr'
                continue
            p = rp.get_paragraph()
            paragraphs.append(p)
        # self.paragraphs = paragraphs
        return paragraphs

    def modify_paragraphs(self):
        # CHANGE PARAGRAPH CLASS
        self.paragraphs = self._modpar_left_alignment()
        self.paragraphs = self._modpar_blank_paragraph_to_space_before()
        # CHANGE VIRTUAL LENGTH
        self.paragraphs = self._modpar_article_title()
        self.paragraphs = self._modpar_section_space_before_and_after()
        self.paragraphs = self._modpar_spaced_and_centered()
        self.paragraphs = self._modpar_length_reviser_to_depth_setter()
        # CHANGE HORIZONTAL LENGTH
        self.paragraphs = self._modpar_one_line_paragraph()
        self.paragraphs = self._modpar_cancel_first_indent()
        # CHANGE VERTICAL LENGTH
        self.paragraphs = self._modpar_vertical_length()
        # ISOLATE FONT REVISERS
        self.paragraphs = self._modpar_isolate_revisers()
        # RETURN
        return self.paragraphs

    def _modpar_left_alignment(self):
        # |                    ->  |
        # |(first indent = 0)  ->  |: 段落
        # |(lef indent = 0)    ->  |
        # |段落                ->  |
        # |                    ->  |
        for i, p in enumerate(self.paragraphs):
            if p.has_removed:
                continue
            if p.paragraph_class == 'sentence':
                if p.length_docx['first indent'] == 0:
                    if p.length_docx['left indent'] == 0:
                        p.paragraph_class = 'alignment'
                        p.alignment = 'left'
                        mt = ''
                        for text in p.md_text.split('\n'):
                            mt += ': ' + re.sub('<br>$', '', text) + '\n'
                        mt = re.sub('\n+$', '', mt)
                        p.md_text = mt
                        p.md_lines_text = p._get_md_lines_text(p.md_text)
                        p.text_to_write = p.get_text_to_write()
                        p.text_to_write_with_reviser \
                            = p.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_blank_paragraph_to_space_before(self):
        # |              ->  |
        # |v=1           ->  |v=2
        # |(blank line)  ->  |段落
        # |段落          ->  |
        # |              ->  |
        for i, p in enumerate(self.paragraphs):
            if p.has_removed:
                continue
            p_next = self._get_next_paragraph(self.paragraphs, i)
            if p.paragraph_class == 'blank':
                v_line = p.md_text.count('\n') + 1.0
                p.md_text = ''
                p.length_supp['space before'] += v_line
                # RENEW
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines_text = p._get_md_lines_text(p.md_text)
                # p.text_to_write = p.get_text_to_write()
                p.text_to_write_with_reviser \
                    = p.get_text_to_write_with_reviser()
                p.paragraph_class = 'empty'
            if p.paragraph_class == 'empty' and p_next is not None:
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
                # p.md_lines_text = p._get_md_lines_text(p.md_text)
                # p.text_to_write = p.get_text_to_write()
                p.text_to_write_with_reviser \
                    = p.get_text_to_write_with_reviser()
                p_next.length_revi = p_next._get_length_revi()
                p_next.length_revisers \
                    = p_next._get_length_revisers(p_next.length_revi)
                # p_next.md_lines_text \
                #     = p_next._get_md_lines_text(p_next.md_text)
                # p_next.text_to_write = p_next.get_text_to_write()
                p_next.text_to_write_with_reviser \
                    = p_next.get_text_to_write_with_reviser()
        return self.paragraphs

    # ARTICLE TITLE (MIMI=EAR)
    def _modpar_article_title(self):
        # |                    ->  |
        # |<!--                ->  |<!--
        # |document_style: j   ->  |document_style: j
        # |space_before:   ,1  ->  |space_before:   ,1
        # |-->                 ->  |-->
        # |                    ->  |
        # |: （条文の耳）      ->  |(space)
        # |                    ->  |: （条文の耳）
        # |(space)             ->  |
        # |## 条文本文         ->  |## 条文本文
        # |                    ->  |
        if Form.document_style != 'j':
            return self.paragraphs
        for i, p in enumerate(self.paragraphs):
            if p.has_removed:
                continue
            p_prev = self._get_prev_paragraph(self.paragraphs, i)
            if p.paragraph_class == 'section' and \
               p.head_section_depth == 2 and \
               p.tail_section_depth == 2 and \
               p_prev is not None and \
               p_prev.paragraph_class == 'alignment' and \
               p_prev.alignment == 'left':
                p_prev.length_conf['space before'] \
                    = p.length_conf['space before']
                p.length_conf['space before'] = 0.0
                # RENEW
                p_prev.length_revi = p_prev._get_length_revi()
                p_prev.length_revisers \
                    = p_prev._get_length_revisers(p_prev.length_revi)
                # p_prev.md_lines_text \
                #     = p_prev._get_md_lines_text(p_prev.md_text)
                # p_prev.text_to_write = p_prev.get_text_to_write()
                p_prev.text_to_write_with_reviser \
                    = p_prev.get_text_to_write_with_reviser()
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines_text = p._get_md_lines_text(p.md_text)
                # p.text_to_write = p.get_text_to_write()
                p.text_to_write_with_reviser \
                    = p.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_section_space_before_and_after(self):
        # |               ->  |
        # |v=+0.2 V=+0.2  ->  |v=+0.1 V=+0.1
        # |# タイトル     ->  |# タイトル
        # |               ->  |
        # |v=-0.5 V=-0.2  ->  ||項目|項目|
        # ||項目|項目|    ->  ||:--:|:--:|
        # ||:--:|:--:|    ->  ||セル|セル|
        # ||セル|セル|    ->  |
        # |               ->  |
        for i, p in enumerate(self.paragraphs):
            if p.has_removed:
                continue
            p_prev = self._get_prev_paragraph(self.paragraphs, i)
            p_next = self._get_next_paragraph(self.paragraphs, i)
            # TITLE
            ds = ParagraphSection._get_section_depths(p.raw_text_doi,
                                                      not p.has_removed)
            if p.paragraph_class == 'section' and ds == (1, 1):
                # BEFORE
                if p_prev is not None:
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
                if p_next is not None:
                    if p_next.length_docx['space before'] >= 0.1:
                        p_next.length_docx['space before'] += 0.1
                    elif p_next.length_docx['space before'] >= 0.0:
                        p_next.length_docx['space before'] *= 2
            # TABLE
            elif p.paragraph_class == 'table':
                if p_prev is None or p_prev.paragraph_class == 'pagebreak':
                    p.length_supp['space before'] += TABLE_SPACE_BEFORE
                else:
                    p.length_docx['space before'] \
                        = p_prev.length_docx['space after']
                    p_prev.length_docx['space after'] = 0.0
                if p_next is None or p_next.paragraph_class == 'pagebreak':
                    p.length_supp['space after'] += TABLE_SPACE_AFTER
                else:
                    p.length_docx['space after'] \
                        = p_next.length_docx['space before']
                    p_next.length_docx['space before'] = 0.0
            # IMAGE
            elif p.paragraph_class == 'image':
                if p_prev is None or p_prev.paragraph_class == 'pagebreak':
                    p.length_supp['space before'] += IMAGE_SPACE_BEFORE
                else:
                    p.length_docx['space before'] \
                        = p_prev.length_docx['space after']
                    p_prev.length_docx['space after'] = 0.0
                if p_next is None or p_next.paragraph_class == 'pagebreak':
                    p.length_supp['space after'] += IMAGE_SPACE_AFTER
                else:
                    p.length_docx['space after'] \
                        = p_next.length_docx['space before']
                    p_next.length_docx['space before'] = 0.0
            else:
                continue
            # RENEW
            if p_prev is not None:
                p_prev.length_revi = p_prev._get_length_revi()
                p_prev.length_revisers \
                    = p_prev._get_length_revisers(p_prev.length_revi)
                # p_prev.md_lines_text \
                #     = p_prev._get_md_lines_text(p_prev.md_text)
                # p_prev.text_to_write = p_prev.get_text_to_write()
                p_prev.text_to_write_with_reviser \
                    = p_prev.get_text_to_write_with_reviser()
            if True:
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines_text = p._get_md_lines_text(p.md_text)
                # p.text_to_write = p.get_text_to_write()
                p.text_to_write_with_reviser \
                    = p.get_text_to_write_with_reviser()
            if p_next is not None:
                p_next.length_revi = p_next._get_length_revi()
                p_next.length_revisers \
                    = p_next._get_length_revisers(p_next.length_revi)
                # p_next.md_lines_text \
                #     = p_next._get_md_lines_text(p_next.md_text)
                # p_next.text_to_write = p_next.get_text_to_write()
                p_next.text_to_write_with_reviser \
                    = p_next.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_spaced_and_centered(self):
        # |              ->  |
        # |v=1           ->  |v=1
        # |: 添付資料 :  ->  |# ##=1 ###=1
        # |              ->  |
        # |###=1         ->  |: 添付資料 :
        # |### 資料1     ->  |
        # |              ->  |### 資料1
        # |### 資料2     ->  |
        # |              ->  |### 資料2
        # |              ->  |
        # self.paragraphs = self._modpar_blank_paragraph_to_space_before()
        Paragraph.previous_head_section_depth = 0
        Paragraph.previous_tail_section_depth = 0
        for i, p in enumerate(self.paragraphs):
            if p.has_removed:
                continue
            p_next = self._get_next_paragraph(self.paragraphs, i)
            if p.paragraph_class == 'alignment' and \
               p.alignment == 'center' and \
               p.length_revi['space before'] == 1.0:
                Paragraph.previous_head_section_depth = 1
                Paragraph.previous_tail_section_depth = 1
                p.pre_text_to_write += 'v=+1.0\n#'
                if p_next is not None:
                    if p_next.paragraph_class == 'section' and \
                       p_next.head_section_depth == 3 and \
                       p_next.tail_section_depth == 3 and \
                       p_next.section_states[1][0] == 0 and \
                       p_next.section_states[2][0] == 1 and \
                       p_next.section_states[2][1] == 0:
                        p.pre_text_to_write += ' ##=1'
                        p.pre_text_to_write += ' ###=1'
                        if '##=1' in p_next.numbering_revisers:
                            p_next.numbering_revisers.remove('##=1')
                        if '###=1' in p_next.numbering_revisers:
                            p_next.numbering_revisers.remove('###=1')
                p.pre_text_to_write += '\n'
                p.length_supp['space before'] -= 1.0
            p.head_section_depth, p.tail_section_depth \
                = p._get_section_depths(p.raw_text_doi, not p.has_removed)
            p.length_clas = p._get_length_clas()
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            # p.md_lines_text = p._get_md_lines_text(p.md_text)
            # p.text_to_write = p.get_text_to_write()
            p.text_to_write_with_reviser = p.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_length_reviser_to_depth_setter(self):
        # |               ->  |
        # |## セクション  ->  |## セクション
        # |               ->  |
        # |<=+1.0         ->  |#
        # |段落           ->  |
        # |               ->  |段落
        # |               ->  |
        # self.paragraphs = self._modpar_spaced_and_centered()
        res_gg = '^<<=(' + RES_NUMBER + ')$'
        res_g = '^<=(' + RES_NUMBER + ')$'
        res_l = '^>=(' + RES_NUMBER + ')$'
        for i, p in enumerate(self.paragraphs):
            if i == 0:
                continue
            p_prev = self._get_prev_paragraph(self.paragraphs, i)
            if p.paragraph_class != 'sentence':
                continue
            is_in_reviser = False
            for j in range(i - 1, 0, -1):
                p_tmp = self.paragraphs[j]
                if p_tmp.paragraph_class == 'section':
                    break
                if p_tmp.paragraph_class == 'sentence':
                    if re.match('^#+\n$', p_tmp.pre_text_to_write):
                        is_in_reviser = True
                        break
            left_indent = int(p.length_revi['left indent'])
            if not is_in_reviser:
                if p.length_revi['space before'] != 0.0 or \
                   p.length_revi['space after'] != 0.0 or \
                   p.length_revi['line spacing'] != 0.0 or \
                   p.length_revi['first indent'] != 0.0 or \
                   p.length_revi['right indent'] != 0.0 or \
                   p.length_revi['left indent'] >= 0.0 or \
                   not p.length_revi['left indent'].is_integer():
                    continue
                if p.head_section_depth + left_indent < 1:
                    continue
            p.head_section_depth += left_indent
            p.tail_section_depth += left_indent
            if p.section_states[1][0] == 0 and \
               p.section_states[2][0] > 0 and \
               p.head_section_depth + left_indent == 2:
                p.head_section_depth -= 1
                p.tail_section_depth -= 1
            p.length_clas['left indent'] = p.head_section_depth
            p.pre_text_to_write = '#' * p.head_section_depth + ' \n'
            # REMOVE SAME AS BEFORE
            for j in range(i - 1, 0, -1):
                p_tmp = self.paragraphs[j]
                if p_tmp.paragraph_class == 'section':
                    break
                if p_tmp.paragraph_class == 'sentence':
                    if re.match('^#+\n$', p_tmp.pre_text_to_write):
                        if p.pre_text_to_write == p_tmp.pre_text_to_write:
                            p.pre_text_to_write = ''
            # RENEW
            p.length_clas = p._get_length_clas()
            # p.length_conf = p._get_length_conf()
            # p.length_supp = p._get_length_supp()
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            # ParagraphList.reset_states(p.paragraph_class)
            # p.md_lines_text = p._get_md_lines_text(p.md_text)
            # p.text_to_write = p.get_text_to_write()
            p.text_to_write_with_reviser = p.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_one_line_paragraph(self):
        # |                ->  |
        # |<<=-1.0 <=+1.0  ->  |1行の段落
        # |1行の段落       ->  |
        # |                ->  |
        paper_size = Form.paper_size
        left_margin = Form.left_margin
        right_margin = Form.right_margin
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
                    p.md_lines_text = p._get_md_lines_text(p.md_text)
                    # p.text_to_write = p.get_text_to_write()
                    p.text_to_write_with_reviser \
                        = p.get_text_to_write_with_reviser()
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
            # p.md_lines_text = p._get_md_lines_text(p.md_text)
            # p.text_to_write = p.get_text_to_write()
            p.text_to_write_with_reviser = p.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_cancel_first_indent(self):
        # |             ->  |
        # |<<=+1.0      ->  |1行の段落
        # |　1行の段落  ->  |
        # |             ->  |
        res = '^([ \t\u3000]+)((?:.|\n)*)$'
        for p in self.paragraphs:
            if len(p.chars_data) == 0:
                continue
            if not re.match(res, p.chars_data[0].chars):
                continue
            fds = p.chars_data[0].raw_pre_fds + p.chars_data[0].raw_pos_fds
            if '--' in fds or '---' in fds or '++' in fds or '+++' in fds or \
               '>>' in fds or '>>>' in fds or '<<' in fds or '<<<' in fds:
                continue
            hsps = re.sub(res, '\\1', p.chars_data[0].chars)
            w = 0
            for c in hsps:
                if c == ' ':
                    w += 0.5
                elif c == '\t':
                    w += 4.0
                elif c == '\u3000':
                    w += 1.0
            if w + p.length_revi['first indent'] != 0:
                continue
            p.length_supp['first indent'] = w
            p.text_to_write = re.sub('^' + hsps, '', p.text_to_write)
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            # p.md_lines_text = p._get_md_lines_text(p.md_text)
            # p.text_to_write = p.get_text_to_write()
            p.text_to_write_with_reviser = p.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_vertical_length(self):
        # |                  ->  |
        # |<!--              ->  |<!--
        # |space_before: ,1  ->  |space_before: ,1
        # |space_after: ,1   ->  |space_after: ,1
        # |-->               ->  |-->
        # |                  ->  |
        # |V=+1.0            ->  |V=+1.0
        # |## 前段落1        ->  |## 前段落1
        # |                  ->  |
        # |v=-1.0            ->  |## 後段落1
        # |## 後段落2        ->  |
        # |                  ->  |## 前段落2
        # |V=-1.0            ->  |
        # |## 前段落3        ->  |v=+1.0
        # |                  ->  |## 後段落2
        # |v=+1.0            ->  |
        # |## 後段落4        ->  |
        # |                  ->  |
        m = len(self.paragraphs) - 1
        for i, p in enumerate(self.paragraphs):
            p_prev = self._get_prev_paragraph(self.paragraphs, i)
            p_next = self._get_next_paragraph(self.paragraphs, i)
            for lr in p.length_revisers:
                # PREV
                if p_prev is not None and re.match('^v=-.*', lr):
                    must_remove = True
                    for plr in p_prev.length_revisers:
                        if re.match('^V=-.*', plr):
                            must_remove = False
                    if must_remove:
                        p.length_revisers.remove(lr)
                # NEXT
                if p_next is not None and re.match('^V=-.*', lr):
                    must_remove = True
                    for nlr in p_next.length_revisers:
                        if re.match('^v=-.*', nlr):
                            must_remove = False
                    if must_remove:
                        p.length_revisers.remove(lr)
            # RENEW
            p.text_to_write_with_reviser = p.get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_isolate_revisers(self):
        # |           ->  |
        # |**段落1**  ->  |**
        # |           ->  |
        # |**段落2**  ->  |段落1
        # |           ->  |
        # |           ->  |段落2
        # |           ->  |
        # |           ->  |**
        # |           ->  |
        base_cd = None
        for i, p in enumerate(self.paragraphs):
            curr_head = p.head_font_revisers
            curr_tail = p.tail_font_revisers
            curr_cd = CharsDatum(curr_head, '', curr_tail)
            next_head = []
            next_tail = []
            p_next = self._get_next_paragraph(self.paragraphs, i)
            if p_next is not None:
                next_head = p_next.head_font_revisers
                next_tail = p_next.tail_font_revisers
            next_cd = CharsDatum(next_head, '', next_tail)
            is_cc, is_cn, is_nn = False, False, False
            if CharsDatum.are_consecutive(curr_cd, curr_cd):
                is_cc = True
            if CharsDatum.are_consecutive(curr_cd, next_cd):
                is_cn = True
            if CharsDatum.are_consecutive(next_cd, next_cd):
                is_nn = True
            # SINGLE LINE
            if (base_cd is None) and is_cc and (not (is_cn and is_nn)):
                p.head_font_revisers = []
                p.tail_font_revisers = []
                tex_fd = CharsDatum(curr_head, '', []).get_chars_with_fd()
                if tex_fd != '':
                    p.text_to_write = tex_fd + '\n' + p.text_to_write
                tex_fd = CharsDatum([], '', curr_tail).get_chars_with_fd()
                if tex_fd != '':
                    p.text_to_write = p.text_to_write + '\n' + tex_fd
            # MULTI LINES
            if base_cd is None:
                if is_cc and is_cn and is_nn:
                    p.head_font_revisers = []
                    p.tail_font_revisers = []
                if is_cc and is_cn and is_nn:
                    tex_fd = CharsDatum(curr_head, '', []).get_chars_with_fd()
                    if tex_fd != '':
                        p.pre_text_to_write \
                            = re.sub('\\s*\n$', ' ', p.pre_text_to_write)
                        p.pre_text_to_write += tex_fd + '\n'
                    base_cd = curr_cd
            else:
                if True:
                    p.head_font_revisers = []
                    p.tail_font_revisers = []
                if not (is_cc and is_cn and is_nn):
                    tex_fd = CharsDatum([], '', curr_tail).get_chars_with_fd()
                    if tex_fd != '':
                        p.post_text_to_write \
                            = re.sub('^\n', ' ', p.post_text_to_write)
                        p.post_text_to_write += '\n' + tex_fd
                    base_cd = None
            # RENEW
            p.text_to_write_with_reviser = p.get_text_to_write_with_reviser()
        return self.paragraphs

    @staticmethod
    def _get_prev_paragraph(paras, base):
        for i in range(base - 1, -1, -1):
            if not paras[i].has_removed:
                return paras[i]
        return None

    @staticmethod
    def _get_next_paragraph(paras, base):
        for i in range(base + 1, len(paras)):
            if not paras[i].has_removed:
                return paras[i]
        return None

    def get_document(self):
        dcmt = ''
        for p in self.paragraphs:
            dcmt += p.get_document()
            if p.paragraph_class != 'empty' and p.paragraph_class != 'remarks':
                dcmt += '\n'
        return dcmt

    def get_images(self):
        return self.images
        # imgs = {}
        # for p in self.paragraphs:
        #     tmp_imgs = p.get_images()
        #     imgs.update(tmp_imgs)
        # return imgs


class Style:

    """A class to handle style"""

    def __init__(self, number, xml_lines):
        self.number = number
        self.xml_lines = xml_lines
        self.type = None
        self.styleid = None
        self.name = None
        self.font = None
        self.font_size = None
        self.font_italic = False
        self.font_bold = False
        self.font_strike = False
        self.font_underline = None
        self.font_color = None
        self.font_highlight_color = None
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
        f_2s = None
        f_it = False
        f_bd = False
        f_sk = False
        f_ul = None
        f_cl = None
        f_hc = None
        alig = None
        rl = {'sb': None, 'sa': None, 'ls': None,
              'fi': None, 'hi': None, 'li': None, 'ri': None}
        for xl in self.xml_lines:
            type = XML.get_value('w:style', 'w:type', type, xl)
            stid = XML.get_value('w:style', 'w:styleId', stid, xl)
            name = XML.get_value('w:name', 'w:val', name, xl)
            font = XML.get_value('w:rFonts', 'w:ascii', font, xl)
            font = XML.get_value('w:rFonts', 'w:eastAsia', font, xl)
            # font = XML.get_value('w:rFonts', '*', font, xl)
            f_2s = XML.get_value('w:sz', 'w:val', f_2s, xl)
            f_it = XML.is_this_tag('w:i', f_it, xl)
            f_bd = XML.is_this_tag('w:b', f_bd, xl)
            f_sk = XML.is_this_tag('w:strike', f_sk, xl)
            f_ul = XML.get_value('w:u', 'w:val', f_ul, xl)
            f_cl = XML.get_value('w:color', 'w:val', f_cl, xl)
            f_hc = XML.get_value('w:highlight', 'w:val', f_hc, xl)
            alig = XML.get_value('w:jc', 'w:val', alig, xl)
            rl['sb'] = XML.get_value('w:spacing', 'w:before', rl['sb'], xl)
            rl['sa'] = XML.get_value('w:spacing', 'w:after', rl['sa'], xl)
            rl['ls'] = XML.get_value('w:spacing', 'w:line', rl['ls'], xl)
            rl['ls'] = XML.get_value('w:spacing', 'w:line', rl['ls'], xl)
            rl['fi'] = XML.get_value('w:ind', 'w:firstLine', rl['fi'], xl)
            rl['hi'] = XML.get_value('w:ind', 'w:hanging', rl['hi'], xl)
            rl['li'] = XML.get_value('w:ind', 'w:left', rl['li'], xl)
            rl['ri'] = XML.get_value('w:ind', 'w:right', rl['ri'], xl)
        self.type = type
        self.styleid = stid
        self.name = name
        self.font = font
        if f_2s is not None:
            self.font_size = round(float(f_2s) / 2, 1)
        self.is_italic = f_it
        self.is_bold = f_bd
        self.has_strike = f_sk
        self.underline = f_ul
        self.font_color = f_cl
        self.highlight_color = f_hc
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

    def __init__(self, xml_lines):
        # DECLARATION
        self.raw_paragraph_number = -1
        self.has_removed = False
        self.xml_lines = []
        self.raw_class = ''
        self.horizontal_line = ''  # 'top'|'bottom'|'textbox'
        self.attached_pagebreak = ''  # 'pgbr' | 'Pgbr'
        self.chars_data = []
        self.images = {}
        self.raw_text = ''
        self.head_space = ''
        self.tail_space = ''
        self.raw_text_del = ''
        self.raw_text_ins = ''
        self.raw_text_doi = ''
        self.remarks = []
        self.style = ''
        self.alignment = ''
        self.paragraph_class = ''
        # SUBSTITUTION
        RawParagraph.raw_paragraph_number += 1
        self.raw_paragraph_number = RawParagraph.raw_paragraph_number
        self.xml_lines = xml_lines
        self.raw_class = self._get_raw_class(self.xml_lines)
        self.horizontal_line \
            = self._get_horizontal_line(self.raw_class, self.xml_lines)
        self.attached_pagebreak = self.get_attached_pagebreak(self.xml_lines)
        self.chars_data, self.images \
            = self._get_chars_data_and_images(self.raw_class, self.xml_lines)
        self.raw_text = self._get_raw_text(self.chars_data)
        self.head_space, self.raw_text \
            = self._separate_head_space(self.raw_text,
                                        '->', '<-', '\\+>', '<\\+')
        rts, rrt \
            = self._separate_head_space(self.raw_text[::-1],
                                        '>-', '-<', '>\\+', '\\+<')
        self.raw_text = rrt[::-1]
        self.tail_space = rts[::-1]
        self.raw_text_del = self._get_raw_text_del(self.raw_text)
        self.raw_text_ins = self._get_raw_text_ins(self.raw_text)
        if self.raw_text_ins != '':
            self.raw_text_doi = self.raw_text_ins
        else:
            self.raw_text_doi = self.raw_text_del
        if self.raw_text_del != '' and self.raw_text_ins == '':
            self.has_removed = True
        self.remarks = self._get_remarks(xml_lines)
        self.style = self._get_style(xml_lines)
        self.alignment = self._get_alignment(self.xml_lines)
        self.paragraph_class = self._get_paragraph_class()

    @staticmethod
    def _get_raw_class(xml_lines):
        res = '^<(\\S+)( .*)?>$'
        xlz = xml_lines[0]
        if re.match(res, xlz):
            return re.sub(res, '\\1', xlz)
        else:
            return None

    @staticmethod
    def _get_horizontal_line(raw_class, xml_lines):
        for xl in xml_lines:
            # HORIZONTAL LINE
            if raw_class != 'w:tbl' and re.match('^<w:top( .*)?>$', xl):
                # HORIZONTAL LINE (TOPLINE)
                return 'top'
            if raw_class != 'w:tbl' and re.match('^<w:bottom( .*)?>$', xl):
                # HORIZONTAL LINE (BOTTOMLINE)
                return 'bottom'
            res = '^<v:rect( .*)? style="width:0;height:1.5pt"( .*)?>$'
            if re.match(res, xl):
                # HORIZONTAL LINE (TEXTBOX)
                return 'textbox'
        return ''

    @staticmethod
    def get_attached_pagebreak(xml_lines):
        for xl in xml_lines:
            if re.match('^<w:br w:type=[\'"]page[\'"]/>$', xl):
                return 'pgbr'
        return ''

    @classmethod
    def _get_chars_data_and_images(cls, raw_class, xml_lines, type='normal'):
        if raw_class != 'w:tbl':
            font_size = Form.font_size
        else:
            font_size = Form.font_size * TABLE_FONT_SCALE
        chars_data = []
        images = {}
        img_rels = Form.rels
        img_file_name = ''
        img_size = ''
        is_changed = False
        fldchar = ''
        track_changes = ''  # ''|'del'|'ins'
        ruby = ''  # ''|'rub'|'bas'
        width = 100
        cd = CharsDatum([], '', [])
        for xl in xml_lines:
            # EMPTY
            if xl == '':
                continue
            # RPRCHANGE
            if re.match('^<w:rPrChange( .*[^/])?>$', xl):
                is_changed = True
            if re.match('^</w:rPrChange( .*[^/])?>$', xl):
                is_changed = False
            if is_changed:
                continue
            # FOR PAGE NUMBER
            if re.match('^<w:fldChar w:fldCharType="begin"/?>$', xl):
                fldchar = 'begin'
            elif re.match('^<w:fldChar w:fldCharType="separate"/?>$', xl):
                cd.pre_fds, cd.pos_fds = [], []
                fldchar = 'separate'
            elif re.match('^<w:fldChar w:fldCharType="end"/?>$', xl):
                fldchar = 'end'
            if fldchar == 'separate':
                continue
            # MATH
            if 'math_data' not in locals():
                math_data = None
            math_data, chars_data \
                = cls._manage_math_expression(xl, math_data, chars_data)
            if math_data is not None:
                continue
            # IMAGE
            must_continue = False
            if re.match(RES_XML_IMG_MS, xl):
                # IMAGE MS WORD
                img_rel_name, img_file_name \
                    = cls._get_img_file_names_ms(xl, img_rels)
                Document.images[img_rel_name] = img_file_name
                images[img_rel_name] = img_file_name
                must_continue = True
            elif re.match(RES_XML_IMG_PY_ID, xl):
                # IMAGE PYTHON-DOCX ID
                img_rel_name, img_file_name \
                    = cls._get_img_file_names_py(xl, img_rels, img_py_name)
                Document.images[img_rel_name] = img_file_name
                images[img_rel_name] = img_file_name
                must_continue = True
            elif re.match(RES_XML_IMG_PY_NAME, xl):
                # IMAGE PYTHON-DOCX NAME
                img_py_name = re.sub(RES_XML_IMG_PY_NAME, '\\2', xl)
                must_continue = True
            elif re.match(RES_XML_IMG_SIZE, xl):
                # IMAGE SIZE
                img_size = cls._get_img_size(xl)
                must_continue = True
            if img_file_name != '' and img_size != '':
                imt = cls._get_img_md_text(img_file_name, img_size, font_size)
                cd_img = CharsDatum([], '', [])
                if track_changes == 'del':
                    cd_img.append_fds('->', '<-')
                elif track_changes == 'ins':
                    cd_img.append_fds('+>', '<+')
                if re.match('^---(.*)---$', imt):
                    imt = re.sub('^---(.*)---$', '\\1', imt)
                    cd_img.append_fds('---', '---')
                elif re.match('^--(.*)--$', imt):
                    imt = re.sub('^--(.*)--$', '\\1', imt)
                    cd_img.append_fds('--', '--')
                elif re.match('^\\+\\+\\+(.*)\\+\\+\\+$', imt):
                    imt = re.sub('^\\+\\+\\+(.*)\\+\\+\\+$', '\\1', imt)
                    cd_img.append_fds('+++', '+++')
                elif re.match('^\\+\\+(.*)\\+\\+$', imt):
                    imt = re.sub('^\\+\\+(.*)\\+\\+$', '\\1', imt)
                    cd_img.append_fds('++', '++')
                cd_img.chars = '<>' + imt  # '<>' is to avoid being escaped
                chars_data.append(cd_img)
                img_file_name = ''
                img_size = ''
            if must_continue:
                continue
            # TRACK CHANGES
            if re.match('^<w:del( .*[^/])?>$', xl):
                track_changes = 'del'
                continue
            elif re.match('^</w:del( .*[^/])?>$', xl):
                track_changes = ''
                continue
            elif re.match('^<w:ins( .*[^/])?>$', xl):
                track_changes = 'ins'
                continue
            elif re.match('^</w:ins( .*[^/])?>$', xl):
                track_changes = ''
                continue
            # RUBY
            if re.match('^<w:ruby>$', xl):
                chars_data.append(CharsDatum([], '^<', []))
                ruby = 'rub'
                continue
            elif re.match('^<w:rubyBase>$', xl):
                chars_data.append(CharsDatum([], '>/<', []))
                ruby = 'bas'
                continue
            elif re.match('^</w:ruby>$', xl):
                chars_data.append(CharsDatum([], '>$', []))
                ruby = ''
                continue
            # FONT
            if re.match('^<w:rFonts .*>$', xl):
                afnt = XML.get_value('w:rFonts', 'w:ascii', '', xl)
                kfnt = XML.get_value('w:rFonts', 'w:eastAsia', '', xl)
                # SYMPTOMATIC TREATMENT
                for mfs in MS_FONTS:
                    if afnt in mfs:
                        afnt = mfs[0]
                    if kfnt in mfs:
                        kfnt = mfs[0]
                font = ''
                if afnt != '' and kfnt != '':
                    if afnt == kfnt:
                        font = afnt
                    else:
                        font = afnt + ' / ' + kfnt
                elif afnt != '' and kfnt == '':
                    font = afnt
                elif afnt == '' and kfnt != '':
                    font = kfnt
                if font != '':
                    if font == Form.mincho_font:
                        pass
                    elif font == Form.gothic_font:
                        cd.append_fds('`', '`')
                    else:
                        cd.append_fds('@' + font + '@', '@' + font + '@')
                continue
            # ITALIC
            if re.match('^<w:i/?>$', xl):
                cd.append_fds('*', '*')
                continue
            # BOLD
            if re.match('^<w:b/?>$', xl):
                cd.append_fds('**', '**')
                continue
            # STRIKETHROUGH
            if re.match('^<w:strike/?>$', xl):
                cd.append_fds('~~', '~~')
                continue
            # FONT SIZE
            s = XML.get_value('w:sz', 'w:val', -1.0, xl) / 2
            if s > 0:
                if s < font_size * 0.7:
                    cd.append_fds('---', '---')
                elif s < font_size * 0.9:
                    cd.append_fds('--', '--')
                elif s > font_size * 1.3:
                    cd.append_fds('+++', '+++')
                elif s > font_size * 1.1:
                    cd.append_fds('++', '++')
                continue
            # FONT WIDTH
            w = XML.get_value('w:w', 'w:val', -1.0, xl)
            if w > 0:
                if w < 70:
                    cd.append_fds('>>>', '<<<')
                elif w < 90:
                    cd.append_fds('>>', '<<')
                elif w > 130:
                    cd.append_fds('<<<', '>>>')
                elif w > 110:
                    cd.append_fds('<<', '>>')
                width = w
                continue
            # UNDERLINE
            if re.match('^<w:u( .*)?>$', xl):
                underline = ''
                res = '^<.* w:val=[\'"]([a-zA-Z]+)[\'"].*>$'
                if re.match(res, xl):
                    val = re.sub(res, '\\1', xl)
                    if val in UNDERLINE:
                        underline = UNDERLINE[val]
                cd.append_fds('_' + underline + '_', '_' + underline + '_')
                continue
            # FONT COLOR
            if re.match('^<w:color w:val="[0-9A-F]+"( .*)?/?>$', xl):
                val = re.sub('^<.* w:val="([0-9A-F]+)".*>$', '\\1', xl, re.I)
                val = val.upper()
                if val == 'FFFFFF':
                    font_color = ''
                elif val in FONT_COLOR:
                    font_color = FONT_COLOR[val]
                else:
                    font_color = val
                cd.append_fds('^' + font_color + '^', '^' + font_color + '^')
                continue
            # HIGHLIGHT COLOR
            if re.match('^<w:highlight w:val="[a-zA-Z]+"( .*)?/?>$', xl):
                val = re.sub('^<.* w:val="([a-zA-Z]+)".*>$', '\\1', xl)
                highlight = val
                cd.append_fds('_' + highlight + '_', '_' + highlight + '_')
                continue
            # SUBSCRIPT OR SUPERSCRIPT
            if xl == '<w:vertAlign w:val="subscript"/>':
                cd.append_fds('_{', '_}')
                continue
            elif xl == '<w:vertAlign w:val="superscript"/>':
                cd.append_fds('^{', '^}')
                continue
            # TEXT
            if not re.match('^<.*>$', xl):
                t = cls._prepare_chars(fldchar, xl, type)
                cd.chars = CharsDatum._concatenate_chars(cd.chars, t)
                continue
            elif re.match('^<w:tab/?>$', xl):
                cd.chars += '\t'
                continue
            elif re.match('^<w:br/?>$', xl):
                cd.chars += '\n'
                continue
            # RUN
            if re.match('^<w:r( .*)?>$', xl):
                continue
            elif re.match('^</w:r>$', xl):
                if cd.chars != '':
                    if track_changes == 'del':
                        cd.append_fds('->', '<-')
                    elif track_changes == 'ins':
                        cd.append_fds('+>', '<+')
                    # RUBY
                    if ruby == 'rub':
                        if '---' in cd.pre_fds:
                            cd.pre_fds.remove('---')
                        if '--' in cd.pre_fds:
                            cd.pre_fds.remove('--')
                        if '---' in cd.pos_fds:
                            cd.pos_fds.remove('---')
                        if '--' in cd.pos_fds:
                            cd.pos_fds.remove('--')
                    # SPACE
                    if re.match('^\u3000+$', cd.chars) and width != 100:
                        n = len(cd.chars)
                        for fd in ['>>>', '>>', '<<', '<<<']:
                            if fd in cd.pre_fds:
                                cd.pre_fds.remove(fd)
                            if fd in cd.pos_fds:
                                cd.pos_fds.remove(fd)
                        cd.chars = '<' + str(float(width * n) / 100) + '>'
                    chars_data.append(cd)
                width = 100
                cd = CharsDatum([], '', [])
                continue
        # self.chars_data = chars_data
        # self.images = images
        return chars_data, images

    @staticmethod
    def _get_img_file_names_ms(xl, img_rels):
        img_id = re.sub(RES_XML_IMG_MS, '\\1', xl)
        img_rel_name = img_rels[img_id]
        img_ext = re.sub('^.*\\.', '', img_rel_name)
        img_base = re.sub(RES_XML_IMG_MS, '\\2', xl)
        img_base = re.sub('\\s', '_', img_base)
        i = 0
        while True:
            img_file_name = img_base + '.' + img_ext
            if i > 0:
                img_file_name = img_base + str(i) + '.' + img_ext
            for j in Document.images:
                if j != img_rel_name:
                    if Document.images[j] == img_file_name:
                        break
            else:
                break
            i += 1
        return img_rel_name, img_file_name

    @staticmethod
    def _get_img_file_names_py(xl, img_rels, img_py_name):
        img_id = re.sub(RES_XML_IMG_PY_ID, '\\1', xl)
        img_rel_name = img_rels[img_id]
        img_ext = re.sub('^.*\\.', '', img_rel_name)
        img_base = re.sub('\\.' + img_ext + '$', '', img_py_name)
        img_base = re.sub('\\s', '_', img_base)
        i = 0
        while True:
            img_file_name = img_base + '.' + img_ext
            if i > 0:
                img_file_name = img_base + str(i) + '.' + img_ext
            for j in Document.images:
                if j != img_rel_name:
                    if Document.images[j] == img_file_name:
                        break
            else:
                break
            i += 1
        return img_rel_name, img_file_name

    @staticmethod
    def _get_img_size(xl):
        sz_w = re.sub(RES_XML_IMG_SIZE, '\\1', xl)
        sz_h = re.sub(RES_XML_IMG_SIZE, '\\2', xl)
        cm_w = float(sz_w) * 2.54 / 72 / 12700
        cm_h = float(sz_h) * 2.54 / 72 / 12700
        if cm_w >= 1:
            cm_w = round(cm_w, 1)
        else:
            cm_w = round(cm_w, 2)
        if cm_h >= 1:
            cm_h = round(cm_h, 1)
        else:
            cm_h = round(cm_h, 2)
        img_size = str(cm_w) + 'x' + str(cm_h)
        return img_size

    @staticmethod
    def _get_img_md_text(img_file_name, img_size, font_size):
        relative_dir = os.path.basename(IO.media_dir)
        m_size_cm = font_size * 2.54 / 72
        xs_size_cm = m_size_cm * 0.6
        s_size_cm = m_size_cm * 0.8
        l_size_cm = m_size_cm * 1.2
        xl_size_cm = m_size_cm * 1.4
        # cm_w = float(re.sub('x.*$', '', img_size))
        cm_h = float(re.sub('^.*x', '', img_size))
        img_md_text = '![' + img_file_name + ']' \
            + '(' + relative_dir + '/' + img_file_name + ')'
        if cm_h >= m_size_cm * 0.98 and cm_h <= m_size_cm * 1.02:
            # MEDIUM
            pass
        elif cm_h >= xs_size_cm * 0.98 and cm_h <= xs_size_cm * 1.02:
            # XSMALL
            img_md_text = '---' + img_md_text + '---'
        elif cm_h >= s_size_cm * 0.98 and cm_h <= s_size_cm * 1.02:
            # SMALL
            img_md_text = '--' + img_md_text + '--'
        elif cm_h >= l_size_cm * 0.98 and cm_h <= l_size_cm * 1.02:
            # LARGE
            img_md_text = '++' + img_md_text + '++'
        elif cm_h >= xl_size_cm * 0.98 and cm_h <= xl_size_cm * 1.02:
            # XLARGE
            img_md_text = '+++' + img_md_text + '+++'
        else:
            # FREE SIZE
            img_md_text = '![' + img_file_name + ':' + img_size + ']' \
                + '(' + relative_dir + '/' + img_file_name + ')'
        return img_md_text

    @classmethod
    def _manage_math_expression(cls, xl, math_data, chars_data):
        if re.match('^<m:oMath>$', xl):
            math_data = []
            math_data.append(CharsDatum([''], '', ['']))
            math_data.append(CharsDatum([''], '', ['']))
            return math_data, chars_data
        elif re.match('^</m:oMath>$', xl):
            md = math_data[0]
            if md.chars == '' and md.pre_fds == [''] and md.pos_fds == ['']:
                math_data.pop(0)
            md = math_data[-1]
            if md.chars == '' and md.pre_fds == [''] and md.pos_fds == ['']:
                math_data.pop(-1)
            math_str = cls._join_math_expression(math_data)
            math_str = cls._shape_math_matrix(math_str)
            math_str = cls._shape_sub_and_sup(math_str)
            math_str = cls._replace_binomial(math_str)
            chars_data.append(CharsDatum([], '\\[' + math_str + '\\]', []))
            math_data = None
            return math_data, chars_data
        if math_data is None:
            return math_data, chars_data
        md_pre = math_data[-2]
        md_cur = math_data[-1]
        if False:
            pass
        # NEXT
        elif (xl == '</w:rPr>' or xl == '<w:rPr/>') and md_cur.chars != '':
            math_data.append(CharsDatum([''], '', ['']))
            md_pre = math_data[-2]
            md_cur = math_data[-1]
        elif xl == '</m:r>' or xl == '<m:r/>':
            math_data.append(CharsDatum([''], '', ['']))
            md_pre = math_data[-2]
            md_cur = math_data[-1]
        # TEXT
        elif not re.match('^<.*>$', xl):
            xl = re.sub('{', '\\{', xl)
            xl = re.sub('}', '\\}', xl)
            xl = re.sub(' ', '\\\\,', xl)
            # FUNCTION
            if re.match('^.*<fun>.*$', md_cur.pre_fds[0]):
                md_cur.pre_fds[0] = re.sub('<fun>', '', md_cur.pre_fds[0])
                md_cur.chars += '\\'
                if '*' in md_cur.pre_fds:
                    md_cur.pre_fds.remove('*')
                if '*' in md_cur.pos_fds:
                    md_cur.pos_fds.remove('*')
            md_cur.chars += xl
        # FUNCTION
        elif xl == '<m:fName>':
            md_cur.pre_fds[0] += '<fun>'
        elif xl == '</m:fName>':
            pass
        # ELEMENT
        elif xl == '<m:e>':
            md_cur.pre_fds[0] += '{'
        elif xl == '</m:e>':
            if md_cur.chars == '':
                md_pre.pos_fds[0] += '}'
            else:
                md_cur.pos_fds[0] += '}'
        elif xl == '<m:e/>':
            md_cur.pre_fds[0] += '{}'
        # SUP AND SUB
        elif xl == '<m:sPre>':
            md_cur.pre_fds[0] += '{}'
        elif xl == '</m:sPre>':
            pass
        elif xl == '<m:sup>':
            md_cur.pre_fds[0] += '^{'
        elif xl == '</m:sup>':
            md_pre.pos_fds[0] += '}'
        elif xl == '<m:sub>':
            md_cur.pre_fds[0] += '_{'
        elif xl == '</m:sub>':
            md_pre.pos_fds[0] += '}'
        # VECTOR
        elif xl == '<m:chr m:val="⃗"/>':
            md_cur.chars += '\\vec'
        # FRACTION, BINOMIAL
        elif xl == '<m:f>':
            md_cur.chars += '\\frac'
        elif xl == '<m:type m:val="noBar"/>':
            if md_cur.chars == '\\frac':
                md_cur.chars = '\\Xbinom'
        elif xl == '</m:f>':
            pass
        elif xl == '<m:num>':
            md_cur.pre_fds[0] += '{'
        elif xl == '</m:num>':
            md_pre.pos_fds[0] += '}'
        elif xl == '<m:den>':
            md_cur.pre_fds[0] += '{'
        elif xl == '</m:den>':
            md_pre.pos_fds[0] += '}'
        # RADICAL ROOT
        elif xl == '<m:rad>':
            md_cur.chars += '\\sqrt'
            math_data.append(CharsDatum([''], '', ['']))
            md_pre = math_data[-2]
            md_cur = math_data[-1]
        elif xl == '<m:deg>':
            md_cur.pre_fds[0] += '['
        elif xl == '</m:deg>':
            md_cur.pos_fds[0] += ']'
            if md_cur.chars == '' and \
               md_cur.pre_fds[0] == '[' and md_cur.pos_fds[0] == ']':
                math_data.pop(-1)
            math_data.append(CharsDatum([''], '', ['']))
            md_pre = math_data[-2]
            md_cur = math_data[-1]
        # LIMIT
        elif xl == '<m:lim>':
            md_cur.pre_fds[0] += '_{'
            if md_pre.chars == '\\lim' and \
               re.match('^.*{$', md_pre.pre_fds[0]) and \
               re.match('^}.*$', md_pre.pos_fds[0]):
                md_pre.pre_fds[0] = re.sub('{$', '', md_pre.pre_fds[0])
                md_pre.pos_fds[0] = re.sub('^}', '', md_pre.pos_fds[0])
        elif xl == '</m:lim>':
            md_pre.pos_fds[0] += '}'
        # INTEGRAL
        elif xl == '<m:nary>':
            md_cur.chars += '\\int'
        # DOUBLE INTEGRAL
        elif xl == '<m:chr m:val="∬"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\iint', md_cur.chars)
        # TRIPLE INTEGRAL
        elif xl == '<m:chr m:val="∭"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\iiint', md_cur.chars)
        # LINE INTEGRAL
        elif xl == '<m:chr m:val="∮"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\oint', md_cur.chars)
        # SIGMA
        elif xl == '<m:chr m:val="∑"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\sum', md_cur.chars)
        # PI
        elif xl == '<m:chr m:val="∏"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\prod', md_cur.chars)
        # LINE BREAK
        elif re.match('^<m:brk( .*)?/>$', xl):
            md_cur.pre_fds[0] += '\\\\'
        # MATRIX
        elif xl == '<m:m>':
            md_cur.chars += '\\Xbegin{matrix}'
        elif xl == '</m:mr>':
            md_cur.pre_fds[0] += '\\\\'
        elif xl == '</m:m>':
            md_cur.chars += '\\Xend{matrix}'
            math_data.append(CharsDatum([''], '', ['']))
            md_pre = math_data[-2]
            md_cur = math_data[-1]
        # PARENTHESES
        elif xl == '<m:d>':
            md_cur.chars += '(<()>'
        elif re.match('^<m:begChr m:val=""/>$', xl):
            md_cur.chars = re.sub('\\(<(.)(.)>$', '(<>', md_cur.chars)
        elif re.match('^<m:begChr m:val="(.)"/>$', xl):
            bc = re.sub('^<m:begChr m:val="(.)"/>$', '\\1', xl)
            md_cur.chars = re.sub('\\(<(.)(.)>$', '(<' + bc + '\\2>',
                                  md_cur.chars)
        elif re.match('<m:endChr m:val=""/>', xl):
            md_cur.chars = re.sub('\\(<(.)(.)>$', '(<>', md_cur.chars)
        elif re.match('<m:endChr m:val="(.)"/>', xl):
            ec = re.sub('^<m:endChr m:val="(.)"/>$', '\\1', xl)
            md_cur.chars = re.sub('\\(<(.)(.)>$', '(<\\g<1>' + ec + '>',
                                  md_cur.chars)
        elif xl == '</m:d>':
            res = '(.*)\\(<(.?)(.?)>(.*)$'
            end = ')'
            md_beg_chars = None
            for i in range(len(math_data) - 1, -1, -1):
                if re.match(res, math_data[i].chars):
                    md_beg_chars = math_data[i].chars
                    pre = re.sub(res, '\\1', md_beg_chars)
                    beg = re.sub(res, '\\2', md_beg_chars)
                    end = re.sub(res, '\\3', md_beg_chars)
                    pos = re.sub(res, '\\4', md_beg_chars)
                    if beg == '{':
                        beg = '\\{'
                    if end == '}':
                        end = '\\}'
                    math_data[i].chars = pre + beg + pos
                    md_cur.pre_fds = ['']
                    for j in range(1, len(math_data[i].pre_fds)):
                        fd = math_data[i].pre_fds[j]
                        md_cur.pre_fds.append(fd)
                    md_cur.pos_fds = ['']
                    for j in range(1, len(math_data[i].pos_fds)):
                        fd = math_data[i].pos_fds[j]
                        md_cur.pos_fds.append(fd)
                    break
            md_cur.chars += end
            math_data.append(CharsDatum([''], '', ['']))
        # FONT SIZE
        size = Form.font_size
        s = XML.get_value('w:sz', 'w:val', -1.0, xl) / 2
        if s > 0:
            if s < size * 0.7:
                md_cur.append_fds('---', '---')
            elif s < size * 0.9:
                md_cur.append_fds('--', '--')
            elif s > size * 1.3:
                md_cur.append_fds('+++', '+++')
            elif s > size * 1.1:
                md_cur.append_fds('++', '++')
        # FONT WIDTH
        w = XML.get_value('w:w', 'w:val', -1.0, xl)
        if w > 0:
            if w < 70:
                md_cur.append_fds('>>>', '<<<')
            elif w < 90:
                md_cur.append_fds('>>', '<<')
            elif w > 130:
                md_cur.append_fds('<<<', '>>>')
            elif w > 110:
                md_cur.append_fds('<<', '>>')
        # STRIKETHROUGH
        elif re.match('^<w:strike/?>$', xl):
            md_cur.append_fds('~~', '~~')
        # UNDERLINE
        elif re.match('^<w:u w:val="(.*)"/>$', xl):
            tx = re.sub('^<w:u w:val="(.*)"/>$', '\\1', xl)
            md_cur.append_fds('_' + UNDERLINE[tx] + '_',
                              '_' + UNDERLINE[tx] + '_')
        # FONT COLOR
        elif re.match('^<w:color w:val="([^"]*)"(?: .*)?/>$', xl):
            fc = re.sub('^<w:color w:val="([^"]*)"(?: .*)?/>$', '\\1', xl)
            if fc == 'FFFFFF':
                fc = ''
            elif fc in FONT_COLOR:
                fc = FONT_COLOR[fc]
            md_cur.append_fds('^' + fc + '^', '^' + fc + '^')
        # HIGILIGHT COLOR
        elif re.match('^<w:highlight w:val="(.*)"/>$', xl):
            hc = re.sub('^<w:highlight w:val="(.*)"/>$', '\\1', xl)
            md_cur.append_fds('_' + hc + '_', '_' + hc + '_')
        # DEL OR INS
        elif xl == '<w:del>':
            md_cur.pre_fds.append('->')
        elif xl == '</w:del>':
            md_cur.pos_fds.append('<-')
        elif xl == '<w:ins>':
            md_cur.pre_fds.append('+>')
        elif xl == '</w:ins>':
            md_cur.pos_fds.append('<+')
        # BOLD OR ROMAN
        elif re.match('<m:sty m:val="([^"]+)"/>', xl):
            tx = re.sub('<m:sty m:val="([^"]+)"/>', '\\1', xl)
            # ROMAN
            if tx == 'p' or tx == 'b':
                md_cur.pre_fds.append('*')
                md_cur.pos_fds.append('*')
            # BOLD
            if tx == 'bi' or tx == 'b':
                md_cur.pre_fds.append('**')
                md_cur.pos_fds.append('**')
        return math_data, chars_data

    @classmethod
    def _join_math_expression(cls, math_data):
        # CANCEL
        fdlist = [['\\*', ''], ['\\*\\*', ''], ['~~', ''], ['__', ''],
                  ['---', ''], ['--', ''], ['\\+\\+', ''], ['\\+\\+\\+', ''],
                  ['>>>', '<<<'], ['>>', '<<'], ['<<', '>>'], ['<<<', '>>>'],
                  ['\\^([0-9A-Za-z]+)\\^', ''], ['_([0-9A-Za-z]+)_', ''],
                  ]
        for fds in fdlist:
            if fds[1] == '':
                fds[1] = fds[0]
            is_set_anywhere = True
            for i in range(len(math_data)):
                for j in range(1, len(math_data[i].pre_fds)):
                    if re.match('^' + fds[0] + '$', math_data[i].pre_fds[j]):
                        break
                else:
                    is_set_anywhere = False
                    break
                for j in range(1, len(math_data[i].pos_fds)):
                    if re.match('^' + fds[1] + '$', math_data[i].pos_fds[j]):
                        break
                else:
                    is_set_anywhere = False
                    break
            if is_set_anywhere:
                for i in range(len(math_data)):
                    if i > 0:
                        pre = math_data[i].pre_fds
                        for j in range(1, len(pre)):
                            if re.match('^' + fds[0] + '$', pre[j]):
                                pre.pop(j)
                                break
                    if i < len(math_data) - 1:
                        pos = math_data[i].pos_fds
                        for j in range(1, len(pos)):
                            if re.match('^' + fds[1] + '$', pos[j]):
                                pos.pop(j)
                                break
        # REPLACE FDS
        for mada in math_data:
            for i in range(len(mada.pre_fds)):
                if mada.pre_fds[i] == '*':
                    mada.pre_fds[i] = '{\\mathrm{'
                elif mada.pre_fds[i] == '**':
                    mada.pre_fds[i] = '{\\mathbf{'
                elif mada.pre_fds[i] == '~~':
                    mada.pre_fds[i] = '{\\sout{'
                elif mada.pre_fds[i] == '__':
                    mada.pre_fds[i] = '{\\underline{'
                elif mada.pre_fds[i] == '---':
                    mada.pre_fds[i] = '{\\footnotesize{'
                elif mada.pre_fds[i] == '--':
                    mada.pre_fds[i] = '{\\small{'
                elif mada.pre_fds[i] == '++':
                    mada.pre_fds[i] = '{\\large{'
                elif mada.pre_fds[i] == '+++':
                    mada.pre_fds[i] = '{\\Large{'
                elif re.match('^\\^([0-9A-Za-z]+)\\^$', mada.pre_fds[i]):
                    c = re.sub('\\^(.+)\\^', '\\1', mada.pre_fds[i])
                    mada.pre_fds[i] = '{\\textcolor{' + c + '}{'
                elif re.match('^_([0-9A-Za-z]+)_$', mada.pre_fds[i]):
                    c = re.sub('_(.+)_', '\\1', mada.pre_fds[i])
                    mada.pre_fds[i] = '{\\colorbox{' + c + '}{'
            for i in range(len(mada.pos_fds)):
                if mada.pos_fds[i] == '*':
                    mada.pos_fds[i] = '}}'
                elif mada.pos_fds[i] == '**':
                    mada.pos_fds[i] = '}}'
                elif mada.pos_fds[i] == '~~':
                    mada.pos_fds[i] = '}}'
                elif mada.pos_fds[i] == '__':
                    mada.pos_fds[i] = '}}'
                elif mada.pos_fds[i] == '---':
                    mada.pos_fds[i] = '}}'
                elif mada.pos_fds[i] == '--':
                    mada.pos_fds[i] = '}}'
                elif mada.pos_fds[i] == '++':
                    mada.pos_fds[i] = '}}'
                elif mada.pos_fds[i] == '+++':
                    mada.pos_fds[i] = '}}'
                elif re.match('^\\^([0-9A-Za-z]+)\\^$', mada.pos_fds[i]):
                    mada.pos_fds[i] = '}}'
                elif re.match('^_([0-9A-Za-z]+)_$', mada.pos_fds[i]):
                    mada.pos_fds[i] = '}}'
        # JOIN
        math_str = ''
        for mada in math_data:
            pre_str = ''.join(mada.pre_fds)
            str_str = mada.chars
            pos_str = ''.join(mada.pos_fds[::-1])
            math_str += pre_str + str_str + pos_str
        return math_str

    @staticmethod
    def _shape_math_matrix(math_str):
        # FONT DECRATIONS
        res = '{.*\\\\Xbegin{matrix}.*\\\\Xend{matrix}}{.*\\)}+'
        math_str = ParagraphMath.shift_paren('\\(', 2, res, math_str)
        math_str = ParagraphMath.cancel_multi_paren(math_str)
        math_str = re.sub('{(\\\\Xbegin{matrix})}', '\\1', math_str)
        math_str = re.sub('(\\\\Xend{matrix}}){\\)}', '\\1)', math_str)
        # CONFIRM TYPE
        tlist = [['\\(', '\\)', 'p'], ['\\[', '\\]', 'b'],
                 ['\\|', '\\|', 'v'], ['‖', '‖', 'V']]
        tmp = ''
        while tmp != math_str:
            tmp = math_str
            for t in tlist:
                beg_fr = '^(.*)' + t[0] + '({)\\\\Xbegin{matrix}(.*)$'
                beg_to = '\\1\\2\\\\Xbegin{' + t[2] + 'matrix}\\3'
                math_str = re.sub(beg_fr, beg_to, math_str)
                end_fr = '^(.*)\\\\Xend{matrix}(})' + t[1] + '(.*)$'
                end_to = '\\1\\\\Xend{' + t[2] + 'matrix}\\2\\3'
                math_str = re.sub(end_fr, end_to, math_str)
        # SHAPE CELL
        res = '^(.*?){' + \
            '\\\\Xbegin({.?matrix})(.*?)\\\\Xend({.?matrix})' + \
            '}(.*?)$'
        while re.match(res, math_str):
            str1 = re.sub(res, '\\1', math_str)
            mtx1 = re.sub(res, '\\2', math_str)
            roco = re.sub(res, '\\3', math_str)
            mtx2 = re.sub(res, '\\4', math_str)
            str2 = re.sub(res, '\\5', math_str)
            d = 0
            s = ''
            for c in roco:
                s += c
                if c == '{':
                    d += 1
                if c == '}':
                    d -= 1
                if d == 0 and c == '}':
                    s += '&'
                if re.match('.*&\\\\\\\\$', s):
                    s = re.sub('&\\\\\\\\$', '\\\\\\\\', s)
            roco = re.sub('\\\\\\\\$', '', s)
            math_str = str1 + '\\begin' + mtx1 \
                + roco \
                + '\\end' + mtx2 + str2
        return math_str

    @staticmethod
    def _shape_sub_and_sup(math_str):
        res = '{}_{([^{}]*(?:{[^{}]*})?[^{}]*)}' \
            + '{{([^{}]*(?:{[^{}]*})?[^{}]*)}' \
            + '_{([^{}]*(?:{[^{}]*})?[^{}]*)}}'
        tmp = ''
        while tmp != math_str:
            tmp = math_str
            math_str = re.sub(res, '{}_{\\1}{\\2}_{\\3}', math_str)
        return math_str

    @staticmethod
    def _replace_binomial(math_str):
        res = '\\({\\\\Xbinom{(.*?)}{(.*?)}}\\)'
        tmp = ''
        while tmp != math_str:
            tmp = math_str
            math_str = re.sub(res, '\\\\binom{\\1}{\\2}', math_str)
        return math_str

    @staticmethod
    def _prepare_chars(fldchar, input_chars, type='normal'):
        chars = input_chars
        # ESCAPE
        chars = chars.replace('\\', '\\\\')
        chars = chars.replace('*', '\\*')
        chars = chars.replace('`', '\\`')
        chars = chars.replace('~~', '\\~\\~')
        chars = chars.replace('//', '\\/\\/')  # italic
        chars = re.sub('([a-z]+:)\\\\/\\\\/', '\\1//', chars)  # http https ftp
        chars = chars.replace('---', '\\-\\-\\-')
        chars = chars.replace('--', '\\-\\-')
        chars = chars.replace('+++', '\\+\\+\\+')
        chars = chars.replace('++', '\\+\\+')
        chars = chars.replace('>>>', '\\>\\>\\>')
        chars = chars.replace('>>', '\\>\\>')
        chars = chars.replace('<<<', '\\<\\<\\<')
        chars = chars.replace('<<', '\\<\\<')
        # chars = chars.replace('__', '\\_\\_')
        chars = re.sub('@([^@]{1,66})@', '\\\\@\\1\\\\@', chars)
        chars = re.sub('_([\\$=\\.#\\-~\\+]*)_', '\\\\_\\1\\\\_', chars)
        chars = re.sub('\\^([0-9a-zA-Z]+)\\^', '\\\\^\\1\\\\^', chars)
        chars = re.sub('_([0-9a-zA-Z]+)_', '\\\\_\\1\\\\_', chars)
        chars = chars.replace('->', '\\->')
        chars = chars.replace('<-', '\\<-')
        chars = chars.replace('+>', '\\+>')
        chars = chars.replace('<+', '\\<+')
        chars = chars.replace('\\[', '\\[')
        chars = chars.replace('\\]', '\\]')
        chars = chars.replace('{{', '\\{{')
        chars = chars.replace('}}', '\\}}')
        chars = chars.replace('&lt;', '\\&lt;')
        chars = chars.replace('&gt;', '\\&gt;')
        chars = chars.replace('\\&lt-;', '\\&lt;\\-')  # "<-"
        chars = chars.replace('-\\&gt;', '\\-\\&gt;')  # "->"
        chars = chars.replace('\\&lt;+', '\\&lt;\\+')  # "<+"
        chars = chars.replace('+\\&gt;', '\\+\\&gt;')  # "+>"
        # PAGE NUMBER
        if type == 'footer':
            if fldchar == 'begin':
                res = '^ ?(\\S*)\\s*\\\\\\\\\\\\\\* MERGEFORMAT ?$'
                if re.match(res, chars):
                    chars = re.sub(res, '\\1', chars)
                if re.match('^ ?PAGE ?$', chars, re.I):
                    chars = 'n'
                elif re.match('^ ?SECTIONPAGES ?$', chars, re.I):
                    # "SECTIONPAGES" IS NOT SUPPORTOD BY LIBREOFFICE
                    chars = 'N'
                elif re.match('^ ?NUMPAGES ?$', chars, re.I):
                    chars = 'M'
            else:
                chars = re.sub('(n|N|M)', '\\\\\\1', chars)
        # RETURN
        return chars

    @classmethod
    def _get_raw_text(cls, chars_data):
        raw_text = cls._join_data(chars_data)
        raw_text = cls._escape_symbols(raw_text)
        raw_text = cls._convert_ivs(raw_text)  # IDEOGRAPHIC VARIATION SEQUENCE
        raw_text = cls._restore_charcters(raw_text)
        raw_text = cls._shrink_meaningless_font_decorations(raw_text)
        # RUBY
        res = '\\^<([^<>]{1,37})>/<([^<>]{1,37})>\\$'
        raw_text = re.sub(res, '<\\2/\\1>', raw_text)
        # SPACE
        res = NOT_ESCAPED + '<((?:[0-9]*\\.)?[0-9]+)>' * 2 + '((?:.|\n)*)$'
        while re.match(res, raw_text):
            head_text = re.sub(res, '\\1', raw_text)
            num1_text = re.sub(res, '\\2', raw_text)
            num2_text = re.sub(res, '\\3', raw_text)
            tail_text = re.sub(res, '\\4', raw_text)
            numb_text = str(round(float(num1_text) + float(num2_text), 2))
            raw_text = head_text + '<' + numb_text + '>' + tail_text
        for i in range(1, 6):
            j = str(round(0.6 * i, 1))
            res = NOT_ESCAPED + '<<<<><' + j + '><>>>>' + '((?:.|\n)*)$'
            while re.match(res, raw_text):
                raw_text = re.sub(res, '\\1' + '\u3000' * i + '\\2', raw_text)
            j = str(round(0.8 * i, 1))
            res = NOT_ESCAPED + '<<<><' + j + '><>>>' + '((?:.|\n)*)$'
            while re.match(res, raw_text):
                raw_text = re.sub(res, '\\1' + '\u3000' * i + '\\2', raw_text)
            j = str(round(1.2 * i, 1))
            res = NOT_ESCAPED + '>><' + j + '><<' + '((?:.|\n)*)$'
            while re.match(res, raw_text):
                raw_text = re.sub(res, '\\1' + '\u3000' * i + '\\2', raw_text)
            j = str(round(1.4 * i, 1))
            res = NOT_ESCAPED + '>>><' + j + '><<<' + '((?:.|\n)*)$'
            while re.match(res, raw_text):
                raw_text = re.sub(res, '\\1' + '\u3000' * i + '\\2', raw_text)
        # self.raw_text = raw_text
        return raw_text

    @classmethod
    def _join_data(cls, chars_data):
        chars_data = cls._cancel_fds(chars_data)
        raw_text = ''
        for cd in chars_data:
            cwf = cd.get_chars_with_fd()
            raw_text = CharsDatum._concatenate_chars(raw_text, cwf)
        return raw_text

    @classmethod
    def _cancel_fds(cls, chars_data):
        for i, cd in enumerate(chars_data):
            if i < len(chars_data) - 1:
                j = i + 1
                chars_data[i], chars_data[j] \
                    = CharsDatum.cancel_fds(chars_data[i], chars_data[j])
            if (cd.chars == '\n') and (i > 0) and (i < len(chars_data) - 1):
                j, k = i - 1, i + 1
                chars_data[j], chars_data[k] \
                    = CharsDatum.cancel_fds(chars_data[j], chars_data[k])
        return chars_data

    @staticmethod
    def _escape_symbols(raw_text):
        # SPACE
        raw_text = re.sub('(\n)([ \t\u3000]+)', '\\1\\\\\\2', raw_text)
        raw_text = re.sub('([ \t\u3000]+)(\n)', '\\1\\\\\\2', raw_text)
        # LENGTH REVISER
        if re.match('^(v|V|X|<<|<|>)=\\s*[0-9]+', raw_text):
            raw_text = '\\' + raw_text
        # REMARKS
        if re.match('^&quot;&quot;(\\s|$)', raw_text):
            raw_text = '\\' + raw_text
        if re.match('^""(\\s|$)', raw_text):
            raw_text = '\\' + raw_text
        # CHAPTER AND SECTION
        if re.match('^(\\$+(\\-\\$)*|#+(\\-#)*)=[0-9]+(\\s|$)', raw_text):
            raw_text = '\\' + raw_text
        if re.match('^(\\$+(\\-\\$)*|#+(\\-#)*)(\\s|$)', raw_text):
            raw_text = '\\' + raw_text
        # LIST
        if re.match('^(\\-|\\+|[0-9]+\\.|[0-9]+\\))\\s+', raw_text):
            raw_text = '\\' + raw_text
        # TABLE
        if re.match('^\\|((.|\n)*)\\|$', raw_text):
            raw_text = re.sub('^\\|((.|\n)*)\\|$', '\\\\|\\1\\\\|', raw_text)
        # IMAGE
        if re.match('(.|\n)*(' + RES_IMAGE + ')', raw_text):
            raw_text = re.sub('(' + RES_IMAGE + ')', '\\\\\\1', raw_text)
        if re.match('(.|\n)*<>\\\\(' + RES_IMAGE + ')', raw_text):
            raw_text = re.sub('<>\\\\(' + RES_IMAGE + ')', '\\1', raw_text)
        # ALIGNMENT
        res = '^:(\\s*(.|\n)*\\s*):$'
        if re.match(res, raw_text):
            raw_text = re.sub(res, '\\\\:\\1\\\\:', raw_text)
        if re.match('^:(\\s*(.|\n)*)$', raw_text):
            raw_text = re.sub('^:(\\s*(.|\n)*)$', '\\\\:\\1', raw_text)
        if re.match('^((.|\n)*\\s*):$', raw_text):
            raw_text = re.sub('^((.|\n)*\\s*):$', '\\1\\\\:', raw_text)
        # PREFORMATTED
        res = '^```((.|\n)*)```$'
        if re.match(res, raw_text):
            raw_text = re.sub(res, '\\\\```\\1\\\\```', raw_text)
        # PAGEBREAK
        if re.match('^<pgbr>$', raw_text):
            raw_text = '\\' + raw_text
        # HORIZONTAL LINE
        if re.match('^((\\s*-\\s*)|(\\s*\\*\\s*)){3,}$', raw_text):
            raw_text = '\\' + raw_text
        return raw_text

    # IVS (IDEOGRAPHIC VARIATION SEQUENCE)
    @staticmethod
    def _convert_ivs(raw_text):
        ivs_font = Form.ivs_font
        res = '^(.*[^\\\\0-9])([0-9]+);'
        while re.match(res, raw_text, flags=re.DOTALL):
            raw_text = re.sub(res, '\\1\\\\\\2;', raw_text, flags=re.DOTALL)
        ivs_beg = int('0xE0100', 16)
        ivs_end = int('0xE01EF', 16)
        #
        res = '^(.*)(@' + ivs_font + '@)' + \
            '(.[' + chr(ivs_beg) + '-' + chr(ivs_end) + '])' + \
            '(.*)$'
        while re.match(res, raw_text):
            raw_text = re.sub(res, '\\1\\3\\2\\4', raw_text)
        #
        res = '@' + ivs_font + '@@' + ivs_font + '@'
        raw_text = re.sub(res, '', raw_text)
        #
        res = '^(.*)(.)([' + chr(ivs_beg) + '-' + chr(ivs_end) + '])(.*)$'
        while re.match(res, raw_text, flags=re.DOTALL):
            t1 = re.sub(res, '\\1', raw_text, flags=re.DOTALL)
            t2 = re.sub(res, '\\2', raw_text, flags=re.DOTALL)
            t3 = re.sub(res, '\\3', raw_text, flags=re.DOTALL)
            t4 = re.sub(res, '\\4', raw_text, flags=re.DOTALL)
            ivs_n = ord(t3) - ivs_beg
            raw_text = t1 + t2 + str(ivs_n) + ';' + t4
        return raw_text

    @staticmethod
    def _restore_charcters(raw_text):
        raw_text = raw_text.replace('&lt;', '<')
        raw_text = raw_text.replace('&gt;', '>')
        raw_text = raw_text.replace('&quot;', '"')
        raw_text = raw_text.replace('&amp;', '&')
        return raw_text

    @staticmethod
    def _shrink_meaningless_font_decorations(raw_text):
        tmp_text = ''
        while tmp_text != raw_text:
            tmp_text = raw_text
            for fd in FONT_DECORATORS_INVISIBLE:
                res = '((?:\\s|' + '|'.join(FONT_DECORATORS_VISIBLE) + ')+)'
                raw_text = re.sub(fd + res + fd, '\\1', raw_text)
                raw_text = re.sub('^(' + fd + ')' + res, '\\2\\1', raw_text)
                raw_text = re.sub(res + '(' + fd + ')$', '\\2\\1', raw_text)
        return raw_text

    @staticmethod
    def _separate_head_space(text, del_beg, del_end, ins_beg, ins_end):
        right = text
        res_sp = '^([ \t\u3000]+)((?:.|\n)*)$'
        res_db = '^(' + del_beg + ')((?:.|\n)*)$'
        res_de = '^(' + del_end + ')((?:.|\n)*)$'
        res_ix = '^(' + ins_beg + '|' + ins_end + ')((?:.|\n)*)$'
        res_ch = '^(.|\n)((?:.|\n)*)$'
        left = ''
        space = ''
        level_to_break = 0
        is_in_comment = False
        while level_to_break != 2:
            if re.match(res_sp, right):
                if is_in_comment:
                    if level_to_break == 1:
                        left += re.sub(res_sp, '\\1', right)
                else:
                    space += re.sub(res_sp, '\\1', right)
                right = re.sub(res_sp, '\\2', right)
            elif re.match(res_db, right):
                left += re.sub(res_db, '\\1', right)
                right = re.sub(res_db, '\\2', right)
                is_in_comment = True
            elif re.match(res_de, right):
                left += re.sub(res_de, '\\1', right)
                right = re.sub(res_de, '\\2', right)
                is_in_comment = False
            elif re.match(res_ix, right):
                left += re.sub(res_ix, '\\1', right)
                right = re.sub(res_ix, '\\2', right)
            elif is_in_comment:
                left += re.sub(res_ch, '\\1', right)
                right = re.sub(res_ch, '\\2', right)
                level_to_break = 1
            else:
                level_to_break = 2
        left = re.sub(del_beg + del_end, '', left)
        left = re.sub(ins_beg + ins_end, '', left)
        return space, left + right

    @classmethod
    def _get_raw_text_del(cls, raw_text):
        raw_text_del \
            = cls._get_raw_text_del_or_ins(raw_text,
                                           '\\+>', '<\\+', '->', '<-')
        return raw_text_del

    @classmethod
    def _get_raw_text_ins(cls, raw_text):
        raw_text_ins \
            = cls._get_raw_text_del_or_ins(raw_text,
                                           '->', '<-', '\\+>', '<\\+')
        return raw_text_ins

    @staticmethod
    def _get_raw_text_del_or_ins(raw_text,
                                 beg_erase, end_erase,
                                 beg_leave, end_leave):
        raw_text_erase = ''
        raw_text_leave = ''
        track_changes = ''
        in_to_erase = False
        for c in raw_text:
            if in_to_erase:
                raw_text_erase += c
                if re.match(NOT_ESCAPED + end_erase + '$', raw_text_erase):
                    in_to_erase = False
                raw_text_erase = re.sub(end_erase + '$', '', raw_text_erase)
            else:
                raw_text_leave += c
                if re.match(NOT_ESCAPED + beg_erase + '$', raw_text_leave):
                    in_to_erase = True
                raw_text_leave = re.sub(beg_erase + '$', '', raw_text_leave)
                raw_text_leave = re.sub(beg_leave + '$', '', raw_text_leave)
                raw_text_leave = re.sub(end_leave + '$', '', raw_text_leave)
        return raw_text_leave

    @staticmethod
    def _get_remarks(xml_lines):
        remarks = []
        for xl in xml_lines:
            res = '^<w:commentReference w:id="(.*)"/>$'
            if re.match(res, xl):
                remark_id = re.sub(res, '\\1', xl)
                remarks.append(Form.remarks[remark_id])
        return remarks

    @staticmethod
    def _get_style(xml_lines):
        style = None
        for xl in xml_lines:
            style = XML.get_value('w:pStyle', 'w:val', style, xl)
        for ds in Form.styles:
            if style != ds.name:
                continue
            # REMOVED 23.02.18 >
            # self.alignment = ds.alignment
            # for s in self.length:
            #     if ds.raw_length[s] is not None:
            #         self.length[s] = ds.raw_length[s]
            # <
        # self.style = style
        return style

    @staticmethod
    def _get_alignment(xml_lines):
        alignment = ''
        for xl in xml_lines:
            alignment = XML.get_value('w:jc', 'w:val', alignment, xl)
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
        elif ParagraphMath.is_this_class(self):
            return 'math'
        elif ParagraphAlignment.is_this_class(self):
            return 'alignment'
        elif ParagraphPreformatted.is_this_class(self):
            return 'preformatted'
        elif ParagraphHorizontalLine.is_this_class(self):
            return 'horizontalline'
        elif ParagraphPagebreak.is_this_class(self):
            return 'pagebreak'
        elif ParagraphBreakdown.is_this_class(self):
            return 'breakdown'
        elif ParagraphRemarks.is_this_class(self):
            return 'remarks'
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
        elif paragraph_class == 'math':
            return ParagraphMath(self)
        elif paragraph_class == 'alignment':
            return ParagraphAlignment(self)
        elif paragraph_class == 'preformatted':
            return ParagraphPreformatted(self)
        elif paragraph_class == 'horizontalline':
            return ParagraphHorizontalLine(self)
        elif paragraph_class == 'pagebreak':
            return ParagraphPagebreak(self)
        elif paragraph_class == 'breakdown':
            return ParagraphBreakdown(self)
        elif paragraph_class == 'remarks':
            return ParagraphRemarks(self)
        else:
            return ParagraphSentence(self)


class Paragraph:

    """A class to handle paragraph"""

    paragraph_number = 0

    paragraph_class = None

    previous_head_section_depth = 0
    previous_tail_section_depth = 0

    @classmethod
    def is_this_class(cls, raw_paragraph):
        # rp = raw_paragraph
        # rp_xls = rp.xml_lines
        # rp_rcl = rp.raw_class
        # rp_rtx = rp.raw_text_doi
        # rp_img = rp.images
        # rp_sty = rp.style
        # rp_alg = rp.alignment
        # rp_fsz = Document.font_size
        return False

    def __init__(self, raw_paragraph):
        # RECEIVED
        self.raw_paragraph_number = raw_paragraph.raw_paragraph_number
        self.has_removed = raw_paragraph.has_removed
        self.xml_lines = raw_paragraph.xml_lines
        self.raw_class = raw_paragraph.raw_class
        self.horizontal_line = raw_paragraph.horizontal_line
        self.attached_pagebreak = raw_paragraph.attached_pagebreak
        self.chars_data = raw_paragraph.chars_data
        self.raw_text = raw_paragraph.raw_text
        self.head_space = raw_paragraph.head_space
        self.tail_space = raw_paragraph.tail_space
        self.raw_text_del = raw_paragraph.raw_text_del
        self.raw_text_ins = raw_paragraph.raw_text_ins
        self.raw_text_doi = raw_paragraph.raw_text_doi
        self.images = raw_paragraph.images
        self.remarks = raw_paragraph.remarks
        self.style = raw_paragraph.style
        self.alignment = raw_paragraph.alignment
        self.paragraph_class = raw_paragraph.paragraph_class
        # DECLARATION
        self.paragraph_number = -1
        self.head_section_depth = -1
        self.tail_section_depth = -1
        self.proper_depth = -1
        self.numbering_revisers = []
        self.head_font_revisers = []
        self.tail_font_revisers = []
        self.md_text = ''
        self.section_states = []
        self.length_docx = {}
        self.length_clas = {}
        self.length_conf = {}
        self.length_supp = {}
        self.length_revi = {}
        self.length_revisers = []
        self.pre_text_to_write = ''
        self.post_text_to_write = ''
        self.text_to_write_with_reviser = ''
        # SUBSTITUTION
        Paragraph.paragraph_number += 1
        self.paragraph_number = Paragraph.paragraph_number
        self.head_section_depth, self.tail_section_depth \
            = self._get_section_depths(self.raw_text_doi, not self.has_removed)
        self.proper_depth = self._get_proper_depth(self.raw_text_doi)
        self.raw_text = self._remove_track_change_at_head(self.raw_text)
        self.numbering_revisers, \
            self.head_font_revisers, \
            self.tail_font_revisers, \
            self.md_text \
            = self._get_revisers_and_md_text(self.raw_text)
        ParagraphList.reset_states(self.paragraph_class)
        self.section_states = self._get_section_states()
        self.length_docx = self._get_length_docx()
        self.length_clas = self._get_length_clas()
        self.length_conf = self._get_length_conf()
        self.length_supp = self._get_length_supp()
        self.length_revi = self._get_length_revi()
        self.length_revisers = self._get_length_revisers(self.length_revi)
        self.section_states, self.numbering_revisers, self.length_revisers \
            = self._revise_for_section_depth_2(self.paragraph_class,
                                               self.head_section_depth,
                                               self.tail_section_depth,
                                               self.section_states,
                                               self.numbering_revisers,
                                               self.length_revisers)
        # EXECUTION
        self.md_lines_text = self._get_md_lines_text(self.md_text)
        self.text_to_write = self.get_text_to_write()
        self.text_to_write_with_reviser \
            = self.get_text_to_write_with_reviser()

    @classmethod
    def _get_section_depths(cls, raw_text, should_record=False):
        head_section_depth = 0
        tail_section_depth = 0
        # self.head_section_depth = head_section_depth
        # self.tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    @staticmethod
    def _get_section_states():
        ss = ParagraphSection.states
        states \
            = [[ss[i][j] for j in range(len(ss[i]))] for i in range(len(ss))]
        return states

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
        # PREFORMATTED
        if self.paragraph_class == 'preformatted':
            if '`' in head_font_revisers:
                head_font_revisers.remove('`')
            else:
                head_font_revisers.append('`')
            if '`' in tail_font_revisers:
                tail_font_revisers.remove('`')
            else:
                tail_font_revisers.append('`')
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            md_text

    @staticmethod
    def _get_font_revisers_and_md_text(raw_text):
        head_font_revisers = []
        tail_font_revisers = []
        while True:
            for fd in FONT_DECORATORS:
                res = '^(' + fd + ')((?:.|\n)*)$'
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
                res = '^((?:.|\n)*)(' + fd + ')$'
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
            if Form.document_style == 'j':
                if xdepth == 2:
                    cvalue += 1
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

    def _remove_track_change_at_head(self, raw_text):
        pc = self.paragraph_class
        if pc != 'chapter' and pc != 'section' and pc != 'list':
            return raw_text
        head_text = ''
        track_changes = ''
        tmp_text = raw_text
        for i in range(len(raw_text)):
            if re.match(NOT_ESCAPED + '\\->$', raw_text[:i + 1]):
                head_text = head_text[:-1]
                track_changes = 'del'
                continue
            if re.match(NOT_ESCAPED + '<\\-$', raw_text[:i + 1]):
                # head_text = head_text[:-1]
                track_changes = ''
                continue
            if re.match(NOT_ESCAPED + '\\+>$', raw_text[:i + 1]):
                head_text = head_text[:-1]
                track_changes = 'ins'
                continue
            if re.match(NOT_ESCAPED + '<\\+$', raw_text[:i + 1]):
                head_text = head_text[:-1]
                track_changes = ''
                continue
            if track_changes == 'del':
                continue
            head_text += raw_text[i]
            # ParagraphChapter.res_separator
            # ParagraphSection.r9
            # ParagraphList.res_separator
            if re.match('^.*(?:  ?|\t|\u3000|\\. |．)$', head_text):
                tmp_text = head_text
                if track_changes == 'del':
                    tmp_text += '->'
                elif track_changes == 'ins':
                    tmp_text += '+>'
                if i < len(raw_text) - 1:
                    tmp_text += raw_text[i + 1:]
                while re.match(NOT_ESCAPED + '\\-><\\-', tmp_text):
                    tmp_text \
                        = re.sub(NOT_ESCAPED + '\\-><\\-', '\\1', tmp_text)
                while re.match(NOT_ESCAPED + '\\+><\\+', tmp_text):
                    tmp_text \
                        = re.sub(NOT_ESCAPED + '\\+><\\+', '\\1', tmp_text)
                break
        raw_text = tmp_text
        return raw_text

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
        m_size = Form.font_size
        lnsp = Form.line_spacing
        xls = self.xml_lines
        head_space = self.head_space
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
        is_changed = False
        for xl in xls:
            if re.match('^<w:pPrChange( .*[^/])?>$', xl):
                is_changed = True
            if re.match('^</w:pPrChange( .*[^/])?>$', xl):
                is_changed = False
            if is_changed:
                continue
            sb_xml = XML.get_value('w:spacing', 'w:before', sb_xml, xl)
            sa_xml = XML.get_value('w:spacing', 'w:after', sa_xml, xl)
            ls_xml = XML.get_value('w:spacing', 'w:line', ls_xml, xl)
            fi_xml = XML.get_value('w:ind', 'w:firstLine', fi_xml, xl)
            hi_xml = XML.get_value('w:ind', 'w:hanging', hi_xml, xl)
            li_xml = XML.get_value('w:ind', 'w:left', li_xml, xl)
            ri_xml = XML.get_value('w:ind', 'w:right', ri_xml, xl)
            ti_xml = XML.get_value('w:tblInd', 'w:w', ti_xml, xl)
        length_docx['space before'] = round(sb_xml / 20 / m_size / lnsp, 2)
        length_docx['space after'] = round(sa_xml / 20 / m_size / lnsp, 2)
        ls = 0.0
        if ls_xml > 0.0:
            ls = (ls_xml / 20 / m_size / lnsp) - 1
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
        length_docx['first indent'] = round((fi_xml - hi_xml) / 20 / m_size, 2)
        # （１）, （ア）, （ａ）
        paragraph_class = self.paragraph_class
        raw_text = self.raw_text
        res = '^（([0-9０-９]+|[ｱ-ﾝア-ン]+|[a-zａ-ｚ]+)）'
        if paragraph_class == 'section':
            if re.match(res, raw_text):
                length_docx['first indent'] += 1.0
        length_docx['left indent'] = round((li_xml + ti_xml) / 20 / m_size, 2)
        length_docx['right indent'] = round(ri_xml / 20 / m_size, 2)
        if head_space != '':
            width = 0
            for sp in head_space:
                if sp == ' ':
                    width += 0.5
                elif sp == '\t':
                    width = (int(width / 4) + 1) * 4.0
                else:
                    width += 1.0
            length_docx['first indent'] += width
        # self.length_docx = length_docx
        return length_docx

    def _get_length_clas(self):
        paragraph_class = self.paragraph_class
        head_section_depth = self.head_section_depth
        tail_section_depth = self.tail_section_depth
        section_states = self.section_states
        proper_depth = self.proper_depth
        length_clas \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        if paragraph_class == 'chapter':
            length_clas['first indent'] = -1.0
            length_clas['left indent'] = proper_depth + 0.0
        elif paragraph_class == 'section':
            if head_section_depth > 1:
                length_clas['first indent'] \
                    = head_section_depth - tail_section_depth - 1.0
            if tail_section_depth > 1:
                length_clas['left indent'] = tail_section_depth - 1.0
        elif paragraph_class == 'list':
            length_clas['first indent'] = -1.0
            length_clas['left indent'] = proper_depth + 0.0
            if tail_section_depth > 0:
                length_clas['left indent'] += tail_section_depth - 1.0
        elif paragraph_class == 'table':
            length_clas['space before'] += TABLE_SPACE_BEFORE
            length_clas['space after'] += TABLE_SPACE_AFTER
        elif paragraph_class == 'image':
            length_clas['space before'] += IMAGE_SPACE_BEFORE
            length_clas['space after'] += IMAGE_SPACE_AFTER
        elif paragraph_class == 'preformatted':
            if tail_section_depth > 0:
                length_clas['first indent'] = 0.0
                length_clas['left indent'] = tail_section_depth - 0.0
        elif paragraph_class == 'sentence':
            if tail_section_depth > 0:
                length_clas['first indent'] = 1.0
                length_clas['left indent'] = tail_section_depth - 1.0
        if paragraph_class == 'section' or \
           paragraph_class == 'list' or \
           paragraph_class == 'preformatted' or \
           paragraph_class == 'sentence':
            if section_states[1][0] == 0 and tail_section_depth > 2:
                length_clas['left indent'] -= 1.0
        if Form.document_style == 'j':
            if section_states[1][0] > 0 and tail_section_depth > 2:
                length_clas['left indent'] -= 1.0
        # self.length_clas = length_clas
        return length_clas

    def _get_length_conf(self):
        hd = self.head_section_depth
        td = self.tail_section_depth
        length_conf \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        if self.paragraph_class == 'section':
            sb = (Form.space_before + ',,,,,,,').split(',')
            sa = (Form.space_after + ',,,,,,,').split(',')
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
        length_clas = self.length_clas
        length_revi \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        for ln in length_revi:
            lg = length_docx[ln] - length_clas[ln] \
                - length_conf[ln] + length_supp[ln]
            length_revi[ln] = round(lg, 2)
        # self.length_revi = length_revi
        return length_revi

    @classmethod
    def _get_length_revisers(cls, length_revi):
        length_revisers = []
        vs = cls._get_vlength_string(length_revi['space before'])
        if float(vs) < -0.05 or float(vs) > 0.05:
            length_revisers.append('v=' + vs)
        vs = cls._get_vlength_string(length_revi['space after'])
        if float(vs) < -0.05 or float(vs) > 0.05:
            length_revisers.append('V=' + vs)
        vs = cls._get_vlength_string(length_revi['line spacing'])
        if float(vs) < -0.05 or float(vs) > 0.05:
            length_revisers.append('X=' + vs)
        hs = cls._get_hlength_string(- length_revi['first indent'])
        if float(hs) < -0.05 or float(hs) > 0.05:
            length_revisers.append('<<=' + hs)
        hs = cls._get_hlength_string(- length_revi['left indent'])
        if float(hs) < -0.05 or float(hs) > 0.05:
            length_revisers.append('<=' + hs)
        hs = cls._get_hlength_string(- length_revi['right indent'])
        if float(hs) < -0.05 or float(hs) > 0.05:
            length_revisers.append('>=' + hs)
        # self.length_revisers = length_revisers
        return length_revisers

    @staticmethod
    def _get_vlength_string(length):
        # FRACTION
        if length < 0:
            porm = '-'
        elif length == 0:
            porm = ''
        else:
            porm = '+'
        i_part = str(int(abs(length)))
        d_part = abs(length - int(length))
        if d_part > 0.329 and d_part < 0.340:
            return porm + i_part + '.33'  # 1/3=0.3333...
        if d_part > 0.660 and d_part < 0.671:
            return porm + i_part + '.67'  # 2/3=0.6666...
        if d_part > 0.245 and d_part < 0.255:
            return porm + i_part + '.25'  # 1/4=0.25
        if d_part > 0.745 and d_part < 0.755:
            return porm + i_part + '.75'  # 3/4=0.75
        if d_part > 0.160 and d_part < 0.171:
            return porm + i_part + '.17'  # 1/6=0.1666...
        if d_part > 0.829 and d_part < 0.840:
            return porm + i_part + '.83'  # 5/6=0.8333...
        # DECIMAL
        rounded = round(length, 1)
        if rounded < 0:
            return str(rounded)
        elif rounded == 0:
            return '0.0'
        else:
            return '+' + str(rounded)

    @staticmethod
    def _get_hlength_string(length):
        rounded = round(length * 2) / 2  # half-width units
        if rounded < 0:
            return str(rounded)
        if rounded == 0:
            return '0.0'
        else:
            return '+' + str(rounded)

    def _revise_for_section_depth_2(self,
                                    paragraph_class,
                                    head_section_depth, tail_section_depth,
                                    section_states,
                                    numbering_revisers, length_revisers):
        if paragraph_class == 'section':
            if head_section_depth == 3 and tail_section_depth == 3:
                if section_states[1][0] > 0:
                    if section_states[2][0] == 1 and section_states[2][1] == 0:
                        if '##=1' not in numbering_revisers:
                            if '<=+1.0' in length_revisers:
                                ParagraphSection.states[1][0] = 0
                                section_states[1][0] = 0
                                numbering_revisers.insert(0, '##=1')
                                length_revisers.remove('<=+1.0')
        return section_states, numbering_revisers, length_revisers

    def _get_md_lines_text(self, md_text):
        paragraph_class = self.paragraph_class
        # FOR TRAILING WHITE SPACE
        md_text = re.sub('  \n', '  \\\n', md_text)
        if False:
            pass
        # elif paragraph_class == 'chapter':
        #     md_lines_text = Paragraph._split_into_lines(md_text)
        elif paragraph_class == 'section':
            md_lines_text = Paragraph._split_into_lines(md_text)
        # elif paragraph_class == 'list':
        #     md_lines_text = Paragraph._split_into_lines(md_text)
        elif paragraph_class == 'sentence':
            md_lines_text = Paragraph._split_into_lines(md_text)
        else:
            md_lines_text = md_text
        return md_lines_text

    def get_text_to_write(self):
        paper_size = Form.paper_size
        top_margin = Form.top_margin
        bottom_margin = Form.bottom_margin
        left_margin = Form.left_margin
        right_margin = Form.right_margin
        md_lines_text = self.md_lines_text
        length_docx = self.length_docx
        indent = length_docx['first indent'] \
            + length_docx['left indent'] \
            + length_docx['right indent']
        unit = 12 * 2.54 / 72 / 2
        width_cm = PAPER_WIDTH[paper_size] - left_margin - right_margin \
            - (indent * unit)
        height_cm = PAPER_HEIGHT[paper_size] - top_margin - bottom_margin
        region_cm = (width_cm, height_cm)
        res = '^((?:.|\n)*)(' + RES_IMAGE_WITH_SIZE + ')((?:.|\n)*)$'
        text_to_write = ''
        while re.match(res, md_lines_text):
            text_to_write += re.sub(res, '\\1', md_lines_text)
            img_text = re.sub(res, '\\2', md_lines_text)
            text_to_write \
                += ParagraphImage.replace_with_fixed_size(img_text, region_cm)
            md_lines_text = re.sub(res, '\\7', md_lines_text)
        text_to_write += md_lines_text
        # self.text_to_write = text_to_write
        return text_to_write

    def get_text_to_write_with_reviser(self):
        numbering_revisers = self.numbering_revisers
        length_revisers = self.length_revisers
        head_font_revisers = self.head_font_revisers
        tail_font_revisers = self.tail_font_revisers
        text_to_write = self.text_to_write
        pre_text_to_write = self.pre_text_to_write
        post_text_to_write = self.post_text_to_write
        paragraph_class = self.paragraph_class
        attached_pagebreak = self.attached_pagebreak
        # LEFT SYMBOL
        has_left_sharp = False
        has_left_colon = False
        if re.match('^# (.|\n)*$', text_to_write):
            text_to_write = re.sub('^# ', '', text_to_write)
            has_left_sharp = True
        elif re.match('^: (.|\n)*$', text_to_write):
            text_to_write = re.sub('^: ', '', text_to_write)
            has_left_colon = True
        # RIGHT SYMBOL
        has_right_colon = False
        if re.match('^(.|\n)* :$', text_to_write):
            text_to_write = re.sub(' :$', '', text_to_write)
            has_right_colon = True
        ttwwr = ''
        if pre_text_to_write != '':
            ttwwr += pre_text_to_write + '\n'
        for rev in numbering_revisers:
            ttwwr += rev + ' '
        if re.match('^(.|\n)* $', ttwwr):
            ttwwr = re.sub(' $', '\n', ttwwr)
        for rev in length_revisers:
            ttwwr += rev + ' '
        if re.match('^(.|\n)* $', ttwwr):
            ttwwr = re.sub(' $', '\n', ttwwr)
        # LEFT SYMBOL
        if has_left_sharp:
            ttwwr += '# '
        elif has_left_colon:
            ttwwr += ': '
        for rev in head_font_revisers:
            ttwwr += rev
        ttwwr += text_to_write
        for rev in reversed(tail_font_revisers):
            ttwwr += rev
        # RIGHT SYMBOL
        if has_right_colon:
            ttwwr += ' :'
        if post_text_to_write != '':
            ttwwr += '\n' + post_text_to_write
        # PAGE BREAK
        if paragraph_class != 'pagebreak':
            if attached_pagebreak == 'pgbr':
                ttwwr += '\n\n<pgbr>'
            if attached_pagebreak == 'Pgbr':
                ttwwr += '\n\n<Pgbr>'
        text_to_write_with_reviser = ttwwr
        # self.text_to_write_with_reviser = text_to_write_with_reviser
        return text_to_write_with_reviser

    @classmethod
    def _split_into_lines(cls, md_text):
        md_lines_text = ''
        for line in md_text.split('\n'):
            res = NOT_ESCAPED + '(' + RES_IMAGE + ')(.*)$'
            line = re.sub(res, '\\1\n\\2\n\\5', line)
            line = re.sub('\n+', '\n', line)
            phrases = []
            for text in line.split('\n'):
                if re.match(RES_IMAGE, text):
                    phrases.append(text)
                else:
                    phrases += cls._split_into_phrases(text)
            splited = cls._concatenate_phrases(phrases)
            md_lines_text += splited + '<br>\n'
        md_lines_text = re.sub('<br>\n$', '', md_lines_text)
        return md_lines_text

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
            # '->' or '<-' or '+>' or '<+' (TRACK CHANGES)
            res = '(?:\\->|<\\-|\\+>|<\\+)'
            if re.match(NOT_ESCAPED + 'x$', tmp + 'x') and \
               re.match('^' + res + '.*$', tmp2):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            if re.match('^.*' + res + '$', tmp):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            # '\[' or '\]' (MATH MODE)
            res = '(?:\\\\\\[|\\\\\\])'
            if re.match(NOT_ESCAPED + 'x$', tmp + 'x') and \
               re.match('^' + res + '.*$', tmp2):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
            if re.match('^.*' + res + '$', tmp):
                if tmp != '':
                    phrases.append(tmp)
                    tmp = ''
        return phrases

    @staticmethod
    def _concatenate_phrases(phrases):
        def _extend_tex(extension):
            # JUST TO MAKE SURE
            if extension == '':
                return tex
            if is_in_deleted:
                return tex + '->' + extension + '<-\n'
            if is_in_inserted:
                return tex + '+>' + extension + '<+\n'
            return tex + extension + '\n'
        tex = ''
        tmp = ''
        is_in_deleted = False
        is_in_inserted = False
        is_in_math = False
        for p in phrases:
            # MATH MODE (MUST BE FIRST)
            if (not is_in_math) and p == '\\[':
                tex = _extend_tex(tmp)
                tmp = ''
                is_in_math = True
                continue
            if is_in_math and p == '\\]':
                tex = _extend_tex('\\[' + tmp + '\\]')
                tmp = ''
                is_in_math = False
                continue
            if is_in_math:
                tmp += p
                continue
            # DELETED
            if (not is_in_deleted) and p == '->':
                tex = _extend_tex(tmp)
                tmp = ''
                is_in_deleted = True
                continue
            if is_in_deleted and p == '<-':
                tex = _extend_tex(tmp)
                tmp = ''
                is_in_deleted = False
                continue
            # INSERTED
            if (not is_in_inserted) and p == '+>':
                tex = _extend_tex(tmp)
                tmp = ''
                is_in_inserted = True
                continue
            if is_in_inserted and p == '<+':
                tex = _extend_tex(tmp)
                tmp = ''
                is_in_inserted = False
                continue
            # SECTION WITHOUT A TITLE
            res = '(?:#+(?:\\-#)* )+'
            if tex == '':
                if re.match('^' + res + '$', tmp):
                    if not re.match('^' + res + '.*$', p):
                        if re.match('^.*[.．。]$', phrases[-1]):
                            tex = _extend_tex(re.sub('\\s+$', '', tmp))
                            # tex = _extend_tex(tmp + '\\')
                            tmp = ''
            # IMAGE
            if re.match(RES_IMAGE, p):
                tex = _extend_tex(tmp + '\n' + p)
                tmp = ''
                continue
            # CONJUNCTIONS
            if re.match('^.*[,，、]$', tmp):
                for c in CONJUNCTIONS:
                    if re.match('^' + c + '[,，、]$', tmp):
                        tex = _extend_tex(tmp)
                        tmp = ''
                        break
            # END OF A SENTENCE
            if re.match('^.*[．。]$', tmp):
                tex = _extend_tex(tmp)
                tmp = ''
            # RIGTH LENGTH
            if tmp != '':
                if get_ideal_width(tmp + p) > MD_TEXT_WIDTH:
                    tex = _extend_tex(tmp)
                    tmp = ''
            # TOO LONG
            tmp += p
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
                            tex = _extend_tex(s1)
                            tmp = s2
                            break
                    if re.match('^.*[ぁ-ん，、．。]$', s1) and \
                       re.match('^[^ぁ-ん，、．。].*$', s2):
                        if s1 != '':
                            tex = _extend_tex(s1)
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
                        # '_.*' + '.*_' (UNDERLINE)
                        if re.match('^.*_[\\$=\\.#\\-~\\+]*$', s1) and \
                           re.match('^[\\$=\\.#\\-~\\+]*_.*$', s2):
                            continue
                        # '^.*' + '.*^' (FONT COLOR)
                        if re.match('^.*\\^[0-9A-Za-z]*$', s1) and \
                           re.match('^[0-9A-Za-z]*\\^.*$', s2):
                            continue
                        # '_.+' + '.+_' (HIGHLIGHT COLOR)
                        if re.match('^.*_[0-9A-Za-z]+$', s1) and \
                           re.match('^[0-9A-Za-z]+_.*$', s2):
                            continue
                        # '@.+' + '.+@' (FONT)
                        if re.match('^.*@[^@]{1,66}$', s1) and \
                           re.match('^[^@]{1,66}@.*$', s2):
                            continue
                        # ' ' + ' ' (LINE BREAK)
                        if re.match('^.* $', s1) and re.match('^ .*$', s2):
                            continue
                        # '<' + '[-+]' (TRACK CHANGES)
                        if re.match('^.*<$', s1) and \
                           re.match('^[\\-\\+].*$', s2):
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
        tex = re.sub('(  |\t|\u3000)(\n)', '\\1\\\\\\2', tex)
        return tex

    def get_document(self):
        paragraph_class = self.paragraph_class
        remarks = self.remarks
        ttwwr = self.text_to_write_with_reviser
        dcmt = ''
        if paragraph_class != 'empty':
            if ttwwr != '':
                dcmt = ''
                for r in remarks:
                    dcmt += '"" ' + r + '\n'
                dcmt += ttwwr + '\n'
        return dcmt

    def get_images(self):
        return self.images


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
        rp_rtx = rp.raw_text_doi
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphImage.is_this_class(rp):
            return False
        if ParagraphPagebreak.is_this_class(rp):
            return False
        if ParagraphHorizontalLine.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        if re.match('^\\s*$', rp_rtx):
            return True
        return False


class ParagraphChapter(Paragraph):

    """A class to handle chapter paragraph"""

    paragraph_class = 'chapter'
    paragraph_class_ja = 'チャプター'

    res_branch = '((?:の[0-9０-９]+)*)'
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
    res_rest = '(.*\\S(?:.|\n)*)'
    states = [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１編
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１章
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１節
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],  # 第１款
              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]]  # 第１目

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rtx = rp.raw_text_doi
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        for i in range(len(cls.res_symbols)):
            res = '^(?:\\\\\\s+)?' \
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
        head_tc = ''
        tail_tc = ''
        if re.match('^->(.|\n)*$', raw_text):
            head_tc = '->'
            raw_text = re.sub('^->', '', raw_text)
        elif re.match('^\\+>(.|\n)*$', raw_text):
            head_tc = '+>'
            raw_text = re.sub('^\\+>', '', raw_text)
        if re.match('^(.|\n)*<-$', raw_text):
            tail_tc = '<-'
            raw_text = re.sub('<-$', '', raw_text)
        elif re.match('^(.|\n)*<\\+$', raw_text):
            tail_tc = '<+'
            raw_text = re.sub('<\\+$', '', raw_text)
        head_symbol = ''
        for xdepth in range(len(rss)):
            res = '^' + rss[xdepth] + rre + '$'
            if re.match(res, raw_text):
                head_string, raw_text, state \
                    = self._decompose_text(res, raw_text, -1, -1)
                ydepth = len(state) - 1
                if head_tc != '->':
                    self._step_states(xdepth, ydepth)
                    numbering_revisers \
                        = self._get_numbering_revisers(xdepth, state)
                head_symbol = '$' * (xdepth + 1) + '-$' * ydepth + ' '
                break
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            head_tc + head_symbol + raw_text + tail_tc

    @staticmethod
    def _decompose_text(res, raw_text, num1, num2):
        hdstr = re.sub(res, '\\1', raw_text)
        nmsym = re.sub(res, '\\2', raw_text)
        branc = re.sub(res, '\\3', raw_text)
        rtext = re.sub(res, '\\4', raw_text)
        state = []
        for b in branc.split('の'):
            state.append(c2n_n_arab(b) - 1)
        if re.match('[0-9０-９]+', nmsym):
            state[0] = c2n_n_arab(nmsym)
        return hdstr, rtext, state


class ParagraphSection(Paragraph):

    """A class to handle section paragraph"""

    paragraph_class = 'section'
    paragraph_class_ja = 'セクション'

    # r0 = '((?:' + '|'.join(FONT_DECORATORS) + ')*)'
    r1 = '\\+\\+\\+(.*)\\+\\+\\+'
    r2 = '(?:(第([0-9０-９]+)条?)((?:の[0-9０-９]+)*))'
    r3 = '(?:(([0-9０-９]+))((?:の[0-9０-９]+)*))'
    r4 = '(?:([⑴-⒇]|[\\(（]([0-9０-９]+)[\\)）])((?:の[0-9０-９]+)*))'
    r5 = '(?:((' + RES_KATAKANA + '))((?:の[0-9０-９]+)*))'
    r6 = '(?:([(\\(（](' + RES_KATAKANA + ')[\\)）])((?:の[0-9０-９]+)*))'
    r7 = '(?:(([a-zａ-ｚ]))((?:の[0-9０-９]+)*))'
    r8 = '(?:([⒜-⒵]|[(\\(（]([a-zａ-ｚ])[\\)）])((?:の[0-9０-９]+)*))'
    r9 = '(?:  ?|\t|\u3000|\\. |．)'
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
    res_rest = '(.*\\S(?:.|\n)*)'
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
        rp = raw_paragraph
        rp_rtx = rp.raw_text_doi
        alignment = rp.alignment
        head_section_depth, tail_section_depth \
            = cls._get_section_depths(rp_rtx)
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphMath.is_this_class(rp):
            return False
        if ParagraphImage.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        if tail_section_depth == 1 and alignment == 'center':
            return True
        elif tail_section_depth > 1:
            return True
        return False

    @classmethod
    def _get_section_depths(cls, raw_text, should_record=False):
        # （１）, （ア）, （ａ）
        raw_text = re.sub('^（([0-9０-９]+|[ｱ-ﾝア-ン]+|[a-zａ-ｚ]+)）',
                          '(\\1) ', raw_text)
        rss = cls.res_symbols
        rfd = RES_FONT_DECORATORS
        rre = cls.res_rest
        rnm = cls.res_number
        head_section_depth = 0
        tail_section_depth = 0
        for xdepth in range(1, len(rss)):
            res = '^(?:\\\\\\s+)?' + rfd + rss[xdepth] + rre + '$'
            if re.match(res, raw_text) and not re.match(rnm, raw_text):
                if head_section_depth == 0:
                    head_section_depth = xdepth + 1
                tail_section_depth = xdepth + 1
            if head_section_depth == 0 and tail_section_depth == 0:
                res = '^(?:\\\\\\s+)?' + rfd + rss[0] + rfd + '$'
                if re.match(res, raw_text):
                    head_section_depth = 1
                    tail_section_depth = 1
        if should_record:
            Paragraph.previous_head_section_depth = head_section_depth
            Paragraph.previous_tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    def _get_revisers_and_md_text(self, raw_text):
        # （１）, （ア）, （ａ）
        raw_text = re.sub('^（([0-9０-９]+|[ｱ-ﾝア-ン]+|[a-zａ-ｚ]+)）',
                          '(\\1) ', raw_text)
        m_size = Form.font_size
        xl_size = m_size * 1.4
        xml_lines = self.xml_lines
        rss = self.res_symbols
        rre = self.res_rest
        rnm = self.res_number
        numbering_revisers = []
        head_font_revisers, tail_font_revisers, raw_text \
            = Paragraph._get_font_revisers_and_md_text(raw_text)
        head_tc = ''
        tail_tc = ''
        if re.match('^->(.|\n)*$', raw_text):
            head_tc = '->'
            raw_text = re.sub('^->', '', raw_text)
        elif re.match('^\\+>(.|\n)*$', raw_text):
            head_tc = '+>'
            raw_text = re.sub('^\\+>', '', raw_text)
        if re.match('^(.|\n)*<-$', raw_text):
            tail_tc = '<-'
            raw_text = re.sub('<-$', '', raw_text)
        elif re.match('^(.|\n)*<\\+$', raw_text):
            tail_tc = '<+'
            raw_text = re.sub('<\\+$', '', raw_text)
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
                if head_tc != '->':
                    self._step_states(xdepth, ydepth)
                    numbering_revisers \
                        = self._get_numbering_revisers(xdepth, state)
                head_symbol += '#' * (xdepth + 1) + '-#' * ydepth + ' '
        raw_text = re.sub('^' + ParagraphSection.r9, '', raw_text)
        # raw_text = re.sub('^(?:  ?|\t|\u3000|\\. ?|．)', '', raw_text)
        if head_symbol == '':
            self._step_states(0, 0)
            if '+++' in head_font_revisers:
                head_font_revisers.remove('+++')
            if '+++' in tail_font_revisers:
                tail_font_revisers.remove('+++')
            for xl in xml_lines:
                s = XML.get_value('w:sz', 'w:val', -1.0, xl) / 2
                w = XML.get_value('w:w', 'w:val', -1.0, xl)
                if (s > 0 and s < xl_size * 0.7) or (w > 0 and w < 70):
                    head_font_revisers.insert(0, '---')
                    tail_font_revisers.intert(0, '---')
                    # raw_text = '---' + raw_text + '---'
                elif (s > 0 and s < xl_size * 0.9) or (w > 0 and w < 90):
                    head_font_revisers.insert(0, '--')
                    tail_font_revisers.insert(0, '--')
                    # raw_text = '--' + raw_text + '--'
                elif (s > 0 and s > xl_size * 1.3) or (w > 0 and w > 130):
                    head_font_revisers.insert(0, '+++')
                    tail_font_revisers.insert(0, '+++')
                    # raw_text = '+++' + raw_text + '+++'
                elif (s > 0 and s > xl_size * 1.1) or (w > 0 and w > 110):
                    head_font_revisers.insert(0, '++')
                    tail_font_revisers.insert(0, '++')
                    # raw_text = '++' + raw_text + '++'
                if s > 0 or w > 0:
                    break
            head_symbol = '# '
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            head_tc + head_symbol + raw_text + tail_tc

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
        # REVISE ⑴-⒇
        if re.match('^[⑴-⒇]', hdstr) and nmsym == '':
            nmsym = re.sub('^(.)(.|\n)*$', '\\1', hdstr)
        state = []
        if nmsym == '':
            nmsym = hdstr
        for b in branc.split('の'):
            state.append(c2n_n_arab(b) - 1)
        if nmsym == '':
            nmsym = hdstr
        if re.match('[0-9０-９]+', nmsym):
            state[0] = c2n_n_arab(nmsym)
        elif re.match('[⑴-⒇]', nmsym):
            state[0] = c2n_p_arab(nmsym)
        elif re.match(RES_KATAKANA, nmsym):
            state[0] = c2n_n_kata(nmsym)
        elif re.match('[a-zａ-ｚ]', nmsym):
            state[0] = c2n_n_alph(nmsym)
        elif re.match('[⒜-⒵]', nmsym):
            state[0] = c2n_p_alph(nmsym)
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
        xml_lines = rp.xml_lines
        res_xml_bullet_ms = cls.res_xml_bullet_ms
        res_xml_number_ms = cls.res_xml_number_ms
        res_xml_bullet_lo = cls.res_xml_bullet_lo
        res_xml_number_lo = cls.res_xml_number_lo
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        for xl in xml_lines:
            if re.match(res_xml_bullet_ms, xl):
                return True
            if re.match(res_xml_number_ms, xl):
                return True
            if re.match(res_xml_bullet_lo, xl):
                return True
            if re.match(res_xml_number_lo, xl):
                return True
        return False

    @classmethod
    def _get_section_depths(cls, raw_text, should_record=False):
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
        xml_lines = self.xml_lines
        raw_text = self.raw_text
        list_type = ''
        depth = 1
        for xl in xml_lines:
            if re.match(res_xml_bullet_ms, xl):
                n = re.sub(res_xml_bullet_ms, '\\1', xl)
                depth = int(n) + 1
            if re.match(res_xml_number_ms, xl):
                n = re.sub(res_xml_number_ms, '\\1', xl)
                if n == '10':
                    list_type = 'bullet'
                else:
                    list_type = 'number'
            if re.match(res_xml_bullet_lo, xl):
                list_type = 'bullet'
                n = re.sub(res_xml_bullet_lo, '\\1', xl)
                if n != '':
                    depth = int(n)
            if re.match(res_xml_number_lo, xl):
                list_type = 'number'
                n = re.sub(res_xml_number_lo, '\\1', xl)
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
        xml_lines = self.xml_lines
        raw_text = self.raw_text
        list_type = ''
        depth = 1
        for xl in xml_lines:
            if re.match(res_xml_bullet_ms, xl):
                n = re.sub(res_xml_bullet_ms, '\\1', xl)
                depth = int(n) + 1
            if re.match(res_xml_number_ms, xl):
                n = re.sub(res_xml_number_ms, '\\1', xl)
                if n == '10':
                    list_type = 'bullet'
                else:
                    list_type = 'number'
            if re.match(res_xml_bullet_lo, xl):
                list_type = 'bullet'
                n = re.sub(res_xml_bullet_lo, '\\1', xl)
                if n != '':
                    depth = int(n)
            if re.match(res_xml_number_lo, xl):
                list_type = 'number'
                n = re.sub(res_xml_number_lo, '\\1', xl)
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
    res_rest = '(.*\\S(?:.|\n)*)'
    states = [[0],  # ①
              [0],  # ㋐
              [0],  # ⓐ
              [0]]  # ㊀

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rtx = rp.raw_text_doi
        proper_depth = cls._get_proper_depth(rp_rtx)
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        if proper_depth > 0:
            return True
        return False

    @classmethod
    def _get_section_depths(cls, full_text, should_record=False):
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
            state = [c2n_c_arab(nmsym)]
        elif xdepth == 1:
            state = [c2n_c_kata(nmsym)]
        elif xdepth == 2:
            state = [c2n_c_alph(nmsym)]
        elif xdepth == 3:
            state = [c2n_c_kanj(nmsym)]
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
        if ParagraphConfiguration.is_this_class(rp):
            return False
        if rp_cls == 'w:tbl':
            return True
        return False

    def _get_md_text(self, raw_text):
        m_size = Form.font_size
        t_size = m_size * TABLE_FONT_SCALE
        xml_lines = self.xml_lines
        is_in_row = False
        is_in_cel = False
        geta_height, geta_width = 1.5, 1.5
        tab = []
        wid = []
        hei = []
        for xl in xml_lines:
            if re.match('^<w:gridCol(?: .*)?/>$', xl):
                res = '^<(?:.* )?w:w=.*[\'"]([0-9]+)[\'"](?: .*)?/>$'
                w = 0
                if re.match(res, xl):
                    wt = re.sub(res, '\\1', xl)
                    w = round((float(wt) / t_size / 10) - (geta_width * 2))
                wid.append(w)
            if re.match('^<w:trHeight(?: .*)?/>$', xl):
                h = 0
                res = '^<(?:.* )?w:val=.*[\'"]([0-9]+)[\'"](?: .*)?/>$'
                if re.match(res, xl):
                    ht = re.sub(res, '\\1', xl)
                    h = round((float(ht) / t_size / 10) - (geta_height * 2))
                hei.append(h)
            if is_in_cel:
                cell.append(xl)
            if re.match('<w:tr( .*)?>', xl):
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
        # GET LONGEST ROW
        longest_row = 0
        max_length = 0
        for i, row in enumerate(tab):
            if len(row) > max_length:
                max_length = len(row)
                longest_row = i
        # GET CONFIGURE ROW
        if len(tab) == 1:
            half_row = 0
        else:
            half_row = int(len(tab) / 2) + (len(tab) % 2)
        conf_row = half_row
        if longest_row > 0:
            conf_row = longest_row
        # GET NIL OR DOUBLE LINE (ROW)
        row_line = []
        for i in range(len(tab)):
            for tag in tab[i][0]:
                if re.match('^<w:bottom w:val="nil"( .+)?/>$', tag):
                    row_line.append('nil')
                    break
                if re.match('^<w:bottom w:val="double"( .+)?/>$', tag):
                    row_line.append('double')
                    break
            else:
                row_line.append('')
        # GET NIL OR DOUBLE LINE (COLUMN)
        i = longest_row
        col_line = []
        for j in range(len(tab[i])):
            for tag in tab[i][j]:
                if re.match('^<w:right w:val="nil"( .+)?/>$', tag):
                    col_line.append('nil')
                    break
                if re.match('^<w:right w:val="double"( .+)?/>$', tag):
                    col_line.append('double')
                    break
            else:
                col_line.append('')
        # GET ALIGNMENT
        ali = []
        for row in tab:
            tmp = []
            for j, cell in enumerate(row):
                for xml in cell:
                    if re.match('<w:jc w:val=[\'"]left[\'"]/>', xml):
                        if wid[j] == 0:
                            tmp.append('')
                        if wid[j] == 1:
                            tmp.append('-')
                        else:
                            tmp.append(':' + '-' * (wid[j] - 1))
                        break
                    elif re.match('<w:jc w:val=[\'"]center[\'"]/>', xml):
                        if wid[j] == 1:
                            tmp.append(':')
                        else:
                            tmp.append(':' + '-' * (wid[j] - 2) + ':')
                        break
                    elif re.match('<w:jc w:val=[\'"]right[\'"]/>', xml):
                        tmp.append('-' * (wid[j] - 1) + ':')
                        break
                else:
                    tmp.append(':' + '-' * (wid[j] - 1))
                # NIL OR DOUBLE LINE
                if col_line[j] == 'nil':
                    tmp[-1] += '^'
                elif col_line[j] == 'double':
                    tmp[-1] += '='
            ali.append(tmp)
        # GET MD TEXT
        md_text = ''
        is_in_head = True
        for i, row in enumerate(tab):
            if is_in_head:
                if ali[i] == ali[conf_row]:
                    for cell in ali[conf_row]:
                        md_text += '|' + cell
                    is_in_head = False
                    md_text += '|\n'
            for j, cell in enumerate(row):
                text_data, images \
                    = RawParagraph._get_chars_data_and_images('w:tbl', cell)
                raw_text = RawParagraph._get_raw_text(text_data)
                if is_in_head:
                    if not re.match('^:-+:$', ali[i][j]):
                        if re.match('^:-+$', ali[i][j]):
                            raw_text = ': ' + raw_text
                        elif re.match('^-+:$', ali[i][j]):
                            raw_text = raw_text + ' :'
                else:
                    if j < len(ali[conf_row]) and \
                       ali[i][j] != ali[conf_row][j]:
                        if re.match('^:-+:$', ali[i][j]):
                            raw_text = ': ' + raw_text + ' :'
                        elif re.match('^:-+$', ali[i][j]):
                            raw_text = ': ' + raw_text
                        elif re.match('^-+:$', ali[i][j]):
                            raw_text = raw_text + ' :'
                if re.match('^:-+:$', ali[i][j]):
                    raw_text = re.sub('^(\\\\\\s+)', ' \\1', raw_text)
                    raw_text = re.sub('(\\s+\\\\)$', '\\1 ', raw_text)
                else:
                    raw_text = re.sub('^\\\\', '', raw_text)
                    raw_text = re.sub('\\\\$', '', raw_text)
                raw_text = re.sub('\\|', '\\\\|', raw_text)
                raw_text = re.sub('\n', '<br>', raw_text)
                md_text += '|' + raw_text
            md_text += '|'
            # HEIGHT
            if hei[i] > 0:
                for xml in cell:
                    if re.match('<w:vAlign w:val=[\'"]top[\'"]/>', xml):
                        if hei[i] == 1:
                            md_text += '-'
                        else:
                            md_text += ':' + '-' * (hei[i] - 1)
                    elif re.match('<w:vAlign w:val=[\'"]center[\'"]/>', xml):
                        if hei[i] == 1:
                            md_text += ':'
                        else:
                            md_text += ':' + '-' * (hei[i] - 2) + ':'
                    elif re.match('<w:vAlign w:val=[\'"]bottom[\'"]/>', xml):
                        md_text += '-' * (hei[i] - 1) + ':'
            # NIL OR DOUBLE LINE
            if row_line[i] == 'double':
                md_text += '='
            elif row_line[i] == 'nil':
                md_text += '^'
            md_text += '\n'
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
                md_text = re.sub('<br>(\\s+)', '<br>\\\\\\1', md_text)
                md_text = re.sub('<br>([^\\|])', '<br>\\\n    \\1', md_text)
                break
        return md_text

    def get_text_to_write_with_reviser(self):
        self.head_font_revisers = []
        self.tail_font_revisers = []
        text_to_write_with_reviser = super().get_text_to_write_with_reviser()
        # self.text_to_write_with_reviser = text_to_write_with_reviser
        return text_to_write_with_reviser


class ParagraphImage(Paragraph):

    """A class to handle image paragraph"""

    paragraph_class = 'image'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rtx = rp.raw_text_doi
        rp_img = rp.images
        rp_txt = re.sub(RES_IMAGE, '', rp_rtx)
        rp_txt = re.sub('\n.*$', '', rp_txt)  # for caption
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        if rp_txt == '' and len(rp_img) > 0:
            return True
        return False

    def _get_md_text(self, raw_text):
        # CAPTION
        if re.match('^.*\\(.*\\)\n.*$', raw_text):
            caption = re.sub('^.*\n', '', raw_text)
            raw_text = re.sub('\n.*$', '', raw_text)
            raw_text \
                = re.sub('\\((.*)\\)$', '(\\1 "' + caption + '")', raw_text)
        alignment = self.alignment
        text_w = PAPER_WIDTH[Form.paper_size] \
            - Form.left_margin - Form.right_margin
        text_h = PAPER_HEIGHT[Form.paper_size] \
            - Form.top_margin - Form.bottom_margin
        text_size = (text_w, text_h)
        md_text = ParagraphImage.replace_with_fixed_size(raw_text, text_size)
        if alignment == 'left':
            md_text = ': ' + md_text
        elif alignment == 'right':
            md_text = md_text + ' :'
        return md_text

    @staticmethod
    def replace_with_fixed_size(img_text, fixed):
        res = RES_IMAGE_WITH_SIZE
        if re.match(res, img_text):
            alte = re.sub(res, '\\1', img_text)
            cm_w = float(re.sub(res, '\\2', img_text))
            cm_h = float(re.sub(res, '\\3', img_text))
            path = re.sub(res, '\\4', img_text)
            if cm_w >= fixed[0] * 0.98 and cm_w <= fixed[0] * 1.02:
                cm_w = -1
            if cm_w >= fixed[0] * 0.48 and cm_w <= fixed[0] * 0.52:
                cm_w = -0.5
            if cm_h >= fixed[1] * 0.98 and cm_h <= fixed[1] * 1.02:
                cm_h = -1
            if cm_h >= fixed[1] * 0.48 and cm_h <= fixed[1] * 0.52:
                cm_h = -0.5
            if cm_w < 0 and cm_h < 0:
                img_text = '!' \
                    + '[' + alte + ':' + str(cm_w) + 'x' + str(cm_h) + ']' \
                    + '(' + path + ')'
            elif cm_w < 0:
                img_text = '!' \
                    + '[' + alte + ':' + str(cm_w) + 'x' + ']' \
                    + '(' + path + ')'
            elif cm_h < 0:
                img_text = '!' \
                    + '[' + alte + ':' + 'x' + str(cm_h) + ']' \
                    + '(' + path + ')'
        return img_text


class ParagraphMath(Paragraph):

    """A class to handle math paragraph"""

    paragraph_class = 'math'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_rtx = rp.raw_text_doi
        rfd = RES_FONT_DECORATORS
        res = '^' + rfd + '\\\\\\[(.*)\\\\\\]' + rfd + '$'
        if re.match('^' + rfd + '\\\\\\[.*$', rp_rtx):
            if re.match(NOT_ESCAPED + '\\\\\\]' + rfd + '$', rp_rtx):
                tmp = re.sub(res, '\\2', rp_rtx)
                if not re.match(NOT_ESCAPED + '\\\\[\\[\\]].*$', tmp):
                    return True
        return False

    def get_text_to_write(self):
        ttw = super().get_text_to_write()
        alignment = self.alignment
        if alignment == 'left':
            ttw = re.sub('^\\\\\\[', '\\\\\\[: ', ttw)
        elif alignment == 'right':
            ttow = re.sub('\\\\\\]$', ' :\\\\\\]', ttw)
        com = '\\\\(?:int|iint|iiint|oint|sum|prod)'
        ttw = self.shift_paren(com, 5, '_{.*}\\^{.*}{.*}', ttw)
        com = '\\\\(?:int|iint|iiint|oint|sum|prod)'
        ttw = self.shift_paren(com, 1, '{.*}', ttw)
        com = '\\\\(?:log|lim)'
        ttw = self.shift_paren(com, 3, '_{.*}{.*}', ttw)
        com = '\\\\(?:sin|cos|tan)'
        ttw = self.shift_paren(com, 3, '\\^{.*}{.*}', ttw)
        com = '\\\\(?:log|sin|cos|tan|exp|vec)'
        ttw = self.shift_paren(com, 1, '{.*}', ttw)
        ttw = self.cancel_multi_paren(ttw)
        text_to_write = ttw
        return text_to_write

    @staticmethod
    def shift_paren(com, cnt, res, unit):
        res_com = NOT_ESCAPED + '(' + com + ')(}+)$'
        tmp = ''
        while tmp != unit:
            tmp = unit
            tj = -1
            for j in range(len(unit)):
                if re.match(res_com, unit[:j]) and unit[j] != '}':
                    tj = j
                    break
            if tj == -1:
                break
            tk = -1
            dep = []
            d = 0
            for k in range(tj, len(unit)):
                if unit[k] == '{':
                    d += 1
                if unit[k] == '}':
                    d -= 1
                dep.append(d)
                if cnt == -1 and re.match(res, unit[tj:k]):
                    tk = k
                    break
                if dep.count(0) == cnt and re.match(res, unit[tj:k]):
                    tk = k
                    break
            if tk == -1:
                break
            pre_bpa_fds_com_epa = unit[:tj]
            pre_bpa_fds = re.sub(res_com, '\\1', pre_bpa_fds_com_epa)
            com = re.sub(res_com, '\\2', pre_bpa_fds_com_epa)
            epa = re.sub(res_com, '\\3', pre_bpa_fds_com_epa)
            arg = unit[tj:tk]
            pos = unit[tk:]
            ti = -1
            d = - len(epa)
            for i in range(len(pre_bpa_fds) - 1, -1, -1):
                if unit[i] == '{':
                    d += 1
                if unit[i] == '}':
                    d -= 1
                if d == 0:
                    ti = i
                    break
            if ti == -1:
                break
            bpa_fds = pre_bpa_fds[ti:]
            r = '^(.*)(\\\\[A-Za-z]+(?:{[^{}]+})?)(.*)$'
            while re.match(r, bpa_fds):
                f = re.sub(r, '\\2', bpa_fds)
                f = '\\' + re.sub('{[^{}]*}', '{[^{}]*}', f)
                arg = re.sub(f, '', arg)
                bpa_fds = re.sub(r, '\\1\\3', bpa_fds)
            unit = pre_bpa_fds + com + arg + epa + pos
        return unit

    @staticmethod
    def cancel_multi_paren(unit):
        rm = []
        for i in range(len(unit) - 1):
            if unit[i] != '{' or unit[i + 1] != '{':
                continue
            dep = [0]
            d = 0
            for j in range(i, len(unit)):
                if unit[j] == '{':
                    d += 1
                if unit[j] == '}':
                    d -= 1
                dep.append(d)
                if d == 0:
                    if unit[j - 1] == '}' or unit[j] == '}':
                        dep.pop(0)
                        dep.pop(0)
                        dep.pop(-1)
                        dep.pop(-1)
                        if 1 not in dep:
                            rm.append(i)
                            rm.append(j)
                    break
        rm.sort()
        rm.reverse()
        u = list(unit)
        for r in rm:
            u.pop(r)
        unit = ''.join(u)
        return unit


class ParagraphAlignment(Paragraph):

    """A class to handle alignment paragraph"""

    paragraph_class = 'alignment'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_alg = rp.alignment
        if ParagraphChapter.is_this_class(rp):
            return False
        if ParagraphSection.is_this_class(rp):
            return False
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphImage.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
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
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        if rp_sty == 'makdo-g':
            return True
        return False

    @classmethod
    def _get_section_depths(cls, full_text, should_record=False):
        head_section_depth = Paragraph.previous_tail_section_depth
        tail_section_depth = Paragraph.previous_tail_section_depth
        # self.head_section_depth = head_section_depth
        # self.tail_section_depth = tail_section_depth
        return head_section_depth, tail_section_depth

    def _get_md_text(self, raw_text):
        md_text = raw_text
        md_text = re.sub('^`', '', md_text)
        md_text = re.sub('`$', '', md_text)
        res = '^\\[(.*)\\]'
        if re.match(res, md_text):
            md_text = re.sub(res, '\\1', md_text)
        else:
            md_text = '\n' + md_text
        md_text = '``` ' + md_text + '\n```'
        return md_text


class ParagraphHorizontalLine(Paragraph):

    """A class to handle horizontalline paragraph"""

    paragraph_class = 'horizontalline'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        if raw_paragraph.horizontal_line != '':
            return True
        return False

    def _get_length_docx(self):
        if self.horizontal_line == 'textbox':
            return super()._get_length_docx()
        m_size = Form.font_size
        lnsp = Form.line_spacing
        xls = self.xml_lines
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
        for xl in xls:
            sb_xml = XML.get_value('w:spacing', 'w:before', sb_xml, xl)
            sa_xml = XML.get_value('w:spacing', 'w:after', sa_xml, xl)
            ls_xml = XML.get_value('w:spacing', 'w:line', ls_xml, xl)
            fi_xml = XML.get_value('w:ind', 'w:firstLine', fi_xml, xl)
            hi_xml = XML.get_value('w:ind', 'w:hanging', hi_xml, xl)
            li_xml = XML.get_value('w:ind', 'w:left', li_xml, xl)
            ri_xml = XML.get_value('w:ind', 'w:right', ri_xml, xl)
            ti_xml = XML.get_value('w:tblInd', 'w:w', ti_xml, xl)
        # VERTICAL SPACE
        tmp_ls = 0.0
        tmp_sb = (sb_xml / 20)
        tmp_sa = (sa_xml / 20)
        tmp_sb = tmp_sb - ((lnsp - 1) * 0.75 + 0.5) * m_size
        tmp_sa = tmp_sa - ((lnsp - 1) * 0.25 + 0.5) * m_size
        tmp_sb = tmp_sb / lnsp / m_size
        tmp_sa = tmp_sa / lnsp / m_size
        tmp_sb = round(tmp_sb, 2)
        tmp_sa = round(tmp_sa, 2)
        if tmp_sb == tmp_sa:
            tmp_ls = tmp_sb + tmp_sa
            tmp_sb = 0.0
            tmp_sa = 0.0
        length_docx['line spacing'] = tmp_ls
        length_docx['space before'] = tmp_sb
        length_docx['space after'] = tmp_sa
        # HORIZONTAL SPACE
        length_docx['first indent'] = round((fi_xml - hi_xml) / 20 / m_size, 2)
        length_docx['left indent'] = round((li_xml + ti_xml) / 20 / m_size, 2)
        length_docx['right indent'] = round(ri_xml / 20 / m_size, 2)
        # length_docx = self.length_docx
        return length_docx

    def get_text_to_write_with_reviser(self):
        xml_lines = self.xml_lines
        tmp_ttw = self.text_to_write
        self.text_to_write = '----------------'
        ttwwr = super().get_text_to_write_with_reviser()
        self.text_to_write = tmp_ttw
        if xml_lines[-1] == '<horizontalLine:top>':
            if tmp_ttw != '':
                ttwwr = ttwwr + '\n\n' + tmp_ttw
        else:
            if tmp_ttw != '':
                ttwwr = tmp_ttw + '\n\n' + ttwwr
        text_to_write_with_reviser = ttwwr
        # self.text_to_write_with_reviser = text_to_write_with_reviser
        return text_to_write_with_reviser


class ParagraphPagebreak(Paragraph):

    """A class to handle pagebreak paragraph"""

    paragraph_class = 'pagebreak'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_text = rp.raw_text
        rp_xl = rp.xml_lines
        if rp_text != '':
            return False
        for xl in rp_xl:
            if re.match('^<w:br w:type=[\'"]page[\'"]/>$', xl):
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
        rp = raw_paragraph
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphConfiguration.is_this_class(rp):
            return False
        return False


class ParagraphRemarks(Paragraph):

    """A class to handle remarks paragraph"""

    paragraph_class = 'remarks'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_sty = rp.style
        if rp_sty == 'makdo-r':
            return True
        return False

    def get_text_to_write_with_reviser(self):
        md_lines_text = self.md_lines_text
        ttwwr = md_lines_text
        ttwwr = re.sub('^●', '"" ', ttwwr)
        ttwwr = re.sub('\n●', '\n"" ', ttwwr)
        text_to_write_with_reviser = ttwwr
        return text_to_write_with_reviser


class ParagraphSentence(Paragraph):

    """A class to handle sentence paragraph"""

    paragraph_class = 'sentence'

    @classmethod
    def _get_section_depths(cls, full_text, should_record=False):
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
        rp_rtx = rp.raw_text_doi
        rp_xls = rp.xml_lines
        if rp.raw_class == 'w:sectPr':
            return True
        if rp_rtx == '':
            for xl in rp_xls:
                if re.match('<w:sectPr( .*)?>', xl):
                    return True
        return False


class Docx2Md:

    """A class to make a Markdown file from a MS Word file"""

    def __init__(self, inputed_docx_file, args=None):
        self.io = IO()
        io = self.io
        self.doc = Document()
        doc = self.doc
        self.frm = Form()
        frm = self.frm
        # RAED MS WORD FILE
        io.set_docx_file(inputed_docx_file)
        io.unpack_docx_file()
        document_xml_lines = io.read_xml_file('/word/document.xml')
        core_xml_lines = io.read_xml_file('/docProps/core.xml')
        styles_xml_lines = io.read_xml_file('/word/styles.xml')
        header1_xml_lines = io.read_xml_file('/word/header1.xml')
        header2_xml_lines = io.read_xml_file('/word/header2.xml')
        footer1_xml_lines = io.read_xml_file('/word/footer1.xml')
        footer2_xml_lines = io.read_xml_file('/word/footer2.xml')
        rels_xml_lines = io.read_xml_file('/word/_rels/document.xml.rels')
        comments_xml_lines = io.read_xml_file('/word/comments.xml')
        # CONFIGURE
        frm.document_xml_lines = document_xml_lines
        frm.core_xml_lines = core_xml_lines
        frm.styles_xml_lines = styles_xml_lines
        frm.header1_xml_lines = header1_xml_lines
        frm.header2_xml_lines = header2_xml_lines
        frm.footer1_xml_lines = footer1_xml_lines
        frm.footer2_xml_lines = footer2_xml_lines
        frm.rels_xml_lines = rels_xml_lines
        frm.comments_xml_lines = comments_xml_lines
        frm.args = args
        frm.configure()
        # IMAGE LIST
        Form.rels = Form.get_rels(rels_xml_lines)
        # REMARKS
        Form.remarks = Form.get_remarks(comments_xml_lines)
        # STYLE LIST
        Form.styles = Form.get_styles(styles_xml_lines)
        # PRESERVE
        doc.document_xml_lines = document_xml_lines

    def make_md(self, inputed_md_file):
        io = self.io
        doc = self.doc
        document_xml_lines = doc.document_xml_lines
        # SET MARKDOWN FILE NAME
        io.set_md_file(inputed_md_file)
        IO.media_dir = io.get_media_dir()
        # MAKE DOCUMUNT
        doc.raw_paragraphs = doc.get_raw_paragraphs(document_xml_lines)
        doc.paragraphs = doc.get_paragraphs(doc.raw_paragraphs)
        doc.paragraphs = doc.modify_paragraphs()

    def save(self, inputed_md_file):
        io = self.io
        doc = self.doc
        frm = self.frm
        # MAKE MD
        self.make_md(inputed_md_file)
        # SAVE MARKDOWN FILE
        io.open_md_file()
        cfgs = frm.get_configurations()
        io.write_md_file(cfgs)
        dcmt = doc.get_document()
        io.write_md_file(dcmt)
        imgs = doc.get_images()
        io.save_images(imgs)
        io.close_md_file()

    @staticmethod
    def set_document_title(value):
        return Form.set_document_title(value)

    @staticmethod
    def get_document_title():
        return Form.document_title

    @staticmethod
    def set_document_style(value):
        return Form.set_document_style(value)

    @staticmethod
    def get_document_style():
        return Form.document_style

    @staticmethod
    def set_paper_size(value):
        return Form.set_paper_size(value)

    @staticmethod
    def get_paper_size():
        return Form.paper_size

    @staticmethod
    def set_top_margin(value):
        return Form.set_top_margin(str(value))

    @staticmethod
    def get_top_margin():
        return Form.top_margin

    @staticmethod
    def set_bottom_margin(value):
        return Form.set_bottom_margin(str(value))

    @staticmethod
    def get_bottom_margin():
        return Form.bottom_margin

    @staticmethod
    def set_left_margin(value):
        return Form.set_left_margin(str(value))

    @staticmethod
    def get_left_margin():
        return Form.left_margin

    @staticmethod
    def set_right_margin(value):
        return Form.set_right_margin(str(value))

    @staticmethod
    def get_right_margin():
        return Form.right_margin

    @staticmethod
    def set_header_string(value):
        return Form.set_header_string(value)

    @staticmethod
    def get_header_string():
        return Form.header_string

    @staticmethod
    def set_page_number(value):
        return Form.set_page_number(value)

    @staticmethod
    def get_page_number():
        return Form.page_number

    @staticmethod
    def set_line_number(value):
        return Form.set_line_number(value)

    @staticmethod
    def get_line_number():
        return Form.line_number

    @staticmethod
    def set_mincho_font(value):
        return Form.set_mincho_font(value)

    @staticmethod
    def get_mincho_font():
        return Form.mincho_font

    @staticmethod
    def set_gothic_font(value):
        return Form.set_gothic_font(value)

    @staticmethod
    def get_gothic_font():
        return Form.gothic_font

    @staticmethod
    def set_ivs_font(value):
        return Form.set_ivs_font(value)

    @staticmethod
    def get_ivs_font():
        return Form.ivs_font

    @staticmethod
    def set_font_size(value):
        return Form.set_font_size(str(value))

    @staticmethod
    def get_font_size():
        return Form.font_size

    @staticmethod
    def set_line_spacing(value):
        return Form.set_line_spacing(str(value))

    @staticmethod
    def get_line_spacing():
        return Form.line_spacing

    @staticmethod
    def set_space_before(value):
        return Form.set_space_before(value)

    @staticmethod
    def get_space_before():
        return Form.space_before

    @staticmethod
    def set_space_after(value):
        return Form.set_space_after(value)

    @staticmethod
    def get_space_after():
        return Form.space_after

    @staticmethod
    def set_auto_space(value):
        return Form.set_auto_space(str(value))

    @staticmethod
    def get_auto_space():
        return Form.auto_space

    @staticmethod
    def set_version_number(value):
        return Form.set_version_number(value)

    @staticmethod
    def get_version_number():
        return Form.version_number

    @staticmethod
    def set_content_status(value):
        return Form.set_content_status(value)

    @staticmethod
    def get_content_status():
        return Form.content_status

    @staticmethod
    def set_has_completed(value):
        return Form.set_has_completed(str(value))

    @staticmethod
    def get_has_completed():
        return Form.has_completed


############################################################
# MAIN


def main():
    args = get_arguments()
    d2m = Docx2Md(args.docx_file, args)
    d2m.save(args.md_file)
    sys.exit(0)


if __name__ == '__main__':
    main()
