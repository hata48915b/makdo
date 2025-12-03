#!/usr/bin/python3
# Name:         docx2md.py
# Version:      v08 Omachi
# Time-stamp:   <2025.12.03-14:43:26-JST>

# docx2md.py
# Copyright (C) 2022-2025  Seiichiro HATA
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
# 2025.01.04 v08 Omachi

__version__ = 'v08 Omachi'


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

# document -> paragraph -> text -> chars -> imm


############################################################
# SETTING


import sys
import os
import shutil
import argparse     # Python Software Foundation License
import re
import unicodedata
import datetime     # Zope Public License
import tempfile


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
        choices=['A3', 'A3L', 'A3P', 'A4', 'A4L', 'A4P', 'slide'],
        help='用紙設定（A3、A3L、A3P、A4、A4L、A4P、slide）')
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
        help='行間隔（単位文字）')
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
                'A4': 29.7, 'A4L': 21.0, 'A4P': 29.7,
                'slide': 14.2875}
PAPER_WIDTH = {'A3': 42.0, 'A3L': 42.0, 'A3P': 29.7,
               'A4': 21.0, 'A4L': 29.7, 'A4P': 21.0,
               'slide': 25.4}

DEFAULT_TOP_MARGIN = 3.5
DEFAULT_BOTTOM_MARGIN = 2.2
DEFAULT_LEFT_MARGIN = 3.0
DEFAULT_RIGHT_MARGIN = 2.3  # 21.0 - (2.54/72*12*37) - 3.0 = 2.3366666...
# DEFAULT_RIGHT_MARGIN = 2.3  # 21.0 - (2.54/72*12*38) - 3.0 = 1.9133333...
# DEFAULT_RIGHT_MARGIN = 2.0

DEFAULT_HEADER_STRING = ''

DEFAULT_PAGE_NUMBER = ': n :'

DEFAULT_LINE_NUMBER = False

DEFAULT_MINCHO_FONT = 'Times New Roman / ＭＳ 明朝'
DEFAULT_GOTHIC_FONT = 'ＭＳ ゴシック'
DEFAULT_IVS_FONT = 'IPAmj明朝'  # IPAmjMincho
DEFAULT_MATH_FONT = 'Cambria Math'
# DEFAULT_MATH_FONT = 'Liberation Serif'
DEFAULT_FONT_SIZE = 12.0

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
TABLE_LINE_SPACING = 1.5

DEFAULT_CHAR_SPACING = 0.0
# DEFAULT_CHAR_SPACING = 0.0208  # 5/12/20=.0208333...

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

BASIC_TABLE_CELL_HEIGHT = 1.5
BASIC_TABLE_CELL_WIDTH = 1.5  # >= 1.1068

NOT_ESCAPED = '^((?:(?:.|\n)*?[^\\\\])??(?:\\\\\\\\)*?)?'

RES_NUMBER = '(?:[-\\+]?(?:[0-9]*\\.)?[0-9]+)'
RES_NUMBER6 = '(?:' + RES_NUMBER + '?,){,5}' + RES_NUMBER + '?,?'

RES_KATAKANA = '[' + 'ｦｱ-ﾝ' + \
    'アイウエオカキクケコサシスセソタチツテトナニヌネノ' + \
    'ハヒフヘホマミムメモヤユヨラリルレロワヰヱヲン' + ']'

RES_FORCED_TO_BE_FULL_WIDTH = '[' + \
    '⓪' + \
    '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳' + \
    '㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵' + \
    '㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿' + \
    ']'

RES_IMAGE = '! *\\[([^\\[\\]]*)\\] *\\(([^\\(\\)]+)\\)'
RES_IMAGE_WITH_SIZE \
    = '!' \
    + ' *' \
    + '\\[([^\\[\\]]+?)\\s*@(' + RES_NUMBER + ')x(' + RES_NUMBER + ')\\]' \
    + ' *' \
    + '\\(([^\\(\\)]+)\\)'

# MS OFFICE
RES_XML_IMG_MS \
    = '^<v:imagedata r:id=[\'"]([^\'"]+)[\'"] o:title=[\'"]([^\'"]+)[\'"]/>$'
# PYTHON-DOCX AND LIBREOFFICE
RES_XML_IMG_PY_ID \
    = '^<a:blip r:embed=[\'"]([^\'"]+)[\'"](?: .*)?/?>$'
RES_XML_IMG_PY_NAME \
    = '^<pic:cNvPr id=[\'"]([^\'"]+)[\'"] name=[\'"]([^\'"]+)[\'"](?: .*)?/?>$'
RES_XML_IMG_SIZE \
    = '^<wp:extent cx=[\'"]([0-9]+)[\'"] cy=[\'"]([0-9]+)[\'"](?: .*)?/>$'

FONT_DECORATORS_INVISIBLE = [
    '\\*\\*\\*',                     # italic and bold
    '\\*\\*',                        # bold
    '\\*',                           # italic
    '//',                            # italic
    '\\^[0-9A-Za-z]{0,11}\\^',       # font color
    '\\->', '<\\-', '\\+>', '<\\+',  # track changes
]
FONT_DECORATORS_VISIBLE = [
    '\\-\\-\\-',                     # xsmall
    '\\-\\-',                        # small
    '\\+\\+\\+',                     # xlarge
    '\\+\\+',                        # large
    '>>>',                           # xnarrow or reset
    '>>',                            # narrow or reset
    '<<<',                           # xwide or reset
    '<<',                            # wide or reset
    '~~',                            # strikethrough
    '\\[\\|', '\\|\\]',              # frame
    '_[\\$=\\.#\\-~\\+]{,4}_',       # underline
    '_[0-9A-Za-z]{1,11}_',           # highlight color
    '`',                             # preformatted
    '@' + RES_NUMBER + '@',          # font scale
    '@[^@]{1,66}@',                  # font
]
FONT_DECORATORS = FONT_DECORATORS_INVISIBLE + FONT_DECORATORS_VISIBLE
RES_FONT_DECORATORS = '((?:' + '|'.join(FONT_DECORATORS) + ')*)'

RELAX_SYMBOL = '<>'

TAB_WIDTH = 4

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
    'FF0000': 'red',          # 'R'
    '770000': 'darkRed',      # 'DR'
    'FFFF00': 'yellow',       # 'Y'
    '777700': 'darkYellow',   # 'DY'
    '00FF00': 'green',        # 'G'
    '007700': 'darkGreen',    # 'DG'
    '00FFFF': 'cyan',         # 'C'
    '007777': 'darkCyan',     # 'DC'
    '0000FF': 'blue',         # 'B'
    '000077': 'darkBlue',     # 'DB'
    'FF00FF': 'magenta',      # 'M'
    '770077': 'darkMagenta',  # 'DM'
    'BBBBBB': 'lightGray',    # 'G1'
    '777777': 'darkGray',     # 'G2'
    '000000': 'black',        # 'BK'
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
        if c == '\t':
            wid = (int(wid / TAB_WIDTH) + 1) * TAB_WIDTH
            continue
        w = unicodedata.east_asian_width(c)
        if c == '':
            wid += 0.0
        elif re.match('^[☐☑]$', c):
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
        if c == '\t':
            wid = (int(wid / TAB_WIDTH) + 1) * TAB_WIDTH
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
        return int(i)
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


def n2c_p_arab(n, md_line=None):
    return '(' + str(n) + ')'


def n2c_n_kata(n, md_line=None):
    if n >= 1 and n <= 5:
        # アイウエオ
        return chr(12448 + (2 * n))
    elif n >= 1 and n <= 17:
        # カキクケコサシスセソタチ
        return chr(12448 + (2 * n) - 1)
    elif n >= 1 and n <= 20:
        # ツテト
        return chr(12448 + (2 * n))
    elif n >= 1 and n <= 25:
        # ナニヌネノ
        return chr(12448 + (1 * n) + 21)
    elif n >= 1 and n <= 30:
        # ハヒフヘホ
        return chr(12448 + (3 * n) - 31)
    elif n >= 1 and n <= 35:
        # マミムメモ
        return chr(12448 + (1 * n) + 31)
    elif n >= 1 and n <= 38:
        # ヤユヨ
        return chr(12448 + (2 * n) - 4)
    elif n >= 1 and n <= 43:
        # ラリルレロ
        return chr(12448 + (1 * n) + 34)
    elif n >= 1 and n <= 48:
        # ワヰヱヲン
        return chr(12448 + (1 * n) + 35)
    else:
        msg = '※ 警告: ' \
            + 'カタカナ番号は範囲を超えています'
        # msg = 'warning: ' \
        #     + 'overflowed katakana'
        if md_line is None:
            sys.stderr.write(msg + '\n\n')
        else:
            md_line.append_warning_message(msg)
        return '〓'


def n2c_n_alph(n, md_line=None):
    if n >= 1 and n <= 26:
        # ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ
        return chr(65344 + n)
    else:
        msg = '※ 警告: ' \
            + 'アルファベット番号は範囲を超えています'
        # msg = 'warning: ' \
        #     + 'overflowed alphabet'
        if md_line is None:
            sys.stderr.write(msg + '\n\n')
        else:
            md_line.append_warning_message(msg)
        return '〓'


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
        self.temp_dir_instance = None
        self.temp_dir = None
        self.docx_input = None
        self.md_file_instance = None
        # SUBSTITUTE
        self.temp_dir_instance = tempfile.TemporaryDirectory()
        self.temp_dir = self.temp_dir_instance.name

    def set_docx_file(self, inputed_docx_file):
        docx_file = inputed_docx_file
        if not self.__verify_input_file(docx_file):
            return False
        self.inputed_docx_file = inputed_docx_file
        self.docx_file = docx_file
        return True

    @staticmethod
    def __verify_input_file(input_file):
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
        if not self.__verify_output_file(md_file):
            return False
        if not self.__verify_older(docx_file, md_file):
            return False
        self.inputed_md_file = inputed_md_file
        self.md_file = md_file
        return True

    @staticmethod
    def __verify_output_file(output_file):
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
    def __verify_older(input_file, output_file):
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

    def open_md_file(self):
        self.md_file_instance = MdFile(self.md_file)
        self.md_file_instance.open_file()

    def write_md_file(self, article):
        self.md_file_instance.write_file(article)

    def close_md_file(self):
        self.md_file_instance.close_file()

    def save_images(self, images):
        media_dir = self.media_dir
        if len(images) == 0:
            return
        if media_dir == '':
            return
        self.__make_media_dir(media_dir)
        self.__copy_images(images)

    @staticmethod
    def __make_media_dir(media_dir):
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

    def __copy_images(self, images):
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
            #     + 'failde to unpack a input file "' + docx_file + '"'
            sys.stderr.write(msg + '\n\n')
            raise BaseException('failed to unpack docx file')
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
            raise BaseException('is not a MS Word file')
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
            #     + 'failed to read "' + xml_file + '"'
            sys.stderr.write(msg + '\n\n')
            raise BaseException('failed to read xml file')
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

    def open_file(self):
        md_file = self.md_file
        # OPEN
        if md_file == '-':
            md_output = sys.stdout
        else:
            self.__save_old_file(md_file)
            try:
                md_output = open(md_file, 'w', encoding='utf-8', newline='\n')
            except BaseException:
                msg = '※ エラー: ' \
                    + '出力ファイル「' + md_file + '」の書き込みに失敗しました'
                # msg = 'error: ' \
                #     + 'failed to write "' + md_file + '"'
                sys.stderr.write(msg + '\n\n')
                raise BaseException('failed to write output file')
                if __name__ == '__main__':
                    sys.exit(204)
                return False
        self.md_output = md_output
        return True

    def write_file(self, article):
        self.md_output.write(article)

    def close_file(self):
        self.md_output.close()

    @staticmethod
    def __save_old_file(output_file):
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
                #     + 'failed to remove "' + backup_file + '"'
                raise BaseException('failed to remove backup file')
                sys.stderr.write(msg + '\n\n')
                if __name__ == '__main__':
                    sys.exit(205)
                return False
            os.rename(output_file, backup_file)
        if os.path.exists(output_file):
            msg = '※ エラー: ' \
                + '古いファイル「' + output_file + '」を改名できません'
            # msg = 'error: ' \
            #     + 'failed to rename "' + output_file + '"'
            raise BaseException('failed to rename old file')
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
    created_time = ''
    modified_time = ''

    styles = None
    rels = None
    remarks = None
    auto_numbering_styles = None
    footnotes = None

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
        self.numbering_xml_lines = None
        self.args = None

    def configure(self):
        # PAPER SIZE, MARGIN, LINE NUMBER, DOCUMENT STYLE, FONT
        self._configure_by_document_xml(self.document_xml_lines)
        # DOCUMENT TITLE, DOCUMENT STYLE, VERSION NUMBER, CONTENT STATUS,
        # CREATED TIME, MODIFIED TIME
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
        # STATISTICS
        afonts, jfonts, fsizes = {}, {}, {}
        afonts[''], jfonts[''], fsizes[''] = 0, 0, 0
        for xl in xml_lines:
            width_x = XML.get_value('w:pgSz', 'w:w', width_x, xl)
            height_x = XML.get_value('w:pgSz', 'w:h', height_x, xl)
            top_x = XML.get_value('w:pgMar', 'w:top', top_x, xl)
            bottom_x = XML.get_value('w:pgMar', 'w:bottom', bottom_x, xl)
            left_x = XML.get_value('w:pgMar', 'w:left', left_x, xl)
            right_x = XML.get_value('w:pgMar', 'w:right', right_x, xl)
            if re.match('^<w:rPr( .*)?>$', xl):
                af, jf, fs, fsc = '', '', '', ''
            elif re.match('^</w:rPr( .*)?>$', xl):
                if re.match('^.* w:ascii=[\'"]([^\'"]*)[\'"].*$', af):
                    afonts = XML.count_values('w:rFonts', 'w:ascii',
                                              afonts, af)
                elif re.match('^.* w:cs=[\'"]([^\'"]*)[\'"].*$', af):
                    afonts = XML.count_values('w:rFonts', 'w:cs',
                                              afonts, af)
                else:
                    afonts[''] += 1
                if re.match('^.* w:eastAsia=[\'"]([^\'"]*)[\'"].*$', jf):
                    jfonts = XML.count_values('w:rFonts', 'w:eastAsia',
                                              jfonts, jf)
                elif re.match('^.* w:cs=[\'"]([^\'"]*)[\'"].*$', af):
                    jfonts = XML.count_values('w:rFonts', 'w:cs',
                                              jfonts, jf)
                else:
                    jfonts[''] += 1
                if fs != '':
                    fsizes = XML.count_values('w:sz', 'w:val', fsizes, fs)
                elif fsc != '':
                    fsizes = XML.count_values('w:szCs', 'w:val', fsizes, fsc)
                else:
                    fsizes[''] += 1
            else:
                if re.match('^<w:rFonts( .*)/>$', xl):
                    if re.match('^.* w:ascii=[\'"]([^\'"]*)[\'"].*$', xl):
                        af = xl
                    elif re.match('^.* w:cs=[\'"]([^\'"]*)[\'"].*$', xl):
                        af = xl
                    if re.match('^.* w:eastAsia=[\'"]([^\'"]*)[\'"].*$', xl):
                        jf = xl
                    elif re.match('^.* w:cs=[\'"]([^\'"]*)[\'"].*$', xl):
                        jf = xl
                elif re.match('^<w:sz( .*)/>$', xl):
                    fs = xl
                elif re.match('^<w:szCs( .*)/>$', xl):
                    fsc = xl
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
        if 25.3 <= width and width <= 25.5:
            if 14.1875 <= height and height <= 14.3875:
                Form.paper_size = 'slide'
        # MARGIN
        if top_x >= 0:
            Form.top_margin = round(top_x / 567, 1)
        if bottom_x >= 0:
            Form.bottom_margin = round(bottom_x / 567, 1)
        if left_x >= 0:
            Form.left_margin = round(left_x / 567, 1)
        if right_x >= 0:
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
        # FONT
        afont = self.__get_max(afonts)
        jfont = self.__get_max(jfonts)
        for mfs in MS_FONTS:
            if afont in mfs:
                afont = mfs[0]
            if jfont in mfs:
                jfont = mfs[0]
        if afont == jfont:
            Form.mincho_font = '= / ' + jfont
        else:
            Form.mincho_font = afont + ' / ' + jfont
        fsize = self.__get_max(fsizes)
        if re.match('^[0-9]+$', fsize):
            Form.font_size = round(float(fsize) / 2, 1)

    @staticmethod
    def __get_max(values):
        maximum, value = 0, ''
        for v in values:
            if maximum < values[v]:
                maximum = values[v]
                value = v
        return value

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
            # CREATED TIME
            resb = '^<dcterms:created( .*)?>$'
            rese = '^</dcterms:created>$'
            if i > 0 and re.match(resb, xml_lines[i - 1], re.I):
                if not re.match(rese, xl, re.I):
                    jst = datetime.timezone(datetime.timedelta(hours=+9))
                    d = xl
                    d = re.sub('\\.[0-9]+', '', d)  # '%Y-%m-%dT%H:%M:%S.%f%z'
                    dt = datetime.datetime.strptime(d, '%Y-%m-%dT%H:%M:%S%z')
                    dt = dt.astimezone(jst)
                    Form.created_time = dt.isoformat()
            # MODIFIED TIME
            resb = '^<dcterms:modified( .*)?>$'
            rese = '^</dcterms:modified>$'
            if i > 0 and re.match(resb, xml_lines[i - 1], re.I):
                if not re.match(rese, xl, re.I):
                    jst = datetime.timezone(datetime.timedelta(hours=+9))
                    d = xl
                    d = re.sub('\\.[0-9]+', '', d)  # '%Y-%m-%dT%H:%M:%S.%f%z'
                    dt = datetime.datetime.strptime(d, '%Y-%m-%dT%H:%M:%S%z')
                    dt = dt.astimezone(jst)
                    Form.modified_time = dt.isoformat()

    def _configure_by_styles_xml(self, xml_lines):
        # FONT
        fmf_afnt, fmf_jfnt \
            = RawParagraph._get_ascii_and_kanji_font(Form.mincho_font)
        sty_afnt = ''
        sty_jfnt = ''
        is_in_default = False
        for xl in xml_lines:
            if xl == '<w:docDefaults>':
                is_in_default = True
            elif xl == '</w:docDefaults>':
                break
            if not is_in_default:
                continue
            sty_afnt = XML.get_value('w:rFonts', 'w:ascii', sty_afnt, xl)
            sty_jfnt = XML.get_value('w:rFonts', 'w:eastAsia', sty_jfnt, xl)
        def_afnt, def_jfnt \
            = RawParagraph._get_ascii_and_kanji_font(DEFAULT_MINCHO_FONT)
        if fmf_afnt != '':
            afnt = fmf_afnt
        elif sty_afnt != '':
            afnt = sty_afnt
        else:
            afnt = def_afnt
        if fmf_jfnt != '':
            jfnt = fmf_jfnt
        elif sty_jfnt != '':
            jfnt = sty_jfnt
        else:
            jfnt = def_jfnt
        Form.mincho_font = FontDecorator.get_font_name(afnt, jfnt)
        # BLOCKS
        xml_body = XML.get_body('w:styles', xml_lines)
        xml_blocks = XML.get_blocks(xml_body)
        sb = ['0.0', '0.0', '0.0', '0.0', '0.0', '0.0']
        sa = ['0.0', '0.0', '0.0', '0.0', '0.0', '0.0']
        for xb in xml_blocks:
            name = ''
            afnt = ''
            jfnt = ''
            sz_x = -1.0
            f_it = False
            f_bd = False
            f_sk = False
            f_fr = False
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
                jfnt = XML.get_value('w:rFonts', 'w:eastAsia', jfnt, xl)
                sz_x = XML.get_value('w:sz', 'w:val', sz_x, xl)
                sz_x = XML.get_value('w:szCz', 'w:val', sz_x, xl)
                f_it = XML.is_this_tag('w:i', f_it, xl)
                f_bd = XML.is_this_tag('w:b', f_bd, xl)
                f_sk = XML.is_this_tag('w:strike', f_sk, xl)
                f_fr = XML.is_this_tag('w:bdr', f_sk, xl)
                f_ul = XML.get_value('w:u', 'w:val', f_ul, xl)
                f_cl = XML.get_value('w:color', 'w:val', f_cl, xl)
                f_hc = XML.get_value('w:highlight', 'w:val', f_hc, xl)
                alig = XML.get_value('w:jc', 'w:val', alig, xl)
                ls_x = XML.get_value('w:spacing', 'w:line', ls_x, xl)
                ase = XML.get_value('w:autoSpaceDE', 'w:val', ase, xl)
                asn = XML.get_value('w:autoSpaceDN', 'w:val', asn, xl)
            if name == 'makdo':
                # MINCHO FONT
                if afnt != '' and jfnt != '':
                    if afnt == jfnt:
                        Form.mincho_font = '= / ' + jfnt
                    else:
                        Form.mincho_font = afnt + ' / ' + jfnt
                elif afnt != '' and jfnt == '':
                    Form.mincho_font = afnt
                elif afnt == '' and jfnt != '':
                    Form.mincho_font = jfnt
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
                if afnt != '' and jfnt != '':
                    if afnt == jfnt:
                        Form.gothic_font = '= / ' + jfnt
                    else:
                        Form.gothic_font = afnt + ' / ' + jfnt
                elif afnt != '' and jfnt == '':
                    Form.gothic_font = afnt
                elif afnt == '' and jfnt != '':
                    Form.gothic_font = jfnt
            elif name == 'makdo-i':
                # IVS FONT
                if jfnt != '':
                    Form.ivs_font = jfnt
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
        style, alignment, chars_data, raw_text, images, footnotes \
            = RawParagraph.get_raw_text_and_etc(xml_lines, 'header')
        if alignment == 'center':
            raw_text = ': ' + raw_text + ' :'
        elif alignment == 'right':
            raw_text = raw_text + ' :'
        if raw_text != '':
            Form.header_string = raw_text

    @staticmethod
    def _configure_by_footerX_xml(xml_lines):
        # PAGE NUMBER
        style, alignment, chars_data, raw_text, images, footnotes \
            = RawParagraph.get_raw_text_and_etc(xml_lines, 'footer')
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
        elif value == 'slide' or value == 'スライド':
            Form.paper_size = 'slide'
            return True
        msg = '※ 警告: ' \
            + '「' + item + '」の値は' \
            + '"A3横"、"A3縦"、"A4横"、"A4縦"又は"スライド"で' \
            + 'なければなりません'
        # msg = 'warning: ' \
        #     + '"' + item + '" must be "A3", "A3P", "A4", "A4L" or "slide"'
        sys.stderr.write(msg + '\n\n')
        return False

    @staticmethod
    def set_top_margin(value, item='top_margin'):
        return Form.__set_margin(value, item)

    @staticmethod
    def set_bottom_margin(value, item='bottom_margin'):
        return Form.__set_margin(value, item)

    @staticmethod
    def set_left_margin(value, item='left_margin'):
        return Form.__set_margin(value, item)

    @staticmethod
    def set_right_margin(value, item='right_margin'):
        return Form.__set_margin(value, item)

    @staticmethod
    def __set_margin(value, item):
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
        return cls._get_configurations_in_japanese()
        # return cls._get_configurations_in_english()

    @classmethod
    def _get_configurations_in_english(cls):
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
        cfgs += 'created_time:   ' + cls.created_time + '\n'
        cfgs += 'modified_time:  ' + cls.modified_time + '\n'
        cfgs += \
            '---------------------------------------------------------------->'
        cfgs += '\n'
        cfgs += '\n'
        return cfgs

    @classmethod
    def _get_configurations_in_japanese(cls):
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
            '# 用紙のサイズ（A3横、A3縦、A4横、A4縦、スライド）を指定できます。'
        cfgs += '\n'
        if cls.paper_size == 'slide':
            cfgs += '用紙サ: スライド\n'
        elif cls.paper_size == 'A3L' or cls.paper_size == 'A3':
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
        if cls.font_size.is_integer():
            cfgs += '文字サ: ' + str(int(cls.font_size)) + ' pt\n'
        else:
            cfgs += '文字サ: ' + str(round(cls.font_size, 1)) + ' pt\n'
        cfgs += '\n'

        cfgs += \
            '# 行間隔を基本の文字の高さの何倍にするかを指定できます。'
        cfgs += '\n'
        cfgs += '行間隔: ' + str(round(cls.line_spacing, 2)) + ' 倍\n'
        cfgs += '\n'

        cfgs += \
            '# セクションタイトル前後の余白を行間隔の倍数で指定できます。'
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
            '# 原稿の作成日時と更新日時が自動で記録されます。'
        cfgs += '\n'
        cfgs += '作成時: ' + cls.created_time + '\n'
        cfgs += '更新時: ' + cls.modified_time + '\n'
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
        res_head, res_tail = '^<Relationship(?: .*)?', '(?: .*)?/>$'
        res_id, res_tg = ' Id=[\'"](.*?)[\'"]', ' Target=[\'"](.*?)[\'"]'
        res1 = res_head + res_id + '(?: .*)?' + res_tg + res_tail
        res2 = res_head + res_tg + '(?: .*)?' + res_id + res_tail
        for xl in xml_lines:
            if re.match(res1, xl):
                rel_id = re.sub(res1, '\\1', xl)
                rel_tg = re.sub(res1, '\\2', xl)
                rels[rel_id] = rel_tg
            elif re.match(res2, xl):
                rel_id = re.sub(res2, '\\2', xl)
                rel_tg = re.sub(res2, '\\1', xl)
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

    @staticmethod
    def get_auto_numbering_styles(xml_lines):
        # 1ST STEP
        s1_styles = {}
        res_s1a_beg = '^<w:abstractNum w:abstractNumId="([0-9]+)"(?: .+)?>$'
        res_s1a_end = '^</w:abstractNum>$'
        res_s1b_beg = '^<w:lvl w:ilvl="([0-9]+)"(?: .+)?>$'
        res_s1b_end = '^</w:lvl>$'
        res_s1_str = '^<w:start w:val="([0-9]+)"/>$'
        res_s1_fmt = '^<w:numFmt w:val="(.+)"/>$'
        res_s1_txt = '^<w:lvlText w:val="(.+)"/>$'
        res_s1_ind = '^<w:ind .+/>$'
        res_s1_fir = '^.+ w:firstLine="([0-9]+)".+$'
        res_s1_han = '^.+ w:hanging="([0-9]+)".+$'
        res_s1_lef = '^.+ w:left="([0-9]+)".+$'
        # 2ND STEP
        s2_styles = {}
        res_s2_beg = '^<w:num w:numId="([0-9]+)"(?: .*)?>$'
        res_s2_end = '^</w:num(?: .*)?>$'
        res_s2_num = '^<w:abstractNumId w:val="([0-9]+)"/>$'
        s2_num = -1
        for xl in xml_lines:
            if False:
                pass
            # 1ST STEP
            elif re.match(res_s1a_beg, xl):
                s1a_num = int(re.sub(res_s1a_beg, '\\1', xl))
            elif re.match(res_s1b_beg, xl):
                s1_str, s1_fmt, s1_txt = None, '', ''
                s1_fir, s1_han, s1_lef = 0.0, 0.0, 0.0
                s1b_num = int(re.sub(res_s1b_beg, '\\1', xl))
            elif re.match(res_s1_str, xl):
                s1_str = int(re.sub(res_s1_str, '\\1', xl))
            elif re.match(res_s1_fmt, xl):
                s1_fmt = re.sub(res_s1_fmt, '\\1', xl)
            elif re.match(res_s1_txt, xl):
                s1_txt = re.sub(res_s1_txt, '\\1', xl)
            elif re.match(res_s1_ind, xl):
                if re.match(res_s1_fir, xl):
                    s1_fir = int(re.sub(res_s1_fir, '\\1', xl))
                if re.match(res_s1_han, xl):
                    s1_han = int(re.sub(res_s1_han, '\\1', xl))
                if re.match(res_s1_lef, xl):
                    s1_lef = int(re.sub(res_s1_lef, '\\1', xl))
            elif re.match(res_s1b_end, xl):
                paragraph_class, proper_depth \
                    = AutoNumberingStyle.get_class_and_depth(s1_fmt, s1_txt)
                if paragraph_class is not None and proper_depth is not None:
                    ans = AutoNumberingStyle()
                    ans.paragraph_class = paragraph_class
                    ans.proper_depth = proper_depth
                    ans.number_format = s1_fmt
                    ans.head_string = s1_txt
                    ans.start = s1_str
                    ans.raw_first_indent = s1_fir - s1_han
                    # ans.raw_firstline_indent = s1_fir
                    # ans.raw_hanging_indent = s1_han
                    ans.raw_left_indent = s1_lef
                    s1_key = str(s1a_num) + '-' + str(s1b_num)
                    s1_styles[s1_key] = ans
                s1b_num = -1
            elif re.match(res_s1a_end, xl):
                s1a_num = -1
            # 2ND STEP
            elif re.match(res_s2_beg, xl):
                s2_num = int(re.sub(res_s2_beg, '\\1', xl))
            elif re.match(res_s2_num, xl):
                tmp_num = int(re.sub(res_s2_num, '\\1', xl))
                for s1_key in s1_styles:
                    res = '^' + str(tmp_num) + '-([0-9]+)$'
                    if re.match(res, s1_key):
                        s2_key = str(s2_num) + '-' + re.sub(res, '\\1', s1_key)
                        s2_styles[s2_key] = s1_styles[s1_key]
            elif re.match(res_s2_end, xl):
                s2_num = -1
        auto_numbering_styles = s2_styles
        return auto_numbering_styles

    @staticmethod
    def get_footnotes(xml_lines):
        footnotes = {}
        _fnid = None
        for xl in xml_lines:
            if re.match('^<w:footnote( .*)>$', xl):
                _fnid = XML.get_value('w:footnote', 'w:id', '', xl)
                footnotes[_fnid] = ''
            if _fnid is not None:
                if not re.match('^<.*>$', xl):
                    footnotes[_fnid] += xl
        return footnotes


class AutoNumberingStyle:

    """A class to handle an auto numbering style"""

    def __init__(self):
        self.paragraph_class = None
        self.proper_depth = None
        self.number_format = None
        self.head_string = None
        self.start = None
        self.state = 0
        self.raw_first_indent = None
        # self.raw_firstline_indent = None
        # self.raw_hanging_indent = None
        self.raw_left_indent = None

    @staticmethod
    def get_class_and_depth(fmt, txt):
        # CHAPTER
        res_sp = '(?:  ?|\t|\u3000)'
        res_c1_a = '^第%[1-9]編' + res_sp + '?$'
        res_c1_b = '^第[0-9０-９]+編(の[0-9０-９]+)*の%[1-9]' + res_sp + '?$'
        res_c2_a = '^第%[1-9]章' + res_sp + '?$'
        res_c2_b = '^第[0-9０-９]+章(の[0-9０-９]+)*の%[1-9]' + res_sp + '?$'
        res_c3_a = '^第%[1-9]節' + res_sp + '?$'
        res_c3_b = '^第[0-9０-９]+節(の[0-9０-９]+)*の%[1-9]' + res_sp + '?$'
        res_c4_a = '^第%[1-9]款' + res_sp + '?$'
        res_c4_b = '^第[0-9０-９]+款(の[0-9０-９]+)*の%[1-9]' + res_sp + '?$'
        res_c5_a = '^第%[1-9]目' + res_sp + '?$'
        res_c5_b = '^第[0-9０-９]+目(の[0-9０-９]+)*の%[1-9]' + res_sp + '?$'
        if fmt == 'decimal' or fmt == 'decimalFullWidth':
            if re.match(res_c1_a, txt) or re.match(res_c1_b, txt):
                return 'chapter', 1
            if re.match(res_c2_a, txt) or re.match(res_c2_b, txt):
                return 'chapter', 2
            if re.match(res_c3_a, txt) or re.match(res_c3_b, txt):
                return 'chapter', 3
            if re.match(res_c4_a, txt) or re.match(res_c4_b, txt):
                return 'chapter', 4
            if re.match(res_c5_a, txt) or re.match(res_c5_b, txt):
                return 'chapter', 5
        # SECTION
        res_sp = '(?:  ?|\t|\u3000|\\. |．)'
        res_s2_a = '^第%[1-9]条?' + res_sp + '?$'
        res_s2_b = '^第[0-9０-９]+条?(の[0-9０-９]+)*の%[1-9]' + res_sp + '?$'
        res_s3_a = '^%[1-9]' + res_sp + '?$'
        res_s3_b = '^[0-9０-９]+(の[0-9０-９]+)*の%[1-9]' + res_sp + '?$'
        res_s4_a = '^[\\(（]%[1-9][\\)）]' + res_sp + '?$'
        res_s5_a = '^%[1-9]' + res_sp + '?$'
        res_s6_a = '^[\\(（]%[1-9][\\)）]' + res_sp + '?$'
        res_s7_a = '^%[1-9]' + res_sp + '?$'
        res_s8_a = '^[\\(（]%[1-9][\\)）]' + res_sp + '?$'
        if fmt == 'decimal' or fmt == 'decimalFullWidth':
            if re.match(res_s2_a, txt) or re.match(res_s2_b, txt):
                return 'section', 2
            if re.match(res_s3_a, txt) or re.match(res_s3_b, txt):
                return 'section', 3
            if re.match(res_s4_a, txt):
                return 'section', 4
        if fmt == 'decimalEnclosedParen':
            if re.match('^%[1-9]' + res_sp + '?$', txt):
                return 'section', 4
        if fmt == 'aiueo' or fmt == 'aiueoFullWidth':
            if re.match(res_s5_a, txt):
                return 'section', 5
            if re.match(res_s6_a, txt):
                return 'section', 6
        if fmt == 'lowerLetter':
            if re.match(res_s7_a, txt):
                return 'section', 7
            if re.match(res_s8_a, txt):
                return 'section', 8
        return None, None

    @staticmethod
    def get_style_key_from_xml_lines(xml_lines):
        res_xml_number_ms = '^<w:numId w:val=[\'"]([0-9]+)[\'"]/>$'
        res_xml_number_lo = '^<w:pStyle w:val=[\'"]ListNumber([0-9]?)[\'"]/>$'
        res_xml_ilvl = '<w:ilvl w:val="([0-9]+)"/>'
        numid = -1
        ilvl = 0
        for xl in xml_lines:
            if re.match(res_xml_number_ms, xl):
                numid = int(re.sub(res_xml_number_ms, '\\1', xl))
            elif re.match(res_xml_number_lo, xl):
                numid = int(re.sub(res_xml_number_lo, '\\1', xl))
            elif re.match(res_xml_ilvl, xl):
                ilvl = re.sub(res_xml_ilvl, '\\1', xl)
        ans_key = str(numid) + '-' + str(ilvl)
        if ans_key in Form.auto_numbering_styles:
            return ans_key
        return None


class CharsDatum:

    """A class to keep characters data"""

    def __init__(self, fr_fd_lst, raw_chars, bk_fd_lst):
        self.chars = ''
        self.fr_fd_cls = FontDecorator([])
        self.bk_fd_cls = FontDecorator([])
        self.set_chars(raw_chars)
        self.set_fds(fr_fd_lst, bk_fd_lst)

    def set_chars(self, raw_chars):
        self.chars = raw_chars

    def set_fds(self, fr_fd_lst, bk_fd_lst):
        # str -> list
        if type(fr_fd_lst) == str:
            fr_fd_lst = [fr_fd_lst]
        if type(bk_fd_lst) == str:
            bk_fd_lst = [bk_fd_lst]
        # '}' -> '_}'
        if ('_{' in fr_fd_lst) and ('^{' not in fr_fd_lst):
            if ('}' in bk_fd_lst) and ('_}' not in bk_fd_lst):
                bk_fd_lst = ['_}' if fd == '}' else fd for fd in bk_fd_lst]
        # '}' -> '^}'
        if ('^{' in fr_fd_lst) and ('_{' not in fr_fd_lst):
            if ('}' in bk_fd_lst) and ('^}' not in bk_fd_lst):
                bk_fd_lst = ['^}' if fd == '}' else fd for fd in bk_fd_lst]
        self.__set_fr_fd_cls(fr_fd_lst)
        self.__set_bk_fd_cls(bk_fd_lst)

    def __set_fr_fd_cls(self, fr_fd_lst):
        if type(fr_fd_lst) == str:
            fr_fd_lst = [fr_fd_lst]
        self.fr_fd_cls.set_fds(fr_fd_lst)

    def __set_bk_fd_cls(self, bk_fd_lst):
        if type(bk_fd_lst) == str:
            bk_fd_lst = [bk_fd_lst]
        self.bk_fd_cls.set_fds(bk_fd_lst)

    def reset_fds(self):
        self.fr_fd_cls.reset_fds()
        self.bk_fd_cls.reset_fds()

    def get_chars_with_fds(self):
        cwf = self.chars
        frs = self.__get_fr_fds_chars()
        bks = self.__get_bk_fds_chars()
        cwf = self.concatenate_imm(frs, cwf)
        cwf = self.concatenate_imm(cwf, bks)
        return cwf

    @staticmethod
    def prepare_imm(fldchar, imm, type='normal'):
        # ESCAPE
        imm = imm.replace('\\', '\\\\')
        imm = imm.replace('*', '\\*')
        imm = imm.replace('`', '\\`')
        imm = imm.replace('~~', '\\~\\~')
        imm = imm.replace('//', '\\/\\/')  # italic
        imm = re.sub('([a-z]+:)\\\\/\\\\/', '\\1//', imm)  # http https ftp
        imm = imm.replace('--', '\\-\\-')          # --
        imm = imm.replace('\\-\\--', '\\-\\-\\-')  # ---
        imm = imm.replace('++', '\\+\\+')          # ++
        imm = imm.replace('\\+\\++', '\\+\\+\\+')  # +++
        imm = imm.replace('>>', '\\>\\>')          # >>
        imm = imm.replace('\\>\\>>', '\\>\\>\\>')  # >>>
        imm = imm.replace('<<', '\\<\\<')          # <<
        imm = imm.replace('\\<\\<<', '\\<\\<\\<')  # <<<
        # imm = imm.replace('__', '\\_\\_')
        imm = re.sub('@([^@]{1,66})@', '\\\\@\\1\\\\@', imm)
        imm = re.sub('_([\\$=\\.#\\-~\\+]*)_', '\\\\_\\1\\\\_', imm)
        imm = re.sub('\\^([0-9a-zA-Z]+)\\^', '\\\\^\\1\\\\^', imm)
        imm = re.sub('_([0-9a-zA-Z]+)_', '\\\\_\\1\\\\_', imm)
        imm = imm.replace('->', '\\->')
        imm = imm.replace('<-', '\\<-')
        imm = imm.replace('+>', '\\+>')
        imm = imm.replace('<+', '\\<+')
        # imm = imm.replace('\\[', '\\[')
        # imm = imm.replace('\\]', '\\]')
        imm = imm.replace('{{', '\\{{')
        imm = imm.replace('}}', '\\}}')
        imm = imm.replace('&lt;', '\\&lt;')
        imm = imm.replace('&gt;', '\\&gt;')
        imm = imm.replace('\\&lt-;', '\\&lt;\\-')  # "<-"
        imm = imm.replace('-\\&gt;', '\\-\\&gt;')  # "->"
        imm = imm.replace('\\&lt;+', '\\&lt;\\+')  # "<+"
        imm = imm.replace('+\\&gt;', '\\+\\&gt;')  # "+>"
        # PAGE NUMBER
        if type == 'footer':
            if fldchar == 'begin':
                res = '^ ?(\\S*)\\s*\\\\\\\\\\\\\\* MERGEFORMAT ?$'
                if re.match(res, imm):
                    imm = re.sub(res, '\\1', imm)
                if re.match('^ ?PAGE ?$', imm, re.I):
                    imm = 'n'
                elif re.match('^ ?SECTIONPAGES ?$', imm, re.I):
                    # "SECTIONPAGES" IS NOT SUPPORTOD BY LIBREOFFICE
                    imm = 'N'
                elif re.match('^ ?NUMPAGES ?$', imm, re.I):
                    imm = 'M'
            else:
                imm = re.sub('(n|N|M)', '\\\\\\1', imm)
        # RETURN
        return imm

    @staticmethod
    def concatenate_imm(imm1, imm2):
        # "~" + "~"
        if re.match(NOT_ESCAPED + '~$', imm1) and re.match('^~', imm2):
            return imm1 + '<>' + imm2
        # "/" + "/"
        if re.match(NOT_ESCAPED + '/$', imm1) and re.match('^/', imm2):
            return imm1 + '<>' + imm2
        # "-" + "-"
        if re.match(NOT_ESCAPED + '-$', imm1) and re.match('^-', imm2):
            return imm1 + '<>' + imm2
        # "+" + "+"
        if re.match(NOT_ESCAPED + '\\+$', imm1) and re.match('^\\+', imm2):
            return imm1 + '<>' + imm2
        # ">" + ">"
        if re.match(NOT_ESCAPED + '>$', imm1) and re.match('^>', imm2):
            return imm1 + '<>' + imm2
        # "<" + "<"
        if re.match(NOT_ESCAPED + '<$', imm1) and re.match('^<', imm2):
            return imm1 + '<>' + imm2
        # "<" + ">"
        if re.match(NOT_ESCAPED + '<$', imm1) and re.match('^>', imm2):
            return imm1 + '<>' + imm2
        # "@.*" + ".*@"
        if re.match(NOT_ESCAPED + '@([^@]{0,66})$', imm1) and \
           not re.match(NOT_ESCAPED + '@([^@]{1,66})@[^@]*$', imm1) and \
           re.match('^([^@]{0,66})@(.|\n)*', imm2) and \
           not re.match('^[^@]*@([^@]{1,66})@(.|\n)*', imm2):
            c1 = re.sub(NOT_ESCAPED + '@([^@]{0,66})$', '\\2', imm1)
            c2 = re.sub('^([^@]{0,66})@(.|\n)*', '\\1', imm2)
            if len(c1 + c2) <= 66:
                return imm1 + '<>' + imm2
        # "_.*" + ".*_"
        if re.match(NOT_ESCAPED + '_([\\$=\\.#\\-~\\+]*)$', imm1) and \
           re.match('^([\\$=\\.#\\-~\\+]*)_(.|\n)*', imm2):
            c1 = re.sub(NOT_ESCAPED + '_([\\$=\\.#\\-~\\+]*)$', '\\2', imm1)
            c2 = re.sub('^([\\$=\\.#\\-~\\+]*)_(.|\n)*', '\\1', imm2)
            for ul in UNDERLINE:
                if c1 + c2 == UNDERLINE[ul]:
                    return imm1 + '<>' + imm2
        # "^.*" + ".*^"
        if re.match(NOT_ESCAPED + '\\^([0-9a-zA-Z]*)$', imm1) and \
           re.match('^([0-9a-zA-Z]*)\\^(.|\n)*', imm2):
            c1 = re.sub(NOT_ESCAPED + '\\^([0-9a-zA-Z]*)$', '\\2', imm1)
            c2 = re.sub('^([0-9a-zA-Z]*)\\^(.|\n)*', '\\1', imm2)
            if re.match('^([0-9A-F]{3})([0-9A-F]{3})?$', c1 + c2):
                return imm1 + '<>' + imm2
            for fc in FONT_COLOR:
                if c1 + c2 == FONT_COLOR[fc]:
                    return imm1 + '<>' + imm2
        # "_.*" + ".*_"
        if re.match(NOT_ESCAPED + '_([0-9a-zA-Z]*)$', imm1) and \
           re.match('^([0-9a-zA-Z]*)_(.|\n)*', imm2):
            c1 = re.sub(NOT_ESCAPED + '_([0-9a-zA-Z]*)$', '\\2', imm1)
            c2 = re.sub('^([0-9a-zA-Z]*)_(.|\n)*', '\\1', imm2)
            for hc in HIGHLIGHT_COLOR:
                if (c1 + c2 == hc) or (c1 + c2 == HIGHLIGHT_COLOR[hc]):
                    return imm1 + '<>' + imm2
        # "-|+" + ">"
        # if re.match(NOT_ESCAPED + '(-|\\+)$', imm1) and \
        #    re.match('^>', imm2):
        #     return imm1 + '<>' + imm2
        # "<" + "-|+"
        # if re.match(NOT_ESCAPED + '<$', imm1) and \
        #    re.match('^(-|\\+)', imm2):
        #     return imm1 + '<>' + imm2
        # "\" + "[|]"
        # if re.match(NOT_ESCAPED + '\\\\$', imm1) and \
        #    re.match('^(\\[|\\])', imm2):
        #     return imm1 + '<>' + imm2
        return imm1 + imm2

    def __get_fr_fds_chars(self):
        return self.fr_fd_cls.get_ord_fds()

    def __get_bk_fds_chars(self):
        return self.bk_fd_cls.get_rev_fds()

    # def is_empty(self):
    #     if self.chars != '':
    #         return False
    #     if not self.__is_same_fds():
    #         return False
    #     return True

    # def __is_same_fds(self):
    #     if FontDecorators.is_same(self.fr_fd_cls, self.bk_fd_cls):
    #         return True
    #     return False

    @staticmethod
    def cancel_fd_cls(lft_cd, rgt_cd):
        rgt_cd.fr_fd_cls, lft_cd.bk_fd_cls \
            = FontDecorator.cancel_fds(rgt_cd.fr_fd_cls, lft_cd.bk_fd_cls)
        return lft_cd, rgt_cd

    @staticmethod
    def are_consecutive(lft_cd, rgt_cd):
        return FontDecorator.is_same(rgt_cd.fr_fd_cls, lft_cd.bk_fd_cls)

    def apply_style(self, style):
        fr, bk = self.fr_fd_cls, self.bk_fd_cls
        if style is None:
            return
        if style.font is not None:
            if fr.font_name == '' and bk.font_name == '':
                if style.font != Form.mincho_font:
                    fd = FontDecorator.get_font_name_fd(style.font)
                    fr.font_name, bk.font_name = fd, fd
        if style.font_size is not None:
            if fr.font_scale == '' and bk.font_scale == '':
                fd = FontDecorator.get_font_scale_fd(style.font_size)
                if fd is not None:
                    fr.font_name, bk.font_name = fd, fd
        if style.is_italic:
            if fr.italic == '' and bk.italic == '':
                fr.font_name, bk.font_name = '*', '*'
        if style.is_bold:
            if fr.bold == '' and bk.bold == '':
                fr.font_name, bk.font_name = '**', '**'
        if style.has_strike:
            if fr.strike == '' and bk.strike == '':
                fr.font_name, bk.font_name = '~~', '~~'
        if style.has_frame:
            if fr.frame == '' and bk.frame == '':
                fr.font_name, bk.font_name = '[|', '|]'
        if style.underline is not None:
            if fr.underline == '' and bk.underline == '':
                fd = FontDecorator.get_underline_fd(style.underline)
                fr.font_name, bk.font_name = fd, fd
        if style.font_color is not None and style.font_color != 'auto':
            if fr.font_color == '' and bk.font_color == '':
                fd = FontDecorator.get_font_color_fd(style.font_color)
                fr.font_name, bk.font_name = fd, fd


class FontDecorator:

    def __init__(self, fds):
        self.reset_fds()
        self.set_fds(fds)
        self.track_changes = ''     # TRACK CHANGES (-> / <- / +> / <+)

    def reset_fds(self):
        self.font_name = ''         # FONT NAME (` / @.+@)
        # self.font_size = ''         # FONT SIZE
        self.font_scale = ''        # FONT SCALE (@.+@ / --- / -- / ++ / +++)
        self.font_width = ''        # FONT WIDTH (>>> / >> / << / <<<)
        self.italic = ''            # ITALIC (*)
        self.bold = ''              # BOLD (**)
        self.strike = ''            # STRIKETHROUGH (~~)
        self.frame = ''             # FRAME ([| / |])
        self.underline = ''         # UNDERLINE (_.+_)
        self.font_color = ''        # FONT COLOR (^.*^)
        self.highlight_color = ''   # HIGHLIGHT COLOR (_.+_)
        self.sub_or_sup = ''        # SUB OR SUP (_{ / _} / ^{ / ^})
        # self.track_changes = ''     # TRACK CHANGES (-> / <- / +> / <+)

    @staticmethod
    def get_font_name(afont, jfont):
        f_afont = re.sub('\\s*/.*$', '', Form.mincho_font)
        f_jfont = re.sub('^.*/\\s*', '', Form.mincho_font)
        if afont is None or afont == '':
            afont = f_afont
        if jfont is None or jfont == '':
            jfont = f_jfont
        for mfs in MS_FONTS:
            if afont in mfs:
                afont = mfs[0]
            if jfont in mfs:
                jfont = mfs[0]
        if jfont == Form.ivs_font:
            return jfont
        if afont == jfont:
            return '= / ' + jfont
        return afont + ' / ' + jfont

    @staticmethod
    def get_font_name_fd(name):
        if name is None:
            return None
        elif name == Form.mincho_font:
            return None
        elif name == Form.gothic_font:
            return '`'
        elif re.match('^=\\s*/\\s*', name):
            return '@' + re.sub('^=\\s*/\\s*', '', name) + '@'
        else:
            return '@' + name + '@'

    @staticmethod
    def get_font_scale_fd(size):
        if size is None:
            return None
        elif size < Form.font_size * 0.4:     # changed from "0.5" to "0.4"
            if size.is_integer():
                size = int(size)
            return '@' + str(size) + '@'
        elif size < Form.font_size * 0.7:
            return '---'
        elif size < Form.font_size * 0.9:
            return '--'
        elif size <= Form.font_size * 1.1:
            return None
        elif size <= Form.font_size * 1.3:
            return '++'
        elif size <= Form.font_size * 1.6:  # changed from "1.5" to "1.6"
            return '+++'
        else:
            if size.is_integer():
                size = int(size)
            return '@' + str(size) + '@'

    @staticmethod
    def get_font_width_fd(width):
        if width is None:
            return None, None
        elif width < 70:
            return '>>>', '<<<'
        elif width < 90:
            return '>>', '<<'
        elif width > 130:
            return '<<<', '>>>'
        elif width > 110:
            return '<<', '>>'
        else:
            return None, None

    @staticmethod
    def get_underline_fd(underline):
        if underline is None:
            return None
        elif underline == '':
            return '__'
        elif underline in UNDERLINE:
            return '_' + UNDERLINE[underline] + '_'
        return None

    @staticmethod
    def get_font_color_fd(color):
        if color is None:
            return None
        color = color.upper()
        if color == '000000':
            return None
        if color == 'FFFFFF':
            return '^^'
        if color in FONT_COLOR:
            return '^' + FONT_COLOR[color] + '^'
        # res = '^(00|11|22|33|44|55|66|77|88|99|AA|BB|CC|DD|EE|FF){3}$'
        # if re.match(res, color):
        #     return '^' +  re.sub('^.(.).(.).(.)$', '\\1\\2\\3', color) + '^'
        return '^' + color + '^'

    @staticmethod
    def get_highlight_color_fd(color):
        if color is None:
            return None
        return '_' + color + '_'

    def set_fds(self, fds):
        for fd in fds:
            self.__set_fd(fd)

    def __set_fd(self, fd_str):
        if re.match('^`$', fd_str):
            self.font_name = fd_str        # FONT NAME (`)
        elif (re.match('^@.+@$', fd_str) and
              not re.match('^@' + RES_NUMBER + '@$', fd_str)):
            self.font_name = fd_str        # FONT NAME (@.+@)
        elif re.match('^@' + RES_NUMBER + '@$', fd_str):
            self.font_scale = fd_str       # FONT SCALE (@.+@)
        elif re.match('^\\-\\-\\-|\\-\\-|\\+\\+|\\+\\+\\+$', fd_str):
            self.font_scale = fd_str       # FONT SCALE (--- / -- / ++ / +++)
        elif re.match('^>>>|>>|<<|<<<$', fd_str):
            self.font_width = fd_str       # FONT WIDTH (>>> / >> / << / <<<)
        elif fd_str == '*':
            self.italic = fd_str           # ITALIC (*)
        elif fd_str == '**':
            self.bold = fd_str             # BOLD (**)
        elif fd_str == '***':
            self.italic = '*'              # ITALIC (*)
            self.bold = '**'               # BOLD (**)
        elif fd_str == '~~':
            self.strike = fd_str           # STRIKETHROUGH (~~)
        elif fd_str == '[|' or fd_str == '|]':
            self.frame = fd_str            # FRAME ([| / |])
        elif re.match('_[\\$=\\.#\\-~\\+]{,4}_', fd_str):
            self.underline = fd_str        # UNDERLINE (_.+_)
        elif re.match('\\^[0-9A-Za-z]{0,11}\\^', fd_str):
            self.font_color = fd_str       # FONT COLOR (^.*^)
        elif re.match('_[0-9A-Za-z]{1,11}_', fd_str):
            self.highlight_color = fd_str  # HIGHLIGHT COLOR (_.+_)
        elif re.match('^_{|_}|\\^{|\\^}$', fd_str):
            self.sub_or_sup = fd_str       # SUB OR SUP (_{ / _} / ^{ / ^})
        elif re.match('^\\->|<\\-|\\+>|<\\+$', fd_str):
            self.track_changes = fd_str    # TRACK CHANGES (-> / <- / +> / <+)

    # def is_empty(self):
    #     if self.get_ord_fds() != '':
    #         return False
    #     if self.get_rev_fds() != '':
    #         return False
    #     return True

    def get_ord_fds(self):
        return ''.join(self.__get_ordered_list())

    def get_rev_fds(self):
        return ''.join(self.__get_ordered_list()[::-1])

    def __get_ordered_list(self):
        return [self.font_name,
                # self.font_size,
                self.font_scale,
                self.font_width,
                self.italic,
                self.bold,
                self.strike,
                self.frame,
                self.underline,
                self.font_color,
                self.highlight_color,
                re.sub('^[_\\^]}$', '}', self.sub_or_sup),
                self.track_changes]

    @staticmethod
    def get_partner(fd_in: str) -> str:
        partners = [['>>>', '<<<'], ['>>', '<<'], ['<<', '>>'], ['<<<', '>>>'],
                    ['[|', '|]'],
                    ['_{', '_}'], ['^{', '^}'],
                    ['->', '<-'], ['+>', '<+']]
        for p in partners:
            if fd_in == p[0]:
                fd_out = p[1]
                return fd_out
            if fd_in == p[1]:
                fd_out = p[0]
                return fd_out
        fd_out = fd_in
        return fd_out

    @staticmethod
    def escape_chars(fd_in: str) -> str:
        fd_out = fd_in
        fd_out = re.sub('\\*', '\\\\*', fd_out)
        fd_out = re.sub('\\+', '\\\\+', fd_out)
        fd_out = re.sub('\\^', '\\\\^', fd_out)
        fd_out = re.sub('\\|', '\\\\|', fd_out)
        fd_out = re.sub('\\[', '\\\\[', fd_out)
        fd_out = re.sub('\\]', '\\\\]', fd_out)
        return fd_out

    @staticmethod
    def is_same(fr_fd_cls, bk_fd_cls):
        fr, bk = fr_fd_cls, bk_fd_cls
        if fr.font_name != bk.font_name:
            return False
        # if fr.font_size != bk.font_size:
        #     return False
        if fr.font_scale != bk.font_scale:
            return False
        if fr.font_width != '' or bk.font_width != '':
            if fr.font_width != '>>>' or bk.font_width != '<<<':
                if fr.font_width != '>>' or bk.font_width != '<<':
                    if fr.font_width != '<<' or bk.font_width != '>>':
                        if fr.font_width != '<<<' or bk.font_width != '>>>':
                            return False
        if fr.italic != bk.italic:
            return False
        if fr.bold != bk.bold:
            return False
        if fr.strike != bk.strike:
            return False
        if fr.frame != bk.frame:
            return False
        if fr.underline != bk.underline:
            return False
        if fr.font_color != bk.font_color:
            return False
        if fr.highlight_color != bk.highlight_color:
            return False
        if fr.sub_or_sup != '' or bk.sub_or_sup != '':
            if fr.sub_or_sup != '_{' or bk.sub_or_sup != '_}':
                if fr.sub_or_sup != '^{' or bk.sub_or_sup != '^}':
                    return False
        if fr.track_changes != '' or bk.track_changes != '':
            if fr.track_changes != '->' or bk.track_changes != '<-':
                if fr.track_changes != '+>' or bk.track_changes != '<+':
                    return False
        return True

    @staticmethod
    def cancel_fds(fr_fd_cls, bk_fd_cls):
        fr, bk = fr_fd_cls, bk_fd_cls
        if fr.font_name == bk.font_name:
            fr.font_name, bk.font_name = '', ''
        # if fr.font_size == bk.font_size:
        #     fr.font_size, bk.font_size = '', ''
        if fr.font_scale == bk.font_scale:
            fr.font_scale, bk.font_scale = '', ''
        if fr.font_width == '<<<' and bk.font_width == '>>>' or \
           fr.font_width == '<<' and bk.font_width == '>>' or \
           fr.font_width == '>>' and bk.font_width == '<<' or \
           fr.font_width == '>>>' and bk.font_width == '<<<':
            fr.font_width, bk.font_width = '', ''
        if fr.italic == bk.italic:
            fr.italic, bk.italic = '', ''
        if fr.bold == bk.bold:
            fr.bold, bk.bold = '', ''
        if fr.strike == bk.strike:
            fr.strike, bk.strike = '', ''
        if fr.frame == bk.frame:
            fr.frame, bk.frame = '', ''
        if fr.underline == bk.underline:
            fr.underline, bk.underline = '', ''
        if fr.font_color == bk.font_color:
            fr.font_color, bk.font_color = '', ''
        if fr.highlight_color == bk.highlight_color:
            fr.highlight_color, bk.highlight_color = '', ''
        if (fr.sub_or_sup == '_{' and bk.sub_or_sup == '_}') or \
           (fr.sub_or_sup == '^{' and bk.sub_or_sup == '^}'):
            fr.sub_or_sup, bk.sub_or_sup = '', ''
        if (fr.track_changes == '->' and bk.track_changes == '<-') or \
           (fr.track_changes == '+>' and bk.track_changes == '<+'):
            fr.track_changes, bk.track_changes = '', ''
        return fr_fd_cls, bk_fd_cls


class MathDatum:

    """A class to keep math characters data"""

    def __init__(self):
        self.chars = ''
        # FONT NAME (NOT IMPLEMENTED)
        # FONT SIZE AND SCALE (s-4, s-3, s-2, s-1, s+1, s+2, s+3, s+4, s+5)
        # FONT WIDTH (w-4, w-3, w-2, w-1, w+1, w+2, w+3, w+4, w+5)
        # ROMAN (r)
        # BOLD (b)
        # STRIKETHROUGH (s)
        # FRAME (f)
        # UNDERLINE (u)
        # FONT COLOR (c=...)
        # HIGILIGHT COLOR (h=...)
        # DEL OR INS (d, d, i, i)
        self.fr_fd_lst = []
        self.bk_fd_lst = []
        self.is_math_function = False

    def set_chars(self, raw_chars):
        self.chars = raw_chars

    def append_fr_and_bk_fds(self, fr_fd_str, bk_fd_str):
        if fr_fd_str != '':
            if re.match('^[swrbsuchdi]', fr_fd_str):
                if fr_fd_str not in self.fr_fd_lst:
                    self.fr_fd_lst.append(fr_fd_str)
            else:
                if True:
                    self.fr_fd_lst.append(fr_fd_str)
        if bk_fd_str != '':
            if re.match('^[swrbsuchdi]', fr_fd_str):
                if bk_fd_str not in self.bk_fd_lst:
                    self.bk_fd_lst.append(bk_fd_str)
            else:
                if True:
                    self.bk_fd_lst.append(bk_fd_str)

    def remove_fr_and_bk_fds(self, fr_fd_str, bk_fd_str):
        if fr_fd_str != '':
            while fr_fd_str in self.fr_fd_lst:
                self.fr_fd_lst.remove(fr_fd_str)
        if bk_fd_str != '':
            while bk_fd_str in self.bk_fd_lst:
                self.bk_fd_lst.remove(bk_fd_str)

    # def reset_fds(self):
    #     self.fr_fd_lst = []
    #     self.bk_fd_lst = []

    def is_empty(self):
        if self.chars == '':
            if self.fr_fd_lst == []:
                if self.bk_fd_lst == []:
                    return True
        return False

    def get_chars_with_fds(self):
        chars_with_fds \
            = self.__get_fr_chars() + self.chars + self.__get_bk_chars()
        return chars_with_fds

    def __get_fr_chars(self):
        fr_chars = ''
        for fd in self.fr_fd_lst:
            if False:
                pass
            elif fd == 's-4':
                fr_chars += '{\\tiny{'
            elif fd == 's-3':
                fr_chars += '{\\scriptsize{'
            elif fd == 's-2':
                fr_chars += '{\\footnotesize{'
            elif fd == 's-1':
                fr_chars += '{\\small{'
            elif fd == 's+1':
                fr_chars += '{\\large{'
            elif fd == 's+2':
                fr_chars += '{\\Large{'
            elif fd == 's+3':
                fr_chars += '{\\LARGE{'
            elif fd == 's+4':
                fr_chars += '{\\huge{'
            elif fd == 's+5':
                fr_chars += '{\\Huge{'
            elif fd == 'w-4':
                fr_chars += '{\\scalebox{0.2}[1]{'
            elif fd == 'w-3':
                fr_chars += '{\\scalebox{0.4}[1]{'
            elif fd == 'w-2':
                fr_chars += '{\\scalebox{0.6}[1]{'
            elif fd == 'w-1':
                fr_chars += '{\\scalebox{0.8}[1]{'
            elif fd == 'w+1':
                fr_chars += '{\\scalebox{1.2}[1]{'
            elif fd == 'w+2':
                fr_chars += '{\\scalebox{1.4}[1]{'
            elif fd == 'w+3':
                fr_chars += '{\\scalebox{1.6}[1]{'
            elif fd == 'w+4':
                fr_chars += '{\\scalebox{1.8}[1]{'
            elif fd == 'w+5':
                fr_chars += '{\\scalebox{2.0}[1]{'
            elif fd == 'r':
                fr_chars += '{\\mathrm{'
            elif fd == 'b':
                fr_chars += '{\\mathbf{'
            elif fd == 's':
                fr_chars += '{\\sout{'
            elif fd == 'f':
                fr_chars += '{\\boxed{'
            elif fd == 'u':
                fr_chars += '{\\underline{'
            elif re.match('^c=.*$', fd):
                c = re.sub('^c=', '', fd)
                fr_chars += '{\\textcolor{' + c + '}{'
            elif re.match('^h=.*$', fd):
                c = re.sub('^h=', '', fd)
                fr_chars += '{\\colorbox{' + c + '}{'
            elif fd == 'd':
                fr_chars += '->'
            elif fd == 'i':
                fr_chars += '+>'
            else:
                fr_chars += fd
        return fr_chars

    def __get_bk_chars(self):
        bk_chars = ''
        for fd in self.bk_fd_lst[::-1]:
            if False:
                pass
            elif re.match('^s-[1-4]$', fd) or re.match('^s\\+[1-5]$', fd):
                bk_chars += '}}'  # size
            elif re.match('w-[1-4]', fd) or re.match('w\\+[1-5]', fd):
                bk_chars += '}}'  # width
            elif fd == 'r':
                bk_chars += '}}'  # roman
            elif fd == 'b':
                bk_chars += '}}'  # bold
            elif fd == 's':
                bk_chars += '}}'  # strikethrough
            elif fd == 'f':
                bk_chars += '}}'  # frame
            elif fd == 'u':
                bk_chars += '}}'  # underline
            elif re.match('^c=.*$', fd):
                bk_chars += '}}'  # fort color
            elif re.match('^h=.*$', fd):
                bk_chars += '}}'  # highlight color
            elif fd == 'd':
                bk_chars += '<-'  # delete
            elif fd == 'i':
                bk_chars += '<+'  # insert
            elif fd == '_}' or fd == '^}':
                bk_chars += '}'  # sub or sup
            else:
                bk_chars += fd
        return bk_chars

    @staticmethod
    def cancel_fd_lst(math_data):
        if len(math_data) == 0:
            return math_data
        for i in range(len(math_data)):
            if i == 0:
                continue
            bk_fd_lst = math_data[i - 1].bk_fd_lst
            fr_fd_lst = math_data[i].fr_fd_lst
            for fd in fr_fd_lst[::-1]:
                if fd in bk_fd_lst:
                    if re.match('^[swrbsuchdi]', fd):
                        while fd in fr_fd_lst:
                            fr_fd_lst.remove(fd)
                        while fd in bk_fd_lst:
                            bk_fd_lst.remove(fd)
        return math_data

    @classmethod
    def get_math_data(cls, xl, math_data):
        f_size = Form.font_size
        # BEGINNING
        if re.match('^<m:oMath>$', xl):
            math_data = []
            math_data.append(MathDatum())
            math_data.append(MathDatum())
            return math_data, None
        # END
        elif re.match('^</m:oMath>$', xl):
            if math_data[0].is_empty():
                math_data.pop(0)
            if math_data[-1].is_empty():
                math_data.pop(-1)
            math_chars_datum = cls._get_math_chars_datum(math_data)
            math_data = None
            return None, math_chars_datum
        if math_data is None:
            return None, None
        md_pre = math_data[-2]
        md_cur = math_data[-1]
        # FONT NAME (NOT IMPLEMENTED)
        # FONT SIZE AND SCALE
        v = XML.get_value('w:sz', 'w:val', -1.0, xl)
        # (FOR COMPLEX SCRIPT)
        v = XML.get_value('w:szCs', 'w:val', v, xl)
        if v > 0:
            s = round(v / 2, 1)
            if s < f_size * 0.3:
                md_cur.append_fr_and_bk_fds('s-4', 's-4')  # tiny
            elif s < f_size * 0.5:
                md_cur.append_fr_and_bk_fds('s-3', 's-3')  # scriptsize
            elif s < f_size * 0.7:
                md_cur.append_fr_and_bk_fds('s-2', 's-2')  # footnotesize
            elif s < f_size * 0.9:
                md_cur.append_fr_and_bk_fds('s-1', 's-1')  # small
            elif s <= f_size * 1.1:
                pass                                       # normalsize
            elif s <= f_size * 1.3:
                md_cur.append_fr_and_bk_fds('s+1', 's+1')  # large
            elif s <= f_size * 1.5:
                md_cur.append_fr_and_bk_fds('s+2', 's+2')  # Large
            elif s <= f_size * 1.7:
                md_cur.append_fr_and_bk_fds('s+3', 's+3')  # LARGE
            elif s <= f_size * 1.9:
                md_cur.append_fr_and_bk_fds('s+4', 's+4')  # huge
            else:
                md_cur.append_fr_and_bk_fds('s+5', 's+5')  # Huge
            return math_data, None
        # FONT WIDTH
        v = XML.get_value('w:w', 'w:val', -1.0, xl)
        if v > 0:
            if v < 30:
                md_cur.append_fr_and_bk_fds('w-4', 'w-4')
            elif v < 50:
                md_cur.append_fr_and_bk_fds('w-3', 'w-3')
            elif v < 70:
                md_cur.append_fr_and_bk_fds('w-2', 'w-2')
            elif v < 90:
                md_cur.append_fr_and_bk_fds('w-1', 'w-1')
            elif v <= 110:
                pass
            elif v <= 130:
                md_cur.append_fr_and_bk_fds('w+1', 'w+1')
            elif v <= 150:
                md_cur.append_fr_and_bk_fds('w+2', 'w+2')
            elif v <= 170:
                md_cur.append_fr_and_bk_fds('w+3', 'w+3')
            elif v <= 190:
                md_cur.append_fr_and_bk_fds('w+4', 'w+4')
            else:
                md_cur.append_fr_and_bk_fds('w+5', 'w+5')
            return math_data, None
        # BOLD OR ROMAN
        v = XML.get_value('m:sty', 'm:val', '', xl)
        if v != '':
            # ROMAN
            if v == 'p' or v == 'b':
                md_cur.append_fr_and_bk_fds('r', 'r')
            # BOLD
            if v == 'bi' or v == 'b':
                md_cur.append_fr_and_bk_fds('b', 'b')
            return math_data, None
        # STRIKETHROUGH
        if re.match('^<w:strike/?>$', xl):
            md_cur.append_fr_and_bk_fds('s', 's')
            return math_data, None
        # FRAME
        if re.match('^<w:bdr( .*)?/?>$', xl):
            md_cur.append_fr_and_bk_fds('f', 'f')
            return math_data, None
        # UNDERLINE
        v = XML.get_value('w:u', 'w:val', '', xl)
        if v != '':
            md_cur.append_fr_and_bk_fds('u', 'u')
            return math_data, None
        # FONT COLOR
        v = XML.get_value('w:color', 'w:val', '', xl)
        if v != '':
            if v == 'FFFFFF':
                c = 'white'
            elif v in FONT_COLOR:
                c = FONT_COLOR[v]
            else:
                c = v
            md_cur.append_fr_and_bk_fds('c=' + c, 'c=' + c)
            return math_data, None
        # HIGILIGHT COLOR
        v = XML.get_value('w:highlight', 'w:val', '', xl)
        if v != '':
            md_cur.append_fr_and_bk_fds('h=' + v, 'h=' + v)
            return math_data, None
        # DEL OR INS
        if xl == '<w:del>':
            md_cur.append_fr_and_bk_fds('d', '')
            return math_data, None
        if xl == '</w:del>':
            md_cur.append_fr_and_bk_fds('', 'd')
            return math_data, None
        elif xl == '<w:ins>':
            md_cur.append_fr_and_bk_fds('i', '')
            return math_data, None
        elif xl == '</w:ins>':
            md_cur.append_fr_and_bk_fds('', 'i')
            return math_data, None
        # ELEMENT, RUN PROPERTY, RUN, TEXT, LINE BREAK, FUNCTION NAME
        # --------------------------------------------------
        # <m:e> or <m:fName>
        #     <m:r>
        #         <m:rPr/>
        #           or
        #         <m:rPr>...</m:rPr>
        #         A
        #     </m:r>
        # </m:e> or </m:fName>
        # --------------------------------------------------
        # ELEMENT
        if xl == '<m:e>':
            md_cur.append_fr_and_bk_fds('{', '')
            return math_data, None
        if xl == '</m:e>':
            md_pre.append_fr_and_bk_fds('', '}')
            return math_data, None
        if xl == '<m:e/>':
            md_cur.append_fr_and_bk_fds('{}', '')
            return math_data, None
        # RUN PROPERTY (SAVE FUNCTION NAME)
        if (xl == '</w:rPr>' or xl == '<w:rPr/>') and md_cur.chars != '':
            math_data.append(MathDatum())
            return math_data, None
        # RUN (SAVE TEXT)
        if xl == '</m:r>' or xl == '<m:r/>':
            math_data.append(MathDatum())
            return math_data, None
        # TEXT
        if not re.match('^<.*>$', xl):
            # FUNCTION (lim)
            if md_cur.is_math_function:
                md_cur.chars += '\\'
                md_cur.remove_fr_and_bk_fds('r', 'r')
            xl = re.sub('{', '\\{', xl)    # "{" -> "\{"
            xl = re.sub('}', '\\}', xl)    # "}" <- "\}"
            xl = re.sub(' ', '\\\\,', xl)  # " " -> "\," (space)
            md_cur.chars += xl
            return math_data, None
        # LINE BREAK
        if re.match('^<m:brk( .*)?/>$', xl):
            md_cur.chars += '\\\\'
            return math_data, None
        # FUNCTION NAME
        if xl == '<m:fName>':
            md_cur.is_math_function = True
            return math_data, None
        if xl == '</m:fName>':
            # md_pre.is_math_function = False
            return math_data, None
        # SUP AND SUB
        # --------------------------------------------------
        # <m:e>
        #     <m:r>
        #         A
        #     </m:r>
        # </m:e>
        # <m:sub>  # SUBSCIPT
        #   or
        # <m:sup>  # SUPERSCRIPT
        #     <m:r>
        #         B
        #     </m:r>
        # </m:sub>  # SUBSCIPT
        #   or
        # </m:sup>  # SUPERSCRIPT
        # = A_{B}  # SUBSCIPT
        #   or
        # = A^{B}  # SUPERSCRIPT
        # --------------------------------------------------
        # <m:sPre>
        #     <m:sub>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:sub>
        #     <m:sup/>
        #     <m:e>
        #         <m:r>
        #             B
        #         </m:r>
        #     </m:e>
        #     <m:sub>
        #         <m:r>
        #             C
        #         </m:r>
        #     </m:sub>
        # </m:sPre>
        # = {}_{A}B_{C}
        # --------------------------------------------------
        if xl == '<m:sPre>':
            md_cur.append_fr_and_bk_fds('{}', '')  # "{}_{...}" or "{}^{...}"
            return math_data, None
        if xl == '</m:sPre>':
            return math_data, None
        if xl == '<m:sub>':
            md_cur.append_fr_and_bk_fds('_{', '')
            return math_data, None
        if xl == '</m:sub>':
            md_pre.append_fr_and_bk_fds('', '_}')
            return math_data, None
        if xl == '<m:sup>':
            md_cur.append_fr_and_bk_fds('^{', '')
            return math_data, None
        if xl == '</m:sup>':
            md_pre.append_fr_and_bk_fds('', '^}')
            return math_data, None
        # VECTOR
        # --------------------------------------------------
        # <m:acc>
        #     <m:chr m:val="→"/>  # MS OFFICE (\u2192)
        #       or
        #     <m:chr m:val="⃗"/>    # LIBREOFFICE (\u20D7)
        #     <w:rPr/> or <w:rPr>...</w:rPr>
        #     <m:e>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:e>
        # </m:acc>
        # = \vec{A}
        # --------------------------------------------------
        if re.match('<m:chr m:val="(\u2192|\u20D7)"/>', xl):
            md_cur.chars += '\\vec'
            return math_data, None
        # DOT
        # --------------------------------------------------
        # <m:acc>
        #     <m:chr m:val="̇"/>  # \u0307 / \u0308 / \u20DB
        #     <w:rPr/> or <w:rPr>...</w:rPr>
        #     <m:e>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:e>
        # </m:acc>
        # = \dot{A}
        # --------------------------------------------------
        if re.match('<m:chr m:val="(\u0307)"/>', xl):
            md_cur.chars += '\\dot'
            return math_data, None
        if re.match('<m:chr m:val="(\u0308)"/>', xl):
            md_cur.chars += '\\ddot'
            return math_data, None
        if re.match('<m:chr m:val="(\u20DB)"/>', xl):
            md_cur.chars += '\\dddot'
            return math_data, None
        # FRACTION, BINOMIAL
        # --------------------------------------------------
        # <m:f>
        #     (none)                   # FRACTION
        #       or
        #     <m:type m:val="noBar"/>  # BINOMIAL
        #     <w:rPr/> or <w:rPr>...</w:rPr>
        #     <m:num>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:num>
        #     <m:den>
        #         <m:r>
        #             B
        #         </m:r>
        #     </m:den>
        # </m:f>
        # = \frac{A}{B}
        #   or
        # = \binom{A}{B}
        # --------------------------------------------------
        if xl == '<m:f>':
            md_cur.chars += '\\frac'
            return math_data, None
        if xl == '<m:type m:val="noBar"/>':
            if md_cur.chars == '\\frac':
                md_cur.chars = '\\Xbinom'
            return math_data, None
        if xl == '</m:f>':
            return math_data, None
        # NUMERATOR
        if xl == '<m:num>':
            md_cur.append_fr_and_bk_fds('{', '')
            return math_data, None
        if xl == '</m:num>':
            md_pre.append_fr_and_bk_fds('', '}')
            return math_data, None
        # DENOMINATOR
        if xl == '<m:den>':
            md_cur.append_fr_and_bk_fds('{', '')
            return math_data, None
        if xl == '</m:den>':
            md_pre.append_fr_and_bk_fds('', '}')
            return math_data, None
        # RADICAL ROOT
        # --------------------------------------------------
        # <m:rad>
        #     <w:rPr/> or <w:rPr>...</w:rPr>
        #     <m:deg/>
        #       or
        #     <m:deg>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:deg>
        #     <m:e>
        #         <m:r>
        #             B
        #         </m:r>
        #     </m:e>
        # </m:rad>
        # = \sqrt{B}
        #   or
        # = \sqrt[A]{B}
        # --------------------------------------------------
        if xl == '<m:rad>':
            md_cur.chars += '\\sqrt'
            return math_data, None
        if xl == '<m:deg>':
            md_cur.append_fr_and_bk_fds('[', '')
            return math_data, None
        if xl == '</m:deg>':
            md_pre.append_fr_and_bk_fds('', ']')
            return math_data, None
        # LIMIT
        # --------------------------------------------------
        # <m:fName>
        #     <m:e>
        #         <m:r>
        #             lim
        #         </m:r>
        #     </m:e>
        #     <m:lim/>
        #       or
        #     <m:lim>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:lim>
        #     <m:e>
        #         <m:r>
        #             B
        #         </m:r>
        #     </m:e>
        # </m:fName>
        # = \lim{B}
        #   or
        # = \lim_{A}{B}
        # --------------------------------------------------
        if xl == '<m:lim>':
            md_cur.append_fr_and_bk_fds('_{', '')
            # "{\lim}" -> "\lim"
            if md_pre.chars == '\\lim':
                if '{' in md_pre.fr_fd_lst:
                    md_pre.fr_fd_lst.remove('{')
                if '}' in md_pre.bk_fd_lst:
                    md_pre.bk_fd_lst.remove('}')
            return math_data, None
        if xl == '</m:lim>':
            md_pre.append_fr_and_bk_fds('', '}')
            return math_data, None
        # INTEGRAL, DOUBLE INTEGRAL, TRIPLE INTEGRAL, LINE INTEGRAL
        # --------------------------------------------------
        # <m:nary>
        #     (none)
        #       or
        #     <m:chr m:val="∬"/>  # \u222C
        #       or
        #     <m:chr m:val="∭"/>  # \u222D
        #       or
        #     <m:chr m:val="∭"/>  # \u222D
        #       or
        #     <m:chr m:val="∮"/>  # \u222e
        #     <w:rPr/> or <w:rPr>...</w:rPr>
        #     <m:sub/>
        #       or
        #     <m:sub>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:sub>
        #     <m:sup/>
        #       or
        #     <m:sup>
        #         <m:r>
        #             B
        #         </m:r>
        #     </m:sup>
        #     <m:e>
        #         <m:r>
        #             C
        #         </m:r>
        #     </m:e>
        # </m:nary>
        # = \int{A}
        #   or
        # = \int{C}
        #   or
        # = \int_{A}^{B}{C}
        #   or
        # = \iint{A}
        #   or
        # = \iint_{A}^{B}{C}
        #   or
        # = \iiint{A}
        #   or
        # = \iiint_{A}^{B}{C}
        #   or
        # = \oint{A}
        #   or
        # = \oint_{A}^{B}{C}
        # --------------------------------------------------
        if xl == '<m:nary>':
            md_cur.chars += '\\int'
            return math_data, None
        if xl == '<m:chr m:val="∬"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\iint', md_cur.chars)
            return math_data, None
        if xl == '<m:chr m:val="∭"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\iiint', md_cur.chars)
            return math_data, None
        if xl == '<m:chr m:val="∮"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\oint', md_cur.chars)
            return math_data, None
        # SIGMA, PRODUCT
        # --------------------------------------------------
        # <m:nary>
        #     <m:chr m:val="∑"/>  # \u2211
        #       or
        #     <m:chr m:val="∏"/>  # \u220F
        #     <w:rPr/> or <w:rPr>...</w:rPr>
        #     <m:subHide m:val="1"/>
        #     <m:supHide m:val="1"/>
        #       or
        #     <m:sub>
        #         <m:r>
        #             A
        #         </m:r>
        #     </m:sub>
        #     <m:sup>
        #         <m:r>
        #             B
        #         </m:r>
        #     </m:sup>
        #     <m:e>
        #         <m:r>
        #             C
        #         </m:r>
        #     </m:e>
        # </m:nary>
        # = \sum{C}
        #   or
        # = \sum_{A}^{B}{C}
        #   or
        # = \prod{C}
        #   or
        # = \prod_{A}^{B}{C}
        # --------------------------------------------------
        if xl == '<m:chr m:val="∑"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\sum', md_cur.chars)
            return math_data, None
        if xl == '<m:chr m:val="∏"/>':
            md_cur.chars = re.sub('\\\\int$', '\\\\prod', md_cur.chars)
            return math_data, None
        # MATRIX
        # --------------------------------------------------
        # <m:d>
        #     (NONE)
        #     or
        #     <m:begChr m:val="["/>
        #     <m:endChr m:val="]"/>
        #     <w:rPr/> or <w:rPr>...</w:rPr>
        #     <m:e>
        #         <m:m>
        #             <m:mr>
        #                 <m:e>
        #                     <m:r>
        #                         A
        #                     </m:r>
        #                 </m:e>
        #                 <m:e>
        #                     <m:r>
        #                         B
        #                     </m:r>
        #                 </m:e>
        #             </m:mr>
        #             <m:mr>
        #                 <m:e>
        #                     <m:r>
        #                         C
        #                     </m:r>
        #                 </m:e>
        #                 <m:e>
        #                     <m:r>
        #                         D
        #                     </m:r>
        #                 </m:e>
        #             </m:mr>
        #         </m:m>
        #     </m:e>
        # </m:d>
        # = \begin{pmatrix}A&B\\C&D\\\end{pmatrix}
        # or
        # = \begin{bmatrix}A&B\\C&D\\\end{bmatrix}
        # --------------------------------------------------
        # MATRIX (BEGINNING)
        if xl == '<m:m>':
            md_cur.chars += '\\Xbegin{matrix}'
            return math_data, None
        # MATRIX (END)
        if xl == '</m:m>':
            md_cur.chars += '\\Xend{matrix}'
            math_data.append(MathDatum())
            return math_data, None
        # MATRIX (PARENTHESES BEGINNING)
        if xl == '<m:d>':
            md_cur.chars += '(<()>'
            return math_data, None
        if re.match('^<m:begChr m:val="(.?)"/>$', xl):
            bc = re.sub('^<m:begChr m:val="(.?)"/>$', '\\1', xl)
            md_cur.chars = re.sub('\\(<(.)(.?)>$', '(<' + bc + '\\2>',
                                  md_cur.chars)
            return math_data, None
        if re.match('<m:endChr m:val="(.?)"/>', xl):
            ec = re.sub('^<m:endChr m:val="(.?)"/>$', '\\1', xl)
            md_cur.chars = re.sub('\\(<(.?)(.)>$', '(<\\g<1>' + ec + '>',
                                  md_cur.chars)
            return math_data, None
        # MATRIX (PARENTHESES END)
        if xl == '</m:d>':
            res = '(.*)\\(<(.?)(.?)>(.*)$'
            end = ')'
            md_beg_chars = None
            for i in range(len(math_data) - 1, -1, -1):
                if re.match(res, math_data[i].chars):
                    # CHARS
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
                    # FRONT AND BACK LIST
                    for fd in math_data[i].fr_fd_lst:
                        if fd != '{':
                            md_cur.append_fr_and_bk_fds(fd, '')
                    for fd in math_data[i].bk_fd_lst:
                        if fd != '}':
                            md_cur.append_fr_and_bk_fds('', fd)
                    break
            md_cur.chars += end
            math_data.append(MathDatum())  # no "<w:rPr/>" or "</w:rPr>"
            return math_data, None
        # BREAK ROW
        if xl == '</m:mr>':
            md_cur.fr_fd_lst.insert(0, '\\\\')
            return math_data, None
        return math_data, None

    @classmethod
    def _get_math_chars_datum(cls, math_data):
        math_data = cls.cancel_fd_lst(math_data)
        fr_fd_cls, math_data[0].fr_fd_lst \
            = MathDatum.extract_fr_fd_cls(math_data[0].fr_fd_lst)
        bk_fd_cls, math_data[-1].bk_fd_lst \
            = MathDatum.extract_bk_fd_cls(math_data[-1].bk_fd_lst)
        math_str = ''
        for md in math_data:
            math_str += md.get_chars_with_fds()
        math_str = cls._shape_math_matrix(math_str)
        math_str = cls._shape_math_binomial(math_str)
        math_str = cls._shape_sub_and_sup(math_str)
        math_str = re.sub('\\\\mathrm{([=\\-\\+\\±])}', '\\1', math_str)
        math_str \
            = re.sub('\\\\mathrm{(' + RES_NUMBER + ')}', '\\1', math_str)
        math_str = re.sub('{([=\\-\\+\\±])}', '\\1', math_str)
        math_chars_datum = CharsDatum([], '\\[' + math_str + '\\]', [])
        math_chars_datum.fr_fd_cls = fr_fd_cls
        math_chars_datum.bk_fd_cls = bk_fd_cls
        return math_chars_datum

    @staticmethod
    def _shape_math_matrix(math_str):
        # FONT DECRATIONS
        res = '{.*\\\\Xbegin{matrix}.*\\\\Xend{matrix}}{.*\\)}+'
        math_str = MathDatum.shift_paren('\\(', 2, res, math_str)
        math_str = MathDatum.cancel_multi_paren(math_str)
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
    def shift_paren(com, cnt, res, math_str):
        res_com = NOT_ESCAPED + '(' + com + ')(}+)$'
        tmp = ''
        while tmp != math_str:
            tmp = math_str
            tj = -1
            for j in range(len(math_str)):
                if re.match(res_com, math_str[:j]) and math_str[j] != '}':
                    tj = j
                    break
            if tj == -1:
                break
            tk = -1
            dep = []
            d = 0
            for k in range(tj, len(math_str)):
                if math_str[k] == '{':
                    d += 1
                if math_str[k] == '}':
                    d -= 1
                dep.append(d)
                if cnt == -1 and re.match(res, math_str[tj:k]):
                    tk = k
                    break
                if dep.count(0) == cnt and re.match(res, math_str[tj:k]):
                    tk = k
                    break
            if tk == -1:
                break
            pre_bpa_fds_com_epa = math_str[:tj]
            pre_bpa_fds = re.sub(res_com, '\\1', pre_bpa_fds_com_epa)
            com = re.sub(res_com, '\\2', pre_bpa_fds_com_epa)
            epa = re.sub(res_com, '\\3', pre_bpa_fds_com_epa)
            arg = math_str[tj:tk]
            pos = math_str[tk:]
            ti = -1
            d = - len(epa)
            for i in range(len(pre_bpa_fds) - 1, -1, -1):
                if math_str[i] == '{':
                    d += 1
                if math_str[i] == '}':
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
            math_str = pre_bpa_fds + com + arg + epa + pos
        return math_str

    @staticmethod
    def cancel_multi_paren(math_str):
        # {{..}} -> {}
        rm = []
        for i in range(len(math_str) - 1):
            if math_str[i] != '{' or math_str[i + 1] != '{':
                continue
            dep = [0]
            d = 0
            for j in range(i, len(math_str)):
                if math_str[j] == '{':
                    d += 1
                if math_str[j] == '}':
                    d -= 1
                dep.append(d)
                if d == 0:
                    if math_str[j - 1] == '}' or math_str[j] == '}':
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
        u = list(math_str)
        for r in rm:
            u.pop(r)
        math_str = ''.join(u)
        return math_str

    @staticmethod
    def _shape_math_binomial(math_str):
        res = '\\({\\\\Xbinom{(.*?)}{(.*?)}}\\)'
        tmp = ''
        while tmp != math_str:
            tmp = math_str
            math_str = re.sub(res, '\\\\binom{\\1}{\\2}', math_str)
        return math_str

    @staticmethod
    def _shape_sub_and_sup(math_str):
        # {}_{A}{{B}_{C}} -> {}_{A}{B}_{C}
        res = '{}_{([^{}]*(?:{[^{}]*})?[^{}]*)}' \
            + '{{([^{}]*(?:{[^{}]*})?[^{}]*)}' \
            + '_{([^{}]*(?:{[^{}]*})?[^{}]*)}}'
        tmp = ''
        while tmp != math_str:
            tmp = math_str
            math_str = re.sub(res, '{}_{\\1}{\\2}_{\\3}', math_str)
        return math_str

    @staticmethod
    def extract_fr_fd_cls(fr_fd_lst):
        fr_fd_cls, fr_fd_lst = MathDatum.extract_xx_fd_cls(fr_fd_lst)
        for fd in fr_fd_lst[::-1]:
            if False:
                pass
            elif fd == 'w-2':
                fr_fd_cls.font_width = '>>>'
                fr_fd_lst.remove(fd)
            elif fd == 'w-1':
                fr_fd_cls.font_width = '>>'
                fr_fd_lst.remove(fd)
            elif fd == 'w+1':
                fr_fd_cls.font_width = '<<'
                fr_fd_lst.remove(fd)
            elif fd == 'w+2':
                fr_fd_cls.font_width = '<<<'
                fr_fd_lst.remove(fd)
            elif fd == 'd':
                fr_fd_cls.font_width = '->'
                fr_fd_lst.remove(fd)
            elif fd == 'i':
                fr_fd_cls.font_width = '+>'
                fr_fd_lst.remove(fd)
            elif fd == 'f':
                fr_fd_cls.frame = '[|'
                fr_fd_lst.remove(fd)
        return fr_fd_cls, fr_fd_lst

    @staticmethod
    def extract_bk_fd_cls(bk_fd_lst):
        bk_fd_cls, bk_fd_lst = MathDatum.extract_xx_fd_cls(bk_fd_lst)
        for fd in bk_fd_lst[::-1]:
            if False:
                pass
            elif fd == 'w-2':
                bk_fd_cls.font_width = '<<<'
                bk_fd_lst.remove(fd)
            elif fd == 'w-1':
                bk_fd_cls.font_width = '<<'
                bk_fd_lst.remove(fd)
            elif fd == 'w+1':
                bk_fd_cls.font_width = '>>'
                bk_fd_lst.remove(fd)
            elif fd == 'w+2':
                bk_fd_cls.font_width = '>>>'
                bk_fd_lst.remove(fd)
            elif fd == 'd':
                bk_fd_cls.font_width = '<-'
                bk_fd_lst.remove(fd)
            elif fd == 'i':
                bk_fd_cls.font_width = '<+'
                bk_fd_lst.remove(fd)
            elif fd == 'f':
                bk_fd_cls.frame = '|]'
                bk_fd_lst.remove(fd)
        return bk_fd_cls, bk_fd_lst

    @staticmethod
    def extract_xx_fd_cls(xx_fd_lst):
        xx_fd_cls = FontDecorator([])
        for fd in xx_fd_lst[::-1]:
            if False:
                pass
            elif fd == 's-2':
                xx_fd_cls.font_scale = '---'
                xx_fd_lst.remove(fd)
            elif fd == 's-1':
                xx_fd_cls.font_scale = '--'
                xx_fd_lst.remove(fd)
            elif fd == 's+1':
                xx_fd_cls.font_scale = '++'
                xx_fd_lst.remove(fd)
            elif fd == 's+2':
                xx_fd_cls.font_scale = '+++'
                xx_fd_lst.remove(fd)
            elif fd == 'b':
                xx_fd_cls.bold = '**'
                xx_fd_lst.remove(fd)
            elif fd == 's':
                xx_fd_cls.strike = '~~'
                xx_fd_lst.remove(fd)
            elif fd == 'u':
                xx_fd_cls.underline = '__'
                xx_fd_lst.remove(fd)
            elif re.match('^c=', fd):
                c = re.sub('^c=', '', fd)
                if c == 'white':
                    xx_fd_cls.font_color = '^^'
                else:
                    xx_fd_cls.font_color = '^' + c + '^'
                xx_fd_lst.remove(fd)
            elif re.match('^h=', fd):
                c = re.sub('^h=', '', fd)
                xx_fd_cls.font_color = '_' + c + '_'
                xx_fd_lst.remove(fd)
        return xx_fd_cls, xx_fd_lst


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
    def get_value(tag_name, value_name, cur_value, tag):
        if re.match('<' + tag_name + ' .+>', tag):
            res = '^.* ' + value_name + '=[\'"]([^\'"]*)[\'"].*$'
            if re.match(res, tag):
                new_value = re.sub(res, '\\1', tag)
                if type(cur_value) is int:
                    # INT
                    if re.match('^[-\\+]?[0-9]+$', new_value):
                        return int(new_value)
                    if re.match('^true$', new_value, re.IGNORECASE):
                        return 1
                    if re.match('^false$', new_value, re.IGNORECASE):
                        return -1
                    return cur_value  # bad value
                if type(cur_value) is float:
                    # FLOAT
                    if re.match('^' + RES_NUMBER + '$', new_value):
                        return float(new_value)
                    return cur_value  # bad value
                if type(cur_value) is bool:
                    # BOOL
                    if re.match('^true$', new_value, re.IGNORECASE):
                        return True
                    if re.match('^false$', new_value, re.IGNORECASE):
                        return False
                    if new_value == '1':
                        return True
                    if new_value == '-1':
                        return False
                    return cur_value  # bad value
                # STRING
                return new_value
        return cur_value

    @staticmethod
    def count_values(tag_name, value_name, value_dict, tag):
        if re.match('<' + tag_name + ' .+>', tag):
            res = '^.* ' + value_name + '=[\'"]([^\'"]*)[\'"].*$'
            if re.match(res, tag):
                value = re.sub(res, '\\1', tag)
                if value in value_dict:
                    value_dict[value] += 1
                else:
                    value_dict[value] = 1
        return value_dict

    @staticmethod
    def is_this_tag(tag_name, init_value, tag):
        if re.match('<' + tag_name + '( .*)?/?>', tag):
            return True
        else:
            return init_value


class LineTruncation:

    """A class to truncate lines"""

    def __init__(self, md_text):
        self.old_text = md_text
        indent = self._get_indent(md_text)
        parens = self.Paren._get_parens(md_text)
        phrases = self._split_into_phrases(md_text, parens)
        new_text = self._concatenate_phrases(phrases, indent)
        self.new_text = self._indent_text(new_text, indent)

    def get_truncated_md_text(self):
        return self.new_text

    @staticmethod
    def _get_indent(md_text: str) -> int:
        res_chapter = '^(\\$+(?:-\\$+)*\\s+)((?:.|\n)*)$'
        res_section = '^(#+(?:-#+)*\\s+)((?:.|\n)*)$'
        res_list = '^(\\s*(1\\.|-)\\s+)((?:.|\n)*)$'
        res_alignment = '^(:\\s+)((?:.|\n)*)$'
        if re.match(res_chapter, md_text):
            head_string = re.sub(res_chapter, '\\1', md_text)
            head_string = re.sub('-.*$', '-', head_string)
        elif re.match(res_section, md_text):
            head_string = re.sub(res_section, '\\1', md_text)
            head_string = re.sub('-.*$', '-', head_string)
            if LineTruncation._is_sentence(md_text):
                head_string = ''
        elif re.match(res_list, md_text):
            head_string = re.sub(res_list, '\\1', md_text)
        elif re.match(res_alignment, md_text):
            head_string = re.sub(res_alignment, '\\1', md_text)
        else:
            head_string = ''
        indent = len(head_string)
        return indent

    class Paren:

        @staticmethod
        def is_paren(char):
            if LineTruncation.Paren.is_left_paren(char):
                return True
            if LineTruncation.Paren.is_right_paren(char):
                return True
            return False

        @staticmethod
        def is_left_paren(char):
            if len(char) > 1:
                return False
            if LineTruncation.Paren.are_left_parens(char):
                return True
            return False

        @staticmethod
        def is_right_paren(char):
            if len(char) > 1:
                return False
            if LineTruncation.Paren.are_right_parens(char):
                return True
            return False

        @staticmethod
        def are_parens(chars):
            if LineTruncation.Paren.are_left_parens(chars):
                return True
            if LineTruncation.Paren.are_right_parens(chars):
                return True
            return False

        @staticmethod
        def are_left_parens(chars):
            if re.match('^[\\(（「『]+$', chars):
                return True
            return False

        @staticmethod
        def are_right_parens(chars):
            if re.match('^[\\)）」』]+$', chars):
                return True
            return False

        @staticmethod
        def separate_parens(text):
            res = '^((?:.|\n)*?)([\\(（「『]+)$'
            if re.match(res, text):
                t1 = re.sub(res, '\\1', text)
                t2 = re.sub(res, '\\2', text)
                return t1, t2
            res = '^((?:.|\n)*?)([\\)）」』]+)$'
            if re.match(res, text):
                t1 = re.sub(res, '\\1', text)
                t2 = re.sub(res, '\\2', text)
                return t1, t2
            return text, ''

        depth_list = [0, 0, 0]  # = ["()" or "（）", "「」", "『』"]

        def __init__(self, pos, cha, bef, aft):
            # POSITION
            self.position = pos
            # PARTNER
            self.partner = None
            # PAREN CODE
            self.paren_code = 0
            if cha == ')' or cha == '）':
                self.paren_code = -1
            elif cha == '」':
                self.paren_code = -2
            elif cha == '』':
                self.paren_code = -3
            elif cha == '(' or cha == '（':
                self.paren_code = +1
            elif cha == '「':
                self.paren_code = +2
            elif cha == '『':
                self.paren_code = +3
            # DEPTH LIST
            if self.paren_code < 0:
                self.step_depth()
            self.depth_list = [i for i in LineTruncation.Paren.depth_list]
            if self.paren_code > 0:
                self.step_depth()
            # PARENTHESIS CHARACTER
            self.char = cha
            # ANOTHER POSSIBILITY
            self.has_another_possibility = False
            if cha == ')' or cha == '）':
                if re.match('^[0-9０-９a-zａ-ｚA-ZＡ-Ｚ]$', bef):
                    self.has_another_possibility = True

        def step_depth(self):
            pos = abs(self.paren_code) - 1
            if self.paren_code > 0:
                LineTruncation.Paren.depth_list[pos] += 1
            else:
                LineTruncation.Paren.depth_list[pos] -= 1

        def get_individual_depth(self):
            pos = abs(self.paren_code) - 1
            dep = self.depth_list[pos]
            return dep

        def is_inconsistent(self):
            if self.paren_code == -1:  # ')'
                if not self.has_another_possibility:  # not '1)', 'a)', 'A)'...
                    if self.partner is None:
                        return True
            return False

        def may_be_inconsistent(self):
            if self.paren_code == -1:  # ')'
                if self.has_another_possibility:  # '1)', 'a)', 'A)'...
                    if self.partner is not None:
                        return True
            return False

        @staticmethod
        def _get_parens(md_text):
            parens = []
            m = len(md_text) - 1
            for pos, cha in enumerate(md_text):
                bef = md_text[pos - 1] if pos > 0 else ''
                aft = md_text[pos + 1] if pos < m else ''
                if LineTruncation.Paren.is_paren(cha):
                    p = LineTruncation.Paren(pos, cha, bef, aft)
                    p_cod = p.paren_code
                    p_dep = p.get_individual_depth()
                    if p_cod < 0:
                        for q in parens[::-1]:
                            q_cod = q.paren_code
                            q_dep = q.get_individual_depth()
                            if p_cod + q_cod == 0:
                                if p_dep == q_dep:
                                    p.partner = q.position
                                    q.partner = p.position
                                    break
                    if p.is_inconsistent():
                        if len(parens) > 0:
                            for q in parens[::-1]:
                                if q.paren_code != -1:
                                    break
                                if q.may_be_inconsistent():
                                    for r in parens:
                                        if r.partner == q.position:
                                            r.partner = p.position
                                            break
                                    p.partner = q.partner
                                    q.partner = None
                                    break
                    parens.append(p)
            return parens

    @staticmethod
    def __save_one(phrases, res, tmp1):
        m1 = re.sub(res, '\\1', tmp1)
        m2 = re.sub(res, '\\2', tmp1)
        phrases.append(m1)
        return phrases, m2

    @staticmethod
    def __save_two(phrases, res, tmp1):
        m1 = re.sub(res, '\\1', tmp1)
        m2 = re.sub(res, '\\2', tmp1)
        phrases.append(m1)
        phrases.append(m2)
        return phrases, ''

    @staticmethod
    def __must_continue(res1, res2, tmp1, tmp2):
        if re.match(NOT_ESCAPED + res1 + '$', tmp1):
            if re.match('^' + res2 + '(?:.|\n)*$', tmp2):
                return True
        return False

    @classmethod
    def _split_into_phrases(cls, old_text, parens):
        phrases = []
        fds = ''
        tmp1 = ''
        closing_point = -1
        m = len(old_text) - 1
        for i in range(len(old_text)):
            j = i + 1
            c1 = (old_text + '\0')[i]
            c2 = (old_text + '\0')[j]
            tmp1 += c1
            tmp2 = (old_text + '\0')[j:]
            # FONT DECORATORS
            must_continue = False
            if not must_continue:
                res = NOT_ESCAPED + '(`)$'
                if re.match(res, tmp1):
                    # "`"
                    phrases, m2 = cls.__save_one(phrases, res, tmp1)
                    fds += m2
                    tmp1 = ''
                    must_continue = True
            if not must_continue:
                for c in ['\\*', '\\-', '\\+', '>', '<', '~', '/']:
                    res3 = NOT_ESCAPED + '(' + c * 3 + ')$'
                    res2 = NOT_ESCAPED + '(' + c * 2 + ')$'
                    res1 = NOT_ESCAPED + '(' + c * 1 + ')$'
                    if re.match(res3, tmp1):
                        if c != '~' and c != '/':
                            # "***", "---", "+++", ">>>", "<<<"
                            phrases, m2 = cls.__save_one(phrases, res3, tmp1)
                            fds += m2
                            tmp1 = ''
                            must_continue = True
                            break
                    elif re.match(res2, tmp1):
                        if c == '~' or c == '/':
                            # "~~", "//"
                            phrases, m2 = cls.__save_one(phrases, res2, tmp1)
                            fds += m2
                            tmp1 = ''
                            must_continue = True
                            break
                        elif re.match('^' + c, tmp2):
                            must_continue = True
                            break
                        else:
                            # **, --, ++, >>, <<
                            phrases, m2 = cls.__save_one(phrases, res2, tmp1)
                            fds += m2
                            tmp1 = ''
                            must_continue = True
                            break
                    elif re.match(res1, tmp1):
                        if re.match('^' + c, tmp2):
                            must_continue = True
                            break
                        elif c == '\\*':
                            # *
                            phrases, m2 = cls.__save_one(phrases, res1, tmp1)
                            fds += m2
                            tmp1 = ''
                            must_continue = True
                            break
            if not must_continue:
                for ress in [[NOT_ESCAPED + '(@[^@]{1,66}@)$',
                              NOT_ESCAPED + '@[^@]{,66}$',
                              '^([^@]{,66}@)'],
                             [NOT_ESCAPED + '(\\^[0-9A-Za-z]{,11}\\^)$',
                              NOT_ESCAPED + '\\^[0-9A-Za-z]{,11}$',
                              '^[0-9A-Za-z]{,11}\\^'],
                             [NOT_ESCAPED + '(_[0-9A-Za-z]{1,11}_)$',
                              NOT_ESCAPED + '_[0-9A-Za-z]{,11}$',
                              '^[0-9A-Za-z]{,11}_'],
                             [NOT_ESCAPED + '(_[\\$=\\.#\\-~\\+]{,4}_)$',
                              NOT_ESCAPED + '_[\\$=\\.#\\-~\\+]{,4}$',
                              '^[\\$=\\.#\\-~\\+]{,4}_']]:
                    if re.match(ress[0], tmp1):
                        # @.+@, _.+_, ^.+^
                        phrases, m2 = cls.__save_one(phrases, ress[0], tmp1)
                        fds += m2
                        tmp1 = ''
                        must_continue = True
                        break
                    elif re.match(ress[1], tmp1) and re.match(ress[2], tmp2):
                        must_continue = True
                        break
            if (not must_continue) or (i == m):
                if fds != '':
                    phrases.append(fds)
                    fds = ''
            if must_continue:
                continue
            # SPACE
            if re.match('^[ \t\u3000](?:.|\n)$', tmp2):
                continue
            if re.match('^(?:.|\n)*[\t\u3000]$', tmp1):
                continue
            # SUB OR SUP
            if re.match('^[_\\^]{[^{}]*}', tmp2):
                continue
            if cls.__must_continue('[_\\^]', '{', tmp1, tmp2):
                continue
            if cls.__must_continue('[_\\^]{[^{}]*', '[^{}]*}', tmp1, tmp2):
                continue
            # LINE BREAK
            res = '^((?:.|\n)*)(\n)$'
            if re.match(res, tmp1):
                phrases, tmp1 = cls.__save_one(phrases, res, tmp1)
                phrases.append('<br>')
                tmp1 = ''
                continue
            # IMAGE
            res = NOT_ESCAPED + '(' + RES_IMAGE + ')$'
            if re.match(res, tmp1):
                phrases, tmp1 = cls.__save_two(phrases, res, tmp1)
                continue
            if cls.__must_continue('!',
                                   '\\[[^\\[\\]]*\\]' + '\\([^\\(\\)]*\\)',
                                   tmp1, tmp2):
                continue  # ! + [....](....)
            if cls.__must_continue('!' + '\\[[^\\[\\]]*',
                                   '[^\\[\\]]*\\]' + '\\([^\\(\\)]*\\)',
                                   tmp1, tmp2):
                continue  # ![.. + ..](....)
            if cls.__must_continue('!' + '\\[[^\\[\\]]*\\]',
                                   '\\([^\\(\\)]*\\)',
                                   tmp1, tmp2):
                continue  # ![....] + (....)
            if cls.__must_continue('!' + '\\[[^\\[\\]]*\\]\\([^\\(\\)]*',
                                   '[^\\(\\)]*\\)',
                                   tmp1, tmp2):
                continue  # ![....](.. + ..)
            # NUMBER
            if re.match('^.*[0-9０-９]+[,\\.，．]$', tmp1):
                if re.match('^[0-9０-９]+.*$', tmp2):
                    continue
            # MATH
            res = NOT_ESCAPED + '(\\\\\\[)$'
            if re.match(res, tmp1):
                t, tex = old_text[j:], ''
                res_tex = NOT_ESCAPED + '\\\\\\]((?:.|\n)*)'
                if re.match(res_tex, t):
                    tex = re.sub(res_tex, '\\1', t)
                wid = get_ideal_width('\\[' + tex + '\\]')
                if wid <= int(MD_TEXT_WIDTH / 2):
                    phrases, tmp1 = cls.__save_one(phrases, res, tmp1)
                else:
                    phrases, tmp1 = cls.__save_two(phrases, res, tmp1)
                continue
            res = NOT_ESCAPED + '(\\\\\\])$'
            if re.match(res, tmp1):
                t = old_text[:j]
                res_tex = NOT_ESCAPED + '\\\\\\[((?:.|\n)*)'
                while re.match(res_tex, t):
                    t = re.sub(res_tex, '\\2', t)
                wid = get_ideal_width('\\[' + t)
                if wid <= int(MD_TEXT_WIDTH / 2):
                    phrases.append(tmp1)
                    tmp1 = ''
                else:
                    phrases, tmp1 = cls.__save_two(phrases, res, tmp1)
                continue
            if cls.__must_continue('\\\\', '[\\[\\]]', tmp1, tmp2):
                continue
            # TRACK CHANGES
            res = NOT_ESCAPED + '([\\-\\+]>)$'
            if re.match(res, tmp1):
                phrases, tmp1 = cls.__save_two(phrases, res, tmp1)
                continue
            res = NOT_ESCAPED + '(<[\\-\\+])$'
            if re.match(res, tmp1):
                phrases, tmp1 = cls.__save_two(phrases, res, tmp1)
                continue
            if cls.__must_continue('[\\-\\+]', '>', tmp1, tmp2):
                continue
            if cls.__must_continue('<', '[\\-\\+]', tmp1, tmp2):
                continue
            # PARENTHESES
            if cls.Paren.is_paren(c1):
                par = None
                for p in parens:
                    if p.position == i:
                        par = p
                        break
                if (par is not None) and (par.partner is not None):
                    t_not, t_par = cls.Paren.separate_parens(tmp1)
                    if par.paren_code > 0:
                        # OPEN PARENTHESES "(（「『"
                        b = par.position
                        e = par.partner
                        s = old_text[b:e + 1]
                        w = get_ideal_width(s)
                        if closing_point < 0:
                            phrases.append(t_not)
                            if w <= int(MD_TEXT_WIDTH / 2):
                                tmp1 = t_par
                                closing_point = e
                            else:
                                k = len(phrases) - 1
                                while k >= 0:
                                    if phrases[k] != '':
                                        break
                                    k -= 1
                                if k >= 0 and \
                                   cls.Paren.are_left_parens(phrases[k]):
                                    phrases[k] += t_par
                                else:
                                    phrases.append(t_par)
                                tmp1 = ''
                                closing_point = -1
                    else:
                        # CLOSE PARENTHESES "』」）)"
                        if closing_point == i:
                            phrases.append(tmp1)
                            tmp1 = ''
                            closing_point = -1
                        elif (closing_point < 0 and
                              not cls.Paren.is_right_paren(c2)):
                            phrases.append(t_not)
                            phrases.append(t_par)
                            tmp1 = ''
                            closing_point = -1
                continue
            # PUNCTUATION
            res_pun = '[,\\.，、．。]'
            if re.match('^(.|\n)*' + res_pun + '$', tmp1):
                if not re.match('^' + res_pun, tmp2) and \
                   not LineTruncation.Paren.is_right_paren(c2):
                    phrases.append(tmp1)
                    tmp1 = ''
                    continue
            # SPACE
            if re.match('^(.|\n)* $', tmp1) and (not re.match('^ ', tmp1)):
                if re.match('^@[^@]{1,66}$', tmp1):
                    continue  # font scale or name
                phrases.append(tmp1)
                tmp1 = ''
                continue
            # REMOVED 24.09.26 >
            # END
            # if i == m:
            #     if tmp1 != '':
            #         phrases.append(tmp1)
            #         tmp1 = ''
            #     break
            # <
        if tmp1 != '':
            phrases.append(tmp1)
            tmp1 = ''
        # REMOVE EMPTY
        while '' in phrases:
            phrases.remove('')
        return phrases

    @classmethod
    def _concatenate_phrases(cls, phrases: list, indent: int) -> str:
        def __extend_tex(extension):
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
        is_in_math = False  # not in use now
        for p in phrases:
            # INDENT
            if '\n' not in tex:
                # FIRST LINE
                md_text_width = MD_TEXT_WIDTH
            else:
                # SECOND LINE AND ONWARDS
                md_text_width = MD_TEXT_WIDTH - indent
            if md_text_width < 2:
                md_text_width = 2  # width of full width character
            # MATH MODE (MUST BE FIRST)
            if p == '\\[' and not is_in_math:
                tex = __extend_tex(tmp)
                tex = __extend_tex(p)
                tmp = ''
                is_in_math = True
                continue
            if p == '\\]' and is_in_math:
                tex = __extend_tex(tmp)
                tex = __extend_tex(p)
                tmp = ''
                is_in_math = False
                continue
            # DELETED
            if (not is_in_deleted) and p == '->':
                tex = __extend_tex(tmp)
                tmp = ''
                is_in_deleted = True
                continue
            if is_in_deleted and p == '<-':
                tex = __extend_tex(tmp)
                tmp = ''
                is_in_deleted = False
                continue
            # INSERTED
            if (not is_in_inserted) and p == '+>':
                tex = __extend_tex(tmp)
                tmp = ''
                is_in_inserted = True
                continue
            if is_in_inserted and p == '<+':
                tex = __extend_tex(tmp)
                tmp = ''
                is_in_inserted = False
                continue
            # LINE BREAK
            if p == '<br>':
                tex = __extend_tex(tmp)
                tex = __extend_tex(p)
                tmp = ''
                continue
            # NUMBERED
            if re.match('^.*[,，、]$', tmp):
                n1 = '0-9０-９ｱ-ﾝA-ZＡ-Ｚa-zａ-ｚ' \
                    + 'アイウエオカキクケコサシスセソタチツテトナニヌネノ' \
                    + 'ハヒフヘホマミムメモヤユヨラリルレロワヰヱヲン' \
                    + 'あいうえおかきくけこさしすせそたちつてとなにぬねの' \
                    + 'はひふへほまみむめもやゆよらりるれろわゐゑをん'
                if re.match('^[\\(（][' + n1 + ']+[\\)）]', p):
                    tex = __extend_tex(tmp)
                    tmp = p
                    continue
                n2 = '⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇' \
                    + '🄐🄑🄒🄓🄔🄕🄖🄗🄘🄙🄚🄛🄜🄝🄞🄟🄠🄡🄢🄣🄤🄥🄦🄧🄨🄩' \
                    + '⒜⒝⒞⒟⒠⒡⒢⒣⒤⒥⒦⒧⒨⒩⒪⒫⒬⒭⒮⒯⒰⒱⒲⒳⒴⒵' \
                    + '㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩' \
                    + '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳' \
                    + '㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵' \
                    + '㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿' \
                    + '❶❷❸❹❺❻❼❽❾❿⓫⓬⓭⓮⓯⓰⓱⓲⓳⓴' \
                    + '㋐㋑㋒㋓㋔㋕㋖㋗㋘㋙㋚㋛㋜㋝㋞㋟㋠㋡㋢㋣㋤㋥㋦㋧㋨' \
                    + '㋩㋪㋫㋬㋭㋮㋯㋰㋱㋲㋳㋴㋵㋶㋷㋸㋹㋺㋻㋼㋽㋾' \
                    + 'ⒶⒷⒸⒹⒺⒻⒼⒽⒾⒿⓀⓁⓂⓃⓄⓅⓆⓇⓈⓉⓊⓋⓌⓍⓎⓏ' \
                    + 'ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙⓚⓛⓜⓝⓞⓟⓠⓡⓢⓣⓤⓥⓦⓧⓨⓩ' \
                    + '㊀㊁㊂㊃㊄㊅㊆㊇㊈㊉' \
                    + 'ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ' \
                    + 'ⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅺⅻ' \
                    + '⒈⒉⒊⒋⒌⒍⒎⒏⒐⒑⒒⒓⒔⒕⒖⒗⒘⒙⒚⒛'
                if re.match('^[' + n2 + ']', p):
                    tex = __extend_tex(tmp)
                    tmp = p
                    continue
            # FONT DECORATORS
            if re.match('^' + RES_FONT_DECORATORS + '+$', p):
                tex = __extend_tex(tmp)
                tex = __extend_tex(p)
                tmp = ''
                continue
            # PARENTHESES
            if cls.Paren.are_parens(p):
                tex = __extend_tex(tmp)
                tex = __extend_tex(p)
                tmp = ''
                continue
            # SECTION WITHOUT A TITLE
            res = '(?:#+(?:\\-#)* +)'
            if tex == '':
                if re.match('^' + res + '$', tmp):
                    if not re.match('^' + res + '.*$', p):
                        if LineTruncation._is_sentence(phrases[-1]):
                            tex = __extend_tex(tmp + '\\')
                            # tex = __extend_tex(re.sub('\\s+$', '', tmp))
                            tmp = ''
            # IMAGE
            if re.match(RES_IMAGE, p):
                tex = __extend_tex(tmp)
                tex = __extend_tex(p)
                tmp = ''
                continue
            # CONJUNCTIONS
            if re.match('^.*[,，、]$', tmp):
                for c in CONJUNCTIONS:
                    if re.match('^' + c + '[,，、]$', tmp):
                        tex = __extend_tex(tmp)
                        tmp = ''
                        break
            # END OF A SENTENCE
            if re.match('^.*[．。]$', tmp):
                tex = __extend_tex(tmp)
                tmp = ''
            # RIGHT LENGTH
            if tmp != '':
                if get_ideal_width(tmp + p) > md_text_width:
                    tex = __extend_tex(tmp)
                    tmp = ''
            # FONT SCALE (NOT SIZE)
            if re.match('^@.*$', p) and re.match(NOT_ESCAPED + '@$', p):
                if not re.match('^@' + RES_NUMBER + '@$', p):
                    tex = __extend_tex(tmp)
                    tmp = ''
            if re.match('^@.*$', tmp) and re.match(NOT_ESCAPED + '@$', tmp):
                if not re.match('^@' + RES_NUMBER + '@$', tmp):
                    tex = __extend_tex(tmp)
                    tmp = ''
            # CONCATENATE
            tmp += p
            # TOO LONG
            while get_ideal_width(tmp) > md_text_width:
                for i in range(len(tmp), -1, -1):
                    s1 = tmp[:i]
                    s2 = tmp[i:]
                    if get_ideal_width(s1) > md_text_width:
                        continue
                    if re.match('^.*[０-９][，．]$', s1) and \
                       re.match('^[０-９].*$', s2):
                        continue
                    if re.match('^.*を$', s1):
                        if s1 != '':
                            tex = __extend_tex(s1)
                            tmp = s2
                            break
                    if re.match('^.*[ぁ-ん，、．。]$', s1) and \
                       re.match('^[^ぁ-ん，、．。].*$', s2):
                        if s1 != '':
                            tex = __extend_tex(s1)
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
                        # '[|' + '|]' (FRAME)
                        if re.match('^.*\\[$', s1) and re.match('^\\|.*$', s2):
                            continue
                        if re.match('^.*\\|$', s1) and re.match('^\\].*$', s2):
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
                        if get_ideal_width(s1) <= md_text_width:
                            if s1 != '':
                                tex += s1 + '\n'
                                tmp = s2
                                break
                    else:
                        tex += tmp + '\n'
                        tmp = ''
        if tmp != '':
            if is_in_deleted:
                tex += '->' + tmp
            elif is_in_inserted:
                tex += '+>' + tmp
            else:
                tex += tmp + '\n'
            tmp = ''
        tmp = ''
        for t in tex.split('\n'):
            if re.match('^\\s+.*$', t):
                if (tmp != '') or (not re.match('^\\s+(1\\.|-)\\s', t)):
                    t = '\\' + t
            if re.match('^.*\\s+$', t):
                t = t + '\\'
            tmp += t + '\n'
        tex = tmp
        tex = re.sub('\n$', '', tex)
        tex = re.sub('(  |\t|\u3000)(\n)', '\\1\\\\\\2', tex)
        new_text = re.sub('\n+', '\n', tex)
        return new_text

    @staticmethod
    def _indent_text(md_text: str, indent: int) -> str:
        md_text = re.sub('\n', ('\n' + ' ' * indent), md_text)
        return md_text

    @staticmethod
    def _is_sentence(md_text: str) -> bool:
        if re.match('^(.|\n)*[.．。]$', md_text):
            return True
        return False


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
            if rp.paragraph_class == 'empty':
                continue
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
            if re.match('^\\s+', p.text_to_write):
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
                        p.text_to_write = p._get_text_to_write()
                        p.text_to_write_with_reviser \
                            = p._get_text_to_write_with_reviser()
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
            p_next = self.__get_next_paragraph(self.paragraphs, i)
            if p.paragraph_class == 'blank':
                v_line = p.md_text.count('\n') + 1.0
                p.md_text = ''
                p.length_supp['space before'] += v_line
                # RENEW
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines_text = p._get_md_lines_text(p.md_text)
                # p.text_to_write = p._get_text_to_write()
                # p.text_to_write_with_reviser \
                #     = p._get_text_to_write_with_reviser()
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
                # p.text_to_write = p._get_text_to_write()
                p.text_to_write_with_reviser \
                    = p._get_text_to_write_with_reviser()
                p_next.length_revi = p_next._get_length_revi()
                p_next.length_revisers \
                    = p_next._get_length_revisers(p_next.length_revi)
                # p_next.md_lines_text \
                #     = p_next._get_md_lines_text(p_next.md_text)
                # p_next.text_to_write = p_next._get_text_to_write()
                # p_next.text_to_write_with_reviser \
                #     = p_next._get_text_to_write_with_reviser()
        # CANCEL FONT REVISERS
        p_prev = None
        for p_next in self.paragraphs:
            if p_next.has_removed:
                continue
            if p_next.paragraph_class == 'empty':
                continue
            if p_prev is not None:
                for tfr in p_prev.tail_font_revisers:
                    hfr = FontDecorator.get_partner(tfr)
                    if hfr in p_next.head_font_revisers:
                        p_prev.tail_font_revisers.remove(tfr)
                        p_next.head_font_revisers.remove(hfr)
            p_prev = p_next
        # REMAKE TEXT TO WRITE WITH REVISERS
        for p_next in self.paragraphs:
            p.text_to_write_with_reviser \
                = p._get_text_to_write_with_reviser()
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
            p_prev = self.__get_prev_paragraph(self.paragraphs, i)
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
                # p_prev.text_to_write = p_prev._get_text_to_write()
                p_prev.text_to_write_with_reviser \
                    = p_prev._get_text_to_write_with_reviser()
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines_text = p._get_md_lines_text(p.md_text)
                # p.text_to_write = p._get_text_to_write()
                p.text_to_write_with_reviser \
                    = p._get_text_to_write_with_reviser()
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
            p_prev = self.__get_prev_paragraph(self.paragraphs, i)
            p_next = self.__get_next_paragraph(self.paragraphs, i)
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
                    p.length_supp['space before'] \
                        += p.length_clas['space before']
                else:
                    p.length_docx['space before'] \
                        = p_prev.length_docx['space after']
                    p_prev.length_docx['space after'] = 0.0
                if p_next is None or p_next.paragraph_class == 'pagebreak':
                    p.length_supp['space after'] \
                        += p.length_clas['space after']
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
                # p_prev.text_to_write = p_prev._get_text_to_write()
                p_prev.text_to_write_with_reviser \
                    = p_prev._get_text_to_write_with_reviser()
            if True:
                p.length_revi = p._get_length_revi()
                p.length_revisers = p._get_length_revisers(p.length_revi)
                # p.md_lines_text = p._get_md_lines_text(p.md_text)
                # p.text_to_write = p._get_text_to_write()
                p.text_to_write_with_reviser \
                    = p._get_text_to_write_with_reviser()
            if p_next is not None:
                p_next.length_revi = p_next._get_length_revi()
                p_next.length_revisers \
                    = p_next._get_length_revisers(p_next.length_revi)
                # p_next.md_lines_text \
                #     = p_next._get_md_lines_text(p_next.md_text)
                # p_next.text_to_write = p_next._get_text_to_write()
                p_next.text_to_write_with_reviser \
                    = p_next._get_text_to_write_with_reviser()
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
            p_next = self.__get_next_paragraph(self.paragraphs, i)
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
                        while '##=1' in p_next.numbering_revisers:
                            p_next.numbering_revisers.remove('##=1')
                        while '###=1' in p_next.numbering_revisers:
                            p_next.numbering_revisers.remove('###=1')
                p.pre_text_to_write += '\n'
                p.length_supp['space before'] -= 1.0
            p.head_section_depth, p.tail_section_depth \
                = p._get_section_depths(p.raw_text_doi, not p.has_removed)
            p.length_clas = p._get_length_clas()
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            # p.md_lines_text = p._get_md_lines_text(p.md_text)
            # p.text_to_write = p._get_text_to_write()
            p.text_to_write_with_reviser = p._get_text_to_write_with_reviser()
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
        for i, p in enumerate(self.paragraphs):
            if i == 0:
                continue
            p_prev = self.__get_prev_paragraph(self.paragraphs, i)
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
            # p.text_to_write = p._get_text_to_write()
            p.text_to_write_with_reviser = p._get_text_to_write_with_reviser()
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
                    # p.text_to_write = p._get_text_to_write()
                    p.text_to_write_with_reviser \
                        = p._get_text_to_write_with_reviser()
                continue
            rt = p.raw_text
            for fd in FONT_DECORATORS:
                res = NOT_ESCAPED + fd
                while re.match(res, rt):
                    rt = re.sub(res, '\\1', rt)
            while re.match(NOT_ESCAPED + '\\\\', rt):
                rt = re.sub(NOT_ESCAPED + '\\\\', '\\1', rt)
            unit = 12 * 2.54 / 72 / 2
            line_width_in_cm = float(get_real_width(rt)) * unit
            indent = p.length_docx['first indent'] \
                + p.length_docx['left indent'] \
                + p.length_docx['right indent']
            region_width_in_cm = PAPER_WIDTH[paper_size] \
                - left_margin - right_margin \
                - (indent * unit)
            if line_width_in_cm > region_width_in_cm:
                continue
            indent \
                = p.length_revi['first indent'] + p.length_revi['left indent']
            if indent > -.25 and indent < +.25:
                p.length_supp['first indent'] -= p.length_revi['first indent']
                p.length_supp['left indent'] -= p.length_revi['left indent']
            elif re.match('^\\s+', p.text_to_write):
                p.length_supp['first indent'] += p.length_revi['left indent']
                p.length_supp['left indent'] -= p.length_revi['left indent']
            else:
                continue
            # RENEW
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            # p.md_lines_text = p._get_md_lines_text(p.md_text)
            # p.text_to_write = p._get_text_to_write()
            p.text_to_write_with_reviser = p._get_text_to_write_with_reviser()
        return self.paragraphs

    def _modpar_cancel_first_indent(self):
        # |             ->  |
        # |　1行の段落  ->  |<<=-1.0
        # |             ->  |1行の段落
        # |             ->  |
        res = '^([ \t\u3000]+)((?:.|\n)*)$'
        for p in self.paragraphs:
            if p.head_space == '':
                continue
            if p.chars_data[0].fr_fd_cls.font_scale != '' or \
               p.chars_data[0].fr_fd_cls.font_width != '' or \
               p.chars_data[0].bk_fd_cls.font_scale != '' or \
               p.chars_data[0].bk_fd_cls.font_width != '':
                continue
            w = 0
            for c in p.head_space:
                if c == ' ':
                    w += 0.5
                elif c == '\t':
                    w = (int(w / TAB_WIDTH) + 1) * TAB_WIDTH
                elif c == '\u3000':
                    w += 1.0
            p.head_space = ''
            p.length_supp['first indent'] += w
            p.length_revi = p._get_length_revi()
            p.length_revisers = p._get_length_revisers(p.length_revi)
            # p.md_lines_text = p._get_md_lines_text(p.md_text)
            p.text_to_write = p._get_text_to_write()
            p.text_to_write_with_reviser = p._get_text_to_write_with_reviser()
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
            p_prev = self.__get_prev_paragraph(self.paragraphs, i)
            p_next = self.__get_next_paragraph(self.paragraphs, i)
            for lr in p.length_revisers[::-1]:
                # PREV
                if p_prev is not None and re.match('^v=-.*', lr):
                    must_remove = True
                    for plr in p_prev.length_revisers:
                        if re.match('^V=-.*', plr):
                            must_remove = False
                    if must_remove:
                        if lr in p.length_revisers:
                            p.length_revisers.remove(lr)
                # NEXT
                if p_next is not None and re.match('^V=-.*', lr):
                    must_remove = True
                    for nlr in p_next.length_revisers:
                        if re.match('^v=-.*', nlr):
                            must_remove = False
                    if must_remove:
                        if lr in p.length_revisers:
                            p.length_revisers.remove(lr)
            # RENEW
            p.text_to_write_with_reviser \
                = p._get_text_to_write_with_reviser()
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
        for i, p in enumerate(self.paragraphs):
            # CURRENT
            p_curr = p
            if i == 0:
                curr_hfr, curr_tfr = [], []
                for hfr in p_curr.head_font_revisers:
                    curr_hfr.append(hfr)
                for tfr in p_curr.tail_font_revisers:
                    curr_tfr.append(tfr)
            else:
                curr_hfr, curr_tfr = next_hfr, next_tfr
            # NEXT
            p_next = self.__get_next_paragraph(self.paragraphs, i)
            next_hfr, next_tfr = [], []
            if p_next is not None:
                for hfr in p_next.head_font_revisers:
                    next_hfr.append(hfr)
                for tfr in p_next.tail_font_revisers:
                    next_tfr.append(tfr)
            # CANCEL
            if p_next is not None:
                for tfr in curr_tfr:
                    hfr = FontDecorator.get_partner(tfr)
                    if tfr in p_curr.tail_font_revisers and \
                       hfr in p_next.head_font_revisers:
                        efr = FontDecorator.escape_chars(hfr)
                        curr_ttw = p_curr.text_to_write
                        if ((hfr in curr_hfr) and (tfr in curr_tfr)) or \
                           (not re.match(NOT_ESCAPED + efr, curr_ttw)):
                            next_ttw = p_next.text_to_write
                            efr = FontDecorator.escape_chars(tfr)
                            if ((hfr in next_hfr) and (tfr in next_tfr)) or \
                               (not re.match(NOT_ESCAPED + efr, next_ttw)):
                                p_curr.tail_font_revisers.remove(tfr)
                                p_next.head_font_revisers.remove(hfr)
            # ISOLATE HEAD FONT REVISERS
            pttw = ''
            for hfr in curr_hfr:
                if hfr in p.head_font_revisers:
                    tfr = FontDecorator.get_partner(hfr)
                    efr = FontDecorator.escape_chars(tfr)
                    ttw = self.__erase_font_decorator(tfr, p.text_to_write)
                    if (tfr in curr_tfr) or \
                       (not re.match(NOT_ESCAPED + efr, ttw)):
                        if tfr not in p.tail_font_revisers:
                            pttw += hfr
                            p.head_font_revisers.remove(hfr)
            if pttw != '':
                p.pre_text_to_write \
                    = re.sub('\\s*\n$', ' ', p.pre_text_to_write)
                p.pre_text_to_write += pttw + '\n'
            # ISOLATE TAIL FONT REVISERS
            pttw = ''
            for tfr in curr_tfr:
                if tfr in p.tail_font_revisers:
                    hfr = FontDecorator.get_partner(tfr)
                    efr = FontDecorator.escape_chars(hfr)
                    ttw = self.__erase_font_decorator(hfr, p.text_to_write)
                    if (hfr in curr_hfr) or \
                       (not re.match(NOT_ESCAPED + efr, ttw)):
                        if hfr not in p.head_font_revisers:
                            pttw += tfr
                            p.tail_font_revisers.remove(tfr)
            if pttw != '':
                p.post_text_to_write \
                    = re.sub('^\n', ' ', p.post_text_to_write)
                p.post_text_to_write += '\n' + pttw
            # RENEW
            if True:
                p_curr.text_to_write_with_reviser \
                    = p_curr._get_text_to_write_with_reviser()
            if p_next is not None:
                p_next.text_to_write_with_reviser \
                    = p_next._get_text_to_write_with_reviser()
        return self.paragraphs

    @staticmethod
    def __erase_font_decorator(fd: str, text: str) -> str:
        if fd == '*':
            while re.match(NOT_ESCAPED + '\\*\\*', text):
                text = re.sub(NOT_ESCAPED + '\\*\\*', '\\1X', text)
        elif fd == '--':
            while re.match(NOT_ESCAPED + '<\\-', text):
                text = re.sub(NOT_ESCAPED + '<\\-', '\\1X', text)
            while re.match(NOT_ESCAPED + '\\-\\-\\-', text):
                text = re.sub(NOT_ESCAPED + '\\-\\-\\-', '\\1X', text)
        elif fd == '++':
            while re.match(NOT_ESCAPED + '<\\+', text):
                text = re.sub(NOT_ESCAPED + '<\\+', '\\1X', text)
            while re.match(NOT_ESCAPED + '\\+\\+\\+', text):
                text = re.sub(NOT_ESCAPED + '\\+\\+\\+', '\\1X', text)
        elif fd == '>>':
            while re.match(NOT_ESCAPED + '>>>', text):
                text = re.sub(NOT_ESCAPED + '>>>', '\\1X', text)
        elif fd == '<<':
            while re.match(NOT_ESCAPED + '<<<', text):
                text = re.sub(NOT_ESCAPED + '<<<', '\\1X', text)
        elif fd == '->':
            while re.match(NOT_ESCAPED + '\\-\\-\\-', text):
                text = re.sub(NOT_ESCAPED + '\\-\\-\\-', '\\1X', text)
            while re.match(NOT_ESCAPED + '\\-\\-', text):
                text = re.sub(NOT_ESCAPED + '\\-\\-', '\\1X', text)
        elif fd == '<-':
            while re.match(NOT_ESCAPED + '<<<', text):
                text = re.sub(NOT_ESCAPED + '<<<', '\\1X', text)
            while re.match(NOT_ESCAPED + '<<', text):
                text = re.sub(NOT_ESCAPED + '<<', '\\1X', text)
        elif fd == '+>':
            while re.match(NOT_ESCAPED + '\\+\\+\\+', text):
                text = re.sub(NOT_ESCAPED + '\\+\\+\\+', '\\1X', text)
            while re.match(NOT_ESCAPED + '\\+\\+', text):
                text = re.sub(NOT_ESCAPED + '\\+\\+', '\\1X', text)
        elif fd == '<+':
            while re.match(NOT_ESCAPED + '<<<', text):
                text = re.sub(NOT_ESCAPED + '<<<', '\\1X', text)
            while re.match(NOT_ESCAPED + '<<', text):
                text = re.sub(NOT_ESCAPED + '<<', '\\1X', text)
        return text

    @staticmethod
    def __get_prev_paragraph(paras, base):
        for i in range(base - 1, -1, -1):
            if paras[i].has_removed:
                continue
            if paras[i].paragraph_class == 'empty':
                continue
            return paras[i]
        return None

    @staticmethod
    def __get_next_paragraph(paras, base):
        for i in range(base + 1, len(paras)):
            if paras[i].has_removed:
                continue
            if paras[i].paragraph_class == 'empty':
                continue
            return paras[i]
        return None

    def get_document(self):
        mcols = []
        for p in self.paragraphs:
            if p.paragraph_class == 'multicolumns':
                mcols.append(p.md_text)
        dcmt = ''
        posi = 0
        if len(mcols) > 0:
            if self.paragraphs[0].paragraph_class != 'multicolumns':
                # IF MULTICOLUMNS AT THE BEGINNIG OF THE DOCUMENT
                dcmt += mcols[posi] + '\n\n'
            posi += 1
        for p in self.paragraphs:
            if p.paragraph_class == 'multicolumns':
                if posi < len(mcols):
                    dcmt += mcols[posi] + '\n\n'
                    posi += 1
                continue
            dcmt += p.get_document()  # main process
            if p.paragraph_class != 'empty' and p.paragraph_class != 'remarks':
                dcmt += '\n'
        res = '^((?:.|\n)*?\n\n\\|(?:-+\\|)+\n\n)((?:\\|(?:-+\\|)+\n\n)*)$'
        if re.match(res, dcmt):
            dcmt = re.sub(res, '\\1', dcmt)
        dcmt = re.sub('\n+$', '\n', dcmt)
        # SINGLE COLUMN DOCUMENT
        if len(mcols) == 1:
            if self.paragraphs[-1].paragraph_class == 'multicolumns':
                if re.match('^\\|-\\|\n\n', dcmt):
                    dcmt = re.sub('^\\|-\\|\n\n', '', dcmt)
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
        self.style_id = None
        self.name = None
        self.font = None
        self.font_size = None
        # self.font_width = None
        self.is_italic = False
        self.is_bold = False
        self.has_strike = False
        self.has_frame = False
        self.underline = None
        self.font_color = None
        # self.highlight_color = None
        self.alignment = None
        self.raw_length = {'sb': 0.0, 'sa': 0.0, 'sl': 0.0, 'if': 0.0,
                           'ih': 0.0, 'il': 0.0, 'ir': 0.0, 'tw': 0.0}
        self._substitute_values()

    def _substitute_values(self):
        type = None
        stid = None
        name = None
        afnt = None
        jfnt = None
        f_2s = None
        f_it = False
        f_bd = False
        f_sk = False
        f_fr = False
        f_ul = None
        f_cl = None
        alig = None
        rl = {'sb': 0.0, 'sa': 0.0, 'sl': 0.0, 'if': 0.0,
              'ih': 0.0, 'il': 0.0, 'ir': 0.0, 'tw': 0.0}
        for xl in self.xml_lines:
            type = XML.get_value('w:style', 'w:type', type, xl)
            stid = XML.get_value('w:style', 'w:styleId', stid, xl)
            name = XML.get_value('w:name', 'w:val', name, xl)
            afnt = XML.get_value('w:rFonts', 'w:ascii', afnt, xl)
            jfnt = XML.get_value('w:rFonts', 'w:eastAsia', jfnt, xl)
            # font = XML.get_value('w:rFonts', '*', font, xl)
            f_2s = XML.get_value('w:sz', 'w:val', f_2s, xl)
            # f_2s = XML.get_value('w:szCs', 'w:val', f_2s, xl)
            f_it = XML.is_this_tag('w:i', f_it, xl)
            f_bd = XML.is_this_tag('w:b', f_bd, xl)
            f_sk = XML.is_this_tag('w:strike', f_sk, xl)
            f_fr = XML.is_this_tag('w:bdr', f_fr, xl)
            f_ul = XML.get_value('w:u', 'w:val', f_ul, xl)
            f_cl = XML.get_value('w:color', 'w:val', f_cl, xl)
            alig = XML.get_value('w:jc', 'w:val', alig, xl)
            rl['sb'] = XML.get_value('w:spacing', 'w:before', rl['sb'], xl)
            rl['sa'] = XML.get_value('w:spacing', 'w:after', rl['sa'], xl)
            rl['sl'] = XML.get_value('w:spacing', 'w:line', rl['sl'], xl)
            rl['if'] = XML.get_value('w:ind', 'w:firstLine', rl['if'], xl)
            rl['ih'] = XML.get_value('w:ind', 'w:hanging', rl['ih'], xl)
            rl['il'] = XML.get_value('w:ind', 'w:left', rl['il'], xl)
            rl['ir'] = XML.get_value('w:ind', 'w:right', rl['ir'], xl)
            rl['tw'] = XML.get_value('w:tblInd', 'w:w', rl['tw'], xl)
        self.type = type
        self.style_id = stid
        self.name = name
        self.font = FontDecorator.get_font_name(afnt, jfnt)
        if f_2s is not None:
            self.font_size = round(float(f_2s) / 2, 1)
        self.is_italic = f_it
        self.is_bold = f_bd
        self.has_strike = f_sk
        self.has_frame = f_fr
        self.underline = f_ul
        self.font_color = f_cl
        self.alignment = alig
        self.raw_length = rl


class RawParagraph:

    """A class to handle raw paragraph"""

    raw_paragraph_number = 0

    def __init__(self, xml_lines):
        # DECLARATION
        self.raw_paragraph_number = -1
        self.has_removed = False
        self.xml_lines = []
        self.raw_class = ''
        self.style = None
        self.alignment = ''
        self.horizontal_line = ''  # 'top'|'bottom'|'textbox'
        self.attached_pagebreak = ''  # 'pgbr' | 'Pgbr'
        self.chars_data = []
        self.images = {}
        self.footnotes = {}
        self.raw_text = ''
        self.head_space = ''
        self.tail_space = ''
        self.raw_text_del = ''
        self.raw_text_ins = ''
        self.raw_text_doi = ''
        self.remarks = []
        self.paragraph_class = ''
        # SUBSTITUTION
        RawParagraph.raw_paragraph_number += 1
        self.raw_paragraph_number = RawParagraph.raw_paragraph_number
        self.xml_lines = xml_lines
        self.raw_class = self._get_raw_class(xml_lines)
        self.style, self.alignment, self.chars_data, self.raw_text, \
            self.images, self.footnotes = self.get_raw_text_and_etc(xml_lines)
        self.horizontal_line \
            = self._get_horizontal_line(self.raw_class, xml_lines)
        self.attached_pagebreak = self._get_attached_pagebreak(xml_lines)
        self.head_space, self.raw_text \
            = self._separate_head_space(self.raw_text,
                                        '->', '<-', '\\+>', '<\\+')
        rts, rrt \
            = self._separate_head_space(self.raw_text[::-1],
                                        '-<', '>-', '>\\+', '\\+<')
        self.raw_text = rrt[::-1]
        self.tail_space = rts[::-1]
        self.raw_text_del = self._get_raw_text_del(self.raw_text)
        self.raw_text_ins = self._get_raw_text_ins(self.raw_text)
        if Paragraph.get_font_revisers_and_md_text(self.raw_text_ins)[2] != '':
            self.raw_text_doi = self.raw_text_ins
        else:
            self.raw_text_doi = self.raw_text_del
        if self.raw_text_del != '' and self.raw_text_ins == '':
            self.has_removed = True
        self.remarks = self._get_remarks(xml_lines)
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
    def _get_attached_pagebreak(xml_lines):
        for xl in xml_lines:
            if re.match('^<w:br w:type=[\'"]page[\'"]/>$', xl):
                return 'pgbr'
        return ''

    @staticmethod
    def get_raw_text_and_etc(xml_lines, type='normal'):
        style = RawParagraph._get_style(xml_lines)
        alignment = RawParagraph.get_alignment(style, xml_lines)
        chars_data, images, footnotes \
            = RawParagraph._get_chars_data_and_etc(xml_lines, style, type)
        chars_data = RawParagraph._reduce_font_name(chars_data)
        # chars_data.reverse()
        # chars_data = RawParagraph._reduce_font_name(chars_data)
        # chars_data.reverse()
        raw_text = RawParagraph.get_raw_text(chars_data)
        return style, alignment, chars_data, raw_text, images, footnotes

    @staticmethod
    def _get_style(xml_lines):
        style_id = None
        for xl in xml_lines:
            style_id = XML.get_value('w:pStyle', 'w:val', style_id, xl)
        style = None
        if style_id is not None:
            for fs in Form.styles:
                if fs.type == 'paragraph' and style_id == fs.style_id:
                    style = fs
                    break
        # self.style = style
        return style

    @staticmethod
    def get_alignment(style, xml_lines):
        alignment = ''
        if style is not None and style.alignment is not None:
            alignment = style.alignment
        for xl in xml_lines:
            alignment = XML.get_value('w:jc', 'w:val', alignment, xl)
            if not re.match('^(left|center|right)$', alignment):
                alignment = ''
        # self.alignment = alignment
        return alignment

    @classmethod
    def _get_chars_data_and_etc(cls, xml_lines, style, type):
        # MARKUP COMPATIBILITY
        m = len(xml_lines) - 1
        must_drop = False
        for i in range(m - 1, -1, -1):
            cur, nex = xml_lines[i], xml_lines[i + 1]
            if re.match('^</mc:Choice( .*[^/])?>$', cur):
                if re.match('^<mc:Fallback( .*[^/])?>$', nex):
                    must_drop = True
            if must_drop:
                xml_lines.pop(i)
            if re.match('^<mc:Choice( .*[^/])?>$', cur):
                must_drop = False
        font_size = Form.font_size
        chars_data = []
        images = {}
        footnotes = {}
        img_rels = Form.rels
        img_file_name = ''
        img_size = ''
        is_changed = False
        fldchar = ''
        track_changes = ''  # ''|'del'|'ins'
        ruby = ''  # ''|'rub'|'bas'
        width = 100
        cd = CharsDatum([], '', [])
        cd.apply_style(style)
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
                cd.reset_fds()
                fldchar = 'separate'
            elif re.match('^<w:fldChar w:fldCharType="end"/?>$', xl):
                fldchar = 'end'
            if fldchar == 'separate':
                continue
            # MATH
            if 'math_data' not in locals():
                math_data = None
            math_data, math_chars_datum \
                = MathDatum.get_math_data(xl, math_data)
            if math_chars_datum is not None:
                chars_data.append(math_chars_datum)
                math_chars_datum = None
            if math_data is not None:
                continue
            # IMAGE
            must_continue = False
            if re.match(RES_XML_IMG_MS, xl):
                # IMAGE MS WORD
                img_rel_name, img_file_name \
                    = cls.__get_img_file_names_ms(xl, img_rels)
                Document.images[img_rel_name] = img_file_name
                images[img_rel_name] = img_file_name
                must_continue = True
            elif re.match(RES_XML_IMG_PY_ID, xl):
                # IMAGE PYTHON-DOCX ID
                img_rel_name, img_file_name \
                    = cls.__get_img_file_names_py(xl, img_rels, img_py_name)
                Document.images[img_rel_name] = img_file_name
                images[img_rel_name] = img_file_name
                must_continue = True
            elif re.match(RES_XML_IMG_PY_NAME, xl):
                # IMAGE PYTHON-DOCX NAME
                img_py_name = re.sub(RES_XML_IMG_PY_NAME, '\\2', xl)
                must_continue = True
            elif re.match(RES_XML_IMG_SIZE, xl):
                # IMAGE SIZE
                img_size = cls.__get_img_size(xl)
                must_continue = True
            if img_file_name != '' and img_size != '':
                fr, imt, bk \
                    = cls.__get_img_md_text(img_file_name, img_size, font_size)
                cd_img = CharsDatum([fr], '', [bk])
                if track_changes == 'del':
                    cd_img.fr_fd_cls.del_or_ins = '->'
                    cd_img.bk_fd_cls.del_or_ins = '<-'
                elif track_changes == 'ins':
                    cd_img.fr_fd_cls.del_or_ins = '+>'
                    cd_img.bk_fd_cls.del_or_ins = '<+'
                if re.match('^---(.*)---$', imt):
                    imt = re.sub('^---(.*)---$', '\\1', imt)
                    cd_img.fr_fd_cls.scale = '---'
                    cd_img.bk_fd_cls.scale = '---'
                elif re.match('^--(.*)--$', imt):
                    imt = re.sub('^--(.*)--$', '\\1', imt)
                    cd_img.fr_fd_cls.scale = '--'
                    cd_img.bk_fd_cls.scale = '--'
                elif re.match('^\\+\\+\\+(.*)\\+\\+\\+$', imt):
                    imt = re.sub('^\\+\\+\\+(.*)\\+\\+\\+$', '\\1', imt)
                    cd_img.fr_fd_cls.scale = '+++'
                    cd_img.bk_fd_cls.scale = '+++'
                elif re.match('^\\+\\+(.*)\\+\\+$', imt):
                    imt = re.sub('^\\+\\+(.*)\\+\\+$', '\\1', imt)
                    cd_img.fr_fd_cls.scale = '++'
                    cd_img.bk_fd_cls.scale = '++'
                cd_img.chars = '<>' + imt  # '<>' is to avoid being escaped
                chars_data.append(cd_img)
                img_file_name = ''
                img_size = ''
            if must_continue:
                continue
            # TEXTBOX
            if re.match('^<v:textbox( .*[^/])?>$', xl):
                chars_data.append(CharsDatum([], '$[', []))
                continue
            elif re.match('^</v:textbox( .*[^/])?>$', xl):
                chars_data.append(CharsDatum([], ']$', []))
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
            # RESET
            if xl == '<w:rPr>':
                cd.reset_fds()
                cd.apply_style(style)
            # FONT
            if re.match('^<w:rFonts .*>$', xl):
                afnt = XML.get_value('w:rFonts', 'w:ascii', '', xl)
                if re.match('^.* w:eastAsia=[\'"]([^\'"]*)[\'"].*$', xl):
                    jfnt = XML.get_value('w:rFonts', 'w:eastAsia', '', xl)
                else:
                    # (FOR COMPLEX SCRIPT)
                    jfnt = XML.get_value('w:rFonts', 'w:cs', '', xl)
                font = FontDecorator.get_font_name(afnt, jfnt)
                fd = FontDecorator.get_font_name_fd(font)
                if fd is not None:
                    cd.fr_fd_cls.font_name = fd
                    cd.bk_fd_cls.font_name = fd
                continue
            # FONT SIZE AND SCALE
            v = XML.get_value('w:sz', 'w:val', -1.0, xl)
            v = XML.get_value('w:szCs', 'w:val', v, xl)  # (for complex script)
            if v > 0:
                s = round(v / 2, 1)
                fd = FontDecorator.get_font_scale_fd(s)
                if fd is not None:
                    cd.fr_fd_cls.font_scale = fd
                    cd.bk_fd_cls.font_scale = fd
                continue
            # FONT WIDTH
            w = XML.get_value('w:w', 'w:val', -1.0, xl)
            if w > 0:
                fd1, fd2 = FontDecorator.get_font_width_fd(w)
                if fd1 is not None and fd2 is not None:
                    cd.fr_fd_cls.font_width = fd1
                    cd.bk_fd_cls.font_width = fd2
                width = w
                continue
            # ITALIC
            if re.match('^<w:i/?>$', xl):
                cd.fr_fd_cls.italic = '*'
                cd.bk_fd_cls.italic = '*'
                continue
            # BOLD
            if re.match('^<w:b/?>$', xl):
                cd.fr_fd_cls.bold = '**'
                cd.bk_fd_cls.bold = '**'
                continue
            # STRIKETHROUGH
            if re.match('^<w:strike/?>$', xl):
                cd.fr_fd_cls.strike = '~~'
                cd.bk_fd_cls.strike = '~~'
                continue
            # STRIKETHROUGH
            if re.match('^<w:bdr( .*)?/?>$', xl):
                cd.fr_fd_cls.strike = '[|'
                cd.bk_fd_cls.strike = '|]'
                continue
            # UNDERLINE
            if re.match('^<w:u( .*)?>$', xl):
                val = ''
                res = '^<.* w:val=[\'"]([a-zA-Z]+)[\'"].*>$'
                if re.match(res, xl):
                    val = re.sub(res, '\\1', xl)
                fd = FontDecorator.get_underline_fd(val)
                if fd is not None:
                    cd.fr_fd_cls.underline = fd
                    cd.bk_fd_cls.underline = fd
                continue
            # FONT COLOR
            if re.match('^<w:color w:val="[0-9A-F]+"( .*)?/?>$', xl):
                val = re.sub('^<.* w:val="([0-9A-F]+)".*>$', '\\1', xl, re.I)
                fd = FontDecorator.get_font_color_fd(val)
                cd.fr_fd_cls.font_color = fd
                cd.bk_fd_cls.font_color = fd
                continue
            # HIGHLIGHT COLOR
            if re.match('^<w:highlight w:val="[a-zA-Z]+"( .*)?/?>$', xl):
                val = re.sub('^<.* w:val="([a-zA-Z]+)".*>$', '\\1', xl)
                fd = FontDecorator.get_highlight_color_fd(val)
                cd.fr_fd_cls.highlight_color = fd
                cd.bk_fd_cls.highlight_color = fd
                continue
            # SUBSCRIPT OR SUPERSCRIPT
            if xl == '<w:vertAlign w:val="subscript"/>':
                cd.fr_fd_cls.sub_or_sup = '_{'
                cd.bk_fd_cls.sub_or_sup = '_}'
                continue
            elif xl == '<w:vertAlign w:val="superscript"/>':
                cd.fr_fd_cls.sub_or_sup = '^{'
                cd.bk_fd_cls.sub_or_sup = '^}'
                continue
            # AUTO NUMBERING STYLE
            res_number_ms = '^<w:numId w:val=[\'"]([0-9]+)[\'"]/>$'
            res_number_lo = '^<w:pStyle w:val=[\'"]ListNumber([0-9]?)[\'"]/>$'
            res_ilvl = '<w:ilvl w:val="([0-9]+)"/>'
            if xl == '<w:numPr>':
                numid, ilvl = -1, -1
                continue
            elif re.match(res_number_ms, xl):
                numid = re.sub(res_number_ms, '\\1', xl)
                continue
            elif re.match(res_number_lo, xl):
                numid = re.sub(res_number_lo, '\\1', xl)
                continue
            elif re.match(res_ilvl, xl):
                ilvl = re.sub(res_ilvl, '\\1', xl)
                continue
            elif xl == '</w:numPr>':
                ans_key = str(numid) + '-' + str(ilvl)
                if ans_key in Form.auto_numbering_styles:
                    ans = Form.auto_numbering_styles[ans_key]
                    n = ans.start + ans.state
                    if re.match('^decimal(?:FullWidth)?$', ans.number_format):
                        hs = re.sub('%[1-9]', str(n), ans.head_string)
                        cd.chars += hs + ' '
                    elif re.match('^decimalEnclosedParen$', ans.number_format):
                        hs = re.sub('%[1-9]', n2c_p_arab(n), ans.head_string)
                        cd.chars += hs + ' '
                    elif re.match('^aiueo(?:FullWidth)?$', ans.number_format):
                        hs = re.sub('%[1-9]', n2c_n_kata(n), ans.head_string)
                        cd.chars += hs + ' '
                    elif re.match('lowerLetter', ans.number_format):
                        hs = re.sub('%[1-9]', n2c_n_alph(n), ans.head_string)
                        cd.chars += hs + ' '
                    ans.state += 1
                continue
            # FOOTNOTE
            if re.match('^<w:footnoteReference( .*)>$', xl):
                _fnid = XML.get_value('w:footnoteReference', 'w:id', '', xl)
                cd.chars += '[^' + _fnid + ']'
                footnotes[_fnid] = Form.footnotes[_fnid]
            # TEXT
            if not re.match('^<.*>$', xl):
                imm = CharsDatum.prepare_imm(fldchar, xl, type)
                cd.chars = CharsDatum.concatenate_imm(cd.chars, imm)
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
                        cd.fr_fd_cls.track_changes = '->'
                        cd.bk_fd_cls.track_changes = '<-'
                    elif track_changes == 'ins':
                        cd.fr_fd_cls.track_changes = '+>'
                        cd.bk_fd_cls.track_changes = '<+'
                    # RUBY
                    if ruby == 'rub':
                        cd.fr_fd_cls.font_scale = ''
                        cd.bk_fd_cls.font_scale = ''
                    # SPACE
                    if re.match('^\u3000+$', cd.chars) and width != 100:
                        n = len(cd.chars)
                        cd.fr_fd_cls.font_width = ''
                        cd.bk_fd_cls.font_width = ''
                        w = float(width * n) / 100
                        if w.is_integer():
                            w = int(w)
                        cd.chars = '<' + str(w) + '>'
                    chars_data.append(cd)
                width = 100
                cd = CharsDatum([], '', [])
                cd.apply_style(style)
                continue
        # RUBY (PUT OUT OR CANCEL FONT DECORATORS)
        for i in range(len(chars_data)):
            if i < 4:
                continue
            if chars_data[i - 4].chars != '^<':
                continue
            if chars_data[i - 2].chars != '>/<':
                continue
            if chars_data[i - 0].chars != '>$':
                continue
            beg_cd = chars_data[i - 4]
            rub_cd = chars_data[i - 3]
            bas_cd = chars_data[i - 1]
            end_cd = chars_data[i - 0]
            beg_cd.fr_fd_cls = bas_cd.fr_fd_cls
            end_cd.bk_fd_cls = bas_cd.bk_fd_cls
            bas_cd.fr_fd_cls = FontDecorator([])
            bas_cd.bk_fd_cls = FontDecorator([])
            rub_cd.fr_fd_cls = FontDecorator([])
            rub_cd.bk_fd_cls = FontDecorator([])
        # FORCE TO BE FULL_WIDTH
        for i in range(len(chars_data)):
            cd = chars_data[i]
            if re.match('^' + RES_FORCED_TO_BE_FULL_WIDTH + '+$', cd.chars):
                pre_f_deco = ''
                if i > 0:
                    pre_f_deco = chars_data[i - 1].bk_fd_cls.font_name
                pre_f_mint = Form.mincho_font
                if pre_f_deco != '':
                    pre_f_mint = re.sub('^@(.*)@$', '\\1', pre_f_deco)
                pre_f_full = re.sub('^.*/\\s+', '', pre_f_mint)
                if cd.fr_fd_cls.font_name == '@' + pre_f_full + '@' and \
                   cd.bk_fd_cls.font_name == '@' + pre_f_full + '@':
                    cd.fr_fd_cls.font_name = pre_f_deco
                    cd.bk_fd_cls.font_name = pre_f_deco
        # self.chars_data = chars_data
        # self.images = images
        return chars_data, images, footnotes

    @staticmethod
    def __get_img_file_names_ms(xl, img_rels):
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
    def __get_img_file_names_py(xl, img_rels, img_py_name):
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
    def __get_img_size(xl):
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
    def __get_img_md_text(img_file_name, img_size, font_size):
        relative_dir = os.path.basename(IO.media_dir)
        m_size_cm = font_size * 2.54 / 72
        xs_size_cm = m_size_cm * 0.6
        s_size_cm = m_size_cm * 0.8
        l_size_cm = m_size_cm * 1.2
        xl_size_cm = m_size_cm * 1.4
        # cm_w = float(re.sub('x.*$', '', img_size))
        cm_h = float(re.sub('^.*x', '', img_size))
        fr, bk = '', ''
        img_md_text = '![' + img_file_name + ']' \
            + '(' + relative_dir + '/' + img_file_name + ')'
        if cm_h >= m_size_cm * 0.98 and cm_h <= m_size_cm * 1.02:
            # MEDIUM
            pass
        elif cm_h >= xs_size_cm * 0.98 and cm_h <= xs_size_cm * 1.02:
            # XSMALL
            fr, bk = '---', '---'
        elif cm_h >= s_size_cm * 0.98 and cm_h <= s_size_cm * 1.02:
            # SMALL
            fr, bk = '--', '--'
        elif cm_h >= l_size_cm * 0.98 and cm_h <= l_size_cm * 1.02:
            # LARGE
            fr, bk = '++', '++'
        elif cm_h >= xl_size_cm * 0.98 and cm_h <= xl_size_cm * 1.02:
            # XLARGE
            fr, bk = '+++', '+++'
        else:
            # FREE SIZE
            img_md_text = '![' + img_file_name + ' @' + img_size + ']' \
                + '(' + relative_dir + '/' + img_file_name + ')'
        return fr, img_md_text, bk

    @classmethod
    def _reduce_font_name(cls, chars_data):
        # FORM
        frm_font = Form.mincho_font
        frm_afont, frm_jfont = cls._get_ascii_and_kanji_font(frm_font)
        for i, cur_cd in enumerate(chars_data):
            # PREVIOUS
            pre_font = ''
            if i > 0:
                pre_cd = chars_data[i - 1]
                pre_state = cls.__get_chars_state(pre_cd.chars)
                pre_font = pre_cd.fr_fd_cls.font_name
                pre_afont, pre_jfont = cls._get_ascii_and_kanji_font(pre_font)
                pre_fd = '@' + pre_font + '@'
            # CURRENT
            if True:
                cur_state = cls.__get_chars_state(cur_cd.chars)
                cur_font = cur_cd.fr_fd_cls.font_name
                cur_afont, cur_jfont = cls._get_ascii_and_kanji_font(cur_font)
                cur_fd = '@' + cur_font + '@'
            # REDUCE
            if cur_font != '':
                if cur_cd.bk_fd_cls.font_name != cur_fd:
                    continue
            if (cur_state == 'only ascii' and cur_afont == frm_afont) or \
               (cur_state == 'only kanji' and cur_jfont == frm_jfont) or \
               (cur_state == 'mix' and cur_font == frm_font):
                cur_cd.fr_fd_cls.font_name = ''
                cur_font, cur_afont, cur_jfont = '', '', ''
            # REDUCE
            if pre_font != '':
                if pre_cd.bk_fd_cls.font_name != pre_fd:
                    continue
            if pre_font == '':
                continue
            if pre_font == cur_font:
                continue
            for j in range(i + 1, len(chars_data)):
                if chars_data[j].fr_fd_cls.font_name == pre_fd:
                    if chars_data[j].bk_fd_cls.font_name == pre_fd:
                        break  # There is same font backward
            else:
                continue  # There is no same font backward
            if cur_font != '':
                tmp_font, tmp_afont, tmp_jfont = cur_font, cur_afont, cur_jfont
            else:
                tmp_font, tmp_afont, tmp_jfont = frm_font, frm_afont, frm_jfont
            if (cur_state == 'only ascii' and tmp_afont == pre_afont) or \
               (cur_state == 'only kanji' and tmp_jfont == pre_jfont) or \
               (cur_state == 'mix' and tmp_font == pre_font):
                cur_cd.fr_fd_cls.font_name = pre_fd
                cur_cd.bk_fd_cls.font_name = pre_fd
        return chars_data

    @staticmethod
    def __get_chars_state(chars):
        if re.match('^[\t -~]*$', chars):
            state = 'only ascii'
        elif re.match('^[^\t -~]*$', chars):
            state = 'only kanji'
        else:
            state = 'mix'
        return state

    @staticmethod
    def _get_ascii_and_kanji_font(font):
        if re.match('^(.*) / (.*)$', font):
            ascii_font = re.sub('^(.*) / (.*)$', '\\1', font)
            kanji_font = re.sub('^(.*) / (.*)$', '\\2', font)
        else:
            ascii_font = font
            kanji_font = font
        if ascii_font == '=':
            ascii_font = kanji_font
        return ascii_font, kanji_font

    @classmethod
    def get_raw_text(cls, chars_data):
        chars_data = cls.__cancel_fds(chars_data)
        raw_text = cls.__join_data(chars_data)
        raw_text = cls.__escape_symbols(raw_text)
        # IVS (IDEOGRAPHIC VARIATION SEQUENCE)
        raw_text = cls.__convert_ivs(raw_text)
        raw_text = cls.__restore_charcters(raw_text)
        raw_text = cls.__shrink_meaningless_font_decorations(raw_text)
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
        # CANCEL FONT FONT DECORATORS
        raw_text = cls._cancel_font_font_decorators(raw_text)
        # self.raw_text = raw_text
        return raw_text

    @staticmethod
    def _cancel_font_font_decorators(raw_text):
        # ...(=new) + @...@(=beg) + ...(=mid) + @...@(=end) + ...(=old)
        new, beg, mid, end, old = '', '', '', '', raw_text
        res = NOT_ESCAPED + '(@[^@]{1,66}@)' + '((?:.|\n)*)$'
        while re.match(res, old):
            pre = re.sub(res, '\\1', old)
            com = re.sub(res, '\\2', old)
            old = re.sub(res, '\\3', old)
            if beg == '':
                if not re.match('@' + RES_NUMBER + '@', com):
                    new += pre
                    beg = com
                else:
                    new += pre + com
            else:
                if not re.match('@' + RES_NUMBER + '@', com):
                    mid += pre
                    end = com
                    t = None
                    for c in mid:
                        if t is None:
                            if c.isascii():
                                t = 'Ascii'
                            else:
                                t = 'Nonascii'
                        elif t == 'Ascii':
                            if not c.isascii():
                                t = 'Multi'
                        elif t == 'Nonascii':
                            if c.isascii():
                                t = 'Multi'
                    f_afnt = re.sub(' */.*$', '', Form.mincho_font)
                    f_jfnt = re.sub('^.*/ *', '', Form.mincho_font)
                    if f_afnt == '=':
                        f_afnt = f_jfnt
                    t_afnt = re.sub('^@(.*?) */ *(.*)@$', '\\1', beg)
                    t_jfnt = re.sub('^@(.*?) */ *(.*)@$', '\\2', beg)
                    if t_afnt == '=':
                        t_afnt = t_jfnt
                    if beg == end and t == 'Ascii' and t_afnt == f_afnt:
                        new += mid
                    elif beg == end and t == 'Nonascii' and t_jfnt == f_jfnt:
                        new += mid
                    else:
                        new += beg + mid + end
                    beg, mid, end = '', '', ''
                else:
                    mid += pre + com
        raw_text = new + beg + mid + old
        return raw_text

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

    @classmethod
    def __cancel_fds(cls, chars_data):
        for i, cd in enumerate(chars_data):
            if i < len(chars_data) - 1:
                j = i + 1
                chars_data[i], chars_data[j] \
                    = CharsDatum.cancel_fd_cls(chars_data[i], chars_data[j])
            if (cd.chars == '\n') and (i > 0) and (i < len(chars_data) - 1):
                j, k = i - 1, i + 1
                chars_data[j], chars_data[k] \
                    = CharsDatum.cancel_fd_cls(chars_data[j], chars_data[k])
        return chars_data

    @classmethod
    def __join_data(cls, chars_data):
        raw_text = ''
        for cd in chars_data:
            cwf = cd.get_chars_with_fds()
            raw_text = CharsDatum.concatenate_imm(raw_text, cwf)
        return raw_text

    @staticmethod
    def __escape_symbols(raw_text):
        # SPACE
        raw_text = re.sub('(\n)([ \t\u3000]+)', '\\1\\\\\\2', raw_text)
        raw_text = re.sub('([ \t\u3000]+)(\n)', '\\1\\\\\\2', raw_text)
        # LENGTH REVISER
        if re.match('^(v|V|X|x|<<|<|>)=\\s*(\\-|\\+)?[0-9]+', raw_text):
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
    def __convert_ivs(raw_text):
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
    def __restore_charcters(raw_text):
        raw_text = raw_text.replace('&lt;', '<')
        raw_text = raw_text.replace('&gt;', '>')
        raw_text = raw_text.replace('&quot;', '"')
        raw_text = raw_text.replace('&amp;', '&')
        return raw_text

    @staticmethod
    def __shrink_meaningless_font_decorations(raw_text):
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
            # JUST IN CASE
            if right == '':
                level_to_break = 2
        left = re.sub(del_beg + del_end, '', left)
        left = re.sub(ins_beg + ins_end, '', left)
        return space, left + right

    @staticmethod
    def _get_remarks(xml_lines):
        remarks = []
        for xl in xml_lines:
            res = '^<w:commentReference w:id="(.*)"/>$'
            if re.match(res, xl):
                remark_id = re.sub(res, '\\1', xl)
                remarks.append(Form.remarks[remark_id])
        return remarks

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
        elif ParagraphMultiColumns.is_this_class(self):
            return 'multicolumns'
        elif ParagraphPagebreak.is_this_class(self):
            return 'pagebreak'
        elif ParagraphBreakdown.is_this_class(self):
            return 'breakdown'
        elif ParagraphRemarks.is_this_class(self):
            return 'remarks'
        elif ParagraphFootnotes.is_this_class(self):
            return 'footnotes'
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
        elif paragraph_class == 'multicolumns':
            return ParagraphMultiColumns(self)
        elif paragraph_class == 'pagebreak':
            return ParagraphPagebreak(self)
        elif paragraph_class == 'breakdown':
            return ParagraphBreakdown(self)
        elif paragraph_class == 'remarks':
            return ParagraphRemarks(self)
        elif paragraph_class == 'footnotes':
            return ParagraphFootnotes(self)
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
        # rp_sty = rp.style
        # rp_alg = rp.alignment
        # rp_rtx = rp.raw_text_doi
        # rp_img = rp.images
        # rp_fsz = Document.font_size
        return False

    def __init__(self, raw_paragraph):
        # RECEIVED
        self.raw_paragraph_number = raw_paragraph.raw_paragraph_number
        self.has_removed = raw_paragraph.has_removed
        self.xml_lines = raw_paragraph.xml_lines
        self.raw_class = raw_paragraph.raw_class
        self.style = raw_paragraph.style
        self.alignment = raw_paragraph.alignment
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
        self.footnotes = raw_paragraph.footnotes
        self.remarks = raw_paragraph.remarks
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
        self.tab_config = []
        self.pre_text_to_write = ''
        self.post_text_to_write = ''
        self.text_to_write_with_reviser = ''
        self.char_spacing = 0.0
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
        self.char_spacing = self._get_char_spacing(self.xml_lines)
        self.tab_config = self._get_tab_config(self.xml_lines)
        # EXECUTION
        self.md_lines_text = self._get_md_lines_text(self.md_text)
        self.text_to_write = self._get_text_to_write()
        self.text_to_write_with_reviser \
            = self._get_text_to_write_with_reviser()

    @classmethod
    def _get_section_depths(cls, raw_text, should_record=False):
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
            = Paragraph.get_font_revisers_and_md_text(raw_text)
        # IMAGE (CAPTION)
        if self.paragraph_class == 'image':
            raw_text += ''.join(reversed(tail_font_revisers))
            tail_font_revisers = []
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
    def get_font_revisers_and_md_text(raw_text):
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

    @staticmethod
    def _get_section_states():
        ss = ParagraphSection.states
        states \
            = [[ss[i][j] for j in range(len(ss[i]))] for i in range(len(ss))]
        return states

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
        f_size = Form.font_size
        lnsp = Form.line_spacing
        xls = self.xml_lines
        style = self.style
        paragraph_class = self.paragraph_class
        head_font_revisers = self.head_font_revisers
        tail_font_revisers = self.tail_font_revisers
        length_docx \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        rl = {'sb': 0.0, 'sa': 0.0, 'sl': 0.0, 'if': 0.0,
              'ih': 0.0, 'il': 0.0, 'ir': 0.0, 'tw': 0.0}
        if style is not None:
            for k in style.raw_length:
                rl[k] = style.raw_length[k]
        is_changed = False
        for xl in xls:
            if re.match('^<w:pPrChange( .*[^/])?>$', xl):
                is_changed = True
            if re.match('^</w:pPrChange( .*[^/])?>$', xl):
                is_changed = False
            if is_changed:
                continue
            rl['sb'] = XML.get_value('w:spacing', 'w:before', rl['sb'], xl)
            rl['sa'] = XML.get_value('w:spacing', 'w:after', rl['sa'], xl)
            rl['sl'] = XML.get_value('w:spacing', 'w:line', rl['sl'], xl)
            rl['if'] = XML.get_value('w:ind', 'w:firstLine', rl['if'], xl)
            rl['ih'] = XML.get_value('w:ind', 'w:hanging', rl['ih'], xl)
            rl['il'] = XML.get_value('w:ind', 'w:left', rl['il'], xl)
            rl['ir'] = XML.get_value('w:ind', 'w:right', rl['ir'], xl)
            rl['tw'] = XML.get_value('w:tblInd', 'w:w', rl['tw'], xl)
        length_docx['space before'] = rl['sb'] / 20 / f_size / lnsp
        length_docx['space after'] = rl['sa'] / 20 / f_size / lnsp
        ls = 0.0
        if rl['sl'] > 0:
            if paragraph_class != 'table':
                length_docx['line spacing'] \
                    = (rl['sl'] / 20 / f_size / lnsp) - 1
            else:
                sc = 1.0
                if '---' in head_font_revisers:
                    sc = 0.6
                elif '--' in head_font_revisers:
                    sc = 0.8
                elif '++' in head_font_revisers:
                    sc = 1.2
                elif '+++' in head_font_revisers:
                    sc = 1.4
                for fr in head_font_revisers:
                    res = '^@(' + RES_NUMBER + ')@$'
                    if re.match(res, fr):
                        c_size = float(re.sub(res, '\\1', fr))
                        if c_size > 0:
                            sc = c_size / Form.font_size
                length_docx['line spacing'] \
                    = (rl['sl'] / 20 / f_size / sc / TABLE_LINE_SPACING) - 1
        # MODIFY SPACE BEFORE AND AFTER
        if rl['sb'] != 0 or rl['sa'] != 0 or rl['sl'] != 0:
            ls = (rl['sl'] / 20 / f_size / lnsp) - 1
            ls80 = ls * .80
            ls20 = ls * .20
            if length_docx['space before'] >= ls80 * 0.33333:
                length_docx['space before'] += ls80
            else:
                length_docx['space before'] *= 4
            if length_docx['space after'] >= ls20 * 0.33333:
                length_docx['space after'] += ls20
            else:
                length_docx['space after'] *= 4
        length_docx['first indent'] = (rl['if'] - rl['ih']) / 20 / f_size
        length_docx['left indent'] = (rl['il'] + rl['tw']) / 20 / f_size
        length_docx['right indent'] = rl['ir'] / 20 / f_size
        # AUTO NUMBERING STYLE
        ans_key = AutoNumberingStyle.get_style_key_from_xml_lines(xls)
        if ans_key is not None:
            ans = Form.auto_numbering_styles[ans_key]
            rl['if'], rl['ih'], rl['il'] = None, None, None
            for xl in xls:
                rl['if'] = XML.get_value('w:ind', 'w:firstLine', rl['if'], xl)
                rl['ih'] = XML.get_value('w:ind', 'w:hanging', rl['ih'], xl)
                rl['il'] = XML.get_value('w:ind', 'w:left', rl['il'], xl)
            if rl['if'] is None and rl['ih'] is None:
                length_docx['first indent'] \
                    = ans.raw_first_indent / 20 / f_size
            if rl['il'] is None:
                length_docx['left indent'] \
                    = ans.raw_left_indent / 20 / f_size
        # （１）, （ア）, （ａ）
        paragraph_class = self.paragraph_class
        raw_text = self.raw_text
        res = '^（([0-9０-９]+|[ｱ-ﾝア-ン]+|[a-zａ-ｚ]+)）'
        if paragraph_class == 'section':
            if re.match(res, raw_text):
                length_docx['first indent'] += 1.0
        for ln in length_docx:
            length_docx[ln] = round(length_docx[ln], 2)
        # self.length_docx = length_docx
        return length_docx

    def _get_length_clas(self):
        length_docx = self.length_docx
        paragraph_class = self.paragraph_class
        head_section_depth = self.head_section_depth
        tail_section_depth = self.tail_section_depth
        section_states = self.section_states
        proper_depth = self.proper_depth
        xml_lines = self.xml_lines
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
        for ln in length_clas:
            length_clas[ln] = round(length_clas[ln], 2)
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
        for ln in length_conf:
            length_conf[ln] = round(length_conf[ln], 2)
        # self.length_conf = length_conf
        return length_conf

    def _get_length_supp(self):
        length_supp \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        for ln in length_supp:
            length_supp[ln] = round(length_supp[ln], 2)
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
            length_revi[ln] = length_docx[ln] \
                - length_clas[ln] - length_conf[ln] + length_supp[ln]
            length_revi[ln] = round(length_revi[ln], 2)
        # self.length_revi = length_revi
        return length_revi

    @classmethod
    def _get_length_revisers(cls, length_revi):
        length_revisers = []
        vs = cls.__get_vlength_string(length_revi['space before'])
        if float(vs) < -0.1 or float(vs) > 0.1:
            length_revisers.append('v=' + vs)
        vs = cls.__get_vlength_string(length_revi['space after'])
        if float(vs) < -0.1 or float(vs) > 0.1:
            length_revisers.append('V=' + vs)
        vs = cls.__get_vlength_string(length_revi['line spacing'])
        if float(vs) < -0.1 or float(vs) > 0.1:
            length_revisers.append('X=' + vs)
        # WHAT YOU SEE IS THE SUM OF 'first indent' AND 'left indent'
        hs1 = cls.__get_hlength_string(- length_revi['first indent']
                                       - length_revi['left indent'])
        hs2 = cls.__get_hlength_string(- length_revi['left indent'])
        hs = str(float(hs1) - float(hs2))
        if float(hs) > 0:
            hs = '+' + hs
        # hs = cls.__get_hlength_string(- length_revi['first indent'])
        if float(hs) < -0.1 or float(hs) > 0.1:
            length_revisers.append('<<=' + hs)
        hs = cls.__get_hlength_string(- length_revi['left indent'])
        if float(hs) < -0.1 or float(hs) > 0.1:
            length_revisers.append('<=' + hs)
        hs = cls.__get_hlength_string(- length_revi['right indent'])
        if float(hs) < -0.1 or float(hs) > 0.1:
            length_revisers.append('>=' + hs)
        # self.length_revisers = length_revisers
        return length_revisers

    @staticmethod
    def __get_vlength_string(length):
        # FRACTION
        if length < 0:
            porm = '-'
        elif length == 0:
            porm = ''
        else:
            porm = '+'
        i_part = str(int(abs(length)))
        d_part = abs(length - int(length))
        if d_part > 0.313 and d_part < 0.353:  # 1/3=0.3333...
            return porm + i_part + '.33'
        if d_part > 0.647 and d_part < 0.687:  # 2/3=0.6666...
            return porm + i_part + '.67'
        if d_part > 0.230 and d_part < 0.270:  # 1/4=0.25
            return porm + i_part + '.25'
        if d_part > 0.730 and d_part < 0.770:  # 3/4=0.75
            return porm + i_part + '.75'
        if d_part > 0.147 and d_part < 0.187:  # 1/6=0.1666...
            return porm + i_part + '.17'
        if d_part > 0.813 and d_part < 0.853:  # 5/6=0.8333...
            return porm + i_part + '.83'
        # DECIMAL
        rounded = round(length, 1)
        if rounded < 0:
            return str(rounded)
        elif rounded == 0:
            return '0.0'
        else:
            return '+' + str(rounded)

    @staticmethod
    def __get_hlength_string(length):
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
                                if self.head_space == '':
                                    ParagraphSection.states[1][0] = 0
                                    section_states[1][0] = 0
                                    numbering_revisers.insert(0, '##=1')
                                    length_revisers.remove('<=+1.0')
        return section_states, numbering_revisers, length_revisers

    @staticmethod
    def _get_char_spacing(xml_lines):
        for xl in xml_lines:
            if xl == '</w:rPr>':
                return 0.0
            sp = XML.get_value('w:spacing', 'w:val', 0.0, xl)
            if sp != 0.0:
                cs = (sp / Form.font_size / 20) - DEFAULT_CHAR_SPACING
                return round(cs, 2)
        return 0.0

    @staticmethod
    def _get_tab_config(xml_lines):
        tab_config = []
        is_in_ppr = False
        is_in_tab = False
        for xl in xml_lines:
            if xl == '<w:pPr>':
                is_in_ppr = True
            elif xl == '</w:pPr>':
                break
            elif xl == '<w:tabs>':
                is_in_tab = True
            elif xl == '</w:tabs>':
                break
            elif is_in_ppr and is_in_tab:
                res = '^<w:tab w:val="([^"]+)" w:pos="([0-9]+)"/>$'
                if re.match(res, xl):
                    ali = re.sub(res, '\\1', xl)
                    wid = int(re.sub(res, '\\2', xl))
                    wc = str(round(wid / Form.font_size / 20, 2))
                    wc = re.sub('\\.?0+$', '', wc)
                    if ali == 'right':
                        tc = '@' + wc + ':'
                    elif ali == 'center':
                        tc = ':' + '@' + wc + ':'
                    else:
                        tc = '@' + wc
                    tab_config.append(tc)
        return tab_config

    def _get_md_lines_text(self, md_text):
        paragraph_class = self.paragraph_class
        # FOR TRAILING WHITE SPACE
        md_text = re.sub('  \n', '  \\\n', md_text)
        if False:
            pass
        elif paragraph_class == 'chapter':
            md_lines_text = LineTruncation(md_text).get_truncated_md_text()
        elif paragraph_class == 'section':
            md_lines_text = LineTruncation(md_text).get_truncated_md_text()
        elif paragraph_class == 'list':
            md_lines_text = LineTruncation(md_text).get_truncated_md_text()
        elif paragraph_class == 'alignment':
            md_lines_text = ''
            for mt in md_text.split('\n'):
                md_lines_text \
                    += LineTruncation(mt).get_truncated_md_text() + '\n'
            md_lines_text = re.sub('\n+$', '', md_lines_text)
        elif paragraph_class == 'sentence':
            md_lines_text = LineTruncation(md_text).get_truncated_md_text()
        else:
            md_lines_text = md_text
        return md_lines_text

    def _get_text_to_write(self):
        paper_size = Form.paper_size
        top_margin = Form.top_margin
        bottom_margin = Form.bottom_margin
        left_margin = Form.left_margin
        right_margin = Form.right_margin
        md_lines_text = self.md_lines_text
        length_docx = self.length_docx
        head_space = self.head_space
        tab_config = self.tab_config
        indent = length_docx['first indent'] \
            + length_docx['left indent'] \
            + length_docx['right indent']
        unit = 12 * 2.54 / 72 / 2
        width_cm = PAPER_WIDTH[paper_size] - left_margin - right_margin \
            - (indent * unit)
        height_cm = PAPER_HEIGHT[paper_size] - top_margin - bottom_margin
        region_cm = (width_cm, height_cm)
        res = '^((?:.|\n)*)(' + RES_IMAGE_WITH_SIZE + ')((?:.|\n)*)$'
        text_to_write = head_space
        while re.match(res, md_lines_text):
            text_to_write += re.sub(res, '\\1', md_lines_text)
            img_text = re.sub(res, '\\2', md_lines_text)
            text_to_write \
                += ParagraphImage.replace_with_fixed_size(img_text, region_cm)
            md_lines_text = re.sub(res, '\\7', md_lines_text)
        text_to_write += md_lines_text
        # TAB
        for tc in tab_config:
            if '\t' not in text_to_write:
                break
            text_to_write = text_to_write.replace('\t', '< ' + tc + ' >', 1)
        # FOOTNOTES
        res = '^((?:.|\n)*)\\^{([0-9]+)）}'
        while re.match(res, text_to_write):
            text_to_write = re.sub(res, '\\1[^\\2]', text_to_write)
        return text_to_write

    def _get_text_to_write_with_reviser(self):
        paragraph_class = self.paragraph_class
        numbering_revisers = self.numbering_revisers
        length_revisers = self.length_revisers
        char_spacing = self.char_spacing
        head_font_revisers = self.head_font_revisers
        tail_font_revisers = self.tail_font_revisers
        text_to_write = self.text_to_write
        pre_text_to_write = self.pre_text_to_write
        post_text_to_write = self.post_text_to_write
        attached_pagebreak = self.attached_pagebreak
        footnotes = self.footnotes
        # FONT REVISERS
        head_pair_font_revisers = []
        head_single_font_revisers = []
        tail_pair_font_revisers = []
        tail_single_font_revisers = []
        for rev in head_font_revisers:
            partner = FontDecorator.get_partner(rev)
            if partner in tail_font_revisers:
                head_pair_font_revisers.append(rev)
            else:
                head_single_font_revisers.append(rev)
        for rev in tail_font_revisers:
            partner = FontDecorator.get_partner(rev)
            if partner in head_font_revisers:
                tail_pair_font_revisers.append(rev)
            else:
                tail_single_font_revisers.append(rev)
        # CUT SYMBOLS
        has_left_sharp = False
        is_left_or_center_alignment = False
        is_center_or_right_alignment = False
        if re.match('^# (.|\n)*$', text_to_write):
            text_to_write = re.sub('^# ', '', text_to_write)
            has_left_sharp = True
        elif re.match('^: (.|\n)*$', text_to_write):
            text_to_write = re.sub('^: ', '', text_to_write)
            is_left_or_center_alignment = True
        if re.match('^(.|\n)* :$', text_to_write):
            text_to_write = re.sub(' :$', '', text_to_write)
            is_center_or_right_alignment = True
        # INITIALIZE
        ttwwr = ''
        # PRE TEXT
        if pre_text_to_write != '':
            ttwwr += pre_text_to_write + '\n'
        # LENGTH REVISER
        for rev in length_revisers:
            ttwwr += rev + ' '
        if char_spacing != 0.0:
            if char_spacing > 0.0:
                ttwwr += 'x=+' + str(char_spacing) + ' '
            else:
                ttwwr += 'x=' + str(char_spacing) + ' '
        if re.match('^(.|\n)* $', ttwwr):
            ttwwr = re.sub(' $', '\n', ttwwr)
        for rev in numbering_revisers:
            ttwwr += rev + ' '
        if re.match('^(.|\n)* $', ttwwr):
            ttwwr = re.sub(' $', '\n', ttwwr)
        if has_left_sharp:
            ttwwr += '# '
        # LEFT SYMBOL
        if len(head_pair_font_revisers) > 0:
            ttwwr += ''.join(head_pair_font_revisers) + '\n'
        if is_left_or_center_alignment:
            ttwwr += ': '
        if len(head_single_font_revisers) > 0:
            ttwwr += ''.join(head_single_font_revisers)
            if not is_left_or_center_alignment:
                ttwwr += '\n'
        # TEXT
        ttwwr += text_to_write
        # RIGHT SYMBOL
        if len(tail_single_font_revisers) > 0:
            if not is_center_or_right_alignment:
                ttwwr += '\n'
            ttwwr += ''.join(tail_single_font_revisers)
        if is_center_or_right_alignment:
            ttwwr += ' :'
        if len(tail_pair_font_revisers) > 0:
            ttwwr += '\n' + ''.join(tail_pair_font_revisers)
        # POST TEXT
        if post_text_to_write != '':
            ttwwr += '\n' + post_text_to_write
        # PAGE BREAK
        if paragraph_class != 'pagebreak':
            if attached_pagebreak == 'pgbr':
                ttwwr += '\n\n<pgbr>'
            if attached_pagebreak == 'Pgbr':
                ttwwr += '\n\n<Pgbr>'
        # FOOTNOTES
        if footnotes != {}:
            ttwwr += '\n'
            for _fnid in footnotes:
                ttwwr += '\n[^' + _fnid + ']: ' + footnotes[_fnid]
        text_to_write_with_reviser = ttwwr
        # self.text_to_write_with_reviser = text_to_write_with_reviser
        return text_to_write_with_reviser

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
        rp = raw_paragraph
        xls = rp.xml_lines
        if ParagraphConfiguration.is_this_class(rp):
            return False
        if ParagraphTable.is_this_class(rp):
            return False
        if ParagraphHorizontalLine.is_this_class(rp):
            return False
        if rp.raw_class == 'w:p':
            return False
        if rp.raw_text == '':
            has_run = False
            for xl in xls:
                if re.match('^<w:r( .*)?>$', xl):
                    has_run = True
            if not has_run:
                return True
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
        hfrs, tfrs, mtx = Paragraph.get_font_revisers_and_md_text(rp_rtx)
        if re.match('^\\s*$', mtx):
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
            = Paragraph.get_font_revisers_and_md_text(raw_text)
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
                    = self.__decompose_text(res, raw_text, -1, -1)
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
    def __decompose_text(res, raw_text, num1, num2):
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
        raw_text = re.sub('（([0-9０-９]+|[ｱ-ﾝア-ン]|[a-zａ-ｚ])）',
                          '(\\1) ', raw_text)
        rss = cls.res_symbols
        rfd = RES_FONT_DECORATORS
        rre = cls.res_rest
        rnm = cls.res_number
        head_section_depth = 0
        tail_section_depth = 0
        for xdepth in range(1, len(rss)):
            res = '^(?:\\\\\\s+)?' + rfd + '\\s*' + rss[xdepth] + rre + '$'
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
        xl_size = Form.font_size * 1.4
        xml_lines = self.xml_lines
        rss = self.res_symbols
        rre = self.res_rest
        rnm = self.res_number
        numbering_revisers = []
        head_font_revisers, tail_font_revisers, raw_text \
            = Paragraph.get_font_revisers_and_md_text(raw_text)
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
        # "　１　…" -> "１　…"
        raw_text = re.sub('^\\s+', '', raw_text)
        # （１）, （ア）, （ａ）
        raw_text = re.sub('^（([0-9０-９]+|[ｱ-ﾝア-ン]|[a-zａ-ｚ])）',
                          '(\\1) ', raw_text)
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
                    = self.__decompose_text(res, raw_text, beg_num, end_num)
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
                    tail_font_revisers.insert(0, '---')
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
    def __decompose_text(res, raw_text, beg_num, end_num):
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
            = Paragraph.get_font_revisers_and_md_text(raw_text)
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
                        = self.__decompose_text(res, raw_text, xdepth, -1)
                    head_symbol = '  ' * xdepth + '1. '
                    self._step_states(xdepth, 0)
                    numbering_revisers \
                        = self._get_numbering_revisers(xdepth, state)
                break
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            head_symbol + raw_text

    @staticmethod
    def __decompose_text(res, raw_text, xdepth, num):
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

    def _get_revisers_and_md_text(self, raw_text):
        xml_lines = self.xml_lines
        tbl_alg = self.__get_table_alignment(xml_lines)
        xml_tbl = self.__get_xml_table(xml_lines)
        mrg_tbl = self.__get_merge_table(xml_tbl)
        num_row, num_clm = self.__get_table_size(xml_tbl)
        v_raw_hgt, h_raw_wid \
            = self.__get_raw_length(xml_lines, num_row, num_clm)
        txt_tbl, h_frs, t_frs \
            = self.__get_txt_table_and_font_revisers(xml_tbl)
        font_size = self.__get_font_size(h_frs, t_frs)
        v_chr_hgt, h_chr_wid \
            = self.__get_length_in_char_units(v_raw_hgt, h_raw_wid,
                                              font_size)
        v_alig_tbl, h_alig_tbl, v_rule_tbl, h_rule_tbl \
            = self.__get_cell_state(xml_tbl)
        std_row, std_clm \
            = self.__get_standard_row_and_column(num_row, v_alig_tbl,
                                                 num_clm, h_alig_tbl)
        v_conf_clm, h_conf_row = self.__get_confs(num_row, num_clm,
                                                  std_row, std_clm,
                                                  v_chr_hgt, h_chr_wid,
                                                  v_alig_tbl, h_alig_tbl,
                                                  v_rule_tbl, h_rule_tbl)
        md_text = self.__get_md_text(tbl_alg, txt_tbl, mrg_tbl,
                                     std_row, std_clm,
                                     v_alig_tbl, h_alig_tbl,
                                     v_rule_tbl, h_rule_tbl,
                                     v_conf_clm, h_conf_row)
        md_text = self.__split_long_lines(md_text)
        numbering_revisers = []
        head_font_revisers, tail_font_revisers = h_frs, t_frs
        return numbering_revisers, head_font_revisers, tail_font_revisers, \
            md_text

    @staticmethod
    def __get_table_alignment(xml_lines):
        tbl_alig = 'center'
        res_tpr_beg = '^<w:tblPr( .*)?>$'
        res_tpr_end = '^</w:tblPr( .*)?>$'
        res_tpr_alg = '^<w:jc(?: .*)? w:val=[\'"]([a-z]*)[\'"](?: .*)?/>$'
        is_in_tpr = False
        for xl in xml_lines:
            if re.match(res_tpr_beg, xl):
                is_in_tpr = True
            elif re.match(res_tpr_end, xl):
                break
            elif re.match(res_tpr_alg, xl):
                if is_in_tpr:
                    tbl_alig = re.sub(res_tpr_alg, '\\1', xl)
        return tbl_alig

    @staticmethod
    def __get_xml_table(xml_lines):
        xml_tbl = []
        res_tbl_beg = '^<w:tbl( .*)?>$'
        res_tbl_end = '^</w:tbl( .*)?>$'
        depth = 0
        res_row_beg = '^<w:tr( .*)?>$'
        res_row_end = '^</w:tr( .*)?>$'
        is_in_row = False
        res_cel_beg = '^<w:tc( .*)?>$'
        res_cel_end = '^</w:tc( .*)?>'
        is_in_cel = False
        max_row = -1
        for xl in xml_lines:
            if re.match(res_tbl_beg, xl):
                depth += 1
            elif re.match(res_tbl_end, xl):
                depth -= 1
            if depth == 1 and re.match(res_row_beg, xl):
                xml_row = []
                is_in_row = True
            elif depth == 1 and re.match(res_row_end, xl):
                max_row = max(max_row, len(xml_row))
                xml_tbl.append(xml_row)
                is_in_row = False
            elif depth == 1 and re.match(res_cel_beg, xl):
                xml_cel = []
                span_h = 1
                is_in_cel = True
            elif depth == 1 and re.match(res_cel_end, xl):
                xml_row.append(xml_cel)
                for i in range(1, span_h):
                    xml_row.append([])
                is_in_cel = False
            elif is_in_cel:
                if ('</w:p>' in xml_cel) and re.match('<w:p( .*)?>', xl):
                    xml_cel.append('<w:br/>')
                xml_cel.append(xl)
                span_h = XML.get_value('w:gridSpan', 'w:val', span_h, xl)
        for xml_row in xml_tbl:
            for i in range(len(xml_row), max_row):
                xml_row.append([])
        return xml_tbl

    @staticmethod
    def __get_merge_table(xml_tbl):
        merge_tbl = []
        for i in range(len(xml_tbl)):
            merge_row = []
            for j in range(len(xml_tbl[i])):
                h_span = 1
                for xml in xml_tbl[i][j]:
                    h_span \
                        = XML.get_value('w:gridSpan', 'w:val', h_span, xml)
                v_span = 1
                if '<w:vMerge w:val="restart"/>' in xml_tbl[i][j]:
                    for k in range(i + 1, len(xml_tbl)):
                        if '<w:vMerge/>' in xml_tbl[k][j]:
                            v_span += 1
                        else:
                            break
                elif '<w:vMerge/>' in xml_tbl[i][j]:
                    v_span = -1
                if (h_span >= 2 and v_span >= 1) or\
                   (h_span >= 1 and v_span >= 2):
                    merge_cel = '@' + str(h_span) + 'x' + str(v_span)
                else:
                    merge_cel = ''
                merge_row.append(merge_cel)
            merge_tbl.append(merge_row)
        return merge_tbl

    @staticmethod
    def __get_table_size(xml_tbl):
        num_row = len(xml_tbl)
        nc = []
        for row in xml_tbl:
            nc.append(len(row))
        num_clm = max(nc)
        return num_row, num_clm

    @staticmethod
    def __get_raw_length(xml_lines, num_row, num_clm):
        v_raw_hgt, h_raw_wid = [], []
        res_tbl_beg = '^<w:tbl( .*)?>$'
        res_tbl_end = '^</w:tbl( .*)?>$'
        depth = 0
        res_row_beg = '^<w:tr( .*)?>$'
        res_v_hgt = '^<w:trHeight(?: .*)? w:val=[\'"]([0-9]+)[\'"](?: .*)?/>$'
        res_h_wid = '^<w:gridCol(?: .*)? w:w=[\'"]([0-9]+)[\'"](?: .*)?/>$'
        n = 0
        for xl in xml_lines:
            if re.match(res_tbl_beg, xl):
                depth += 1
            elif re.match(res_tbl_end, xl):
                depth -= 1
            if depth == 1 and re.match(res_row_beg, xl):
                n += 1
            elif depth == 1 and re.match(res_v_hgt, xl):
                while len(v_raw_hgt) < n - 1:
                    v_raw_hgt.append(0)
                val = re.sub(res_v_hgt, '\\1', xl)
                v_raw_hgt.append(int(val))
            elif depth == 1 and re.match(res_h_wid, xl):
                val = re.sub(res_h_wid, '\\1', xl)
                h_raw_wid.append(int(val))
        while len(v_raw_hgt) < num_row:
            v_raw_hgt.append(0)
        while len(h_raw_wid) < num_clm:
            h_raw_wid.append(0)
        return v_raw_hgt, h_raw_wid

    def __get_txt_table_and_font_revisers(self, xml_tbl):
        par_xml_tbl = self.__get_par_xml_tbl(xml_tbl)
        cd_tbl = []
        for i in range(len(par_xml_tbl)):
            cd_row = []
            for j in range(len(par_xml_tbl[i])):
                cd_cel = []
                for k in range(len(par_xml_tbl[i][j])):
                    par = par_xml_tbl[i][j][k]
                    style, alignment, chars_data, raw_text, \
                        images, footnotes \
                        = RawParagraph.get_raw_text_and_etc(par)
                    cd_cel.append(chars_data)
                cd_row.append(cd_cel)
            cd_tbl.append(cd_row)
        # CANCEL
        pre_cd = None
        for cd_row in cd_tbl:
            for cd_cel in cd_row:
                for cd_par in cd_cel:
                    for cur_cd in cd_par:
                        if pre_cd is not None:
                            pre_cd, cur_cd \
                                = CharsDatum.cancel_fd_cls(pre_cd, cur_cd)
                        pre_cd = cur_cd
        # FONT REVISERS
        h_frs, t_frs = [], []
        fr_fd_cls = None
        bk_fd_cls = None
        for cd_row in cd_tbl:
            for cd_cel in cd_row:
                for cd_par in cd_cel:
                    if len(cd_par) > 0:
                        if fr_fd_cls is None:
                            fr_fd_cls = cd_par[0].fr_fd_cls
                        bk_fd_cls = cd_par[-1].bk_fd_cls
        # fr_fd_cls = cd_tbl[0][0][0][0].fr_fd_cls
        # bk_fd_cls = cd_tbl[-1][-1][-1][-1].bk_fd_cls
        if (fr_fd_cls is not None) and (bk_fd_cls is not None):
            fr, bk = fr_fd_cls.font_name, bk_fd_cls.font_name
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.font_name, bk_fd_cls.font_name = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.font_scale, bk_fd_cls.font_scale
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.font_scale, bk_fd_cls.font_scale = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.font_width, bk_fd_cls.font_width
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.font_width, bk_fd_cls.font_width = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.italic, bk_fd_cls.italic
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.italic, bk_fd_cls.italic = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.bold, bk_fd_cls.bold
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.bold, bk_fd_cls.bold = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.strike, bk_fd_cls.strike
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.strike, bk_fd_cls.strike = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.frame, bk_fd_cls.frame
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.frame, bk_fd_cls.frame = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.underline, bk_fd_cls.underline
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.underline, bk_fd_cls.underline = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.font_color, bk_fd_cls.font_color
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.font_color, bk_fd_cls.font_color = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.highlight_color, bk_fd_cls.highlight_color
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.highlight_color, bk_fd_cls.highlight_color = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
            fr, bk = fr_fd_cls.sub_or_sup, bk_fd_cls.sub_or_sup
            if fr != '' and bk != '' and FontDecorator.get_partner(fr) == bk:
                fr_fd_cls.sub_or_sup, bk_fd_cls.sub_or_sup = '', ''
                h_frs.append(fr)
                t_frs.append(bk)
                t_frs.reverse()
        # TEXT
        txt_tbl = []
        for cd_row in cd_tbl:
            txt_row = []
            for cd_cel in cd_row:
                txt_cel = []
                for cd_par in cd_cel:
                    raw_text = RawParagraph.get_raw_text(cd_par)
                    txt_cel.append(raw_text)
                txt_row.append(txt_cel)
            txt_tbl.append(txt_row)
        # RETURN
        return txt_tbl, h_frs, t_frs

    @staticmethod
    def __get_par_xml_tbl(xml_tbl):
        par_xml_tbl = []
        for xml_row in xml_tbl:
            par_xml_row = []
            for xml_cel in xml_row:
                par_xml_cel = []
                is_in_par = False
                for xml in xml_cel:
                    if re.match('<w:p( .+)?>', xml):
                        is_in_par = True
                        par_xml = []
                        par_xml.append(xml)
                    elif re.match('</w:p( .+)?>', xml):
                        is_in_par = False
                        par_xml.append(xml)
                        par_xml_cel.append(par_xml)
                        par_xml = []
                    elif is_in_par:
                        par_xml.append(xml)
                par_xml_row.append(par_xml_cel)
            par_xml_tbl.append(par_xml_row)
        return par_xml_tbl

    @staticmethod
    def __get_font_revisers(txt_tab):
        # GET CANDIDATES
        head_font_revisers, tail_font_revisers = [], []
        for row in txt_tab:
            for cell in row:
                h_frs, t_frs, _ \
                    = Paragraph.get_font_revisers_and_md_text(cell)
                for fr in h_frs:
                    if fr not in head_font_revisers:
                        head_font_revisers.append(fr)
                for fr in t_frs:
                    if fr not in tail_font_revisers:
                        tail_font_revisers.append(fr)
        tail_font_revisers.reverse()
        # CHECK THE FREQUENCY
        total = 0
        h_freq = [0 for i in head_font_revisers]
        t_freq = [0 for i in tail_font_revisers]
        for row in txt_tab:
            for cell in row:
                if cell == '':
                    continue
                total += 1
                h_frs, t_frs, _ \
                    = Paragraph.get_font_revisers_and_md_text(cell)
                for i, fr in enumerate(head_font_revisers):
                    if fr in h_frs:
                        h_freq[i] += 1
                for i, fr in enumerate(tail_font_revisers):
                    if fr in t_frs:
                        t_freq[i] += 1
        for i in range(len(h_freq) - 1, -1, -1):
            if (h_freq[i] * 2) < total:
                fr = head_font_revisers[i]
                head_font_revisers.remove(fr)
        for i in range(len(t_freq) - 1, -1, -1):
            if (t_freq[i] * 2) < total:
                fr = tail_font_revisers[i]
                tail_font_revisers.remove(fr)
        # FIND THE PARTNER
        for h_fr in head_font_revisers:
            t_fr = h_fr
            t_fr = t_fr.replace('>', '\\>')
            t_fr = t_fr.replace('<', '\\<')
            t_fr = t_fr.replace('\\>', '<')
            t_fr = t_fr.replace('\\<', '>')
            if t_fr not in tail_font_revisers:
                head_font_revisers.remove(h_fr)
        for t_fr in tail_font_revisers:
            h_fr = t_fr
            h_fr = h_fr.replace('>', '\\>')
            h_fr = h_fr.replace('<', '\\<')
            h_fr = h_fr.replace('\\>', '<')
            h_fr = h_fr.replace('\\<', '>')
            if h_fr not in head_font_revisers:
                tail_font_revisers.remove(t_fr)
        return head_font_revisers, tail_font_revisers

    @staticmethod
    def __get_font_size(h_frs, t_frs):
        font_size = Form.font_size * 1.0
        for fr in h_frs:
            res = '^@(' + RES_NUMBER + ')@$'
            if re.match(res, fr):
                c_size = float(re.sub(res, '\\1', fr))
                if c_size > 0:
                    font_size = c_size
            elif fr == '---':
                font_size = Form.font_size * 0.6
            elif fr == '--':
                font_size = Form.font_size * 0.8
            elif fr == '++':
                font_size = Form.font_size * 1.2
            elif fr == '+++':
                font_size = Form.font_size * 1.4
        return font_size

    @staticmethod
    def __get_length_in_char_units(v_raw_hgt, h_raw_wid, font_size):
        v_chr_hgt, h_chr_wid = [], []
        for rh in v_raw_hgt:
            ch = rh / font_size / 10
            # ch = (rh / font_size / 10) - (BASIC_TABLE_CELL_HEIGHT * 2)
            if ch < 0:
                ch = 0
            ch = round(ch)
            v_chr_hgt.append(ch)
        for rw in h_raw_wid:
            cw = (rw / font_size / 10) - (BASIC_TABLE_CELL_WIDTH * 2)
            if cw < 0:
                cw = 0
            cw = round(cw)
            h_chr_wid.append(cw)
        return v_chr_hgt, h_chr_wid

    def __get_cell_state(self, xml_tbl):
        par_xml_tbl = self.__get_par_xml_tbl(xml_tbl)
        v_alig_tbl, h_alig_tbl, v_rule_tbl, h_rule_tbl = [], [], [], []
        res_tcpr_beg = '^<w:tcPr(?: .*)?>$'
        res_tcpr_end = '^</w:tcPr(?: .*)?>$'
        res_v_top = '^<w:vAlign(?: .*)? w:val=[\'"]top[\'"](?: .*)?/>$'
        res_v_cen = '^<w:vAlign(?: .*)? w:val=[\'"]center[\'"](?: .*)?/>$'
        res_v_bot = '^<w:vAlign(?: .*)? w:val=[\'"]bottom[\'"](?: .*)?/>$'
        res_tcborders_beg = '^<w:tcBorders(?: .*)?>$'
        res_tcborders_end = '^</w:tcBorders(?: .*)?>$'
        res_v_nil = '^<w:right(?: .*)? w:val=[\'"]nil[\'"]( .*)?/>$'
        res_v_dbl = '^<w:right(?: .*)? w:val=[\'"]double[\'"]( .*)?/>$'
        res_h_nil = '^<w:bottom(?: .*)? w:val=[\'"]nil[\'"]( .*)?/>$'
        res_h_dbl = '^<w:bottom(?: .*)? w:val=[\'"]double[\'"]( .*)?/>$'
        res_ppr_beg = '^<w:pPr(?: .*)?>$'
        res_ppr_end = '^</w:pPr(?: .*)?>$'
        res_h_lef = '^<w:jc(?: .*)? w:val=[\'"]left[\'"](?: .*)?/>$'
        res_h_cen = '^<w:jc(?: .*)? w:val=[\'"]center[\'"](?: .*)?/>$'
        res_h_rig = '^<w:jc(?: .*)? w:val=[\'"]right[\'"](?: .*)?/>$'
        is_in_tcpr, is_in_tcborders, is_in_ppr = False, False, False
        for i, par_xml_row in enumerate(par_xml_tbl):
            v_alig_row, h_alig_row = [], []
            v_rule_row, h_rule_row = [], []
            for j, par_xml_cel in enumerate(par_xml_row):
                #
                v_alig_val, v_rule_val, h_rule_val = '', '', ''
                for k, xml in enumerate(xml_tbl[i][j]):
                    if re.match(res_tcpr_beg, xml):
                        is_in_tcpr = True
                    elif re.match(res_tcpr_end, xml):
                        is_in_tcpr = False
                    elif is_in_tcpr and re.match(res_v_top, xml):
                        v_alig_val = 'T'  # top
                    elif is_in_tcpr and re.match(res_v_cen, xml):
                        v_alig_val = 'C'  # vertical center
                    elif is_in_tcpr and re.match(res_v_bot, xml):
                        v_alig_val = 'B'  # bottom
                    elif re.match(res_tcborders_beg, xml):
                        is_in_tcborders = True
                    elif re.match(res_tcborders_end, xml):
                        is_in_tcborders = False
                    elif is_in_tcborders and re.match(res_v_nil, xml):
                        v_rule_val = '^'  # nil
                    elif is_in_tcborders and re.match(res_v_dbl, xml):
                        v_rule_val = '='  # double
                    elif is_in_tcborders and re.match(res_h_nil, xml):
                        h_rule_val = '^'  # nil
                    elif is_in_tcborders and re.match(res_h_dbl, xml):
                        h_rule_val = '='  # double
                #
                v_alig_cel, h_alig_cel = [], []
                for k, par_xml_par in enumerate(par_xml_cel):
                    h_alig_val = ''
                    for xml in par_xml_par:
                        if re.match(res_ppr_beg, xml):
                            is_in_ppr = True
                        elif re.match(res_ppr_end, xml):
                            is_in_ppr = False
                        elif is_in_ppr and re.match(res_h_lef, xml):
                            h_alig_val = 'L'  # left
                        elif is_in_ppr and re.match(res_h_cen, xml):
                            h_alig_val = 'C'  # horizontal center
                        elif is_in_ppr and re.match(res_h_rig, xml):
                            h_alig_val = 'R'  # right
                    v_alig_cel.append(v_alig_val)
                    h_alig_cel.append(h_alig_val)
                if v_alig_cel == []:
                    v_alig_cel = ['']
                if h_alig_cel == []:
                    h_alig_cel = ['']
                v_alig_row.append(v_alig_cel)
                h_alig_row.append(h_alig_cel)
                v_rule_row.append(v_rule_val)
                h_rule_row.append(h_rule_val)
            v_alig_tbl.append(v_alig_row)
            h_alig_tbl.append(h_alig_row)
            v_rule_tbl.append(v_rule_row)
            h_rule_tbl.append(h_rule_row)
        return v_alig_tbl, h_alig_tbl, v_rule_tbl, h_rule_tbl

    @staticmethod
    def __get_standard_row_and_column(num_row, v_alig_tbl,
                                      num_clm, h_alig_tbl):
        ha_tbl = [['' for j in range(num_clm)] for i in range(num_row)]
        va_tbl = [['' for i in range(num_row)] for j in range(num_clm)]
        for i in range(num_row):
            for j in range(num_clm):
                ha_tbl[i][j] = h_alig_tbl[i][j][0]
                va_tbl[j][i] = v_alig_tbl[i][j][0]
        hm_dct, ha_dct = {}, {}
        vm_dct, va_dct = {}, {}
        # ROW
        if num_row == 1:
            std_row = 0
        else:
            for i in range(1, num_row):
                ha_str = ''
                for ht in ha_tbl[i]:
                    if ht == '' and len(ha_str) > 0:
                        ha_str += ha_str[-1]
                    else:
                        ha_str += ht
                if ha_str not in ha_dct:
                    hm_dct[ha_str] = i
                    ha_dct[ha_str] = 1
                else:
                    ha_dct[ha_str] += 1
            max_num_ha_str = max(ha_dct, key=ha_dct.get)
            std_row = hm_dct[max_num_ha_str]
        # COLUMN
        if num_clm == 1:
            std_clm = 0
        else:
            for j in range(1, num_clm):
                va_str = ''
                for vt in va_tbl[j]:
                    if vt == '' and len(va_str) > 0:
                        va_str += va_str[-1]
                    else:
                        va_str += vt
                if va_str not in va_dct:
                    vm_dct[va_str] = j
                    va_dct[va_str] = 1
                else:
                    va_dct[va_str] += 1
            max_num_va_str = max(va_dct, key=va_dct.get)
            std_clm = vm_dct[max_num_va_str]
        # RETURN
        return std_row, std_clm

    @staticmethod
    def __get_confs(num_row, num_clm,
                    std_row, std_clm,
                    v_clen_clm, h_clen_row,
                    v_alig_tbl, h_alig_tbl,
                    v_rule_tbl, h_rule_tbl):
        v_conf_clm = ['' for i in range(num_row)]
        h_conf_row = ['' for i in range(num_clm)]
        # VERTUAL CONFIGURATIONS
        for i in range(num_row):
            # ALIGNMENT AND LENGTH
            # ""=C / "-"=T / ":"=C / ":-*"=T / ":-*:"=C / "-*:"=B
            if v_alig_tbl[i][std_clm][0] == 'T':
                if v_clen_clm[i] == 0:
                    pass
                elif v_clen_clm[i] == 1:
                    v_conf_clm[i] += '-'
                else:
                    v_conf_clm[i] += ':' + '-' * (v_clen_clm[i] - 1)
            elif v_alig_tbl[i][std_clm][0] == 'B':
                if v_clen_clm[i] == 0:
                    pass
                elif v_clen_clm[i] == 1:
                    pass
                else:
                    v_conf_clm[i] += '-' * (v_clen_clm[i] - 1) + ':'
            else:  # "" or "C"
                if v_clen_clm[i] == 0:
                    v_conf_clm[i] += ''
                elif v_clen_clm[i] == 1:
                    v_conf_clm[i] += ':'
                else:
                    v_conf_clm[i] += ':' + '-' * (v_clen_clm[i] - 2) + ':'
            # RULE
            v_conf_clm[i] += h_rule_tbl[i][std_clm]
            # REMOVE DEFAULT CONFIGURATION
            default_v_conf = ':' \
                + '-' * round(BASIC_TABLE_CELL_HEIGHT - 0.5) \
                + ':'
            if re.match('^' + default_v_conf, v_conf_clm[i]):
                v_conf_clm[i] = re.sub('^' + default_v_conf, '', v_conf_clm[i])
        # HORIZONTAL CONFIGURATIONS
        for j in range(num_clm):
            # ALIGNMENT AND LENGTH
            # ""=L / "-"=L / ":"=C / ":-*"=L / ":-*:"=C / "-*:"=R
            if h_alig_tbl[std_row][j][0] == 'R':
                if h_clen_row[j] == 0:
                    pass
                elif h_clen_row[j] == 1:
                    pass
                else:
                    h_conf_row[j] += '-' * (h_clen_row[j] - 1) + ':'
            elif h_alig_tbl[std_row][j][0] == 'C':
                if h_clen_row[j] == 0:
                    pass
                elif h_clen_row[j] == 1:
                    h_conf_row[j] += ':'
                else:
                    h_conf_row[j] += ':' + '-' * (h_clen_row[j] - 2) + ':'
            else:  # "" or "L"
                if h_clen_row[j] == 0:
                    h_conf_row[j] += ''
                elif h_clen_row[j] == 1:
                    h_conf_row[j] += '-'
                else:
                    h_conf_row[j] += ':' + '-' * (h_clen_row[j] - 1)
            # RULE
            h_conf_row[j] += v_rule_tbl[std_row][j]
        return v_conf_clm, h_conf_row

    @staticmethod
    def __get_md_text(tbl_alig, txt_tbl, merge_tbl,
                      std_row, std_clm,
                      v_alig_tbl, h_alig_tbl, v_rule_tbl, h_rule_tbl,
                      v_conf_clm, h_conf_row):
        md_text = ''
        is_in_head = True
        for i, txt_row in enumerate(txt_tbl):
            # CONFIGURATION
            if is_in_head:
                if i == std_row:
                    if tbl_alig == 'left':
                        md_text += ': '
                    for j, _ in enumerate(txt_row):
                        md_text += '|' + h_conf_row[j]
                    md_text += '|'
                    if tbl_alig == 'right':
                        md_text += ' :'
                    md_text += '\n'
                    is_in_head = False
            # DATA
            for j, txt_cel in enumerate(txt_row):
                if len(txt_cel) == 0:
                    md_text += '|'
                else:
                    for k, txt_par in enumerate(txt_cel):
                        txt_par = re.sub('\n', '<br>', txt_par)
                        if k == 0:
                            md_text += '|'
                        else:
                            md_text += '<Br>'
                        if h_alig_tbl[i][j][k] == 'R':
                            if is_in_head:
                                md_text += txt_par + ' :'
                            elif merge_tbl[i][j] != '':
                                md_text += txt_par + ' :'
                            elif h_alig_tbl[std_row][j][0] != 'R':
                                md_text += txt_par + ' :'
                            else:
                                md_text += txt_par
                        elif h_alig_tbl[i][j][k] == 'C':
                            if is_in_head:
                                md_text += txt_par
                            elif merge_tbl[i][j] != '':
                                md_text += ': ' + txt_par + ' :'
                            elif h_alig_tbl[std_row][j][0] != 'C':
                                md_text += ': ' + txt_par + ' :'
                            else:
                                md_text += txt_par
                        else:  # "" or "L"
                            if is_in_head:
                                md_text += ': ' + txt_par
                            elif merge_tbl[i][j] != '':
                                md_text += ': ' + txt_par
                            elif (h_alig_tbl[std_row][j][0] != '' and
                                  h_alig_tbl[std_row][j][0] != 'L'):
                                md_text += ': ' + txt_par
                            else:
                                md_text += txt_par
                # MERGE CELLS
                if merge_tbl[i][j] != '':
                    if not re.match('^(.|\n)*\\s:$', md_text):
                        md_text += ' '
                    md_text += merge_tbl[i][j]
            md_text += '|' + v_conf_clm[i] + '\n'
        # md_text = md_text.replace('&lt;', '<')
        # md_text = md_text.replace('&gt;', '>')
        md_text = re.sub('\n$', '', md_text)
        return md_text

    @staticmethod
    def __split_long_lines(md_text):
        for line in md_text.split('\n'):
            if get_ideal_width(line) > MD_TEXT_WIDTH:
                break
        else:
            return md_text  # no long lines
        # LEFT ALIGNMENT
        is_left_alignment = False
        for line in md_text.split('\n'):
            if re.match('^:\\s*\\|(:?-*:?(^|=)?\\|)+:?-*:?(^|=)?$', line):
                is_left_alignment = True
        # SPLIT LINE
        new_text = ''
        for line in md_text.split('\n'):
            new_line = ''
            i = 0 if not is_left_alignment else 1
            for c in line:
                if c == '|' and re.match(NOT_ESCAPED + '\\|$', new_line + c):
                    if re.match('\\s$', new_line) and new_line != ': ':
                        new_line += '\\'
                    if new_line != '' and new_line != ': ':
                        new_line += '\n'
                    if new_line != ': ':
                        new_line += ' ' * (i * 2)
                    new_line += c
                    i += 1
                elif re.match(NOT_ESCAPED + '<(b|B)r>$', new_line + c):
                    new_line += c + '\n' + ' ' * (i * 2 - 1)
                else:
                    new_line += c
            new_text += new_line + '\n'
        md_text = new_text
        return md_text


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
                    + '[' + alte + ' @' + str(cm_w) + 'x' + str(cm_h) + ']' \
                    + '(' + path + ')'
            elif cm_w < 0:
                img_text = '!' \
                    + '[' + alte + ' @' + str(cm_w) + 'x' + ']' \
                    + '(' + path + ')'
            elif cm_h < 0:
                img_text = '!' \
                    + '[' + alte + ' @' + 'x' + str(cm_h) + ']' \
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

    def _get_text_to_write(self):
        ttw = super()._get_text_to_write()
        alignment = self.alignment
        if alignment == 'left':
            ttw = re.sub('^\\\\\\[', '\\\\[:', ttw)
        elif alignment == 'right':
            ttw = re.sub('\\\\\\]$', ':\\\\]', ttw)
        com = '\\\\(?:int|iint|iiint|oint|sum|prod)'
        ttw = MathDatum.shift_paren(com, 5, '_{.*}\\^{.*}{.*}', ttw)
        com = '\\\\(?:int|iint|iiint|oint|sum|prod)'
        ttw = MathDatum.shift_paren(com, 1, '{.*}', ttw)
        com = '\\\\(?:log|lim)'
        ttw = MathDatum.shift_paren(com, 3, '_{.*}{.*}', ttw)
        com = '\\\\(?:sin|cos|tan)'
        ttw = MathDatum.shift_paren(com, 3, '\\^{.*}{.*}', ttw)
        com = '\\\\(?:log|sin|cos|tan|exp|vec)'
        ttw = MathDatum.shift_paren(com, 1, '{.*}', ttw)
        ttw = MathDatum.cancel_multi_paren(ttw)
        ttw = re.sub('(\\\\begin{[^{}]+})', '\\n\\1\n', ttw)
        ttw = re.sub('(\\\\end{[^{}]+})', '\\n\\1\n', ttw)
        ttw = ttw.replace('\\\\', '\\\\\n')
        ttw = re.sub('^(\\\\\\[:?)', '\\1\n', ttw)
        ttw = re.sub('(:?\\\\\\])$', '\n\\1', ttw)
        ttw = re.sub('\n+', '\n', ttw)
        text_to_write = ttw
        return text_to_write


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
        if rp_alg != '':
            return True
        return False

    def _get_md_text(self, raw_text):
        alignment = self.alignment
        head_space = self.head_space
        tail_space = self.tail_space
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
        if head_space != '':
            if alignment == 'left' or alignment == 'center':
                md_text = re.sub('^: ', ': \\' + head_space, md_text)
                self.head_space = ''
        if tail_space != '':
            if alignment == 'center' or alignment == 'right':
                md_text = re.sub(' :$', tail_space + '\\ :', md_text)
                self.tail_space = ''
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
        if rp_sty is not None and rp_sty.style_id == 'makdo-g':
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
        rp = raw_paragraph
        if rp.horizontal_line != '':
            return True
        return False

    def _get_length_docx(self):
        if self.horizontal_line == 'textbox':
            return super()._get_length_docx()
        f_size = Form.font_size
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
        # ((2.14 - 1) * 0.75 * 12) - (2.14 * 12 * 0.5) = -2.58
        tmp_sb = tmp_sb - ((lnsp - 1) * 0.75 * f_size) - 2.580 - 0.0001
        # ((2.14 - 1) * 0.25 * 12) - (2.14 * 12 * 0.2) = -1.716
        tmp_sa = tmp_sa - ((lnsp - 1) * 0.25 * f_size) - 1.716 - 0.0001
        tmp_sb = tmp_sb / lnsp / f_size
        tmp_sa = tmp_sa / lnsp / f_size
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
        length_docx['first indent'] = round((fi_xml - hi_xml) / 20 / f_size, 2)
        length_docx['left indent'] = round((li_xml + ti_xml) / 20 / f_size, 2)
        length_docx['right indent'] = round(ri_xml / 20 / f_size, 2)
        # length_docx = self.length_docx
        return length_docx

    def _get_text_to_write_with_reviser(self):
        xml_lines = self.xml_lines
        tmp_ttw = self.text_to_write
        self.text_to_write = '----------------'
        ttwwr = super()._get_text_to_write_with_reviser()
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


class ParagraphMultiColumns(Paragraph):

    """A class to handle multicolumns paragraph"""

    paragraph_class = 'multicolumns'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_text = rp.raw_text
        rp_xl = rp.xml_lines
        if rp_text != '':
            return False
        for xl in rp_xl:
            if re.match('^<w:cols( .*)?/?>$', xl):
                return True
        return False

    def _get_md_text(self, raw_text):
        xml_lines = self.xml_lines
        num = 0
        wid = []
        for xl in xml_lines:
            if re.match('^<w:cols( .*)?>$', xl):
                num = XML.get_value('w:cols', 'w:num', num, xl)
                wid = []
            if re.match('^<w:col( .*)?>$', xl):
                w = XML.get_value('w:col', 'w:w', -1, xl)
                wid.append(w)
        if num == 0:
            md_text = '|-|'
        else:
            m = 1
            if len(wid) > 0:
                m = min(wid)
            while len(wid) < num:
                wid.append(m)
            md_text = '|'
            for w in wid:
                md_text += ('-' * round(w / m)) + '|'
        return md_text


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
        if rp_sty is not None and rp_sty.style_id == 'makdo-r':
            return True
        return False

    def _get_text_to_write_with_reviser(self):
        md_lines_text = self.md_lines_text
        ttwwr = md_lines_text
        ttwwr = re.sub('^●', '"" ', ttwwr)
        ttwwr = re.sub('\n●', '\n"" ', ttwwr)
        text_to_write_with_reviser = ttwwr
        return text_to_write_with_reviser


class ParagraphFootnotes(Paragraph):

    """A class to handle footnotes paragraph"""

    paragraph_class = 'footnotes'

    @classmethod
    def is_this_class(cls, raw_paragraph):
        rp = raw_paragraph
        rp_sty = rp.style
        if rp_sty is not None and rp_sty.style_id == 'makdo-f':
            return True
        return False

    def _get_text_to_write_with_reviser(self):
        md_lines_text = self.md_lines_text
        ttwwr = md_lines_text
        ttwwr = re.sub('^([0-9]+)）', '[^\\1]: ', ttwwr)
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
        numbering_xml_lines = io.read_xml_file('/word/numbering.xml')
        footnotes_xml_lines = io.read_xml_file('/word/footnotes.xml')
        # IMAGE LIST
        Form.rels = Form.get_rels(rels_xml_lines)
        # REMARKS
        Form.remarks = Form.get_remarks(comments_xml_lines)
        # STYLE LIST
        Form.styles = Form.get_styles(styles_xml_lines)
        # AUTO NUMBERING STYLE
        Form.auto_numbering_styles \
            = Form.get_auto_numbering_styles(numbering_xml_lines)
        # FOOTNOTES
        Form.footnotes = Form.get_footnotes(footnotes_xml_lines)
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
        frm.numbering_xml_lines = numbering_xml_lines
        frm.footnotes_xml_lines = footnotes_xml_lines
        frm.args = args
        frm.configure()
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
