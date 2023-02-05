#!/usr/bin/python3
# Name:         md2docx.py
# Version:      v04 Mitaki
# Time-stamp:   <2023.02.06-06:20:59-JST>

# md2docx.py
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


############################################################
# SETTING


import os
import sys
import argparse
import re
import unicodedata
import datetime
import docx
import chardet
from docx.shared import Cm, Pt
from docx.enum.text import WD_LINE_SPACING
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement, ns
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import socket   # host
import getpass  # user


__version__ = 'v04 Mitaki'


def get_arguments():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description='MarkdownファイルからMS Wordファイルを作ります',
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
        'md_file',
        help='Markdownファイル')
    parser.add_argument(
        'docx_file',
        default='',
        nargs='?',
        help='MS Wordファイル')
    return parser.parse_args()


def floats6(s):
    if not re.match('^' + RES_NUMBER6 + '$', s):
        raise argparse.ArgumentTypeError
    return s


HELP_EPILOG = '''Markdownの記法:
  行頭指示
    [#+=(数字) ]でセクション番号を変えることができます（独自）
    [v=(数字) ]で段落の上の余白を行数だけ増減します（独自）
    [V=(数字) ]で段落の下の余白を行数だけ増減します（独自）
    [X=(数字) ]で段落の改行幅を行数だけ増減します（独自）
    [<<=(数字) ]で段落1行目の左の余白を文字数だけ増減します（独自）
    [<=(数字) ]で段落の左の余白を文字数だけ増減します（独自）
    [>=(数字) ]で段落の右の余白を文字数だけ増減します（独自）
  行中指示
    [;;]から行末まではコメントアウトされます（独自）
    [<>]は何もせず表示もされません（独自）
    [<br>]で改行されます
    [<pgbr>]で改行されます（独自）
  文字装飾
    [*]で挟まれた文字列は斜体になります
    [**]で挟まれた文字列は太字になります
    [***]で挟まれた文字列は斜体かつ太字になります
    [~~]で挟まれた文字列は打消線が引かれます
    [`]で挟まれた文字列はゴシック体になります
    [//]で挟まれた文字列は斜体になります（独自）
    [--]で挟まれた文字列は文字が小さくなります（独自）
    [++]で挟まれた文字列は文字が大きくなります（独自）
    [^^]で挟まれた文字列は白色になって見えなくなります（独自）
    [^XXYYZZ^]で挟まれた文字列はRGB(XX,YY,ZZ)色になります（独自）
    [^foo^]で挟まれた文字列はfoo色になります（独自）
    [__]で挟まれた文字列は下線が引かれます（独自）
    [_foo_]で挟まれた区間の背景はfoo色になります（独自）
      red(R) darkRed(DR) yellow(Y) darkYellow(DY) green(G) darkGreen(DG)
      cyan(C) darkCyan(DC) blue(B) darkBlue(DB) magenta(M) darkMagenta(DM)
      lightGray(G1) darkGray(G2) black(BK)
  エスケープ記号
    [\\]をコマンドの前に書くとコマンドが文字列になります
    [\\\\]で"\\"が表示されます
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

FONT_COLOR = {
    'red':         'FF0000',
    'R':           'FF0000',
    'darkRed':     '7F0000',
    'DR':          '7F0000',
    'yellow':      'FFFF00',
    'Y':           'FFFF00',
    'darkYellow':  '7F7F00',
    'DY':          '7F7F00',
    'green':       '00FF00',
    'G':           '00FF00',
    'darkGreen':   '007F00',
    'DG':          '007F00',
    'cyan':        '00FFFF',
    'C':           '00FFFF',
    'darkCyan':    '007F7F',
    'DC':          '007F7F',
    'blue':        '0000FF',
    'B':           '0000FF',
    'darkBlue':    '00007F',
    'DB':          '00007F',
    'magenta':     'FF00FF',
    'M':           'FF00FF',
    'darkMagenta': '7F007F',
    'DM':          '7F007F',
    'lightGray':   'BFBFBF',
    'G1':          'BFBFBF',
    'darkGray':    '7F7F7F',
    'G2':          '7F7F7F',
    'black':       '000000',
    'BK':          '000000',
    'a000': 'FF5D5D',
    'a010': 'FF603C',
    'a020': 'FF6512',
    'a030': 'E07000',
    'a040': 'BC7A00',
    'a050': 'A08300',
    'a060': '898900',
    'a070': '758F00',
    'a080': '619500',
    'a090': '4E9B00',
    'a100': '38A200',
    'a110': '1FA900',
    'a120': '00B200',
    'a130': '00AF20',
    'a140': '00AC3C',
    'a150': '00AA55',
    'a160': '00A76D',
    'a170': '00A586',
    'a180': '00A2A2',
    'a190': '009FC3',
    'a200': '009AED',
    'a210': '1F8FFF',
    'a220': '4385FF',
    'a230': '5F7CFF',
    'a240': '7676FF',
    'a250': '8A70FF',
    'a260': '9E6AFF',
    'a270': 'B164FF',
    'a280': 'C75DFF',
    'a290': 'E056FF',
    'a300': 'FF4DFF',
    'a310': 'FF50DF',
    'a320': 'FF53C3',
    'a330': 'FF55AA',
    'a340': 'FF5892',
    'a350': 'FF5A79',
}

HIGHLIGHT_COLOR = {
    'red':         WD_COLOR_INDEX.RED,
    'R':           WD_COLOR_INDEX.RED,
    'darkRed':     WD_COLOR_INDEX.DARK_RED,
    'DR':          WD_COLOR_INDEX.DARK_RED,
    'yellow':      WD_COLOR_INDEX.YELLOW,
    'Y':           WD_COLOR_INDEX.YELLOW,
    'darkYellow':  WD_COLOR_INDEX.DARK_YELLOW,
    'DY':          WD_COLOR_INDEX.DARK_YELLOW,
    'green':       WD_COLOR_INDEX.BRIGHT_GREEN,
    'G':           WD_COLOR_INDEX.BRIGHT_GREEN,
    'darkGreen':   WD_COLOR_INDEX.GREEN,
    'DG':          WD_COLOR_INDEX.GREEN,
    'cyan':        WD_COLOR_INDEX.TURQUOISE,
    'C':           WD_COLOR_INDEX.TURQUOISE,
    'darkCyan':    WD_COLOR_INDEX.TEAL,
    'DC':          WD_COLOR_INDEX.TEAL,
    'blue':        WD_COLOR_INDEX.BLUE,
    'B':           WD_COLOR_INDEX.BLUE,
    'darkBlue':    WD_COLOR_INDEX.DARK_BLUE,
    'DB':          WD_COLOR_INDEX.DARK_BLUE,
    'magenta':     WD_COLOR_INDEX.PINK,
    'M':           WD_COLOR_INDEX.PINK,
    'darkMagenta': WD_COLOR_INDEX.VIOLET,
    'DM':          WD_COLOR_INDEX.VIOLET,
    'lightGray':   WD_COLOR_INDEX.GRAY_25,
    'G1':          WD_COLOR_INDEX.GRAY_25,
    'darkGray':    WD_COLOR_INDEX.GRAY_50,
    'G2':          WD_COLOR_INDEX.GRAY_50,
    'black':       WD_COLOR_INDEX.BLACK,
    'BK':          WD_COLOR_INDEX.BLACK,
}

DEFAULT_AUTO_SPACE = False

DEFAULT_LINE_SPACING = 2.14  # (2.0980+2.1812)/2=2.1396

DEFAULT_SPACE_BEFORE = ''
DEFAULT_SPACE_AFTER = ''

ZENKAKU_SPACE = chr(12288)

RES_NUMBER = '([-\\+]?(([0-9]+(\\.[0-9]+)?)|(\\.[0-9]+)))'
RES_NUMBER6 = '(' + RES_NUMBER + '?,){,5}' + RES_NUMBER + '?,?'

RELAX_SYMBOL = '<>'
ORIGINAL_COMMENT_SYMBOL = ';;'
COMMENT_SEPARATE_SYMBOL = ' / '

NOT_ESCAPED = '^((?:(?:.*\n)*.*[^\\\\])?(?:\\\\\\\\)*)?'

HORIZONTAL_BAR = '[ー−—－―‐]'


class ParagraphChapter:

    """A class to handle chapter paragraph"""

    res_par = '^(\\$+)((?:-\\$+)*)\\s*(.*)$'
    res_ins = '^(\\$+)((?:-\\$+)*)=\\s*([0-9]+)$'
    states = [[0, 0, 0, 0, 0],  # 第１編
              [0, 0, 0, 0, 0],  # 第１章
              [0, 0, 0, 0, 0],  # 第１節
              [0, 0, 0, 0, 0],  # 第１款
              [0, 0, 0, 0, 0]]  # 第１目
    post_char = ['編', '章', '節', '款', '目']

    @classmethod
    def is_this_class(cls, full_text):
        if re.match(cls.res_par, full_text):
            return True
        else:
            return False

    @classmethod
    def get_depth(cls, md_text):
        if not re.match(cls.res_par, md_text):
            return 0
        head = re.sub(cls.res_par, '\\1', md_text)
        depth = len(head)
        return depth

    @classmethod
    def set_states(cls, ins):
        if not re.match(cls.res_ins, ins):
            return
        head = re.sub(cls.res_ins, '\\1', ins)
        bran = re.sub(cls.res_ins, '\\2', ins)
        stat = re.sub(cls.res_ins, '\\3', ins)
        xdepth = len(head) - 1
        height = len(bran.replace('$', ''))
        cls.check_xdepth_and_height(xdepth, height, ins)
        if height == 0:
            cls.states[xdepth][height] = int(stat) - 1
        else:
            cls.states[xdepth][height] = int(stat) - 2
        cls.reset_after(xdepth, height, ins)

    @classmethod
    def update_states(cls, md_text):
        if not re.match(cls.res_par, md_text):
            return
        head = re.sub(cls.res_par, '\\1', md_text)
        bran = re.sub(cls.res_par, '\\2', md_text)
        titl = re.sub(cls.res_par, '\\3', md_text)
        xdepth = len(head) - 1
        height = len(bran.replace('$', ''))
        cls.check_xdepth_and_height(xdepth, height, md_text)
        cls.states[xdepth][height] += 1
        cls.reset_after(xdepth, height, md_text)

    @classmethod
    def check_xdepth_and_height(cls, xdepth, height, md_text):
        if xdepth >= len(cls.states):
            msg = '※ 警告: ' \
                + 'チャプターの深さが上限を超えています' \
                + '\n  ' + md_text
            # msg = 'warning: ' \
            #     + 'chapter has too many branches' \
            #     + '\n  ' + md_text
            sys.stderr.write(msg + '\n\n')
        if height >= len(cls.states[xdepth]):
            msg = '※ 警告: ' \
                + 'チャプターの枝が上限を超えています' \
                + '\n  ' + md_text
            # msg = 'warning: ' \
            #     + 'chapter has too many branches' \
            #     + '\n  ' + md_text
            sys.stderr.write(msg + '\n\n')

    @classmethod
    def reset_after(cls, xdepth, height, md_text):
        for i in range(len(cls.states)):
            for j in range(len(cls.states[i])):
                if i < xdepth:
                    pass
                elif i == xdepth:
                    if j <= height:
                        if cls.states[i][j] == 0:
                            print(str(i) + '/' + str(j))
                            msg = '※ 警告: ' \
                                + 'チャプターの枝番が"0"を含んでいます' \
                                + '\n  ' + md_text
                            # msg = 'warning: ' \
                            #     + 'chapter has "0" branch' \
                            #     + '\n  ' + md_text
                            sys.stderr.write(msg + '\n\n')
                    else:
                        cls.states[i][j] = 0
                else:
                    cls.states[i][j] = 0

    @classmethod
    def get_docx_text(cls, md_text):
        if not re.match(cls.res_par, md_text):
            return
        head = re.sub(cls.res_par, '\\1', md_text)
        bran = re.sub(cls.res_par, '\\2', md_text)
        titl = re.sub(cls.res_par, '\\3', md_text)
        depth = len(head) - 1
        heigh = len(bran.replace('$', ''))
        docx_text = '第' + n_int(cls.states[depth][0]) + cls.post_char[depth]
        for j in range(1, heigh + 1):
            docx_text += 'の' + n_int(cls.states[depth][j] + 1)
        docx_text += ZENKAKU_SPACE + titl
        return docx_text

    @classmethod
    def modify_length(cls, depth, length_docx):
        length_docx['space before'] += 0.5
        length_docx['space after'] += 0.5
        if depth > 0:
            length_docx['left indent'] += depth - 1
        return length_docx


class ParagraphSection:

    """A class to handle section paragraph"""

    res_par = '^(#+)((?:-#+)*)\\s*(.*)$'
    res_ins = '^(#+)((?:-#+)*)=\\s*([0-9]+)$'
    states = [[0, 0, 0, 0, 0],  # -
              [0, 0, 0, 0, 0],  # 第１
              [0, 0, 0, 0, 0],  # １
              [0, 0, 0, 0, 0],  # (1)
              [0, 0, 0, 0, 0],  # ア
              [0, 0, 0, 0, 0],  # (ｱ)
              [0, 0, 0, 0, 0],  # ａ
              [0, 0, 0, 0, 0]]  # (a)

    @classmethod
    def is_this_class(cls, full_text):
        if re.match(cls.res_par, full_text):
            return True
        else:
            return False

    @classmethod
    def get_depths(cls, full_text):
        depth_first = -1
        depth = -1
        head = re.sub('[^\\-# ].*', '', full_text)
        for s in head.split(' '):
            if s == '':
                continue
            if re.match(cls.res_par, s):
                depth = len(re.sub('\\-.*', '', s))
                if depth_first == -1:
                    depth_first = depth
                if depth > 6:
                    msg = '※ 警告: ' \
                        + 'セクションの深さが上限を超えています\n' \
                        + '  s'
                    # msg = 'warning: ' \
                    #     + 'section symbol is too deep\n' \
                    #     + '  s'
                    sys.stderr.write(msg + '\n\n')
        return depth_first, depth

    @classmethod
    def set_states(cls, ins):
        if not re.match(cls.res_ins, ins):
            return
        head = re.sub(cls.res_ins, '\\1', ins)
        bran = re.sub(cls.res_ins, '\\2', ins)
        stat = re.sub(cls.res_ins, '\\3', ins)
        xdepth = len(head) - 1
        height = len(bran.replace('#', ''))
        cls.check_xdepth_and_height(xdepth, height, ins)
        if height == 0:
            cls.states[xdepth][height] = int(stat) - 1
        else:
            cls.states[xdepth][height] = int(stat) - 2
        cls.reset_after(xdepth, height, ins)

    @classmethod
    def update_states(cls, md_text):
        if md_text == '':
            return
        head = re.sub(cls.res_par, '\\1', md_text)
        bran = re.sub(cls.res_par, '\\2', md_text)
        titl = re.sub(cls.res_par, '\\3', md_text)
        xdepth = len(head) - 1
        height = len(bran.replace('#', ''))
        if height >= len(cls.states[xdepth]):
            msg = '※ 警告: ' \
                + 'セクションの枝番が上限を超えています' \
                + '\n  ' + md_text
            # msg = 'warning: ' \
            #     + 'section has too many branches' \
            #     + '\n  ' + md_text
            sys.stderr.write(msg + '\n\n')
        cls.check_xdepth_and_height(xdepth, height, md_text)
        cls.states[xdepth][height] += 1
        cls.reset_after(xdepth, height, md_text)

    @classmethod
    def check_xdepth_and_height(cls, xdepth, height, md_text):
        if xdepth >= len(cls.states):
            msg = '※ 警告: ' \
                + 'セクションの深さが上限を超えています' \
                + '\n  ' + md_text
            # msg = 'warning: ' \
            #     + 'seciton has too many branches' \
            #     + '\n  ' + md_text
            sys.stderr.write(msg + '\n\n')
        if height >= len(cls.states[xdepth]):
            msg = '※ 警告: ' \
                + 'セクションの枝が上限を超えています' \
                + '\n  ' + md_text
            # msg = 'warning: ' \
            #     + 'section has too many branches' \
            #     + '\n  ' + md_text
            sys.stderr.write(msg + '\n\n')

    @classmethod
    def reset_after(cls, xdepth, height, md_text):
        for i in range(len(cls.states)):
            for j in range(len(cls.states[i])):
                if i < xdepth:
                    pass
                elif i == xdepth:
                    if j <= height:
                        if cls.states[i][j] == 0:
                            print(str(i) + '/' + str(j))
                            msg = '※ 警告: ' \
                                + 'セクションの枝番が"0"を含んでいます' \
                                + '\n  ' + md_text
                            # msg = 'warning: ' \
                            #     + 'section has "0" branch' \
                            #     + '\n  ' + md_text
                            sys.stderr.write(msg + '\n\n')
                    else:
                        cls.states[i][j] = 0
                else:
                    cls.states[i][j] = 0

    @classmethod
    def get_head_string(cls, md_text):
        if not re.match(cls.res_par, md_text):
            return ''
        head = re.sub(cls.res_par, '\\1', md_text)
        bran = re.sub(cls.res_par, '\\2', md_text)
        titl = re.sub(cls.res_par, '\\3', md_text)
        xdepth = len(head) - 1
        height = len(bran.replace('#', ''))
        head_string = ''
        if xdepth == 0:
            head_string = cls.get_head_1(cls.states[0][0])
        elif xdepth == 1:
            if doc.document_style == 'n':
                head_string = cls.get_head_2(cls.states[1][0])
            else:
                head_string = cls.get_head_2_j_or_J(cls.states[1][0])
        elif xdepth == 2:
            if doc.document_style != 'j' or cls.states[1][0] == 0:
                head_string = cls.get_head_3(cls.states[2][0])
            else:
                head_string = cls.get_head_3(cls.states[2][0] + 1)
        elif xdepth == 3:
            head_string = cls.get_head_4(cls.states[3][0])
        elif xdepth == 4:
            head_string = cls.get_head_5(cls.states[4][0])
        elif xdepth == 5:
            head_string = cls.get_head_6(cls.states[5][0])
        elif xdepth == 6:
            head_string = cls.get_head_7(cls.states[6][0])
        elif xdepth == 7:
            head_string = cls.get_head_8(cls.states[7][0])
        for j in range(1, height + 1):
            head_string += 'の' + n_int(cls.states[xdepth][j] + 1)
        return head_string

    @classmethod
    def get_head_space(cls, depth):
        n = cls.states[depth - 1][0]
        if depth == 1:
            return ''
        elif depth == 4 and ((n == 0) or (n > 20)):
            return ' '
        elif depth == 6:
            return ' '
        elif depth == 8:
            return ' '
        else:
            return ZENKAKU_SPACE

    @staticmethod
    def get_head_1(n):
        return ''

    @staticmethod
    def get_head_2(n):
        return '第' + n_int(n)

    @staticmethod
    def get_head_2_j_or_J(n):
        return '第' + n_int(n) + '条'

    @staticmethod
    def get_head_3(n):
        return n_int(n)

    @staticmethod
    def get_head_4(n):
        return n_paren_int(n)

    # @staticmethod
    # def get_head_4_J(n):
    #     return n_kanji(n)

    @staticmethod
    def get_head_5(n):
        return n_kata(n)

    @staticmethod
    def get_head_6(n):
        return n_paren_kata(n)

    @staticmethod
    def get_head_7(n):
        return n_alph(n)

    @staticmethod
    def get_head_8(n):
        return n_paren_alph(n)


class List:

    @staticmethod
    def get_bullet_head_1(n):
        return '• '  # U+2022 Bullet

    @staticmethod
    def get_bullet_head_2(n):
        return '◦ '  # U+25E6 White Bullet

    @staticmethod
    def get_bullet_head_3(n):
        return '‣ '  # U+2023 Triangular Bullet

    @staticmethod
    def get_bullet_head_4(n):
        return '⁃ '  # U+2043 Hyphen Bullet

    @staticmethod
    def get_number_head_1(n):
        return n_int(n) + '．'

    @staticmethod
    def get_number_head_2(n):
        return n_int(n) + '）'

    @staticmethod
    def get_number_head_3(n):
        return n_alph(n) + '．'

    @staticmethod
    def get_number_head_4(n):
        return n_alph(n) + '）'


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
        if p != '' and p != w:
            wid += 0.5
        p = w
    return wid


def n_int(n):
    if n < 10:
        return chr(65296 + n)
    else:
        return str(n)


def n_paren_int(n):
    if n == 0:
        return '(0)'
    elif n <= 20:
        return chr(9331 + n)
    else:
        return '(' + str(n) + ')'


def n_kanji(n):
    k = str(n)
    if n >= 100:
        k = re.sub('^(.+)(..)$', '\\1百\\2', k)
    if n >= 10:
        k = re.sub('^(.+)(.)$', '\\1十\\2', k)
    k = re.sub('0', '零', k)
    k = re.sub('1', '一', k)
    k = re.sub('2', '二', k)
    k = re.sub('3', '三', k)
    k = re.sub('4', '四', k)
    k = re.sub('5', '五', k)
    k = re.sub('6', '六', k)
    k = re.sub('7', '七', k)
    k = re.sub('8', '八', k)
    k = re.sub('9', '九', k)
    k = re.sub('(.+)零$', '\\1', k)
    k = re.sub('零十', '', k)
    k = re.sub('一十', '十', k)
    k = re.sub('一百', '百', k)
    return k


def n_kata(n):
    if n == 0:
        return chr(12448 + 83)
    elif n <= 5:
        return chr(12448 + (2 * n))
    elif n <= 17:
        return chr(12448 + (2 * n) - 1)
    elif n <= 20:
        return chr(12448 + (2 * n))
    elif n <= 25:
        return chr(12448 + (1 * n) + 21)
    elif n <= 30:
        return chr(12448 + (3 * n) - 31)
    elif n <= 35:
        return chr(12448 + (1 * n) + 31)
    elif n <= 38:
        return chr(12448 + (2 * n) - 4)
    elif n <= 43:
        return chr(12448 + (1 * n) + 34)
    elif n <= 45:
        return chr(12448 + (3 * n) - 53)
    elif n <= 46:
        return chr(12448 + (1 * n) + 37)
    else:
        msg = '※ 警告: ' \
            + 'カタカナ番号"' \
            + str(n) \
            + '"は上限46を超えています'
        # msg = 'warning: ' \
        #     + 'overflowed katakana "' + str(n) + '"'
        sys.stderr.write(msg + '\n\n')
        return '？'


def n_paren_kata(n):
    if n == 0:
        return '(' + chr(65392 + 45) + ')'
    elif n <= 44:
        return '(' + chr(65392 + n) + ')'
    elif n <= 45:
        return '(' + chr(65392 + n - 55) + ')'
    elif n <= 46:
        return '(' + chr(65392 + n - 1) + ')'
    else:
        msg = '※ 警告: ' \
            + '括弧付きカタカナ番号"' \
            + str(n) \
            + '"は上限46を超えています'
        # msg = 'warning: ' \
        #     + 'overflowed parenthesis katakata "' + str(n) + '"'
        sys.stderr.write(msg + '\n\n')
        return '(?)'


def n_alph(n):
    if n == 0:
        return chr(65344 + 26)
    elif n <= 26:
        return chr(65344 + n)
    else:
        msg = '※ 警告: ' \
            + 'アルファベット番号"' \
            + str(n) \
            + '"は上限26を超えています'
        # msg = 'warning: ' \
        #     + 'overflowed alphabet "' + str(n) + '"'
        sys.stderr.write(msg + '\n\n')
        return '？'


def n_paren_alph(n):
    if n == 0:
        return chr(9371 + 26)
    elif n <= 26:
        return chr(9371 + n)
    else:
        msg = '※ 警告: ' \
            + '括弧付きアルファベット番号"' \
            + str(n) \
            + '"は上限26を超えています'
        # msg = 'warning: ' \
        #     + 'overflowed parenthesis alphabet "' + str(n) + '"'
        sys.stderr.write(msg + '\n\n')
        return '(?)'


############################################################
# CLASS


class Document:

    """A class to handle document"""

    def __init__(self):
        self.md_file = ''
        self.docx_file = ''
        self.raw_md_lines = []
        self.md_lines = []
        self.raw_paragraphs = []
        self.paragraphs = []
        self.document_title = DEFAULT_DOCUMENT_TITLE
        self.document_style = DEFAULT_DOCUMENT_STYLE
        self.paper_size = DEFAULT_PAPER_SIZE
        self.top_margin = DEFAULT_TOP_MARGIN
        self.bottom_margin = DEFAULT_BOTTOM_MARGIN
        self.left_margin = DEFAULT_LEFT_MARGIN
        self.right_margin = DEFAULT_RIGHT_MARGIN
        self.header_string = DEFAULT_HEADER_STRING
        self.page_number = DEFAULT_PAGE_NUMBER
        self.line_number = DEFAULT_LINE_NUMBER
        self.mincho_font = DEFAULT_MINCHO_FONT
        self.gothic_font = DEFAULT_GOTHIC_FONT
        self.font_size = DEFAULT_FONT_SIZE
        self.line_spacing = DEFAULT_LINE_SPACING
        self.space_before = DEFAULT_SPACE_BEFORE
        self.space_after = DEFAULT_SPACE_AFTER
        self.auto_space = DEFAULT_AUTO_SPACE
        self.original_file = ''

    def get_raw_md_lines(self, md_file):
        self.md_file = md_file
        raw_md_lines = []
        try:
            if md_file == '-':
                bd = sys.stdin.buffer.read()
            else:
                bd = open(md_file, 'rb').read()
        except BaseException:
            msg = '※ エラー: ' \
                + '入力ファイル「' + md_file + '」を読み込めません'
            # msg = 'error: ' \
            #     + 'file "' + md_file + '" is not found'
            sys.stderr.write(msg + '\n\n')
            sys.exit(0)
        enc = chardet.detect(bd)['encoding']
        if enc is None:
            enc = 'SHIFT_JIS'
        elif (re.match('^utf[-_]?.*$', enc, re.I)) or \
             (re.match('^shift[-_]?jis.*$', enc, re.I)) or \
             (re.match('^cp932.*$', enc, re.I)) or \
             (re.match('^euc[-_]?(jp|jis).*$', enc, re.I)) or \
             (re.match('^iso[-_]?2022[-_]?jp.*$', enc, re.I)) or \
             (re.match('^ascii.*$', enc, re.I)):
            pass
        else:
            # Windows-1252 (Western Europe)
            # MacCyrillic (Macintosh Cyrillic)
            # ...
            msg = '※ 警告: ' \
                + '文字コード「' + enc + '」から「SHIFT_JIS」に修正しました'
            # msg = 'warning: ' \
            #     + 'detected encoding "' + enc + '" may be wrong'
            sys.stderr.write(msg + '\n\n')
            enc = 'SHIFT_JIS'
        try:
            sd = bd.decode(enc)
        except BaseException:
            msg = '※ エラー: ' \
                + '正しい文字コードを取得できませんでした' \
                + '（Markdownでない？）'
            # msg = 'error: ' \
            #     + 'could not detect correct character code '
            #     + '(not markdown?)'
            sys.stderr.write(msg + '\n\n')
            sys.exit(0)
        sd = re.sub('^' + chr(65279), '', sd)  # remove BOM / unnecessary?
        sd = re.sub('\r\n', '\n', sd)  # unnecessary?
        sd = re.sub('\r', '\n', sd)  # unnecessary?
        for rml in sd.split('\n'):
            rml = re.sub('  $', '\n', rml)
            rml = re.sub('[ ' + ZENKAKU_SPACE + '\t]*$', '', rml)
            raw_md_lines.append(rml)
        raw_md_lines.append('')
        # self.raw_md_lines = raw_md_lines
        return raw_md_lines

    def get_md_lines(self, raw_md_lines):
        md_lines = []
        is_in_comment = False
        for i, rml in enumerate(raw_md_lines):
            ml = MdLine(i + 1, rml)
            md_lines.append(ml)
        # self.md_lines = md_lines
        return md_lines

    def get_raw_paragraphs(self, md_lines):
        raw_paragraphs = []
        i = 0
        block = []
        for ml in md_lines:
            if ml.raw_text != '':
                block.append(ml)
                continue
            if len(block) == 0:
                continue
            if len(block) >= 2:
                if re.match('^\\`\\`\\`', block[0].raw_text):
                    if not re.match('^```', block[-1].raw_text):
                        block.append(ml)
                        continue
            p = Paragraph(i + 1, block)
            raw_paragraphs.append(p)
            i += 1
            block = []
        if len(block) > 0:
            p = Paragraph(i + 1, block)
            raw_paragraphs.append(p)
            i += 1
            block = []
        # self.raw_paragraphs = raw_paragraphs
        return raw_paragraphs

    def get_paragraphs(self, raw_paragraphs):
        paragraphs = []
        chapter_instructions = []
        section_instructions = []
        for rp in raw_paragraphs:
            if rp.paragraph_class == 'empty':
                if len(rp.chapter_instructions) > 0:
                    chapter_instructions += rp.chapter_instructions
                if len(rp.section_instructions) > 0:
                    section_instructions += rp.section_instructions
            else:
                paragraphs.append(rp)
                # CHAPTER INSTRUCTIONS
                if len(chapter_instructions) > 0:
                    paragraphs[-1].chapter_instructions += chapter_instructions
                    chapter_instructions = []
                # SECITON INSTRUCTIONS
                if len(section_instructions) > 0:
                    paragraphs[-1].section_instructions += section_instructions
                    section_instructions = []
        # self.paragraphs = paragraphs
        return paragraphs

    def modify_paragraphs(self, paragraphs):
        m = len(paragraphs) - 1
        for i, p in enumerate(paragraphs):
            if i > 0:
                p_prev = paragraphs[i - 1]
            if i < m:
                p_next = paragraphs[i + 1]
            if p.paragraph_class == 'title' and \
               p.section_depth == 1 and \
               re.match('^# .*\\S+.*$', p.full_text):
                if (p.length['space after'] >= 0.2) or \
                   (i < m and p_next.length['space before'] >= 0.2):
                    if i > 0:
                        p_prev.length['space after'] += 0.1
                    if True:
                        p.length['space before'] += 0.1
                    if p.length['space after'] >= 0.2:
                        p.length['space after'] -= 0.1
                    if i < m and p_next.length['space before'] >= 0.2:
                        p_next.length['space before'] -= 0.1
            if p.paragraph_class == 'table' or \
               p.paragraph_class == 'breakdown':
                sb = p.length['space before']
                if i > 0:
                    if p_prev.length['space after'] < sb:
                        p_prev.length['space after'] = sb
                sa = p.length['space after']
                if i < m:
                    if p_next.length['space before'] < sa:
                        p_next.length['space before'] = sa
            if i > 0 and \
               p_prev.paragraph_class == 'title' and \
               p_prev.section_depth == 1 and \
               re.match('^# *$', p_prev.full_text):
                if p.paragraph_class == 'title':
                    sb = (doc.space_before + ',,,,,').split(',')
                    df = self.section_depth_first
                    if sb[df - 1] != '':
                        p.length['space before'] -= float(sb[df - 1])
                p.length['space before'] += p_prev.length['space before']
                p.length['space before'] += p_prev.length['space after']
        return self.paragraphs

    def configure(self, md_lines, args):
        self._configure_by_md_file(md_lines)
        self._configure_by_args(args)
        Paragraph.mincho_font = self.mincho_font
        Paragraph.gothic_font = self.gothic_font
        Paragraph.font_size = self.font_size

    def _configure_by_md_file(self, md_lines):
        for ml in md_lines:
            com = ml.comment
            if com is None:
                break
            if re.match('^\\s*#', com):
                continue
            res = '^\\s*([^:：]+)[:：]\\s*(.*)$'
            if not re.match(res, com):
                continue
            nam = re.sub(res, '\\1', com).rstrip()
            val = re.sub(res, '\\2', com).rstrip()
            if False:
                pass
            elif nam == 'document_title' or nam == '書題名':
                self.document_title = val
            elif nam == 'document_style' or nam == '文書式':
                if val == 'n' or val == '普通' or val == '-':
                    self.document_style = 'n'
                elif val == 'k' or val == '契約':
                    self.document_style = 'k'
                elif val == 'j' or val == '条文':
                    self.document_style = 'j'
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は"普通"、"契約"又は"条文"で' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be "-", "k" or "j"'
                    sys.stderr.write(msg + '\n\n')
            elif nam == 'paper_size' or nam == '用紙サ':
                val = unicodedata.normalize('NFKC', val)
                if val == 'A3':
                    self.paper_size = 'A3'
                elif val == 'A3L' or val == 'A3横':
                    self.paper_size = 'A3L'
                elif val == 'A3P' or val == 'A3縦':
                    self.paper_size = 'A3P'
                elif val == 'A4':
                    self.paper_size = 'A4'
                elif val == 'A4L' or val == 'A4横':
                    self.paper_size = 'A4L'
                elif val == 'A4P' or val == 'A4縦':
                    self.paper_size = 'A4P'
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は' \
                        + '"A3横"、"A3縦"、"A4横"又は"A4縦"で' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be "A3", "A3P", "A4" or "A4L"'
                    sys.stderr.write(msg + '\n\n')
            elif (re.match('^(top|bottom|left|right)_margin$', nam) or
                  re.match('^(上|下|左|右)余白$', nam)):
                val = unicodedata.normalize('NFKC', val)
                val = re.sub('\\s*cm$', '', val)
                if re.match('^' + RES_NUMBER + '$', val):
                    if nam == 'top_margin' or nam == '上余白':
                        self.top_margin = float(val)
                    elif nam == 'bottom_margin' or nam == '下余白':
                        self.bottom_margin = float(val)
                    elif nam == 'left_margin' or nam == '左余白':
                        self.left_margin = float(val)
                    elif nam == 'right_margin' or nam == '右余白':
                        self.right_margin = float(val)
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は整数又は小数で' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be an integer or a decimal'
                    sys.stderr.write(msg + '\n\n')
            elif nam == 'header_string' or nam == '頭書き':
                self.header_string = val
            elif nam == 'page_number' or nam == '頁番号':
                val = unicodedata.normalize('NFKC', val)
                if val == 'True' or val == '有':
                    self.page_number = DEFAULT_PAGE_NUMBER
                elif val == 'False' or val == '無' or val == '-':
                    self.page_number = ''
                else:
                    self.page_number = val
            elif nam == 'line_number' or nam == '行番号':
                val = unicodedata.normalize('NFKC', val)
                if val == 'True' or val == '有':
                    self.line_number = True
                elif val == 'False' or val == '無':
                    self.line_number = False
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は"有"又は"無"で' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be "True" or "False"'
                    sys.stderr.write(msg + '\n\n')
            elif nam == 'mincho_font' or nam == '明朝体':
                self.mincho_font = val
            elif nam == 'gothic_font' or nam == 'ゴシ体':
                self.gothic_font = val
            elif nam == 'font_size' or nam == '文字サ':
                val = unicodedata.normalize('NFKC', val)
                val = re.sub('\\s*pt$', '', val)
                if re.match('^' + RES_NUMBER + '$', val):
                    self.font_size = float(val)
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は整数又は小数で' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be an integer or a decimal'
                    sys.stderr.write(msg + '\n\n')
            elif nam == 'line_spacing' or nam == '行間高':
                val = unicodedata.normalize('NFKC', val)
                val = re.sub('\\s*倍$', '', val)
                if re.match('^' + RES_NUMBER + '$', val):
                    self.line_spacing = float(val)
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は整数又は小数で' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be an integer or a decimal'
                    sys.stderr.write(msg + '\n\n')
            elif (re.match('^space_(before|after)$', nam) or
                  re.match('^(前|後)余白$', nam)):
                val = unicodedata.normalize('NFKC', val)
                val = val.replace('、', ',')
                val = val.replace('倍', '')
                val = val.replace(' ', '')
                if re.match('^' + RES_NUMBER6 + '$', val):
                    if nam == 'space_before' or nam == '前余白':
                        self.space_before = val
                    elif nam == 'space_after'or nam == '後余白':
                        self.space_after = val
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は' \
                        + '整数又は小数をカンマで区切って並べたもので' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be 6 integers or decimals'
                    sys.stderr.write(msg + '\n\n')
            elif nam == 'auto_space' or nam == '字間整':
                val = unicodedata.normalize('NFKC', val)
                if val == 'True' or val == '有':
                    self.auto_space = True
                elif val == 'False' or val == '無':
                    self.auto_space = False
                else:
                    msg = '※ 警告: ' \
                        + '「' + nam + '」の値は"有"又は"無"で' \
                        + 'なければなりません'
                    # msg = 'warning: ' \
                    #     + '"' + nam + '" must be "True" or "False"'
                    sys.stderr.write(msg + '\n\n')
            elif nam == 'original_file' or nam == '元原稿':
                self.original_file = val
            else:
                msg = '※ 警告: ' \
                    + '「' + nam + '」という設定項目は存在しません'
                # msg = 'warning: ' \
                #     + 'configuration name "' + nam + '" does not exist'
                sys.stderr.write(msg + '\n\n')

    def _configure_by_args(self, args):
        if args.document_title is not None:
            self.document_title = args.document_title
        if args.paper_size is not None:
            self.paper_size = args.paper_size
        if args.top_margin is not None:
            self.top_margin = args.top_margin
        if args.bottom_margin is not None:
            self.bottom_margin = args.bottom_margin
        if args.left_margin is not None:
            self.left_margin = args.left_margin
        if args.right_margin is not None:
            self.right_margin = args.right_margin
        if args.mincho_font is not None:
            self.mincho_font = args.mincho_font
        if args.gothic_font is not None:
            self.gothic_font = args.gothic_font
        if args.font_size is not None:
            self.font_size = args.font_size
        if args.document_style is not None:
            self.document_style = args.document_style
        if args.header_string is not None:
            self.header_string = args.header_string
        if args.page_number is not None:
            self.page_number = args.page_number
        if args.line_number:
            self.line_number = True
        if args.line_spacing is not None:
            self.line_spacing = args.line_spacing
        if args.space_before is not None:
            self.space_before = args.space_before
        if args.space_after is not None:
            self.space_after = args.space_after
        if args.auto_space:
            self.auto_space = True

    def get_ms_doc(self):
        size = self.font_size
        ms_doc = docx.Document()
        ms_sec = ms_doc.sections[0]
        ms_sec.page_height = Cm(PAPER_HEIGHT[self.paper_size])
        ms_sec.page_width = Cm(PAPER_WIDTH[self.paper_size])
        ms_sec.top_margin = Cm(self.top_margin)
        ms_sec.bottom_margin = Cm(self.bottom_margin)
        ms_sec.left_margin = Cm(self.left_margin)
        ms_sec.right_margin = Cm(self.right_margin)
        ms_sec.header_distance = Cm(1.0)
        ms_sec.footer_distance = Cm(1.0)
        ms_doc.styles['Footer'].font.size = Pt(size)      # page number
        ms_doc.styles['Normal'].font.size = Pt(size / 2)  # line number
        ms_doc.styles['List Bullet'].font.size = Pt(size)
        ms_doc.styles['List Bullet 2'].font.size = Pt(size)
        ms_doc.styles['List Bullet 3'].font.size = Pt(size)
        ms_doc.styles['List Number'].font.size = Pt(size)
        ms_doc.styles['List Number 2'].font.size = Pt(size)
        ms_doc.styles['List Number 3'].font.size = Pt(size)
        # HEADER
        if self.header_string != '':
            hs = self.header_string
            ms_par = ms_doc.sections[0].header.paragraphs[0]
            ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if re.match('^: (.*) :$', hs):
                hs = re.sub('^: (.*) :', '\\1', hs)
                ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            elif re.match('^: (.*)$', hs):
                hs = re.sub('^: (.*)', '\\1', hs)
                ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            elif re.match('^(.*) :$', hs):
                hs = re.sub('(.*) :$', '\\1', hs)
                ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            tex = ''
            is_large = False
            is_small = False
            for c in hs + '\0':
                tex += c
                if re.match(NOT_ESCAPED + '\\-\\-$', tex) or \
                   re.match(NOT_ESCAPED + '\\+\\+$', tex) or \
                   re.match(NOT_ESCAPED + '\0$', tex):
                    ms_run = ms_par.add_run()
                    if is_small:
                        ms_run.font.size = Pt(self.font_size * 0.8)
                    elif is_large:
                        ms_run.font.size = Pt(self.font_size * 1.2)
                    else:
                        ms_run.font.size = Pt(self.font_size * 1.0)
                    oe = OxmlElement('w:t')
                    oe.set(ns.qn('xml:space'), 'preserve')
                    if re.match(NOT_ESCAPED + '\\-\\-$', tex):
                        oe.text = re.sub('\\-\\-$', '', tex)
                        is_small = not is_small
                        is_large = False
                    elif re.match(NOT_ESCAPED + '\\+\\+$', tex):
                        oe.text = re.sub('\\+\\+$', '', tex)
                        is_small = False
                        is_large = not is_large
                    elif re.match(NOT_ESCAPED + '\0$', tex):
                        oe.text = re.sub('\0$', '', tex)
                    tex = ''
                    ms_run._r.append(oe)
        # FOOTER
        if self.page_number != '':
            pn = self.page_number
            ms_par = ms_doc.sections[0].footer.paragraphs[0]
            ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if re.match('^: (.*) :$', pn):
                pn = re.sub('^: (.*) :', '\\1', pn)
                ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            elif re.match('^: (.*)$', pn):
                pn = re.sub('^: (.*)', '\\1', pn)
                ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            elif re.match('^(.*) :$', pn):
                pn = re.sub('(.*) :$', '\\1', pn)
                ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            tex = ''
            for c in pn + '\0':
                if c != 'n' and c != 'N' and c != '\0':
                    tex += c
                    continue
                if not re.match(NOT_ESCAPED + 'x', tex + 'x') and c != '\0':
                    tex = re.sub('\\\\$', '', tex) + c
                    continue
                # TEXT
                ms_run = ms_par.add_run()
                oe = OxmlElement('w:t')
                oe.set(ns.qn('xml:space'), 'preserve')
                oe.text = tex
                ms_run._r.append(oe)
                tex = ''
                # PAGE OR NUMPAGES
                if c == '\0':
                    continue
                pn = re.sub('^n', '', pn)
                ms_run = ms_par.add_run()
                oe = OxmlElement('w:fldChar')
                oe.set(ns.qn('w:fldCharType'), 'begin')
                ms_run._r.append(oe)
                oe = OxmlElement('w:instrText')
                oe.set(ns.qn('xml:space'), 'preserve')
                if c == 'n':
                    oe.text = 'PAGE'
                elif c == 'N':
                    oe.text = 'NUMPAGES'
                ms_run._r.append(oe)
                oe = OxmlElement('w:fldChar')
                oe.set(ns.qn('w:fldCharType'), 'end')
                ms_run._r.append(oe)
        # LINE NUMBER
        if self.line_number:
            ms_scp = ms_doc.sections[0]._sectPr
            oe = OxmlElement('w:lnNumType')
            oe.set(ns.qn('w:countBy'), '5')
            oe.set(ns.qn('w:restart'), 'newPage')
            oe.set(ns.qn('w:distance'), '567')  # 567≒20*72/2.54=1cm
            ms_scp.append(oe)
        self.make_styles(ms_doc)
        return ms_doc

    def make_styles(self, ms_doc):
        size = self.font_size
        line_spacing = self.line_spacing
        # NORMAL
        ms_doc.styles.add_style('makdo', WD_STYLE_TYPE.PARAGRAPH)
        ms_doc.styles['makdo'].font.name = self.mincho_font
        ms_doc.styles['makdo'].font.size = Pt(size)
        ms_doc.styles['makdo'].paragraph_format.line_spacing \
            = Pt(line_spacing * size)
        if not doc.auto_space:
            pPr = ms_doc.styles['makdo']._element.get_or_add_pPr()
            # KANJI<->ENGLISH
            oe = OxmlElement('w:autoSpaceDE')
            oe.set(ns.qn('w:val'), '0')
            pPr.append(oe)
            # KANJI<->NUMBER
            oe = OxmlElement('w:autoSpaceDN')
            oe.set(ns.qn('w:val'), '0')
            pPr.append(oe)
        # GOTHIC
        ms_doc.styles.add_style('makdo-g', WD_STYLE_TYPE.PARAGRAPH)
        ms_doc.styles['makdo-g'].font.name = self.gothic_font
        # ALIGNMENT
        ms_doc.styles.add_style('makdo-a', WD_STYLE_TYPE.PARAGRAPH)
        # SPACE
        sb = self.space_before.split(',')
        sa = self.space_after.split(',')
        for i in range(6):
            n = 'makdo-' + str(i + 1)
            ms_doc.styles.add_style(n, WD_STYLE_TYPE.PARAGRAPH)
            if len(sb) > i and sb[i] != '':
                ms_doc.styles[n].paragraph_format.space_before \
                    = Pt(float(sb[i]) * line_spacing * size)
            if len(sa) > i and sa[i] != '':
                ms_doc.styles[n].paragraph_format.space_after \
                    = Pt(float(sa[i]) * line_spacing * size)

    def write_property(self, ms_doc):
        host = socket.gethostname()
        if host is None:
            host = '-'
        hh = self._get_hash(host)
        user = getpass.getuser()
        if user is None:
            user = '='
        hu = self._get_hash(user)
        tt = self.document_title
        if self.document_style == 'n':
            ct = '（普通）'
        elif self.document_style == 'k':
            ct = '（契約）'
        elif self.document_style == 'j':
            ct = '（条文）'
        at = hu + '@' + hh + ' (makdo ' + __version__ + ')'
        dt = datetime.datetime.utcnow()
        # utc = datetime.timezone.utc
        # pt = datetime.datetime(1970, 1, 1, 0, 0, 0, tzinfo=utc)
        # TIMEZONE IS NOT SUPPORTED
        # jst = datetime.timezone(datetime.timedelta(hours=9))
        # dt = datetime.datetime.now(jst)
        # pt = datetime.datetime(1970, 1, 1, 9, 0, 0, tzinfo=jst)
        ms_cp = ms_doc.core_properties
        ms_cp.identifier \
            = 'makdo(' + __version__.split()[0] + ');' \
            + hu + '@' + hh + ';' \
            + dt.strftime('%Y-%m-%dT%H:%M:%SZ')
        ms_cp.title = tt               # タイトル
        # ms_cp.subject = ''           # 件名
        # ms_cp.keywords = ''          # タグ
        ms_cp.category = ct            # 分類項目
        # ms_cp.comments = ''          # コメント
        ms_cp.author = at              # 作成者
        # ms_cp.last_modified_by = ''  # 前回保存者
        # ms_cp.revision = 1           # 改訂番号
        # ms_cp.version = ''           # バージョン番号
        ms_cp.created = dt             # コンテンツの作成日時
        ms_cp.modified = dt            # 前回保存時
        # ms_cp.last_printed = pt      # 前回印刷日
        # ms_cp.content_status = ''    # 内容の状態
        # ms_cp.language = ''          # 言語

    @staticmethod
    def _get_hash(st):
        # ''  owicwvnu
        # '-' sojfooxd
        # '=' empzhdhk
        x = 9973
        b = 99999989
        m = 999999999989
        z = int(((4 ** 20) - 1) / (4 - 1) * 2)
        for c in st + ' 2022.05.07 07:31:03':
            x = (x * b + ord(c)) % m
            x = x ^ z
        hs = ''
        for i in range(8):
            hs += chr(x % 26 + 97)
            x = int(x / 26)
        return hs

    def write_document(self, ms_doc):
        for p in self.paragraphs:
            p.write_paragraph(ms_doc)

    def save_docx_file(self, ms_doc, docx_file, md_file):
        if docx_file == '':
            if re.match('^.*\\.md$', md_file):
                docx_file = re.sub('\\.md$', '.docx', md_file)
            else:
                docx_file = md_file + '.docx'
            self.docx_file = docx_file
        if os.path.exists(docx_file):
            if not os.access(docx_file, os.W_OK):
                msg = '※ エラー: ' \
                    + '出力ファイル「' + docx_file + '」に書き込み権限が' \
                    + 'ありません'
                # msg = 'error: ' \
                #     + 'overwriting a unwritable file "' + docx_file + '"'
                sys.stderr.write(msg + '\n\n')
                sys.exit(1)
            if os.path.getmtime(md_file) < os.path.getmtime(docx_file):
                msg = '※ エラー: ' \
                    + '出力ファイル「' + docx_file + '」の方が' \
                    + '入力ファイル「' + md_file + '」よりも新しいです'
                # msg = 'error: ' \
                #     + 'overwriting a newer file "' + docx_file + '"'
                sys.stderr.write(msg + '\n\n')
                sys.exit(1)
            if os.path.exists(docx_file + '~'):
                os.remove(docx_file + '~')
            os.rename(docx_file, docx_file + '~')
        try:
            ms_doc.save(docx_file)
        except BaseException:
            msg = '※ エラー: ' \
                + '出力ファイル「' + docx_file + '」の書込みに失敗しました'
            # msg = 'error: ' \
            #     + 'failed to write output file "' + docx_file + '"'
            sys.stderr.write(msg + '\n\n')
            sys.exit(1)

    def print_warning_messages(self):
        for p in self.paragraphs:
            p.print_warning_messages()


class Paragraph:

    """A class to handle paragraph"""

    mincho_font = None
    font_size = None
    is_preformatted = False
    is_large = False
    is_small = False
    is_italic = False
    is_bold = False
    has_strike = False
    has_underline = False
    font_color = ''
    highlight_color = None

    def __init__(self, paragraph_number, md_lines):
        self.paragraph_number = paragraph_number
        self.md_lines = md_lines
        self.full_text = ''
        self.paragraph_class = None
        self.decoration_instruction = ''
        self.chapter_instructions = []
        self.section_instructions = []
        self.section_states = []
        self.section_depth_first = 0
        self.section_depth = 0
        self.alignment = None
        self.length \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        self.length_ins \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        self.length_sec \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        self.chapter_instructions, \
            self.section_instructions, \
            self.decoration_instruction, \
            self.length_ins, \
            self.md_lines \
            = self.read_first_line_instructions()
        self.full_text = self.get_full_text()
        self.paragraph_class \
            = self.get_paragraph_class()
        self.section_depth_first, \
            self.section_depth \
            = self.get_section_depths()
        self.length_sec \
            = self.get_length_sec()
        self.length \
            = self.get_length()

    def read_first_line_instructions(self):
        chapter_instructions = []
        section_instructions = []
        decoration_instruction = ''
        length_ins \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        md_lines = self.md_lines
        res_cn = '^\\s*(\\$+(?:-\\$)*=\\s*[0-9]+)(.*)$'
        res_sn = '^\\s*(#+(?:-#)*=\\s*[0-9]+)(.*)$'
        res_de = ('^\\s*((?:'
                  + '(?:\\*{1,3})'             # italic, bold
                  + '|(?:~~)'                  # strikethrough
                  + '|(?:`)'                   # preformatted
                  + '|(?://)'                  # italic
                  + '|(?:\\-\\-)'              # small
                  + '|(?:\\+\\+)'              # large
                  + '|(?:\\^[0-9A-Za-z]*\\^)'  # font color
                  + '|(?:_[0-9A-Za-z]*_)'      # highlight color
                  + ')+)(.*)$')
        res_sb = '^\\s*v=\\s*' + RES_NUMBER + '(.*)$'
        res_sa = '^\\s*V=\\s*' + RES_NUMBER + '(.*)$'
        res_ls = '^\\s*X=\\s*' + RES_NUMBER + '(.*)$'
        res_fi = '^\\s*<<=\\s*' + RES_NUMBER + '(.*)$'
        res_li = '^\\s*<=\\s*' + RES_NUMBER + '(.*)$'
        res_ri = '^\\s*>=\\s*' + RES_NUMBER + '(.*)$'
        for ml in md_lines:
            # FOR BREAKDOWN
            if re.match('^-+::-*(::-+)?$', ml.text):
                break
            while True:
                if False:
                    pass
                elif re.match(res_cn, ml.text):
                    chap_ins = re.sub(res_cn, '\\1', ml.text)
                    ml.text = re.sub(res_cn, '\\2', ml.text)
                    chapter_instructions.append(chap_ins)
                elif re.match(res_sn, ml.text):
                    sect_ins = re.sub(res_cn, '\\1', ml.text)
                    ml.text = re.sub(res_sn, '\\2', ml.text)
                    section_instructions.append(sect_ins)
                elif re.match(res_de, ml.text):
                    deco = re.sub(res_de, '\\1', ml.text)
                    ml.text = re.sub(res_de, '\\2', ml.text)
                    decoration_instruction += deco
                elif re.match(res_sb, ml.text):
                    deci = re.sub(res_sb, '\\1', ml.text)
                    ml.text = re.sub(res_sb, '\\6', ml.text)
                    length_ins['space before'] += float(deci)
                elif re.match(res_sa, ml.text):
                    deci = re.sub(res_sa, '\\1', ml.text)
                    ml.text = re.sub(res_sa, '\\6', ml.text)
                    length_ins['space after'] += float(deci)
                elif re.match(res_ls, ml.text):
                    deci = re.sub(res_ls, '\\1', ml.text)
                    ml.text = re.sub(res_ls, '\\6', ml.text)
                    length_ins['line spacing'] += float(deci)
                elif re.match(res_fi, ml.text):
                    deci = re.sub(res_fi, '\\1', ml.text)
                    ml.text = re.sub(res_fi, '\\6', ml.text)
                    length_ins['first indent'] = -float(deci)
                elif re.match(res_li, ml.text):
                    deci = re.sub(res_li, '\\1', ml.text)
                    ml.text = re.sub(res_li, '\\6', ml.text)
                    length_ins['left indent'] = -float(deci)
                elif re.match(res_ri, ml.text):
                    deci = re.sub(res_ri, '\\1', ml.text)
                    ml.text = re.sub(res_ri, '\\6', ml.text)
                    length_ins['right indent'] = -float(deci)
                else:
                    break
            # ml_rawt = ml.raw_text
            # while True:
            #     if False:
            #         pass
            #     elif re.match(res_cn, ml_rawt) and re.match(res_cn, ml.text):
            #         chap_ins = re.sub(res_cn, '\\1', ml_text)
            #         ml_rawt = re.sub(res_cn, '\\2', ml_rawt)
            #         ml.text = re.sub(res_cn, '\\2', ml.text)
            #         chapter_instructions.append(chap_ins)
            #     elif re.match(res_sn, ml_rawt) and re.match(res_sn, ml.text):
            #         sect_ins = re.sub(res_cn, '\\1', ml_text)
            #         ml_rawt = re.sub(res_sn, '\\2', ml_rawt)
            #         ml.text = re.sub(res_sn, '\\2', ml.text)
            #         section_instructions.append(sect_ins)
            #     elif re.match(res_de, ml_rawt) and re.match(res_de, ml.text):
            #         deco = re.sub(res_de, '\\1', ml.text)
            #         ml_rawt = re.sub(res_de, '\\2', ml_rawt)
            #         ml.text = re.sub(res_de, '\\2', ml.text)
            #         decoration_instruction += deco
            #     elif re.match(res_sb, ml_rawt) and re.match(res_sb, ml.text):
            #         deci = re.sub(res_sb, '\\1', ml.text)
            #         ml_rawt = re.sub(res_sb, '\\6', ml_rawt)
            #         ml.text = re.sub(res_sb, '\\6', ml.text)
            #         length_ins['space before'] += float(deci)
            #     elif re.match(res_sa, ml_rawt) and re.match(res_sa, ml.text):
            #         deci = re.sub(res_sa, '\\1', ml.text)
            #         ml_rawt = re.sub(res_sa, '\\6', ml_rawt)
            #         ml.text = re.sub(res_sa, '\\6', ml.text)
            #         length_ins['space after'] += float(deci)
            #     elif re.match(res_ls, ml_rawt) and re.match(res_ls, ml.text):
            #         deci = re.sub(res_ls, '\\1', ml.text)
            #         ml_rawt = re.sub(res_ls, '\\6', ml_rawt)
            #         ml.text = re.sub(res_ls, '\\6', ml.text)
            #         length_ins['line spacing'] += float(deci)
            #     elif re.match(res_fi, ml_rawt) and re.match(res_fi, ml.text):
            #         deci = re.sub(res_fi, '\\1', ml.text)
            #         ml_rawt = re.sub(res_fi, '\\6', ml_rawt)
            #         ml.text = re.sub(res_fi, '\\6', ml.text)
            #         length_ins['first indent'] = -float(deci)
            #     elif re.match(res_li, ml_rawt) and re.match(res_li, ml.text):
            #         deci = re.sub(res_li, '\\1', ml.text)
            #         ml_rawt = re.sub(res_li, '\\6', ml_rawt)
            #         ml.text = re.sub(res_li, '\\6', ml.text)
            #         length_ins['left indent'] = -float(deci)
            #     elif re.match(res_ri, ml_rawt) and re.match(res_ri, ml.text):
            #         deci = re.sub(res_ri, '\\1', ml.text)
            #         ml_rawt = re.sub(res_ri, '\\6', ml_rawt)
            #         ml.text = re.sub(res_ri, '\\6', ml.text)
            #         length_ins['right indent'] = -float(deci)
            #     else:
            #        break
            if ml.text != '':
                break
        if length_ins['line spacing'] < 0:
            length_ins['space before'] -= length_ins['line spacing'] * .75
            length_ins['space after'] -= length_ins['line spacing'] * .25
        elif length_ins['line spacing'] > 0:
            if length_ins['space before'] > length_ins['line spacing'] * .75:
                length_ins['space before'] -= length_ins['line spacing'] * .75
            else:
                length_ins['space before'] = 0
            if length_ins['space after'] > length_ins['line spacing'] * .25:
                length_ins['space after'] -= length_ins['line spacing'] * .25
            else:
                length_ins['space after'] = 0
        # self.chapter_instructions = chapter_instructions
        # self.section_instructions = section_instructions
        # self.decoration_instruction = decoration_instruction
        # self.length_ins = length_ins
        # self.md_lines = md_lines
        return chapter_instructions, section_instructions, \
            decoration_instruction, length_ins, md_lines

    def get_full_text(self):
        full_text = ''
        for ml in self.md_lines:
            full_text += ml.text + ' '
        full_text = re.sub('\t', ' ', full_text)
        full_text = re.sub(' +', ' ', full_text)
        full_text = re.sub('^ ', '', full_text)
        full_text = re.sub(' $', '', full_text)
        # self.full_text = full_text
        return full_text

    def get_paragraph_class(self):
        decoration = self.decoration_instruction
        full_text = self.full_text
        paragraph_class = None
        if decoration + full_text == '':
            paragraph_class = 'empty'
        elif re.match('^\n$', decoration + full_text):
            paragraph_class = 'blank'
        elif ParagraphChapter.is_this_class(full_text):
            paragraph_class = 'chapter'
        elif ParagraphSection.is_this_class(full_text):
            paragraph_class = 'title'
        elif re.match(NOT_ESCAPED + '::', full_text):
            paragraph_class = 'breakdown'
        elif re.match('^ *([-\\+\\*]|([0-9]+\\.)) ', full_text):
            paragraph_class = 'list'
        elif (re.match('^: (.*\n)*.*$', full_text) or
              re.match('^(.*\n)*.* :$', full_text)):
            paragraph_class = 'alignment'
        elif re.match('^\\|.*\\|$', full_text):
            paragraph_class = 'table'
        elif re.match('^(\\s*' +
                      '(! ?\\[[^\\[\\]]*\\] ?\\([^\\(\\)]+\\)|\\+\\+|\\-\\-)' +
                      ')+\\s*$', full_text):
            paragraph_class = 'image'
        elif re.match('^```.*$', full_text):
            paragraph_class = 'preformatted'
        elif re.match('^<div style="break-.*: page;"></div>$', full_text):
            paragraph_class = 'pagebreak'
        elif re.match('^<pgbr/?>$', full_text):
            paragraph_class = 'pagebreak'
        else:
            paragraph_class = 'sentence'
        # self.paragraph_class = paragraph_class
        return paragraph_class

    def get_section_depths(self):
        depth_first = 0
        depth = 0
        for i, pss in enumerate(ParagraphSection.states):
            if pss[0] > 0:
                depth_first = i + 1
                depth = i + 1
        if self.paragraph_class == 'title':
            depth_first, depth = ParagraphSection.get_depths(self.full_text)
        # self.section_depth_first = depth_first
        # self.section_depth = depth
        return depth_first, depth

    def get_length_sec(self):
        length_sec \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        par_class = self.paragraph_class
        states = ParagraphSection.states
        depth_first = self.section_depth_first
        depth = self.section_depth
        if par_class == 'title':
            if depth_first > 1:
                length_sec['first indent'] = depth_first - depth - 1.0
            if depth_first > 1:
                length_sec['left indent'] = depth - 1.0
            if depth_first >= 3 and states[1] == 0:
                length_sec['left indent'] -= 1.0
        elif par_class == 'list' or par_class == 'breakdown':
            length_sec['first indent'] = 0
            if depth_first > 1:
                length_sec['left indent'] = depth - 1.0
            if depth_first >= 3 and states[1] == 0:
                length_sec['left indent'] -= 1.0
        elif par_class == 'sentence':
            if depth_first > 0:
                length_sec['first indent'] = 1.0
            if depth_first > 1:
                length_sec['left indent'] = depth - 1.0
            if depth_first >= 3 and states[1] == 0:
                length_sec['left indent'] -= 1.0
        # self.length_sec = length_sec
        return length_sec

    def get_length(self):
        length \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        for s in length:
            length[s] = self.length_ins[s] + self.length_sec[s]
        if self.paragraph_class == 'chapter':
            depth = ParagraphChapter.get_depth(self.full_text)
            length = ParagraphChapter.modify_length(depth, self.length_ins)
        if self.paragraph_class == 'title':
            sb = (doc.space_before + ',,,,,').split(',')
            sa = (doc.space_after + ',,,,,').split(',')
            df = self.section_depth_first
            dl = self.section_depth
            if sb[df - 1] != '':
                length['space before'] += float(sb[df - 1])
            if sa[dl - 1] != '':
                length['space after'] += float(sa[dl - 1])
        # self.length = length
        return length

    def write_paragraph(self, ms_doc):
        for ci in self.chapter_instructions:
            ParagraphChapter.set_states(ci)
        for si in self.section_instructions:
            ParagraphSection.set_states(si)
        if doc.document_style == 'j':
            if self.paragraph_class == 'title' or \
               self.paragraph_class == 'sentence':
                if ParagraphSection.states[1][0] > 0 and \
                   self.section_depth_first >= 3:
                    self.length['left indent'] -= 1
        paragraph_class = self.paragraph_class
        if paragraph_class == 'empty':
            self._write_empty_paragraph(ms_doc)
        elif paragraph_class == 'blank':
            self._write_blank_paragraph(ms_doc)
        elif paragraph_class == 'chapter':
            self._write_chapter_paragraph(ms_doc)
        elif paragraph_class == 'title':
            self._write_title_paragraph(ms_doc)
        elif paragraph_class == 'breakdown':
            self._write_breakdown_paragraph(ms_doc)
        elif paragraph_class == 'list':
            self._write_list_paragraph(ms_doc)
        elif paragraph_class == 'alignment':
            self._write_alignment_paragraph(ms_doc)
        elif paragraph_class == 'table':
            self._write_table_paragraph(ms_doc)
        elif paragraph_class == 'image':
            self._write_image_paragraph(ms_doc)
        elif paragraph_class == 'preformatted':
            self._write_preformatted_paragraph(ms_doc)
        elif paragraph_class == 'pagebreak':
            self._write_pagebreak_paragraph(ms_doc)
        else:
            self._write_sentence_paragraph(ms_doc)

    def _write_empty_paragraph(self, ms_doc):
        text_to_write = self.decoration_instruction
        for ml in self.md_lines:
            text_to_write += ml.text
        if text_to_write != '':
            ms_par = self._get_ms_par(ms_doc)
            self._write_text(text_to_write, ms_par)
            msg = '※ 警告: ' \
                + '空段落が「' + text_to_write + '」を' \
                + '含んでいます'
            # msg = 'warning: ' \
            #     + 'unexpected state (empty paragraph)' + '\n  ' \
            #     + text_to_write
            sys.stderr.write(msg + '\n\n')

    def _write_blank_paragraph(self, ms_doc):
        text_to_write = self.decoration_instruction
        for ml in self.md_lines:
            text_to_write += ml.text
        ms_par = self._get_ms_par(ms_doc)
        if text_to_write != '\n':
            self._write_text(text_to_write, ms_par)
            msg = '※ 警告: ' \
                + '改行段落が「' + re.sub('\n$', '', text_to_write) + '」を' \
                + '含んでいます'
            # msg = 'warning: ' \
            #     + 'unexpected state (blank paragraph)' + '\n  ' \
            #     + text_to_write
            sys.stderr.write(msg + '\n\n')

    def _write_chapter_paragraph(self, ms_doc):
        for i, ml in enumerate(self.md_lines):
            text_to_write = ''
            if i == 0:
                text_to_write += self.decoration_instruction
            ParagraphChapter.update_states(ml.text)
            text_to_write += ParagraphChapter.get_docx_text(ml.text)
            ms_par = self._get_ms_par(ms_doc)
            self._write_text(text_to_write, ms_par)

    def _write_title_paragraph(self, ms_doc):
        md_lines = self.md_lines
        size = self.font_size
        ll_size = size * 1.4
        depth = self.section_depth
        text_to_write = self.decoration_instruction
        head_symbol, title, text = self._split_title_paragraph(md_lines)
        head_string = ''
        for hs in head_symbol.split(' '):
            ParagraphSection.update_states(hs)
            head_string += ParagraphSection.get_head_string(hs)
        head_string += ParagraphSection.get_head_space(depth)
        if title + text == '':
            return
        ms_par = self._get_ms_par(ms_doc)
        ms_fmt = ms_par.paragraph_format
        if depth == 1:
            Paragraph.font_size = ll_size
            ms_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif title == '' and text != '':
            ms_par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        text_to_write += head_string + title + text
        self._write_text(text_to_write, ms_par)
        if depth == 1:
            Paragraph.font_size = size

    @staticmethod
    def _split_title_paragraph(md_lines):
        head = ''
        title = ''
        text = ''
        is_in_head = True
        res = '(#+(?:-#+)*(?:\\s+#+(?:-#+)*)*)'
        for ml in md_lines:
            if is_in_head:
                if re.match('^' + res + '\\s*$', ml.text):
                    head += ml.text
                    head += ' '
                elif re.match('^' + res + '\\s+(.*)$', ml.text):
                    head += re.sub('^' + res + '\\s+(.*)$', '\\1', ml.text)
                    head += ' '
                    title = re.sub('^' + res + '\\s+(.*)$', '\\2', ml.text)
                else:
                    is_in_head = False
                    text += ml.text
            else:
                text += ml.text
        return head, title, text

    @staticmethod
    def _is_consistent_with_depth(md_line, pre_dep, dep):
        if pre_dep > 0:
            if (pre_dep <= 2) or (pre_dep + 1 != dep):
                msg = '警告: ' \
                    + 'セクションの深さが「' + str(pre_dep) + '」から「' \
                    + str(dep) + '」に飛んでいます'
                # msg = 'warning: bad depth ' + str(pre_dep) \
                #     + ' to ' + str(dep)
                md_line.append_warning_message(msg)
                return False
        return True

    def _write_breakdown_paragraph(self, ms_doc):
        size = self.font_size
        bds = self._get_breakdown_data()
        conf_row, wid_list, hei_list \
            = self._get_breakdown_width_and_height(bds)
        if conf_row >= 0:
            bds.pop(conf_row)
        row = len(bds)
        ms_tab = ms_doc.add_table(row, 3, style='Normal Table')
        ind = self.length['left indent'] * size * 20
        oe = OxmlElement('w:tblInd')
        oe.set(qn('w:w'), str(ind))
        oe.set(qn('w:type'), 'dxa')
        tblpr = ms_tab._element.xpath('w:tblPr')
        tblpr[0].append(oe)
        for i in range(len(bds)):
            ms_tab.rows[i].height_rule = WD_ROW_HEIGHT_RULE.AUTO
            # ms_tab.rows[i].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            ms_tab.rows[i].height = Pt(doc.line_spacing * size * hei_list[i])
        for j in range(len(bds[0])):
            ms_tab.columns[j].width = Pt((wid_list[j] + 2) * size)
        # ms_tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i in range(len(bds)):
            for j in range(len(bds[i])):
                cell = bds[i][j]
                if j == 1:
                    cell = re.sub('^\\s+', '', cell)
                else:
                    cell = re.sub('\\s+$', '', cell)
                ms_cell = ms_tab.cell(i, j)
                ms_cell.width = Pt((wid_list[j] + 2) * size)
                ms_par = ms_cell.paragraphs[0]
                self._write_text(cell, ms_par)
                ms_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                ms_fmt = ms_par.paragraph_format
                ms_fmt.space_before = Pt(0)
                ms_fmt.space_after = Pt(0)
                ms_fmt.line_spacing = Pt(doc.line_spacing * size)
                if i < conf_row:
                    ms_fmt.alignment = WD_TABLE_ALIGNMENT.CENTER
                elif j == 1:
                    ms_fmt.alignment = WD_TABLE_ALIGNMENT.RIGHT
                else:
                    ms_fmt.alignment = WD_TABLE_ALIGNMENT.LEFT

    def _get_breakdown_data(self):
        bds = []
        list_states = 0
        for i, ml in enumerate(self.md_lines):
            if ml.text == '':
                continue
            bd = ml.text.split('::', 2)
            if not re.match('^-+::-*(::-+)?$', self.md_lines[0].text):
                if i == 0:
                    bd[0] = self.decoration_instruction + bd[0]
            else:
                if i == 1:
                    bd[0] = self.decoration_instruction + bd[0]
            while len(bd) < 3:
                bd.append('')
            res_b = '^ *[-\\+\\*] '
            res_n = '^ *[0-9]+\\. '
            if re.match(res_b, bd[0]):
                item = re.sub(res_b, '', bd[0])
                bd[0] = List.get_bullet_head_1(0) + item
            elif re.match(res_n, bd[0]):
                item = re.sub(res_n, '', bd[0])
                list_states += 1
                bd[0] = List.get_number_head_1(list_states) + item
            else:
                bd[0] = ZENKAKU_SPACE + bd[0]
            bds.append(bd)
        return bds

    def _get_breakdown_width_and_height(self, bds):
        # CONFIGURATION ROW
        conf_row = -1
        for i, ml in enumerate(self.md_lines):
            if re.match('^-+::-*(::-+)?$', ml.text):
                conf_row = i
                break
        # WIDTH
        wid_list = []
        if conf_row >= 0:
            for s in bds[conf_row]:
                wid_list.append(float(len(s)) / 2)
        else:
            wid_list = [0, 0, 0]
            for i, bd in enumerate(bds):
                for j, s in enumerate(bd):
                    s = re.sub('<br/>', '<br>', s)
                    lns = s.split('<br>')
                    for ln in lns:
                        w = float(get_real_width(ln)) / 2
                        if wid_list[j] < w:
                            wid_list[j] = w
        # HEIGHT
        hei_list = []
        for i, bd in enumerate(bds):
            if i == conf_row:
                continue
            h = 0
            for s in bd:
                s = re.sub('<br/>', '<br>', s)
                lns = s.split('<br>')
                if h < len(lns):
                    h = len(lns)
            hei_list.append(h)
        # RETURN
        return conf_row, wid_list, hei_list

    # def _write_breakdown_paragraph(self, ms_doc):
    #     size = self.font_size
    #     bds = []
    #     list_states = 0
    #     hei_list = []
    #     wid_list = [0, 0, 0]
    #     for ml in self.md_lines:
    #         if ml.text == '':
    #             continue
    #         bd = ml.text.split('::', 2)
    #         # for i, b in enumerate(bd):
    #         #     if (i % 2) == 0:
    #         #         bd[i] = re.sub('\\s+$', '', bd[i])
    #         #     else:
    #         #         bd[i] = re.sub('^\\s+', '', bd[i])
    #         res_b = '^ *[-\\+\\*] '
    #         res_n = '^ *[0-9]+\\. '
    #         if re.match(res_b, bd[0]):
    #             item = re.sub(res_b, '', bd[0])
    #             bd[0] = List.get_bullet_head_1(0) + item
    #         elif re.match(res_n, bd[0]):
    #             item = re.sub(res_n, '', bd[0])
    #             list_states += 1
    #             bd[0] = List.get_number_head_1(list_states) + item
    #         else:
    #             bd[0] = ZENKAKU_SPACE + bd[0]
    #         while len(bd) < 3:
    #             bd.append('')
    #         bds.append(bd)
    #         h = 0
    #         wl = [0, 0, 0]
    #         for i, b in enumerate(bd):
    #             bd[i] = re.sub('<br/>', '<br>', bd[i])
    #             lns = b.split('<br>')
    #             if h < len(lns):
    #                 h = len(lns)
    #             for ln in lns:
    #                 wl[i] = float(get_real_width(ln)) / 2
    #                 if wid_list[i] < wl[i]:
    #                     wid_list[i] = wl[i]
    #         hei_list.append(h)
    #     ms_tab = ms_doc.add_table(len(bds), 3, style='Normal Table')
    #     ind = self.length['left indent'] * size * 20
    #     oe = OxmlElement('w:tblInd')
    #     oe.set(qn('w:w'), str(ind))
    #     oe.set(qn('w:type'), 'dxa')
    #     tblpr = ms_tab._element.xpath('w:tblPr')
    #     tblpr[0].append(oe)
    #     for i in range(len(bds)):
    #         ms_tab.rows[i].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    #         ms_tab.rows[i].height = Pt(doc.line_spacing * size * hei_list[i])
    #     ms_tab.autofit = True
    #     for j in range(len(bds[0])):
    #         ms_tab.columns[j].width = Pt((wid_list[j] + 2) * size)
    #     for i in range(len(bds)):
    #         for j in range(len(bds[i])):
    #             ms_cell = ms_tab.cell(i, j)
    #             ms_cell.hight = ms_tab.rows[i].height
    #             ms_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    #             ms_cell.width = Pt((wid_list[j] + 2) * size)
    #             ms_par = ms_cell.paragraphs[0]
    #             self._write_text(bds[i][j] + '\n', ms_par)
    #             ms_fmt = ms_par.paragraph_format
    #             ms_fmt.space_before = Pt(0)
    #             ms_fmt.space_after = Pt(0)
    #             ms_fmt.line_spacing = Pt(doc.line_spacing * size)
    #             if j == 1:
    #                 ms_fmt.alignment = WD_TABLE_ALIGNMENT.RIGHT
    #             else:
    #                 ms_fmt.alignment = WD_TABLE_ALIGNMENT.LEFT

    def _write_list_paragraph(self, ms_doc):
        size = self.font_size
        list_states = [0, 0, 0, 0]
        list_depth = -1
        text_to_write = self.decoration_instruction
        for i, ml in enumerate(self.md_lines):
            text = ml.text
            text = re.sub('^\t', ' ' * 4, text)
            res = '^ *([-\\+\\*]|([0-9]+\\.)) '
            if not re.match(res, text):
                text = re.sub('^[ \t]*', '', text)
                text_to_write = self._join_string(text_to_write, text)
                continue
            text_to_write += '\n'
            list_depth = int(len(re.sub('\\S.*$', '', text)) / 2) + 1
            if re.match('^ *[0-9]+\\. ', text):
                list_states[list_depth - 1] += 1
                for dep in range(list_depth, 4):
                    list_states[list_depth] = 0
            n = list_states[list_depth - 1]
            item = re.sub(res, '', text)
            if re.match('^ *[-\\+\\*] ', text):
                if list_depth == 1:
                    text_to_write \
                        += ZENKAKU_SPACE * 0 + List.get_bullet_head_1(n) + item
                elif list_depth == 2:
                    text_to_write \
                        += ZENKAKU_SPACE * 2 + List.get_bullet_head_2(n) + item
                elif list_depth == 3:
                    text_to_write \
                        += ZENKAKU_SPACE * 4 + List.get_bullet_head_3(n) + item
                else:
                    text_to_write \
                        += ZENKAKU_SPACE * 6 + List.get_bullet_head_4(n) + item
            else:
                if list_depth == 1:
                    text_to_write \
                        += ZENKAKU_SPACE * 0 + List.get_number_head_1(n) + item
                elif list_depth == 2:
                    text_to_write \
                        += ZENKAKU_SPACE * 2 + List.get_number_head_2(n) + item
                elif list_depth == 3:
                    text_to_write \
                        += ZENKAKU_SPACE * 4 + List.get_number_head_3(n) + item
                else:
                    text_to_write \
                        += ZENKAKU_SPACE * 6 + number_list_head_4(n) + item
        text_to_write = re.sub('^\n*', '', text_to_write)
        text_to_write = re.sub('\n*$', '', text_to_write)
        ms_par = self._get_ms_par(ms_doc)
        self._write_text(text_to_write, ms_par)
        text_to_write = ''

    def _write_alignment_paragraph(self, ms_doc):
        size = self.font_size
        decoration = self.decoration_instruction
        indent = self.length['first indent'] + self.length['left indent']
        ms_par = self._get_ms_par(ms_doc, 'makdo-a')
        ms_fmt = ms_par.paragraph_format
        ms_fmt.first_line_indent = Pt(0 * size)
        oe = OxmlElement('w:wordWrap')
        oe.set(ns.qn('w:val'), '0')
        pPr = ms_par._p.get_or_add_pPr()
        pPr.append(oe)
        text_to_write = ''
        first_line = self.md_lines[-1].text
        if re.match('^: .* :$', first_line):
            alignment = 'center'
            for ml in self.md_lines:
                if ml.text == '':
                    continue
                text_to_write += '\n' + re.sub('^: (.*) :$', '\\1', ml.text)
                if indent > 0:
                    ms_fmt.left_indent = Pt(indent * 2 * size)
                elif indent < 0:
                    ms_fmt.right_indent = Pt(-indent * 2 * size)
            ms_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif re.match('^.* :$', first_line):
            alignment = 'right'
            for ml in self.md_lines:
                if ml.text == '':
                    continue
                text_to_write += '\n' + re.sub('^(.*) :$', '\\1', ml.text)
                ms_fmt.right_indent = Pt(-indent * size)
            ms_par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            alignment = 'left'
            for ml in self.md_lines:
                if ml.text == '':
                    continue
                text_to_write += '\n' + re.sub('^:(?: (.*))?$', '\\1', ml.text)
                # text_to_write += '\n' + re.sub('^: (.*)$', '\\1', ml.text)
                ms_fmt.left_indent = Pt(+indent * size)
            ms_par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        text_to_write = re.sub('^\n*', '', text_to_write)
        self._write_text(decoration + text_to_write, ms_par)

    def _write_table_paragraph(self, ms_doc):
        size = self.font_size
        s_size = 0.8 * size
        Paragraph.font_size = s_size
        tab = self._get_table_data()
        conf_row, ali_list, wid_list = self._get_table_alignment_and_width(tab)
        if conf_row >= 0:
            tab.pop(conf_row)
        row = len(tab)
        col = len(tab[0])
        ms_tab = ms_doc.add_table(row, col, style='Table Grid')
        # ms_tab.autofit = True
        for i in range(len(tab)):
            ms_tab.rows[i].height_rule = WD_ROW_HEIGHT_RULE.AUTO
        for j in range(len(tab[0])):
            ms_tab.columns[j].width = Pt((wid_list[j] + 2) * s_size)
        ms_tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i in range(len(tab)):
            # ms_tab.rows[i].height = Pt(1.5 * size)
            for j in range(len(tab[i])):
                cell = tab[i][j]
                if ali_list[j] != WD_TABLE_ALIGNMENT.LEFT:
                    cell = re.sub('^\\s+', '', cell)
                if ali_list[j] != WD_TABLE_ALIGNMENT.RIGHT:
                    cell = re.sub('\\s+$', '', cell)
                ms_cell = ms_tab.cell(i, j)
                ms_cell.width = Pt((wid_list[j] + 2) * s_size)
                ms_par = ms_cell.paragraphs[0]
                self._write_text(cell, ms_par)
                ms_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                ms_fmt = ms_par.paragraph_format
                if i < conf_row:
                    ms_fmt.alignment = WD_TABLE_ALIGNMENT.CENTER
                else:
                    ms_fmt.alignment = ali_list[j]
        Paragraph.font_size = size

    def _get_table_data(self):
        tab = []
        for ml in self.md_lines:
            if ml.text == '':
                continue
            tab.append(ml.text.split('|'))
        m = 0
        for rw in tab:
            if m < len(rw) - 1:
                m = len(rw) - 1
        for rw in tab:
            while len(rw) - 1 < m:
                rw.append('')
        # for i in range(len(tab)):
        #     for j in range(len(tab[i])):
        #         tab[i][j] = re.sub('^\\s+', '', tab[i][j])
        #         tab[i][j] = re.sub('\\s+$', '', tab[i][j])
        is_empty_r = True
        is_empty_l = True
        for rw in tab:
            if rw[m] != '':
                is_empty_r = False
            if rw[0] != '':
                is_empty_l = False
        for rw in tab:
            if is_empty_r:
                rw.pop(m)
            if is_empty_l:
                rw.pop(0)
        return tab

    def _get_table_alignment_and_width(self, tab):
        conf_row = -1
        for i in range(len(tab)):
            for j in range(len(tab[i])):
                if not re.match('^ *:?-*:? *$', tab[i][j]):
                    break
            else:
                conf_row = i
                break
        ali_list = []
        wid_list = []
        if conf_row >= 0:
            for s in tab[conf_row]:
                s = s.replace(' ', '')
                if re.match('^:-*:$', s):
                    ali_list.append(WD_TABLE_ALIGNMENT.CENTER)
                elif re.match('^-+:$', s):
                    ali_list.append(WD_TABLE_ALIGNMENT.RIGHT)
                else:
                    ali_list.append(WD_TABLE_ALIGNMENT.LEFT)
                wid_list.append(float(len(s)) / 2)
        else:
            for i in range(len(tab)):
                while len(ali_list) < len(tab[i]):
                    ali_list.append(WD_TABLE_ALIGNMENT.LEFT)
                while len(wid_list) < len(tab[i]):
                    wid_list.append(0.0)
                for j in range(len(tab[i])):
                    s = tab[i][j]
                    w = float(get_real_width(s)) / 2
                    if wid_list[j] < w:
                        wid_list[j] = w
        return conf_row, ali_list, wid_list

    def _write_image_paragraph(self, ms_doc):
        full_text = self.decoration_instruction
        for ml in self.md_lines:
            full_text += ml.text
        image_res = '! *\\[([^\\[\\]]*)\\] *\\(([^\\(\\)]+)\\)'
        full_text = re.sub('\\s*(' + image_res + ')\\s*', '\n\\1\n', full_text)
        full_text = re.sub('\\s*\\+\\+\\s*', '\n++\n', full_text)
        full_text = re.sub('\\s*\\-\\-\\s*', '\n--\n', full_text)
        full_text = re.sub('\n+', '\n', full_text)
        full_text = re.sub('^\n+', '', full_text)
        full_text = re.sub('\n+$', '', full_text)
        is_large = False
        is_small = False
        text_height \
            = PAPER_HEIGHT[doc.paper_size] - doc.top_margin - doc.bottom_margin
        text_width \
            = PAPER_WIDTH[doc.paper_size] - doc.left_margin - doc.right_margin
        for text in full_text.split('\n'):
            if re.match(image_res, text):
                comm = re.sub(image_res, '\\1', text)
                path = re.sub(image_res, '\\2', text)
                try:
                    if is_large:
                        if text_height > text_width:
                            ms_doc.add_picture(path, height=Cm(text_height))
                        else:
                            ms_doc.add_picture(path, width=Cm(text_width))
                    elif is_small:
                        if text_height > text_width:
                            ms_doc.add_picture(path, width=Cm(text_width))
                        else:
                            ms_doc.add_picture(path, height=Cm(text_height))
                    else:
                        ms_doc.add_picture(path)
                    ms_doc.paragraphs[-1].alignment \
                        = WD_ALIGN_PARAGRAPH.CENTER
                except BaseException:
                    e = ms_doc.paragraphs[-1]._element
                    e.getparent().remove(e)
                    ms_par = ms_doc.add_paragraph()
                    ms_par.add_run(text)
                    ms_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    msg = '警告: ' \
                        + '画像「' + path + '」が読み込めません'
                    # msg = 'warning: can\'t open "' + path + '"'
                    r = '^.*! *\\[.*\\] *\\(' + path + '\\).*$'
                    for ml in self.md_lines:
                        if re.match(r, ml.text):
                            if msg not in ml.warning_messages:
                                ml.append_warning_message(msg)
            elif text == '++':
                is_large = not is_large
                if is_small:
                    is_small = False
            elif text == '--':
                is_small = not is_small
                if is_large:
                    is_large = False

    def _write_preformatted_paragraph(self, ms_doc):
        ms_par = ms_doc.add_paragraph(style='makdo-g')
        text_to_write = ''
        md_lines = self.md_lines
        m = len(md_lines) - 1
        for i, ml in enumerate(md_lines):
            if i == 0:
                res = '^``` (\\s*)(.*)?$'
                if re.match(res, ml.raw_text):
                    text_to_write \
                        += re.sub(res, '\\1[\\2]', ml.raw_text)
            elif i == m:
                if re.match('^```( .*)?$', ml.raw_text):
                    continue
            else:
                if text_to_write != '':
                    text_to_write += '\n'
                text_to_write += ml.raw_text
        text_to_write = '`' + text_to_write + '`'
        self._write_text(text_to_write, ms_par)

    def _write_pagebreak_paragraph(self, ms_doc):
        ms_doc.add_page_break()

    def _write_sentence_paragraph(self, ms_doc):
        size = self.font_size
        text_to_write = self.decoration_instruction
        for ml in self.md_lines:
            text = re.sub('^ *', '', ml.text)
            text_to_write = self._join_string(text_to_write, text)
        ms_par = self._get_ms_par(ms_doc)
        ms_fmt = ms_par.paragraph_format
        if not re.match('^.*\n', text_to_write):
            ms_fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self._write_text(text_to_write, ms_par)

    @staticmethod
    def _join_string(string_a, string_b):
        string_joined = string_a + string_b
        if re.match('^.*[0-9A-Za-z,\\.\\)}\\]]$', string_a):
            if re.match('^[0-9A-Za-z\\({\\]].*$', string_b):
                string_joined = string_a + ' ' + string_b
        return string_joined

    def _get_ms_par(self, ms_doc, style=''):
        length = self.length
        size = self.font_size
        if style == '':
            ms_par = ms_doc.add_paragraph(style='makdo')
        else:
            ms_par = ms_doc.add_paragraph(style=style)
        if not doc.auto_space:
            pPr = ms_par._p.get_or_add_pPr()
            oe = OxmlElement('w:autoSpaceDE')
            oe.set(ns.qn('w:val'), '0')
            pPr.append(oe)
            oe = OxmlElement('w:autoSpaceDN')
            oe.set(ns.qn('w:val'), '0')
            pPr.append(oe)
        ms_fmt = ms_par.paragraph_format
        ms_fmt.widow_control = False
        if length['space before'] >= 0:
            ms_fmt.space_before \
                = Pt(length['space before'] * doc.line_spacing * size)
        else:
            ms_fmt.space_before = Pt(0)
            msg = '警告: ' \
                + '段落前の余白「v」の値が少な過ぎます'
            # msg = 'warning: ' \
            #     + '"space before" must be positive'
            self.md_lines[0].append_warning_message(msg)
        if length['space after'] >= 0:
            ms_fmt.space_after \
                = Pt(length['space after'] * doc.line_spacing * size)
        else:
            ms_fmt.space_after = Pt(0)
            msg = '警告: ' \
                + '段落後の余白「V」の値が少な過ぎます'
            # msg = 'warning: ' \
            #     + '"space after" must be positive'
            self.md_lines[0].append_warning_message(msg)
        ms_fmt.first_line_indent = Pt(length['first indent'] * size)
        ms_fmt.left_indent = Pt(length['left indent'] * size)
        ms_fmt.right_indent = Pt(length['right indent'] * size)
        # ms_fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        ls = doc.line_spacing * (1 + length['line spacing'])
        ms_fmt.line_spacing = Pt(ls * size)
        if ls < 1.0:
            msg = '警告: ' \
                + '段落後の余白「X」の値が少な過ぎます'
            # msg = 'warning: ' \
            #     + 'too small line spacing'
            self.md_lines[0].append_warning_message(msg)
        return ms_par

    def _write_text(self, text, ms_par):
        lns = text.split('\n')
        text = ''
        res = NOT_ESCAPED + '<br/?>'
        for ln in lns:
            while re.match(res, ln):
                ln = re.sub(res, '\\1\n', ln)
            text += ln + '\n'
        text = re.sub('\n$', '', text)
        text = Paragraph._remove_relax_symbol(text)
        res_img = '(.*(?:\n.*)*)! ?\\[([^\\[\\]]*)\\] ?\\(([^\\(\\)]+)\\)'
        tex = ''
        for c in text + '\0':
            if False:
                pass
            elif re.match(NOT_ESCAPED + '\\*\\*\\*$', tex + c):
                # *** (ITALIC AND BOLD)
                tex = re.sub('\\*\\*\\*$', '', tex + c)
                tex = self._write_string(tex, ms_par)
                Paragraph.is_italic = not Paragraph.is_italic
                Paragraph.is_bold = not Paragraph.is_bold
                continue
            elif re.match(NOT_ESCAPED + '\\*\\*$', tex) and c != '*':
                # ** (BOLD)
                tex = re.sub('\\*\\*$', '', tex)
                tex = self._write_string(tex, ms_par)
                tex += c
                Paragraph.is_bold = not Paragraph.is_bold
                continue
            elif re.match(NOT_ESCAPED + '\\*$', tex) and c != '*':
                # * (ITALIC)
                tex = re.sub('\\*$', '', tex)
                tex = self._write_string(tex, ms_par)
                tex += c
                Paragraph.is_italic = not Paragraph.is_italic
                continue
            elif re.match(NOT_ESCAPED + '~~$', tex + c):
                # ~~ (STRIKETHROUGH)
                tex = re.sub('~~$', '', tex + c)
                tex = self._write_string(tex, ms_par)
                Paragraph.has_strike = not Paragraph.has_strike
                continue
            elif re.match(NOT_ESCAPED + '`$', tex + c):
                # ` (PREFORMATTED)
                tex = re.sub('`$', '', tex + c)
                tex = self._write_string(tex, ms_par)
                Paragraph.is_preformatted = not Paragraph.is_preformatted
                continue
            elif re.match(NOT_ESCAPED + '//$', tex + c):
                # // (ITALIC)
                if not re.match('[a-z]+://', tex + c):
                    # not http:// https:// ftp:// ...
                    tex = re.sub('//$', '', tex + c)
                    tex = self._write_string(tex, ms_par)
                    Paragraph.is_italic = not Paragraph.is_italic
                    continue
            elif re.match(NOT_ESCAPED + '\\-\\-$', tex + c):
                # -- (SMALL)
                tex = re.sub('\\-\\-$', '', tex + c)
                tex = self._write_string(tex, ms_par)
                Paragraph.is_small = not Paragraph.is_small
                continue
            elif re.match(NOT_ESCAPED + '\\+\\+$', tex + c):
                # ++ (LARGE)
                tex = re.sub('\\+\\+$', '', tex + c)
                tex = self._write_string(tex, ms_par)
                Paragraph.is_large = not Paragraph.is_large
                continue
            elif re.match(NOT_ESCAPED + '\\^([0-9A-Za-z]*)\\^$', tex + c):
                # ^...^ (FONT COLOR)
                col = re.sub(NOT_ESCAPED + '\\^([0-9A-Za-z]*)\\^$', '\\2',
                             tex + c)
                if col == '':
                    col = 'FFFFFF'
                elif re.match('^([0-9A-F])([0-9A-F])([0-9A-F])$', col):
                    col = re.sub('^([0-9A-F])([0-9A-F])([0-9A-F])$',
                                 '\\1\\1\\2\\2\\3\\3', col)
                elif col in FONT_COLOR:
                    col = FONT_COLOR[col]
                if re.match('^[0-9A-F]{6}$', col):
                    tex = re.sub('\\^([0-9A-Za-z]*)\\^$', '', tex + c)
                    tex = self._write_string(tex, ms_par)
                    if Paragraph.font_color == '':
                        Paragraph.font_color = col
                    else:
                        Paragraph.font_color = ''
                    continue
            elif re.match(NOT_ESCAPED + '__$', tex + c):
                # __ (UNDERLINE)
                tex = re.sub('__$', '', tex + c)
                tex = self._write_string(tex, ms_par)
                Paragraph.has_underline = not Paragraph.has_underline
                continue
            elif re.match(NOT_ESCAPED + '_([0-9A-Za-z]+)_$', tex + c):
                # _..._ (HIGHLIGHT COLOR)
                col = re.sub(NOT_ESCAPED + '_([0-9A-Za-z]+)_$', '\\2', tex + c)
                if col in HIGHLIGHT_COLOR:
                    tex = re.sub('_([0-9A-Za-z]+)_$', '', tex + c)
                    tex = self._write_string(tex, ms_par)
                    if Paragraph.highlight_color is None:
                        Paragraph.highlight_color = HIGHLIGHT_COLOR[col]
                    else:
                        Paragraph.highlight_color = None
                    continue
            elif re.match(res_img, tex + c):
                # ![...](...)
                comm = re.sub(res_img, '\\2', tex + c)
                path = re.sub(res_img, '\\3', tex + c)
                tex = re.sub(res_img, '\\1', tex + c)
                tex = self._write_string(tex, ms_par)
                self._write_image(comm, path, ms_par)
            tex += c
        tex = re.sub('\0$', '', tex)
        if tex != '':
            tex = self._write_string(tex, ms_par)

    @staticmethod
    def _remove_relax_symbol(text):
        res = NOT_ESCAPED + RELAX_SYMBOL
        while re.match(res, text):
            text = re.sub(res, '\\1', text)
        return text

    @classmethod
    def _write_string(cls, string, ms_par):
        if string == '':
            return ''
        # REMOVE ESCAPE SYMBOL (BACKSLASH)
        string = re.sub('\\\\', '-\\\\', string)
        string = re.sub('-\\\\-\\\\', '-\\\\\\\\', string)
        string = re.sub('-\\\\', '', string)
        size = cls.font_size
        l_size = 1.2 * size
        s_size = 0.8 * size
        ms_run = ms_par.add_run(string)
        if cls.is_large and not cls.is_small:
            ms_run.font.size = Pt(l_size)
        elif not cls.is_large and cls.is_small:
            ms_run.font.size = Pt(s_size)
        else:
            ms_run.font.size = Pt(size)
        if cls.is_bold:
            ms_run.bold = True
        # else:
        #     ms_run.bold = False
        if cls.is_preformatted:
            ms_run.font.name = cls.gothic_font
        else:
            ms_run.font.name = cls.mincho_font
        ms_run._element.rPr.rFonts.set(qn('w:eastAsia'), ms_run.font.name)
        if cls.is_italic:
            ms_run.italic = True
        # else:
        #     ms_run.italic = False
        if cls.has_strike:
            ms_run.font.strike = True
        # else:
        #     ms_run.font.strike = False
        if cls.has_underline:
            ms_run.underline = True
        # else:
        #     ms_run.underline = False
        if cls.font_color != '':
            r = int(re.sub('^(..)(..)(..)$', '\\1', cls.font_color), 16)
            g = int(re.sub('^(..)(..)(..)$', '\\2', cls.font_color), 16)
            b = int(re.sub('^(..)(..)(..)$', '\\3', cls.font_color), 16)
            ms_run.font.color.rgb = RGBColor(r, g, b)
        if cls.highlight_color is not None:
            ms_run.font.highlight_color = cls.highlight_color
        return ''

    def _write_image(self, comm, path, ms_par):
        size = self.font_size
        l_size = 1.2 * size
        s_size = 0.8 * size
        ms_run = ms_par.add_run()
        try:
            if self.is_large and not self.is_small:
                ms_run.add_picture(path, height=Pt(l_size))
            elif not self.is_large and self.is_small:
                ms_run.add_picture(path, height=Pt(s_size))
            else:
                ms_run.add_picture(path, height=Pt(size))
        except BaseException:
            ms_run.text = '![' + comm + '](' + path + ')'
            msg = '警告: ' \
                + 'インライン画像「' + path + '」が読み込めません'
            # msg = 'warning: can\'t open "' + path + '"'
            r = '^.*! *\\[.*\\] *\\(' + path + '\\).*$'
            for ml in self.md_lines:
                if re.match(r, ml.text):
                    if msg not in ml.warning_messages:
                        ml.append_warning_message(msg)

    def print_warning_messages(self):
        for ml in self.md_lines:
            ml.print_warning_messages()


class MdLine:

    """A class to handle markdown line"""

    is_in_comment = False

    def __init__(self, line_number, raw_text):
        self.line_number = line_number
        self.raw_text = raw_text
        self.text = ''
        self.comment = ''
        self.warning_messages = []
        self.text, self.comment = self.separate_comment()

    def separate_comment(self):
        ori_sym = ORIGINAL_COMMENT_SYMBOL
        com_sep = COMMENT_SEPARATE_SYMBOL
        rt = self.raw_text
        text = ''
        comment = None
        if MdLine.is_in_comment:
            comment = ''
        tmp = ''
        for i, c in enumerate(rt):
            tmp += c
            if not MdLine.is_in_comment:
                if re.match(NOT_ESCAPED + '<!--$', tmp):
                    tmp = re.sub('<!--$', '', tmp)
                    text += tmp
                    tmp = ''
                    if comment is None:
                        comment = ''
                    MdLine.is_in_comment = True
            else:
                if re.match(NOT_ESCAPED + '-->$', tmp):
                    tmp = re.sub('-->$', '', tmp)
                    comment += tmp + com_sep
                    tmp = ''
                    MdLine.is_in_comment = False
            if not MdLine.is_in_comment:
                if re.match(NOT_ESCAPED + ori_sym + '$', tmp):
                    tmp = re.sub(ori_sym + '$', '', tmp)
                    text += tmp
                    tmp = ''
                    if comment is None:
                        comment = ''
                    comment += rt[i + 1:] + com_sep
                    break
        else:
            if tmp != '':
                if not MdLine.is_in_comment:
                    text += tmp
                    tmp = ''
                else:
                    if comment is None:
                        comment = ''
                    comment += tmp + com_sep
                    tmp = ''
        if comment is not None:
            comment = re.sub(com_sep + '$', '', comment)
        text = re.sub(NOT_ESCAPED + '<br/?>$', '\\1\n', text)
        # TRACK CHANGES
        res = NOT_ESCAPED + '<!?\\+>'
        while re.match(res, text):
            text = re.sub(res, '\\1', text)
        text = re.sub('  $', '\n', text)
        text = re.sub(' *$', '', text)
        # self.text = text
        # self.comment = comment
        return text, comment

    def append_warning_message(self, warning_message):
        self.warning_messages.append(warning_message)

    def print_warning_messages(self):
        for wm in self.warning_messages:
            msg = wm + ' (line ' + str(self.line_number) + ')' + '\n  ' \
                + self.raw_text
            sys.stderr.write(msg + '\n\n')


############################################################
# MAIN


if __name__ == '__main__':

    args = get_arguments()

    doc = Document()

    doc.raw_md_lines = doc.get_raw_md_lines(args.md_file)

    doc.md_lines = doc.get_md_lines(doc.raw_md_lines)

    doc.configure(doc.md_lines, args)

    doc.raw_paragraphs = doc.get_raw_paragraphs(doc.md_lines)
    doc.paragraphs = doc.get_paragraphs(doc.raw_paragraphs)
    doc.paragraphs = doc.modify_paragraphs(doc.paragraphs)

    ms_doc = doc.get_ms_doc()

    doc.write_property(ms_doc)

    doc.write_document(ms_doc)

    doc.save_docx_file(ms_doc, args.docx_file, args.md_file)

    doc.print_warning_messages()

    sys.exit(0)
