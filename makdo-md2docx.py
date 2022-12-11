#!/usr/bin/python3
# Name:         md2docx.py
# Version:      v02 Shin-Hakushima
# Time-stamp:   <2022.12.11-11:59:27-JST>

# md2docx.py
# Copyright (C) 2022  Seiichiro HATA
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


__version__ = 'v02 Shin-Hakushima'


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
        choices=['A3', 'A3P', 'A4', 'A4L'],
        help='用紙設定（A3、A3縦、A4、A4横）')
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
        help='文書スタイルの指定')
    parser.add_argument(
        '-N', '--no-page-number',
        action='store_true',
        help='ページ番号を出力しません')
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
        help='行間の幅（単位文字）')
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
    if not re.match('^(' + RES_NUMBER + '?,){,5}' + RES_NUMBER + '?,?$', s):
        raise argparse.ArgumentTypeError
    return s


HELP_EPILOG = '''Markdownの記法:
  行頭指示
    [#+=(数字) ]でセクション番号を変えることができます（独自）
    [v=(数字) ]で段落が下に数字行分ずれます（独自）
    [V=(数字) ]で次の段落が下に数字行分ずれます（独自）
    [X=(数字) ]で改行幅を数字行分増減します（独自）
    [<<=(数字) ]で1行目が左に数字文字分ずれます（独自）
    [<=(数字) ]で全体が左に数字文字分ずれます（独自）
  行中指示
    [;;]から行末まではコメントアウトされます（独自）
    [<>]は何もせず表示もされません（独自）
  文字装飾
    [*]で挟まれた文字列は斜体になります
    [**]で挟まれた文字列は太字になります
    [***]で挟まれた文字列は斜体かつ太字になります
    [~~]で挟まれた文字列は打消線が引かれます
    [__]で挟まれた文字列は下線が引かれます
    [//]で挟まれた文字列は斜体になります（修正）
    [++]で挟まれた文字列は文字が大きくなります（独自）
    [--]で挟まれた文字列は文字が小さくなります（独自）
    [@@]で挟まれた文字列は白色になって見えなくなります（独自）
'''

DEFAULT_DOCUMENT_TITLE = ''

DEFAULT_PAPER_SIZE = 'A4'
PAPER_HEIGHT = {'A3': 29.7, 'A3P': 42.0, 'A4': 29.7, 'A4L': 21.0}
PAPER_WIDTH = {'A3': 42.0, 'A3P': 29.7, 'A4': 21.0, 'A4L': 29.7}

DEFAULT_TOP_MARGIN = 3.5
DEFAULT_BOTTOM_MARGIN = 2.2
DEFAULT_LEFT_MARGIN = 3.0
DEFAULT_RIGHT_MARGIN = 2.0

DEFAULT_DOCUMENT_STYLE = '-'
DEFAULT_NO_PAGE_NUMBER = False
DEFAULT_LINE_NUMBER = False

DEFAULT_MINCHO_FONT = 'ＭＳ 明朝'
DEFAULT_GOTHIC_FONT = 'ＭＳ ゴシック'
DEFAULT_FONT_SIZE = 12.0

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

NOT_ESCAPED = '^((?:.*[^\\\\])?(?:\\\\\\\\)*)?'


class Title:

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
    def get_head_space(section_depth, n):
        if section_depth == 1:
            return ''
        elif section_depth == 4 and ((n == 0) or (n > 20)):
            return ' '
        elif section_depth == 6:
            return ' '
        else:
            return ZENKAKU_SPACE


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
        msg = 'warning: overflowed katakana "' + str(n) + '"\n'
        sys.stderr.write(msg)
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
        msg = 'warning: overflowed parenthesis katakata "' + str(n) + '"\n'
        sys.stderr.write(msg)
        return '(?)'


def n_alph(n):
    if n == 0:
        return chr(65344 + 26)
    elif n <= 26:
        return chr(65344 + n)
    else:
        msg = 'warning: overflowed alphabet "' + str(n) + '"\n'
        sys.stderr.write(msg)
        return '？'


def n_paren_alph(n):
    if n == 0:
        return chr(9371 + 26)
    elif n <= 26:
        return chr(9371 + n)
    else:
        msg = 'warning: overflowed parenthesis alphabet "' + str(n) + '"\n'
        sys.stderr.write(msg)
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
        self.paper_size = DEFAULT_PAPER_SIZE
        self.top_margin = DEFAULT_TOP_MARGIN
        self.bottom_margin = DEFAULT_BOTTOM_MARGIN
        self.left_margin = DEFAULT_LEFT_MARGIN
        self.right_margin = DEFAULT_RIGHT_MARGIN
        self.document_style = DEFAULT_DOCUMENT_STYLE
        self.no_page_number = DEFAULT_NO_PAGE_NUMBER
        self.line_number = DEFAULT_LINE_NUMBER
        self.mincho_font = DEFAULT_MINCHO_FONT
        self.gothic_font = DEFAULT_GOTHIC_FONT
        self.font_size = DEFAULT_FONT_SIZE
        self.line_spacing = DEFAULT_LINE_SPACING
        self.space_before = DEFAULT_SPACE_BEFORE
        self.space_after = DEFAULT_SPACE_AFTER
        self.auto_space = DEFAULT_AUTO_SPACE

    def get_raw_md_lines(self, md_file):
        self.md_file = md_file
        raw_md_lines = []
        with open(md_file, 'rb') as f:
            bf = f.read()
        kanji_code = chardet.detect(bf)['encoding']
        try:
            mfl = self._open_md_file(md_file, kanji_code).readlines()
        except BaseException:
            msg = 'error: not a markdown file "' + md_file + '"\n'
            sys.stderr.write(msg)
            sys.exit(0)
        if len(mfl) > 0 and len(mfl[0]) > 0:
            mfl[0] = re.sub('^' + chr(65279), '', mfl[0])  # remove BOM
        for rml in mfl:
            rml = re.sub('\n$', '', rml)
            rml = re.sub('\r$', '', rml)
            rml = re.sub('  $', '\n', rml)
            rml = re.sub('[ ' + ZENKAKU_SPACE + '\t]*$', '', rml)
            raw_md_lines.append(rml)
        raw_md_lines.append('')
        # self.raw_md_lines = raw_md_lines
        return raw_md_lines

    def _open_md_file(self, md_file, enc='utf-8'):
        if md_file == '-':
            mf = sys.stdin
        else:
            try:
                mf = open(md_file, 'r', encoding=enc)
            except BaseException:
                msg = 'error: can\'t read "' + md_file + '"\n'
                sys.stderr.write(msg)
                sys.exit(0)
        return mf

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
        for rp in raw_paragraphs:
            if rp.paragraph_class != 'empty':
                paragraphs.append(rp)
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
            res = '^\\s*(\\S+)\\s*:\\s*(\\S+)\\s*$'
            # res = '^[ \t]*([^ \t]+)[ \t]*:[ \t]*([^ \t]+)[ \t]*$'
            if re.match(res, com):
                nam = re.sub(res, '\\1', com)
                val = re.sub(res, '\\2', com)
                try:
                    eval('self.' + nam)
                except BaseException:
                    msg = 'no option "' + nam + '"'
                    ml.append_warning_message(msg)
                if nam == 'paper_size' or \
                   nam == 'mincho_font' or nam == 'gothic_font' or \
                   nam == 'document_style':
                    exec('self.' + nam + ' = "' + val + '"')
                elif nam == 'no_page_number' or nam == 'line_number':
                    if val == 'True' or val == 'False':
                        exec('self.' + nam + ' = ' + val)
                    else:
                        msg = 'not boolian "' + val + '"'
                        ml.append_warning_message(msg)
                elif nam == 'space_before' or nam == 'space_after':
                    if re.match('^' + RES_NUMBER6 + '$', val):
                        exec('self.' + nam + ' = "' + val + '"')
                    else:
                        msg = 'not decimals "' + val + '"'
                        ml.append_warning_message(msg)
                elif nam == 'document_title':
                    exec('self.' + nam + ' = "' + val + '"')
                else:
                    try:
                        exec('self.' + nam + ' = float(' + val + ')')
                    except BaseException:
                        msg = 'not decimal "' + val + '"'
                        ml.append_warning_message(msg)

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
        if args.no_page_number:
            self.no_page_number = True
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
        if not self.no_page_number:
            ms_par = ms_doc.sections[0].footer.paragraphs[0]
            ms_par.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            ms_run = ms_par.add_run()
            oe = OxmlElement('w:fldChar')
            oe.set(ns.qn('w:fldCharType'), 'begin')
            ms_run._r.append(oe)
            oe = OxmlElement('w:instrText')
            oe.set(ns.qn('xml:space'), 'preserve')
            oe.text = "PAGE"
            ms_run._r.append(oe)
            oe = OxmlElement('w:fldChar')
            oe.set(ns.qn('w:fldCharType'), 'end')
            ms_run._r.append(oe)
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
        user = os.getenv('USER')
        if user is None:
            user = os.getenv('USERNAME')
        if user is None:
            hs = '-'
        else:
            x = 9973
            b = 99999989
            m = 999999999989
            for c in user + ' 2022.05.07 07:31:03':
                x = (x * b + ord(c)) % m
            hs = ''
            for i in range(8):
                hs += chr(x % 26 + 97)
                x = int(x / 26)
        dt = datetime.datetime.now()
        ms_cp = ms_doc.core_properties
        ms_cp.title = self.document_title
        ms_cp.author = hs + ' (with makdo ' + __version__ + ')'
        ms_cp.created = dt
        ms_cp.modified = dt

    def write_document(self, ms_doc):
        for p in self.paragraphs:
            p.write_paragraph(ms_doc)

    def save_docx_file(self, ms_doc, docx_file, md_file):
        if docx_file == '':
            if re.match('^.*\\.md$', md_file):
                docx_file = re.sub('\\.md$', '.docx', md_file)
                self.docx_file = docx_file
        if os.path.exists(docx_file):
            if os.path.exists(docx_file + '~'):
                os.remove(docx_file + '~')
            os.rename(docx_file, docx_file + '~')
        ms_doc.save(docx_file)

    def print_warning_messages(self):
        for p in self.paragraphs:
            p.print_warning_messages()


class Paragraph:

    """A class to handle paragraph"""

    mincho_font = None
    font_size = None
    section_states = [0, 0, 0, 0, 0, 0]
    is_preformatted = False
    is_large = False
    is_small = False
    is_italic = False
    is_bold = False
    has_strike = False
    has_underline = False
    color = ''

    def __init__(self, paragraph_number, md_lines):
        self.paragraph_number = paragraph_number
        self.md_lines = md_lines
        self.full_text = ''
        self.paragraph_class = None
        self.decoration_instruction = ''
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
        self.section_instructions, \
            self.decoration_instruction, \
            self.length_ins, \
            self.md_lines \
            = self.read_first_line_instructions()
        self.full_text = self.get_full_text()
        self.paragraph_class \
            = self.get_paragraph_class()
        self.section_states, \
            self.section_depth_first, \
            self.section_depth \
            = self.get_section_states_and_depths()
        self.length_sec \
            = self.get_length_sec()
        self.length \
            = self.get_length()

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

    def read_first_line_instructions(self):
        section_instructions = []
        decoration_instruction = ''
        md_lines = self.md_lines
        length_ins \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        res_sn = '^\\s*(#+)=\\s*([0-9]+)(.*)$'
        res_de = '^\\s*' + \
            '((?:' + \
            '(?:\\*+)|(?:~~)|(?:__)|(?://)|(?:\\+\\+)|(?:--)|(?:@@)' + \
            '|(?:@[0-9A-F]*@)' + \
            ')+)' + \
            '(.*)$'
        res_sb = '^\\s*v=\\s*' + RES_NUMBER + '(.*)$'
        res_sa = '^\\s*V=\\s*' + RES_NUMBER + '(.*)$'
        res_ls = '^\\s*X=\\s*' + RES_NUMBER + '(.*)$'
        res_fi = '^\\s*<<=\\s*' + RES_NUMBER + '(.*)$'
        res_li = '^\\s*<=\\s*' + RES_NUMBER + '(.*)$'
        for ml in md_lines:
            # FOR BREAKDOWN
            if re.match('^-+::-*(::-+)?$', ml.text):
                continue
            while True:
                if re.match(res_sn, ml.text):
                    sect = re.sub(res_sn, '\\1', ml.text)
                    numb = re.sub(res_sn, '\\2', ml.text)
                    ml.text = re.sub(res_sn, '\\3', ml.text)
                    sec_dep = len(sect)
                    sec_num = int(numb)
                    section_instructions.append([sec_dep, sec_num])
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
                else:
                    break
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
        # self.section_instructions = section_instructions
        # self.decoration_instruction = decoration_instruction
        # self.length_ins = length_ins
        # self.md_lines = md_lines
        return section_instructions, decoration_instruction, \
            length_ins, md_lines

    def get_paragraph_class(self):
        decoration = self.decoration_instruction
        full_text = self.full_text
        paragraph_class = None
        if decoration + full_text == '':
            paragraph_class = 'empty'
        elif re.match('^\n$', decoration + full_text):
            paragraph_class = 'blank'
        elif re.match('^#+ ', full_text) or re.match('^#+$', full_text):
            paragraph_class = 'title'
        elif re.match('^' + NOT_ESCAPED + '::', full_text):
            paragraph_class = 'breakdown'
        elif re.match('^ *([-\\+\\*]|([0-9]+\\.)) ', full_text):
            paragraph_class = 'list'
        elif re.match('^: .*$', full_text) or re.match('^.* :$', full_text):
            paragraph_class = 'alignment'
        elif re.match('^\\|.*\\|$', full_text):
            paragraph_class = 'table'
        elif re.match('^(! ?\\[[^\\[\\]]*\\] ?\\([^\\(\\)]+\\) ?)+$',
                      full_text):
            paragraph_class = 'image'
        elif re.match('^```.*$', full_text):
            paragraph_class = 'preformatted'
        elif re.match('^<div style="break-.*: page;"></div>$', full_text):
            paragraph_class = 'pagebreak'
        elif re.match('<pgbr/?>', full_text):
            paragraph_class = 'pagebreak'
        else:
            paragraph_class = 'sentence'
        # self.paragraph_class = paragraph_class
        return paragraph_class

    def get_section_states_and_depths(self):
        states = []
        depth_first = 0
        depth = 0
        for si in self.section_instructions:
            dep = si[0]
            num = si[1]
            Paragraph.section_states[dep - 1] = num - 1
            for i in range(dep, len(Paragraph.section_states)):
                Paragraph.section_states[i] = 0
        for i, pss in enumerate(Paragraph.section_states):
            states.append(pss)
            if pss > 0:
                depth_first = i + 1
                depth = i + 1
        if self.paragraph_class == 'title':
            for i, sharps in enumerate(self.full_text.split(' ')):
                if re.match('^#+$', sharps):
                    if i == 0:
                        depth_first = len(sharps)
                    depth = len(sharps)
                    states[depth - 1] += 1
                    for i in range(depth, len(states)):
                        states[i] = 0
                else:
                    break
            for i, s in enumerate(states):
                Paragraph.section_states[i] = s
        # self.section_states = states
        # self.section_depth_first = depth_first
        # self.section_depth = depth
        return states, depth_first, depth

    def get_length_sec(self):
        length_sec \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        par_class = self.paragraph_class
        states = self.section_states
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
        paragraph_class = self.paragraph_class
        if paragraph_class == 'empty':
            self._write_empty_paragraph(ms_doc)
        elif paragraph_class == 'blank':
            self._write_blank_paragraph(ms_doc)
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
            msg = 'warning: unexpected state (empty paragraph)'
            sys.stderr.write(msg + '\n  ' + text_to_write + '\n')

    def _write_blank_paragraph(self, ms_doc):
        text_to_write = self.decoration_instruction
        for ml in self.md_lines:
            text_to_write += ml.text
        ms_par = self._get_ms_par(ms_doc)
        if text_to_write != '\n':
            self._write_text(text_to_write, ms_par)
            msg = 'warning: unexpected state (blank paragraph)'
            sys.stderr.write(msg + '\n  ' + text_to_write + '\n')

    def _write_title_paragraph(self, ms_doc):
        md_lines = self.md_lines
        size = self.font_size
        ll_size = size * 1.4
        depth = self.section_depth
        text_to_write = self.decoration_instruction
        head_symbol, title, text = self._split_title_paragraph(md_lines)
        head_string = self._get_title_head_string(head_symbol)
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
        has_title = False
        pre_dep = -1
        dep = 0
        for ml in md_lines:
            for i, c in enumerate(ml.text):
                if not is_in_head:
                    if i == 0:
                        title = Paragraph._join_string(title, c)
                    else:
                        title += c
                    continue
                if c == '#':
                    head += c
                    dep += 1
                elif c == ' ' or c == '\t':
                    head += ' '
                    Paragraph._is_consistent_with_depth(ml, pre_dep, dep)
                    pre_dep = dep
                    dep = 0
                else:
                    is_in_head = False
                    if re.match('.*#.*', ml.text[:i]):
                        has_title = True
                    title = c
            if is_in_head:
                head += ' '
                Paragraph._is_consistent_with_depth(ml, pre_dep, dep)
                pre_dep = dep
                dep = 0
        head = re.sub('\t', ' ', head)
        head = re.sub('^ *', '', head)
        head = re.sub(' +', ' ', head)
        if re.match('.*[．。]$', title):
            has_title = False
        if not has_title:
            title, text = '', title
        return head, title, text

    @staticmethod
    def _is_consistent_with_depth(md_line, pre_dep, dep):
        if pre_dep > 0:
            if (pre_dep <= 2) or (pre_dep + 1 != dep):
                msg = 'bad depth ' + str(pre_dep) + ' to ' + str(dep)
                md_line.append_warning_message(msg)
                return False
        return True

    def _get_title_head_string(self, head_symbol):
        sec_stat = self.section_states
        sec_dep = -1
        head_string = ''
        for hs in head_symbol.split(' '):
            if hs == '':
                continue
            sec_dep = len(hs)
            if sec_dep == 1:
                head_string += Title.get_head_1(sec_stat[0])
            elif sec_dep == 2:
                if doc.document_style == '-':
                    head_string += Title.get_head_2(sec_stat[1])
                else:
                    head_string += Title.get_head_2_j_or_J(sec_stat[1])
            elif sec_dep == 3:
                if doc.document_style != 'j':
                    head_string += Title.get_head_3(sec_stat[2])
                else:
                    head_string += Title.get_head_3(sec_stat[2] + 1)
            elif sec_dep == 4:
                head_string += Title.get_head_4(sec_stat[3])
            elif sec_dep == 5:
                head_string += Title.get_head_5(sec_stat[4])
            elif sec_dep == 6:
                head_string += Title.get_head_6(sec_stat[5])
        head_string += Title.get_head_space(sec_dep, sec_stat[sec_dep - 1])
        return head_string

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
        res = '^(! ?\\[([^\\[\\]]*)\\] ?\\(([^\\(\\)]+)\\) ?)+$'
        for ml in self.md_lines:
            if not re.match(res, ml.text):
                continue
            comm = re.sub(res, '\\2', ml.text)
            path = re.sub(res, '\\3', ml.text)
            try:
                ms_doc.add_picture(path)
                ms_doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            except BaseException:
                e = ms_doc.paragraphs[-1]._element
                e.getparent().remove(e)
                ms_par = ms_doc.add_paragraph()
                ms_par.add_run(ml.text)
                ms_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                msg = 'can\'t open "' + path + '"'
                self.md_lines[0].append_warning_message(msg)

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
        if re.match('^.*[0-9a-zA-Z,\\.\\)}\\]]$', string_a):
            if re.match('^[0-9a-zA-Z\\({\\]].*$', string_b):
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
            msg = 'warning: "space before" must be positive'
            self.md_lines[0].append_warning_message(msg)
        if length['space after'] >= 0:
            ms_fmt.space_after \
                = Pt(length['space after'] * doc.line_spacing * size)
        else:
            ms_fmt.space_after = Pt(0)
            msg = 'warning: "space after" must be positive'
            self.md_lines[0].append_warning_message(msg)
        ms_fmt.first_line_indent = Pt(length['first indent'] * size)
        ms_fmt.left_indent = Pt(length['left indent'] * size)
        if doc.document_style == 'j' and self.section_depth_first >= 3:
            ms_fmt.left_indent = Pt((length['left indent'] - 1) * size)
        ms_fmt.right_indent = Pt(length['right indent'] * size)
        # ms_fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        ls = doc.line_spacing * (1 + length['line spacing'])
        ms_fmt.line_spacing = Pt(ls * size)
        if ls < 1.0:
            msg = 'warning: too small line spacing'
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
        esc = ''
        tex = ''
        for c in text + '\0':
            if c == '`':
                if esc == '\\':
                    tex += '`'
                else:
                    tex = self._write_string(tex, ms_par)
                    Paragraph.is_preformatted = not Paragraph.is_preformatted
            elif self.is_preformatted:
                tex += c
            elif esc == '':
                esc, tex = self._set_esc_and_tex(c, tex)
                # INLINE IMAGE
                res = '^' \
                    + '(.*(?:\n.*)*)' \
                    + '! ?\\[([^\\[\\]]*)\\] ?\\(([^\\(\\)]+)\\)' \
                    + '$'
                if re.match(res, tex):
                    comm = re.sub(res, '\\2', tex)
                    path = re.sub(res, '\\3', tex)
                    tex = re.sub(res, '\\1', tex)
                    tex = self._write_string(tex, ms_par)
                    self._write_image(comm, path, ms_par)
                elif re.match(NOT_ESCAPED + '@([0-9A-F]*)@$', tex, re.I):
                    col = re.sub('.*@([0-9A-F]*)@$', '\\1', tex)
                    col = re.sub('^([0-9A-F])([0-9A-F])([0-9A-F])$',
                                 '\\1\\1\\2\\2\\3\\3', col)
                    if col == '':
                        col = 'FFFFFF'
                    tex = re.sub('@([0-9A-F]*)@$', '', tex)
                    tex = self._write_string(tex, ms_par)
                    esc = ''
                    if Paragraph.color == '':
                        Paragraph.color = col
                    else:
                        Paragraph.color = ''
            elif esc == '\\':
                if re.match('^[\\\\*~_/+\\-@]$', c):
                    esc = ''
                    tex += c
                else:
                    # tex += '\\'
                    esc, tex = self._set_esc_and_tex(c, tex)
            elif esc == '*':
                if c == '*':
                    esc += c
                else:
                    tex = self._write_string(tex, ms_par)
                    esc, tex = self._set_esc_and_tex(c, tex)
                    Paragraph.is_italic = not Paragraph.is_italic
            elif esc == '~':
                if c == '~':
                    esc = ''
                    tex = self._write_string(tex, ms_par)
                    Paragraph.has_strike = not Paragraph.has_strike
                else:
                    tex += '~'
                    esc, tex = self._set_esc_and_tex(c, tex)
            elif esc == '_':
                if c == '_':
                    esc = ''
                    tex = self._write_string(tex, ms_par)
                    Paragraph.has_underline = not Paragraph.has_underline
                else:
                    tex += '_'
                    esc, tex = self._set_esc_and_tex(c, tex)
            elif esc == '/':
                if c == '/':
                    esc = ''
                    if re.match('^.*[a-z]+:$', tex):
                        # http https ftp ...
                        tex += '//'
                    else:
                        tex = self._write_string(tex, ms_par)
                        Paragraph.is_italic = not Paragraph.is_italic
                else:
                    tex += '/'
                    esc, tex = self._set_esc_and_tex(c, tex)
            elif esc == '+':
                if c == '+':
                    esc = ''
                    tex = self._write_string(tex, ms_par)
                    Paragraph.is_large = not Paragraph.is_large
                else:
                    tex += '+'
                    esc, tex = self._set_esc_and_tex(c, tex)
            elif esc == '-':
                if c == '-':
                    esc = ''
                    tex = self._write_string(tex, ms_par)
                    Paragraph.is_small = not Paragraph.is_small
                else:
                    tex += '-'
                    esc, tex = self._set_esc_and_tex(c, tex)
                    esc, tex = self._set_esc_and_tex(c, tex)
            elif esc == '**':
                if c == '*':
                    esc = ''
                    tex = self._write_string(tex, ms_par)
                    Paragraph.is_italic = not Paragraph.is_italic
                    Paragraph.is_bold = not Paragraph.is_bold
                else:
                    tex = self._write_string(tex, ms_par)
                    esc, tex = self._set_esc_and_tex(c, tex)
                    Paragraph.is_bold = not Paragraph.is_bold
            else:
                esc, tex = self._set_esc_and_tex(c, tex)
        if tex != '':
            tex = self._write_string(tex, ms_par)

    @staticmethod
    def _remove_relax_symbol(text):
        res = NOT_ESCAPED + RELAX_SYMBOL
        while re.match(res, text):
            text = re.sub(res, '\\1', text)
        return text

    @classmethod
    def _set_esc_and_tex(cls, c, tex):
        if c == '\0':
            return '', tex
        elif re.match('^[\\\\*~_/+\\-]$', c):
            return c, tex
        else:
            return '', tex + c

    @classmethod
    def _write_string(cls, string, ms_par):
        if string == '':
            return ''
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
        if cls.color != '':
            r = int(re.sub('^(..)(..)(..)$', '\\1', cls.color), 16)
            g = int(re.sub('^(..)(..)(..)$', '\\2', cls.color), 16)
            b = int(re.sub('^(..)(..)(..)$', '\\3', cls.color), 16)
            ms_run.font.color.rgb = RGBColor(r, g, b)
        return ''

    def _write_image(self, comm, path, ms_par):
        size = self.font_size
        l_size = 1.2 * size
        s_size = 0.8 * size
        ms_run = ms_par.add_run()
        try:
            if self.is_large and not self.is_small:
                ms_run.add_picture(path, Pt(l_size), Pt(l_size))
            elif not self.is_large and self.is_small:
                ms_run.add_picture(path, Pt(s_size), Pt(s_size))
            else:
                ms_run.add_picture(path, Pt(size), Pt(size))
        except BaseException:
            ms_run.text = '![' + comm + '](' + path + ')'
            msg = 'can\'t open "' + path + '"'
            self.md_lines[0].append_warning_message(msg)

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
        self._warning_messages = []
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
        text = re.sub('  $', '\n', text)
        text = re.sub(' *$', '', text)
        # self.text = text
        # self.comment = comment
        return text, comment

    def append_warning_message(self, warning_message):
        self._warning_messages.append(warning_message)

    def print_warning_messages(self):
        for wm in self._warning_messages:
            msg = 'warning: ' + wm + ' (line ' + str(self.line_number) + ')'
            sys.stderr.write(msg + '\n  ' + self.raw_text + '\n')


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
