#!/usr/bin/python3
# Name:         docx2md.py
# Version:      v02 Shin-Hakushima
# Time-stamp:   <2022.12.13-02:56:25-JST>

# docx2md.py
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
import tempfile
import shutil
import argparse
import re
import unicodedata


__version__ = 'v02 Shin-Hakushima'


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
        choices=['A3', 'A3P', 'A4', 'A4L'],
        help='用紙設定（A3、A3縦、A4、A4横）')
    parser.add_argument(
        '-t', '--top-margin',
        type=float,
        help='上余白（単位cm）')
    parser.add_argument(
        '-b', '--bottom-margin',
        type=float,
        help='下余白（単位cm）')
    parser.add_argument(
        '-l', '--left-margin',
        type=float,
        help='左余白（単位cm）')
    parser.add_argument(
        '-r', '--right-margin',
        type=float,
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
        help='明朝フォント')
    parser.add_argument(
        '-g', '--gothic-font',
        type=str,
        help='ゴシックフォント')
    parser.add_argument(
        '-f', '--font-size',
        type=float,
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
        'docx_file',
        help='MS Wordファイル')
    parser.add_argument(
        'md_file',
        default='',
        nargs='?',
        help='Markdownファイル（"-"は標準出力）')
    return parser.parse_args()


def floats6(s):
    if not re.match('^(' + RES_NUMBER + '?,){,5}' + RES_NUMBER + '?,?$', s):
        raise argparse.ArgumentTypeError
    return s


HELP_EPILOG = '''
'''

DEFAULT_DOCUMENT_TITLE = ''

DEFAULT_PAPER_SIZE = 'A4'
PAPER_HEIGHT = {'A3': 29.7, 'A3P': 42.0, 'A4': 29.7, 'A4L': 21.0}
PAPER_WIDTH = {'A3': 42.0, 'A3P': 29.7, 'A4': 21.0, 'A4L': 29.7}

DEFAULT_TOP_MARGIN = 3.5
DEFAULT_BOTTOM_MARGIN = 2.0
DEFAULT_LEFT_MARGIN = 3.0
DEFAULT_RIGHT_MARGIN = 2.0

DEFAULT_MINCHO_FONT = 'ＭＳ 明朝'
DEFAULT_GOTHIC_FONT = 'ＭＳ ゴシック'
DEFAULT_FONT_SIZE = 12.0

DEFAULT_DOCUMENT_STYLE = '-'
DEFAULT_NO_PAGE_NUMBER = False
DEFAULT_LINE_NUMBER = False

DEFAULT_LINE_SPACING = 2.14  # (2.0980+2.1812)/2=2.1396

DEFAULT_SPACE_BEFORE = ''
DEFAULT_SPACE_AFTER = ''

DEFAULT_AUTO_SPACE = False

ZENKAKU_SPACE = chr(12288)

NOT_ESCAPED = '^((?:.*[^\\\\])?(?:\\\\\\\\)*)?'

MD_TEXT_WIDTH = 79


class Title:

    r0 = '((?:__)|(?:\\+\\+)|(?:--)|(?:~~)|(?:\\*+)|(?:@@))*'
    r1 = '(__)?\\+\\+(.*)\\+\\+(__)?'
    r2 = '(第([0-9０-９]+)条?)'
    r3 = '([0-9０-９]+)'
    r4 = '([⑴-⒇]|([\\(（]([0-9０-９]+)[\\)）]))'
    r5 = '([ｱ-ﾝア-ン])'
    r6 = '([(\\(（]([ｱ-ﾝア-ン])[\\)）])'
    r9 = '((  ?)|(\t)|(' + ZENKAKU_SPACE + ')|(\\. ?)|(．))'
    res1 = r1
    res2 = r0 + r2 + r9
    res3 = r0 + r3 + '(((' + r4 + '?)' + r5 + '?)' + r6 + '?)' + r9
    res4 = r0 + '(' + r3 + ')?' + r4 + '((' + r5 + '?)' + r6 + '?)' + r9 + '?'
    res5 = r0 + '((' + r3 + ')?' + r4 + ')?' + r5 + '(' + r6 + '?)' + r9
    res6 = r0 + '(((' + r3 + ')?' + r4 + ')?' + r5 + '?)' + r6 + r9 + '?'
    not3 = r3 + r9 + '.*\n[ \t' + ZENKAKU_SPACE + ']*' + r3 + r9

    @classmethod
    def get_depth(cls, line, alignment):
        if re.match(cls.res1, line) and alignment == 'center':
            return 1
        if re.match(cls.res2, line):
            return 2
        if re.match(cls.res3, line) and not re.match(cls.not3, line):
            return 3
        if re.match(cls.res4, line):
            return 4
        if re.match(cls.res5, line):
            return 5
        if re.match(cls.res6, line):
            return 6
        return -1

    @classmethod
    def decompose(cls, depth, line):
        if depth == 1:
            res = '^' + cls.res1 + '$'
            comm = ''
            head = ''
            rest = ''
            text = re.sub(res, '\\1\\2\\3', line)
            numb = -1
        elif depth == 2:
            res = '^' + cls.res2 + '(.*)$'
            comm = re.sub(res, '\\1', line)
            head = re.sub(res, '\\2', line)
            rest = ''
            text = re.sub(res, '\\10', line)
            numb = inverse_n_int(re.sub(res, '\\3', line))
        elif depth == 3:
            res = '^' + cls.res3 + '(.*)$'
            comm = re.sub(res, '\\1', line)
            head = re.sub(res, '\\2', line)
            rest = re.sub(res, '\\3', line)
            text = re.sub(res, '\\18', line)
            numb = inverse_n_int(re.sub(res, '\\2', line))
        elif depth == 4:
            res = '^' + cls.res4 + '(.*)$'
            comm = re.sub(res, '\\1', line)
            head = re.sub(res, '\\4', line)
            rest = re.sub(res, '\\7', line)
            text = re.sub(res, '\\18', line)
            if re.match('^[⑴-⒇]$', head):
                numb = ord(head) - 9331
            else:
                numb = inverse_n_int(re.sub(res, '\\6', line))
        elif depth == 5:
            res = '^' + cls.res5 + '(.*)$'
            comm = re.sub(res, '\\1', line)
            head = re.sub(res, '\\8', line)
            rest = re.sub(res, '\\9', line)
            text = re.sub(res, '\\18', line)
            numb = inverse_n_kata(re.sub(res, '\\8', line))
        elif depth == 6:
            res = '^' + cls.res6 + '(.*)$'
            comm = re.sub(res, '\\1', line)
            head = re.sub(res, '\\10', line)
            rest = ''
            text = re.sub(res, '\\18', line)
            numb = inverse_n_kata(re.sub(res, '\\11', line))
        else:
            comm = ''
            head = ''
            rest = ''
            text = ''
            numb = -1
        if comm == line:
            comm = ''
        if head == line:
            head = ''
        if rest == line:
            rest = ''
        if rest != '':
            return comm, numb, head, rest + ZENKAKU_SPACE + text
        else:
            return comm, numb, head, text


class List:

    res_b1 = '(•((  ?)|(\t)|(' + ZENKAKU_SPACE + ')))'  # U+2022 Bullet
    res_b2 = '(◦((  ?)|(\t)|(' + ZENKAKU_SPACE + ')))'  # U+25E6 White Bullet
    res_b3 = '(‣((  ?)|(\t)|(' + ZENKAKU_SPACE + ')))'  # U+2023 Triangular Bul
    res_b4 = '(⁃((  ?)|(\t)|(' + ZENKAKU_SPACE + ')))'  # U+2043 Hyphen Bullet
    res_n1 = '(([0-9０-９]+)((\\. ?)|(．)))'
    res_n2 = '(([0-9０-９]+)((\\) ?)|(）)))'
    res_n3 = '(([a-zａ-ｚ]+)((\\. ?)|(．)))'
    res_n4 = '(([a-zａ-ｚ]+)((\\) ?)|(）)))'

    @classmethod
    def get_division_and_depth(cls, line):
        if re.match(cls.res_b1, line):
            return 'b1'
        if re.match(cls.res_b2, line):
            return 'b2'
        if re.match(cls.res_b3, line):
            return 'b3'
        if re.match(cls.res_b4, line):
            return 'b4'
        if re.match(cls.res_n1, line):
            return 'n1'
        if re.match(cls.res_n2, line):
            return 'n2'
        if re.match(cls.res_n3, line):
            return 'n3'
        if re.match(cls.res_n4, line):
            return 'n4'
        return ''

    @classmethod
    def decompose(cls, type_and_depth, line):
        if type_and_depth == 'b1':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_b1 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = -1
        elif type_and_depth == 'b2':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_b2 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = -1
        elif type_and_depth == 'b3':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_b3 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = -1
        elif type_and_depth == 'b4':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_b4 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = -1
        elif type_and_depth == 'n1':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_n1 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = inverse_n_int(re.sub(res, '\\2', line))
        elif type_and_depth == 'n2':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_n2 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = inverse_n_int(re.sub(res, '\\2', line))
        elif type_and_depth == 'n3':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_n3 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = inverse_n_alph(re.sub(res, '\\2', line))
        elif type_and_depth == 'n4':
            res = '^[ \t' + ZENKAKU_SPACE + ']*' + cls.res_n4 + '(.*)$'
            head = re.sub(res, '\\1', line)
            text = re.sub(res, '\\6', line)
            numb = inverse_n_alph(re.sub(res, '\\2', line))
        else:
            head = ''
            text = ''
            numb = -1
        return numb, head, text


############################################################
# FUNCTION


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
        if p != '' and p != c:
            wid += 0.5
        p = w
    return wid


def inverse_n_int(s):
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


def inverse_n_kata(s):
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


def inverse_n_alph(s):
    c = s
    if re.match('^[a-z]$', c):
        return ord(c) - 96
    elif re.match('^[ａ-ｚ]$', c):
        return ord(c) - 65344
    return -1


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

    def __init__(self):
        self.tmpdir = None
        self.media_dir = None
        self.docx_file = None
        self.md_file = None
        self.core_raw_xml_lines = None
        self.footer1_raw_xml_lines = None
        self.styles_raw_xml_lines = None
        self.rels_raw_xml_lines = None
        self.document_raw_xml_lines = None
        self.styles = None
        self.rels = None
        self.images = {}
        self.raw_paragraphs = None
        self.paragraphs = None
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
        self.birthtime = None

    def make_tmpdir(self):
        tmpdir = tempfile.TemporaryDirectory()
        # self.tmpdir = tmpdir
        return tmpdir

    def get_media_dir_name(self, md_file, docx_file):
        media_dir = ''
        if md_file != '':
            if re.match('^.*\\.md$', md_file, re.I):
                media_dir = re.sub('\\.md$', '', md_file, re.I)
        else:
            if re.match('^.*\\.docx$', docx_file, re.I):
                media_dir = re.sub('\\.docx$', '', docx_file, re.I)
        # self.media_dir = media_dir
        return media_dir

    def extract_docx_file(self, docx_file):
        self.docx_file = docx_file
        tmpdir = self.tmpdir.name
        try:
            shutil.unpack_archive(docx_file, tmpdir, 'zip')
        except BaseException:
            msg = 'error: not a ms word file "' + docx_file + '"\n'
            sys.stderr.write(msg)
            sys.exit(1)

    def get_raw_xml_lines(self, xml_file):
        path = self.tmpdir.name + '/' + xml_file
        if not os.path.exists(path):
            return []
        try:
            xf = open(path, 'r', encoding='utf-8')
        except BaseException:
            sys.stderr.write('error: can\'t read "' + xml_file + '"\n')
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
        # DOCUMENT TITLE
        self._configure_by_core_xml(self.core_raw_xml_lines)
        # PAGE NUMBER
        self._configure_by_footer1_xml(self.footer1_raw_xml_lines)
        # FONT, LINE SPACING
        self._configure_by_styles_xml(self.styles_raw_xml_lines)
        # PAPER SIZE, MARGIN, LINE NUMBER
        self._configure_by_document_xml(self.document_raw_xml_lines)
        # REVISE
        self._configure_by_args(args)
        # PARAGRAPH
        Paragraph.mincho_font = self.mincho_font
        Paragraph.gothic_font = self.gothic_font
        Paragraph.font_size = self.font_size

    def _configure_by_core_xml(self, raw_xml_lines):
        for i, rxl in enumerate(raw_xml_lines):
            resb = '^<dc:title>$'
            rese = '^</dc:title>$'
            if i > 0 and re.match(resb, raw_xml_lines[i - 1], re.I):
                if not re.match(rese, rxl, re.I):
                    self.document_title = rxl
            resb = '^<dcterms:modified( .*)?>$'
            rese = '^</dcterms:modified>$'
            if i > 0 and re.match(resb, raw_xml_lines[i - 1], re.I):
                if not re.match(rese, rxl, re.I):
                    self.birthtime = rxl
                    self.birthtime = re.sub('T', ' ', self.birthtime)
                    self.birthtime = re.sub('Z', '', self.birthtime)

    def _configure_by_footer1_xml(self, raw_xml_lines):
        self.no_page_number = True
        for rxl in raw_xml_lines:
            if rxl == 'PAGE':
                self.no_page_number = False

    def _configure_by_styles_xml(self, raw_xml_lines):
        xml_body = self._get_xml_body('w:styles', raw_xml_lines)
        xml_blocks = self._get_xml_blocks(xml_body)
        sb = ['', '', '', '', '', '']
        sa = ['', '', '', '', '', '']
        for xb in xml_blocks:
            name = ''
            font = ''
            sz_x = -1.0
            ls_x = -1.0
            for xl in xb:
                name = get_xml_value('w:name', 'w:val', name, xl)
                font = get_xml_value('w:rFonts', '*', font, xl)
                sz_x = get_xml_value('w:sz', 'w:val', sz_x, xl)
                ls_x = get_xml_value('w:spacing', 'w:line', ls_x, xl)
            if name == 'makdo':
                self.mincho_font = font
                if sz_x > 0:
                    self.font_size = round(sz_x / 2, 1)
                if ls_x > 0:
                    self.line_spacing = round(ls_x / 20 / self.font_size, 2)
            elif name == 'makdo-g':
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
        csb = ',' + ','.join(sb) + ','
        csb = re.sub(',0\\.0,', ',,', csb)
        csb = re.sub('\\.0,', ',', csb)
        csb = re.sub('^,', '', csb)
        csb = re.sub(',+$', '', csb)
        csa = ',' + ','.join(sa) + ','
        csa = re.sub(',0\\.0,', ',,', csa)
        csa = re.sub('\\.0,', ',', csa)
        csa = re.sub('^,', '', csa)
        csa = re.sub(',+$', '', csa)
        if csb != '':
            self.space_before = csb
        if csa != '':
            self.space_after = csa

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
            if re.match('^<w:lnNumType( .*)?>$', rxl):
                self.line_number = True
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
        for n, xb in enumerate(xml_blocks):
            p = Paragraph(n + 1, xb)
            raw_paragraphs.append(p)
        # self.raw_paragraphs = raw_paragraphs
        return raw_paragraphs

    def get_paragraphs(self, raw_paragraphs):
        paragraphs = []
        for rp in raw_paragraphs:
            if rp.paragraph_class != 'configuration':
                paragraphs.append(rp)
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
        # LEFT ALIGNMENT
        self.paragraphs = self._modpar_left_alignment()
        # SPACE BEFORE AND AFTER
        self.paragraphs = self._modpar_title_space_before_and_after()
        self.paragraphs = self._modpar_brank_paragraph_to_space_before()
        # LIST
        # self.paragraphs = self._modpar_title_3_to_list()
        # CENTERING
        self.paragraphs = self._modpar_centering_with_title_1()
        # INDENT
        self.paragraphs = self._modpar_one_line_paragraph()
        # RETURN
        return self.paragraphs

    def _modpar_left_alignment(self):
        for i, p in enumerate(self.paragraphs):
            if p.paragraph_class == 'sentence':
                if p.length['first indent'] == 0:
                    if p.length['left indent'] == 0:
                        p.paragraph_class = 'alignment'
                        p.alignment = 'left'
        return self.paragraphs

    def _modpar_title_space_before_and_after(self):
        m = len(self.paragraphs) - 1
        for i, p in enumerate(self.paragraphs):
            if i > 0:
                p_prev = self.paragraphs[i - 1]
            if i < m:
                p_next = self.paragraphs[i + 1]
            if p.paragraph_class != 'title':
                continue
            if p.section_depth_first == 1:
                if i > 0 and p_prev.paragraph_class == 'blank':
                    p_prev.paragraph_class = 'empty'
                    p.length_ins['space before'] += 1.0
                # if i > 0:
                #     p.length_ins['space before'] -= 1.0
                if i < m and p_next.paragraph_class == 'blank':
                    p_next.paragraph_class = 'empty'
                    p.length_ins['space after'] += 1.0
                # if i < m:
                #     p.length_ins['space after'] -= 1.0
                if (p.length_ins['space after'] >= 0.1) or \
                   (i < m and p_next.length_ins['space before'] >= 0.1):
                    if i > 0 and p_prev.length_ins['space after'] >= 0.1:
                        p_prev.length_ins['space after'] -= 0.1
                    if p.length_ins['space before'] >= 0.1:
                        p.length_ins['space before'] -= 0.1
                    if p.length_ins['space after'] >= 0.1:
                        p.length_ins['space after'] += 0.1
                    if i < m and p_next.length_ins['space before'] >= 0.1:
                        p_next.length_ins['space before'] += 0.1
            # elif p.section_depth_first == 2:
            #     if i > 0 and p_prev.paragraph_class == 'blank':
            #         p.length_ins['space before'] += 1.0
            #         p_prev.paragraph_class = 'empty'
            #     p.length_ins['space before'] -= 1.0
            # elif p.section_depth_first == 3 and p.section_states[1] == 0:
            #     if i > 0 and p_prev.paragraph_class == 'blank':
            #         p.length_ins['space before'] += 1.0
            #         p_prev.paragraph_class = 'empty'
            #     p.length_ins['space before'] -= 1.0
            p.first_line_instructions = p.get_first_line_instructions()
        return self.paragraphs

    def _modpar_brank_paragraph_to_space_before(self):
        for i, p in enumerate(self.paragraphs):
            # if p.paragraph_class == 'table':
            #     continue
            if i == 0:
                continue
            p_prev = self.paragraphs[i - 1]
            if p_prev.paragraph_class == 'blank':
                p.length_ins['space before'] += 1.0
                p_prev.paragraph_class = 'empty'
                p.first_line_instructions = p.get_first_line_instructions()
        return self.paragraphs

    def _modpar_title_3_to_list(self):
        for p in self.paragraphs:
            if p.paragraph_class != 'title':
                continue
            if p.section_depth_first != 3:
                continue
            if p.section_depth != 3:
                continue
            if p.section_states[1] > 0:
                if p.length_ins['space before'] > -0.5 and \
                   p.length_ins['space before'] < +0.5:
                    if p.length_ins['first indent'] > -0.5 and \
                       p.length_ins['first indent'] < +0.5:
                        if p.length_ins['left indent'] > -0.5 and \
                           p.length_ins['left indent'] < +0.5:
                            continue
            else:
                if p.length_ins['space before'] > +0.5 and \
                   p.length_ins['space before'] < +1.5:
                    if p.length_ins['first indent'] > -1.5 and \
                       p.length_ins['first indent'] < +0.5:
                        if p.length_ins['left indent'] > -0.5 and \
                           p.length_ins['left indent'] < +0.5:
                            continue
            p.paragraph_class = 'list_system'
            p.md_text = re.sub('^### ', '1. ', p.md_text)
            p.length_ins['first indent'] -= 1
            p.first_line_instructions = p.get_first_line_instructions()
        return self.paragraphs

    def _modpar_centering_with_title_1(self):
        is_list = False
        m = len(self.paragraphs) - 1
        for i, p in enumerate(self.paragraphs):
            if is_list:
                if p.paragraph_class == 'title' and \
                   p.section_depth_first == 3 and \
                   p.section_depth == 3:
                    p.paragraph_class = 'list_system'
                    p.md_text = re.sub('^### ', '1. ', p.md_text)
                elif p.paragraph_class == 'list_system':
                    pass
                elif p.paragraph_class == 'list':
                    pass
                else:
                    is_list = False
                p.section_depth_first = 1
                p.section_depth = 1
                p.length_sec = p.get_length_sec()
                p.length_spa = p.get_length_spa()
                p.length_ins = p.get_length_ins()
                p.first_line_instructions = p.get_first_line_instructions()
            if i == 0:
                continue
            if i == m:
                continue
            p_prev = self.paragraphs[i - 1]
            p_next = self.paragraphs[i + 1]
            if p.paragraph_class != 'alignment':
                continue
            if p.alignment != 'center':
                continue
            if p_prev.paragraph_class != 'blank' and \
               p.length_ins['space before'] <= 0:
                continue
            if p_next.paragraph_class != 'title' and \
               p_next.paragraph_class != 'list_system' and \
               p_next.paragraph_class != 'list':
                continue
            if p_next.paragraph_class == 'title':
                if p_next.section_depth_first != 3:
                    continue
                if p_next.section_depth != 3:
                    continue
            is_list = True
            if p_prev.paragraph_class == 'blank':
                p_prev.paragraph_class = 'empty'
            if p.first_line_instructions == '':
                p.first_line_instructions = '#\n' + p.first_line_instructions
            else:
                p.first_line_instructions = '#\n\n' + p.first_line_instructions
        return self.paragraphs

    def _modpar_one_line_paragraph(self):
        for p in self.paragraphs:
            rt = p.raw_text
            cms = ['\\*\\*\\*', '\\*\\*', '\\*', '~~', '--', '\\+\\+', '__',
                   '@@']
            for cm in cms:
                while re.match(NOT_ESCAPED + cm, rt):
                    rt = re.sub(NOT_ESCAPED + cm, '\\1', rt)
            while re.match(NOT_ESCAPED + '\\\\', rt):
                rt = re.sub(NOT_ESCAPED + '\\\\', '\\1', rt)
            lm = self.left_margin
            rm = self.right_margin
            fi = p.length['first indent']
            li = p.length['left indent']
            ri = p.length['right indent']
            tx = float(get_real_width(rt))
            w = (float(fi + li + tx + ri) * self.font_size * 2.54 / 72 / 2) \
                + lm + rm
            if w > PAPER_WIDTH[self.paper_size]:
                continue
            ifi = float(round(p.length_ins['first indent'] * 2)) / 2
            ili = float(round(p.length_ins['left indent'] * 2)) / 2
            if ifi + ili != 0:
                continue
            p.length_ins['first indent'] = 0
            p.length_ins['left indent'] = 0
            t1 = ''
            if re.match('^(#\n+).*$', p.first_line_instructions):
                t1 = re.sub('^(#\n+).*$', '\\1', p.first_line_instructions)
            p.first_line_instructions \
                = t1 + p.get_first_line_instructions()
        return self.paragraphs

    def check_section_consistency(self):
        m = len(Paragraph.section_states)
        section_states = []
        for ss in range(m):
            section_states.append(0)
        for i, p in enumerate(self.paragraphs):
            if p.paragraph_class != 'title':
                continue
            depth = -1
            ln = p.md_text
            ln = re.sub('\n', ' ', ln)
            ln = re.sub(' +', ' ', ln)
            res = '^' \
                + '((?:__)|(?:\\+\\+)|(?:--)|(?:~~)|(?:\\*+)|(?:@@))' \
                + '*((#+ )*).*$'
            head = re.sub(res, '\\2', ln + ' ')
            head = re.sub(' +', ' ', head)
            head = re.sub(' $', '', head)
            for sharps in head.split(' '):
                i = len(sharps) - 1
                section_states[i] += 1
                for j in range(i + 1, m):
                    section_states[j] = 0
                dp = i + 1
                if depth == -1:
                    depth = dp
                    continue
                depth += 1
                if depth != dp:
                    msg = 'warning: bad section depth "' + p.md_text + '"\n'
                    sys.stderr.write(msg)
            if i == 0:
                continue
            if self.document_style != 'j' or p.section_depth != 3:
                if section_states[i] != p.section_states[i]:
                    msg = 'warning: bad section number "' + p.md_text + '"\n'
                    sys.stderr.write(msg)
            else:
                if section_states[i] != p.section_states[i] - 1:
                    msg = 'warning: bad section number "' + p.md_text + '"\n'
                    sys.stderr.write(msg)

    def open_md_file(self, md_file, docx_file):
        if md_file == '-':
            mf = sys.stdout
        else:
            if md_file == '':
                if re.match('^.*\\.docx$', docx_file):
                    md_file = re.sub('\\.docx$', '.md', docx_file)
            if os.path.exists(md_file):
                if os.path.exists(md_file + '~'):
                    os.remove(md_file + '~')
                os.rename(md_file, md_file + '~')
            try:
                mf = open(md_file, 'w', encoding='utf-8', newline='\n')
            except BaseException:
                sys.stderr.write('error: can\'t write "' + md_file + '"\n')
                sys.exit(1)
        return mf

    def write_configurations(self, mf):
        mf.write('<!--\n')
        mf.write('document_title: ' + self.document_title + '\n')
        mf.write('document_style: ' + self.document_style + '\n')
        mf.write('no_page_number: ' + str(self.no_page_number) + '\n')
        mf.write('line_number:    ' + str(self.line_number) + '\n')
        mf.write('paper_size:     ' + str(self.paper_size) + '\n')
        mf.write('top_margin:     ' + str(round(self.top_margin, 1)) + '\n')
        mf.write('bottom_margin:  ' + str(round(self.bottom_margin, 1)) + '\n')
        mf.write('left_margin:    ' + str(round(self.left_margin, 1)) + '\n')
        mf.write('right_margin:   ' + str(round(self.right_margin, 1)) + '\n')
        mf.write('mincho_font:    ' + self.mincho_font + '\n')
        mf.write('gothic_font:    ' + self.gothic_font + '\n')
        mf.write('font_size:      ' + str(round(self.font_size, 1)) + '\n')
        mf.write('line_spacing:   ' + str(round(self.line_spacing, 2)) + '\n')
        mf.write('space_before:   ' + self.space_before + '\n')
        mf.write('space_after:    ' + self.space_after + '\n')
        mf.write('auto_space:     ' + str(self.auto_space) + '\n')
        mf.write('birthtime:      ' + self.birthtime + '\n')
        mf.write('-->\n\n')
        return

    def write_md_lines(self, mf):
        ps = self.paragraphs
        for i, p in enumerate(ps):
            p.write_md_lines(mf)

    def make_media_dir(self, media_dir):
        if len(self.images) == 0:
            return
        if media_dir == '':
            sys.stderr.write('error: can\'t make media directory\n')
            return
        if os.path.exists(media_dir):
            if os.path.isdir(media_dir):
                shutil.rmtree(media_dir)
            else:
                sys.stderr.write('error: non-directory "' + media_dir + '"\n')
                return
        os.mkdir(media_dir)
        for rel_img in self.images:
            if rel_img == '':
                continue
            shutil.copy(self.tmpdir.name + '/word/' + rel_img,
                        media_dir + '/' + self.images[rel_img])
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


class Paragraph:

    """A class to handle paragraph"""

    mincho_font = None
    gothic_font = None
    font_size = None
    section_states = [0, 0, 0, 0, 0, 0]

    def __init__(self, paragraph_number, raw_xml_lines):
        self.paragraph_number = paragraph_number
        self.raw_xml_lines = raw_xml_lines
        self.raw_class = None
        self.paragraph_class = None
        self.xml_lines = []
        self.raw_text = ''
        self.beg_space = ''
        self.end_space = ''
        self.raw_md_text = ''
        self.md_text = ''
        self.first_line_instructions = ''
        self.section_states = []
        self.section_depth_first = 0
        self.section_depth = 0
        self.style = None
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
        self.length_spa \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        self.substitute_values()

    def substitute_values(self):
        self.raw_class = self._get_raw_class()
        self.xml_lines = self._get_xml_lines()
        self.raw_text = self._get_raw_text()
        self.beg_space, self.raw_text, self.end_space = self.separate_space()
        self.style = self.get_and_apply_style()
        self.alignment = self.get_alignment()
        self.paragraph_class = self.get_paragraph_class()
        self.section_states, self.section_depth_first, self.section_depth \
            = self.get_section_data()
        self.raw_md_text = self._get_raw_md_text()
        self.length = self.get_length()
        self.length_sec = self.get_length_sec()
        self.length_spa = self.get_length_spa()
        self.length_ins = self.get_length_ins()
        self.first_line_instructions = self.get_first_line_instructions()
        self.md_text = self.get_md_text()

    def _get_raw_class(self):
        res = '^<(\\S+)( .*)?>$'
        rxlz = self.raw_xml_lines[0]
        if re.match(res, rxlz):
            return re.sub(res, '\\1', rxlz)
        else:
            return None

    def _get_xml_lines(self):
        size = Paragraph.font_size
        s_size = 0.9 * size  # not 0.8
        raw_xml_lines = self.raw_xml_lines
        xml_lines = []
        is_large = False
        is_small = False
        is_italic = False
        is_bold = False
        is_gothic = False
        has_strike = False
        has_underline = False
        color = ''
        is_in_text = False
        res_img_ms = \
            '^<v:imagedata r:id=[\'"](.+)[\'"] o:title=[\'"](.+)[\'"]/>$'
        res_img_py_name = \
            '^<pic:cNvPr id=[\'"](.+)[\'"] name=[\'"](.+)[\'"]/>$'
        res_img_py_id = \
            '^<a:blip r:embed=[\'"](.+)[\'"]/>$'
        for rxl in raw_xml_lines:
            if re.match(res_img_ms, rxl):
                # IMAGE MS WORD
                xml_lines.append(rxl)
                continue
            if re.match(res_img_py_name, rxl) or re.match(res_img_py_id, rxl):
                # IMAGE PYTHON-DOCX
                xml_lines.append(rxl)
                continue
            if re.match('^<w:r( .*)?>$', rxl):
                text = ''
                xml_lines.append(rxl)
                is_in_text = True
                continue
            if re.match('^</w:r>$', rxl):
                if is_gothic:
                    text = '`' + text + '`'
                    is_gothic = False
                if is_italic:
                    text = '*' + text + '*'
                    is_italic = False
                if is_bold:
                    text = '**' + text + '**'
                    is_bold = False
                if has_strike:
                    text = '~~' + text + '~~'
                    has_strike = False
                if is_small:
                    text = '--' + text + '--'
                    is_small = False
                if is_large:
                    text = '++' + text + '++'
                    is_large = False
                if color != '':
                    if color == 'FFFFFF':
                        text = '@@' + text + '@@'
                    else:
                        text = '@' + color + '@' + text + '@' + color + '@'
                    color = ''
                if has_underline:
                    text = '__' + text + '__'
                    has_underline = False
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
                if not self._is_table():
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
            elif re.match('^<w:color w:val="[0-9A-F]+"/?>$', rxl):
                color = re.sub('^<.*w:val="([0-9A-F]+)".*>$', '\\1', rxl, re.I)
                color = color.upper()
            elif re.match('^<w:br/?>$', rxl):
                text += '\n'
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
                text += rxl
        # self.xml_lines = xml_lines
        return xml_lines

    def _is_table(self):
        if self.raw_class != 'w:tbl':
            return False
        tbl_type = ''
        col = 0
        for rxl in self.raw_xml_lines:
            if re.match('<w:tblStyle w:val=[\'"].+[\'"]/>', rxl):
                return True
            if re.match('<w:gridCol w:w=[\'"][0-9]+[\'"]/>', rxl):
                col += 1
        if col != 3:
            return True
        return False

    def _get_raw_text(self):
        xml_lines = self.xml_lines
        raw_text = ''
        res_img_ms = \
            '^<v:imagedata r:id=[\'"](.+)[\'"] o:title=[\'"](.+)[\'"]/>$'
        res_img_py_name = \
            '^<pic:cNvPr id=[\'"](.+)[\'"] name=[\'"](.+)[\'"]/>$'
        res_img_py_id = \
            '^<a:blip r:embed=[\'"](.+)[\'"]/>$'
        for xl in xml_lines:
            if re.match(res_img_ms, xl):
                img_id = re.sub(res_img_ms, '\\1', xl)
                img_name = re.sub(res_img_ms, '\\2', xl)
                img_rel_name = doc.rels[img_id]
                img_ext = re.sub('^.*\\.', '', img_rel_name)
                img = img_name + '.' + img_ext
                doc.images[img_rel_name] = img
                raw_text += '![' + img + '](' + doc.media_dir + '/' + img + ')'
            if re.match(res_img_py_name, xl):
                img = re.sub(res_img_py_name, '\\2', xl)
                doc.images[''] = img
                raw_text += '![' + img + '](' + doc.media_dir + '/' + img + ')'
            if re.match(res_img_py_id, xl):
                img_id = re.sub(res_img_py_id, '\\1', xl)
                img_rel_name = doc.rels[img_id]
                doc.images[img_rel_name] = doc.images['']
            if re.match('^<.*>$', xl):
                continue
            while True:
                if re.match('^.*(\n.*)*\\+\\+$', raw_text) and \
                   re.match('^\\+\\+.*$', xl):
                    raw_text = re.sub('\\+\\+$', '', raw_text)
                    xl = re.sub('^\\+\\+', '', xl)
                    continue
                if re.match('^.*(\n.*)*--$', raw_text) and \
                   re.match('^--.*$', xl):
                    raw_text = re.sub('--$', '', raw_text)
                    xl = re.sub('^--', '', xl)
                    continue
                if re.match('^.*(\n.*)*[^\\*]\\*\\*\\*$', raw_text) and \
                   re.match('^\\*\\*\\*[^\\*].*$', xl):
                    raw_text = re.sub('\\*\\*\\*$', '', raw_text)
                    xl = re.sub('^\\*\\*\\*', '', xl)
                    continue
                if re.match('^.*(\n.*)*[^\\*]\\*\\*$', raw_text) and \
                   re.match('^\\*\\*[^\\*].*$', xl):
                    raw_text = re.sub('\\*\\*$', '', raw_text)
                    xl = re.sub('^\\*\\*', '', xl)
                    continue
                if re.match('^.*(\n.*)*[^\\*]\\*$', raw_text) and \
                   re.match('^\\*[^\\*].*$', xl):
                    raw_text = re.sub('\\*$', '', raw_text)
                    xl = re.sub('^\\*', '', xl)
                    continue
                if re.match('^.*(\n.*)*~~$', raw_text) and \
                   re.match('^~~.*$', xl):
                    raw_text = re.sub('~~$', '', raw_text)
                    xl = re.sub('^~~', '', xl)
                    continue
                if re.match('^.*(\n.*)*__$', raw_text) and \
                   re.match('^__.*$', xl):
                    raw_text = re.sub('__$', '', raw_text)
                    xl = re.sub('^__', '', xl)
                    continue
                if re.match('^.*(\n.*)*`$', raw_text) and \
                   re.match('^`.*$', xl):
                    raw_text = re.sub('`$', '', raw_text)
                    xl = re.sub('^`', '', xl)
                    continue
                if re.match('^.*(\n.*)*@[0-9A-F]*@$', raw_text) and \
                   re.match('^@[0-9A-F]*@.*$', xl):
                    ce = re.sub('^.*(?:\n.*)*(@[0-9A-F]*@)$', '\\1', raw_text)
                    cb = re.sub('^(@[0-9A-F]*@).*$', '\\1', xl)
                    if ce == cb:
                        raw_text = re.sub('@[0-9A-F]*@$', '', raw_text)
                        xl = re.sub('^@[0-9A-F]*@', '', xl)
                        continue
                else:
                    break
            raw_text += xl
        raw_text = raw_text.replace('&lt;', '<')
        raw_text = raw_text.replace('&gt;', '>')
        raw_text = raw_text.replace('&amp;', '&')
        # self.raw_text = raw_text
        return raw_text

    def separate_space(self):
        raw_text = self.raw_text
        beg_space = ''
        end_space = ''
        res = '^([ \t' + ZENKAKU_SPACE + ']+)(.*)$'
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

    def get_and_apply_style(self):
        style = None
        for rxl in self.raw_xml_lines:
            style = get_xml_value('w:pStyle', 'w:val', style, rxl)
        for ds in doc.styles:
            if style != ds.name:
                continue
            self.alignment = ds.alignment
            for s in self.length:
                if ds.raw_length[s] is not None:
                    self.length[s] = ds.raw_length[s]
        # self.style = style
        return style

    def get_alignment(self):
        alignment = None
        for rxl in self.raw_xml_lines:
            alignment = get_xml_value('w:jc', 'w:val', alignment, rxl)
            if alignment is not None:
                if not re.match('^((left)|(center)|(right))$', alignment):
                    alignment = None
        # self.alignment = alignment
        return alignment

    def get_paragraph_class(self):
        rt = self.raw_text
        stl = self.style
        aln = self.alignment
        size = Paragraph.font_size
        fs = -1.0
        for rxl in self.raw_xml_lines:
            if fs < 0:
                fs = get_xml_value('w:sz', 'w:val', fs, rxl) / size / 2
        td = Title.get_depth(rt, aln)
        if self.raw_class == 'w:sectPr':
            return 'configuration'
        if self.raw_class == 'w:tbl':
            col = 0
            for rxl in self.raw_xml_lines:
                if re.match('<w:tblStyle w:val=[\'"].+[\'"]/>', rxl):
                    return 'table'
                if re.match('<w:gridCol w:w=[\'"][0-9]+[\'"]/>', rxl):
                    col += 1
            if col != 3:
                return 'table'
            return 'breakdown'
        if rt == '':
            for rxl in self.raw_xml_lines:
                if re.match('^<w:drawing>$', rxl):
                    return 'image'
        if (td == 1 and fs > 1.2) or td > 1:
            return 'title'
        if stl is not None and stl == 'makdo-g':
            return 'preformatted'
        if (stl is not None and stl == 'makdo-a') or (aln is not None):
            return 'alignment'
        for rxl in self.raw_xml_lines:
            if re.match('<w:pStyle w:val=[\'"]ListBullet([0-9]?)[\'"]/>', rxl):
                return 'list'
            if re.match('<w:pStyle w:val=[\'"]ListNumber([0-9]?)[\'"]/>', rxl):
                return 'list'
            if re.match('<w:ilvl w:val=[\'"]([0-9]?)[\'"]/>', rxl):
                return 'list_system'
        if List.get_division_and_depth(rt) != '':
            return 'list'
        for rxl in self.raw_xml_lines:
            if re.match('^<w:br w:type="page"/>$', rxl):
                return 'pagebreak'
        if rt == '':
            return 'blank'
        return 'sentence'

    def get_section_data(self):
        rt = self.raw_text
        aln = self.alignment
        states = []
        depth_first = 0
        depth = 0
        for i, ss in enumerate(Paragraph.section_states):
            states.append(ss)
            if ss != 0:
                depth_first = i + 1
                depth = i + 1
        if self.paragraph_class == 'title':
            depth_first = 0
            depth = 0
            for i, ss in enumerate(Paragraph.section_states):
                dp = i + 1
                if Title.get_depth(rt, aln) == dp:
                    comm, numb, head, rt = Title.decompose(dp, rt)
                    for j in range(i + 1, len(states)):
                        states[j] = 0
                    states[i] = numb
                    if depth_first == 0:
                        depth_first = dp
                    depth = dp
        Paragraph.section_states = states
        # self.section_states = section_states
        # self.section_depth_first = section_depth_first
        # self.section_depth = section_depth
        return states, depth_first, depth

    def _get_raw_md_text(self):
        if self.paragraph_class == 'title':
            return self._get_raw_md_text_of_title_paragraph()
        elif self.paragraph_class == 'list_system':
            return self._get_raw_md_text_of_list_system_paragraph()
        elif self.paragraph_class == 'list':
            return self._get_raw_md_text_of_list_paragraph()
        elif self.paragraph_class == 'alignment':
            return self._get_raw_md_text_of_alignment_paragraph()
        elif self.paragraph_class == 'table':
            return self._get_raw_md_text_of_table_paragraph()
        elif self.paragraph_class == 'image':
            return self._get_raw_md_text_of_image_paragraph()
        elif self.paragraph_class == 'preformatted':
            return self._get_raw_md_text_of_preformatted_paragraph()
        elif self.paragraph_class == 'breakdown':
            return self._get_raw_md_text_of_breakdown_paragraph()
        elif self.paragraph_class == 'pagebreak':
            return '<div style="break-after: page;"></div>'
        return self.raw_text

    def _get_raw_md_text_of_title_paragraph(self):
        rt = self.raw_text
        aln = self.alignment
        head = ''
        for i in range(len(Paragraph.section_states)):
            dp = i + 1
            if Title.get_depth(rt, aln) == dp:
                c, n, h, rt = Title.decompose(dp, rt)
                head += c + '#' * dp + ' '
        if not re.match('^.*[．。]$', rt):
            raw_md_text = head + rt
        else:
            raw_md_text = head + '\n' + rt
        # self.md_text = md_text
        return raw_md_text

    def _get_raw_md_text_of_list_system_paragraph(self):
        typ = ''
        dep = 1
        for rxl in self.raw_xml_lines:
            res = '^<w:pStyle w:val=[\'"]ListBullet([0-9]?)[\'"]/>$'
            if re.match(res, rxl):
                typ = 'bullet'
                n = re.sub(res, '\\1', rxl)
                if n != '':
                    dep = int(n)
            res = '^<w:pStyle w:val=[\'"]ListNumber([0-9]?)[\'"]/>$'
            if re.match(res, rxl):
                typ = 'number'
                n = re.sub(res, '\\1', rxl)
                if n != '':
                    dep = int(n)
            res = '^<w:ilvl w:val=[\'"]([0-9]+)[\'"]/>$'
            if re.match(res, rxl):
                n = re.sub(res, '\\1', rxl)
                dep = int(n) + 1
            res = '^<w:numId w:val=[\'"]([0-9]+)[\'"]/>$'
            if re.match(res, rxl):
                n = re.sub(res, '\\1', rxl)
                if n == '10':
                    typ = 'bullet'
                else:
                    typ = 'number'
        if typ == '':
            return ''
        raw_md_text = '  ' * (dep - 1)
        rt = self.raw_text
        if typ == 'bullet':
            raw_md_text += '- ' + rt
        else:
            raw_md_text += '1. ' + rt
        # self.raw_md_text = raw_md_text
        return raw_md_text

    def _get_raw_md_text_of_list_paragraph(self):
        raw_md_text = ''
        for ln in self.raw_text.split('\n'):
            dd = List.get_division_and_depth(ln)
            numb, head, text = List.decompose(dd, ln)
            if dd == 'b1':
                raw_md_text += '\n' + ' ' * 0 + '- ' + text
            elif dd == 'b2':
                raw_md_text += '\n' + ' ' * 2 + '- ' + text
            elif dd == 'b3':
                raw_md_text += '\n' + ' ' * 4 + '- ' + text
            elif dd == 'b4':
                raw_md_text += '\n' + ' ' * 6 + '- ' + text
            elif dd == 'n1':
                raw_md_text += '\n' + ' ' * 0 + '1. ' + text
            elif dd == 'n2':
                raw_md_text += '\n' + ' ' * 2 + '1. ' + text
            elif dd == 'n3':
                raw_md_text += '\n' + ' ' * 4 + '1. ' + text
            elif dd == 'n4':
                raw_md_text += '\n' + ' ' * 6 + '1. ' + text
            else:
                raw_md_text += rt
        raw_md_text = re.sub('^\n', '', raw_md_text)
        # self.raw_md_text = raw_md_text
        return raw_md_text

    def _get_raw_md_text_of_alignment_paragraph(self):
        aln = self.alignment
        raw_md_text = ''
        for ln in self.raw_text.split('\n'):
            if ln == '':
                continue
            if aln == 'right':
                raw_md_text += ln + ' :\n'
            elif aln == 'center':
                raw_md_text += ': ' + ln + ' :\n'
            else:
                raw_md_text += ': ' + ln + '\n'
        raw_md_text = re.sub('\n$', '', raw_md_text)
        # self.raw_md_text = raw_md_text
        return raw_md_text

    def _get_raw_md_text_of_table_paragraph(self):
        s_size = 0.8 * self.font_size
        is_in_row = False
        is_in_cel = False
        tab = []
        wid = []
        for xl in self.xml_lines:
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
        raw_md_text = ''
        half_row = int(len(tab) / 2)
        is_in_head = True
        for i, row in enumerate(tab):
            if is_in_head:
                if ali[i] == ali[half_row]:
                    for cell in ali[half_row]:
                        raw_md_text += '|' + cell + '|'
                    is_in_head = False
                    raw_md_text += '\n'
            for cell in row:
                tmp = ''
                for lin in cell:
                    if not re.match('<.*>', lin):
                        tmp += lin
                tmp = re.sub('\n', '<br>', tmp)
                raw_md_text += '|' + tmp + '|'
            raw_md_text += '\n'
        raw_md_text = raw_md_text.replace('||', '|')
        raw_md_text = raw_md_text.replace('&lt;', '<')
        raw_md_text = raw_md_text.replace('&gt;', '>')
        raw_md_text = re.sub('\n$', '', raw_md_text)
        # self.raw_md_text = raw_md_text
        return raw_md_text

    def _get_raw_md_text_of_image_paragraph(self):
        image = ''
        res = '^<pic:cNvPr id=[\'"].+[\'"] name=[\'"](.*)[\'"]/>$'
        for rxl in self.raw_xml_lines:
            if re.match(res, rxl):
                image = re.sub(res, '\\1', rxl)
        raw_md_text = '![' + image + '](' + image + ')'
        # self.raw_md_text = raw_md_text
        return raw_md_text

    def _get_raw_md_text_of_preformatted_paragraph(self):
        raw_text = self.raw_text
        raw_text = re.sub('^`', '', raw_text)
        raw_text = re.sub('`$', '', raw_text)
        res = '^(\\s+)\\[(.*)\\]'
        if re.match(res, raw_text):
            raw_text = re.sub(res, '\\1\\2', raw_text)
        else:
            raw_text = '\n' + raw_text
        raw_md_text = '``` ' + raw_text + '\n```'
        # self.raw_md_text = raw_md_text
        return raw_md_text

    def _get_raw_md_text_of_breakdown_paragraph(self):
        size = self.font_size
        is_in_row = False
        is_in_cel = False
        tab = []
        wid = []
        for xl in self.xml_lines:
            res = '^<w:gridCol w:w=[\'"]([0-9]+)[\'"]/>$'
            if re.match(res, xl):
                w = round((float(re.sub(res, '\\1', xl)) / size / 10) - 4)
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
        raw_md_text = ''
        for i, row in enumerate(tab):
            if len(wid) > 0:
                is_header = True
                for j, cell in enumerate(row):
                    for xml in cell:
                        if re.match('<w:jc w:val=[\'"]left[\'"]/>', xml):
                            is_header = False
                        if re.match('<w:jc w:val=[\'"]rigth[\'"]/>', xml):
                            is_header = False
                if not is_header:
                    for w in wid:
                        raw_md_text += ':' + '-' * w + ':'
                    raw_md_text += '\n'
                    wid = []
            for j, cell in enumerate(row):
                tmp = ''
                for lin in cell:
                    if not re.match('<.*>', lin):
                        tmp += lin
                tmp = re.sub('\n', '<br>', tmp)
                if j == 0:
                    dd = List.get_division_and_depth(tmp)
                    if re.match('^b', dd):
                        numb, head, text = List.decompose(dd, tmp)
                        tmp = '- ' + text
                    elif re.match('^n', dd):
                        numb, head, text = List.decompose(dd, tmp)
                        tmp = '1. ' + text
                    else:
                        tmp = re.sub('^' + ZENKAKU_SPACE, '', tmp)
                raw_md_text += ':' + tmp + ':'
            raw_md_text += '\n'
        raw_md_text = re.sub('^:', '', raw_md_text)
        raw_md_text = re.sub('(::)?:\n:', '\n', raw_md_text)
        raw_md_text = re.sub('(::)?:$', '', raw_md_text)
        raw_md_text = raw_md_text.replace('&lt;', '<')
        raw_md_text = raw_md_text.replace('&gt;', '>')
        raw_md_text = re.sub('\n$', '', raw_md_text)
        # self.raw_md_text = raw_md_text
        return raw_md_text

    def get_length(self):
        size = Paragraph.font_size
        lnsp = doc.line_spacing
        length \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        sb_x = 0.0
        sa_x = 0.0
        ls_x = 0.0
        fi_x = 0.0
        hi_x = 0.0
        li_x = 0.0
        ri_x = 0.0
        ti_x = 0.0
        for rxl in self.raw_xml_lines:
            sb_x = get_xml_value('w:spacing', 'w:before', sb_x, rxl)
            sa_x = get_xml_value('w:spacing', 'w:after', sa_x, rxl)
            ls_x = get_xml_value('w:spacing', 'w:line', ls_x, rxl)
            fi_x = get_xml_value('w:ind', 'w:firstLine', fi_x, rxl)
            hi_x = get_xml_value('w:ind', 'w:hanging', hi_x, rxl)
            li_x = get_xml_value('w:ind', 'w:left', li_x, rxl)
            ri_x = get_xml_value('w:ind', 'w:right', ri_x, rxl)
            ti_x = get_xml_value('w:tblInd', 'w:w', ti_x, rxl)
        length['space before'] = sb_x / 20 / size / lnsp
        length['space after'] = sa_x / 20 / size / lnsp
        ls = 0.0
        if ls_x > 0.0:
            ls = (ls_x / 20 / size / lnsp) - 1
        length['space before'] += ls * .75
        length['space after'] += ls * .25
        # if round(ls, 1) < 0:
        #     length['space before'] += ls * .75
        #     length['space after'] += ls * .25
        # elif round(ls, 1) > 0:
        #     length['space before'] -= ls * .25
        #     length['space after'] -= ls * .75
        length['line spacing'] = ls
        length['first indent'] = (fi_x - hi_x) / 20 / size
        length['left indent'] = (li_x + ti_x) / 20 / size
        length['right indent'] = ri_x / 20 / size
        # self.length = length
        return length

    def get_length_sec(self):
        depth_first = self.section_depth_first
        depth = self.section_depth
        states = self.section_states
        pclass = self.paragraph_class
        length_sec \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        if pclass == 'title':
            if depth_first > 1:
                length_sec['first indent'] = depth_first - depth - 1
                length_sec['left indent'] = depth - 1
            if depth_first > 2 and states[1] == 0:
                length_sec['left indent'] -= 1
            if doc.document_style == 'j':
                if self.section_depth >= 3:
                    length_sec['left indent'] -= 1
        elif re.match('^((list_system)|(list)|(breakdown))$', pclass):
            length_sec['first indent'] = 0
            if depth_first > 1:
                length_sec['left indent'] = depth - 1
            if depth_first == 3 and states[1] == 0:
                length_sec['left indent'] -= 1
            if doc.document_style == 'j':
                if self.section_depth >= 3:
                    length_sec['left indent'] -= 1
        elif pclass == 'sentence':
            if depth_first > 0:
                length_sec['first indent'] = 1
            if depth_first > 1:
                length_sec['left indent'] = depth - 1
            if depth_first == 3 and states[1] == 0:
                length_sec['left indent'] -= 1
            if doc.document_style == 'j':
                if self.section_depth >= 3:
                    length_sec['left indent'] -= 1
        if pclass == 'breakdown':
            length_sec['first indent'] += 1
        # self.length_sec = length_sec
        return length_sec

    def get_length_spa(self):
        length_spa \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        length_spa['first indent'] = float(get_real_width(self.beg_space)) / 2
        # self.length_spa = length_spa
        return length_spa

    def get_length_ins(self):
        length_ins \
            = {'space before': 0.0, 'space after': 0.0, 'line spacing': 0.0,
               'first indent': 0.0, 'left indent': 0.0, 'right indent': 0.0}
        for s in length_ins:
            length_ins[s] \
                = self.length[s] - self.length_sec[s] + self.length_spa[s]
        if self.paragraph_class == 'title':
            d = self.section_depth_first
            sb = (doc.space_before + ',,,,,').split(',')[d - 1]
            if sb != '':
                length_ins['space before'] -= float(sb)
            d = self.section_depth
            sa = (doc.space_after + ',,,,,').split(',')[d - 1]
            if sa != '':
                length_ins['space after'] -= float(sa)
        # self.length_ins = length_ins
        return length_ins

    def get_first_line_instructions(self):
        sb = self.length_ins['space before']
        sa = self.length_ins['space after']
        ls = self.length_ins['line spacing']
        fi = self.length_ins['first indent']
        li = self.length_ins['left indent']
        ri = self.length_ins['right indent']
        if self.paragraph_class == 'alignment':
            if self.alignment == 'center':
                if li > 0:
                    fi = li / 2
                else:
                    fi = ri / 2
            if self.alignment == 'right':
                fi = -self.length_ins['right indent']
        rsb = round(sb, 1)
        rsa = round(sa, 1)
        rls = round(ls, 1)
        # rsb = float(round(sb * 2)) / 2
        # rsa = float(round(sa * 2)) / 2
        # rls = float(round(ls * 2)) / 2
        rfi = float(round(fi * 2)) / 2
        rli = float(round(li * 2)) / 2
        ins = ''
        if rsb < 0:
            ins += 'v=' + str(rsb) + ' '
        elif rsb > 0:
            ins += 'v=+' + str(rsb) + ' '
        if rsa < 0:
            ins += 'V=' + str(rsa) + ' '
        elif rsa > 0:
            ins += 'V=+' + str(rsa) + ' '
        if rls < 0:
            ins += 'X=' + str(rls) + ' '
        if rls > 0:
            ins += 'X=+' + str(rls) + ' '
        if rfi != 0:
            ins += '<<=' + str(-rfi) + ' '
        if rli != 0:
            ins += '<=' + str(-rli) + ' '
        ins = re.sub(' $', '', ins)
        first_line_instructions = ins
        # self.first_line_instructions = ins
        return first_line_instructions

    def get_md_text(self):
        rmt = self.raw_md_text
        if self.paragraph_class == 'title':
            if len(rmt.split('\n')) == 1:
                return rmt
        if self.paragraph_class == 'list':
            return rmt
        if self.paragraph_class == 'alignment':
            return rmt
        if self.paragraph_class == 'table':
            return rmt
        if self.paragraph_class == 'image':
            return rmt
        if self.paragraph_class == 'preformatted':
            return rmt
        if self.paragraph_class == 'pagebreak':
            return rmt
        if self.paragraph_class == 'title':
            res = '^((#+ \n)*#+ )(#+ .*)$'
            while re.match(res, rmt):
                rmt = re.sub(res, '\\1\n\\3', rmt)
        md_text = ''
        for line in rmt.split('\n'):
            splited = self._split_by_punctuate(line)
            md_text += self._reconcatenate(splited)
            splited = []
        md_text = re.sub('\n$', '', md_text)
        return md_text

    @staticmethod
    def _split_by_punctuate(line):
        splited = []
        tmp = ''
        m = len(line) - 1
        for i, c in enumerate(line):
            tmp += c
            res = '^[^\\(\\[]$'
            if re.match('^ $', line[i]):
                if (i == m) or re.match(res, line[i + 1]):
                    splited.append(tmp)
                    tmp = ''
            res = '^[^（「]$'
            if re.match(res, line[i]):
                if (i == m) or (not re.match(res, line[i + 1])):
                    splited.append(tmp)
                    tmp = ''
            res = '^[,\\.\\)\\]]$'
            if re.match(res, line[i]):
                if (i == m) or re.match('^ $', line[i + 1]):
                    splited.append(tmp)
                    tmp = ''
            res = '^[，、．。）」]$'
            if re.match(res, line[i]):
                if ((i == m) or (not re.match(res, line[i + 1]))) and \
                   ((i > 0) and not re.match('^[０-９]$', line[i - 1])) and \
                   ((i < m) and not re.match('^[０-９]$', line[i + 1])):
                    splited.append(tmp)
                    tmp = ''
        if tmp != '':
            splited.append(tmp + '\n')
            tmp = ''
        return splited

    @staticmethod
    def _reconcatenate(splited):
        tex = ''
        tmp = ''
        for p in splited:
            if get_ideal_width(tmp + p) > MD_TEXT_WIDTH:
                if tmp != '':
                    tex += tmp + '\n'
                    tmp = ''
            tmp += p
            if get_ideal_width(tmp) <= MD_TEXT_WIDTH:
                if re.match('^.*[．。]$', tmp):
                    if tmp != '':
                        tex += tmp + '\n'
                        tmp = ''
            else:
                while get_ideal_width(tmp) > MD_TEXT_WIDTH:
                    for i in range(len(tmp), -1, -1):
                        s1 = tmp[:i]
                        s2 = tmp[i:]
                        if get_ideal_width(s1) > MD_TEXT_WIDTH:
                            continue
                        if not re.match('^.*[ぁ-ん。]$', s1):
                            continue
                        if not re.match('^[^ぁ-ん。].*$', s2):
                            continue
                        if s1 != '':
                            tex += s1 + '\n'
                            tmp = s2
                            break
                    else:
                        for i in range(len(tmp), -1, -1):
                            s1 = tmp[:i]
                            s2 = tmp[i:]
                            if re.match('.*[\\\\*~_/+-]$', s1):
                                continue
                            if re.match('^[\\\\*~_/+-].*', s2):
                                continue
                            if get_ideal_width(s1) < MD_TEXT_WIDTH:
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
        return tex

    def write_md_lines(self, mf):
        pc = self.paragraph_class
        fli = self.first_line_instructions
        mt = self.md_text
        n = self.paragraph_number - 1
        m = len(doc.paragraphs) - 1
        if n > 0:
            p_prev = doc.paragraphs[n - 1]
        if n < m:
            p_next = doc.paragraphs[n + 1]
        if pc == 'empty':
            return
        if pc == 'blank':
            for i in range(n + 1, m + 1):
                if doc.paragraphs[i].paragraph_class != 'blank' and \
                   doc.paragraphs[i].paragraph_class != 'configuration':
                    break
            else:
                return
        if pc == 'configuration':
            return
        if pc == 'list_system':
            text_to_write = ''
            if n == 0 or p_prev.paragraph_class != 'list_system':
                if fli != '':
                    text_to_write += fli + '\n'
            text_to_write += mt + '\n'
            if n == m or p_next.paragraph_class != 'list_system':
                text_to_write += '\n'
            mf.write(text_to_write)
            return
        if re.match('^\\s*(#+|v|V|X|<<|<)=\\s*[0-9]+', mt):
            mt = '<!---->' + mt
        if mt == '':
            text_to_text = '  \n\n'
        elif fli == '':
            text_to_text = mt + '\n\n'
        else:
            text_to_text = fli + '\n' + mt + '\n\n'
        mf.write(text_to_text)


############################################################
# MAIN


if __name__ == '__main__':

    args = get_arguments()

    doc = Document()

    doc.tmpdir = doc.make_tmpdir()

    doc.media_dir = doc.get_media_dir_name(args.md_file, args.docx_file)

    doc.extract_docx_file(args.docx_file)

    doc.core_raw_xml_lines = doc.get_raw_xml_lines('/docProps/core.xml')
    doc.footer1_raw_xml_lines = doc.get_raw_xml_lines('/word/footer1.xml')
    doc.styles_raw_xml_lines = doc.get_raw_xml_lines('/word/styles.xml')
    doc.rels_raw_xml_lines \
        = doc.get_raw_xml_lines('/word/_rels/document.xml.rels')
    doc.document_raw_xml_lines = doc.get_raw_xml_lines('/word/document.xml')

    doc.configure(args)

    doc.styles = doc.get_styles(doc.styles_raw_xml_lines)
    doc.rels = doc.get_rels(doc.rels_raw_xml_lines)

    doc.raw_paragraphs = doc.get_raw_paragraphs(doc.document_raw_xml_lines)
    doc.paragraphs = doc.get_paragraphs(doc.raw_paragraphs)
    doc.paragraphs = doc.modify_paragraphs()

    doc.check_section_consistency()

    mf = doc.open_md_file(args.md_file, args.docx_file)
    doc.write_configurations(mf)
    doc.write_md_lines(mf)

    doc.make_media_dir(doc.media_dir)

    sys.exit(0)
