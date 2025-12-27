#!/usr/bin/python3
# Name:         jpchars.py
# Version:      v01
# Time-stamp:   <2025.12.22-10:53:36-JST>

# jpchars.py
# Copyright (C) 2025  Seiichiro HATA
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

# EXAMPLE
#
# >>> import jpchars
# >>> c = jpchars.chars['應']
# >>> c.char
# '應'
# >>> c.id
# '01337'
# >>> c.sub_id
# 'D0108'
# >>> c.utf8
# '61C9'
# >>> c.char_class
# '常用'
# >>> c.level
# '第二水準'
# >>> c.width
# '全角'
# >>> c.stroke_count
# 17
# >>> c.bushu
# 'こころ・りっしんべん・したごころ(4)'
# >>> c.school_grade
# ''（←義務教育では習わないため空）
# >>> c.pronunciation
# オウ{応答、応用、呼応}/こた-える{応える}⇔答える
# >>> c.usage
# '「反応」、「順応」などは、「ハンノウ」、「ジュンノウ」。'
# >>> c.notes
# '「応」の康熙字典体'
# >>> c.group
# ['応', '應']
#
# >>> jpchars.convert_half_to_full('123ABCabcｱｲｳ')
# '１２３ＡＢＣａｂｃアイウ'
#
# >>> jpchars.convert_full_to_half('１２３ＡＢＣａｂｃアイウ')
# '123ABCabcｱｲｳ'
#
# >>> jpchars.adjust_width('１２３ＡＢＣａｂｃアイウ')
# '123ABCabcアイウ'


import sys
import re


HALF_FULL_TABLE = [
    [' ', '\u3000'],
    ['!', '！'], ['"', '”'], ['#', '＃'], ['$', '＄'], ['%', '％'],
    ['&', '＆'], ["'", '’'], ['(', '（'], [')', '）'], ['*', '＊'],
    ['+', '＋'], [',', '，'], ['-', '－'], ['.', '．'], ['/', '／'],
    ['0', '０'], ['1', '１'], ['2', '２'], ['3', '３'], ['4', '４'],
    ['5', '５'], ['6', '６'], ['7', '７'], ['8', '８'], ['9', '９'],
    [':', '：'], [';', '；'], ['<', '＜'], ['=', '＝'], ['>', '＞'],
    ['?', '？'], ['@', '＠'],
    ['A', 'Ａ'], ['B', 'Ｂ'], ['C', 'Ｃ'], ['D', 'Ｄ'], ['E', 'Ｅ'],
    ['F', 'Ｆ'], ['G', 'Ｇ'], ['H', 'Ｈ'], ['I', 'Ｉ'], ['J', 'Ｊ'],
    ['K', 'Ｋ'], ['L', 'Ｌ'], ['M', 'Ｍ'], ['N', 'Ｎ'], ['O', 'Ｏ'],
    ['P', 'Ｐ'], ['Q', 'Ｑ'], ['R', 'Ｒ'], ['S', 'Ｓ'], ['T', 'Ｔ'],
    ['U', 'Ｕ'], ['V', 'Ｖ'], ['W', 'Ｗ'], ['X', 'Ｘ'], ['Y', 'Ｙ'],
    ['Z', 'Ｚ'],
    ['[', '［'], ['\\', '＼'], [']', '］'], ['^', '＾'], ['_', '＿'],
    ['`', '｀'],
    ['a', 'ａ'], ['b', 'ｂ'], ['c', 'ｃ'], ['d', 'ｄ'], ['e', 'ｅ'],
    ['f', 'ｆ'], ['g', 'ｇ'], ['h', 'ｈ'], ['i', 'ｉ'], ['j', 'ｊ'],
    ['k', 'ｋ'], ['l', 'ｌ'], ['m', 'ｍ'], ['n', 'ｎ'], ['o', 'ｏ'],
    ['p', 'ｐ'], ['q', 'ｑ'], ['r', 'ｒ'], ['s', 'ｓ'], ['t', 'ｔ'],
    ['u', 'ｕ'], ['v', 'ｖ'], ['w', 'ｗ'], ['x', 'ｘ'], ['y', 'ｙ'],
    ['z', 'ｚ'],
    ['{', '｛'], ['|', '｜'], ['}', '｝'], ['~', '〜'],
    #
    ['ｳﾞ', 'ヴ'],
    ['ｶﾞ', 'ガ'], ['ｷﾞ', 'ギ'], ['ｸﾞ', 'グ'], ['ｹﾞ', 'ゲ'], ['ｺﾞ', 'ゴ'],
    ['ｻﾞ', 'ザ'], ['ｼﾞ', 'ジ'], ['ｽﾞ', 'ズ'], ['ｾﾞ', 'ゼ'], ['ｿﾞ', 'ゾ'],
    ['ﾀﾞ', 'ダ'], ['ﾁﾞ', 'ヂ'], ['ﾂﾞ', 'ヅ'], ['ﾃﾞ', 'デ'], ['ﾄﾞ', 'ド'],
    ['ﾊﾞ', 'バ'], ['ﾋﾞ', 'ビ'], ['ﾌﾞ', 'ブ'], ['ﾍﾞ', 'ベ'], ['ﾎﾞ', 'ボ'],
    ['ﾊﾟ', 'パ'], ['ﾋﾟ', 'ピ'], ['ﾌﾟ', 'プ'], ['ﾍﾟ', 'ペ'], ['ﾎﾟ', 'ポ'],
    ['ﾜﾞ', 'ヷ'], ['ｦﾞ', 'ヺ'],
    ['｡', '。'], ['｢', '「'], ['｣', '」'], ['､', '、'], ['･', '・'],
    ['ｦ', 'ヲ'],
    ['ｧ', 'ァ'], ['ｨ', 'ィ'], ['ｩ', 'ゥ'], ['ｪ', 'ェ'], ['ｫ', 'ォ'],
    ['ｬ', 'ャ'], ['ｭ', 'ュ'], ['ｮ', 'ョ'], ['ｯ', 'ッ'], ['ｰ', 'ー'],
    ['ｱ', 'ア'], ['ｲ', 'イ'], ['ｳ', 'ウ'], ['ｴ', 'エ'], ['ｵ', 'オ'],
    ['ｶ', 'カ'], ['ｷ', 'キ'], ['ｸ', 'ク'], ['ｹ', 'ケ'], ['ｺ', 'コ'],
    ['ｻ', 'サ'], ['ｼ', 'シ'], ['ｽ', 'ス'], ['ｾ', 'セ'], ['ｿ', 'ソ'],
    ['ﾀ', 'タ'], ['ﾁ', 'チ'], ['ﾂ', 'ツ'], ['ﾃ', 'テ'], ['ﾄ', 'ト'],
    ['ﾅ', 'ナ'], ['ﾆ', 'ニ'], ['ﾇ', 'ヌ'], ['ﾈ', 'ネ'], ['ﾉ', 'ノ'],
    ['ﾊ', 'ハ'], ['ﾋ', 'ヒ'], ['ﾌ', 'フ'], ['ﾍ', 'ヘ'], ['ﾎ', 'ホ'],
    ['ﾏ', 'マ'], ['ﾐ', 'ミ'], ['ﾑ', 'ム'], ['ﾒ', 'メ'], ['ﾓ', 'モ'],
    ['ﾔ', 'ヤ'], ['ﾕ', 'ユ'], ['ﾖ', 'ヨ'],
    ['ﾗ', 'ラ'], ['ﾘ', 'リ'], ['ﾙ', 'ル'], ['ﾚ', 'レ'], ['ﾛ', 'ロ'],
    ['ﾜ', 'ワ'], ['ﾝ', 'ン'],
    ['ﾞ', '゛'], ['ﾟ', '゜']]


def convert_half_to_full(half_str):
    full_str = half_str
    for hf in HALF_FULL_TABLE:
        full_str = full_str.replace(hf[0], hf[1])
    return full_str


def convert_full_to_half(full_str):
    half_str = full_str
    for hf in HALF_FULL_TABLE:
        half_str = half_str.replace(hf[1], hf[0])
    return half_str


def adjust_width(old_str):
    new_str = old_str
    # NUMBER
    for i in range(0, 10):
        new_str = new_str.replace(chr(i + 65296), chr(i + 48))
    new_str = re.sub('＋([0-9]+)', '+\\1', new_str)
    new_str = re.sub('－([0-9]+)', '-\\1', new_str)
    new_str = re.sub('([0-9]+)．([0-9]+)', '\\1.\\2', new_str)
    new_str = re.sub('([0-9]+)，([0-9]{3})', '\\1,\\2', new_str)
    new_str = re.sub('([0-9]+)，([0-9]{3})', '\\1,\\2', new_str)
    # ALPHABET
    for i in range(0, 26):
        new_str = new_str.replace(chr(i + 65313), chr(i + 65))
        new_str = new_str.replace(chr(i + 65345), chr(i + 97))
    # COMMA
    new_str = new_str.replace('，', '、')
    return new_str


class Char:

    char_class = {'S': '記号',
                  'I': '数字',
                  'E': '英字',
                  'K': '仮名',
                  'B': '部首',
                  'D': '常用',
                  'N': '人名',
                  'M': 'ＭＪ',
                  }

    level = {'1': '第一水準',
             '2': '第二水準',
             '3': '第三水準',
             '4': '第四水準',
             }

    width = {'H': '半角',
             '': '全角',
             }

    school_grade = {'': '',
                    '1': '小一',
                    '2': '小二',
                    '3': '小三',
                    '4': '小四',
                    '5': '小五',
                    '6': '小六',
                    'J': '中学',
                    }

    def __init__(self):
        self.id = '00000'            # 0 [0-9]{5}
        self.char = ''               # 1 .|[0-9A-F]{4,5}_E01[0-9A-F]{2}
        self.char_class = ''         # 2 記号|数字|英字|仮名|常用|人名|ＭＪ
        self.sub_id = ''             # 3 .{2}[0-9]{4}
        self.level = ''              # 4 第(一|二|三)水準
        self.width = ''              # 5 (半|全)角
        self.stroke_count = -1       # 6 (-1|[1-9][0-9]*)
        self.bushu = ''              # 7 .*([1-9][0-9]*)
        self.school_grade = ''       # 8 小一|小二|小三|小四|小五|小六|中学
        self.group = ''              # 9 [0-9]+
        self.pronunciation = ''      # 10 .*{.*、.*、...}/.*{.*、.*、...}/...
        self.usage = ''              # 11 .*
        self.notes = ''              # 12 .*

    def get_utf8_code(self):
        s = self.char
        utf8 = ''
        if len(s) == 1:
            # ONE CHAR
            utf8 = hex(ord(s))[2:].upper()
        elif len(s) == 2 and (s[1] == '' or s[1] == ''):
            # COMPOUND KATAKANA
            utf8 = hex(ord(s[0]))[2:].upper() + ',' \
                + hex(int(s[1:-1]))[2:].upper()
        elif re.match('^.[0-9]+;$', s):
            # MJ CODE
            utf8 = hex(ord(s[0]))[2:].upper() + ',' \
                + 'E01' + hex(int(s[1:-1]))[2:].upper().zfill(2)
        else:
            # JUST IN CASE
            for c in s:
                utf8 += hex(ord(c))[2:].upper() + ','
            utf8 = utf8[:-1]
        return utf8


class Chars:

    def __init__(self):
        lines = self._read_csv_file()
        self.bushus, self.chars, self.groups = self._get_data(lines)

    @staticmethod
    def _read_csv_file() -> list:
        try:
            csv_file = re.sub('\\.py$', '.csv', __file__)
            with open(csv_file, 'r') as f:
                lines = f.readlines()
            return lines
        except BaseException:
            sys.stderr.write("error: can't read a csv file\n")
            return None

    def _get_data(self, lines):
        bushus, chars, groups = {}, {}, {}
        if lines is None:
            return bushus, groups, chars
        line_class = 'bushus'
        for line in lines:
            line = line.rstrip()
            if line == '':
                continue
            if line[0] == '#':
                if line == '# [chars]':
                    line_class = 'chars'
                elif line == '# [groups]':
                    line_class = 'groups'
                continue
            if line_class == 'bushus':
                k, d = self._get_bushu_data(line)
                bushus[k] = d
            elif line_class == 'chars':
                k, d = self._get_char_data(line)
                chars[k] = d
            elif line_class == 'groups':
                k, d = self._get_group_data(line)
                groups[k] = d
        return bushus, chars, groups

    @staticmethod
    def _get_bushu_data(line) -> dict:
        bs = line.split(',')
        k, b = bs[0], bs[1]
        if k[0] == '0':
            if k[1] == '0':
                k = k[2:]
            else:
                k = k[1:]
        return k, b

    @staticmethod
    def _get_char_data(line) -> dict:
        char = Char()
        cs = line.split(',')
        #
        char.id = cs[0]
        #
        if cs[1] == '""""':
            cs[1] = '"'  # "
        elif cs[1] == '"' and cs[2] == '"':
            cs[1] = ','  # ,
            cs.pop(2)
        elif cs[1] == '"#"':
            cs[1] = '#'  # #
        char.char = cs[1]
        #
        char.sub_id = cs[2]
        #
        if cs[2] != '':
            char.char_class = Char.char_class[cs[2][0]]
        #
        if cs[3] != '':
            char.level = Char.level[cs[3][0]]
        #
        char.width = Char.width[cs[4]]
        #
        if cs[5] != '':
            char.stroke_count = int(cs[5])
        #
        char.school_grade = Char.school_grade[cs[6]]
        #
        if cs[7] != '':
            char.bushu = cs[7]
        #
        if cs[8] != '':
            char.group = cs[8]
        #
        if cs[9] != '':
            char.pronunciation = cs[9]
        #
        if cs[10] != '':
            char.usage = cs[10]
        #
        if cs[11] != '':
            char.notes = cs[11]
        #
        return cs[1], char

    @staticmethod
    def _get_group_data(line) -> list:
        ds = line.split(',')
        k = ds[0]
        if k[0] == '0':
            if k[1] == '0':
                if k[2] == '0':
                    k = k[3:]
                else:
                    k = k[2:]
            else:
                k = k[1:]
        if ds[1][0] == '"':
            if len(ds) == 3:
                # 「"」 + 「/，"」
                ds[1] += ',' + ds[2]
                # ds.pop(2)
            if ds[1][-1] == '"':
                if ds[1][1] != '"':
                    # 「",/，"」, 「"#/＃"」
                    ds[1] = ds[1][1:-1]
                else:
                    # 「"""/″/“/”"」
                    ds[1] = ds[1][2:-1]
        g = ds[1].split('/')
        if g[0] == '' and g[1] == '':
            g[0] = '/'
            g.pop(1)
        return k, g

    @staticmethod
    def _apply_groups_to_chars(groups, chars) -> dict:
        for c in chars:
            for g in groups:
                if c in g or '@' + c in g:
                    if len(chars[c].group) > 0:
                        m = 'warning: "' \
                            + c + '" is a member of "' + ','.join(g) + '"\n'
                        sys.stderr.write(m)
                    else:
                        chars[c].group = g
                    break
        return chars

    def is_joyokanji(self):
        if self.char_class == '常用':
            return True
        return False

    def is_jimmeikanji(self):
        if self.char_class == '人名':
            return True
        return False


cs = Chars()
bushus = cs.bushus
chars = cs.chars
groups = cs.groups
