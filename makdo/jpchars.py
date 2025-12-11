#!/usr/bin/python3
# Name:         jpchars.py
# Version:      v01
# Time-stamp:   <2025.12.11-12:27:39-JST>

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

    def __init__(self):
        self.id = '00000'            # [0-9]{5}
        self.char = ''               # .|[0-9A-F]{4,5}_E01[0-9A-F]{2}
        self.utf8 = ''               # [0-9]{2,5}
        self.char_class = ''         # 記号|英数|仮名|常用|人名|ＭＪ
        self.sub_id = ''             # .{2}[0-9]{4}
        self.level = ''              # 第(一|二|三)水準
        self.width = ''              # (半|全)角
        self.stroke_count = -1       # (-1|[1-9][0-9]*)
        self.bushu = ''              # .*([1-9][0-9]*)
        self.school_grade = ''       # 小一|小二|小三|小四|小五|小六|中学
        self.pronunciation = ''      # .*{.*、.*、...}/.*{.*、.*、...}/...
        self.usage = ''              # .*
        self.notes = ''              # .*
        self.group = []              # ['.', '.', '@.', ...]


class Chars:

    def __init__(self):
        csv_lines = self._read_csv_file()
        if csv_lines is None:
            self.chars = []
        else:
            chars_lines, groups_lines = self._split_lines(csv_lines)
            chars = self._get_chars(chars_lines)
            groups = self._get_groups(groups_lines)
            self.chars = self._apply_groups_to_chars(groups, chars)
            self.groups = groups

    @staticmethod
    def _read_csv_file():
        csv_file = re.sub('\\.py$', '.csv', __file__)
        try:
            with open(csv_file, 'r') as f:
                csv_data = f.readlines()
        except BaseException:
            sys.stderr.write("error: can't read a csv file\n")
            return None
        return csv_data

    @staticmethod
    def _split_lines(csv_lines):
        chars_lines, groups_lines = [], []
        is_in_chars = True
        for line in csv_lines:
            line = line.rstrip()
            if len(line) > 0 and line[0] == ' ':
                line += '\t'
            if line == '# [groups]':
                is_in_chars = False
            if line == '' or line[0] == '#':
                continue
            if is_in_chars:
                chars_lines.append(line)
            else:
                groups_lines.append(line)
        return chars_lines, groups_lines

    @staticmethod
    def _get_chars(lines) -> dict:
        chars = {}
        for line in lines:
            if len(line) > 1 and line[0] == '#':
                continue
            cs = line.split(',')
            if cs[1] == '""""':
                cs[1] = '"'  # "
            elif cs[1] == '"' and cs[2] == '"':
                cs[1] = ','  # ,
                cs.pop(2)
            elif cs[1] == '"#"':
                cs[1] = '#'  # #
            if cs[1] in chars:
                m = 'warning: "' + cs[1] + '" is already registered\n'
                sys.stderr.write(m)
            chars[cs[1]] = Char()

            chars[cs[1]].id = cs[0]

            chars[cs[1]].char = cs[1]

            chars[cs[1]].utf8 = cs[2]

            chars[cs[1]].sub_id = cs[3]

            if cs[3] != '':
                if cs[3][0] == 'S':
                    chars[cs[1]].char_class = '記号'
                elif cs[3][0] == 'E':
                    chars[cs[1]].char_class = '英数'
                elif cs[3][0] == 'K':
                    chars[cs[1]].char_class = '仮名'
                elif cs[3][0] == 'D':
                    chars[cs[1]].char_class = '常用'
                elif cs[3][0] == 'N':
                    chars[cs[1]].char_class = '人名'
                elif cs[3][0] == 'M':
                    chars[cs[1]].char_class = 'ＭＪ'

            if cs[4] != '':
                if cs[4][0] == '1':
                    chars[cs[1]].level = '第一水準'
                elif cs[4][0] == '2':
                    chars[cs[1]].level = '第二水準'
                elif cs[4][0] == '3':
                    chars[cs[1]].level = '第三水準'
                elif cs[4][0] == '4':
                    chars[cs[1]].level = '第四水準'

            if cs[5] == '1':
                chars[cs[1]].width = '半角'
            else:
                chars[cs[1]].width = '全角'

            if cs[6] != '':
                chars[cs[1]].stroke_count = int(cs[6])

            if cs[7] == '1':
                chars[cs[1]].school_grade = '小一'
            elif cs[7] == '2':
                chars[cs[1]].school_grade = '小二'
            elif cs[7] == '3':
                chars[cs[1]].school_grade = '小三'
            elif cs[7] == '4':
                chars[cs[1]].school_grade = '小四'
            elif cs[7] == '5':
                chars[cs[1]].school_grade = '小五'
            elif cs[7] == '6':
                chars[cs[1]].school_grade = '小六'
            elif cs[7] == 'J':
                chars[cs[1]].school_grade = '中学'

            if cs[8] != '':
                chars[cs[1]].bushu = cs[8]

            if cs[9] != '':
                chars[cs[1]].pronunciation = cs[9]

            if cs[10] != '':
                chars[cs[1]].usage = cs[10]

            if cs[11] != '':
                chars[cs[1]].notes = cs[11]

        return chars

    @staticmethod
    def _get_groups(lines) -> list:
        groups = []
        for line in lines:
            if len(line) > 1 and line[0] == '#':
                continue
            gs = line.split(',')
            if gs[0] == '""""':
                gs[0] = '"'  # "
            elif gs[0] == '"' and gs[1] == '"':
                gs[0] = ','  # ,
                gs.pop(1)
            elif gs[0] == '"#"':
                gs[0] = '#'  # #
            groups.append(gs)
        return groups

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


chars = Chars().chars
groups = Chars().groups
