#!/usr/bin/python3
# Name:         makdo_gui.py
# Version:      v07 Furuichibashi
# Time-stamp:   <2024.08.04-11:00:36-JST>

# makdo_gui.py
# Copyright (C) 2024-2024  Seiichiro HATA
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


import os
import sys
import argparse
import re
import datetime
import tkinter
import tkinter.filedialog
import tkinter.messagebox
# import tkinter.font
from tkinterdnd2 import TkinterDnD, DND_FILES
import zipfile
import tempfile
import importlib
import makdo.makdo_md2docx
import makdo.makdo_docx2md


__version__ = 'v07 Furuichibashi'

WINDOW_SIZE = '900x900'
GOTHIC_FONT = 'ＭＳ ゴシック'
MINCHO_FONT = 'ＭＳ 明朝'
# GOTHIC_FONT = 'IPAゴシック'
# MINCHO_FONT = 'IPA明朝'
# GOTHIC_FONT = 'BIZ UDゴシック'
# MINCHO_FONT = 'BIZ UD明朝 Medium'
# GOTHIC_FONT = 'Noto Sans Mono CJK JP'
# MINCHO_FONT = 'Noto Serif CJK JP'
NOT_ESCAPED = '^((?:(?:.|\n)*?[^\\\\])??(?:\\\\\\\\)*?)??'

WHITE_SPACE = ('#C0C000', '#FFFFFF', '#F7F700')

# Y=            0.5        0.7        0.9
COLOR_SPACE = (('#FF5D5D', '#FF9E9E', '#FFDFDF'),  # 000 * comment
               ('#FF5F4E', '#FF9F95', '#FFDFDC'),  # 005
               ('#FF603C', '#FFA08A', '#FFDFD8'),  # 010
               ('#FF6229', '#FFA17F', '#FFE0D4'),  # 015
               ('#FF6512', '#FFA271', '#FFE0D0'),  # 020 * del
               ('#F76900', '#FFA461', '#FFE1CA'),  # 025
               ('#E07000', '#FFA64D', '#FFE1C4'),  # 030 * sect1
               ('#CC7600', '#FFA933', '#FFE2BB'),  # 035
               ('#BC7A00', '#FFAC10', '#FFE3AF'),  # 040 * sect2 第１
               ('#AD7F00', '#F2B200', '#FFE59E'),  # 045
               ('#A08300', '#E0B700', '#FFE882'),  # 050 * sect3 １
               ('#948600', '#D0BC00', '#FFEE4A'),  # 055
               ('#898900', '#C0C000', '#F7F700'),  # 060 * sect4 (1), 判断者
               ('#7F8D00', '#B2C500', '#E5FD00'),  # 065
               ('#758F00', '#A4C900', '#D5FF1A'),  # 070 * sect5 ア
               ('#6B9200', '#96CD00', '#CAFF3A'),  # 075
               ('#619500', '#88D100', '#C2FF50'),  # 080 * sect6 (ｱ), paren1
               ('#589800', '#7BD500', '#BCFF62'),  # 085
               ('#4E9B00', '#6DD900', '#B8FF70'),  # 090 * sect7 ａ,  paren2
               ('#439E00', '#5EDE00', '#B4FF7C'),  # 095
               ('#38A200', '#4FE200', '#B0FF86'),  # 100 * sect8 (a), paren3
               ('#2CA500', '#3EE700', '#ADFF8F'),  # 105
               ('#1FA900', '#2CED00', '#AAFF97'),  # 110
               ('#11AD00', '#17F300', '#A8FF9F'),  # 115
               ('#00B200', '#00FA00', '#A5FFA5'),  # 120 * br, pgbr, fd
               ('#00B111', '#00F718', '#A3FFAC'),  # 125
               ('#00AF20', '#00F52D', '#A1FFB2'),  # 130
               ('#00AE2F', '#00F341', '#9FFFB9'),  # 135
               ('#00AC3C', '#00F154', '#9DFFBF'),  # 140
               ('#00AB49', '#00EF66', '#9BFFC5'),  # 145
               ('#00AA55', '#00EE77', '#98FFCC'),  # 150 * length reviser
               ('#00A861', '#00EC88', '#96FFD3'),  # 155
               ('#00A76D', '#00EA99', '#94FFDA'),  # 160
               ('#00A67A', '#00E8AA', '#91FFE2'),  # 165
               ('#00A586', '#00E7BC', '#8EFFEA'),  # 170
               ('#00A394', '#00E5CF', '#8BFFF4'),  # 175
               ('#00A2A2', '#00E3E3', '#87FFFF'),  # 180 *  algin, 申立人
               ('#00A0B1', '#00E1F8', '#A4F6FF'),  # 185
               ('#009FC3', '#21D6FF', '#B5F1FF'),  # 190
               ('#009DD6', '#42CCFF', '#C0EEFF'),  # 195
               ('#009AED', '#59C5FF', '#C8ECFF'),  # 200 * hsp, ins
               ('#0896FF', '#6BC0FF', '#CEEAFF'),  # 205
               ('#1F8FFF', '#79BCFF', '#D2E9FF'),  # 210 * chap1 第１編, hnumb
               ('#3389FF', '#84B8FF', '#D6E7FF'),  # 215
               ('#4385FF', '#8EB6FF', '#D9E7FF'),  # 220 * chap2 第１章, tab
               ('#5280FF', '#97B3FF', '#DCE6FF'),  # 225
               ('#5F7CFF', '#9FB1FF', '#DFE5FF'),  # 230 * chap3 第１節
               ('#6B79FF', '#A6AFFF', '#E1E4FF'),  # 235
               ('#7676FF', '#ADADFF', '#E4E4FF'),  # 240 * chap4 第１款, fsp
               ('#8072FF', '#B3ABFF', '#E6E3FF'),  # 245
               ('#8A70FF', '#B9A9FF', '#E8E2FF'),  # 250 * chap5 第１目
               ('#946DFF', '#BFA7FF', '#EAE2FF'),  # 255
               ('#9E6AFF', '#C5A5FF', '#ECE1FF'),  # 260 *
               ('#A767FF', '#CAA4FF', '#EDE1FF'),  # 265
               ('#B164FF', '#D0A2FF', '#EFE0FF'),  # 270 *
               ('#BC61FF', '#D7A0FF', '#F2DFFF'),  # 275
               ('#C75DFF', '#DD9EFF', '#F4DFFF'),  # 280 *
               ('#D35AFF', '#E49CFF', '#F6DEFF'),  # 285
               ('#E056FF', '#EC9AFF', '#F9DDFF'),  # 290
               ('#EE52FF', '#F597FF', '#FCDCFF'),  # 295
               ('#FF4DFF', '#FF94FF', '#FFDBFF'),  # 300 * 相手方
               ('#FF4EEE', '#FF95F5', '#FFDCFC'),  # 305
               ('#FF50DF', '#FF96EC', '#FFDCF9'),  # 310
               ('#FF51D0', '#FF97E3', '#FFDCF6'),  # 315
               ('#FF53C3', '#FF98DB', '#FFDDF3'),  # 320
               ('#FF54B6', '#FF98D3', '#FFDDF0'),  # 325
               ('#FF55AA', '#FF99CC', '#FFDDEE'),  # 330 * list, fnumb
               ('#FF579E', '#FF9AC5', '#FFDDEC'),  # 335
               ('#FF5892', '#FF9BBE', '#FFDEE9'),  # 340
               ('#FF5985', '#FF9BB6', '#FFDEE7'),  # 345
               ('#FF5A79', '#FF9CAE', '#FFDEE4'),  # 350
               ('#FF5C6B', '#FF9DA6', '#FFDEE1'))  # 355

KEYWORDS = [
    ['(加害者' +
     '|被告|本訴被告|反訴原告|被控訴人|被上告人' +
     '|相手方' +
     '|被疑者|被告人|弁護人|対象弁護士)',
     'c-300-50-g-x'],
    ['(被害者' +
     '|原告|本訴原告|反訴被告|控訴人|上告人' +
     '|申立人' +
     '|検察官|検察事務官|懲戒請求者)',
     'c-180-50-g-x'],
    ['(裁判官|審判官|調停官|調停委員|司法委員|専門委員|書記官|事務官|訴外)',
     'c-60-50-g-x']]

PARAGRAPH_SAMPLE = ['\t',
                    '<!--コメント-->',
                    '# <!--タイトル-->', '## <!--第１-->', '### <!--１-->',
                    '#### <!--(1)-->', '##### <!--ア-->', '###### <!--(ｱ)-->',
                    '####### <!--ａ-->', '######## <!--(a)-->',
                    ': <!--左寄せ-->', ': <!--中寄せ--> :', '<!--右寄せ--> :',
                    '|<!--表のセル-->|<!--表のセル-->|',
                    '![<!--画像の名前-->](<!--画像のファイル名-->)',
                    '$ <!--第１編-->', '$$ <!--第１章-->', '$$$ <!--第１節-->',
                    '$$$$ <!--第１款-->', '$$$$$ <!--第１目-->',
                    '\\[<!--数式-->\\]', '<pgbr><!--改ページ-->',
                    '<!--------------------------vv-------' +
                    '---------------------vv------------->',
                    '\t']

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
                         '\t']


class CharsState:

    def __init__(self):
        self.del_or_ins = ''
        self.is_in_comment = False
        self.parentheses = []
        self.has_underline = False
        self.has_specific_font = False
        self.is_length_reviser = False
        self.chapter_depth = 0
        self.section_depth = 0

    def __eq__(self, other):
        if self.del_or_ins == other.del_or_ins:
            if self.is_in_comment == other.is_in_comment:
                return True
        return False

    def copy(self):
        copy = CharsState()
        copy.del_or_ins = self.del_or_ins
        copy.is_in_comment = self.is_in_comment
        for p in self.parentheses:
            copy.parentheses.append(p)
        copy.has_underline = self.has_underline
        copy.has_specific_font = self.has_specific_font
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

    def apply_parenthesis(self, parenthesis):
        ps = self.parentheses
        p = parenthesis
        if p == '「' or p == '『' or p == '（' or p == '(':
            ps.append(p)
        if p == ')' or p == '）' or p == '』' or p == '」':
            if len(ps) > 0:
                if ps[-1] == '(' and p == ')' or \
                   ps[-1] == '（' and p == '）' or \
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
        elif self.is_in_comment:
            key += '-0'
        elif chars == ' ':
            key += '-200'
        elif chars == '\t':
            return 'tab_tag'
        elif chars == '\u3000':
            key += '-240'
        elif chars == 'font decorator':
            key += '-120'
        elif chars == 'half number':
            key += '-210'
        elif chars == 'full number':
            key += '-330'
        elif chars == 'list':
            key += '-330'
        elif chars == 'alignment':
            key += '-180'
        elif len(self.parentheses) == 1:
            key += '-80'
        elif len(self.parentheses) == 2:
            key += '-120'
        elif len(self.parentheses) >= 3:
            key += '-160'
        elif chars == '<br>' or chars == '<pgbr>':
            key += '-120'
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
            key += '-30'
        elif self.del_or_ins == 'ins':
            key += '-70'
        else:
            key += '-50'
        # FONT
        if chars == 'mincho':
            key += '-m'  # mincho
        else:
            key += '-g'  # gothic
        # UNDERLINE
        if chars == 'font decorator':
            key += '-x'  # no underline
        elif (not self.is_in_comment and
              (chars == ' ' or chars == '\t' or chars == '\u3000')):
            key += '-u'  # underline
        elif not self.is_in_comment and self.has_underline:
            key += '-u'  # underline
        elif not self.is_in_comment and self.has_specific_font:
            key += '-u'  # underline
        else:
            key += '-x'  # no underline
        # RETURN
        return key


class LineDatum:

    def __init__(self):
        self.line_number = 0
        self.line_text = ''
        self.beg_chars_state = CharsState()
        self.end_chars_state = CharsState()
        self.should_paint_keywords = False

    def paint_line(self, txt, should_paint_keywords=False):
        # PREPARE
        i = self.line_number
        line_text = self.line_text
        chars_state = self.beg_chars_state.copy()
        self.should_paint_keywords = should_paint_keywords
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
                key = 'c-20-50-g-x'                                     # 1.key
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
                key = 'c-200-50-g-x'                                    # 1.key
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
            # FONT DECORATOR ("---", "+++", ">>>", "<<<")
            if (not chars_state.is_in_comment) and \
               (s3 == '---' or s3 == '+++' or s3 == '>>>' or s3 == '<<<') and \
               (c4 != '\\' or re.match(NOT_ESCAPED + '...$', tmp)):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j - 2)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '---' or '+++' or '>>>' or '<<<'                # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
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
                # tmp = '--' or '++' or '>>' or '<<'                    # 5.tmp
                beg = end                                               # 6.beg
                key = chars_state.get_key('font decorator')             # 1.key
                end = str(i + 1) + '.' + str(j + 1)                     # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                tmp = ''                                                # 5.tmp
                beg = end                                               # 6.beg
                continue
            # FONT DECORATOR ("@.+@", "^.*^", "_.*_")
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
            # PARENTHESES
            if c == '「' or c == '『' or c == '（' or c == '(':
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                chars_state.apply_parenthesis(c)                        # 4.set
                tmp = c                                                 # 5.tmp
                beg = end                                               # 6.beg
                continue
            if c == ')' or c == '）' or c == '』' or c == '」':
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
            if re.match('[０-９零一二三四五六七八九十]', c):
                key = chars_state.get_key('')                           # 1.key
                end = str(i + 1) + '.' + str(j)                         # 2.end
                txt.tag_add(key, beg, end)                              # 3.tag
                #                                                       # 4.set
                # tmp = '[０-９]'                                       # 5.tmp
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
            if self.should_paint_keywords:
                for kw in KEYWORDS:
                    if re.match('^(.*?)' + kw[0] + '$', tmp):
                        t1 = re.sub('^(.*?)' + kw[0] + '$', '\\1', tmp)
                        t2 = re.sub('^(.*?)' + kw[0] + '$', '\\2', tmp)
                        if t2 == '原告' or t2 == '被告':
                            if re.match('^(?:.|\n)*(本|反)訴$', t1):
                                continue
                        if t2 == '被告' and c0 == '人':
                            continue
                        key = chars_state.get_key('')                   # 1.key
                        end = str(i + 1) + '.' + str(j - len(t2) + 1)   # 2.end
                        txt.tag_add(key, beg, end)                      # 3.tag
                        #                                               # 4.set
                        # tmp = t2                                      # 5.tmp
                        beg = end                                       # 6.beg
                        key = kw[1]                                     # 1.key
                        end = str(i + 1) + '.' + str(j + 1)             # 2.end
                        txt.tag_add(key, beg, end)                      # 3.tag
                        #                                               # 4.set
                        tmp = ''                                        # 5.tmp
                        beg = end                                       # 6.beg
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


class Makdo:

    search_word = ''

    def __init__(self, args):
        self.tmep_dir = ''
        self.file_path = args.input_file
        self.init_text = ''
        self.file_lines = []
        self.must_make_backup_file = args.input_file
        self.font_size = None
        self.number_of_period = 0
        self.line_data = []
        self.cursor_line = 0
        self.global_line_to_paint = 0
        self.local_line_to_paint = 0
        self.akauni_history = ['', '', '', '', '']
        # WINDOW
        self.win = TkinterDnD.Tk()  # drag and drop
        # self.win = tkinter.Tk()
        self.win.title('MAKDO')
        self.win.geometry(WINDOW_SIZE)
        self.win.protocol("WM_DELETE_WINDOW", self.quit_makdo)
        # FRAME
        self.frm = tkinter.Frame()
        self.frm.pack(expand=True, fill=tkinter.BOTH)
        self.txt = tkinter.Text(self.frm, width=80, height=30, undo=True)
        self.txt.drop_target_register(DND_FILES)  # drag and drop
        self.txt.dnd_bind('<<Drop>>', self.drop)  # drag and drop
        # MENU BAR
        self.mnb = tkinter.Menu()
        #
        self.mc1 = tkinter.Menu(self.mnb, tearoff=False)
        self.mc1.add_command(label='ファイルを開く',
                             command=self.open_file)
        self.mc1.add_command(label='ファイルを閉じる',
                             command=self.close_file)
        self.mc1.add_separator()
        self.mc1.add_command(label='ファイルを保存する',
                             command=self.save_file,
                             accelerator='Ctrl+S')
        self.mc1.add_command(label='名前を付けて保存する',
                             command=self.name_and_save)
        self.mc1.add_separator()
        self.mc1.add_command(label='終了',
                             command=self.quit_makdo,
                             accelerator='Ctrl+Q')
        self.mnb.add_cascade(label='ファイル', menu=self.mc1)
        #
        self.mc2 = tkinter.Menu(self.mnb, tearoff=False)
        self.mc2.add_command(label='元に戻す',
                             command=self.txt.edit_undo,
                             accelerator='Ctrl+Z')
        self.mc2.add_command(label='やり直す',
                             command=self.txt.edit_redo,
                             accelerator='Ctrl+Y')
        self.mc2.add_separator()
        self.mc2.add_command(label='切り取り',
                             command=self.cut_text,
                             accelerator='Ctrl+X')
        self.mc2.add_command(label='コピー',
                             command=self.copy_text,
                             accelerator='Ctrl+C')
        self.mc2.add_command(label='貼り付け',
                             command=self.paste_text,
                             accelerator='CTrl+V')
        self.mc2.add_separator()
        self.mc2.add_command(label='計算',
                             command=self.calculate)
        self.mnb.add_cascade(label='編集', menu=self.mc2)
        #
        self.mc3 = tkinter.Menu(self.mnb, tearoff=False)
        self.mc3.add_command(label='微サイズ',
                             command=self.set_size_ss)
        self.mc3.add_command(label='小サイズ',
                             command=self.set_size_s)
        self.mc3.add_command(label='中サイズ',
                             command=self.set_size_m)
        self.mc3.add_command(label='大サイズ',
                             command=self.set_size_l)
        self.mc3.add_command(label='巨サイズ',
                             command=self.set_size_ll)
        self.mnb.add_cascade(label='文字サイズ', menu=self.mc3)
        #
        self.mc4 = tkinter.Menu(self.mnb, tearoff=False)
        self.should_paint_keywords = tkinter.BooleanVar()
        if args.paint_keywords:
            self.should_paint_keywords.set(True)
        else:
            self.should_paint_keywords.set(False)
        self.mc4.add_checkbutton(label='キーワードに色付け',
                                 variable=self.should_paint_keywords)
        self.mc4.add_separator()
        self.is_read_only = tkinter.BooleanVar()
        if args.read_only:
            self.is_read_only.set(True)
        else:
            self.is_read_only.set(False)
        self.mc4.add_checkbutton(label='読み取り専用',
                                 variable=self.is_read_only)
        self.mc4.add_separator()
        self.sb1 = tkinter.Menu(self.mnb, tearoff=False)
        self.digit_separator = tkinter.StringVar(value='4')
        self.sb1.add_radiobutton(label='桁区切りなし（12345678）',
                                 value='0', variable=self.digit_separator)
        self.sb1.add_radiobutton(label='3桁区切り（12,345,678）',
                                 value='3', variable=self.digit_separator)
        self.sb1.add_radiobutton(label='4桁区切り（1234万5678）',
                                 value='4', variable=self.digit_separator)
        self.mc4.add_cascade(label='数式計算結果', menu=self.sb1)
        self.mnb.add_cascade(label='設定', menu=self.mc4)
        #
        self.mc5 = tkinter.Menu(self.mnb, tearoff=False)
        self.mc5.add_command(label='基本',
                             command=self.insert_basis)
        self.mc5.add_command(label='民法',
                             command=self.insert_law)
        self.mc5.add_command(label='訴状',
                             command=self.insert_petition)
        self.mc5.add_command(label='証拠説明書',
                             command=self.insert_evidence)
        self.mc5.add_command(label='和解契約書',
                             command=self.insert_settlement)
        self.mnb.add_cascade(label='サンプル', menu=self.mc5)
        self.win['menu'] = self.mnb
        # TEXT
        self.txt.pack(expand=True, fill=tkinter.BOTH)
        self.txt.bind('<Key>', self.process_key_press)
        self.txt.bind('<Key-Tab>', self.process_key_tab)
        self.txt.bind('<KeyRelease>', self.process_key_release)
        self.txt.bind('<ButtonRelease-1>', self.process_button1_release)
        self.txt.bind('<ButtonRelease-2>', self.process_button2_release)
        self.txt.bind('<Button-3>', self.process_button3)
        self.txt.config(bg='black', fg='white')
        self.txt.config(insertbackground='#FF7777', blockcursor=True)  # CURSOR
        # SCROLL BAR
        scb = tkinter.Scrollbar(self.txt, orient=tkinter.VERTICAL,
                                command=self.txt.yview)
        scb.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.txt['yscrollcommand'] = scb.set
        # STATUS BAR
        self.stb = tkinter.Frame(self.frm)
        self.stb.pack(anchor=tkinter.W)
        self.stb_pos1 = tkinter.Label(self.stb, text='1x0')
        self.stb_pos1.pack(side=tkinter.LEFT)
        self.stb_spc1 = tkinter.Label(self.stb, text='\u3000')
        self.stb_spc1.pack(side=tkinter.LEFT)
        self.stb_sor1 = tkinter.Entry(self.stb, width=20)
        self.stb_sor1.pack(side=tkinter.LEFT)
        self.stb_sor1.insert(0, '（検索語）')
        self.stb_sor2 = tkinter.Entry(self.stb, width=20)
        self.stb_sor2.pack(side=tkinter.LEFT)
        self.stb_sor2.insert(0, '（置換語）')
        self.stb_sor3 = tkinter.Button(self.stb, text='前',
                                       command=self.search_or_replace_backward)
        self.stb_sor3.pack(side=tkinter.LEFT)
        self.stb_sor4 = tkinter.Button(self.stb, text='次',
                                       command=self.search_or_replace_forward)
        self.stb_sor4.pack(side=tkinter.LEFT)
        self.stb_sor5 = tkinter.Button(self.stb, text='消',
                                       command=self.clear_search_word)
        self.stb_sor5.pack(side=tkinter.LEFT)
        self.stb_spc2 = tkinter.Label(self.stb, text='\u3000')
        self.stb_spc2.pack(side=tkinter.LEFT)
        self.stb_fnm1 = tkinter.Label(self.stb, text='')
        self.stb_fnm1.pack(side=tkinter.LEFT)
        self.stb_spc3 = tkinter.Label(self.stb, text='\u3000')
        self.stb_spc3.pack(side=tkinter.LEFT)
        self.stb_msg1 = tkinter.Label(self.stb, text='')
        self.stb_msg1.pack(side=tkinter.LEFT)
        # FONT
        # tkinter.font.families()
        self.set_size_m()
        self.txt.tag_config('error_tag', foreground='#FF0000')
        self.txt.tag_config('search_tag', background='#777777')
        self.txt.tag_config('akauni_tag', background='#777777')
        self.txt.tag_config('tab_tag', background='#5280FF')  # (0.5, 225, max)
        # OPEN FILE
        if args.input_file is not None:
            self.just_open_file(args.input_file)
        # LOOP
        self.txt.focus_set()
        self.run_periodically()
        self.win.mainloop()

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
                    n, m = 'エラー', '自動保存ファイルの削除に失敗しました'
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

    # INPUT AND OUTPUT

    def drop(self, event):                                # drag and drop
        ans = self.close_file()                           # drag and drop
        if ans is None:                                   # drag and drop
            return None                                   # drag and drop
        file_path = event.data                            # drag and drop
        file_path = re.sub('^{(.*)}$', '\\1', file_path)  # drag and drop
        self.just_open_file(file_path)                    # drag and drop

    def open_file(self):
        ans = self.close_file()
        if ans is None:
            return None
        typ = [('可能な形式', '.md .docx'),
               ('Markdown', '.md'), ('MS Word', '.docx')]
        file_path = tkinter.filedialog.askopenfilename(filetypes=typ)
        if file_path == ():
            return None
        self.just_open_file(file_path)

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
            with open(md_path, 'r', encoding='utf-8') as f:
                init_text = f.read()
        except BaseException:
            return
        self.file_path = file_path
        self.init_text = init_text
        self.file_lines = init_text.split('\n')
        self.txt.delete('1.0', 'end')
        self.txt.insert('1.0', init_text)
        self.txt.focus_set()
        file_name = re.sub('^.*[/\\\\]', '', file_path)
        self.win.title(file_name + ' - MAKDO')
        self.save_auto_file(self.file_path)
        self.stb_fnm1['text'] = file_name
        self.txt.mark_set('insert', '1.0')
        self.line_data = [LineDatum() for line in init_text.split('\n')]
        for i, line in enumerate(self.file_lines):
            self.line_data[i].line_number = i
            self.line_data[i].line_text = line + '\n'
            if i > 0:
                self.line_data[i].beg_chars_state \
                    = self.line_data[i - 1].end_chars_state.copy()
                self.line_data[i].beg_chars_state.reset_partially()
            self.line_data[i].paint_line(self.txt)
        self.txt.edit_reset()

    def close_file(self):
        # SAVE FILE
        if self.has_edited():
            ans = self._ask_to_save('保存しますか？')
            if ans is None:
                return None
            elif ans is True:
                if not self.save_file():
                    return None
        if self.has_edited():
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
        self.stb_fnm1['text'] = ''
        return True

    def has_edited(self):
        file_text = self.txt.get('1.0', 'end-1c')
        if file_text != '':
            if self.init_text != file_text:
                return True
        return False

    def _ask_to_save(self, message):
        tkinter.Tk().withdraw()
        n, m, d = '確認', message, 'yes'
        return tkinter.messagebox.askyesnocancel(n, m, default=d)

    def save_file(self):
        if self.has_edited():
            file_text = self.txt.get('1.0', 'end-1c')
            self._stamp_time(file_text)
            if file_text == '' or file_text[-1] != '\n':
                self.txt.insert('end', '\n')
            file_text = self.txt.get('1.0', 'end-1c')
            if (self.file_path is None) or (self.file_path == ''):
                typ = [('Markdown', '*.md')]
                file_path = tkinter.filedialog.asksaveasfilename(filetypes=typ)
                if file_path == ():
                    return False
                self.file_path = file_path
            if self.must_make_backup_file:
                if os.path.exists(self.file_path) and \
                   not os.path.islink(self.file_path):
                    try:
                        os.rename(self.file_path, self.file_path + '~')
                        self.must_make_backup_file = False
                    except BaseException:
                        n, m = 'エラー', 'バックアップに失敗しました'
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
                n, m = 'エラー', 'ファイルの保存に失敗しました'
                tkinter.messagebox.showerror(n, m)
                return False
            # SAVE DOCX FILE
            if re.match('^(?:.|\n)+.docx$', self.file_path):
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
                    n = 'エラー'
                    tkinter.messagebox.showerror(n, msg)
                    return
            self.set_message('保存しました')
            self.init_text = file_text
            return True

    def _stamp_time(self, file_text):
        if not re.match('^\\s*<!--', file_text):
            return
        file_text = re.sub('-->(.|\n)*$', '', file_text)
        now = datetime.datetime.utcnow() + datetime.timedelta(hours=+9)
        jst = datetime.timezone(datetime.timedelta(hours=+9))
        now = now.replace(tzinfo=jst)
        res = '^(\\S+:\\s*)(.*)$'
        for i, line in enumerate(file_text.split('\n')):
            # CREATED TIME
            if re.match('^作成時:', line) or re.match('^created_time:', line):
                cfg = re.sub(res, '\\1', line)
                val = re.sub(res, '\\2', line)
                j = len(cfg)
                beg = str(i + 1) + '.' + str(j)
                end = str(i + 1) + '.end'
                res_jst = '^' + '[0-9]{4}-[0-9]{2}-[0-9]{2}' + \
                    'T[0-9]{2}:[0-9]{2}:[0-9]{2}\\+09:00' + '$'
                if not re.match(res_jst, val):
                    val = ''
                try:
                    dt = datetime.datetime.fromisoformat(val)
                except BaseException:
                    self.txt.delete(beg, end)
                    self.txt.insert(beg, now.isoformat(timespec='seconds'))
            if re.match('^更新時:', line) or re.match('^modified_time:', line):
                cfg = re.sub(res, '\\1', line)
                val = re.sub(res, '\\2', line)
                j = len(cfg)
                beg = str(i + 1) + '.' + str(j)
                end = str(i + 1) + '.end'
                self.txt.delete(beg, end)
                self.txt.insert(beg, now.isoformat(timespec='seconds'))

    def name_and_save(self):
        typ = [('可能な形式', '.md .docx'),
               ('Markdown', '.md'), ('MS Word', '.docx')]
        file_path = tkinter.filedialog.asksaveasfilename(filetypes=typ)
        self.file_path = file_path
        self.save_file()

    def quit_makdo(self):
        ans = self.close_file()
        if ans is None:
            return None
        self.win.quit()
        self.win.destroy()
        sys.exit(0)

    # EDIT

    def cut_text(self):
        if self.txt.tag_ranges('sel'):
            c = self.txt.get('sel.first', 'sel.last')
            self.win.clipboard_clear()
            self.win.clipboard_append(c)
            self.txt.delete('sel.first', 'sel.last')  # delete
        return

    def copy_text(self):
        if self.txt.tag_ranges('sel'):
            c = self.txt.get('sel.first', 'sel.last')
            self.win.clipboard_clear()
            self.win.clipboard_append(c)
        return

    def paste_text(self):
        try:
            c = self.win.clipboard_get()
            self.txt.insert('insert', c)
            # self.txt.yview('insert -20 line')
        except BaseException:
            pass
        return

    def calculate(self):
        inse = self.txt.index('insert')
        numb = re.sub('\\..*$', '', inse)
        line = self.txt.get(numb + '.0', numb + '.end')
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
        r = str(eval(math))
        if not re.match('^-?([0-9]+\\.)?[0-9]+', r):
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
        beg = numb + '.' + str(len(line_head + line_math))
        end = numb + '.' + str(len(line_head + line_math + line_rslt))
        self.txt.delete(beg, end)
        self.txt.insert(beg, '=' + r)
        self.win.clipboard_clear()
        self.win.clipboard_append(r)

    # STATUS BAR

    def search_or_replace_backward(self):
        word1 = self.stb_sor1.get()
        word2 = self.stb_sor2.get()
        if Makdo.search_word != word1:
            Makdo.search_word = word1
            self._highlight_search_word()
        pos = self.txt.index('insert')
        tex = self.txt.get('1.0', pos)
        res = '^((?:.|\n)*)(' + word1 + '(?:.|\n)*)$'
        if re.match(res, tex):
            t1 = re.sub(res, '\\1', tex)
            t2 = re.sub(res, '\\2', tex)
            # SEARCH
            self.txt.mark_set('insert', pos + '-' + str(len(t2)) + 'c')
            self.txt.yview(pos + '-' + str(len(t2)) + 'c')
            if word2 != '' and word2 != '（置換語）':
                # REPLACE
                self.txt.delete('insert', 'insert+' + str(len(word1)) + 'c')
                self.txt.insert('insert', word2)
                # self.stb_msg1['text'] = '置換しました'
        self.set_message('')
        self.txt.focus_set()

    def search_or_replace_forward(self):
        word1 = self.stb_sor1.get()
        word2 = self.stb_sor2.get()
        if Makdo.search_word != word1:
            Makdo.search_word = word1
            self._highlight_search_word()
        pos = self.txt.index('insert')
        tex = self.txt.get(pos, 'end-1c')
        res = '^((?:.|\n)*?' + word1 + ')((?:.|\n)*)$'
        if re.match(res, tex):
            t1 = re.sub(res, '\\1', tex)
            t2 = re.sub(res, '\\2', tex)
            # SEARCH
            self.txt.mark_set('insert', pos + '+' + str(len(t1)) + 'c')
            self.txt.yview(pos + '+' + str(len(t1)) + 'c')
            if word2 != '' and word2 != '（置換語）':
                # REPLACE
                self.txt.delete('insert-' + str(len(word1)) + 'c', 'insert')
                self.txt.insert('insert', word2)
                # self.stb_msg1['text'] = '置換しました'
        self.set_message('')
        self.txt.focus_set()

    def clear_search_word(self):
        self.stb_sor1.delete('0', 'end')
        self.stb_sor2.delete('0', 'end')
        self.txt.tag_remove('search_tag', '1.0', 'end')
        Makdo.search_word = ''

    def set_message(self, msg):
        self.stb_msg1['text'] = msg

    # FONT SIZE

    def set_size_ss(self):
        self.__set_font_size(8)

    def set_size_s(self):
        self.__set_font_size(12)

    def set_size_m(self):
        self.__set_font_size(18)

    def set_size_l(self):
        self.__set_font_size(26)

    def set_size_ll(self):
        self.__set_font_size(36)

    def __set_font_size(self, size):
        self.font_size = size
        self.txt['font'] = (GOTHIC_FONT, size)
        self.stb_sor1['font'] = (GOTHIC_FONT, size)
        self.stb_sor2['font'] = (GOTHIC_FONT, size)
        for u in ['-x', '-u']:
            und = False if u == '-x' else True
            for f in ['-g', '-m']:
                fon = (GOTHIC_FONT, size) if f == '-g' else (MINCHO_FONT, size)
                # WHITE
                for i in range(3):
                    a = '-XXX'
                    y = '-' + str(30 + (i * 20))
                    tag = 'c' + a + y + f + u
                    col = WHITE_SPACE[i]
                    self.txt.tag_config(tag, font=fon,
                                        foreground=col, underline=und)
                # COLOR
                for i in range(3):  # lightness
                    y = '-' + str(30 + (i * 20))
                    for j, c in enumerate(COLOR_SPACE):  # angle
                        a = '-' + str(j * 5)
                        tag = 'c' + a + y + f + u  # example: c-125-50-g-x
                        col = c[i]
                        self.txt.tag_config(tag, font=fon,
                                            foreground=col, underline=und)

    # RECURSIVE CALL

    def run_periodically(self):
        interval = 10
        # IF THE REGION IS SET
        if self.txt.tag_ranges(tkinter.SEL):
            self.win.after(interval, self.run_periodically)  # NEXT
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
            self.win.after(interval, self.run_periodically)  # NEXT
            return
        # CHECK WHETHER TO PAINT KEYWORDS
        should_paint_keywords = self.should_paint_keywords.get()
        # GLOBAL
        i = self.global_line_to_paint
        if i < len(self.line_data):
            self.line_data[i].line_number = i
            lt = self.file_lines[i] + '\n'
            if i == 0:
                cs = CharsState()
            else:
                cs = self.line_data[i - 1].end_chars_state.copy()
                cs.reset_partially()
            spk = should_paint_keywords
            if self.line_data[i].line_text != lt or \
               self.line_data[i].should_paint_keywords != spk or \
               self.line_data[i].beg_chars_state != cs:
                self.line_data[i].line_text = lt
                self.line_data[i].beg_chars_state = cs
                self.line_data[i].end_chars_state = CharsState()
                self.line_data[i].paint_line(self.txt, should_paint_keywords)
        # LOCAL
        i = self.cursor_line + self.local_line_to_paint - 10
        if i >= 0 and i < len(self.line_data):
            self.line_data[i].line_number = i
            lt = self.file_lines[i] + '\n'
            if i == 0:
                cs = CharsState()
            else:
                cs = self.line_data[i - 1].end_chars_state.copy()
                cs.reset_partially()
            spk = should_paint_keywords
            if self.line_data[i].line_text != lt or \
               self.line_data[i].should_paint_keywords != spk or \
               self.line_data[i].beg_chars_state != cs:
                self.line_data[i].line_text = lt
                self.line_data[i].beg_chars_state = cs
                self.line_data[i].end_chars_state = CharsState()
                self.line_data[i].paint_line(self.txt, should_paint_keywords)
        # POINT
        i = self.cursor_line
        if i < len(self.line_data):
            self.line_data[i].line_number = i
            lt = self.file_lines[i] + '\n'
            if i == 0:
                cs = CharsState()
            else:
                cs = self.line_data[i - 1].end_chars_state.copy()
                cs.reset_partially()
            spk = should_paint_keywords
            if self.line_data[i].line_text != lt or \
               self.line_data[i].should_paint_keywords != spk or \
               self.line_data[i].beg_chars_state != cs:
                self.line_data[i].line_text = lt
                self.line_data[i].beg_chars_state = cs
                self.line_data[i].end_chars_state = CharsState()
                self.line_data[i].paint_line(self.txt, should_paint_keywords)
        # STEP (GLOBAL)
        self.global_line_to_paint += 1
        if self.global_line_to_paint >= m:
            self.global_line_to_paint = 0
        # STEP (LOCAL)
        self.local_line_to_paint += 1
        if self.local_line_to_paint >= 100:
            i = self.txt.index('insert')
            self.cursor_line = int(re.sub('\\..*$', '', i)) - 1
            self.local_line_to_paint = 0
        # READ ONLY
        is_read_only = self.is_read_only.get()
        if self.txt['state'] == 'normal' and is_read_only:
            self.txt.configure(state='disabled')
        if self.txt['state'] == 'disabled' and not is_read_only:
            self.txt.configure(state='normal')
        # AUTO SAVE
        if (self.number_of_period % 6000) == 0:
            self.save_auto_file(self.file_path)
        # TO NEXT
        self.number_of_period += 1
        self.win.after(interval, self.run_periodically)  # NEXT
        return

    @staticmethod
    def __get_end_position(i, j, k):
        return s1, s2, ''

    def _highlight_search_word(self):
        word = Makdo.search_word
        tex = self.txt.get('1.0', 'end-1c')
        beg = 0
        res = '^((?:.|\n)*?)' + word + '((?:.|\n)*)$'
        while re.match(res, tex):
            pre = re.sub(res, '\\1', tex)
            tex = re.sub(res, '\\2', tex)
            beg += len(pre)
            end = beg + len(word)
            self.txt.tag_add('search_tag',
                             '1.0+' + str(beg) + 'c',
                             '1.0+' + str(end) + 'c',)
            beg = end

    def set_current_position(self):
        pos = self.txt.index('insert')
        self.stb_pos1['text'] = str(pos).replace('.', 'x')

    def insert_document(self, document_data):
        chs = CharsState()
        tmp = ''
        for c in document_data + '\0':
            tmp += c

    def process_key_press(self, key):
        # FOR AKAUNI
        self.akauni_history.append(key.keysym)
        self.akauni_history.pop(0)
        if key.keysym == 'F19':              # x (ctrl)
            return 'break'
        elif key.keysym == 'Left':
            if 'akauni' in self.txt.mark_names():
                self.txt.tag_remove('akauni_tag', '1.0', 'end')
                self.txt.tag_add('akauni_tag', 'akauni', 'insert-1c')
                self.txt.tag_add('akauni_tag', 'insert-1c', 'akauni')
        elif key.keysym == 'Right':
            if 'akauni' in self.txt.mark_names():
                self.txt.tag_remove('akauni_tag', '1.0', 'end')
                self.txt.tag_add('akauni_tag', 'akauni', 'insert+1c')
                self.txt.tag_add('akauni_tag', 'insert+1c', 'akauni')
        elif key.keysym == 'Up':
            if 'akauni' in self.txt.mark_names():
                self.txt.tag_remove('akauni_tag', '1.0', 'end')
                self.txt.tag_add('akauni_tag', 'akauni', 'insert-1l')
                self.txt.tag_add('akauni_tag', 'insert-1l', 'akauni')
        elif key.keysym == 'Down':
            if 'akauni' in self.txt.mark_names():
                self.txt.tag_remove('akauni_tag', '1.0', 'end')
                self.txt.tag_add('akauni_tag', 'akauni', 'insert+1l')
                self.txt.tag_add('akauni_tag', 'insert+1l', 'akauni')
        elif key.keysym == 'F17':            # } (, calc)
            if self.akauni_history[-2] == 'F13':
                self.calculate()
                return 'break'
        elif key.keysym == 'F21':            # w (undo)
            self.txt.edit_undo()
            return 'break'
        elif key.keysym == 'XF86AudioMute':  # W (redo)
            self.txt.edit_redo()
            return 'break'
        elif key.keysym == 'F22':            # f (mark, save)
            if self.akauni_history[-2] == 'F19':
                self.save_file()
                return 'break'
            else:
                if 'akauni' in self.txt.mark_names():
                    self.txt.mark_unset('akauni')
                self.txt.mark_set('akauni', 'insert')
                return 'break'
        elif key.keysym == 'Delete':         # d (delete, quit)
            if self.akauni_history[-2] == 'F19':
                self.quit_makdo()
                return 'break'
            elif 'akauni' in self.txt.mark_names():
                self.txt.tag_remove('akauni_tag', '1.0', 'end')
                akn = self.txt.index('akauni')
                pos = self.txt.index('insert')
                beg = re.sub('\\..*$', '.0', akn)
                if akn == pos and akn != beg:
                    c = self.txt.get(beg, akn)
                    self.win.clipboard_clear()
                    self.win.clipboard_append(c)
                    self.txt.delete(beg, akn)
                else:
                    c = ''
                    c += self.txt.get('akauni', 'insert')
                    c += self.txt.get('insert', 'akauni')
                    self.win.clipboard_clear()
                    self.win.clipboard_append(c)
                    self.txt.delete('akauni', 'insert')
                    self.txt.delete('insert', 'akauni')
                self.txt.mark_unset('akauni')
                return 'break'
        elif key.keysym == 'F14':            # v (quit)
            if 'akauni' in self.txt.mark_names():
                self.txt.tag_remove('akauni_tag', '1.0', 'end')
                self.txt.mark_unset('akauni')
                return 'break'
        elif key.keysym == 'F15':            # g (paste)
            c = self.win.clipboard_get()
            self.txt.insert('insert', c)
            # self.txt.yview('insert -20 line')
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
        if key.char == '\x11':    # ctrl-q
            self.quit_makdo()
        elif key.char == '\x13':  # ctrl-s
            self.save_file()
        elif key.keysym == 'Delete':
            if self.txt.tag_ranges('sel'):
                self.cut_text()
            else:
                pos = self.txt.index('insert')
                end = re.sub('\\..*$', '.end', pos)
                c = self.txt.get(pos, end)
                if self.akauni_history[-2] != 'Delete':
                    self.win.clipboard_clear()
                if c == '':
                    self.win.clipboard_append('\n')
                    self.txt.delete(pos, end)
                else:
                    self.win.clipboard_append(c)
                    self.txt.delete(pos, end + '-1c')
        self.set_current_position()

    def process_key_release(self, key):
        # FOR AKAUNI
        if 'akauni' in self.txt.mark_names():
            self.txt.tag_remove('akauni_tag', '1.0', 'end')
            self.txt.tag_add('akauni_tag', 'akauni', 'insert')
            self.txt.tag_add('akauni_tag', 'insert', 'akauni')
        self.set_current_position()

    def process_key_tab(self, key):
        # CALCULATE
        text_beg_to_cur = self.txt.get('1.0', 'insert')
        text = text_beg_to_cur
        res_open = '^((?:.|\n)*)(<!--(?:.|\n)*)'
        res_close = '^((?:.|\n)*)(-->(?:.|\n)*)'
        if re.match(res_open, text):
            text = re.sub(res_open, '\\2', text)
            if not re.match(res_close, text):
                self.calculate()
                return 'break'
        # INSERT
        point_cur = self.txt.index('insert')
        point_beg = re.sub('\\..*$', '.0', point_cur)
        point_end = re.sub('\\..*$', '.end', point_cur)
        line_cur_to_end = self.txt.get(point_cur, point_end)
        line_beg_to_end = self.txt.get(point_beg, point_end)
        if re.match('^.*\\.0$', point_cur):
            if line_beg_to_end == '':
                self.txt.insert('insert', PARAGRAPH_SAMPLE[0])
                self.txt.mark_set('insert', point_beg)
                return 'break'
            for i, sample in enumerate(PARAGRAPH_SAMPLE):
                if line_beg_to_end == sample:
                    self.txt.delete(point_beg, point_end)
                    self.txt.insert('insert', PARAGRAPH_SAMPLE[i + 1])
                    self.txt.mark_set('insert', point_beg)
                    return 'break'
        else:
            for i, sample in enumerate(FONT_DECORATOR_SAMPLE):
                sample_esc = sample
                sample_esc = sample_esc.replace('*', '\\*')
                sample_esc = sample_esc.replace('+', '\\+')
                sample_esc = sample_esc.replace('^', '\\^')
                if re.match('^' + sample_esc, line_cur_to_end):
                    point_tmp = point_cur + '+' + str(len(sample)) + 'c'
                    self.txt.delete(point_cur, point_tmp)
                    self.txt.insert('insert', FONT_DECORATOR_SAMPLE[i + 1])
                    self.txt.mark_set('insert', point_cur)
                    return 'break'
            else:
                self.txt.insert('insert', FONT_DECORATOR_SAMPLE[0])
                self.txt.mark_set('insert', point_cur)
                return 'break'

    def process_button1_release(self, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        self.set_current_position()

    def process_button2_release(self, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        self.paste_text()

    def process_button3(self, click):
        try:
            self.bt3.destroy()
        except BaseException:
            pass
        self.bt3 = tkinter.Menu(self.win, tearoff=False)
        self.bt3.add_command(label='切り取り', command=self.cut_text)
        self.bt3.add_command(label='コピー', command=self.copy_text)
        self.bt3.add_command(label='貼り付け', command=self.paste_text)
        self.bt3.post(click.x_root, click.y_root)

    # SAMPLE

    def insert_configuration(self, document_style, space_before):
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
作成時: -
更新時: -

---------------------------------------------------------------->
'''
        return document

    def insert_basis(self):
        document = self.insert_configuration('普通', '0.0') + \
            '''
# ★（タイトル）

v=+1.0
### ★（第1項）

★

### ★（第2項）

★
'''
        self.insert_sample(document)

    def insert_law(self):
        document = self.insert_configuration('条文', '0.0') + \
            '''
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
        self.insert_sample(document)

    def insert_petition(self):
        document = self.insert_configuration('普通', '1.0') + \
            '''
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
        self.insert_sample(document)

    def insert_evidence(self):
        document = self.insert_configuration('普通', '0.0') + \
            '''
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
        self.insert_sample(document)

    def insert_settlement(self):
        document = self.insert_configuration('契約', '1.0') + \
            '''
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
        self.insert_sample(document)

    def insert_sample(self, sample_document):
        current_document = self.txt.get('1.0', 'end-1c')
        if current_document != '':
            return
        self.txt.insert('1.0', sample_document)


if __name__ == '__main__':

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
        '-c', '--paint-keywords',
        action='store_true',
        help='キーワードに色を付けます')
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

    Makdo(args)
