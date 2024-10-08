#!/usr/bin/python3
# Name:         mddiff.py
# Version:      v07 Furuichibashi
# Time-stamp:   <2024.10.09-07:33:02-JST>

# md2docx.py
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


import sys
import argparse     # Python Software Foundation License
import re
import chardet      # GNU Lesser General Public License v2 or later (LGPLv2+)
import Levenshtein  # GNU General Public License v2 or later (GPLv2+)
import hashlib


__version__ = 'v01'


def get_arguments():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description='Markdownファイルを比べたり、違いを適用したりします',
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
        '-a', '--apply-difference',
        type=str,
        metavar='DEFFERENCE_ID',
        help='違いを適用します')
    parser.add_argument(
        '-r', '--print-reverse-id',
        type=str,
        metavar='DEFFERENCE_ID',
        help='逆の違いIDを表示します')
    parser.add_argument(
        '-H', '--html',
        action='store_true',
        help='違いをHTMLで表示します')
    parser.add_argument(
        '-V', '--verbose',
        action='store_true',
        help='違いがない段落も表示します')
    parser.add_argument(
        'main_md_file',
        help='主Markdownファイル')
    parser.add_argument(
        'sub_md_file',
        help='副Markdownファイル')
    return parser.parse_args()


SALT = "L'essentiel est invisible pour les yeux."


class File:

    """A class to handle a file"""

    def __init__(self, file_name=None):
        self.file_name = None
        self.file_text = None
        self.raw_paragraphs = None
        self.cmp_paragraphs = None
        self.app_paragraphs = None
        self.configs = {}
        if file_name is not None:
            self.file_name = file_name
            self.set_up_from_file(file_name)

    def set_up_from_file(self, file_name):
        self.file_name = file_name
        raw_data = self.get_raw_data(self.file_name)
        encoding = self._get_encoding(raw_data)
        file_text = self._decode_data(encoding, raw_data)
        self.set_up_from_text(file_text)

    def set_up_from_text(self, file_text):
        self.file_text = file_text
        self.raw_paragraphs = self.get_raw_paragraphs(file_text)
        self.cmp_paragraphs = self.get_cmp_paragraphs(self.raw_paragraphs)

    @staticmethod
    def get_raw_data(file_name):
        try:
            if file_name == '-':
                file_text = sys.stdin.buffer.read()
            else:
                file_text = open(file_name, 'rb').read()
            return file_text
        except BaseException:
            sys.stderr.write('error: bad file "' + file_name + '"\n')
            sys.exit(1)

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
            msg = '※ 警告: ' \
                + '文字コードを「SHIFT_JIS」に修正しました'
            # msg = 'warning: ' \
            #     + 'changed encoding to "SHIFT_JIS"'
            sys.stderr.write(msg + '\n\n')
        return encoding

    @staticmethod
    def _decode_data(encoding, raw_data):
        try:
            decoded_data = raw_data.decode(encoding)
        except BaseException:
            msg = '※ エラー: ' \
                + 'データを読みません（Markdownでないかも？）'
            # msg = 'error: ' \
            #     + 'can\'t read data (maybe not Markdown?)'
            sys.stderr.write(msg + '\n\n')
            raise BaseException('failed to read data')
            if __name__ == '__main__':
                sys.exit(105)
            return ''
        return decoded_data

    @staticmethod
    def get_raw_paragraphs(text):
        raw_paragraphs = []
        block = ''
        is_in_block = False
        for line in (text + '=').split('\n'):
            # CONFIGURATIONS
            if 'is_in_conf' not in locals():
                if re.match('^\\s*<!--.*$', line):
                    is_in_conf = True
                else:
                    is_in_conf = False
            if is_in_conf:
                block += line + '\n'
                if re.match('^.*-->\\s*$', line):
                    is_in_conf = False
                continue
            # BODY
            if not is_in_block and line != '':
                raw_paragraphs.append(block)
                block = ''
                is_in_block = True
            block += line + '\n'
            if re.match('^(.|\n)*\n\n$', block):
                is_in_block = False
        if block != '':
            raw_paragraphs.append(block)
        # CORRECTION OF LAST LINE
        raw_paragraphs[-1] = re.sub('=\n$', '', raw_paragraphs[-1])
        if raw_paragraphs[-1] == '':
            raw_paragraphs.pop(-1)
        return raw_paragraphs

    @staticmethod
    def get_cmp_paragraphs(raw_paragraphs):
        cmp_paragraphs = []
        for p in raw_paragraphs:
            p = re.sub('\n+$', '', p)
            cmp_paragraphs.append(p)
        return cmp_paragraphs

    @staticmethod
    def print_paragraphs(paragraphs):
        for i, p in enumerate(paragraphs):
            p = re.sub('\n+$', '', p)
            if i == 0 and p == '':
                continue
            elif i != len(paragraphs) - 1:
                print(p + '\n')
            else:
                print(p)

    @staticmethod
    def get_configs(paragraphs):
        configs = {}
        for item in ['版番号|version_number', '更新時|modified_time']:
            res = '^((?:.|\n)*\n(?:' + item + '):)(.*)(\n(?:.|\n)*)$'
            if re.match(res, paragraphs[0]):
                value = re.sub(res, '\\2', paragraphs[0])
                configs[item] = value
        return configs

    @staticmethod
    def set_configs(paragraphs, configs):
        for item in configs:
            res = '^((?:.|\n)*\n(?:' + item + '):)(.*)(\n(?:.|\n)*)$'
            txt = '\\1' + configs[item] + '\\3'
            if re.match(res, paragraphs[0]):
                paragraphs[0] = re.sub(res, txt, paragraphs[0])
        return paragraphs

    @staticmethod
    def reset_configs(paragraphs):
        for item in ['版番号|version_number', '更新時|modified_time']:
            res = '^((?:.|\n)*\n(?:' + item + '):)(.*)(\n(?:.|\n)*)$'
            txt = '\\1-\\3'
            if re.match(res, paragraphs[0]):
                paragraphs[0] = re.sub(res, txt, paragraphs[0])
        return paragraphs


class Paragraph:

    """A class to handle a paragraph"""

    # (main_paragraph)  >>>>(diff_id)>>>  (sub_paragraph)
    #                   <<<<(rev_id)<<<<
    #               (has_applied=False/True)
    #
    # (main_number=0)   (ses_symbor='.')  (sub_number=0)    (diff_text)
    # <!--                                <!--               | <!--
    # aaaa: bbbb                          aaaa: bbbb         | aaaa: bbbb
    # -->                                 -->                | -->
    #
    # (main_number=1)   (ses_symbor='.')  (sub_number=1)
    # cccccccccccccccc                    cccccccccccccccc   | cccccccccccccccc
    # cccccccccccccccc                    cccccccccccccccc   | cccccccccccccccc
    # cccccccccccccccc                    cccccccccccccccc   | cccccccccccccccc
    #
    # (main_number=2)   (ses_symbor='&')  (sub_number=2)
    # dddddddddddddddd                    dddddddddddddddd   | dddddddddddddddd
    # eeeeeeeeeeeeeeee                    ffffffffffffffff  x| eeeeeeeeeeeeeeee
    # dddddddddddddddd                    dddddddddddddddd  o| ffffffffffffffff
    #                                                        | dddddddddddddddd
    #
    # (main_number=3)   (ses_symbor='-')  (sub_number=-1)
    # gggggggggggggggg                                      X| gggggggggggggggg
    # gggggggggggggggg                                      X| gggggggggggggggg
    # gggggggggggggggg                                      X| gggggggggggggggg
    #
    # (main_number=-1)  (ses_symbor='+')  (sub_number=3)
    #                                     hhhhhhhhhhhhhhhh  O| hhhhhhhhhhhhhhhh
    #                                     hhhhhhhhhhhhhhhh  O| hhhhhhhhhhhhhhhh
    #                                     hhhhhhhhhhhhhhhh  O| hhhhhhhhhhhhhhhh

    def __init__(self, ses_symbol,
                 diff_id, rev_id, diff_text,
                 main_number, main_paragraph, sub_number, sub_paragraph):
        self.ses_symbol = ses_symbol
        self.has_applied = False
        self.diff_id = diff_id
        self.rev_id = rev_id
        self.diff_text = diff_text
        self.main_number = main_number
        self.main_paragraph = main_paragraph
        self.sub_number = sub_number
        self.sub_paragraph = sub_paragraph

    def get_current_paragraph(self):
        if not self.has_applied:
            return self.main_paragraph
        else:
            return self.sub_paragraph


class Comparison:

    """A class to compare paragraphs"""

    def __init__(self, strs_x, strs_y):
        dire_mx, dist_mx = Comparison.get_matrices(strs_x, strs_y)
        edit_distance = dist_mx[0][0]
        shortest_edit_script = self.get_shortest_edit_script(dire_mx)
        self.paragraphs \
            = self.get_paragraphs(shortest_edit_script, strs_x, strs_y)

    @staticmethod
    def get_matrices(strs_x, strs_y):
        numb_x = len(strs_x)
        numb_y = len(strs_y)
        dire_mx = [[0 for y in range(numb_y + 1)] for x in range(numb_x + 1)]
        dist_mx = [[0 for y in range(numb_y + 1)] for x in range(numb_x + 1)]
        dire_mx[numb_x][numb_y] = '/'
        dist_mx[numb_x][numb_y] = 0
        for x in range(numb_x - 1, -1, -1):
            dire_mx[x][numb_y] = '-'  # RIGHT -
            dist_mx[x][numb_y] = dist_mx[x + 1][numb_y] + len(strs_x[x])
        for y in range(numb_y - 1, -1, -1):
            dire_mx[numb_x][y] = '+'  # -     DOWN
            dist_mx[numb_x][y] = dist_mx[numb_x][y + 1] + len(strs_y[y])
        for x in range(numb_x - 1, -1, -1):
            for y in range(numb_y - 1, -1, -1):
                d11 = Levenshtein.distance(strs_x[x], strs_y[y])
                d10, d01 = len(strs_x[x]), len(strs_y[y])
                # MODIFY
                threshold = 4  # "aabb" "aacc" -> 2*4 = 4+4
                if d11 * threshold > d10 + d01:
                    d11 = d10 + d01 + 1
                t11 = dist_mx[x + 1][y + 1] + d11
                t10 = dist_mx[x + 1][y] + d10
                t01 = dist_mx[x][y + 1] + d01
                if d11 == 0:
                    dire_mx[x][y] = '.'  # RIGHT DOWN
                    dist_mx[x][y] = t11
                elif t11 <= t01 and t11 <= t10:
                    dire_mx[x][y] = '&'  # RIGHT DOWN
                    dist_mx[x][y] = t11
                elif t10 <= t01:
                    dire_mx[x][y] = '-'  # RIGHT -
                    dist_mx[x][y] = t10
                else:
                    dire_mx[x][y] = '+'  # -     DOWN
                    dist_mx[x][y] = t01
        return dire_mx, dist_mx

    @staticmethod
    def get_shortest_edit_script(dire_mx):
        shortest_edit_script = ''
        x = 0
        y = 0
        while True:
            d = dire_mx[x][y]
            if x == 0 and y == 0 and d != '.':
                d = '&'  # <- for configuration
            shortest_edit_script += d
            if d == '.' or d == '&' or d == '-':
                x += 1
            if d == '.' or d == '&' or d == '+':
                y += 1
            if dire_mx[x][y] == '/':
                break
        return shortest_edit_script

    def get_paragraphs(self, shortest_edit_script, strs_x, strs_y):
        paragraphs = []
        x, y, nx, ny = 0, 0, 0, 0
        strs_x += ['']
        strs_y += ['']
        for s in shortest_edit_script:
            diff_text = ''
            nx, ny = self._step_or_reset_nz(s, nx, ny)
            diff_id = Comparison._get_hash(x, ny, s, strs_x[x], strs_y[y])
            if s == '-':
                rev_id = Comparison._get_hash(y, nx, '+', strs_y[y], strs_x[x])
            elif s == '+':
                rev_id = Comparison._get_hash(y, nx, '-', strs_y[y], strs_x[x])
            else:
                rev_id = Comparison._get_hash(y, nx, s, strs_y[y], strs_x[x])
            if s == '.':
                pn = str(x) + '/' + str(y)
                diff_text += '【第' + pn + '段落】\n'
                diff_text += re.sub('\n', '\n | ', ' | ' + strs_y[y]) + '\n'
            elif s == '&':
                d11 = Levenshtein.distance(strs_x[x], strs_y[y])
                d10, d01 = len(strs_x[x]), len(strs_y[y])
                concordance_rate = round((1 - (d11 / (d10 + d01))) * 100, 2)
                pn = str(x) + '/' + str(y)
                diff_text += '【第' + pn + '段落】' \
                    + ' 編集（' + diff_id + '）' \
                    + '一致率=' + str(concordance_rate) + '%\n'
                if strs_x[x] != '' and strs_y[y] != '':
                    str_x, str_y = strs_x[x].split('\n'), strs_y[y].split('\n')
                    dire, dist = Comparison.get_matrices(str_x, str_y)
                    tx, ty = 0, 0
                    while True:
                        d = dire[tx][ty]
                        if d == '.':
                            diff_text += ' | ' + str_x[tx] + '\n'
                        elif d == '&':
                            diff_text += 'x| ' + str_x[tx] + '\n'
                            diff_text += 'o| ' + str_y[ty] + '\n'
                        elif d == '-':
                            diff_text += 'X| ' + str_x[tx] + '\n'
                        elif d == '+':
                            diff_text += 'O| ' + str_y[ty] + '\n'
                        tx, ty = self._step_z(d, tx, ty)
                        if dire[tx][ty] == '/':
                            break
                elif strs_x[x] != '':  # <- for configuration
                    diff_text += re.sub('\n', '\nX| ', 'X| ' + strs_x[x]) \
                        + '\n'
                elif strs_y[y] != '':  # <- for configuration
                    diff_text += re.sub('\n', '\nO| ', 'O| ' + strs_y[y]) \
                        + '\n'
            elif s == '-':
                pn = str(x) + '/' + str(y - 1) + '+' + str(nx)
                diff_text += '【第' + pn + '段落】' \
                    + ' 削除（' + diff_id + '）\n'
                diff_text += re.sub('\n', '\nX| ', 'X| ' + strs_x[x]) + '\n'
            elif s == '+':
                pn = str(x - 1) + '+' + str(ny) + '/' + str(y)
                diff_text += '【第' + pn + '段落】' \
                    + ' 追加（' + diff_id + '）\n'
                diff_text += re.sub('\n', '\nO| ', 'O| ' + strs_y[y]) + '\n'
            if s == '-':
                p = Paragraph(s, diff_id, rev_id, diff_text,
                              x, strs_x[x], -1, '')
            elif s == '+':
                p = Paragraph(s, diff_id, rev_id, diff_text,
                              -1, '', y, strs_y[y])
            else:
                p = Paragraph(s, diff_id, rev_id, diff_text,
                              x, strs_x[x], y, strs_y[y])
            paragraphs.append(p)
            x, y = self._step_z(s, x, y)
        return paragraphs

    @staticmethod
    def _step_or_reset_nz(s, nx, ny):
        if s == '-':
            nx += 1
        else:
            nx = 0
        if s == '+':
            ny += 1
        else:
            ny = 0
        return nx, ny

    @staticmethod
    def _step_z(s, x, y):
        if s == '.' or s == '&' or s == '-':
            x += 1
        if s == '.' or s == '&' or s == '+':
            y += 1
        return x, y

    @staticmethod
    def _get_hash(x, ny, s, str_x, str_y):
        if s == '-':
            str_y = ''
        if s == '+':
            str_x = ''
        s = SALT + '\n' \
            + str(x) + '+' + str(ny) + '(' + s + ')\n' \
            + str_x + '\n\n' + str_y
        return hashlib.md5(s.encode()).hexdigest()

    def print_current_paragraphs(self):
        m = len(self.paragraphs) - 1
        for i, p in enumerate(self.paragraphs):
            par = p.get_current_paragraph()
            if par != '':
                if i < m:
                    print(par + '\n')
                else:
                    print(par)

    def apply_difference(self, diff_id):
        for p in self.paragraphs:
            if p.diff_id == diff_id:
                p.has_applied = True
                return 0
        return 1

    def print_reverse_id(self, diff_id):
        for p in self.paragraphs:
            if p.diff_id == diff_id:
                print(p.rev_id)
                return 0
        return 1

    def print_diff_html(self):
        retval = 0
        print('<pre style="border:3pt solid; padding:9pt; font-size:18px;">'
              + '<code>\n')
        for p in self.paragraphs:
            if p.ses_symbol != '.':
                retval = 1
                lines = p.diff_text.split('\n')
                lines.pop(0)
                for line in lines:
                    if len(line) <= 3:
                        continue
                    c = line[0]
                    if c == ' ':
                        self._print_html_line('.', line[3:])
                    elif c == 'x' or c == 'X':
                        self._print_html_line('-', line[3:])
                    elif c == 'o' or c == 'O':
                        self._print_html_line('+', line[3:])
                self._print_button(p.diff_id)
                print('')
        print('</code></pre>')
        return retval

    @staticmethod
    def _print_html_line(s, line):
        line = re.sub('&', '&amp;', line)
        line = re.sub('<', '&lt;', line)
        line = re.sub('>', '&gt;', line)
        line = re.sub('"', '&quot;', line)
        if s == '-':
            print('<span style="color:red;">- ' + line + '</span>')
        elif s == '+':
            print('<span style="color:blue;">+ ' + line + '</span>')
        else:
            print('<span style="color:black;">  ' + line + '</span>')

    @staticmethod
    def _print_button(hash):
        print('<form action="{{ URL }}">' +
              '<input type="submit" value="適用する" id="' + hash + '" />' +
              '</form>')

    def print_all_diff_text(self):
        retval = 0
        for p in self.paragraphs:
            if p.ses_symbol != '.':
                retval = 1
            print(p.diff_text)
        return retval

    def print_diff_text(self):
        retval = 0
        for p in self.paragraphs:
            if p.ses_symbol != '.':
                retval = 1
                print(p.diff_text)
        return retval


def main():
    args = get_arguments()
    main_file = File(args.main_md_file)
    sub_file = File(args.sub_md_file)
    # PUT OUT CONFIGS >>>
    configs = File.get_configs(main_file.raw_paragraphs)
    main_file.cmp_paragraphs = File.reset_configs(main_file.cmp_paragraphs)
    sub_file.cmp_paragraphs = File.reset_configs(sub_file.cmp_paragraphs)
    # <<<
    comp = Comparison(main_file.cmp_paragraphs, sub_file.cmp_paragraphs)
    # PUT IN CONFIGS >>>
    comp.paragraphs[0].main_paragraph \
        = File.set_configs([comp.paragraphs[0].main_paragraph], configs)[0]
    comp.paragraphs[0].sub_paragraph \
        = File.set_configs([comp.paragraphs[0].sub_paragraph], configs)[0]
    # <<<
    if args.apply_difference is not None:
        diff_id = args.apply_difference
        retval = comp.apply_difference(diff_id)
        comp.print_current_paragraphs()
        sys.exit(retval)
    elif args.print_reverse_id is not None:
        diff_id = args.print_reverse_id
        retval = comp.print_reverse_id(diff_id)
        sys.exit(retval)
    elif args.html:
        retval = comp.print_diff_html()
        sys.exit(retval)
    elif args.verbose:
        retval = comp.print_all_diff_text()
        sys.exit(retval)
    else:
        retval = comp.print_diff_text()
        sys.exit(retval)


if __name__ == '__main__':
    main()
