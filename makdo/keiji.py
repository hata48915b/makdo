#!/usr/bin/python3
# Name:         keiji.py
# Version:      v02
# Time-stamp:   <2025.08.11-16:53:12-JST>

# keiji.py
# Copyright (C) 2017-2025  Seiichiro HATA
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


############################################################
# SETTING


import os
import sys
import getopt
import re
import datetime
import locale
from decimal import Decimal


__version__ = 'v02'


GENGO = [['M', datetime.date(1868,  1, 25)],  # MEIJI
         ['T', datetime.date(1912,  7, 30)],  # TAISHO
         ['S', datetime.date(1926, 12, 25)],  # SHOWA
         ['H', datetime.date(1989,  1,  8)],  # HEISEI
         ['R', datetime.date(2019,  5,  1)],  # REIWA
         ]


STATUTORY_RATE = [['5', datetime.date(1,     1,  1)],
                  ['3', datetime.date(2020,  4,  1)],
                  ]


RESTRICTED_RATE = [['20',       '0'],
                   ['18',  '100000'],
                   ['15', '1000000']]


TABLE_HEADER = ['日付',
                '借入',
                '返済',
                '年利',
                '日数',
                '利息',
                '増減',
                '残利息',
                '残元金']


TABLE_FOOTER = ['',
                '',
                '',
                '',
                '',
                '',
                '',
                '【合計】',
                '$total']


TABLE_WIDTH = [9, 9, 8, 4, 4, 8, 9, 9, 10]


HELP_MESSAGE = '''\
Usage: keiji [オプション]... [ファイル]

取引履歴から利息等を計算します

オプション
  -h, --help         この使い方を表示して終了します
  -v, --version      このプログラムのバージョンを表示して終了します
  -e, --exclude      初日を参入しません(デフォルトは算入する)
  -d, --daily        年未満は閏年を考慮せずに日歩で計算します
  -D, --Daily        全期間を閏年を考慮せずに日歩で計算します
  -n, --no-comma     桁区切りのコンマを出力しません
  -3  --3jc          日付を和暦(元号イニシャル+2桁年数)で出力します
  -2  --2wc          日付を西暦2桁年数で出力します
  -4  --4wc          日付を西暦4桁年数で出力します
  -t, --tex          LaTeX形式で出力します
  -c, --csv          CSV形式で出力します
  -w, --web          HTML形式で出力します
  -m, --md           Markdown形式で出力します
  -M, --math         数式形式で出力します
      --sample       サンプルデータを出力します
      --debug        デバッグのためのメッセージを出力します

ファイルの指定がなかったり、"-"であった場合、標準入力から読み込みます

入力データの仕様は、次のとおりです
  日付 借入額 返済額 年利 設定
    日付：    年月日の区切りは、"-"を使います
              元号は"MTSHR"が使えます。
              元号がない場合、69未満は2000年代、70以上100未満は1900年代です
              元号がない場合の100以上は、そのまま西暦です
              例：  H01-02-03（平成元年2月3日）
                     69-02-03（  2069年2月3日）
                     70-02-03（  1970年2月3日）
                   2001-02-03（  2001年2月3日）
    借入額：  金額又は"-"を指定してください
              桁区切りの","があっても構いません
              金額の先頭に"_"を付けると、利息ではなく、元金に充当します
              例：   10000（過払利息、過払元金の順に、1万円を充当します）
                    _10000（過払元金に、1万円を充当します）
    返済額：  金額又は"-"を指定してください
              桁区切りの","があっても構いません
              金額の先頭に"_"を付けると、利息ではなく、元金に充当します
              例：   10000（利息、元金の順に、1万円を充当します）
                    _10000（元金に、1万円を充当します）
    年利：    パーセント単位で指定してください
              "="にすると、利息制限法及び民法所定の利率になります
              "*"にすると、遅延損害金となり1.46倍になります
              2行目以下で指定がない場合は、前の設定を引き継ぎます
    設定：    初日だけ指定できます
              "+"を指定すると初日参入です（デフォルトと同じ）
              "-"を指定すると初日不参入です（"-e"又は"--exclude"と同じ）
              "?"を指定すると年未満を日歩計算します（"-d"又は"--daily"と同じ）
              "!"を指定すると全期間を日歩計算します（"-D"又は"--Daily"と同じ）
              "t"を指定するとLaTeX形式で出力します（"-t"又は"--tex"と同じ）
              "c"を指定するとCSV形式で出力します（"-c"又は"--csv"と同じ）
              "w"を指定するとHTML形式で出力します（"-w"又は"--web"と同じ）
              "m"を指定するとMarkdown形式で出力します（"-m"又は"--markdown"と同じ）
              ";"を指定すると日付を和暦で出力します（"-3"又は"--3jc"と同じ）
              "."を指定すると日付を西暦2桁で出力します（"-2"又は"--2wc"と同じ）
              ":"を指定すると日付を西暦4桁で出力します（"-4"又は"--4wc"と同じ）
              "_"を指定すると桁区切りのコンマを出力しません

  例：  S64-01-02    10,000        -    = +!
        H01-02-03         -    5,000   5%
        H01-03-04         -   _5,000   5%

行頭の"#"を除き、行の"#"以降は、無視されますので、コメントになります

"---"を含む行は、無視されますので、コメント行に使えます

年計算の場合、利息の計算は、次のとおりです（東京地裁等の取扱い）
  年単位にできるものは年単位にし、残りは非閏年と閏年に分けます
  年単位は年利で、非閏年は365日の日割り、閏年は366日の日割りで計算します
  例：  H02-09-01 ------> H03-08-31 ------> H04-01-01 ------> H04-05-01
        年利 -----------> 日割(非閏年) ---> 日割(閏年) ----->
日歩計算の場合、利息制限法の利率を指定しても、同法に違反する場合があります

よくわからない方は、「keiji --sample | keiji」を実行してみてください\
'''


MANUAL_FOR_MAKDO = '''\
<!--
各行は「日付 借入額 返済額 年利 設定」から構成されます
- 日付
    69未満は2000年代、70以上は1900年代
    M??は明治、T??は大正、S??は昭和、H??は平成、R??は令和
- 借入額
    先頭の"_"に付けると元金に充当
- 返済額
    先頭の"_"に付けると元金に充当
- 年利
    "="は利息制限法及び民法所定の金利
    "*"は遅延損害金で1.46倍
- 設定
    - 初日算入
        "+"は初日算入（デフォルト）
        "-"は初日不算入
    - 閏年の扱い
        "?"は年未満を日歩計算
        "!"は全期間を日歩計算
    - 日付出力
        ";"は日付を和暦で出力
        "."は日付を西暦2桁で出力
        ":"は日付を西暦4桁で出力
-->\
'''


SAMPLE_DATA = '''\
H24-01-01    10,000        - = +
H25-01-01   100,000        -
H26-01-01 1,000,000        -
H27-01-01         -  300,000
H28-01-01         -  300,000
H29-01-01         -  300,000
H30-01-01         -  300,000
H31-01-01         -  300,000
R02-01-01         -  300,000\
'''


WARNING_MESSAGE_NUMBER_OF_DAYS_IS_NEGATIVE \
    = \
    'WARNING: The number of days is negative. ' + \
    '-------------------------------------'


# LOCALE
locale.setlocale(locale.LC_NUMERIC, 'ja_JP.UTF-8')


############################################################
# FUNCTIONS


def width(string):
    if(sys.version_info[0] == 2):
        st = string.decode('utf-8')
    else:
        st = string
    le = 0
    for c in st:
        if(re.match('^[ -~]$', c)):
            le += 1
        else:
            le += 2
    return le


def to_date(date):
    if(not (isinstance(date, str))):
        sys.stderr.write('bad date type "' + str(date) + '"\n')
        return None
    [sgy, sm, sd] = date.replace('.', '-').split('-')
    if(re.match('^[A-Z]', sgy)):
        # JAPANESE CALENDER
        for i in range(len(GENGO)):
            if(re.match(GENGO[i][0], sgy)):
                ny = int(re.sub(GENGO[i][0], '', sgy)) + \
                     GENGO[i][1].year - 1
        if(not ('ny' in locals())):
            sys.stderr.write('no gengo "' + str(date) + '"\n')
            return None
    else:
        # WESTERN CALENDER
        if(int(sgy) < 70):        # 20XX
            ny = int(sgy) + 2000
        elif(int(sgy) < 100):     # 19XX
            ny = int(sgy) + 1900
        else:                     # XXXX
            ny = int(sgy)
    nm = int(sm)
    nd = int(sd)
    try:
        return datetime.date(ny, nm, nd)
    except BaseException:
        return None


def to_str(date, style='3jc'):
    if(not (isinstance(date, datetime.date))):
        sys.stderr.write('bad date type "' + str(date) + '"\n')
        return None
    sg = ''
    ny = date.year
    if(style == '3jc'):
        # JAPANESE CALENDER
        for i in range(len(GENGO)):
            if(date >= GENGO[i][1]):
                sg = GENGO[i][0]
                ny = date.year - GENGO[i][1].year + 1
        if(sg == ''):
            sys.stderr.write('no gengo "' + str(date) + '"\n')
            return None
    elif(style == '2wc'):
        # WESTERN CALENDER (2digit)
        if(ny >= 1970):
            if(ny < 2070):
                ny = ny % 100
    sy = str(ny).zfill(2)
    sm = str(date.month).zfill(2)
    sd = str(date.day).zfill(2)
    return sg + sy + '-' + sm + '-' + sd


def count_years_and_days(first_day, last_day,
                         has_to_include_first_day=False,
                         has_to_include_last_day=True):
    # PREPARE
    if(has_to_include_first_day):
        fo = to_date(first_day) - datetime.timedelta(days=1)
    else:
        fo = to_date(first_day)
    if(has_to_include_last_day):
        lo = to_date(last_day)
    else:
        lo = to_date(last_day) - datetime.timedelta(days=1)
    fy = fo.year
    fm = fo.month
    fd = fo.day
    ly = lo.year
    lm = lo.month
    ld = lo.day
    # COUNT
    ds = str((lo - fo).days)  # days
    if((lm * 100 + ld) - (fm * 100 + fd) >= 0):
        # SAME YEAR
        if((fm == 2) and (fd == 29) and (not is_leap_year(ly))):
            bi = datetime.date(ly, 2, 28)
        else:
            bi = datetime.date(ly, fm, fd)
        ys = str(ly - fy)  # years
        if(is_leap_year(ly)):
            dn = '0'                  # days in normal year
            dl = str((lo - bi).days)  # days in leap year
        else:
            dn = str((lo - bi).days)  # days in normal year
            dl = '0'                  # days in leap year
    else:
        # OVER YEAR
        if((fm == 2) and (fd == 29) and (not is_leap_year(ly - 1))):
            bi = datetime.date(ly - 1, 2, 28)
        else:
            bi = datetime.date(ly - 1, fm, fd)
        ys = str(ly - fy - 1)  # years
        ei = datetime.date(ly - 1, 12, 31)
        if(is_leap_year(ly - 1)):
            dn = str((lo - ei).days)  # days in normal year
            dl = str((ei - bi).days)  # days in leap year
        elif(is_leap_year(ly)):
            dn = str((ei - bi).days)  # days in normal year
            dl = str((lo - ei).days)  # days in leap year
        else:
            dn = str((lo - bi).days)  # days in normal year
            dl = '0'                  # days in leap year
    # RETURN
    return ds, ys, dn, dl


def is_leap_year(year):
    if((year % 400) == 0):
        return True
    elif((year % 100) == 0):
        return False
    if((year % 4) == 0):
        return True


def get_statutory_rate(this_principal, prev_principal='0',
                       prev_rate='0',
                       date=datetime.date.today().strftime("%Y-%m-%d"),
                       this_standard='=', prev_standard='='):
    # CHECK
    if((this_standard != '=') and (this_standard != '*')):
        sys.stderr.write('bad standard "' + str(this_standard) + '"\n')
        if __name__ == '__main__':
            sys.exit(1)
    # PREPARE
    tp = Decimal(this_principal)
    pp = Decimal(prev_principal)
    pr = Decimal(prev_rate)
    da = to_date(date)
    ts = this_standard
    ps = prev_standard
    ch = True
    # JUDGE
    if(tp < 0):
        # OVERPAYMENT CASE
        if(((ps != '=') and (ps != '*')) or (pp >= 0)):
            for i in STATUTORY_RATE:
                if(da >= i[1]):
                    tr = Decimal(i[0])
        else:
            tr = pr
            ch = False
    elif(tp == 0):
        # FULL PAYMENT CASE
        tr = Decimal('0')
    else:
        # NORMAL CASE
        for i in reversed(RESTRICTED_RATE):
            am = Decimal(i[1])
            if((tp >= am) and
               (((ps != '=') and (ps != '*')) or ((pp == 0) or (pp < am)))):
                tr = Decimal(i[0])
                break  # Do not remove
        else:
            tr = pr
            ch = False
    # DELAYED DAMAGE
    if((ts == '*') and ((ch) or (ps != '*'))):
        tr = tr * Decimal('1.0000') * Decimal('1.46')
    elif((ts != '*') and (not ch) and (ps == '*')):
        tr = tr * Decimal('1.0000') / Decimal('1.46')
    # RETURN
    sr = str(tr)
    if('.' in sr):
        sr = sr.rstrip('0').rstrip('.')
    return sr


def calculate_interest(principal, interest_rate,
                       days, years, days_in_normal_year, days_in_leap_year,
                       calculating_unit='yearly'):
    pr = Decimal(principal)
    ir = Decimal(interest_rate)
    da = Decimal(days)
    yr = Decimal(years)
    dn = Decimal(days_in_normal_year)
    dl = Decimal(days_in_leap_year)
    if(calculating_unit == 'daily'):
        # DAILY TOTAL
        ny = da / Decimal(365)
    elif(calculating_unit == 'yearly+daily'):
        # DAILY PARTIAL
        ny = yr + dn / Decimal(365) + dl / Decimal(365)
    else:
        # YEAR
        ny = yr + dn / Decimal(365) + dl / Decimal(366)
    it = pr * ir * ny / Decimal(100)
    return str(int(it))


############################################################
# CLASS


class Trade:
    """A class of a trade"""

    ##############################################
    # CLASS FUNCTION AND VARIABLE

    ####################################
    # _ERROR

    @classmethod
    def _error(cls, _message):
        sys.stderr.write(_message + '\n')
        if __name__ == '__main__':
            sys.exit(1)

    ####################################
    # TOTAL AMOUNT

    _total_amount = Decimal(0)

    @classmethod
    def set_total_amount(cls, amount):
        if(re.match('^-?[0-9]+$', amount)):
            cls._total_amount = amount
        else:
            cls._error('bad total amount "' + str(amount) + '"')

    @classmethod
    def get_total_amount(cls):
        return cls._total_amount

    ####################################
    # COMMENT OUT SYMBOL

    _comment_out_symbol = ''

    @classmethod
    def set_comment_out_symbol(cls, symbol):
        if(re.match(r'^([^0-9a-zA-Z]+\s*)?$', symbol)):
            cls._comment_out_symbol = symbol
        else:
            cls._error('bad comment out symbol "' + str(symbol) + '"')

    @classmethod
    def get_comment_out_symbol(cls):
        return cls._comment_out_symbol

    ####################################
    # SHOULD INCLUDE FIRST DAY

    _should_include_first_day = True

    @classmethod
    def set_should_include_first_day(cls, should):
        if(isinstance(should, bool)):
            cls._should_include_first_day = should
        else:
            cls._error('bad should include first day "' + str(should) + '"')

    @classmethod
    def should_include_first_day(cls):
        return cls._should_include_first_day

    ####################################
    # DATE STYLE

    _output_date_style = ''

    @classmethod
    def set_output_date_style(cls, style):
        if(style == ''):
            return
        elif(style == '3jc'):
            cls._output_date_style = '3jc'
        elif(style == '2wc'):
            cls._output_date_style = '2wc'
        elif(style == '4wc'):
            cls._output_date_style = '4wc'
        else:
            cls._error('bad output_date_style "' + str(unit) + '"')

    @classmethod
    def get_output_date_style(cls):
        return cls._output_date_style

    ####################################
    # CALCULATING UNIT

    _calculating_unit = 'yearly'

    @classmethod
    def set_calculating_unit(cls, unit):
        if(unit == 'year'):
            cls._calculating_unit = 'yearly'
        elif(unit == 'yearly+daily'):
            cls._calculating_unit = 'yearly+daily'
        elif(unit == 'daily'):
            cls._calculating_unit = 'daily'
        else:
            cls._error('bad calculatign unit "' + str(unit) + '"')

    @classmethod
    def get_calculating_unit(cls):
        return cls._calculating_unit

    ####################################
    # SHOULD OUTPUT COMMA

    _should_insert_comma = True

    @classmethod
    def set_should_insert_comma(cls, should):
        if(isinstance(should, bool)):
            cls._should_insert_comma = should
        else:
            cls._error('bad should output comma "' + str(should) + '"')

    @classmethod
    def should_insert_comma(cls):
        return cls._should_insert_comma

    ####################################
    # OUTPUT FORMAT

    _output_style = 'text'

    @classmethod
    def set_output_style(cls, style):
        if(style == 'text'):
            cls._output_style = 'text'
        elif(style == 'tex'):
            cls._output_style = 'tex'
        elif(style == 'csv'):
            cls._output_style = 'csv'
        elif(style == 'web'):
            cls._output_style = 'web'
        elif(style == 'markdown'):
            cls._output_style = 'markdown'
        elif(style == 'math'):
            cls._output_style = 'math'
        else:
            cls._error('bad output style "' + str(style) + '"')

    @classmethod
    def get_output_style(cls):
        return cls._output_style

    ####################################
    # PRINT

    @classmethod
    def get_header(cls):
        of = cls.get_output_style()
        th = TABLE_HEADER
        tw = TABLE_WIDTH
        header = ''
        if(of == 'math'):
            return ''
        elif(of == 'tex'):
            return ('\\begin{tabular}{lrrrrrrrrl}\n' +
                    th[0] + '&' +  # 日付
                    th[1] + '&' +  # 借入
                    th[2] + '&' +  # 返済
                    th[3] + '&' +  # 年利
                    th[4] + '&' +  # 日数
                    th[5] + '&' +  # 利息
                    th[6] + '&' +  # 増減
                    th[7] + '&' +  # 残利息
                    th[8] + '&' +  # 残元金
                    '\\\\')
        elif(of == 'csv'):
            return ('"' + th[0] + '",' +  # 日付
                    '"' + th[1] + '",' +  # 借入
                    '"' + th[2] + '",' +  # 返済
                    '"' + th[3] + '",' +  # 年利
                    '"' + th[4] + '",' +  # 日数
                    '"' + th[5] + '",' +  # 利息
                    '"' + th[6] + '",' +  # 増減
                    '"' + th[7] + '",' +  # 残利息
                    '"' + th[8] + '",' +  # 残元金
                    '""')
        elif(of == 'web'):
            return ('<table>\n' +
                    '<tr>' +
                    '<th>' + th[0] + '</th>' +  # 日付
                    '<th>' + th[1] + '</th>' +  # 借入
                    '<th>' + th[2] + '</th>' +  # 返済
                    '<th>' + th[3] + '</th>' +  # 年利
                    '<th>' + th[4] + '</th>' +  # 日数
                    '<th>' + th[5] + '</th>' +  # 利息
                    '<th>' + th[6] + '</th>' +  # 増減
                    '<th>' + th[7] + '</th>' +  # 残利息
                    '<th>' + th[8] + '</th>' +  # 残元金
                    '<th></th>' +
                    '</tr>')
        elif(of == 'markdown'):
            ds = cls.get_output_date_style()
            dw = tw[0]
            if(ds == '2wc'):
                th_0 = '' + th[0] + ''
                cf_0 = '-------:'
                dw -= 1
            elif(ds == '4wc'):
                th_0 = ' ' + th[0] + ' '
                cf_0 = '---------:'
                dw += 1
            else:
                th_0 = '' + th[0] + ' '
                cf_0 = '--------:'
            return ('|  ' + th_0 + '  ' +    # 日付
                    '|  ' + th[1] + '   ' +  # 借入
                    '|  ' + th[2] + '  ' +   # 返済
                    '|' + th[3] + '' +       # 年利
                    '|' + th[4] + '' +       # 日数
                    '|  ' + th[5] + '  ' +   # 利息
                    '|  ' + th[6] + '   ' +  # 増減
                    '| ' + th[7] + '  ' +    # 残利息
                    '|  ' + th[8] + '  ' +   # 残元金
                    '|\n' +
                    '|' + cf_0 +
                    '|--------:' +
                    '|-------:' +
                    '|---:' +
                    '|---:' +
                    '|-------:' +
                    '|--------:' +
                    '|--------:' +
                    '|---------:' +
                    '|\n' +
                    '=' * (71 + dw))
        else:
            ds = cls.get_output_date_style()
            dw = tw[0]
            if(ds == '2wc'):
                dw -= 1
            elif(ds == '4wc'):
                dw += 1
            return (cls.get_comment_out_symbol() +
                    th[0] + '-' * (dw - width(th[0])) + ' ' +     # 日付
                    th[1] + '-' * (tw[1] - width(th[1])) + ' ' +  # 借入
                    th[2] + '-' * (tw[2] - width(th[2])) + ' ' +  # 返済
                    th[3] + '-' * (tw[3] - width(th[3])) + ' ' +  # 年利
                    th[4] + '-' * (tw[4] - width(th[4])) + ' ' +  # 日数
                    th[5] + '-' * (tw[5] - width(th[5])) + ' ' +  # 利息
                    th[6] + '-' * (tw[6] - width(th[6])) + ' ' +  # 増減
                    th[7] + '-' * (tw[7] - width(th[7])) + ' ' +  # 残利息
                    th[8] + '-' * (tw[8] - width(th[8])))         # 残元金

    @classmethod
    def get_footer(cls):
        of = cls.get_output_style()
        tf = TABLE_FOOTER
        tw = TABLE_WIDTH
        fa = cls.get_total_amount()
        if(cls.should_insert_comma()):
            fa = locale.format_string('%d', int(fa), True)
        for i, t in enumerate(tf):
            tf[i] = t.replace('$total', fa)
        if(of == 'math'):
            return ''
        elif(of == 'tex'):
            return (tf[0] + '&' +
                    tf[1] + '&' +
                    tf[2] + '&' +
                    tf[3] + '&' +
                    tf[4] + '&' +
                    tf[5] + '&' +
                    tf[6] + '&' +
                    tf[7] + '&' +  # 【合計】
                    tf[8] + '&\\\\\n' +
                    '\\end{tabular}')
        elif(of == 'csv'):
            return ('"","","","","","","","合計","' + fa + '",""')
        elif(of == 'web'):
            return ('<tr>' +
                    '<td>' + tf[0] + '</td>' +
                    '<td>' + tf[1] + '</td>' +
                    '<td>' + tf[2] + '</td>' +
                    '<td>' + tf[3] + '</td>' +
                    '<td>' + tf[4] + '</td>' +
                    '<td>' + tf[5] + '</td>' +
                    '<td>' + tf[6] + '</td>' +
                    '<td>' + tf[7] + '</td>' +  # 【合計】
                    '<td>' + tf[8] + '</td>' +
                    '<td></td>' +
                    '</tr>\n' +
                    '</table>')
        elif(of == 'markdown'):
            ds = cls.get_output_date_style()
            dw = tw[0]
            if(ds == '2wc'):
                dw -= 1
            elif(ds == '4wc'):
                dw += 1
            tf_7 = '合計'
            return ('=' * (71 + dw) + '\n' +
                    '|' + tf[0] + ' ' * (dw - width(tf[0])) +
                    '|' + tf[1] + ' ' * (tw[1] - width(tf[1])) +
                    '|' + tf[2] + ' ' * (tw[2] - width(tf[2])) +
                    '|' + tf[3] + ' ' * (tw[3] - width(tf[3])) +
                    '|' + tf[4] + ' ' * (tw[4] - width(tf[4])) +
                    '|' + tf[5] + ' ' * (tw[5] - width(tf[5])) +
                    '|' + tf[6] + ' ' * (tw[6] - width(tf[6])) +
                    '|' + ' ' * (tw[7] - width(tf_7)) + tf_7 +  # 合計
                    '|' + ' ' * (tw[8] - len(tf[8])) + tf[8] +
                    '|')
        else:
            ds = cls.get_output_date_style()
            dw = tw[0]
            if(ds == '2wc'):
                dw -= 1
            elif(ds == '4wc'):
                dw += 1
            return (cls.get_comment_out_symbol() +
                    tf[0] + '-' * (dw - width(tf[0])) + ' ' +
                    tf[1] + '-' * (tw[1] - width(tf[1])) + ' ' +
                    tf[2] + '-' * (tw[2] - width(tf[2])) + ' ' +
                    tf[3] + '-' * (tw[3] - width(tf[3])) + ' ' +
                    tf[4] + '-' * (tw[4] - width(tf[4])) + ' ' +
                    tf[5] + '-' * (tw[5] - width(tf[5])) + ' ' +
                    tf[6] + '-' * (tw[6] - width(tf[6])) + ' ' +
                    tf[7] + ' ' * (tw[7] - width(tf[7])) + ' ' +  # 【合計】
                    ' ' * (tw[8] - len(tf[8])) + tf[8])

    ##############################################
    # INSTANCE FUNCTION AND VARIABLE

    ####################################
    # HEAD SYMBOL

    def set_head_symbol(self, line):
        fo = re.search(re.compile(r'^[^0-9a-zA-Z\\|]+\s*'), line)
        if(fo):
            self._head_symbol = fo.group(0)
        else:
            self._head_symbol = ''

    def get_head_symbol(self):
        return self._head_symbol

    ####################################
    # DATE

    def set_input_date_style(self, date):
        date = date.replace('.', '-')
        mw = r'^[A-Z]?[0-9]{1,4}-[0-9]{1,2}-[0-9]{1,2}$'  # G?YY.MM.DD
        if(not re.match(mw, date)):
            self._error('bad date "' + str(date) + '"')
        if(re.match('^[A-Z]', date)):
            self._input_date_style = '3jc'
        elif(re.match('^[0-9]{1,2}', date)):
            self._input_date_style = '2wc'
        else:
            self._input_date_style = '4wc'

    def get_input_date_style(self):
        return self._input_date_style

    def set_this_date(self, date):
        date = date.replace('.', '-')
        mw = r'^[A-Z]?[0-9]{1,4}-[0-9]{1,2}-[0-9]{1,2}$'  # G?YY.MM.DD
        if(not re.match(mw, date)):
            self._error('bad this date "' + str(date) + '"')
        self._this_date = to_date(date)

    def get_this_date(self, style='3jc'):
        return to_str(self._this_date, style)

    def set_prev_date(self, _date):
        _date = _date.replace('.', '-')
        mw = r'^[A-Z]?[0-9]{1,4}-[0-9]{1,2}-[0-9]{1,2}$'  # G?YY.MM.DD
        if(not re.match(mw, _date)):
            self._error('bad prev date "' + str(_date) + '"')
        self._prev_date = to_date(_date)

    def get_prev_date(self, style='3jc'):
        return to_str(self._prev_date, style)

    ####################################
    # BORROWING AND REPAYMENT AMOUNT

    def _shape_amount(self, amount):
        amount = re.sub('^_', '', amount)
        if(amount == '-'):
            return Decimal('0')
        else:
            return Decimal(amount.replace(',', ''))

    def set_applies_borrowing_to_principal(self, amount):
        mw = '^(-)|(_?[0-9]{1,3}(,?[0-9]{3})*)$'
        if(not re.match(mw, amount)):
            self._error('bad borrowing amount "' + str(amount) + '"')
        if(re.match('^_', amount)):
            self._applies_borrowing_to_principal = True
        else:
            self._applies_borrowing_to_principal = False

    def applies_borrowing_to_principal(self):
        return self._applies_borrowing_to_principal

    def set_this_borrowing_amount(self, amount):
        mw = '^(-)|(_?[0-9]{1,3}(,?[0-9]{3})*)$'
        if(not re.match(mw, amount)):
            self._error('bad this borrowing amount "' + str(amount) + '"')
        self._this_borrowing_amount = self._shape_amount(amount)

    def get_this_borrowing_amount(self):
        return str(self._this_borrowing_amount)

    def set_prev_borrowing_amount(self, amount):
        mw = '^(-)|(_?[0-9]{1,3}(,?[0-9]{3})*)$'
        if(not re.match(mw, amount)):
            self._error('bad prev borrowing amount "' + str(amount) + '"')
        self._prev_borrowing_amount = self._shape_amount(amount)

    def get_prev_borrowing_amount(self):
        return str(self._prev_borrowing_amount)

    def set_applies_repayment_to_principal(self, amount):
        mw = '^(-)|(_?[0-9]{1,3}(,?[0-9]{3})*)$'
        if(not re.match(mw, amount)):
            self._error('bad repayment amount "' + str(amount) + '"')
        if(re.match('^_', amount)):
            self._applies_repayment_to_principal = True
        else:
            self._applies_repayment_to_principal = False

    def applies_repayment_to_principal(self):
        return self._applies_repayment_to_principal

    def set_this_repayment_amount(self, amount):
        mw = '^(-)|(_?[0-9]{1,3}(,?[0-9]{3})*)$'
        if(not re.match(mw, amount)):
            self._error('bad this repayment amount "' + str(amount) + '"')
        self._this_repayment_amount = self._shape_amount(amount)

    def get_this_repayment_amount(self):
        return str(self._this_repayment_amount)

    def set_prev_repayment_amount(self, amount):
        mw = '^(-)|(_?[0-9]{1,3}(,?[0-9]{3})*)$'
        if(not re.match(mw, amount)):
            self._error('bad prev repayment amount "' + str(amount) + '"')
        self._prev_repayment_amount = self._shape_amount(amount)

    def get_prev_repayment_amount(self):
        return str(self._prev_repayment_amount)

    ####################################
    # INTEREST RATE

    def set_this_interest_rate_standard(self, rate):
        mw = r'^((=)|(\*)|(((=)|(\*))?[0-9]+(\\.[0-9]+)?%?))?$'
        if(not re.match(mw, rate)):
            self._error('bad this interest rate standard "' + str(rate) + '"')
        if(re.match('^=', rate)):
            self._this_interest_rate_standard = '='
        elif(re.match(r'^\*', rate)):
            self._this_interest_rate_standard = '*'
        else:
            self._this_interest_rate_standard = rate.rstrip('%')

    def get_this_interest_rate_standard(self):
        return self._this_interest_rate_standard

    def set_prev_interest_rate_standard(self, rate):
        mw = r'^(=)|(\*)|([0-9]+(\\.[0-9]+)?%?)$'
        if(not re.match(mw, rate)):
            self._error('bad prev interest rate standard "' + str(rate) + '"')
        self._prev_interest_rate_standard = rate

    def get_prev_interest_rate_standard(self):
        return self._prev_interest_rate_standard

    def check_and_set_this_interest_rate_standard(self):
        ts = self.get_this_interest_rate_standard()
        if(ts == ''):
            ps = self.get_prev_interest_rate_standard()
            self.set_this_interest_rate_standard(ps)

    def set_this_interest_rate(self, rate):
        mw = r'^((=)|(\*)|(((=)|(\*))?[0-9]+(\.[0-9]+)?%?))?$'
        if(not re.match(mw, rate)):
            self._error('bad this interest rate "' + str(rate) + '"')
        rate = rate.lstrip('=')
        rate = rate.lstrip('*')
        rate = rate.rstrip('%')
        if(rate != ''):
            self._this_interest_rate = Decimal(rate)

    def get_this_interest_rate(self):
        return str(self._this_interest_rate)

    def set_prev_interest_rate(self, rate):
        mw = r'^[0-9]+(\.[0-9]+)?%?$'
        if(not re.match(mw, rate)):
            self._error('bad prev interest rate "' + str(rate) + '"')
        self._prev_interest_rate = Decimal(rate)

    def get_prev_interest_rate(self):
        return str(self._prev_interest_rate)

    def calc_and_set_this_interest_rate(self):
        ts = self.get_this_interest_rate_standard()
        ps = self.get_prev_interest_rate_standard()
        if(ts == ''):
            ts = ps
            self.set_this_interest_rate_standard(ts)
        if((ts == '=') or (ts == '*')):
            tp = self.get_this_remaining_principal()
            pp = self.get_prev_remaining_principal()
            pr = self.get_prev_interest_rate()
            td = self.get_this_date('4wc')
            tr = get_statutory_rate(tp, pp, pr, td, ts, ps)
        else:
            tr = ts
        self.set_this_interest_rate(tr)

    ####################################
    # DAYS

    def set_options(self, days):
        if(not re.match(r'^([\_\;\.\:\+\-\?\!tcwm]+)|([0-9]+)', days)):
            self._error('bad options "' + str(days) + '"')
        self.options = days

    def get_options(self):
        sc = self.should_insert_comma()
        cu = self.get_calculating_unit()
        si = self.should_include_first_day()
        ds = self.get_output_date_style()
        op = ''
        if(not sc):
            op = op + '_'
        if(ds == '3jc'):
            op = op + ';'
        elif(ds == '2wc'):
            op = op + '.'
        elif(ds == '4wc'):
            op = op + ':'
        if(si):
            op = op + '+'
        else:
            op = op + '-'
        if(cu == 'yearly+daily'):
            op = op + '?'
        if(cu == 'daily'):
            op = op + '!'
        return op

    def set_days(self, days):
        if(not re.match('^-?[0-9]+', days)):
            self._error('bad days "' + str(days) + '"')
        self._days = Decimal(days)

    def get_days(self):
        return str(self._days)

    def set_years(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad years "' + str(amount) + '"')
        self._years = Decimal(amount)

    def get_years(self):
        return str(self._years)

    def set_days_in_normal_year(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad days in normal year "' + str(amount) + '"')
        self._days_in_normal_year = Decimal(amount)

    def get_days_in_normal_year(self):
        return str(self._days_in_normal_year)

    def set_days_in_leap_year(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad days in leap year "' + str(amount) + '"')
        self._days_in_leap_year = Decimal(amount)

    def get_days_in_leap_year(self):
        return str(self._days_in_leap_year)

    def calc_and_set_years_and_days(self):
        pd = self.get_prev_date('4wc')
        td = self.get_this_date('4wc')
        ph = self.has_to_include_prev_day()
        th = self.has_to_include_this_day()
        dy, yr, nd, ld = count_years_and_days(pd, td, ph, th)
        self.set_days(dy)
        self.set_years(yr)
        self.set_days_in_normal_year(nd)
        self.set_days_in_leap_year(ld)

    ####################################
    # INTEREST

    def set_interest(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad interest "' + str(amount) + '"')
        self._interest = Decimal(amount)

    def get_interest(self):
        return str(self._interest)

    def calc_and_set_interest(self):
        pp = self.get_prev_remaining_principal()
        pr = self.get_prev_interest_rate()
        da = self.get_days()
        yr = self.get_years()
        dn = self.get_days_in_normal_year()
        dl = self.get_days_in_leap_year()
        cu = self.get_calculating_unit()
        it = calculate_interest(pp, pr, da, yr, dn, dl, cu)
        self.set_interest(it)

    ####################################
    # CHANGE OF PRINCIPLE

    def set_change_of_principal(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad change of principal "' + str(amount) + '"')
        self._change_of_principal = Decimal(amount)

    def get_change_of_principal(self):
        return str(self._change_of_principal)

    ####################################
    # REMAINING INTEREST

    def set_this_remaining_interest(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad this remaining interest "' + str(amount) + '"')
        self._this_remaining_interest = Decimal(amount)

    def get_this_remaining_interest(self):
        return str(self._this_remaining_interest)

    def set_prev_remaining_interest(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad prev remaining interest "' + str(amount) + '"')
        self._prev_remaining_interest = Decimal(amount)

    def get_prev_remaining_interest(self):
        return str(self._prev_remaining_interest)

    ####################################
    # REMAINING PRINCIPAL

    def set_this_remaining_principal(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad this remaining principal "' + str(amount) +
                        '"')
        self._this_remaining_principal = Decimal(amount)

    def get_this_remaining_principal(self):
        return str(self._this_remaining_principal)

    def set_prev_remaining_principal(self, amount):
        if(not re.match('^-?[0-9]+', amount)):
            self._error('bad prev remaining principal "' + str(amount) +
                        '"')
        self._prev_remaining_principal = Decimal(amount)

    def get_prev_remaining_principal(self):
        return str(self._prev_remaining_principal)

    ####################################
    # CALC AND SET CHANGE AND REMAINING

    def calc_and_set_change_and_remaining(self):
        pp = Decimal(self.get_prev_remaining_principal())
        rp = pp
        pi = Decimal(self.get_prev_remaining_interest())
        ti = Decimal(self.get_interest())
        ri = pi + ti
        ba = Decimal(self.get_this_borrowing_amount())
        ra = Decimal(self.get_this_repayment_amount())
        di = ba - ra
        if(rp > 0):
            # NARMAL CASE
            if((di >= 0) or (self.applies_repayment_to_principal())):
                rp += di  # apply to principal
            else:
                ri += di  # apply to interest
            if(rp < 0):
                ri += rp
                rp = Decimal(0)
            if(ri < 0):
                rp += ri
                ri = Decimal(0)
        elif(rp == 0):
            # FULL PAYMENT CASE
            if((ri * di) >= 0):
                rp += di  # apply to principal
            elif((ri * (ri + di)) <= 0):
                ri += di  # apply to interest
                rp = ri
                ri = Decimal(0)
            else:
                ri += di  # apply to interest
        else:
            # OVERPAYMENT CASE
            if((di <= 0) or (self.applies_borrowing_to_principal())):
                rp += di  # apply to principal
            else:
                ri += di  # apply to interest
            if(rp > 0):
                ri += rp
                rp = Decimal(0)
            if(ri > 0):
                rp += ri
                ri = Decimal(0)
        self.set_change_of_principal(str(rp - pp))
        self.set_this_remaining_interest(str(ri))
        self.set_this_remaining_principal(str(rp))

    ####################################
    # REMARKS

    def set_remarks(self, line):
        fo = re.search(re.compile('#.*$'), line)
        if(fo):
            self._remarks = fo.group(0)
        else:
            self._remarks = ''

    def get_remarks(self):
        return self._remarks

    ####################################
    # HAS TO INCLUDE DAY

    def set_has_to_include_this_day(self, has_to):
        self._has_to_include_this_day = has_to

    def has_to_include_this_day(self):
        return self._has_to_include_this_day

    def set_has_to_include_prev_day(self, has_to):
        self._has_to_include_prev_day = has_to

    def has_to_include_prev_day(self):
        return self._has_to_include_prev_day

    def calc_and_set_has_to_include_this_day(self):
        pp = Decimal(self.get_prev_remaining_principal())
        ba = Decimal(self.get_this_borrowing_amount())
        ra = Decimal(self.get_this_repayment_amount())
        di = ba - ra
        if(self.should_include_first_day()):
            th = (not self.judge_has_to_include_first_day(pp, di))
        else:
            th = True
        self.set_has_to_include_this_day(th)

    def calc_and_set_has_to_include_prev_day(self):
        pp = Decimal(self.get_prev_remaining_principal())
        ba = Decimal(self.get_prev_borrowing_amount())
        ra = Decimal(self.get_prev_repayment_amount())
        di = ba - ra
        if(self.should_include_first_day()):
            ph = self.judge_has_to_include_first_day(pp, di)
        else:
            ph = False
        self.set_has_to_include_prev_day(ph)

    def judge_has_to_include_first_day(self, prev_remaining, trading_amount):
        pr = prev_remaining
        ta = trading_amount
        # pr > 0
        #   ta > 0 -> True
        #   ta = 0 -> False
        #   ta < 0 -> False
        # pr = 0
        #   ta > 0 -> False
        #   ta = 0 -> False
        #   ta < 0 -> False
        # pr < 0
        #   ta > 0 -> False
        #   ta = 0 -> False
        #   ta < 0 -> True
        if((pr * ta) > 0):
            return True
        else:
            return False

    ##############################################
    # PRINT

    def get_trade(self, i):
        # SHOULD INSERT COMMA
        sc = self.should_insert_comma()
        # COMMENT OUT SYMBOL
        co = self.get_comment_out_symbol()
        # DATE
        ds = self.get_output_date_style()
        if(ds == ''):
            ds = self.get_input_date_style()
        td = self.get_this_date(ds)
        # BORROWING AMOUNT
        ba = self.get_this_borrowing_amount()
        if(sc):
            ba = locale.format_string('%d', int(ba), True)
        if(self.applies_borrowing_to_principal()):
            ba = '_' + ba
        # REPAYMENT AMOUNT
        ra = self.get_this_repayment_amount()
        if(sc):
            ra = locale.format_string('%d', int(ra), True)
        if(self.applies_repayment_to_principal()):
            ra = '_' + ra
        # INTEREST RATE
        ts = self.get_this_interest_rate_standard()
        tr = self.get_this_interest_rate()
        # tr = tr + '%'
        if((ts == '=') or (ts == '*')):
            tr = ts + tr
        # DAYS
        dy = self.get_days()
        if(i == 0):
            dy = self.get_options()
        # INTEREST
        it = self.get_interest()
        if(sc):
            it = locale.format_string('%d', int(it), True)
        if(i == 0 and it == '0'):
            it = '-'
        # CHANGE OF PRINCIPAL
        cp = self.get_change_of_principal()
        if(sc):
            cp = locale.format_string('%d', int(cp), True)
        # REMAINING INTEREST
        ti = self.get_this_remaining_interest()
        if(sc):
            ti = locale.format_string('%d', int(ti), True)
        # REMAINING PRINCIPAL
        tp = self.get_this_remaining_principal()
        if(sc):
            tp = locale.format_string('%d', int(tp), True)
        # REMARKS
        rm = self.get_remarks()
        # OUTPUT
        trade_line = ''
        of = self.get_output_style()
        if(of == 'math'):
            if(i > 0):
                trade_line += self.get_trade_math()
        elif(of == 'tex'):
            ba = ba.replace('_', '\\_')
            ra = ra.replace('_', '\\_')
            dy = dy.replace('_', '\\_')
            trade_line += \
                self.get_trade_tex(co, td, ba, ra, tr, dy, it, cp, ti, tp, rm)
        elif(of == 'csv'):
            trade_line += \
                self.get_trade_csv(co, td, ba, ra, tr, dy, it, cp, ti, tp, rm)
        elif(of == 'web'):
            trade_line += \
                self.get_trade_web(co, td, ba, ra, tr, dy, it, cp, ti, tp, rm)
        elif(of == 'markdown'):
            trade_line += \
                self.get_trade_mkd(co, td, ba, ra, tr, dy, it, cp, ti, tp, rm)
        else:
            trade_line += \
                self.get_trade_txt(co, td, ba, ra, tr, dy, it, cp, ti, tp, rm)
        # ERRER
        if(re.match('^-[0-9]+', dy)):
            trade_line += co + WARNING_MESSAGE_NUMBER_OF_DAYS_IS_NEGATIVE
        return trade_line

    def get_trade_math(self):
        sc = self.should_insert_comma()
        co = self.get_comment_out_symbol()
        ds = self.get_output_date_style()
        if(ds == ''):
            ds = self.get_input_date_style()
        td = self.get_this_date(ds)
        pd = self.get_prev_date(ds)
        pp = self.get_prev_remaining_principal()
        if(sc):
            pp = locale.format_string('%d', int(pp), True)
        pr = self.get_prev_interest_rate()
        dy = self.get_days()
        yr = self.get_years()
        dn = self.get_days_in_normal_year()
        dl = self.get_days_in_leap_year()
        dr = str(Decimal(dn) + Decimal(dl))
        cu = self.get_calculating_unit()
        if(cu == 'daily'):
            yd = ' ' * (4 - len(dy)) + dy + '/365'
        elif(cu == 'yearly+daily'):
            yd = '(' + \
                 ' ' * (2 - len(yr)) + yr + ' + ' + \
                 ' ' * (3 - len(dr)) + dr + '/365' + \
                 ')'
        else:
            yd = '(' + \
                 ' ' * (2 - len(yr)) + yr + ' + ' + \
                 ' ' * (3 - len(dn)) + dn + '/365 + ' + \
                 ' ' * (3 - len(dl)) + dl + '/366' + \
                 ')'
        it = self.get_interest()
        if(sc):
            it = locale.format_string('%d', int(it), True)
        return (co + pd + '-' + td + ': ' +
                ' ' * (10 - len(pp)) + pp + ' * ' +
                ' ' * (2 - len(pr)) + pr + '/100 * ' +
                yd + ' = ' +
                ' ' * (8 - len(it)) + it)

    def get_trade_tex(self, co, td, ba, ra, tr, dy, it, cp, ti, tp, rm):
        rm = rm.replace('\\', '{\\textbackslash}')
        rm = rm.replace('#', '\\#')
        rm = rm.replace('&', '\\&')
        return (td + '&' +
                ba + '&' +
                ra + '&' +
                tr + '&' +
                dy + '&' +
                it + '&' +
                cp + '&' +
                ti + '&' +
                tp + '&' +
                rm + '\\\\')

    def get_trade_csv(self, co, td, ba, ra, tr, dy, it, cp, ti, tp, rm):
        return ('"' + td + '",' +
                '"' + ba + '",' +
                '"' + ra + '",' +
                '"' + tr + '",' +
                '"' + dy + '",' +
                '"' + it + '",' +
                '"' + cp + '",' +
                '"' + ti + '",' +
                '"' + tp + '",' +
                '"' + rm + '"')

    def get_trade_web(self, co, td, ba, ra, tr, dy, it, cp, ti, tp, rm):
        return ('<tr>' +
                '<td>' + td + '</td>' +
                '<td align="right">' + ba + '</td>' +
                '<td align="right">' + ra + '</td>' +
                '<td align="right">' + tr + '</td>' +
                '<td align="right">' + dy + '</td>' +
                '<td align="right">' + it + '</td>' +
                '<td align="right">' + cp + '</td>' +
                '<td align="right">' + ti + '</td>' +
                '<td align="right">' + tp + '</td>' +
                '<td>' + rm + '</td>' +
                '</tr>')

    def get_trade_mkd(self, co, td, ba, ra, tr, dy, it, cp, ti, tp, rm):
        if(rm != ''):
            rm = ' ' + rm
        return (co +
                '|' + td +
                '|' + ' ' * (9 - len(ba)) + ba +
                '|' + ' ' * (8 - len(ra)) + ra +
                '|' + ' ' * (4 - len(tr)) + tr +
                '|' + ' ' * (4 - len(dy)) + dy +
                '|' + ' ' * (8 - len(it)) + it +
                '|' + ' ' * (9 - len(cp)) + cp +
                '|' + ' ' * (9 - len(ti)) + ti +
                '|' + ' ' * (10 - len(tp)) + tp +
                '|' + rm)

    def get_trade_txt(self, co, td, ba, ra, tr, dy, it, cp, ti, tp, rm):
        if(rm != ''):
            rm = ' ' + rm
        return (co +
                td +
                ' ' + ' ' * (9 - len(ba)) + ba +
                ' ' + ' ' * (8 - len(ra)) + ra +
                ' ' + ' ' * (4 - len(tr)) + tr +
                ' ' + ' ' * (4 - len(dy)) + dy +
                ' ' + ' ' * (8 - len(it)) + it +
                ' ' + ' ' * (9 - len(cp)) + cp +
                ' ' + ' ' * (9 - len(ti)) + ti +
                ' ' + ' ' * (10 - len(tp)) + tp +
                rm)

    ##############################################
    # MAIN

    ####################################
    # CONSTRUCTOR

    def __init__(self, line):
        self.line = line
        # SET HEAD SYMBOL
        self.set_head_symbol(line)
        line = re.sub('^' + self.get_head_symbol(), '', line)
        # SET REMARKS
        self.set_remarks(line)
        line = re.sub(self.get_remarks() + '$', '', line)
        # SPLIT TO WORD
        if re.match('^\\|', line):
            words = line.replace(' ', '').split('|')
            words.pop(0)
        else:
            words = line.split()
        # SET INPUT DATE STYLE AND THIS DATE
        if(len(words) > 0):
            self.set_input_date_style(words[0])
            self.set_this_date(words[0])
        else:
            self._error('no date "' + str(self.line) + '"')
        # SET APPLIES BORROWING TO CAPITAL AND THIS BORROWING AMOUNT
        self.set_applies_borrowing_to_principal('-')  # False
        self.set_this_borrowing_amount('0')
        if(len(words) > 1):
            self.set_applies_borrowing_to_principal(words[1])
            self.set_this_borrowing_amount(words[1])
        # SET APPLIES REPAYMENT TO CAPITAL AND THIS REPAYMENT AMOUNT
        self.set_applies_repayment_to_principal('-')  # False
        self.set_this_repayment_amount('0')
        if(len(words) > 2):
            self.set_applies_repayment_to_principal(words[2])
            self.set_this_repayment_amount(words[2])
        # SET THIS INTEREST RATE STANDARD
        self.set_this_interest_rate_standard('')
        if(len(words) > 3):
            self.set_this_interest_rate_standard(words[3])
        # SET OPTIONS
        self.set_options('+')  # should include first day
        if(len(words) > 4):
            self.set_options(words[4])
        # SET INTEREST
        self.set_interest('0')

    ####################################
    # RESET OPTIONS

    def reset_options(self):
        self.set_comment_out_symbol(self.get_head_symbol())
        if('_' in self.options):
            self.set_should_insert_comma(False)
        if('?' in self.options):
            self.set_calculating_unit('yearly+daily')
        if('!' in self.options):
            self.set_calculating_unit('daily')
        if('+' in self.options):
            self.set_should_include_first_day(True)
        if('-' in self.options):
            self.set_should_include_first_day(False)
        if(';' in self.options):
            self.set_output_date_style('3jc')
        if('.' in self.options):
            self.set_output_date_style('2wc')
        if(':' in self.options):
            self.set_output_date_style('4wc')
        if('t' in self.options):
            self.set_output_style('tex')
        if('c' in self.options):
            self.set_output_style('csv')
        if('w' in self.options):
            self.set_output_style('web')
        if('m' in self.options):
            self.set_output_style('math')

    ####################################
    # INHERIT PREV DATA

    def inherit_prev_data_for_first_trade(self):
        # DATE
        self.set_prev_date(self.get_this_date())
        # BORROWING AMOUNT
        self.set_prev_borrowing_amount('-')
        # REPAYMENT AMOUNT
        self.set_prev_repayment_amount('-')
        # INTEREST RATE STANDARD
        self.set_prev_interest_rate_standard('=')
        # INTEREST RATE
        self.set_prev_interest_rate('0')
        # REMAINING INTEREST
        self.set_prev_remaining_interest('0')
        # REMAINING PRINCIPAL
        self.set_prev_remaining_principal('0')
        # HAS TO INCLUDE PREV DAY
        self.set_has_to_include_prev_day(True)  # False causes an error

    def inherit_prev_data_for_second_and_subsequent_trade(self, prev):
        # DATE
        pd = prev.get_this_date()
        self.set_prev_date(pd)
        # BORROWING AMOUNT
        ba = prev.get_this_borrowing_amount()
        self.set_prev_borrowing_amount(ba)
        # REPAYMENT AMOUNT
        ra = prev.get_this_repayment_amount()
        self.set_prev_repayment_amount(ra)
        # INTEREST RATE STANDARD
        ps = prev.get_this_interest_rate_standard()
        self.set_prev_interest_rate_standard(ps)
        # INTEREST RATE
        pr = prev.get_this_interest_rate()
        self.set_prev_interest_rate(pr)
        # REMAINING INTEREST
        pi = prev.get_this_remaining_interest()
        self.set_prev_remaining_interest(pi)
        # REMAINING PRINCIPAL
        pp = prev.get_this_remaining_principal()
        self.set_prev_remaining_principal(pp)
        # HAS TO INCLUDE PREV DAY
        ph = not prev.has_to_include_this_day()
        self.set_has_to_include_prev_day((ph))

    ####################################
    # CHECK CONSISTENCY

    def check_consistency(self):
        ds = self.get_output_date_style()
        if(ds == ''):
            ds = self.get_input_date_style()
        td = self.get_this_date(ds)
        ba = Decimal(self.get_this_borrowing_amount())
        ra = Decimal(self.get_this_repayment_amount())
        it = Decimal(self.get_interest())
        pi = Decimal(self.get_prev_remaining_interest())
        ti = Decimal(self.get_this_remaining_interest())
        pp = Decimal(self.get_prev_remaining_principal())
        tp = Decimal(self.get_this_remaining_principal())
        if(((ba - ra + it) - (tp - pp) - (ti - pi)) != 0):
            self._error('inconsistent "' + td + '"')

############################################################
# OPTION PROCESS


debug_mode = False

if __name__ == '__main__':
    options = ['help', 'version',
               'daily', 'Daily',
               'exclude',
               '3jc', '2wc', '4wc',
               'no-comma'
               'tex', 'csv', 'web', 'markdown', 'math',
               'sample', 'debug']
    try:
        opts, args = getopt.getopt(sys.argv[1:], 'hvdDe324ntcwmM', options)
    except getopt.GetoptError:
        sys.exit(1)
    for opt, arg in opts:
        if opt in ('-h', '--help'):
            print(HELP_MESSAGE)
            sys.exit(0)
        elif opt in ('-v', '--version'):
            print('keiji ' + __version__)
            sys.exit(0)
        elif opt in ('-d', '--daily'):
            Trade.set_calculating_unit('yearly+daily')
        elif opt in ('-D', '--Daily'):
            Trade.set_calculating_unit('daily')
        elif opt in ('-e', '--exclude'):
            Trade.set_should_include_first_day(False)
        elif opt in ('-3', '--3jc'):
            Trade.set_output_date_style('3jc')
        elif opt in ('-2', '--2wc'):
            Trade.set_output_date_style('2wc')
        elif opt in ('-4', '--4wc'):
            Trade.set_output_date_style('4wc')
        elif opt in ('-n', '--no-comma'):
            Trade.set_should_insert_comma(False)
        elif opt in ('-t', '--tex'):
            Trade.set_output_style('tex')
        elif opt in ('-c', '--csv'):
            Trade.set_output_style('csv')
        elif opt in ('-w', '--web'):
            Trade.set_output_style('web')
        elif opt in ('-m', '--markdown'):
            Trade.set_output_style('markdown')
        elif opt in ('-M', '--math'):
            Trade.set_output_style('math')
        elif opt in ('--sample'):
            print(SAMPLE_DATA)
            sys.exit(0)
        elif opt in ('--debug'):
            debug_mode = True


############################################################
# MAIN


def open_input_source(args):
    if(len(args) == 0):
        return sys.stdin
    elif(len(args) == 1):
        if(args[0] == '-'):
            return sys.stdin
        elif(os.path.isfile(args[0])):
            return open(args[0], 'r')
        else:
            sys.stderr.write('no such file "' + args[0] + '"\n')
            sys.exit(1)
    else:
        sys.stderr.write('too many arguments "' + args[0] + ' ..."\n')
        sys.exit(1)


def is_data_line(line):
    res = '^[#%\\|]?\\s*' \
        + '((?:[A-Z]?[0-9]{2}|[0-9]{4})[\\.\\-][0-9]{2}[\\.\\-][0-9]{2})' \
        + '(?:\\s*[\\| ]\\s*_?(?:[,0-9]|-)\\s*)?' \
        + '(?:\\s*[\\| ]\\s*_?(?:[,0-9]|-)\\s*)?' \
        + '.*$'
    if not re.match(res, line):
        return False
    sd = re.sub(res, '\\1', line)
    if to_date(sd) is None:
        return False
    return True


def main(raw_data):
    trades = []
    bad_lines = []
    lines = raw_data.split('\n')
    for i, line in enumerate(lines):
        # MESSAGE FOR DEBUGGING
        if debug_mode:
            sys.stderr.write('reading line ' + str(i) + '\n')
        # ACCEPT DATA
        line = line.rstrip()
        if not is_data_line(line):
            if line != '' and \
               ('日付' not in line) and ('合計' not in line) and \
               ('---:' not in line) and not re.match('^=+$', line):
                bad_lines.append(line)
            continue
        trades.append(Trade(line))
    # CALC AMOUNT
    for i, this in enumerate(trades):
        # MESSAGE FOR DEBUGGING
        if(debug_mode):
            sys.stderr.write('calculating line ' + str(i) + '\n')
        # INHERIT DATA
        if(i == 0):
            this.reset_options()
            this.inherit_prev_data_for_first_trade()
        else:
            prev = trades[i - 1]
            this.inherit_prev_data_for_second_and_subsequent_trade(prev)
        # THIS INTEREST RATE STANDARD
        this.check_and_set_this_interest_rate_standard()
        # HAS TO INCLUDE PREV DAY AND THIS DAY
        this.calc_and_set_has_to_include_prev_day()
        this.calc_and_set_has_to_include_this_day()
        if i == (len(trades) - 1):
            this.set_has_to_include_this_day(True)  # include last day
        # YEARS AND DAYS
        this.calc_and_set_years_and_days()
        # INTEREST
        this.calc_and_set_interest()
        # CHANGE AND REMAINING
        this.calc_and_set_change_and_remaining()
        # INTEREST RATE
        this.calc_and_set_this_interest_rate()
        # TOTAL AMOUNT
        ti = Decimal(this.get_this_remaining_interest())
        tp = Decimal(this.get_this_remaining_principal())
        Trade.set_total_amount(str(tp + ti))
        # CHECK CONSISTENCY
        # this.check_consistency()
    # MAKE OUTPUT
    output = ''
    output += Trade.get_header() + '\n'
    for i, tr in enumerate(trades):
        # MESSAGE FOR DEBUGGING
        if debug_mode:
            sys.stderr.write('printing line ' + str(i) + '\n')
        output += tr.get_trade(i) + '\n'
    output += Trade.get_footer() + '\n'
    if len(bad_lines) > 0:
        output += '次の行は除外しました。\n'
        for line in bad_lines:
            output += line + '\n'
    return output


if __name__ == '__main__':
    # IMPORT DATA
    input = open_input_source(args)
    raw_data = input.read()
    output = main(raw_data)
    # PRINT
    print(output, end='')
