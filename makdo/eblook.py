#!/usr/bin/python3
# Name:         eblook.py
# Version:      v01
# Time-stamp:   <2025.01.07-07:23:04-JST>

# eblook.py
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


import os
import argparse
import subprocess
import re


__version__ = 'v01'


if os.path.exists('/usr/bin/eblook'):
    EBLOOK = '/usr/bin/eblook'
if os.path.exists('/usr/local/bin/eblook'):
    EBLOOK = '/usr/local/bin/eblook'

GAIJI_KOJIEN = {
    'za422': '【文】',
    'za423': '１',
    'za425': '２',
    # 'za426': '',
    'za427': '【一】', 'za428': '【二】',
    # 'za429': '',
    'za42a': '３',
    # 'za42b': '',
    # 'za42c': '',
    'za42d': '（〱の横書き）】', 'za42e': '【漢字',  # 2つで1つ
    # 'za42f': '',
    # 'za430': '',
    # 'za431': '',
    'za432': '４',
    # 'za433': '',
    # 'za434': '',
    # 'za435': '',
    'za436': '５',
    # 'za437': '',
    'za438': '𣑥',
    # 'za439': '',
    # 'za43a': '',
    # 'za43b': '',
    'za43c': '６',
    # 'za43d': '',
    # 'za43e': '',
    # 'za43f': '',
    'za440': '７',
    # 'za441': '',
    # 'za442': '',
    # 'za443': '',
    'za444': '【三】',
    'za648': '８',
    # 'za4': '９',
    'za744': '𩊱',
    'zac49': '𫒒',  # 金丘
    'zac6e': '𑖀',  # 梵字（阿の音訳となった字）
    'zaf57': '㣺',
    'zb476': '⻞',
    'zb76e': '晌',
    'zb773': '﨑',
    'zb847': '（季）', 'zb848': '（呉）', 'zb849': '（漢）', 'zb84a': '（唐）',
    'zb84b': '（慣）', 'zb84c': '（同）', 'zb84d': '（対）',
    'zb84f': '【A】',
    'zb850': '【漢】',
    'zb851': '【意', 'zb852': '味】',    # 【意味】
    'zb853': '【解', 'zb854': '字】',    # 【解字】
    'zb855': '【下', 'zb856': 'つき】',  # 【下つき】
    'zb857': '【難', 'zb858': '読】',    # 【難読】
    'zb929': '弎',
    'zb931': '卋',
    'zb93f': '⽏',
    'zb956': '𠆢',  # ひとやね
    'zb95a': '𫝆', 'zb95b': '㠯', 'zb95c': '仡',
    'zb97c': '每',
    'zba2b': '𠊳',  # ⺅㪅
    'zba4e': '【漢字（傷のつくり）】', 'zba4f': '【漢字（修の彡が羽）】',
    'zba58': '僨', 'zba59': '菐',
    'zba62': '【漢字（修の⼺が革）】',
    'zba72': '【漢字（兼のソが八）】', 'zba73': '【漢字（六の点なし）】',
    'zbb24': '【漢字（最の異字体（本字））】',
    'zbb2d': '冣',
    'zbb42': '逬',
    'zbb4e': '㝴',
    'zbb61': '【漢字（肖の⺌が小）】',
    'zbc42': '【漢字（卑の点なし）】', 'zbc43': '劦',
    'zbc5f': '【漢字（難解文字）】', 'zbc60': '【漢字（原の小が水）】',
    'zbc61': '厝',
    'zbc72': '⺈', 'zbc73': '𧰼',
    'zbc7c': '𠮷',
    'zbd22': '吳', 'zbd23': '吴', 'zbd24': '吿', 'zbd25': '呏',
    'zbd3d': '哿',
    'zbd5e': '喿',
    'zbe2a': '囟',
    'zbe37': '阫',
    'zbf21': '【漢字（夢の夕が目）】',
    'zbf22': '㝱', 'zbf23': '夣', 'zbf24': '夤',
    'zbf28': '𡗗',
    'zbf30': '妒',
    'zbf3b': '娀',
    'zbf79': '寘', 'zbf7a': '【漢字（帚の巾が又）】',
    'zc024': '【漢字（対の点なし）】',
    'zc23d': '【漢字（難解文字）】',
    'zc02b': '⺌',
    'zc02d': '尙',
    'zc04a': '㟢',
    'zc137': '【漢字（珍獣の名）】',
    'zc145': '辡',
    'zc17b': '狀', 'zc17c': '【漢字（爿羊）】',
    'zc224': '𥝕',  # 禾亡
    'zc23c': '羑',
    'zc37b': '䍃', 'zc37c': '【漢字（違の2点しんにょう）】',
    'zc46b': '阽',
    'zc46e': '陔',
    'zc524': '𨺬',  # 阝界
    'zc530': '⺍',
    'zc534': '【漢字（月の異字体）】',
    'zc72c': '𢫦',  # 扌百
    'zc771': '撾',
    'zc828': '數', 'zc829': '鼔',
    'zc835': '斋',
    'zc83b': '【漢字（旁の異字体）】',
    'zc845': '时',
    'zc84b': '【漢字（時の異字体）】',
    'zc874': '【漢字（月の異字体）】',
    'zc965': '枒',
    'zcb53': '【漢字（気の異字体）】',
    'zcb6c': '【漢字（法の異字体）】', 'zcb6d': '㳒',
    'zcb73': '海',
    'zcc29': '【漢字（消の⺌が小）】',
    'zcd4d': '煁',
    'zcf2a': '⺪',
    'zcf68': '眜', 'zcf69': '眢',
    'zd031': '𨥨',  # 金矛
    'zd27b': '⺷',
    'zd358': '𦧟',  # 舌沓
    'zd465': '衕',
    'zd57c': '諭',  # 言兪
    'zd65d': '【漢字（起の己が巳）】',
    'zd759': '辤',
    'zd873': '雱',
    'zd92d': '韒',
    'zd94e': '頖',
    'zd970': '【漢字（食の異字体）】', 'zd971': '【漢字（食の異字体）】',
    'zda38': '馱',
    'zda62': '髙',
    'zdb33': '鯋',
    'zdb35': '鯈',
    'zdc35': '壴',
}

GAIJI_KANJIGEN = {
    'ha121': 'ā', 'ha122': 'á', 'ha123': 'ǎ', 'ha124': 'à',
    'ha125': 'ē', 'ha126': 'é', 'ha127': 'ě', 'ha128': 'è',
    'ha129': 'ī', 'ha12a': 'í', 'ha12b': 'ǐ', 'ha12c': 'ì',
    'ha12d': 'ō', 'ha12e': 'ó', 'ha12f': 'ǒ', 'ha130': 'ò',
    'ha131': 'ū', 'ha132': 'ú', 'ha133': 'ǔ', 'ha134': 'ù',
    'ha137': 'ǚ',
    # 'ha138': '',
    'ha139': 'ü',
    'ha13a': 'ヮ', 'ha13b': 'ヰ', 'ha13c': 'ヱ',
    'za13a': 'ヮ', 'za13b': 'ヰ', 'za13c': 'ヱ',
    'za121': 'ā',
    'za141': '忄',
    'za143': '⺘',
    'za149': '⽧',
    'za158': '⻌',
    'za160': '❶', 'za161': '❷', 'za162': '❸', 'za163': '❹', 'za164': '❺',
    'za165': '❻', 'za166': '❼', 'za167': '❽', 'za168': '❾', 'za169': '❿',
    'za16a': '⓫', 'za16b': '⓬', 'za16c': '⓭', 'za16d': '⓮',
    'za16f': '【一】', 'za170': '【二】', 'za171': '【三】',
    'za173': '①', 'za174': '②', 'za175': '③', 'za176': '④', 'za177': '⑤',
    'za178': '⑥', 'za179': '⑦', 'za17a': '⑧', 'za17b': '⑨', 'za17c': '⑩',
    'za229': '⏋',
    'za22d': '—',
    'za233': '【呉】', 'za234': '【漢】', 'za235': '【唐】', 'za236': '【慣】',
}

GAIJI_CHUJITEN = {
    'ha121': '・', 'ha122': '：', 'ha123': '︙',
    'ha172': '（？）',
    'ha174': '（？）',
    'ha26b': '･',
    'za321': '【名】', 'za322': '【代】', 'za323': '【形】', 'za324': '【動】',
    'za325': '【副】', 'za326': '【接】', 'za327': '【前】', 'za328': '【冠】',
    'za329': '【間】',
    'za32a': '【助', 'za32b': '動】',  # 【助動】
    'za32c': '【接', 'za32d': '頭】', 'za32e': '尾】',  # 【接頭】/【接尾】
    'za32f': '【Ｕ】', 'za330': '【Ｃ】',
    'za332': '（復）',
    'za333': '【Ａ】', 'za334': '【Ｐ】', 'za335': '（自）', 'za336': '（他）',
    'za337': '【成', 'za338': '句】',  # 【成句】
    'za339': '♪', 'za33a': '✓',
    'za33c': '≡',
    'za33f': '→',
    'za34f': '⇔',
    'za37e': '♮',
}

GAIJI_GENIUS = {
    'ha12d': '(ə)',
    # 'ha174': '',
    'ha270': 'ɪ̀',
    'zb430': '【Ｃ】', 'zb431': '【Ｕ】',
    'zb478': '→',
    'zb434': '↝',
    'zb43b': '︙',
    'zb44e': '♮',
}

GAIJI_BIZTEC = {
    'ha13c': 'é',
    'za143': 'II',
    'za148': '▶',
    'za149': '((', 'za14a': '))',
    'za14d': '⇔',
    'za153': '-',
    'za154': '【Ｕ】',
}

GAIJI_NANMED = {
    'hb124': 'Ö',
    'hb127': 'ê',
    'hb129': 'ä',
    'hb12b': 'é', 'hb12c': 'ê', 'hb12d': 'è', 'hb12e': 'ë',
    'hb136': 'ö',
    'hb138': 'ü',
    'za122': '①', 'za123': '②', 'za124': '③', 'za125': '④', 'za126': '⑤',
    'za127': '⑥', 'za128': '⑦', 'za129': '⑧',
}


class Dictionary:

    def __init__(self) -> None:
        self.number: int = -1
        self.a_name: str = ''
        self.k_name: str = ''
        self.gaiji: dict = {}


class Item:

    def __init__(self) -> None:
        self.dictionary = None
        self.number = -1
        self.code = ''
        self.title = ''
        self.content = ''

    def make_up_title(self, title: str) -> str:
        gaiji: dict = self.dictionary.gaiji
        for g in gaiji:
            while re.match('^(.|\n)*<gaiji=' + g + '>(.|\n)*$', title, re.I):
                title = re.sub('<gaiji=' + g + '>', gaiji[g], title, re.I)
        return title

    def get_content(self, dictionary_directory: str) -> str:
        gaiji = self.dictionary.gaiji
        number = self.dictionary.number
        command = 'echo "' \
            + 'select ' + str(number) + '\n' \
            + 'content ' + self.code \
            + '" | ' \
            + EBLOOK + ' ' + dictionary_directory
        try:
            sr = subprocess.run(command,
                                check=True,
                                shell=True,
                                stdout=subprocess.PIPE,
                                encoding="utf-8")
        except subprocess.CalledProcessError:
            return None
        so = sr.stdout
        so = re.sub('^eblook> eblook> ', '', so)
        so = re.sub('\neblook> ', '', so)
        for g in gaiji:
            while re.match('^(.|\n)*<gaiji=' + g + '>(.|\n)*$', so, re.I):
                so = re.sub('<gaiji=' + g + '>', gaiji[g], so, re.I)
        so = re.sub('<prev>(.*?)</prev>', '前：\\1', so)
        so = re.sub('<next>(.*?)</next>', '次：\\1', so)
        so = re.sub('<reference>(.*?)</reference=([0-9]+:[0-9]+)>',
                    '\\1<' + str(number) + ':\\2>',
                    so)
        return so

    def print_item(self) -> None:
        print('## 【' + self.dictionary.k_name
              + '\u3000' + self.title + '】')
        print(self.content + '\n')


class Eblook:

    def __init__(self) -> None:
        self.dictionary_directory: str = ''
        self.dictionaries: list[Dictionary] = []
        self.search_word: str = ''
        self.items: list[Item] = []

    def set_dictionaries(self, dictionary_directory: str) -> bool:
        self.dictionary_directory = dictionary_directory
        command = 'echo "' \
            + 'list' \
            + '" | ' \
            + EBLOOK + ' ' + self.dictionary_directory
        try:
            sr = subprocess.run(command,
                                check=True,
                                shell=True,
                                stdout=subprocess.PIPE,
                                encoding='utf-8')
        except subprocess.CalledProcessError:
            return False
        so = sr.stdout
        so = re.sub('^eblook> ', '', so)
        so = re.sub('\neblook> $', '', so)
        dictionaries: list[Dictionary] = []
        for sos in so.split('\n'):
            res = '\\s*([0-9]+)\\.\\s+(\\S+)\\s+(.*)$'
            if re.match(res, sos):
                dic = Dictionary()
                dic.number = int(re.sub(res, '\\1', sos))
                dic.a_name = re.sub(res, '\\2', sos)
                dic.k_name = re.sub(res, '\\3', sos)
                dic.gaiji = self._get_gaiji(dic.a_name)
                if re.match('^kojien', dic.a_name):
                    for g in GAIJI_KOJIEN:
                        dic.gaiji[g] = GAIJI_KOJIEN[g]
                if dic.a_name == 'kanjigen':
                    for g in GAIJI_KANJIGEN:
                        dic.gaiji[g] = GAIJI_KANJIGEN[g]
                if dic.a_name == 'chujiten':
                    for g in GAIJI_CHUJITEN:
                        dic.gaiji[g] = GAIJI_CHUJITEN[g]
                if dic.a_name == 'genius4':
                    for g in GAIJI_GENIUS:
                        dic.gaiji[g] = GAIJI_GENIUS[g]
                if dic.a_name == 'biztec4a':
                    for g in GAIJI_BIZTEC:
                        dic.gaiji[g] = GAIJI_BIZTEC[g]
                if dic.a_name == 'nanmed18':
                    for g in GAIJI_NANMED:
                        dic.gaiji[g] = GAIJI_NANMED[g]
                dictionaries.append(dic)
        self.dictionaries = dictionaries
        return True

    def _get_gaiji(self, a_name) -> dict:
        gaiji_directory = self.dictionary_directory + '/GAIJI_XML'
        if not os.path.exists(gaiji_directory):
            return {}
        if not os.path.isdir(gaiji_directory):
            return {}
        for dne in os.listdir(gaiji_directory):
            if not re.match('^.*\\.plist$', dne):
                continue
            x = a_name.lower()
            y = dne.lower()
            y = re.sub('\\.plist$', '', y)
            if x.upper() == y.upper():
                with open(gaiji_directory + '/' + dne, 'r') as f:
                    gaiji = {}
                    for line in f.readlines():
                        line = line.rstrip()
                        res = '^.*<key>(.+)</key><string>(.+)</string>.*$'
                        if re.match(res, line):
                            k = re.sub(res, '\\1', line).lower()
                            s = re.sub(res, '\\2', line)
                            gaiji[k] = s
                    return gaiji
        for dne in os.listdir(gaiji_directory):
            if not re.match('^.*\\.plist$', dne):
                continue
            x = a_name.lower()
            x = re.sub('[0-9]+', '', x)
            y = dne.lower()
            y = re.sub('\\.plist$', '', y)
            y = re.sub('[0-9]+', '', y)
            if x.upper() == y.upper():  #
                with open(gaiji_directory + '/' + dne, 'r') as f:
                    gaiji = {}
                    for line in f.readlines():
                        line = line.rstrip()
                        res = '^.*<key>(.+)</key><string>(.+)</string>.*$'
                        if re.match(res, line):
                            k = re.sub(res, '\\1', line).lower()
                            s = re.sub(res, '\\2', line)
                            gaiji[k] = s
                    return gaiji
        return {}

    def set_search_word(self, search_word: str) -> bool:
        self.search_word = search_word
        items: [Item] = []
        if re.match('^([0-9]+):([0-9]+:[0-9]+)$', search_word):
            dc = re.sub('^([0-9]+):([0-9]+:[0-9]+)$', '\\1', search_word)
            cc = re.sub('^([0-9]+):([0-9]+:[0-9]+)$', '\\2', search_word)
            for d in self.dictionaries:
                if d.number == int(dc):
                    i = Item()
                    i.dictionary = d
                    i.code = cc
                    i.content = i.get_content(self.dictionary_directory)
                    items.append(i)
        else:
            for d in self.dictionaries:
                command = 'echo "' \
                    + 'select ' + str(d.number) + '\n' \
                    + 'search ' + search_word \
                    + '" | ' \
                    + EBLOOK + ' ' + self.dictionary_directory
                try:
                    sr = subprocess.run(command,
                                        check=True,
                                        shell=True,
                                        stdout=subprocess.PIPE,
                                        encoding="utf-8")
                except subprocess.CalledProcessError:
                    return False
                so = sr.stdout
                so = re.sub('^eblook> eblook> ', '', so)
                so = re.sub('\neblook> ', '', so)
                for sos in so.split('\n'):
                    res = '\\s*([0-9]+)\\.\\s+(\\S+)\\s+(.*)$'
                    if re.match(res, sos):
                        i = Item()
                        i.dictionary = d
                        i.number = int(re.sub(res, '\\1', sos))
                        i.code = re.sub(res, '\\2', sos)
                        i.title = re.sub(res, '\\3', sos)
                        i.title = i.make_up_title(i.title)
                        i.content = i.get_content(self.dictionary_directory)
                        items.append(i)
        self.items = items
        return True


if __name__ == '__main__':

    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description='eblookのランチャーです',
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
        'dictionary_directory',
        help='辞書ディレクトリー')
    parser.add_argument(
        'search_word',
        help='検索する言葉')
    args = parser.parse_args()

    eb = Eblook()
    eb.set_dictionaries(args.dictionary_directory)
    eb.set_search_word(args.search_word)

    for ei in eb.items:
        ei.print_item()
