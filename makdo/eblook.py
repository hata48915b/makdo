#!/usr/bin/python3


import os
import argparse
import subprocess
import re


__version__ = 'v01'


EBLOOK = '/usr/bin/eblook'

GAIJI_KOJIEN = {
    'za423': '１', 'za425': '２',
    'za427': '【一】', 'za428': '【二】', 'za444': '【三】',
    'za42a': '３', 'za432': '４',
    # 'za4': '５',
    'za43c': '６', 'za440': '７', 'za648': '８',
    # 'za4': '９',
    'zb848': '（呉）', 'zb849': '（漢）',
    'zb84b': '（慣）', 'zb84c': '（同）', 'zb84d': '（対）',
    'zb850': '【漢】',
    'zb851': '【意', 'zb852': '味】',    # 【意味】
    'zb853': '【解', 'zb854': '字】',    # 【解字】
    'zb855': '【下', 'zb856': 'つき】',  # 【下つき】
    'zb857': '【難', 'zb858': '読】',    # 【難読】
    'zb956': '𠆢',  # ひとやね
    'zb97c': '每',
    'zba58': '僨',
    'zba59': '菐',
    'zbc7c': '𠮷',
    'zb773': '﨑',
    'zbf3b': '娀',
    'zcb73': '海',
    'zc02b': '⺌',
    'zc04a': '㟢',
    'zc145': '辡',
    'zc828': '數',
    'zc829': '鼔',
    'zc845': '时',
    'zc965': '枒',
    'zda62': '髙',
    'zdc35': '壴',
}

GAIJI_KANJIGEN = {
    'ha121': 'ā',
    'ha122': 'á',
    'ha123': 'ǎ',
    'ha124': 'à',
    'ha125': 'ē',
    'ha126': 'é',
    'ha127': 'ě',
    'ha128': 'è',
    'ha129': 'ī',
    'ha12a': 'í',
    # 'ha12b': '',
    # 'ha12c': '',
    'ha12d': 'ō',
    # 'ha12e': '',
    'ha12f': 'ǒ',
    'ha130': 'ò',
    # 'ha131': '',
    'ha132': 'ú',
    'ha133': 'ǔ',
    'ha13a': 'ヮ',
    'za13a': 'ヮ',
    'za13c': 'ヱ',
    'za160': '❶', 'za161': '❷', 'za162': '❸', 'za163': '❹', 'za164': '❺',
    'za165': '❻', 'za166': '❼', 'za167': '❽', 'za168': '❾', 'za169': '❿',
    'za173': '①', 'za174': '②', 'za175': '③', 'za176': '④', 'za177': '⑤',
    'za178': '⑥', 'za179': '⑦', 'za17a': '⑧', 'za17b': '⑨', 'za17c': '⑩',
    'za233': '【呉】', 'za234': '【漢】',
    'za236': '【慣】',
    'za22d': '—'
}

GAIJI_CHUJITEN = {
    'ha121': '・',
    'ha122': '：',
    'ha123': '︙',
    'ha172': '（？）',
    'ha174': '（？）',
    'ha26b': '･',
    'za321': '【名】',
    'za323': '【形】',
    'za324': '【動】',
    'za325': '【副】',
    'za329': '【間】',
    'za32a': '【助', 'za32b': '動】',  # 【助動】
    'za32c': '【接', 'za32d': '頭】',  # 【接頭】
    'za32f': '【Ｕ】', 'za330': '【Ｃ】',
    'za332': '（復）',
    'za333': '【Ａ】', 'za334': '【Ｐ】', 'za335': '（自）', 'za336': '（他）',
    'za337': '【成', 'za338': '句】',  # 【成句】
    'za339': '♪', 'za33a': '✓',
    'za33c': '≡',
    'za33f': '→',
    'za34f': '⇔',
}

GAIJI_GENIUS = {
    'ha12d': '(ə)',
    'ha270': 'ɪ̀',
    'zb430': '【Ｃ】',
    'zb431': '【Ｕ】',
    'zb478': '→',
    'zb434': '↝',
    'zb43b': '︙',
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


class Eblook:

    def __init__(self):
        self.dictionary_directory = ''
        self.dictionaries = []
        self.search_word = ''
        self.items = []

    def set_dictionary_directory(self, dictionary_directory):
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
                                encoding="utf-8")
        except subprocess.CalledProcessError:
            sys.exit(1)
        so = sr.stdout
        so = re.sub('^eblook> ', '', so)
        so = re.sub('\neblook> ', '', so)
        dictionaries = []
        for sos in so.split('\n'):
            res = '\\s*([0-9]+)\\.\\s+(\\S+)\\s+(.*)$'
            if re.match(res, sos):
                d = Dictionary()
                d.number = int(re.sub(res, '\\1', sos))
                d.a_name = re.sub(res, '\\2', sos)
                d.k_name = re.sub(res, '\\3', sos)
                d.gaiji = self._get_gaiji(d.a_name)
                if d.a_name == 'kojien5':
                    for g in GAIJI_KOJIEN:
                        d.gaiji[g] = GAIJI_KOJIEN[g]
                if d.a_name == 'kanjigen':
                    for g in GAIJI_KANJIGEN:
                        d.gaiji[g] = GAIJI_KANJIGEN[g]
                if d.a_name == 'chujiten':
                    for g in GAIJI_CHUJITEN:
                        d.gaiji[g] = GAIJI_CHUJITEN[g]
                if d.a_name == 'genius4':
                    for g in GAIJI_GENIUS:
                        d.gaiji[g] = GAIJI_GENIUS[g]
                if d.a_name == 'biztec4a':
                    for g in GAIJI_BIZTEC:
                        d.gaiji[g] = GAIJI_BIZTEC[g]

                dictionaries.append(d)
        self.dictionaries = dictionaries

    def _get_gaiji(self, a_name):
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
        return {}

    def set_search_word(self, search_word):
        self.search_word = search_word
        items = []
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
                    sys.exit(1)
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
                        i.title = i.get_title(i.title)
                        i.content = i.get_content(self.dictionary_directory)
                        items.append(i)
        self.items = items


class Dictionary:

    def __init__(self):
        self.number = -1
        self.a_name = ''
        self.k_name = ''
        self.gaiji = {}


class Item:

    def __init__(self):
        self.dictionary = None
        self.number = -1
        self.code = ''
        self.title = ''
        self.content = ''

    def get_title(self, title):
        gaiji = self.dictionary.gaiji
        for g in gaiji:
            while re.match('^(.|\n)*<gaiji=' + g + '>(.|\n)*$', title):
                title = re.sub('<gaiji=' + g + '>', gaiji[g], title, re.I)
        return title

    def get_content(self, dictionary_directory):
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
            while re.match('^(.|\n)*<gaiji=' + g + '>(.|\n)*$', so):
                so = re.sub('<gaiji=' + g + '>', gaiji[g], so, re.I)
        so = re.sub('<prev>(.*?)</prev>', '前：\\1', so)
        so = re.sub('<next>(.*?)</next>', '次：\\1', so)
        so = re.sub('<reference>(.*?)</reference=([0-9]+:[0-9]+)>',
                    '\\1<' + str(number) + ':\\2>',
                    so)
        return so

    def print_item(self):
        print('=====================================' +
              '=====================================')
        print('●\u3000' + self.dictionary.k_name + '\u3000' + self.title)
        print(self.content + '\n')


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
    eb.set_dictionary_directory(args.dictionary_directory)
    eb.set_search_word(args.search_word)

    for ei in eb.items:
        ei.print_item()
