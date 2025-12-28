#!/usr/bin/python3
# Name:         genai.py
# Version:      v01
# Time-stamp:   <2025.12.28-18:49:17-JST>

# genai.py
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


import os
import sys
import re
import tkinter
import threading
import datetime  # save_exchanges


class GenAI:

    genai_name = None
    system_message = 'あなたは誠実で優秀な日本人のアシスタントです。\n' \
        + '特に指示が無い場合は、常に日本語で回答してください。'

    def __init__(self, makdo):
        self.makdo = makdo
        self.model = None

    def _get_genai_head(self) -> tuple[str, str, str]:
        n = MD_TEXT_WIDTH - get_real_width('## 【' + self.genai_name + 'にＸＸ】')
        cnf_head = '## 【' + self.genai_name + 'の設定】' + ('-' * n)
        que_head = '## 【' + self.genai_name + 'に質問】' + ('-' * n)
        ans_head = '## 【' + self.genai_name + 'の回答】' + ('-' * n)
        return cnf_head, que_head, ans_head

    def _get_messages(self, qanda) -> list[dict]:
        cnf_head, que_head, ans_head = self._get_genai_head()
        messages = []
        role, mc = '', ''
        doc = qanda + '\n\n' + ans_head
        for line in doc.split('\n'):
            if line == cnf_head or line == que_head or line == ans_head:
                if role != '' and mc != '':
                    mc = re.sub('^\n+', '', mc)
                    mc = re.sub('\n+$', '', mc)
                    messages.append({'role': role, 'content': mc})
                mc = ''
            if line == cnf_head:
                role = 'system'
            elif line == que_head:
                role = 'user'
            elif line == ans_head:
                role = 'assistant'
            else:
                mc += line + '\n'
        return messages

    def _set_message_on_status_bar(self) -> bool:
        mes = self.genai_name + '（' + self.model + '）に質問しています'
        res = '^Thread-[0-9]+ \\(_' + self.genai_name.lower() + '_.*\\)$'
        if len(threading.enumerate()) > 1:
            for te in threading.enumerate():
                if re.match(res, te.name):
                    self.makdo.set_message_on_status_bar(mes, True)
                    self.makdo.win.after(1000, self._set_message_on_status_bar)
                    return True
        self.makdo.set_message_on_status_bar('', True)
        return False

    def _write_formal_answer(self, answer) -> None:
        if answer == '':
            return -1
        cnf_head, que_head, ans_head = self._get_genai_head()
        self.makdo.sub['autoseparators'] = False
        self.makdo.sub.edit_separator()
        doc = self.makdo.sub.get('1.0', 'end-1c')
        if not re.match('^(.|\n)*\n$', doc):
            self.makdo.sub.insert('end', '\n')
        if not re.match('^(.|\n)*\n\n$', doc):
            self.makdo.sub.insert('end', '\n')
        self.makdo.sub.insert('end', ans_head + '\n\n')
        self.makdo.sub.insert('end', answer + '\n\n')
        self.makdo.sub.edit_separator()
        self.makdo.sub.insert('end', que_head + '\n\n')
        self.makdo.sub.mark_set('insert', 'end-1c')
        self.makdo._put_back_cursor_to_pane(self.makdo.sub)
        self.makdo.sub['autoseparators'] = True
        self.makdo.sub.edit_separator()
        self._paint_genai_lines()

    def _paint_genai_lines(self) -> None:
        for tag in self.makdo.sub.tag_names():
            self.makdo.sub.tag_remove(tag, '1.0', 'end-1c')
        n = 0
        pos = self.makdo.sub.get('1.0', 'end-1c')
        res = '^((?:.|\n)*?)(## 【' + self.genai_name + '...】-+)(\n(?:.|\n)*)$'
        while re.match(res, pos):
            pre = re.sub(res, '\\1', pos)
            key = re.sub(res, '\\2', pos)
            pos = re.sub(res, '\\3', pos)
            beg = '1.0+' + str(n + len(pre)) + 'c'
            end = '1.0+' + str(n + len(pre) + len(key)) + 'c'
            n += len(pre) + len(key)
            self.makdo.sub.tag_add('c-40-1-g-x', beg, end)

    def _warning_dialog(self):
        n, m = '警告', '外部と通信しますか？'
        if not tkinter.messagebox.askyesno(n, m, default='no'):
            return False
        return True


class OpenAI(GenAI):

    genai_name = 'OpenAI'
    notes = '- 外部処理ですので、個人情報の流出に注意してください。\n' \
        + '- 有料ですので、料金に注意してください。\n\n' \

    # KEY

    def _has_key_openai(self) -> bool:
        if 'openai_key' not in vars(self.makdo):
            t = 'OpenAIのキー'
            m = 'OpenAIのキーを入力してください．'
            ok = PasswordDialog(self.makdo.txt, self.makdo, t, m).get_value()
            if ok is None:
                return False
            self.makdo.openai_key = self._enchant(ok)
            self.makdo.show_config_help_message()
        return True

    # 00-94 (32-126)
    pepper = [
        42, 84, 62, 18, 76,  4, 68, 37, 53, 26, 59, 71, 23, 27, 33, 29,
        51, 20, 25, 76, 89, 30, 90, 86, 45, 74,  6, 42, 14,  7, 34, 51,
        31, 31, 13, 74, 68, 32, 41, 44, 17, 39, 34,  4, 41, 25, 79, 94,
        56, 61, 23, 42, 58, 44, 79, 91, 38,  7, 42, 14,  7, 11, 12, 75,
        43, 71,  5,  1,  4, 42, 45, 32, 68, 83, 42,  5, 52, 13, 32, 47,
        39,  7, 48, 90,  1,  1, 53, 80, 42, 57, 64, 56,  5, 82, 30, 15,
        12, 82, 51, 48, 43, 63, 12, 14, 20, 62, 93, 55, 13, 24, 68, 63,
        71, 30, 79, 20, 22, 42, 29, 81, 56, 61, 70, 37, 35, 37, 27, 37,
        57, 82, 58, 71, 83,  4, 57, 62,  3, 31, 40, 48, 21, 51, 87, 49,
        38, 27, 48,  7, 54, 35, 45, 58, 85, 35, 39, 11, 88, 37, 18, 90,
        90, 21, 66, 56, 18, 91, 36, 71, 63, 48, 46, 75, 52, 65, 12, 33,
        42, 72, 41, 31, 86, 59, 24, 56, 27, 94, 23, 47, 92, 42, 15, 15,
        40, 27, 62, 53, 65, 59, 36, 38, 93, 21, 37, 32, 43, 55, 77, 64,
        17, 67, 48, 88, 74, 75, 67,  9, 94, 84,  4,  0, 90, 48, 24, 50,
        22,  6, 27, 39, 38, 10, 68, 46, 90,  5, 66, 34,  4, 40, 50, 31,
        93,  5, 54, 89, 43, 44, 54, 57, 90, 26, 60, 61, 33, 33, 45, 28,
    ]

    @staticmethod
    def _enchant(dechant_word):
        m = len(dechant_word)
        ns = []
        for i in range(m):
            j = i - m // 2
            if j < 0:
                j += m
            # j = i - 1
            # if j == -1:
            #     j = -1
            c_i = dechant_word[i]
            c_j = dechant_word[j]
            n_i = (ord(c_i) - 32) // 5  # 0-18
            n_j = (ord(c_j) - 32) % 5   # 0-4
            n = (n_j * 19) + n_i        # (4 * 19) + 18 = 94
            # n = (n_i * 5) + n_j        # (18 * 5) + 4 = 94
            ns.append(n)
        enchant_word = ''
        for i in range(m):
            n = ns[i]
            n += OpenAI.pepper[i % len(OpenAI.pepper)]
            if n >= 95:
                n -= 95
            e = chr(n + 32)
            enchant_word += e
        return enchant_word

    @staticmethod
    def _dechant(enchant_word):
        m = len(enchant_word)
        ns = []
        for i in range(m):
            e = enchant_word[i]
            n = ord(e) - 32
            n -= OpenAI.pepper[i % len(OpenAI.pepper)]
            if n < 0:
                n += 95
            ns.append(n)
        dechant_word = ''
        for i in range(m):
            j = i + m // 2
            if j >= m:
                j -= m
            # j = i + 1
            # if j == m:
            #     j = 0
            n_i = ns[i] % 19     # 0 -> 18
            n_j = ns[j] // 19    # 0 -> 4
            # n_i = ns[i] // 5    # 0 -> 18
            # n_j = ns[j] % 5     # 0 -> 4
            n = (n_i * 5) + n_j  # (18 * 5) + 4 = 94
            d = chr(n + 32)
            dechant_word += d
        return dechant_word

    # IMPORT

    def _import_openai(self) -> bool:
        try:
            import openai
        except BaseException:
            n, m = 'エラー', '"openai"を\nインポートできませんでした．'
            tkinter.messagebox.showerror(n, m)
            return False
        ok = self._dechant(self.makdo.openai_key)
        self.makdo.set_message_on_status_bar('"OpenAI"を起動しています', True)
        self.openai = openai.OpenAI(api_key=ok)
        self.makdo.set_message_on_status_bar('')
        return True

    # MODEL

    def set_openai_model(self, mother=None) -> bool:
        pane = self.makdo._get_pane()
        if mother is None:
            mother = pane
        mol = self._get_openai_models()
        if mol is None:
            return False
        tit, mes = 'OpenAIのモデルを選択', 'OpenAIのモデルを選択してください．'
        num = -1
        if 'openai_model' in vars(self.makdo):
            om = self.makdo.openai_model
            if om in mol:
                num = mol.index(om)
        rd = RadiobuttonDialog(mother, self.makdo, tit, mes, mol, num)
        val = rd.get_value()
        if (val is not None) and (val != self.makdo.openai_model):
            self.makdo.openai_model = val
            m = 'Openaiのモデルを"' + val + '"に設定しました'
            self.makdo.set_message_on_status_bar(m)
            self.makdo.show_config_help_message()
            return True
        return False

    def _validate_openai_model(self) -> bool:
        models = self._get_openai_models()
        if models is None:
            return False
        if 'openai_model' not in vars(self.makdo):
            self.makdo.openai_model = None
        if self.makdo.openai_model not in models:
            if not self.set_openai_model():
                return False
        self.model = self.makdo.openai_model
        return True

    def _get_openai_models(self) -> list:
        try:
            tmp = []
            for om in self.openai.models.list():
                tmp.append([om.created, om.id])
            models = []
            for m in reversed(sorted(tmp)):
                models.append(m[1])
            return models
        except BaseException:
            n, m = 'エラー', '"openai"のモデルを\n取得できませんでした．'
            tkinter.messagebox.showerror(n, m)
            return None

    # ASK

    def open_openai(self) -> bool:
        m = 'モデルは"' + self.makdo.openai_model + '"が設定されています'
        self.makdo.set_message_on_status_bar(m)
        # PROMPT
        cnf_head, que_head, ans_head = self._get_genai_head()
        if 'openai_qanda' not in vars(self):
            self.openai_qanda = self.notes \
                + cnf_head + '\n\n' + self.system_message + '\n\n' \
                + que_head + '\n\n'
        self.makdo.txt.focus_force()
        self.makdo._execute_sub_pane = self.ask_openai
        self.makdo._close_sub_pane = self.close_openai
        self.makdo._open_sub_pane(self.openai_qanda, False, '質問')
        self.makdo.sub.mark_set('insert', 'end-1c')
        self.makdo.sub.edit_separator()
        self._paint_genai_lines()
        return True

    def ask_openai(self) -> None:
        thread_1 = threading.Thread(
            target=self._openai_ask_openai, daemon=True)
        thread_2 = threading.Thread(
            target=self._set_message_on_status_bar, daemon=True)
        thread_1.start()
        thread_2.start()

    def _openai_ask_openai(self) -> bool:
        model = self.makdo.openai_model
        qanda = self.makdo.sub.get('1.0', 'end-1c')
        messages = self._get_messages(qanda)
        if not self._warning_dialog():
            return False
        self.makdo.set_message_on_status_bar('OpenAIに質問しています', True)
        try:
            output = self.openai.chat.completions.create(
                model=model, messages=messages, n=1,  # one answer
            )
        except BaseException:
            n, m = 'エラー', '"openal"から\n回答を得られませんでした．'
            tkinter.messagebox.showerror(n, m)
            return
        self.makdo.set_message_on_status_bar('')
        answer = output.choices[0].message.content  # one answer
        # answer = adjust_line(answer)
        self._write_formal_answer(answer)

    def close_openai(self) -> None:
        del self.makdo._execute_sub_pane
        del self.makdo._close_sub_pane
        self.openai_qanda = self.makdo.sub.get('1.0', 'end-1c')
        self.makdo.set_message_on_status_bar('')
        self.makdo._close_sub_pane()


class Ollama(GenAI):

    genai_name = 'Ollama'
    notes = '- 内部処理又はクラウドですので、情報を外部に出しません。\n' \
        + '- 無料ですので、料金は発生しません。\n\n' \

    # IMPORT

    def _import_ollama(self) -> bool:
        if 'ollama' in sys.modules:
            return True
        try:
            import ollama
        except BaseException:
            n, m = 'エラー', '"ollama"を\nインポートできませんでした．'
            tkinter.messagebox.showerror(n, m)
            return False
        self.ollama = ollama
        return True

    # TEST

    def _test_ollama(self) -> bool:
        try:
            self.ollama.ps()
        except BaseException:
            n, m = 'エラー', '"ollama"を\n起動できませんでした．'
            tkinter.messagebox.showerror(n, m)
            return False
        return True

    # MODEL

    def set_ollama_model(self, mother=None) -> bool:
        pane = self.makdo._get_pane()
        if mother is None:
            mother = pane
        # GET INSTALLED MODELS
        mol = self._get_installed_ollama_models()
        if mol is None:
            return False
        tmp = []
        for m in mol:
            t = ['', '', 0, m]
            if re.match('^.*-cloud$', m):
                t[0], t[3] = '@', '@' + m  # for cloud
            m = re.sub('-[^-]+$', '', m)
            lt = m.split(':')
            if len(lt) > 0:
                t[1] = lt[0]
            if len(lt) > 1:
                s = lt[1]
                res = '([A-Za-z]*)([0-9]*\\.?[0-9])([mbt])$'
                if re.match(res, s):
                    n = re.sub(res, '\\2', s)
                    u = re.sub(res, '\\3', s)
                    if u == 'm':
                        t[2] = int(float(n) * 1_000_000)
                    elif u == 'b':
                        t[2] = int(float(n) * 1_000_000_000)
                    elif u == 't':
                        t[2] = int(float(n) * 1_000_000_000_000)
                elif s != '':
                    t[2] = 999_999_999_999_999
            tmp.append(t)
        tmp.sort()
        mol = []
        for t in tmp:
            mol.append(t[3])
        # SET TITLE AND MESSAGE
        tit, mes = 'Ollamaのモデルを選択', 'Ollamaのモデルを選択してください．'
        # GET THE CURRENT MODEL
        num = -1
        if 'ollama_model' in vars(self.makdo):
            om = self.makdo.ollama_model
            if re.match('^.*-cloud$', om):
                om = '@' + om  # for cloud
            if om in mol:
                num = mol.index(om)
        # GET A NEW MODEL
        rd = RadiobuttonDialog(mother, self.makdo, tit, mes, mol, num)
        val = rd.get_value()
        if (val is not None) and (val != self.makdo.ollama_model):
            val = re.sub('^@', '', val)  # for cloud
            self.makdo.ollama_model = val
            m = 'Ollamaのモデルを"' + val + '"に設定しました'
            self.makdo.set_message_on_status_bar(m)
            self.makdo.show_config_help_message()
            return True
        return False

    def _validate_ollama_model(self) -> bool:
        models = self._get_installed_ollama_models()
        if models is None:
            return False
        if 'ollama_model' not in vars(self.makdo):
            self.makdo.ollama_model = None
        if self.makdo.ollama_model not in models:
            if not self.set_ollama_model():
                return False
        self.model = self.makdo.ollama_model
        return True

    def _get_installed_ollama_models(self) -> list:
        try:
            models = []
            for om in self.ollama.list().models:
                models.append(om.model)
        except BaseException:
            n, m = 'エラー', '"ollama"のモデルを\n取得できませんでした．'
            tkinter.messagebox.showerror(n, m)
            return None
        return models

    # OPEN OLLAMA

    def open_ollama(self) -> bool:
        m = 'モデルは"' + self.model + '"が設定されています'
        self.makdo.set_message_on_status_bar(m)
        # PROMPT
        cnf_head, que_head, ans_head = self._get_genai_head()
        if 'ollama_qanda' not in vars(self):
            self.ollama_qanda = self.notes \
                + cnf_head + '\n\n' + self.system_message + '\n\n' \
                + que_head + '\n\n'
        self.makdo.txt.focus_force()
        self.makdo._execute_sub_pane = self.ask_ollama
        self.makdo._close_sub_pane = self.close_ollama
        self.makdo._open_sub_pane(self.ollama_qanda, False, '質問')
        self.makdo.sub.mark_set('insert', 'end-1c')
        self.makdo.sub.edit_separator()
        self._paint_genai_lines()
        return True

    def ask_ollama(self) -> None:
        if self.makdo.current_pane != 'sub':
            self.ask_ollama_on_main_pane()
        else:
            self.ask_ollama_on_sub_pane()

    def ask_ollama_on_sub_pane(self) -> None:
        thread_1 = threading.Thread(
            target=self._ollama_ask_ollama_on_sub_pane, daemon=True)
        thread_2 = threading.Thread(
            target=self._set_message_on_status_bar, daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_ask_ollama_on_sub_pane(self) -> bool:
        qanda = self.makdo.sub.get('1.0', 'end-1c')
        txt = self.makdo.txt.get('1.0', 'end-1c')
        qanda = qanda.replace('\n%[本文]%\n', '\n' + txt + '\n')
        qanda = self._insert_files(qanda)
        messages = self._get_messages(qanda)
        response = self._execute_ollama(None, None, messages)
        if response is None:
            return False
        answer = response.message.content
        self._write_formal_answer(answer)
        return True

    def close_ollama(self) -> None:
        del self.makdo._execute_sub_pane
        del self.makdo._close_sub_pane
        self.ollama_qanda = self.makdo.sub.get('1.0', 'end-1c')
        self.makdo.set_message_on_status_bar('')
        self.makdo._close_sub_pane()

    # ASK OLLAMA

    def ask_ollama_on_main_pane(self):
        thread_1 = threading.Thread(
            target=self._ollama_ask_ollama_on_main_pane, daemon=True)
        thread_2 = threading.Thread(
            target=self._set_message_on_status_bar, daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_ask_ollama_on_main_pane(self) -> bool:
        self.makdo.txt.mark_set('ollama', 'insert')
        sc = self.system_message
        doc = self._get_document(self.makdo.txt)
        doc = self._insert_files(doc)
        uc = ''
        for line in doc.split('\n'):
            uc += line
        response = self._execute_ollama(sc, uc + '\n' + doc)
        if response is None:
            return False
        answer = response.message.content
        self._write_simple_answer(self.makdo.txt, answer)
        self.makdo.txt.tag_remove('ollama', '1.0', 'end')
        self.makdo.cancel_region(self.makdo.txt)

    # PICK UP PROPER NOUNS

    def pick_up_proper_nouns(self):
        thread_1 = threading.Thread(target=self._ollama_pick_up_proper_nouns,
                                    daemon=True)
        thread_2 = threading.Thread(target=self._set_message_on_status_bar,
                                    daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_pick_up_proper_nouns(self) -> bool:
        pane = self.makdo._get_pane()
        pane.mark_set('ollama', 'insert')
        doc = self._get_document(pane)
        sc = self.system_message
        uc = '次の文章から固有名詞を抽出してください。\n' \
            + '回答はjson形式で回答してください。\n'
        response = self._execute_ollama(sc, uc + '\n' + doc)
        if response is None:
            return False
        answer_json = response.message.content
        answer_list = self._json_to_list(answer_json)
        answer = ''
        for a in answer_list:
            answer += a + '\n'
        self._write_simple_answer(self.makdo.txt, answer)
        pane.tag_remove('ollama', '1.0', 'end')
        self.makdo.cancel_region(pane)
        return True

    @staticmethod
    def _json_to_list(answer_json):
        tmp = answer_json
        tmp = tmp.replace('\n', '')
        tmp = re.sub('^(.|\n)*?\\[', '', tmp)
        tmp = re.sub('\\](.|\n)*?$', '', tmp)
        tmp = re.sub('^\\s*"', '', tmp)
        tmp = re.sub('"\\s*$', '', tmp)
        tmp = re.sub('",\\s+"', '","', tmp)
        answer_list = tmp.split('","')
        return answer_list

    # FIND TYPOS

    def find_typos(self):
        thread_1 = threading.Thread(target=self._ollama_find_typos,
                                    daemon=True)
        thread_2 = threading.Thread(target=self._set_message_on_status_bar,
                                    daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_find_typos(self) -> bool:
        pane = self.makdo._get_pane()
        pane.mark_set('ollama', 'insert')
        doc = self._get_document(pane)
        sc = self.system_message
        uc = '次の文章に誤字脱字があれば、指摘してください。\n'
        response = self._execute_ollama(sc, uc + '\n' + doc)
        answer = response.message.content
        self._write_simple_answer(self.makdo.txt, answer)
        pane.tag_remove('ollama', '1.0', 'end')
        self.makdo.cancel_region(pane)
        return True

    # TOOLS

    @staticmethod
    def _get_document(pane) -> str:
        if pane.tag_ranges('sel'):
            doc = pane.get('sel.first', 'sel.last')
        elif 'akauni' in pane.mark_names():
            doc = ''
            doc += pane.get('akauni', 'insert')
            doc += pane.get('insert', 'akauni')
        else:
            doc = pane.get('insert', 'end-1c')
        return doc

    @staticmethod
    def _insert_files(doc):
        new = ''
        res = '^%\\[(.+)\\]%$'
        for line in doc.split('\n'):
            if re.match(res, line):
                fname = re.sub(res, '\\1', line)
                try:
                    with open(fname, 'r') as f:
                        t = f.read()
                        line = t
                except BaseException:
                    n, m = 'エラー', '"fn"を\n挿入できませんでした．'
                    tkinter.messagebox.showerror(n, m)
            new += line + '\n'
        return new

    def _execute_ollama(self, system_content, user_content, messages=None):
        if re.match('.*-cloud$', self.makdo.ollama_model):
            if not self._warning_dialog():
                return None
        if messages is None:
            messages = [
                {'role': 'system', 'content': system_content},
                {'role': 'user', 'content': user_content}
            ]
        try:
            response = self.ollama.chat(
                model=self.makdo.ollama_model,
                messages=messages,
                think=False,  # for reasoning model
                # options={ "temperature": 0, "num_ctx": 512 }
            )
        except BaseException:
            n, m = 'エラー', '"ollama"を\n実行できませんでした．'
            tkinter.messagebox.showerror(n, m)
            return None
        return response

    def _write_simple_answer(self, pane, answer) -> None:
        answer = re.sub('^\n+', '', answer)
        answer = re.sub('\n+$', '', answer)
        if answer == '':
            answer = '（空）'
        pre = pane.get('1.0', 'ollama')
        pos = pane.get('ollama', 'end-1c')
        rmd = re.sub('^((.|\n)*?<!--(.|\n)*?-->)*', '', pre)
        if re.match('^(.|\n)*<!--(.|\n)*$', rmd):
            answer = '-----\n' + answer + '\n-----'
        else:
            answer = '<!--\n' + answer + '\n-->'
        self.makdo._insert_line_break_as_necessary('ollama')
        pane.insert('ollama', answer)

    def save_ollama_exchanges(self) -> bool:
        if 'ollama_qanda' not in vars(self):
            return False
        if not os.path.exists(CONFIG_DIR + '/ollama'):
            os.mkdir(CONFIG_DIR + '/ollama')
        datetime.datetime.now()
        time = datetime.datetime.now().strftime("%y%m%d%H%M%S")
        sc = self.system_message
        uc = '次の対話のタイトルを1行かつ20文字程度で考えてください。\n' \
            + '「Ollama」という単語は入れないでください。\n' \
            + '答えはMarkdownではなくText形式でお願いします。\n\n' \
            + self.ollama_qanda
        answer = self._execute_ollama(sc, uc).message.content
        answer = answer.replace('\n', '')
        fn = CONFIG_DIR + '/ollama/' + time + '-' + answer + '.md'
        with open(fn, 'w') as f:
            f.write(re.sub('^[^#]+', '', self.ollama_qanda))
        return True

    def open_ollama_exchanges(self) -> bool:
        if 'ollama_qanda' in vars(self):
            n, m = '確認', '今までの対話を保存しますか？'
            if tkinter.messagebox.askyesno(n, m, default='yes'):
                self.save_ollama_exchanges()
        ti = 'ファイルを読み込む'
        ft = [('可能な形式', '.md .docx'), ('Markdown', '.md')]
        id = CONFIG_DIR + '/ollama'
        filename = tkinter.filedialog.askopenfilename(
            title=ti, filetypes=ft, initialdir=id)
        doc = self.notes
        with open(filename, 'r') as f:
            doc += f.read()
        self.ollama_qanda = doc
        self.open_ollama()
        return True
