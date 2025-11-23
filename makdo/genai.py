#!/usr/bin/python3
# Name:         genai.py
# Version:      v01
# Time-stamp:   <2025.11.24-04:19:58-JST>

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


def _get_genai_head(genai) -> tuple[str, str, str]:
    n = MD_TEXT_WIDTH - get_real_width('## 【' + genai + 'にＸＸ】')
    cnf_head = '## 【' + genai + 'の設定】' + ('-' * n)
    que_head = '## 【' + genai + 'に質問】' + ('-' * n)
    ans_head = '## 【' + genai + 'の回答】' + ('-' * n)
    return cnf_head, que_head, ans_head


def _get_messages(qanda, genai) -> list[dict]:
    cnf_head, que_head, ans_head = _get_genai_head(genai)
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


def _write_answer(makdo, genai, answer) -> None:
    if answer == '':
        return -1
    cnf_head, que_head, ans_head = _get_genai_head(genai)
    makdo.sub['autoseparators'] = False
    makdo.sub.edit_separator()
    doc = makdo.sub.get('1.0', 'end-1c')
    if not re.match('^(.|\n)*\n$', doc):
        makdo.sub.insert('end', '\n')
    if not re.match('^(.|\n)*\n\n$', doc):
        makdo.sub.insert('end', '\n')
    makdo.sub.insert('end', ans_head + '\n\n')
    makdo.sub.insert('end', answer + '\n\n')
    makdo.sub.edit_separator()
    makdo.sub.insert('end', que_head + '\n\n')
    makdo.sub.mark_set('insert', 'end-1c')
    makdo._put_back_cursor_to_pane(makdo.sub)
    makdo.sub['autoseparators'] = True
    makdo.sub.edit_separator()
    _paint_genai_lines(makdo, genai)


def _paint_genai_lines(makdo, genai) -> None:
    for tag in makdo.sub.tag_names():
        makdo.sub.tag_remove(tag, '1.0', 'end-1c')
    n = 0
    pos = makdo.sub.get('1.0', 'end-1c')
    res = '^((?:.|\n)*?)(## 【' + genai + '...】-+)(\n(?:.|\n)*)$'
    while re.match(res, pos):
        pre = re.sub(res, '\\1', pos)
        key = re.sub(res, '\\2', pos)
        pos = re.sub(res, '\\3', pos)
        beg = '1.0+' + str(n + len(pre)) + 'c'
        end = '1.0+' + str(n + len(pre) + len(key)) + 'c'
        n += len(pre) + len(key)
        makdo.sub.tag_add('c-40-1-g-x', beg, end)


class Ollama:

    def __init__(self, makdo):
        self.makdo = makdo

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

    def _validate_ollama_model(self) -> bool:
        models = self._get_installed_ollama_models()
        if models is None:
            return False
        if 'ollama_model' not in vars(self.makdo):
            self.makdo.ollama_model = None
        if self.makdo.ollama_model not in models:
            return self.set_ollama_model()
        return True

    def _get_installed_ollama_models(self) -> list:
        try:
            models = []
            for om in self.ollama.list().models:
                models.append(om.model)
        except BaseException:
            n, m = 'エラー', '"ollama"を\n実行できませんでした．'
            tkinter.messagebox.showerror(n, m)
            return None
        return models

    # OPEN OLLAMA

    def open_ollama(self) -> bool:
        m = 'モデルは"' + self.makdo.ollama_model + '"が設定されています'
        self.makdo.set_message_on_status_bar(m)
        # PROMPT
        cnf_head, que_head, ans_head = _get_genai_head('Ollama')
        if 'ollama_qanda' not in vars(self):
            self.ollama_qanda \
                = '- 内部処理ですので、情報を外部に送信しません。\n' \
                + '- 無料ですので、料金は発生しません。\n\n' \
                + cnf_head + '\n\n' \
                + 'あなたは誠実で優秀な日本人のアシスタントです。\n' \
                + '特に指示が無い場合は、常に日本語で回答してください。\n\n' \
                + que_head + '\n\n'
        self.makdo.txt.focus_force()
        self.makdo._execute_sub_pane = self.ask_ollama
        self.makdo._close_sub_pane = self.close_ollama
        self.makdo._open_sub_pane(self.ollama_qanda, False, '質問')
        self.makdo.sub.mark_set('insert', 'end-1c')
        self.makdo.sub.edit_separator()
        _paint_genai_lines(self.makdo, 'Ollama')
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
            target=self._set_message_ollama, daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_ask_ollama_on_sub_pane(self) -> bool:
        qanda = self.makdo.sub.get('1.0', 'end-1c')
        txt = self.makdo.txt.get('1.0', 'end-1c')
        qanda = qanda.replace('\n%[本文]%\n', '\n' + txt + '\n')
        qanda = self._insert_files(qanda)
        messages = _get_messages(qanda, 'Ollama')
        response = self._execute_ollama(None, None, messages)
        if response is None:
            return False
        answer = response.message.content
        _write_answer(self.makdo, 'Ollama', answer)
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
            target=self._set_message_ollama, daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_ask_ollama_on_main_pane(self) -> bool:
        self.makdo.txt.mark_set('ollama', 'insert')
        sc = 'あなたは誠実で優秀な日本人のアシスタントです。\n' \
            + '特に指示が無い 場合は、常に日本語で回答してください。'
        doc = self._get_document(self.makdo.txt)
        doc = self._insert_files(doc)
        uc = ''
        for line in doc.split('\n'):
            uc += line
        response = self._execute_ollama(sc, uc + '\n' + doc)
        if response is None:
            return False
        answer = response.message.content
        answer = re.sub('\n+$', '', answer)
        pre = self.makdo.txt.get('1.0', 'ollama')
        pos = self.makdo.txt.get('ollama', 'end-1c')
        rmd = re.sub('^((.|\n)*?<!--(.|\n)*?-->)*', '', pre)
        if re.match('^(.|\n)*<!--(.|\n)*$', rmd):
            answer = '-----\n' + answer + '\n-----'
        else:
            answer = '<!--\n' + answer + '\n-->'
        if len(pre) > 0 and pre[-1] != '\n':
            answer = '\n' + answer
        if len(pos) > 0 and pos[0] != '\n':
            answer = answer + '\n'
        self.makdo.txt.insert('ollama', answer)
        self.makdo.txt.tag_remove('ollama', '1.0', 'end')
        self.makdo.cancel_region(self.makdo.txt)

    # SET MODEL

    def set_ollama_model(self, mother=None) -> bool:
        pane = self.makdo._get_pane()
        if mother is None:
            mother = pane
        mol = self._get_installed_ollama_models()
        if mol is None:
            return False
        tit = 'Ollamaのモデルを選択'
        mes = 'Ollamaのモデルを選択してください．'
        num = -1
        if 'ollama_model' in vars(self.makdo):
            om = self.makdo.ollama_model
            if om in mol:
                num = mol.index(om)
            rd = RadiobuttonDialog(mother, self.makdo, tit, mes, mol, num)
        val = rd.get_value()
        if (val is not None) and (val != self.makdo.ollama_model):
            self.makdo.ollama_model = val
            m = 'Ollamaのモデルを"' + val + '"に設定しました'
            self.makdo.set_message_on_status_bar(m)
            self.makdo.show_config_help_message()
            return True
        return False

    # PICK UP PROPER NOUNS

    def pick_up_proper_nouns(self):
        thread_1 = threading.Thread(target=self._ollama_pick_up_proper_nouns,
                                    daemon=True)
        thread_2 = threading.Thread(target=self._set_message_ollama,
                                    daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_pick_up_proper_nouns(self) -> bool:
        pane = self.makdo._get_pane()
        pane.mark_set('ollama', 'insert')
        doc = self._get_document(pane)
        sc = 'あなたは誠実で優秀な日本人のアシスタントです。\n' \
            + '特に指示が無い 場合は、常に日本語で回答してください。'
        uc = '次の文章から固有名詞を抽出してください。\n' \
            + '回答はjson形式で回答してください。\n'
        response = self._execute_ollama(sc, uc + '\n' + doc)
        if response is None:
            return False
        answer_json = response.message.content
        answer_list = self._json_to_list(answer_json)
        self.makdo._insert_line_break_as_necessary()
        pane.insert('ollama', '<!--\n')
        for a in answer_list:
            pane.insert('ollama', a + '\n')
        pane.insert('ollama', '-->')
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
        thread_2 = threading.Thread(target=self._set_message_ollama,
                                    daemon=True)
        thread_1.start()
        thread_2.start()

    def _ollama_find_typos(self) -> bool:
        pane = self.makdo._get_pane()
        pane.mark_set('ollama', 'insert')
        doc = self._get_document(pane)
        sc = 'あなたは誠実で優秀な日本人のアシスタントです。\n' \
            + '特に指示が無い 場合は、常に日本語で回答してください。'
        uc = '次の文章に誤字脱字があれば、指摘してください。\n'
        response = self._execute_ollama(sc, uc + '\n' + doc)
        if response is None:
            return False
        answer = response.message.content
        answer = re.sub('\n+$', '', answer)
        self.makdo._insert_line_break_as_necessary()
        pane.insert('ollama', '<!--\n')
        pane.insert('ollama', answer + '\n')
        pane.insert('ollama', '-->')
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

    def _set_message_ollama(self) -> bool:
        message = 'Ollama（' + self.makdo.ollama_model + '）に質問しています'
        if len(threading.enumerate()) > 1:
            for te in threading.enumerate():
                if re.match('^Thread-[0-9]+ \\(_ollama_.*\\)$', te.name):
                    self.makdo.set_message_on_status_bar(message, True)
                    self.makdo.win.after(1_000, self._set_message_ollama)
                    return True
        self.makdo.set_message_on_status_bar('', True)
        return False

    def _execute_ollama(self, system_content, user_content, messages=None):
        if not self._cloud_dialog():
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

    def _cloud_dialog(self):
        if re.match('.*-cloud$', self.makdo.ollama_model):
            n, m = '警告', '外部と通信しますか？'
            if not tkinter.messagebox.askyesno(n, m, default='no'):
                return False
        return True
