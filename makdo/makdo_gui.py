#!/usr/bin/python3
# Name:         makdo-gui.py
# Version:      v06 Shimo-Gion
# Time-stamp:   <2024.02.17-18:54:39-JST>

# makdo-gui.py
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


# from makdo_gui import Makdo
# Makdo()


import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import sys
import os
import re
import tempfile
from makdo_md2docx import Md2Docx
from makdo_docx2md import Docx2Md


WINDOW_SIZE = "500x300"


class Makdo:

    def __init__(self):

        def drop(event):
            textarea.delete('1.0', 'end')
            filename = event.data
            filename = re.sub('^{(.*)}$', '\\1', filename)
            basename = os.path.basename(filename)
            textarea.insert('end', '"' + basename + '"を受け取りました\n\n')
            stderr = sys.stderr
            sys.stderr = tempfile.TemporaryFile(mode='w+')
            if re.match('^.*\\.(m|M)(d|D)$', filename):
                textarea.insert('end', 'docxファイルを作成します\n\n')
                try:
                    m2d = Md2Docx(filename)
                    m2d.save('')
                    textarea.insert('end', 'docxファイルを作成しました\n\n')
                except BaseException:
                    sys.stderr.seek(0)
                    textarea.insert('end', sys.stderr.read())
                    textarea.insert('end', 'docxファイルを作成できませんでした\n\n')
            elif re.match('^.*\\.(d|D)(o|O)(c|C)(x|X)$', filename):
                textarea.insert('end', 'mdファイルを作成します\n\n')
                try:
                    d2m = Docx2Md(filename)
                    d2m.save('')
                    textarea.insert('end', 'mdファイルを作成しました\n\n')
                except BaseException:
                    sys.stderr.seek(0)
                    textarea.insert('end', sys.stderr.read())
                    textarea.insert('end', 'mdファイルを作成できませんでした\n\n')
            else:
                textarea.insert('end', '不適切なファイルです\n\n')
            sys.stderr = stderr
            textarea.insert('end', 'ここにmdファイル又はdocxファイルをドロップしてください\n\n')

        win = TkinterDnD.Tk()
        win.geometry(WINDOW_SIZE)
        win.title("MAKDO（mdファイル又はdocxファイルをドロップしてください）")

        frame = ttk.Frame(win)
        textarea = tk.Text(frame, height=20, width=65)
        textarea.drop_target_register(DND_FILES)
        textarea.insert('end', 'ここにmdファイル又はdocxファイルをドロップしてください\n\n')
        textarea.dnd_bind('<<Drop>>', drop)

        frame.pack(expand=True, fill=tk.X, padx=16, pady=8)
        textarea.pack(side=tk.LEFT)

        win.mainloop()


if __name__ == '__main__':
    Makdo()