#!/usr/bin/python3
# Name:         makdo-gui.py
# Version:      v07 Furuichibashi
# Time-stamp:   <2024.04.04-08:49:47-JST>

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
# 2024.04.02 v07 Furuichibashi


# USAGE
# from makdo.makdo_gui import Makdo
# Makdo()


import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import sys
import os
import re
import tempfile
import importlib
import makdo.makdo_md2docx
import makdo.makdo_docx2md


__version__ = 'v07 Furuichibashi'

VERSION = 'v07.04'

WINDOW_SIZE = "601x276"


class Makdo:

    receipt_number = 0

    def __init__(self):

        def drop(event):
            textarea.delete('1.0', 'end')
            Makdo.receipt_number += 1
            textarea.insert('end', '受付番号：' + str(Makdo.receipt_number) + '\n')
            filename = event.data
            filename = re.sub('^{(.*)}$', '\\1', filename)
            basename = os.path.basename(filename)
            textarea.insert('end', '"' + basename + '"を受け取りました\n')
            stderr = sys.stderr
            sys.stderr = tempfile.TemporaryFile(mode='w+')
            if re.match('^.*\\.(m|M)(d|D)$', filename):
                textarea.insert('end', 'docxファイルを作成します\n')
                try:
                    importlib.reload(makdo.makdo_md2docx)
                    m2d = makdo.makdo_md2docx.Md2Docx(filename)
                    m2d.save('')
                    textarea.insert('end', 'docxファイルを作成しました\n')
                except BaseException:
                    sys.stderr.seek(0)
                    textarea.insert('end', sys.stderr.read())
                    textarea.insert('end', 'docxファイルを作成できませんでした\n')
            elif re.match('^.*\\.(d|D)(o|O)(c|C)(x|X)$', filename):
                textarea.insert('end', 'mdファイルを作成します\n')
                try:
                    importlib.reload(makdo.makdo_docx2md)
                    d2m = makdo.makdo_docx2md.Docx2Md(filename)
                    d2m.save('')
                    textarea.insert('end', 'mdファイルを作成しました\n')
                except BaseException:
                    sys.stderr.seek(0)
                    textarea.insert('end', sys.stderr.read())
                    textarea.insert('end', 'mdファイルを作成できませんでした\n')
            else:
                textarea.insert('end', '不適切なファイルです\n')
            sys.stderr = stderr
            textarea.insert('end', '\nここにmdファイル又はdocxファイルをドロップしてください\n')

        win = TkinterDnD.Tk()
        win.geometry(WINDOW_SIZE)
        win.title("MAKDO " + VERSION
                  + " （mdファイルをdocxファイルに、docxファイルをmdファイルに変換します）")

        frame = ttk.Frame(win)
        textarea = tk.Text(frame, width=120, height=30)
        # textarea = tk.Text(frame, width=80, height=20)
        textarea.drop_target_register(DND_FILES)
        textarea.insert('end', 'ここにmdファイル又はdocxファイルをドロップしてください\n')
        textarea.dnd_bind('<<Drop>>', drop)

        frame.pack(expand=True, fill=tk.X, padx=16, pady=8)
        textarea.pack(side=tk.LEFT)

        win.mainloop()


if __name__ == '__main__':
    Makdo()
