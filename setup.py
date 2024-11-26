import sys
from setuptools import setup

VERSION = '07.17'

INSTALL_REQUIRES = ['python-docx', 'chardet', 'Levenshtein', 'openpyxl', 'openai', 'llama_cpp_python']
if sys.platform == 'win32':
    INSTALL_REQUIRES.append('pywin32')
if sys.platform != 'darwin':
    INSTALL_REQUIRES.append('tkinterdnd2')

setup(
    name='makdo',
    version=VERSION,
    description='MS WordのファイルをMarkdownで作成・編集します',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    author='Seiichiro HATA',
    author_email='hata48915b@post.nifty.jp',
    url='https://github.com/hata48915b/makdo/',
    license='GPLv3+',
    install_requires=INSTALL_REQUIRES,
    packages=['makdo'],
)
