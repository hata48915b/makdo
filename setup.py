import sys
from setuptools import setup

VERSION = '07.19'

INSTALL_REQUIRES = ['python-docx', 'chardet', 'Levenshtein', 'openpyxl']
if sys.platform == 'win32':
    INSTALL_REQUIRES.append('pywin32')
if sys.platform != 'darwin':
    INSTALL_REQUIRES.append('tkinterdnd2')
# INSTALL_REQUIRES.append('openai')
# INSTALL_REQUIRES.append('llama_cpp_python')

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
