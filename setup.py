import sys
from setuptools import setup

VERSION = '07.12'

INSTALL_REQUIRES = ['python-docx', 'chardet', 'tkinterdnd2', 'openpyxl', 'openai', 'Levenshtein']
if sys.platform == 'win32':
    INSTALL_REQUIRES.append('pywin32')

setup(
    name='makdo',
    version=VERSION,
    description='日本の公用文書（司法文書、行政文書）をMarkdown形式とMicrosoft Word形式との間で変換します',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    author='Seiichiro HATA',
    author_email='hata48915b@post.nifty.jp',
    url='https://github.com/hata48915b/makdo/',
    license='GPLv3+',
    install_requires=INSTALL_REQUIRES,
    packages=['makdo'],
)
