import os
from io import open
from setuptools import setup

CUR_DIR = os.path.abspath(os.path.dirname(__file__))
README = os.path.join(CUR_DIR, "README.md")
with open(README, 'r', encoding='utf-8') as fd:
    long_description = fd.read()

setup(
    name = 'excel-exporter',
    version = "0.1.2",
    author = 'Zhang Yu',
    author_email = 'zhangyu836@gmail.com',
    url = 'https://github.com/zhangyu836/django-excel-export',
    packages = ['excel_exporter'],
    include_package_data = True,
    install_requires = ['xltpl >= 0.9.2', 'pydocxtpl >= 0.2.1', 'six'],
    description = ( 'A Django library for exporting data.' ),
    long_description = long_description,
    long_description_content_type = "text/markdown",
    platforms = ["Any platform "],
    license = 'BSD License',
    keywords = ['django', 'Excel', 'xls', 'xlsx', 'spreadsheet', 'workbook', 'template']
)