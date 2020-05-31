# coding:utf-8

from setuptools import setup


setup(
        name='cmp_reader',     # 包名字
        version='0.0.1',   # 包版本
        description='this is test release',   # 简单描述
        author='mofant',  # 作者
        author_email='helehappy@126.com',  # 作者邮箱
        packages=['cmp_reader'],                 # 包
        install_requires = [
            "xlrd"
        ]
)