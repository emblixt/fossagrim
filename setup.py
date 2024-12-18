from setuptools import setup
import re

verstr = '0.2'

setup(
    name='fossagrim',
    version=verstr,
    packages=[
        'Fossagrim',
        'Fossagrim.io',
        'Fossagrim.plotting',
        'Fossagrim.unit_tests',
        'Fossagrim.utils'
    ],
    url='https://github.com/emblixt/fossagrim',
    license='GNU 3.0',
    author='Erik Mårten Blixt',
    author_email='marten.blixt@fossagrim.com',
    description='Some scripts for modeling forest growth using Heureka',
    long_description='',
    zip_safe=False,
    platforms='any',
    install_requires=[
        'numpy>=1.16.0',
        'matplotlib>=3.0.2',
        'scipy>=1.4.1',
        'setuptools>=47.1.0'
    ],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Intended Audience :: Science/Research',
        'License :: OSI Approved :: GNU Software License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Topic :: Scientific/Engineering',
    ],
)