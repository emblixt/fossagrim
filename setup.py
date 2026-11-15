from setuptools import setup

verstr = '0.3'

setup(
    name='fossagrim',
    version=verstr,
    packages=[
        'Fossagrim',
        'Fossagrim.io',
        'Fossagrim.legacy',
        'Fossagrim.plotting',
        'Fossagrim.unit_tests',
        'Fossagrim.utils'
    ],
    url='https://github.com/emblixt/fossagrim',
    license='GNU 3.0',
    author='Erik Mårten Blixt',
    author_email='marten.blixt@fossagrim.com',
    description='Some scripts for modeling forest growth using Heureka',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    zip_safe=False,
    platforms='any',
    python_requires='>=3.9',
    install_requires=[
        'numpy>=1.16.0',
        'pandas>=2.0',
        'matplotlib>=3.0.2',
        'openpyxl>=3.1',
        'xlsxwriter>=3.0',
        'pypdf>=3.0',
        'setuptools>=47.1.0'
    ],
    extras_require={
        'plots': ['bokeh>=3.0'],  # plotting.misc_plots.plot_collected_stand_data()
        'legacy': ['scipy>=1.4.1'],  # legacy/interpolate.py
        'test': ['pytest>=7.0'],
    },
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Intended Audience :: Science/Research',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Topic :: Scientific/Engineering',
    ],
)
