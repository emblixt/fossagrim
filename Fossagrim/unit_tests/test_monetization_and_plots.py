"""Smoke tests: the monetization file and the QC plots can be created without errors"""
import os

import openpyxl
import pandas as pd

from Fossagrim.io.common import read_excel
from Fossagrim.io.monetization import modify_monetization_file
from Fossagrim.plotting.misc_plots import plot_raw_data, qc_stand_data, list_duplicates


def test_modify_monetization_file(tmp_path):
    f = str(tmp_path / 'Monetization.xlsx')
    writer = pd.ExcelWriter(f, engine='xlsxwriter')
    for sheet in ['Rearranged results', 'Parameters', 'Monetization']:
        writer.book.add_worksheet(sheet)
    writer.close()

    kwargs = {'Position of total area': 400., 'Position of Batch 1 total volume': 1000.,
              'Position of Batch 2 total volume': 500., 'Position of passive forest total volume': 100.,
              'Batch 1 start date': '2025-01-01', 'Batch 2 delay': 5, 'Root net': 0.9, 'Contract length': 30,
              'Rent': 0.02, 'Price growth': 0.01, 'Buffer': 0.2, 'Reserve years': 5, 'Net price': 100,
              'Gross price': 120, 'NIBOR 10yr': 0.04}
    modify_monetization_file(f, **kwargs)

    wb = openpyxl.load_workbook(f)
    assert wb['Monetization']['F1'].value == 40.  # daa -> ha
    assert wb['Monetization']['AI3'].value == 1600.
    assert wb['Parameters']['A2'].value == 'hpap'
    assert set(wb.defined_names.keys()) >= {'hpap', 'hsawn', 'SFsawn', 'k'}


def test_plot_raw_data_creates_folder(tmp_path):
    qc_dir = str(tmp_path / 'QC_plots')
    plot_raw_data({'Year': [2025, 2030], 'A': [1, 2], 'Treatment': ['x', 'y']}, 'sheet', 'Raw Data', qc_dir)
    assert os.path.isfile(os.path.join(qc_dir, 'raw_data_sheet.png'))


def test_qc_stand_data(forvaltningsplan_file, tmp_path):
    table = read_excel(forvaltningsplan_file, 7, 'data')
    qc_dir = str(tmp_path / 'QC_plots')
    qc_stand_data(table, 'test', qc_dir)
    assert sorted(os.listdir(qc_dir)) == ['bestand qc 1.png', 'bestand qc 2.png', 'bestand qc 3.png']


def test_list_duplicates():
    assert list_duplicates(['A', 'B', 'C', 'A']) == ['A']
    assert sorted(list_duplicates([1, 2, 1, 3, 2])) == [1, 2]
