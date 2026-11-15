"""Tests of the sheet/file naming conventions and of the two command line scripts"""
import os
import subprocess
import sys

import openpyxl
import pytest

from Fossagrim.io.common import read_csv_file
from Fossagrim.utils.project_setup import build_result_sheets, build_combine_sheets, sheet_name_prefix
from Fossagrim.unit_tests.conftest import raw_heureka_rows, write_heureka_results

PACKAGE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_build_result_sheets():
    assert build_result_sheets('P1', ['Spruce', 'Pine']) == [
        'P1 Spruce BAU', 'P1 Spruce PRES', 'P1 Pine BAU', 'P1 Pine PRES']


def test_long_project_names_are_truncated_in_sheet_names(capsys):
    name = 'FHF24-0047-01 Long name'
    assert sheet_name_prefix(name) == 'FHF24-0047'
    assert 'too long' in capsys.readouterr().out
    assert all(len(sheet) <= 31 for sheet in build_result_sheets(name, ['Spruce', 'Pine', 'Birch']))
    assert sheet_name_prefix('FHF23-007 v01') == 'FHF23-007 v01'


def test_build_combine_sheets():
    sheets = build_result_sheets('P1', ['Spruce', 'Pine'])
    combine = build_combine_sheets('P1', ['Spruce', 'Pine'], [0.3, 0.7], sheets)
    assert combine == {
        'P1 Combined Stands Spruce-and-Pine BAU': ['P1 Spruce BAU', 0.3, 'P1 Pine BAU', 0.7],
        'P1 Combined Stands Spruce-and-Pine PRES': ['P1 Spruce PRES', 0.3, 'P1 Pine PRES', 0.7]}


def test_build_combine_sheets_single_forest_type():
    sheets = build_result_sheets('P1', ['Pine'])
    combine = build_combine_sheets('P1', ['Pine'], [0.99], sheets)
    assert combine == {'P1 Combined Stands Pine BAU': ['P1 Pine BAU', 1.],
                       'P1 Combined Stands Pine PRES': ['P1 Pine PRES', 1.]}


def test_project_settings_creates_input_files(forvaltningsplan_file, tmp_path):
    from Fossagrim.create_heureka_input_files import project_settings

    folder, _, average_over, stand_id_key, result_file, result_sheets, combine_sheets, csv_file = \
        project_settings('P1', str(tmp_path), forvaltningsplan_file, create_output=True, _unique_id='7')

    assert average_over == {'Spruce': ['1-1', '2-1'], 'Pine': ['3-1']}
    assert stand_id_key == 'Fossagrim ID'
    assert result_file == os.path.join(str(tmp_path), 'P1 Heureka results.xlsx')
    assert openpyxl.load_workbook(result_file).sheetnames == result_sheets
    assert list(combine_sheets.values())[0] == ['P1 Spruce BAU', pytest.approx(2 / 3), 'P1 Pine BAU',
                                                pytest.approx(1 / 3)]

    data, _ = read_csv_file(csv_file)
    assert data['StandId'] == ['P1 Spruce', 'P1 Pine']
    assert data['UniqueID'] == [7, 7]


def test_project_settings_without_output(forvaltningsplan_file, tmp_path):
    from Fossagrim.create_heureka_input_files import project_settings

    project_settings('P1', str(tmp_path), forvaltningsplan_file, create_output=False)
    assert os.listdir(str(tmp_path)) == [os.path.basename(forvaltningsplan_file)]


def run_script(script, *args, cwd):
    """Runs one of the command line scripts from a folder outside the package"""
    return subprocess.run([sys.executable, os.path.join(PACKAGE_DIR, script)] + list(args), cwd=cwd,
                          capture_output=True, text=True)


def test_command_line_scripts_end_to_end(forvaltningsplan_file, tmp_path):
    folder = str(tmp_path)
    created = run_script('create_heureka_input_files.py', 'P1', folder, forvaltningsplan_file, cwd=folder)
    assert created.returncode == 0, created.stderr
    result_file = os.path.join(folder, 'P1 Heureka results.xlsx')
    assert os.path.isfile(os.path.join(folder, 'P1 Averaged stand data.csv'))

    # "Run Heureka" and paste the results into the results file
    years = [2025, 2030, 2035]
    sheets = {}
    for sheet in openpyxl.load_workbook(result_file).sheetnames:
        sheets[sheet] = raw_heureka_rows(years, {
            'Treatment': ('', ['None'] * 3),
            'Total Carbon Stock (dead wood, soil, trees, stumps and roots)': ('ton C/ha', [100, 110, 120]),
        })
    write_heureka_results(result_file, sheets)

    rearranged = run_script('rearrange_heureka_results.py', 'P1', folder, forvaltningsplan_file, result_file,
                            cwd=folder)
    assert rearranged.returncode == 0, rearranged.stderr
    assert 'Rearranged results' in openpyxl.load_workbook(result_file).sheetnames
