"""
Manual checks that used to be unittest.TestCase classes inside fossagrim_io.py.

They depend on files in the Fossagrim OneDrive and mostly print results instead of asserting anything, so
they are NOT part of the automatic test suite (see Fossagrim/unit_tests for that). Run single checks manually,
e.g. from PyCharm, when debugging with real project files.
"""
import os
import unittest

import matplotlib.pyplot as plt
import pandas as pd

import Fossagrim.plotting.misc_plots as fpp
import Fossagrim.io.fossagrim_io as fio_shim
import Fossagrim.utils.projects_new as projects_new
from Fossagrim.io.fossagrim_io import *  # noqa: F401,F403
from Fossagrim.legacy.legacy_io import *  # noqa: F401,F403
from Fossagrim.utils.definitions import heureka_standdata_keys, heureka_standdata_desc
from Fossagrim.utils.monetization_parameters import variables_used_in_monetization


class IoChecks(unittest.TestCase):
    def test_arrange_export(self):
        stand_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0047 Åmli\\FHF24-0047 Forvaltningsplan_v02.xlsx"
        stand_id_key = 'Fossagrim ID'
        sheet_name='data'
        header = 7
        table = read_excel(stand_file, header, sheet_name)
        # print(list(table.keys()))
        keyword_args = (
            arrange_export_of_one_stand(1,
                                        table,
                                        ['PlantDensity', 'UserDefinedVariable2_RotationPeriod',
                                         'UserDefinedVariable3_ThinningYear']))
        print(keyword_args)

    def test_read_fossagrim_treatment(self):
        file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0014 Arne Tag\\FHF24-0014 Averaged treatment.csv"
        print(read_fossagrim_treatment(file, 'FHF24-0014', 'Spruce'))
        print(read_fossagrim_treatment(file, 'FHF24-0014', 'Rubbish'))

    def test_read_raw_heureka_results(self):
        file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-999 Testing only\\FHF23-999 Heureka results.xlsx"
        sheet_name = 'FHF23-999 Pine BAU'
        result = read_raw_heureka_results(file, sheet_name, variables_used_in_monetization, False)
        print(result.keys())
        # print(result)

    def test_modify_monetization_file(self):
        mf = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-999 Testing only\\Test Monetization.xlsx"
        sf = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-999 Testing only\\FHF23-999 Bestandsutvalg.xlsx"
        pf = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings - WIP.xlsx"
        tag = 'FHF23-999'

        # Create empty monetization file
        writer = pd.ExcelWriter(mf, engine='xlsxwriter')
        wb = writer.book
        writer.close()

        _kwargs, combine_fractions = get_kwargs_from_stand(sf, pf, tag)

        modify_monetization_file(mf, **_kwargs)

    def test_write_excel_with_equations(self):
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill
        from openpyxl.styles import Alignment

        colors = ['00660066', '00FFFFCC',
                  '00FF0000', '0000FF00', '00660066']
        fillers = []

        for color in colors:
            temp = PatternFill(patternType='solid',
                               fgColor=color)
            fillers.append(temp)

        dir_path = os.path.dirname(os.path.realpath(__file__))
        efile = os.path.join(dir_path, 'test.xlsx')
        sheet_names = ['One', 'Two', 'Three']

        writer = pd.ExcelWriter(efile, engine='xlsxwriter')
        wb = writer.book
        for sheet in sheet_names:
            ws = wb.add_worksheet(sheet)
        writer.close()

        wb = load_workbook(efile)
        ws = wb['Three']
        ws.cell(1, 1).value = 'Constant'
        ws.cell(1, 2).value = 4

        ws = wb['Two']
        ws.cell(1, 1).value = '=One!A1*One!B1'
        ws.cell(1, 2).value = '=One!A1*Three!$B$1'
        for i in range(4):
            ws.cell(i + 2, 1).value = '=One!A{}*One!B{}'.format(i + 2, i + 2)
            ws.cell(i + 2, 2).value = '=One!A{}*Three!$B$1'.format(i + 2)
            ws.cell(i + 2, 1).fill = fillers[0]
            ws.cell(i + 2, 2).fill = fillers[1]
        for my_row in ws['C6:D12']:
            for my_cell in my_row:
                my_cell.fill = fillers[3]

        ws.cell(1, 4).value = 'TEST'
        ws.cell(1, 4).alignment = Alignment(horizontal='center')
        ws.cell(1, 4).fill = fillers[2]
        ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=8)
        wb.save(efile)

    def test_rearrange(self):
        file = "C:\\User\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-999 Testing only\\FHF23-999 Heureka results.xlsx"
        rearrange_raw_heureka_results(file, ['FHF23-999 Avg Stand-Pine BAU', 'FHF23-999 Avg Stand-Pine PRES'])

    def test_write_csv_file(self):
        dir_path = os.path.dirname(os.path.realpath(__file__))
        write_csv_file(os.path.join(dir_path, 'test.csv'),
                       heureka_standdata_keys,
                       heureka_standdata_desc,
                       test=8, DGV=46, SiteIndexSpecies='G')
        write_csv_file(os.path.join(dir_path, 'test.csv'),
                       heureka_standdata_keys,
                       heureka_standdata_desc,
                       append=True, SiteIndexSpecies='T')
        write_csv_file(os.path.join(dir_path, 'test2.csv'),
                       heureka_standdata_keys,
                       heureka_standdata_desc)

    def test_export_fossagrim_stand(self):
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # read_from_file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra\\FHF23-003 Bestandsutvalg.xlsx"
        read_from_file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-005 Kvistaul\\FHF23-005 Bestandsutvalg.xlsx"
        write_to_file = os.path.join(dir_path, 'standdata.csv')

        export_fossagrim_stand_to_heureka(read_from_file, write_to_file,
                                          stand_id_key='Fossagrim ID',
                                          average_over={
                                              'Spruce': ['FHF23-005-2', 'FHF23-005-13'],
                                              'Pine': ['FHF23-005-26', 'FHF23-005-27']
                                          }
                                          )
        # export_fossagrim_stand_to_heureka(read_from_file, write_to_file, average_over={'Spruce': [67, 70]},
        #                                   this_stand_only=67)

    def test_export_fossagrim_treatment(self):
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # read_from_file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra\\FHF23-003 Bestandsutvalg.xlsx"
        read_from_file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-005 Kvistaul\\FHF23-005 Bestandsutvalg.xlsx"
        write_to_file = os.path.join(dir_path, 'treatment.csv')

        export_fossagrim_treatment(read_from_file, write_to_file,
                                   stand_id_key='Fossagrim ID',
                                   average_over={
                                       'Spruce': ['FHF23-005-2', 'FHF23-005-13'],
                                       'Pine': ['FHF23-005-26', 'FHF23-005-27']
                                   }
                                   )
        # export_fossagrim_treatment(read_from_file, write_to_file, average_over={'Spruce': [67, 70]},
        #                            this_stand_only=67)

    def test_get_kwargs_from_stand(self):
        kwargs, combine_fractions = get_kwargs_from_stand(
            # "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-999 Testing only\\FHF23-999 Bestandsutvalg.xlsx",
            # "C:\\Users\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0014 Arne Tag\\Bestandsoversikt 16052024.xlsx",
            "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0047 Åmli\\FHF24-0047 Forvaltningsplan_v02.xlsx",
            "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings.xlsx",
            "FHF24-0047 v02"
        )
        for _key in kwargs:
            print(_key, ':', kwargs[_key])
        print('-x-')
        print(combine_fractions)
        if None in combine_fractions:
            print('None is found')

        # f = 'C:\\tmp\\test.xlsx'
        # wb = openpyxl.load_workbook(f)
        # ws = wb['Sheet1']
        # for i, _key in enumerate(kwargs):
        #     ws['A{}'.format(i+1)] = _key
        #     ws['B{}'.format(i+1)] = kwargs[_key]
        # wb.save(f)

    def test_get_carbon_effect(self):
        # result_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0016 Margrete Folsland\\FHF24-0016 Heureka results.xlsx"
        # project_tag = "FHF24-0016"
        # wood_species = "Pine"
        result_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0047 Åmli\\FHF24-0047 v01 Heureka results.xlsx"
        project_tag = "FHF24-0047 v01 Pine"
        save_plot_to = result_file.replace("xlsx", "png")
        print(save_plot_to)
        # avg = get_carbon_effect(result_file, project_tag, wood_species=wood_species, save_plot_to=save_plot_to)
        avg = get_carbon_effect(result_file, project_tag, save_plot_to=save_plot_to)
        print(avg)
        plt.show()

    def test_combine_raw_herureka_results(self):
        file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-999 Testing only\\FHF23-999 Heureka results.xlsx"
        sheet_name = 'FHF23-999 Pine BAU'

        def func(x, y):
            return x - y

        variables = {'x': 'Total Carbon Stock (dead wood, soil, trees, stumps and roots)',
                     'y': 'Soil Carbon Stock'}

        result = combine_raw_heureka_results(file, sheet_name, func, variables)
        print(result)

    def test_get_nature_carbon(self):
        file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0047 Åmli\\FHF24-0047 v01 Heureka results.xlsx"
        stand_id = 'FHF24-0047 v01 Spruce'
        print(get_nature_and_climate_effect(file, stand_id, average_years=30, verbose=True))

    def test_collect_all_stand(self):
        collect_all_stand_data(out_file="C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\TEST COLLECT STAND DATA.csv",
                               dry_run=False)

    def test_read_rearranged_results(self):
        filename = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0047 Åmli\\FHF24-0047 v02 Heureka results.xlsx"
        result = read_rearranged_heureka_results(filename)
        fpp.plot_raw_data(result['FHF24-0047 v02 Spruce BAU'], 'XXX', 'test',
                          "C:\\Users\\marte\\Downloads")

    def test_read_csv_file(self):
        filename = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0047 Åmli\\FHF24-0047 v02 Averaged stand data.csv"
        data, d = read_csv_file(filename)
        print(data)
        print(d)


class PlotChecks(unittest.TestCase):
    def test_list_duplicates(self):
        print(fpp.list_duplicates(['A', 'B', 'C', 'A']))
        print(fpp.list_duplicates([1, 2, 3, 1]))

    def test_qc_stand_data(self):
        f = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0014 Arne Tag\\Bestandsoversikt 16052024.xlsx"
        table = read_excel(f, 7, 0)
        fpp.qc_stand_data(table, os.path.basename(f), 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-014 Arne Tag\\QC_plots')
        plt.show()

    def test_plot_heureka_results(self):
        result_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0016 Margrete Folsland\\FHF24-0016 Heureka results.xlsx"
        sheets = ["FHF24-0016 Pine PRES", "FHF24-0016 Pine BAU"]
        params = "Total Carbon Stock (dead wood, soil, trees, stumps and roots)"
        diff_sheets = True
        diff = fpp.plot_from_heureka_results(
            result_file,
            sheets=sheets,
            params=params,
            diff_sheets=diff_sheets
        )
        print(diff)
        plt.show()

    def test_plot_collected_stand_data(self):
       csv_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\TEST COLLECT STAND DATA.csv"
       fpp.plot_collected_stand_data(csv_file)


# Moved from the old Fossagrim/unit_tests/test_io.py and test_projects.py
class OldIoChecks(unittest.TestCase):
    def test_active_forests(self):
        new_stand_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\04-Beregning\\FHF24-0027 Kambo\\FHF24-0027-01 Forvaltningsplan.xlsx"
        stand_id_key = 'Fossagrim ID'
        active_forests, active_prod_areal, active_prod_areal_fractions = (
            fio_shim.get_active_forests_from_stand(new_stand_file, stand_id_key=stand_id_key, sheet_name='data'))
        for key in list(active_forests.keys()):
            print(key, active_forests[key])
        self.assertEqual(True, True)  # add assertion here

    def test_decimal_signs(self):
        table = fio_shim.read_excel('TestDecimalSigns.xlsx', 1, 'Sheet1')
        for key in list(table.keys()):
            for value in table[key]:
                print(fio_shim.my_float(value))

    def test_average_over_stands(self):
        forvaltnings_plan_fil = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\03-Modelering\\FHF24-0046 Tokke\\FHF24-0046 Forvaltningsplan.xlsx"
        average_over = {'Spruce': ['1-7', '1-10', '1-13', '1-14']}
        stand_id_key = 'Fossagrim ID'
        average_name = 'TEST'

        table = fio_shim.read_excel(forvaltnings_plan_fil, 7, 'data')

        avg_table = fio_shim.average_over_stands(average_over, table, stand_id_key, average_name, True)

        print(list(avg_table.keys()))


class OldProjectChecks(unittest.TestCase):
    def test_project_settings(self):
        project_name = 'FHF24-0027-01 Kambo'
        project_settings_file = 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings.xlsx'
        project_folder, stand_file, average_over, stand_id_key, result_file, result_sheets, combine_sheets, \
            monetization_file, csv_stand_file, csv_treatment_file, kwargs = \
            projects_new.project_settings(project_name, project_settings_file)
        print(project_folder)
        print(stand_file)
        print(average_over)
        print(stand_id_key)
        print(result_file)
        print(result_sheets)
        print(combine_sheets)
        print(monetization_file)
        print(csv_stand_file)
        print(csv_treatment_file)
        print(kwargs)
        self.assertEqual(True, True)  # add assertion here
