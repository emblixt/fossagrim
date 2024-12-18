import unittest
import Fossagrim.io.fossagrim_io as fio
import Fossagrim.utils.projects_new as fproj


class MyTestCase(unittest.TestCase):
    def test_project_settings(self):
        project_name = 'FHF24-0027-01 Kambo'
        project_settings_file = 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings.xlsx'
        project_folder, stand_file, average_over, stand_id_key, result_file, result_sheets, combine_sheets, \
            monetization_file, csv_stand_file, csv_treatment_file, kwargs = \
            fproj.project_settings(project_name, project_settings_file)
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


if __name__ == '__main__':
    unittest.main()
