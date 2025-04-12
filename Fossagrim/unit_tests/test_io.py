import unittest
import Fossagrim.io.fossagrim_io as fio


class MyTestCase(unittest.TestCase):
    def test_active_forests(self):
        new_stand_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\04-Beregning\\FHF24-0027 Kambo\\FHF24-0027-01 Forvaltningsplan.xlsx"
        stand_id_key = 'Fossagrim ID'
        active_forests, active_prod_areal, active_prod_areal_fractions = (
            fio.get_active_forests_from_stand(new_stand_file, stand_id_key=stand_id_key, sheet_name='data'))
        for key in list(active_forests.keys()):
            print(key, active_forests[key])
        self.assertEqual(True, True)  # add assertion here

    def test_decimal_signs(self):
        table = fio.read_excel('TestDecimalSigns.xlsx', 1, 'Sheet1')
        for key in list(table.keys()):
            for value in table[key]:
                print(fio.my_float(value))

    def test_average_over_stands(self):
        forvaltnings_plan_fil = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\03-Modelering\\FHF24-0046 Tokke\\FHF24-0046 Forvaltningsplan.xlsx"
        average_over = {'Spruce': ['1-7', '1-10', '1-13', '1-14']}
        stand_id_key = 'Fossagrim ID'
        average_name = 'TEST'

        table = fio.read_excel(forvaltnings_plan_fil, 7, 'data')

        avg_table = fio.average_over_stands(average_over, table, stand_id_key, average_name, True)

        print(list(avg_table.keys()))


if __name__ == '__main__':
    unittest.main()
