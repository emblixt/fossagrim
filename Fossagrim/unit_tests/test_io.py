import unittest
import Fossagrim.io.fossagrim_io as fio


class MyTestCase(unittest.TestCase):
    def test_active_forests(self):
        new_stand_file = 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\03-Modelering\\FHF24-0027 Kambo\\Skogdata\\FHF24-0027-01 Bestandsutvalg.xlsx'
        active_forests = fio.get_active_forests_from_stand(new_stand_file)
        for key in list(active_forests.keys()):
            print(key, active_forests[key])
        self.assertEqual(True, True)  # add assertion here


if __name__ == '__main__':
    unittest.main()
