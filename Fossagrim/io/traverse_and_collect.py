"""
These scripts are meant to traverse the set of input and output files in Fossagrim and collect interesting
nature and climate parameters that then can be plotted

It should replace some older functions in fossagrim_io.py and misc_plots.py
"""
import matplotlib.pyplot as plt
import openpyxl
import pandas as pd
import numpy as np
import os
import unittest

default_parameters = ['Total CO2 stock', 'Volume large deadwood', 'Mean age (before)',
                      'Mean age (after)', 'Recreation index']
default_descs = ['Ton C02/ha, ignoring soil', 'm3/ha lying and standing deadwood above 20 cm diam.',
                 'Basal area weighted mean age before project', 'Basal area weighted mean age after project',
                 'Recreation index: https://www.heurekaslu.se/wiki/Recreation_Value_Model']


def collect_data(parameters: list, descs: list):
    pass


class TraverseDirectory:
    """
    An iterator function that should return the fundamental input and output functions of Fossagrims Heureka modeling

    :param base_dir:
    :return:
    """
    def __init__(self,
                 base_dir: str,
                 heureka_input='Averaged stand data.csv',
                 heureka_output='Heureka results.xlsx',
                 verbose=False):
        from pathlib import Path
        self.base_dir = base_dir
        self.heureka_input = heureka_input
        self.heureka_output = heureka_output
        self.verbose = verbose
        self.file_list = Path(base_dir).rglob('*' + heureka_input)

    def __iter__(self):
        return self

    def __next__(self):
        this_input_file = next(self.file_list)
        this_output_file = str(this_input_file).replace(self.heureka_input, self.heureka_output)
        if os.path.isfile(this_output_file):
            if self.verbose:
                print(' Heureka result file exists')
        else:
            return None, None
        return this_input_file, this_output_file


class TestCases(unittest.TestCase):
    def test_traverse(self):
        base_dir = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger"
        files = TraverseDirectory(base_dir)
        for file1, file2 in iter(files):
            print(file1, file2)

