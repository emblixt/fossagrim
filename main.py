import os
import Fossagrim.io.fossagrim_io as fio


def arrange_tags_and_files(_project_tag, _project_folder):
    _bau_tag = '{} Business as usual'.format(_project_tag)
    _preserv_tag = '{} Preservation'.format(_project_tag)

    _stand_file = os.path.join(_project_folder, '{} Bestandsutvalg.xlsx'.format(_project_tag))
    _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))

    _csv_stand_file = os.path.join(_project_folder, '{} Averaged stand data.csv'.format(_project_tag))
    _csv_treatment_file = os.path.join(_project_folder, '{} Averaged treatment.csv'.format(_project_tag))

    return _bau_tag, _preserv_tag, _stand_file, _result_file, _csv_stand_file, _csv_treatment_file


def arrange_import(_stand_file, _csv_stand_file, _csv_treatment_file, _average_over, _stand_id_key):
    fio.export_fossagrim_stand_to_heureka(
        _stand_file,
        _csv_stand_file,
        average_over=_average_over,
        stand_id_key=_stand_id_key
    )

    fio.export_fossagrim_treatment(
        _stand_file,
        _csv_treatment_file,
        average_over=_average_over,
        stand_id_key = _stand_id_key
    )


def arrange_results(_result_file, _bau_tag, _preserv_tag):
    fio.rearrange_raw_heureka_results(_result_file, [_bau_tag, _preserv_tag])


def project_settings(_project_tag):
    if _project_tag == 'FHF23-003':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra"
        _average_over = \
            {'Spruce': [67, 68, 70, 72, 75, 79, 80, 85, 88, 90, 92, 93, 95, 97, 98, 99, 104, 105, 109]}
        _stand_id_key = 'Bestand'

    elif _project_tag == 'FHF23-005':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-005 Kvistaul"
        _average_over = {
            'Spruce': ['FHF23-005-{}'.format(_x) for _x in
                       ['1', '2', '3A', '3B', '4', '11', '13', '14', '15', '19', '28',
                        '30', '31', '32', '33', '37', '38', '39', '41', '23']],
            'Pine': ['FHF23-005-{}'.format(_x) for _x in [18, 26, 27, 29]]
        }
        _stand_id_key = 'Fossagrim ID'

    else:
        raise IOError('Project {} is not known'.format(_project_tag))

    return _project_folder, _average_over, _stand_id_key


if __name__ == '__main__':
    project_tag = 'FHF23-005'
    fix_import = True  # If False, it arranges the Heureka results instead (Heureka must be run first)

    project_folder, average_over, stand_id_key = project_settings(project_tag)
    print(stand_id_key)

    bau_tag, preserv_tag, stand_file, result_file, csv_stand_file, csv_treatment_file = \
        arrange_tags_and_files(project_tag, project_folder)

    if fix_import:
        arrange_import(stand_file, csv_stand_file, csv_treatment_file, average_over, stand_id_key)
    else:
        arrange_results(result_file, bau_tag, preserv_tag)

