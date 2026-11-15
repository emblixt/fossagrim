"""
Project workflow driven by the project settings Excel file ("ProjectForestsSettings.xlsx").

Edit the variables under `if __name__ == '__main__':` and run this file:
    1. fix_import = True:   QC the Forvaltningsplan (plots, and pdf consistency checks if the Hovedtallsrapport
                            and Sumtallsrapport are given in the settings file), create the stand register csv
                            file for Heureka and an empty results file
    2. Run the simulations in Heureka and paste the results into the results file
    3. fix_import = False:  rearrange the Heureka results
"""
import os

import Fossagrim.plotting.misc_plots as fpp
import Fossagrim.io.fossagrim_io as fio
from Fossagrim.utils.definitions import FORVALTNINGSPLAN_DATA_SHEET, FORVALTNINGSPLAN_HEADER_ROW
from Fossagrim.utils.project_setup import (
    averaged_stand_csv_file, build_result_sheets, build_combine_sheets, create_empty_results_file)


def arrange_import(_stand_file, _csv_stand_file, _csv_treatment_file, _average_over, _stand_id_key, _project_tag,
                   _hrapp, _srapp,
                   _verbose=False, **_kwargs):
    """
    QC the Forvaltningsplan and create the stand register csv file that is imported into Heureka.
    The arguments are the outputs of project_settings(). _csv_treatment_file and _kwargs are not used.
    """
    from Fossagrim.qc_input_pdfs import pdf_consistency

    # QC input table
    table = fio.read_excel(_stand_file, FORVALTNINGSPLAN_HEADER_ROW, FORVALTNINGSPLAN_DATA_SHEET)
    fpp.qc_stand_data(
        table,
        os.path.basename(_stand_file),
        os.path.join(os.path.dirname(_stand_file), 'QC_plots')
    )

    # QC input
    if _hrapp is not None and _srapp is not None:
        print('Checking the Forvaltningsplan against the Hovedtallsrapport and Sumtallsrapport')
        pdf_consistency(_stand_file, _hrapp, _srapp, _project_tag)

    fio.export_fossagrim_stand_to_heureka(
        _stand_file,
        _csv_stand_file,
        average_over=_average_over,
        stand_id_key=_stand_id_key,
        average_name='{} '.format(_project_tag),
        sheet_name=FORVALTNINGSPLAN_DATA_SHEET,
        verbose=_verbose
    )

    # # The treatment is now added to the csv stand file instead, among the UserDefinedVariables,
    # then we don't need this file
    # fio.export_fossagrim_treatment(
    #     _stand_file,
    #     _csv_treatment_file,
    #     average_over=_average_over,
    #     stand_id_key=_stand_id_key,
    #     average_name='{} '.format(_project_tag),
    #     sheet_name='data'
    # )


def arrange_results(_result_file, _sheet_names, _combine_sheets, _monetization_file, _verbose=False, **_kwargs):
    """
    Rearrange the Heureka results (see heureka_results.rearrange_raw_heureka_results()), and fill in the
    monetization file if one is given.
    """
    write_monetization_file = \
        fio.rearrange_raw_heureka_results(
            _result_file, _sheet_names, _combine_sheets, monetization_file=_monetization_file, verbose=_verbose)

    if write_monetization_file:
        print('Try to modify Monetization file')
        fio.modify_monetization_file(_monetization_file, **_kwargs)


def project_settings(_project_name, _project_settings_file, _fix_import: bool = True, _verbose: bool = False):
    """
    Reads the project settings Excel file, and extracts information from it to be used in the Heureka simulation
    and sets up files necessary for storing results and monetization calculations for the given project

    Keys expected in project settings files:
    "Project name", "Project folder", "Status", "Stand id key", "Stands file", "Results file",

    :param _project_name:
        str
        Name that identifies the current project, e.g. "FHF23-007 Gudmund Aaker"

    :param _project_settings_file:
        str
        Full path name of the project settings Excel sheet.

    :param _fix_import:
        Bool
        If True it goes through the whole process and creates results files,
        If False, it only returns the first few arguments before creating any files and thereby
        avoiding the possibility of overwriting the results file.
    :param _verbose:
        Bool
    :return:
        tuple of
        (project folder, stand file, average_over, stand id key, result file, result sheets, combine sheets,
         monetization file (always None), csv stand file, csv treatment file (no longer created),
         hovedtallsrapport or None, sumtallsrapport or None, kwargs (always empty))
    """
    p_tabl = fio.read_excel(_project_settings_file, 1, 'Settings')

    i = None
    tag_found = False
    _hrapp = None
    _srapp = None
    for i, p_name in enumerate(p_tabl['Project name']):
        if not isinstance(p_name, str):
            continue
        if p_tabl['Status'][i] not in ['Active', 'Prospective']:
            continue
        if _project_name == p_name.strip():
            tag_found = True
            if isinstance(p_tabl['Hovedtallsrapport'][i], str) and isinstance(p_tabl['Sumtallsrapport'][i], str):
                _hrapp = p_tabl['Hovedtallsrapport'][i]
                _srapp = p_tabl['Sumtallsrapport'][i]
            break

    if not tag_found:
        raise IOError('Project name {} not found in {}'.format(
            _project_name, os.path.basename(_project_settings_file)
        ))

    if _verbose:
        if _fix_import:
            print('Preparing import for project {} using projects settings from {}'.format(
                _project_name, _project_settings_file
            ))
            print('Using project folder {},\n with stand ID key {},\n and stand file {},\n '
                  'to create the empty results file {}'.format(
                p_tabl['Project folder'][i],
                p_tabl['Stand id key'][i],
                p_tabl['Stands file'][i],
                p_tabl['Results file'][i]))
            if _hrapp is not None:
                print(" AND Hovedtallsrapport: '{}' and Sumtallsrapport: '{}' to QC the input".format(
                    os.path.basename(_hrapp), os.path.basename(_srapp)))

    _project_folder = p_tabl['Project folder'][i]
    _qc_folder = os.path.join(_project_folder, 'QC_plots')
    _stand_id_key = p_tabl['Stand id key'][i]
    _stand_file = os.path.join(_project_folder, p_tabl['Stands file'][i])
    _result_file = os.path.join(_project_folder, p_tabl['Results file'][i])
    # _monetization_file = os.path.join(_project_folder, p_tabl['Monetization file'][i])  # not calculated in Python
    _csv_stand_file = averaged_stand_csv_file(_project_folder, _project_name)
    _csv_treatment_file = os.path.join(_project_folder, '{} Averaged treatment.csv'.format(_project_name))

    _average_over, _active_prod_areas, _active_prod_areas_fract = fio.get_active_forests_from_stand(
        _stand_file, stand_id_key=_stand_id_key, sheet_name=FORVALTNINGSPLAN_DATA_SHEET)

    # The combine fractions are the fractions of the active productive area of each wood species
    forest_types = list(_active_prod_areas_fract.keys())
    _result_sheets = build_result_sheets(_project_name, forest_types, verbose=_verbose)

    for key, value in _active_prod_areas.items():
        print('{} stands have a total area of {} daa, which corresponds to a fraction of {:.3} of the total area'.format(
            key, sum(value), _active_prod_areas_fract[key]))

    _combine_sheets = build_combine_sheets(
        _project_name, forest_types, list(_active_prod_areas_fract.values()), _result_sheets)

    if not _fix_import:
        return (_project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets,
                _combine_sheets, None, _csv_stand_file, _csv_treatment_file, _hrapp, _srapp,  {})

    # Create empty QC folder if it doesn't exist from before
    if not os.path.exists(_qc_folder):
        os.makedirs(_qc_folder)

    # Create empty results file
    create_results_file = False
    if os.path.isfile(_result_file):
        print("WARNING result file {} already exists.".format(
            os.path.split(_result_file)[-1]))
        _ans = input("Do you want to overwrite? Y/[N]:")
        if _ans.upper() == 'Y':
            create_results_file = True
    else:
        create_results_file = True

    if create_results_file:
        create_empty_results_file(_result_file, _result_sheets)

    return _project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets, _combine_sheets, \
        None, _csv_stand_file, _csv_treatment_file, _hrapp, _srapp, {}


if __name__ == '__main__':
    project_name = 'FHF-0026 t06_v2'
    project_settings_file = 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings.xlsx'

    # Set to False after Heureka simulation results have been saved in result_file, and you want to
    # rearrange the results so that they are easier to include in climate benefit calculations
    fix_import = True

    verbose = True

    project_folder, stand_file, average_over, stand_id_key, result_file, result_sheets, combine_sheets, \
        monetization_file, csv_stand_file, csv_treatment_file, hrapp, srapp, kwargs = \
        project_settings(project_name, project_settings_file, fix_import, verbose)

    if fix_import:
        arrange_import(stand_file, csv_stand_file, csv_treatment_file, average_over, stand_id_key, project_name,
                       hrapp, srapp,
                       _verbose=verbose, **kwargs)
    elif not fix_import:
        arrange_results(result_file, result_sheets, combine_sheets, monetization_file, _verbose=verbose,  **kwargs)

