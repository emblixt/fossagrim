"""
Quality check the Forvaltningsplan Excel file relative to the pdf files 'Hovedtallsrapport' and 'Sumtallsrapport'

Usage:
    python qc_input_pdfs.py <forvaltningsplan.xlsx> <hovedtallsrapport.pdf> <sumtallsrapport.pdf> <project_name>

Writes, in the folder of the Forvaltningsplan:
    "qc_log <project_name>.txt"    the values that were compared
    "WARNING <project_name>.txt"   only if inconsistencies were found

Checks (the numbers refer to Fossagrim's QC list):
     1. Drivbart areal matches hovedtallsrapport                      check_hovedtallsrapport()
     2. Total timber volume matches hovedtallsrapport                 check_hovedtallsrapport()
     3. No harvest in MIS                                             check_forvaltningsplan()
     4. No harvest in hogstklasse < IV                                check_forvaltningsplan()
     5. Harvest in hogstklasse IV are justified                       check_forvaltningsplan()
     6. Project productive area matches sumtallsrapport               check_sumtallsrapport()
     7. Project stand numbers match sumtallsrapport                   (skipped, the Sumtallsrapport contains
                                                                       non-productive stands)
     8. Project total volume matches sumtallsrapport                  check_sumtallsrapport()
    Also: the selected stands match the sumtallsrapport, and the numeric columns of the Forvaltningsplan do not
    contain numbers written as text (typically a mixture of ',' and '.' as decimal sign)

Not (yet) implemented:
     9. FHF code matches project code
    10. Inventory year matches forest data
    11. Altitude matches location
    12. Latitude matches (Google Earth Pro?)
    13. Forest parameters are within reasonable intervals
    14. Property harvest volumes represent a typical "business-as-usual" scenario

The script can be called from any folder.
"""
import argparse
import os
import sys

import numpy as np

# Make the Fossagrim package importable when this file is run as a script
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Fossagrim.io.common import read_excel  # noqa: E402
from Fossagrim.utils.definitions import FORVALTNINGSPLAN_QC_SHEET, FORVALTNINGSPLAN_QC_HEADER_ROW  # noqa: E402

# Numeric columns of the Forvaltning sheet that must be possible to sum
NUMERIC_COLUMNS = ['Prod.areal', 'Gran', 'Furu', 'Lauv', 'Total', 'Total volum', 'Total hogst', 'Gran hogst',
                   'Furu hogst', 'Lauv hogst', 'Proj areal', 'Proj vol', 'Productive', 'Active', 'Furu.1',
                   'Bjørk', 'Total.1', 'Active.1']


def read_forvaltningsplan(_fplan):
    """The 'Forvaltning' sheet of the Forvaltningsplan. The first data row contains the totals."""
    return read_excel(_fplan, FORVALTNINGSPLAN_QC_HEADER_ROW, FORVALTNINGSPLAN_QC_SHEET)


def read_pdf_lines(pdf_file):
    """Extracts the text of all pages in a pdf file, and returns it as a list of lines"""
    import pypdf

    text = ''
    with open(pdf_file, 'rb') as file:
        reader = pypdf.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text()
    return text.split('\n')


def number_after(lines, label):
    """
    The number on the line after the line that is equal to 'label'. Spaces (thousand separators) are removed
    and a decimal comma is accepted, e.g. ['Tømmervolum', '12 345,6'] -> 12345.6
    Raises ValueError if 'label' is not found.
    """
    i = lines.index(label)
    return float(lines[i + 1].strip().replace(' ', '').replace(',', '.'))


def check_hovedtallsrapport(_fplan: str, _hrapp: str, log_text: str, warn_text: str) -> tuple:
    """
    Compares the productive area and total volume of the Forvaltningsplan with the Hovedtallsrapport (checks 1, 2)

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param _hrapp:
       full pathname of pdf file that contains the 'hovedtallsrapport'
    :param log_text:
       str, the log text so far. The compared values are appended
    :param warn_text:
       str, the warnings so far. Any new warnings are appended
    :return: tuple
        (log_text, warn_text)
    """
    fplan_table = read_forvaltningsplan(_fplan)
    fplan_prod_areal = float(fplan_table['Prod.areal'][0])
    fplan_total_volume = float(fplan_table['Total'][0])

    str_list = read_pdf_lines(_hrapp)
    hrapp_total_volume = number_after(str_list, 'Total kubikkmasse')
    hrapp_prod_areal = number_after(str_list, 'Produktivt skogareal')

    log_text += '  -Productive areal from Forvaltningsplan: {:.2f}\n'.format(fplan_prod_areal)
    log_text += '  -Productive areal from Hovedtallsrappport: {:.2f}\n'.format(hrapp_prod_areal)
    log_text += '  -Total volume from Forvaltningsplan: {:.2f}\n'.format(fplan_total_volume)
    log_text += '  -Total volume from Hovedtallsrappport: {:.2f}\n'.format(hrapp_total_volume)

    # 1. Drivbart areal matches hovedtallsrapport
    if np.floor(fplan_prod_areal) != np.floor(hrapp_prod_areal):
        warn_text += 'WARNING: Productive areal from Forvaltningsplan ({:.2f}) does not match that from ' \
                     'Hovedtallsrapport ({:.2f})\n'.format(fplan_prod_areal, hrapp_prod_areal)

    # 2. Total timber volume matches hovedtallsrapport
    if np.floor(fplan_total_volume) != np.floor(hrapp_total_volume):
        warn_text += 'WARNING: Total volume from Forvaltningsplan ({:.2f}) does not match that from ' \
                     'Hovedtallsrapport ({:.2f})\n'.format(fplan_total_volume, hrapp_total_volume)

    return log_text, warn_text


def _is_summable(values):
    """False if the values contain text, e.g. numbers written with the wrong decimal sign"""
    try:
        np.sum(values)
    except TypeError:
        return False
    return True


def check_forvaltningsplan(_fplan: str, log_text: str, warn_text: str) -> tuple:
    """
    Checks the Forvaltningsplan itself (checks 3, 4, 5, and numbers written as text)

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param log_text:
       str, the log text so far
    :param warn_text:
       str, the warnings so far. Any new warnings are appended
    :return: tuple
        (log_text, warn_text)
    """
    fplan_table = read_forvaltningsplan(_fplan)
    return check_forvaltningsplan_table(fplan_table, log_text, warn_text)


def check_forvaltningsplan_table(fplan_table, log_text: str, warn_text: str) -> tuple:
    """Same as check_forvaltningsplan(), but takes the already read 'Forvaltning' sheet as a DataFrame"""
    # 3. No harvest (or more correctly - no area with MIS included) in MIS
    nr_miljo_fig = 0
    for _i, _ans in enumerate(fplan_table['Miljøfig']):
        if not isinstance(_ans, str):
            continue
        if float(fplan_table['Proj areal'][_i]) > 0.0 and _ans.strip() == 'Ja':
            nr_miljo_fig += 1
            warn_text += 'WARNING: Bestand {} has a mismatch between "misfig" ({}) and inclusion in project ' \
                         'areal ({})\n'.format(fplan_table['GBTBestand'][_i], _ans, fplan_table['Proj areal'][_i])

    log_text += '  -Project contain {} number of mis figs\n'.format(nr_miljo_fig)

    # 4. No harvest in hogstklasse < IV
    nr_hogstklasse_below = 0
    for _i, _ans in enumerate(fplan_table['H.kl']):
        if 4.0 > float(_ans) > 0 and float(fplan_table['Proj vol'][_i]) > 0.0:
            nr_hogstklasse_below += 1
            warn_text += 'WARNING: Bestand {} is cutting in hogstklasse < 4 ({})\n'.format(
                fplan_table['GBTBestand'][_i], _ans)

    log_text += '  -Project contain {} stand in hogstklass < 4\n'.format(nr_hogstklasse_below)

    # 5. Harvest in hogstklasse IV are justified
    nr_hogstklasse_4 = 0
    for _i, _ans in enumerate(fplan_table['H.kl']):
        if float(_ans) == 4.0:
            if float(fplan_table['Proj vol'][_i]) > 0.0 and not isinstance(fplan_table['Begrunnelse'][_i], str):
                nr_hogstklasse_4 += 1
                warn_text += 'WARNING: Bestand {} is cutting in hogstklasse 4 ({}) without "Begrunnelse"\n'.format(
                    fplan_table['GBTBestand'][_i], _ans)

    log_text += '  -Project contain {} stand in hogstklass = 4\n'.format(nr_hogstklasse_4)

    # Check if excel sums are possible to calculate in Python. If not, there are likely a sloppy usage
    # of ',' vs. '.' in the excel file. The first row contains the totals, and is skipped.
    for column in NUMERIC_COLUMNS:
        if column not in fplan_table.keys():
            continue
        data = fplan_table[column]
        if data.ndim == 2:
            # Several columns end up with the same name, e.g. 'Gran\n' and 'Gran' are both 'Gran' after stripping
            for j in range(data.shape[1]):
                if not _is_summable(data.iloc[1:, j].values):
                    warn_text += "WARNING: Column {}:{} in Forvaltningsplan likely contains a mixture of " \
                                 "',' and '.'\n".format(column, j + 1)
        elif not _is_summable(data.values[1:]):
            warn_text += "WARNING: Column {} in Forvaltningsplan likely contains a mixture of ',' and '.'\n".format(
                column)

    return log_text, warn_text


def check_sumtallsrapport(_fplan: str, _srapp: str, log_text: str, warn_text: str) -> tuple:
    """
    Compares the selected stands, productive area and total volume of the project with the
    Sumtallsrapport (checks 6, 8)

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param _srapp:
       full pathname of pdf file that contains the 'sumtallsrapport'
    :param log_text:
       str, the log text so far
    :param warn_text:
       str, the warnings so far. Any new warnings are appended
    :return: tuple
        (log_text, warn_text)
    """
    fplan_table = read_forvaltningsplan(_fplan)
    fplan_prod_areal = 0.
    fplan_total_volume = 0.
    selected_stands = []
    for _i, _proj_area in enumerate(fplan_table['Proj areal']):
        if _i == 0:
            continue  # totals
        if float(_proj_area) > 0:
            selected_stands.append(str(int(fplan_table['Bestand'][_i])))
            fplan_prod_areal += float(fplan_table['Prod.areal'][_i])
            fplan_total_volume += float(fplan_table['Total'][_i])
    fplan_selected_stands = ', '.join(selected_stands)

    fplan_nr_stands = sum(1 for _i, _area in enumerate(fplan_table['Prod.areal']) if _i > 0 and float(_area) > 0)

    str_list = read_pdf_lines(_srapp)
    _i = str_list.index(' Bestand')
    _j = str_list.index('Utvalgte bestand ')
    srapp_selected_stands = ''.join(str_list[_i + 4:_j])
    srapp_total_volume = number_after(str_list, 'Tømmervolum')
    srapp_prod_areal = number_after(str_list, 'Totalt produktivt areal')
    srapp_nr_stands = number_after(str_list, 'Antall bestand')

    log_text += '  -Selected stands from Forvaltningsplan: {}\n'.format(fplan_selected_stands)
    log_text += '  -Selected stands from Sumtallsrapport: {}\n'.format(srapp_selected_stands)
    log_text += '  -Productive areal from Forvaltningsplan: {:.2f}\n'.format(fplan_prod_areal)
    log_text += '  -Productive areal from Sumtallsrappport: {:.2f}\n'.format(srapp_prod_areal)
    log_text += '  -Total volume from Forvaltningsplan: {:.2f}\n'.format(fplan_total_volume)
    log_text += '  -Total volume from Sumtallsrappport: {:.2f}\n'.format(srapp_total_volume)
    log_text += '  -Number of stands from Forvaltningsplan: {}\n'.format(float(fplan_nr_stands))
    log_text += '  -Number of stands from Sumtallsrappport: {}\n'.format(srapp_nr_stands)

    if fplan_selected_stands != srapp_selected_stands:
        warn_text += 'WARNING: The selected stands in Forvaltningsplan\n {}\nare not equal to the selected stands ' \
                     'in Sumtallsrapport\n {}\n'.format(fplan_selected_stands, srapp_selected_stands)

    # 6. Project productive area matches sumtallsrapport
    if np.round(fplan_prod_areal) != np.round(srapp_prod_areal):
        warn_text += 'WARNING: Productive areal from Forvaltningsplan ({}) does not match that from ' \
                     'Sumtallsrapport ({})\n'.format(fplan_prod_areal, srapp_prod_areal)

    # 7. Project stand numbers match sumtallsrapport: skipped, because the Sumtallsrapport contain non-productive
    # stands

    # 8. Project total volume matches sumtallsrapport
    if np.round(fplan_total_volume) != np.round(srapp_total_volume):
        warn_text += 'WARNING: Total volume from Forvaltningsplan ({}) does not match that from ' \
                     'Sumtallsrapport ({})\n'.format(fplan_total_volume, srapp_total_volume)

    return log_text, warn_text


def pdf_consistency(_fplan: str, _hrapp: str, _srapp: str, _project_name: str):
    """
    Runs all checks and writes the log file, and the warning file if any inconsistencies were found
    (see the module doc string). Old log and warning files of the same project are removed first.

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param _hrapp:
       full pathname of pdf file that contains the 'hovedtallsrapport'
    :param _srapp:
       full pathname of pdf file that contains the 'sumtallsrapport'
    :param _project_name:
       text string that identifies this particular project forest, e.g. 'FHF-0001-02 t06_v2'
    :return:
        None
    """
    from datetime import datetime

    _log_file = os.path.join(os.path.dirname(_fplan), 'qc_log {}.txt'.format(_project_name))
    warn_file = os.path.join(os.path.dirname(_fplan), 'WARNING {}.txt'.format(_project_name))
    for old_file in [_log_file, warn_file]:
        if os.path.exists(old_file):
            os.remove(old_file)

    log_text = "QC pdf's vs Forvaltningsplan done {}\n".format(str(datetime.now()))
    log_text += ' Using Forvaltningsplan: {}\n'.format(os.path.basename(_fplan))
    log_text += ' and Hovedtallsrapport: {}\n'.format(os.path.basename(_hrapp))
    log_text += ' and Sumtallsrapport: {}\n'.format(os.path.basename(_srapp))

    warn_text = ''
    log_text, warn_text = check_hovedtallsrapport(_fplan, _hrapp, log_text, warn_text)
    log_text, warn_text = check_forvaltningsplan(_fplan, log_text, warn_text)
    log_text, warn_text = check_sumtallsrapport(_fplan, _srapp, log_text, warn_text)

    with open(_log_file, 'w', encoding="utf-8") as log_file:
        log_file.write(log_text)

    if len(warn_text) > 0:
        with open(warn_file, 'w', encoding="utf-8") as w:
            w.write(warn_text)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description='Quality check the Forvaltningsplan against the Hovedtallsrapport and Sumtallsrapport')
    parser.add_argument("forvaltnings_plan_file", help='Full path name of the Forvaltningsplan Excel file')
    parser.add_argument("hovedtalls_rapport_file", help='Full path name of the Hovedtallsrapport pdf file')
    parser.add_argument("sumtalls_rapport_file", help='Full path name of the Sumtallsrapport pdf file')
    parser.add_argument("project_name", help='Name of the project, used in the names of the log files')
    args = parser.parse_args(argv)
    pdf_consistency(args.forvaltnings_plan_file, args.hovedtalls_rapport_file, args.sumtalls_rapport_file,
                    args.project_name)


if __name__ == '__main__':
    main()
