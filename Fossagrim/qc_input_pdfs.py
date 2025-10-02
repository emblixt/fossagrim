"""
Quality check the forvaltningsplan excel file relative to the pdf files 'hovedtallsrapport' and 'sumtallsrapport'

 1. Drivbart areal matches hovedtallsrapport

 2. Total timber volume matches hovedtallsrapport

 3. No harvest in MIS

 4. No harvest in hogstklasse < IV

 5. Harvest in hogstklasse IV are justified

 6. Project productive area matches sumtallsrapport

 7. Project stand numbers match sumtallsrapport

 8. Project total volume matches sumtallsrapport

 9. FHF code matches project code

 10. Inventory year matches forest data (fra hvor?)

 11. Altitude matches location

 12. Latitude matches (Google Earth Pro?)

 13. Forest parameters are withing reasonable intervals

 14. Property harvest volumes represent a typical “business-as-usual” scenario


 So we need to read the following files:
    1. forvaltningsplan
    2. hovedtallsrapport
    3. sumtallsrapport

We can access a pdf file using pydf
with open(f, 'rb') as file:
    reader = pypdf.PdfReader(file)
    text = ""
    for i in range(reader.get_num_pages()):
        page = reader.get_page(i)
        text += page.extract_text()
"""
import argparse
import sys


def check_hovedtallsrapport(_fplan: str, _hrapp: str, log_text: str, warn_text: str) -> tuple:
    """

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param _hrapp:
       full pathname of pdf file that contains the 'hovedtallsrapport'
    :param warn_text:
       Either string or None.
       If a string, it contains a warning message that needs to be reported
    :return: tuple
        Tuple of information text and warning text
    """
    import pypdf
    import numpy as np
    sys.path.append('..')

    from Fossagrim.io.fossagrim_io import read_excel

    fplan_table = read_excel(_fplan, 6, 'Forvaltning')
    fplan_prod_areal = fplan_table['Prod.areal'][0]
    fplan_total_volume = fplan_table['Total'][0]

    _str = ''
    with open(_hrapp, 'rb') as file:
        reader = pypdf.PdfReader(file)
        for i in range(reader.get_num_pages()):
            page = reader.get_page(i)
            _str += page.extract_text()

    str_list = _str.split('\n')
    _i = str_list.index('Total kubikkmasse')
    hrapp_total_volume = float(str_list[_i+1].strip().replace(' ', ''))

    _i = str_list.index('Produktivt skogareal')
    hrapp_prod_areal = float(str_list[_i+1].strip().replace(' ', ''))

    log_text += '  -Productive areal from Forvaltningsplan: {:.2f}\n'.format(float(fplan_prod_areal))
    log_text += '  -Productive areal from Hovedtallsrappport: {:.2f}\n'.format(float(hrapp_prod_areal))
    log_text += '  -Total volume from Forvaltningsplan: {:.2f}\n'.format(float(fplan_total_volume))
    log_text += '  -Total volume from Hovedtallsrappport: {:.2f}\n'.format(float(hrapp_total_volume))

    _warn_text = ''
    # 1. Drivbart areal matches hovedtallsrapport
    if np.floor(float(fplan_prod_areal)) != np.floor(float(hrapp_prod_areal)):
        _warn_text += 'WARNING: Productive areal from Forvaltningsplan ({:.2}) does not match that from Hovedtallsrapport ({:.2f})\n'.format(float(fplan_prod_areal), float(hrapp_prod_areal))

    # 2. Total timber volume matches hovedtallsrapport
    if np.floor(float(fplan_total_volume)) != np.floor(float(hrapp_total_volume)):
        _warn_text += 'WARNING: Total volume from Forvaltningsplan ({:.2}) does not match that from Hovedtallsrapport ({:.2f})\n'.format(float(fplan_total_volume), float(hrapp_total_volume))

    if len(_warn_text) > 0:
        warn_text += _warn_text

    return log_text, warn_text


def check_forvaltningsplan(_fplan: str, log_text: str,  warn_text: str) -> tuple:
    """

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param warn_text:
       Either string or None.
       If a string, it contains a warning message that needs to be reported

    :return: tuple
        Tuple of information text and warning text
    """
    import numpy as np
    sys.path.append('..')

    from Fossagrim.io.fossagrim_io import read_excel


    fplan_table = read_excel(_fplan, 6, 'Forvaltning')
    _warn_text = ''

    # 3. No harvest in MIS nr_miljo_fig = 0
    nr_miljo_fig = 0
    for _i, _ans in enumerate(fplan_table['Miljøfig']):
        if _ans == 'Ja':
            nr_miljo_fig += 1
            if float(fplan_table['Proj vol'][_i]) > 0.0:
                _warn_text += 'WARNING: Bestand {} has a mismatch between "misfig" ({}) and cutting ({})'.format(
                    fplan_table['GBTBestand'][_i], _ans, fplan_table['Total hogst%'][_i])

    log_text += '  -Project contain {} number of mis figs\n'.format(nr_miljo_fig)

    # 4. No harvest in hogstklasse < IV
    nr_hogstklasse_below = 0
    for _i, _ans in enumerate(fplan_table['H.kl']):
        if 4.0 > float(_ans) > 0:
            nr_hogstklasse_below += 1
            if float(fplan_table['Proj vol'][_i]) > 0.0:
                _warn_text += ('WARNING: Bestand {} is cutting in hogstklasse < 4 ({})\n').format(
                    fplan_table['GBTBestand'][_i], _ans)

    log_text += '  -Project contain {} stand in hogstklass < 4\n'.format(nr_hogstklasse_below)

    # 5. Harvest in hogstklasse IV are justified
    nr_hogstklasse_4 = 0
    for _i, _ans in enumerate(fplan_table['H.kl']):
        if float(_ans) == 4.0:
            nr_hogstklasse_4 += 1

            if float(fplan_table['Active'][_i]) != 0.0 and not isinstance(fplan_table['Begrunnelse'][_i], str):
                _warn_text += ('Bestand {} is cutting in hogstklasse 4 ({}) without "Begrunnelse"\n').format(
                    fplan_table['GBTBestand'][_i], _ans)

    log_text += '  -Project contain {} stand in hogstklass = 4\n'.format(nr_hogstklasse_4)

    # Check if excel sums are possible to calculate in Python. If not, there are likely a sloppy usage
    # of ',' vs. '.' in the excel file
    for column in ['Prod.areal', 'Gran', 'Furu', 'Lauv', 'Total', 'Total volum', 'Total hogst', 'Gran hogst',
                   'Furu hogst', 'Lauv hogst', 'Proj areal', 'Proj vol', 'Productive', 'Active', 'Furu.1',
                   'Bjørk', 'Total.1', 'Active.1']:
        if column == 'Gran':  # There are two columns named 'Gran'
            _continue = False
            try:
                x_ = np.sum(fplan_table['Gran'].values[1:, 0])
            except TypeError as _e:
                _warn_text += "WARNING: Column {}:1 in Forvaltningsplan likely contains a mixture of ',' and '.'\n".format(
                    column)
                _continue = True
            try:
                x_ = np.sum(fplan_table['Gran'].values[0, 1])
            except TypeError as _e:
                _warn_text += "WARNING: Column {}:2 in Forvaltningsplan likely contains a mixture of ',' and '.'\n".format(
                    column)
                _continue = True
            if _continue:
                continue
        else:
            try:
               _x = np.sum(fplan_table[column][1:])
            except TypeError as _e:
                _warn_text += "WARNING: Column {} in Forvaltningsplan likely contains a mixture of ',' and '.'\n".format(column)
                continue

    if len(_warn_text) > 0:
        warn_text += _warn_text

    return log_text, warn_text


def check_sumtallsrapport(_fplan: str, _srapp: str, log_text: str, warn_text: str) -> tuple:
    """

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param _srapp:
       full pathname of pdf file that contains the 'sumtallsrapport'
    :param warn_text:
       Either string or None.
       If a string, it contains a warning message that needs to be reported
    :return: tuple
        Tuple of information text and warning text
    """
    import pypdf
    import numpy as np
    sys.path.append('..')

    from Fossagrim.io.fossagrim_io import read_excel

    fplan_table = read_excel(_fplan, 6, 'Forvaltning')
    fplan_prod_areal = 0.
    fplan_total_volume = 0.
    fplan_nr_stands = 0
    fplan_selected_stands = ''
    for _i, _proj_vol in enumerate(fplan_table['Proj vol']):
        if _i == 0:
            continue
        if float(_proj_vol) > 0:
            # print(_i, _proj_vol, fplan_table['Prod.areal'][_i], fplan_table['Total'][_i])
            fplan_nr_stands += 1
            fplan_selected_stands += ', {}'.format(int(fplan_table['Bestand'][_i]))
            fplan_prod_areal += float(fplan_table['Prod.areal'][_i])
            fplan_total_volume += float(fplan_table['Total'][_i])
    fplan_selected_stands = fplan_selected_stands[2:]

    _str = ''
    with open(_srapp, 'rb') as file:
        reader = pypdf.PdfReader(file)
        for i in range(reader.get_num_pages()):
            page = reader.get_page(i)
            _str += page.extract_text()

    str_list = _str.split('\n')

    _i = str_list.index(' Bestand')
    _j = str_list.index('Utvalgte bestand ')
    srapp_selected_stands = ''.join(str_list[_i+4:_j])

    _i = str_list.index('Tømmervolum')
    srapp_total_volume = float(str_list[_i+1].strip().replace(' ', '').replace(',','.'))

    _i = str_list.index('Totalt produktivt areal')
    srapp_prod_areal = float(str_list[_i+1].strip().replace(' ', '').replace(',','.'))

    _i = str_list.index('Antall bestand')
    srapp_nr_stands = float(str_list[_i+1].strip().replace(' ', '').replace(',','.'))

    log_text += '  -Selected stands from Forvaltningsplan: {}\n'.format(fplan_selected_stands)
    log_text += '  -Selected stands from Sumtallsrapport: {}\n'.format(srapp_selected_stands)
    log_text += '  -Productive areal from Forvaltningsplan: {:.2f}\n'.format(float(fplan_prod_areal))
    log_text += '  -Productive areal from Sumtallsrappport: {:.2f}\n'.format(float(srapp_prod_areal))
    log_text += '  -Total volume from Forvaltningsplan: {:.2f}\n'.format(float(fplan_total_volume))
    log_text += '  -Total volume from Sumtallsrappport: {:.2f}\n'.format(float(srapp_total_volume))
    log_text += '  -Number of stands from Forvaltningsplan: {}\n'.format(float(fplan_nr_stands))
    log_text += '  -Number of stands from Sumtallsrappport: {}\n'.format(float(srapp_nr_stands))

    _warn_text = ''

    # Test which stands that are selected:
    if fplan_selected_stands != srapp_selected_stands:
        _warn_text += 'WARNING: The selected stands in Forvaltningsplan\n {}\nare not equal to the selected stands in Sumtallsrapport\n {}\n'.format(fplan_selected_stands, srapp_selected_stands)

    # 6. Project productive area matches sumtallsrapport
    if np.floor(float(fplan_prod_areal)) != np.floor(float(srapp_prod_areal)):
        _warn_text += 'WARNING: Productive areal from Forvaltningsplan ({}) does not match that from Sumtallsrapport ({})\n'.format(float(fplan_prod_areal), float(srapp_prod_areal))

    # 7. Project stand numbers match sumtallsrapport
    if fplan_nr_stands != srapp_nr_stands:
        _warn_text += 'WARNING: Number of stands from Forvaltningsplan ({}) does not match that from Sumtallsrapport ({})\n'.format(float(fplan_nr_stands), float(srapp_nr_stands))

    # 8. Project total volume matches sumtallsrapport
    if np.floor(float(fplan_total_volume)) != np.floor(float(srapp_total_volume)):
        _warn_text += 'WARNING: Total volume from Forvaltningsplan ({}) does not match that from Sumtallsrapport ({})\n'.format(float(fplan_total_volume), float(srapp_total_volume))

    if len(_warn_text) > 0:
        warn_text += _warn_text

    return log_text, warn_text


def pdf_consistency(_fplan: str, _hrapp: str, _srapp: str, _log_file: str):
    """

    :param _fplan:
       full pathname of excel file that contains the 'forvaltningsplan'
    :param _hrapp:
       full pathname of pdf file that contains the 'hovedtallsrapport'
    :param _srapp:
       full pathname of pdf file that contains the 'sumtallsrapport'
    :param _log_file:
       full pathname of test file that contains the results of the qc

    :return:
        None
        Creates one, potentially two, file(s):
            _log_file  which contains a listing of potential inconsistencies between
            the 'forvaltningsplan' and the two 'hovedtallsrapport' and 'sumtallsrapport'
        AND OPTIONALLY
            a "warning" file in the same directory as _log_file, which flags potential inconsistencies

    """
    from datetime import datetime
    import os

    warn_file = os.path.join(os.path.dirname(_log_file), 'WARNING.txt')
    if os.path.exists(warn_file):
        os.remove(warn_file)

    log_text = "QC pdf's vs Forvaltningsplan done {}\n".format(str(datetime.now()))
    log_text += ' Using Forvaltningsplan: {}\n'.format(os.path.basename(_fplan))
    log_text += ' and Hovedtallsrapport: {}\n'.format(os.path.basename(_hrapp))
    log_text += ' and Sumtallsrapport: {}\n'.format(os.path.basename(_srapp))

    warn_text = ''

    log_text, warn_text = check_hovedtallsrapport(_fplan, _hrapp, log_text, warn_text)

    log_text, warn_text = check_forvaltningsplan(_fplan, log_text, warn_text)

    log_text, warn_text = check_sumtallsrapport(_fplan, _srapp, log_text, warn_text)

    with open(_log_file, 'w',  encoding="utf-8") as log_file:
        log_file.write(log_text)

    if len(warn_text) > 0:
        with open(warn_file, 'w') as w:
            w.write(warn_text)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    """
    :param forvaltnings_plan_file:
        str
        Full path name of the excel file that contains the forvaltningsplan
        
    :param hovedtalls_rapport_file:
       full pathname of pdf file that contains the 'hovedtallsrapport'
       
    :param sumtalls_rapport_file:
       full pathname of pdf file that contains the 'sumtallsrapport'
       
    :param log_file:
       full pathname of text file that contains the results of the quality control
       
    Use it by calling:
    > python qc_input_pdfs.py <forvaltningsplan_file.xlsx> <hovedtallsrapport_file.pdf <log_file.txt>> <sumtallsrapport_file.pdf> <log_file.txt>
        
    """
    parser.add_argument("forvaltnings_plan_file")
    parser.add_argument("hovedtalls_rapport_file")
    parser.add_argument("sumtalls_rapport_file")
    parser.add_argument("log_file")
    args = parser.parse_args()
    pdf_consistency(args.forvaltnings_plan_file, args.hovedtalls_rapport_file, args.sumtalls_rapport_file, args.log_file)


