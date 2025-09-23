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
        Creates one file:
            "<_project_name> report consistency QC.txt"  which contains a listing of potential inconsistencies between
            the 'forvaltningsplan' and the two 'hovedtallsrapport' and 'sumtallsrapport'

    """
    from datetime import datetime
    import os
    with open(_log_file, 'w') as log_file:
        log_file.write('QC done {}\n'.format(str(datetime.now())))
        log_file.write(' Using Forvaltningsplan: {}\n'.format(os.path.basename(_fplan)))
        log_file.write(' and Hovedtallsrapport: {}\n'.format(os.path.basename(_hrapp)))
        log_file.write(' and Sumtallsrapport: {}\n'.format(os.path.basename(_srapp)))


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
    > python qc_input.py <forvaltningsplan_file.xlsx> <hovedtallsrapport_file.pdf <log_file.txt>> <sumtallsrapport_file.pdf> <log_file.txt>
        
    """
    parser.add_argument("forvaltnings_plan_file")
    parser.add_argument("hovedtalls_rapport_file")
    parser.add_argument("sumtalls_rapport_file")
    parser.add_argument("log_file")
    args = parser.parse_args()
    pdf_consistency(args.forvaltnings_plan_file, args.hovedtalls_rapport_file, args.sumtalls_rapport_file, args.log_file)


