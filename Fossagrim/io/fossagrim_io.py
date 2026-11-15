"""
Backwards compatible entry point to the Fossagrim io functions.

The functions used to live in this one (very long) module, and have been split into:
    Fossagrim/io/common.py           Excel and Heureka csv file helpers
    Fossagrim/io/stands.py           Forvaltningsplan -> Heureka stand register (averaging, unit conversion)
    Fossagrim/io/heureka_results.py  Reading, rearranging and combining Heureka results
    Fossagrim/io/monetization.py     The monetization Excel file

Existing code that does
    import Fossagrim.io.fossagrim_io as fio
    fio.read_excel(...)
keeps working. Functions that are no longer in use have been moved to Fossagrim/legacy/legacy_io.py.
"""
from Fossagrim.io.common import (  # noqa: F401
    my_float, my_str, my_str_to_float, get_row_index, read_excel, write_csv_file, read_csv_file)
from Fossagrim.io.stands import (  # noqa: F401
    arrange_export_of_one_stand, average_over_stands, matched_average_groups, export_fossagrim_stand_to_heureka,
    get_active_forests_from_stand)
from Fossagrim.io.heureka_results import (  # noqa: F401
    read_raw_heureka_results, combine_raw_heureka_results, combine_tables, rearrange_raw_heureka_results,
    read_rearranged_heureka_results, get_nature_and_climate_effect, collect_all_stand_data)
from Fossagrim.io.monetization import (  # noqa: F401
    get_kwargs_from_stand, modify_monetization_file, style_monetization_file, qc_plots)
