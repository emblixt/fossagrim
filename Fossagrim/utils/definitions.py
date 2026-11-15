"""
Shared constants for the Fossagrim <-> Heureka workflow.

Contents
--------
* Column names and layout of the Fossagrim "Forvaltningsplan" (a.k.a. "Bestandsutvalg") Excel file
* Column names (and descriptions) of the Heureka stand register import file ("StandData.csv"), see
  https://www.heurekaslu.se/wiki/Import_of_stand_register
* Rules for how each Forvaltningsplan column is treated when stands are averaged
* Colors, line styles and line widths used in plots and in the monetization Excel file
"""

# Unit conversion: cubic meter solid volume under bark (m3fub) -> cubic meter standing volume over bark (m3sk)
m3fub_to_m3sk = 1.19

# ---------------------------------------------------------------------------------------------------------------
# Layout of the Fossagrim "Forvaltningsplan" Excel file
# ---------------------------------------------------------------------------------------------------------------
FORVALTNINGSPLAN_DATA_SHEET = 'data'  # Sheet with one row per stand
FORVALTNINGSPLAN_HEADER_ROW = 7  # Zero-based row index of the column headers in the data sheet
FORVALTNINGSPLAN_QC_SHEET = 'Forvaltning'  # Sheet used by the pdf consistency checks (qc_input_pdfs.py)
FORVALTNINGSPLAN_QC_HEADER_ROW = 6
DEFAULT_STAND_ID_KEY = 'Fossagrim ID'  # Column that uniquely identifies a stand
SPECIES_KEY = 'Bonitering\ntreslag'  # Column with the dominant tree species of a stand
UNPRODUCTIVE_SPECIES = ['Uproduktiv', '-']  # Values in SPECIES_KEY that mark stands which are never exported
SPECIES_NAMES = {
    # English species name used in file/sheet names: values in SPECIES_KEY that map to it
    'Spruce': ['Gran', 'G', 'Spruce'],
    'Pine': ['Furu', 'F', 'Pine'],
    'Birch': ['Bjørk', 'B', 'Birch'],
}
HEUREKA_SITE_INDEX_SPECIES = {'Gran': 'G', 'Furu': 'T', 'Bjørk': 'B'}  # Norwegian name -> Heureka species code

# ---------------------------------------------------------------------------------------------------------------
# Heureka simulations and result files
# ---------------------------------------------------------------------------------------------------------------
METHODS = ['BAU', 'PRES']  # "Business as usual" and "Preservation"
HEUREKA_PERIOD_LENGTH = 5  # Years per Heureka simulation period
REARRANGED_RESULTS_SHEET = 'Rearranged results'
# Encoding of the semicolon separated csv files that are imported into Heureka. Heureka is a Windows program,
# and these files have always been written with the Windows default (cp1252) encoding.
HEUREKA_CSV_ENCODING = 'cp1252'
EXCEL_MAX_SHEET_NAME_LENGTH = 31

standard_colors = {
    'time_1': '#D9E1F2',
    'products_0': '#A5A5A5',
    'products_1': '#D1D1D1',
    'products_2': '#E8E8E8',
    'substitution_0': '#ED7D31',
    'substitution_1': '#F7C5A3',
    'substitution_2': '#FCE7D8',
    'bau_0': '#FFC000',
    'bau_1': '#FFD85B',
    'bau_2': '#FFEFBD',
    'project_case_0': '#70AD47',
    'project_case_1': '#BBDAA6',
    'project_case_2': '#E0EED6',
    'climate_benefit_0': '#5B9BD5',
    'climate_benefit_1': '#A9CBE9',
    'climate_benefit_2': '#E1EDF7',
    'table_1': '#FCE4D6',
    'cbo_1': '#FFF2CC',
    'buffer_1': '#9BC2E6'
}

standard_linestyles = {
    'time_1': '-',
    'products_0': '--',
    'products_1': '--',
    'products_2': '--',
    'substitution_0': ':',
    'substitution_1': ':',
    'substitution_2': ':',
    'bau_0': '-',
    'bau_1': '-',
    'bau_2': '-',
    'project_case_0': '-.',
    'project_case_1': '-.',
    'project_case_2': '-.',
    'climate_benefit_0': '-',
    'climate_benefit_1': '-',
    'climate_benefit_2': '-',
    'table_1': '-',
    'cbo_1': '-',
    'buffer_1': '-'
}

standard_linewidths = {
    'time_1': 1,
    'products_0': 2,
    'products_1': 1,
    'products_2': 0.5,
    'substitution_0': 2,
    'substitution_1': 1,
    'substitution_2': 0.5,
    'bau_0': 2,
    'bau_1': 1,
    'bau_2': 0.5,
    'project_case_0': 2,
    'project_case_1': 1,
    'project_case_2': 0.5,
    'climate_benefit_0': 2,
    'climate_benefit_1': 1,
    'climate_benefit_2': 0.5,
    'table_1': 1,
    'cbo_1': 1,
    'buffer_1': 1
}

heureka_mandatory_standdata_keys = [
    'StandId',
    'TotalArea',
    'ProdArea',
    'CountyCode',
    'Altitude',
    'Latitude',
    'SiteIndexSpecies',
    'SIS',
    'SoilMoistureCode',
    'VegetationType',
    'Peat',
    'InventoryYear',
    'DGV',
    'DG',
    'H',
    'MeanAge',
    'N',
    'G',
    'V',
    'PropPine',
    'PropSpruce',
    'PropBirch',
    'PropAspen',
    'PropOak',
    'PropBeech',
    'PropSouthernBroadleaf',
    'PropContorta',
    'PropOtherBroadleaf'
]

heureka_standdata_keys = heureka_mandatory_standdata_keys + [
    'AreaLevel2',
    'AreaLevel3',
    'ParentStandId',
    'ImpArea',
    'NCArea',
    'Layer',
    'Register',
    'CoordEast',
    'CoordNorth',
    'DistanceToCoast',
    'RoadId1',
    'RoadId2',
    'ClimateCode',
    'SKSManagementClass',
    'MaturityClass',
    'BlueTargetClass',
    'SIH',
    'SI_Management',
    'BottomLayer',
    'SoilDepth',
    'SoilWater',
    'Texture',
    'SlopeDirectionNorthEast',
    'Ditch',
    'SoilBearingCapacity',
    'Surface',
    'SlopeType',
    'TerrRoadSlope',
    'OwnerType',
    'EvenAgedCode',
    'CAI',
    'DiameterType',
    'DGPine',
    'DGSpruce',
    'DGBirch',
    'DGAspen',
    'DGOak',
    'DGBeech',
    'DGSouthernBroadleaf',
    'DGContorta',
    'DGOtherBroadleaf',
    'HPine',
    'HSpruce',
    'HBirch',
    'HAspen',
    'HOak',
    'HBeech',
    'HSouthernBroadleaf',
    'HContorta',
    'HOtherBroadleaf',
    'SpeciesUser',
    'ProportionSpeciesUser',
    'MeanDiameterSpeciesUser',
    'MeanHeightSpeciesUser',
    'DeadWoodTotal',
    'DeadWoodDecayClass1',
    'DeadWoodDecayClass2',
    'DeadWoodDecayClass3',
    'DeadWoodDecayClass4',
    'DeadWoodDecayClass5',
    'TerrainTransportDistance',
    'LastClearcutYear',
    'LastThinningYear',
    'LastFertilizationYear',
    'LastRegenerationYear',
    'RegenerationMethod',
    'RegenerationSpecies',
    'RegenerationBreeded',
    'Note',
    'UserDefinedVariable1_TaxType',
    'UserDefinedVariable2_RotationPeriod',
    'UserDefinedVariable3_ThinningYear',
    # 'UserDefinedVariable4_PlantDensity',
    'PlantDensity',
    'UserDefinedVariable5_ClimateEffect',
    'UserDefinedVariable6_DeadWoodEffect',
    'UserDefinedVariable7_RecreationEffect',
    'UniqueID',
    'UserDefinedVariable9',
    'UserDefinedVariable10',
    'SetAsideType'
]

heureka_standdata_desc = [
    '','','','','','','','','','','','',
    'Grundytevägd medeldiameter (cm)',
    'Grundytemedelstammens diameter (cm)',
    'Medelhöjd (grundytevägd alt. huvudstammarnas aritmetiska medelhöjd i ungskog)',
    'Medelålder (grundytevägd totalålder. år)',
    'Stamantal (träd per ha. produktiv areal)',
    'Grundyta (m2/ha. produktiv areal)',
    'Volym (m3sk/ha. produktiv areal)',
    '','','','','','','','','','','','','','','','','','','','','','','','','','', '','','','',
    '','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','',
    '','','','','','','','','','','','År','År','Planter/ha',
    '30 yrs sum of diff. in Total carbon (PRES-BAU), ignoring Soil. ton C/ha',
    '30 yrs sum of diff. in Deadwood >= 20cm (PRES-BAU), m3/h',
    '30 yrs sum of diff. in Recreation Index After (PRES-BAU)','Unique version string','','',''
]

heureka_treatment_keys = [
    # https://www.heurekaslu.se/wiki/Import_of_stand_register
    # https://www.heurekaslu.se/help/index.htm?importera_atgardsforslag.htm
    'AreaLevel2',
    'AreaLevel3',
    'StandId',
    'Treatment',
    'Year',
    'Species',
    'ThinningForm',
    'ThinningGrade',
    'PlantDensity',
    'Note'
]

heureka_treatment_desc = [
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    'Plants/ha',
    'Description'
]

fossagrim_standdata_keys = [
    'HovedNr',
    'Gnr',
    'Bnr',
    'Teig',
    'Bestand',
    'Miljøfig',
    'Bonitering\ntreslag',
    'Markslag',
    'H.kl',
    'Tetthet',
    'År',
    'Prod.areal',
    'Gran',
    'Furu',
    'Lauv',
    'Total',
    'Vol\n/\ndaa',
    'Tilvekst\n/\ndaa',
    'Fossagrim ID',
    'InventoryYear',
    'CountyCode',
    'Altitude',
    'Latitude',
    'SoilMoistureCode',
    'VegetationType',
    'Peat',
    'Antall trær pr mål',
    'Gjennomsnitts diameter',
    'Middelhøyde',
    'Grunnflatesum',
    'Tretype_Alder_høyde',
    'Svensk bonitet',
    'Plantetetthet',
    'Rotasjonsperiode',
    'Tynnings år'
]

fossagrim_keys_to_average_over = [
    'År',
    'Vol\n/\ndaa',
    'Tilvekst\n/\ndaa',
    'Altitude',
    'Latitude',
    'Antall trær pr mål',
    'Gjennomsnitts diameter',
    'Middelhøyde',
    'Grunnflatesum',
    'Svensk bonitet',
    'Plantetetthet',
    'Rotasjonsperiode',
    'Tynnings år'
]

fossagrim_keys_to_sum_over = [
    'Prod.areal',
    'Gran',
    'Furu',
    'Lauv',
    'Total'
]

fossagrim_keys_to_most_of = [
    'Markslag',
    'SoilMoistureCode',
    'VegetationType',
    'Peat'
]


def translate_keys_from_fossagrim_to_heureka():
    translation = {}
    for key in fossagrim_standdata_keys:
        if key == 'Bonitering\ntreslag':
            translation['SiteIndexSpecies'] = key
        if key == 'År':
            translation['MeanAge'] = key
        if key == 'Prod.areal':
            translation['ProdArea'] = key
        if key == 'Vol\n/\ndaa':
            translation['V'] = key
        if key in ['InventoryYear', 'CountyCode', 'Altitude', 'Latitude', 'SoilMoistureCode', 'VegetationType', 'Peat']:
            translation[key] = key
        if key == 'Antall trær pr mål':
            translation['N'] = key
        if key == 'Gjennomsnitts diameter':
            translation['DGV'] = key
        if key == 'Middelhøyde':
            translation['H'] = key
        if key == 'Grunnflatesum':
            translation['G'] = key
        if key == 'Svensk bonitet':
            translation['SIS'] = key
        if key == 'Plantetetthet':
            translation['PlantDensity'] = key
        if key == 'Rotasjonsperiode':
            translation['UserDefinedVariable2_RotationPeriod'] = key
        if key == 'Tynnings år':
            translation['UserDefinedVariable3_ThinningYear'] = key
    return translation


def validate_key_definitions():
    """
    Sanity check of the key lists in this module. Raises ValueError if they are inconsistent.
    Called from the unit tests.
    """
    if len(heureka_standdata_keys) != len(heureka_standdata_desc):
        raise ValueError('Length of heureka keys ({}) is not same as description ({})'.format(
            len(heureka_standdata_keys), len(heureka_standdata_desc)))

    for key, desc in [('DGV', 'Grundytevägd medeldiameter (cm)'), ('V', 'Volym (m3sk/ha. produktiv areal)')]:
        if heureka_standdata_keys.index(key) != heureka_standdata_desc.index(desc):
            raise ValueError('Position of {} ({}) does not match with position of its description ({})'.format(
                key, heureka_standdata_keys.index(key), heureka_standdata_desc.index(desc)))

    if len(heureka_treatment_keys) != len(heureka_treatment_desc):
        raise ValueError('Length of heureka treatment keys and descriptions differ')

    for key_list in [fossagrim_keys_to_average_over, fossagrim_keys_to_sum_over, fossagrim_keys_to_most_of]:
        for key in key_list:
            if key not in fossagrim_standdata_keys:
                raise ValueError('Key "{}" is not among default Fossagrim keys'.format(key))
