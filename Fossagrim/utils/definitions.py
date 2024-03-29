# https://www.heurekaslu.se/wiki/Import_of_stand_register
import pandas as pd
import datetime
import numpy as np

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
    'UserDefinedVariable2',
    'UserDefinedVariable3',
    'UserDefinedVariable4',
    'UserDefinedVariable5',
    'UserDefinedVariable6',
    'UserDefinedVariable7',
    'UserDefinedVariable8',
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
    '','','','','','','','','','','','','','','','','','','','',''
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
    return translation


def test_length_of_keys():
    if len(heureka_standdata_keys) != len(heureka_standdata_desc):
        raise ValueError('Length of heureka keys ({}) is not same as description ({})'.format(
            len(heureka_standdata_keys), len(heureka_standdata_desc)))

    if heureka_standdata_keys.index('DGV') != heureka_standdata_desc.index('Grundytevägd medeldiameter (cm)'):
        raise ValueError('Position of DGV ({}) does not match with position of its description ({})'.format(
            heureka_standdata_desc.index('DGV'), heureka_standdata_desc.index('Grundytevägd medeldiameter (cm)')
        ))

    if heureka_standdata_keys.index('V') != heureka_standdata_desc.index('Volym (m3sk/ha. produktiv areal)'):
        raise ValueError('Position of V ({}) does not match with position of its description ({})'.format(
            heureka_standdata_desc.index('V'), heureka_standdata_desc.index('Volym (m3sk/ha. produktiv areal)')
        ))


def test_key_existance():
    for key_list in [fossagrim_keys_to_average_over, fossagrim_keys_to_sum_over, fossagrim_keys_to_most_of]:
        for key in key_list:
            if key not in fossagrim_standdata_keys:
                raise ValueError('Key "{}" is not among default Fossagrim keys'.format(key))
