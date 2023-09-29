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

_sep = ','  # ';'
_dec = '.'  # ','
monetization_parameters = pd.DataFrame([
	['hpap', 2, 'Half time for paper [yrs]'],
	['hsawn', 35, 'Half time for sawn products [yrs]'],
	['SFsawn', '0{}58'.format(_dec), 'Substitutionsfaktor  för långlivade produkter, excl end-of-life energiåtervinning'],
	['SFpp', '=0.87*B4', 'Substitutionsfaktor för massa och papper'],
	['SFfuel', '=0.4*B4', 'Substitutionsfaktor för biomassa som avänds som bränsle '],
	['p', '0{}233'.format(_dec), 'Andelen av skördat timmer som blir sågade produkter'],
	['pp', '0{}301'.format(_dec), 'Andelen av skördad massaved som blir papper'],
	['pf', '0{}159'.format(_dec), 'Proportion av skörden som går direkt till energiproduktion'],
	['psfuel', '0{}8'.format(_dec), 'Proportion av förluster från poolen sågade trävaror som ger end of life subst'],
	['k', '0{}21'.format(_dec), 'Omvandlingsfaktor m3sk till ton C']
])

monetization_calculation_part1_header = pd.DataFrame([
    ['Sheet updated', datetime.date.today().isoformat(), '', 'Productive, active area, ha', '', 0, '',
     'Base case pools, ton CO2', '', '', '', '', '', '', '', 'Project case pools, ton CO2', '',
     'Climate benefit, ton CO2', '', '', '', '', ''],
    ['By', 'Python script', '', 'Measure, Carbon or CO2?', '', 'CO2', '=IF(F2=\"C\"{}1{}44/12)'.format(_sep, _sep),
     'Unit area: 1 ha', '', '', '', '=CONCATENATE(\"Active area: \"{}$F$1{}\" ha\")'.format(_sep, _sep), '', '', '',
     '=H2', 'Active area', '=H2', '', '', '=CONCATENATE(\"Active area: \"{}$F$1{}\" ha\")'.format(_sep, _sep), '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['del t', 'year 0', 'Total extracted', 'Total extracted', 'Total extrated to sawn prod', 'Prod pool change',
     'Substitution', '=CONCATENATE(\"Ton \"{}$F$2{}\"/ha\")'.format(_sep, _sep), '', '', '',
     '=CONCATENATE(\"Total ton \"{}$F$2{}\"/ha\")'.format(_sep, _sep), '', '', '',
     '=CONCATENATE(\"Ton \"{}$F$2{}\"/ha\")'.format(_sep, _sep), '=CONCATENATE(\"Total ton \"{}$F$2)'.format(_sep),
     '=CONCATENATE(\"Ton \"{}$F$2{}\"/ha\")'.format(_sep, _sep),
     '=CONCATENATE(\"Ton \"{}$F$2{}\"/ha/\"{}A5{}\"yr\")'.format(_sep, _sep, _sep, _sep),
     '=CONCATENATE(\"Ton \"{}$F$2{}\"/ha/yr\")'.format(_sep, _sep), '=CONCATENATE(\"Ton \"{}$F$2)'.format(_sep),
     '=CONCATENATE(\"Ton \"{}$F$2{}\"/\"{}A5{}\"yr\")'.format(_sep, _sep, _sep, _sep),
     '=CONCATENATE(\"Ton \"{}$F$2{}\"/yr\")'.format(_sep, _sep)],
    [5, 2024, 'vol fub m3/ha', 'C ton/ha', 'C ton/ha', 'C ton/ha/ 5yr', 'C ton/ha', 'Forest', 'Product',
     'Substitution', 'Base case', 'Forest', 'Product', 'Substitution', 'Base case', 'Project', 'Project',
     'Accumulated', 'Interval', 'Yearly', 'Accumulated', 'Interval', 'Yearly'],
    ['t', 'year', 'COPY OVER!', '', '', '', '', 'COPY OVER!', '', '', '', '', '', '', '', 'COPY OVER!', '', '', '',
     '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
])


monetization_calculation_part1 = monetization_calculation_part1_header.append(
    pd.DataFrame([['=$A$5 *  (row() - 8)',
                   '=B$5+A{}'.format(_row),
                   '=\'Rearranged results\'!G{}*$G$2'.format(_row - 4),
                   '=C{}*k'.format(_row),
                   '=D{}*p'.format(_row),
                   '=IF(H{}>0{}I{}-I{}{})'.format(_row, _sep, _row, _row - 1, _sep),
                   '=D{}*(p*SFsawn+pp*SFpp+pf*SFfuel)+(E{}-F{})*psfuel*SFfuel'.format(_row, _row, _row),
                   '=\'Rearranged results\'!F{}*$G$2'.format(_row - 4),
                   '=IF(H{}>0{}E{}+I{}*0{}5^(5/hsawn){})'.format(_row, _sep, _row, _row - 1, _dec, _sep),
                   '=IF(H{}>0{}SUM($G$8:G{}){})'.format(_row, _sep, _row, _sep),
                   '=IF(H{}>0{}H{}+I{}+J{}{})'.format(_row, _sep, _row, _row - 1, _row - 1, _sep),
                   '=H{}*$F$1'.format(_row),
                   '=I{}*$F$1'.format(_row),
                   '=J{}*$F$1'.format(_row),
                   '=K{}*$F$1'.format(_row),
                   '=\'Rearranged results\'!Q{}*$G$2'.format(_row - 4),
                   '=P{}*$F$1'.format(_row),
                   '=P{}-K{}'.format(_row, _row),
                   '=IF(H{}>0{}R{}-R{}{})'.format(_row, _sep, _row, _row - 1, _sep),
                   '=S{}/$A$5'.format(_row),
                   '=Q{}-O{}'.format(_row, _row),
                   '=IF(H{}>0{}U{}-U{}{})'.format(_row, _sep, _row, _row - 1, _sep),
                   '=V{}/$A$5'.format(_row)] for _row in np.arange(8, 49)]))


monetization_resampled_section = pd.DataFrame([
    ['Resampled Climate benefit, ton CO2', '', '', '', 'Climate benefit', ''],
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['', '', '=CONCATENATE(\"Ton \"{}$F$2{}\"/yr")'.format(_sep, _sep),
     '=CONCATENATE(\"Ton \"{}$F$2{}\"/yr")'.format(_sep, _sep), 'Ton CO2/yr', 'Ton CO2'],
    ['', '', 'Linear intpol / 5yr', 'Running average', 'Annual Climate benefit', 'Accumulated Climate benefit'],
    ['t', 'year', '', '', '', ''],
    ['', '', '', '', '', '']
])

monetization_resampled_section = monetization_resampled_section.append(
    pd.DataFrame([
       ['{}'.format(_row),
        '=$B$5+Y{}'.format(_row + 8),
        '=$W{}'.format(9 + int((_row - np.mod(_row, 5))/5)),  # +1 in every 5 iteration
        '=AVERAGE(AA{}:AA{})'.format(_row + 8 - 2,  _row + 8 + 2),
        '=AB{}*IF($F$2=\"C\"{}44/12{}1)'.format(_row + 8, _sep, _sep),
        '=SUM($AC$8:AC{})'.format(_row + 8)
        ] for _row in np.arange(103)])
)
monetization_resampled_section.iloc[7, 3] = '{} + (AB10-AB12)'.format(monetization_resampled_section.iloc[7, 3])
monetization_resampled_section.iloc[8, 3] = '{} + (AB10-AB11)'.format(monetization_resampled_section.iloc[8, 3])


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
