# https://www.heurekaslu.se/wiki/Import_of_stand_register
import pandas as pd
import datetime

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

monetization_parameters = pd.DataFrame([
	['hpap', 2, 'Half time for paper [yrs]'],
	['hsawn', 35, 'Half time for sawn products [yrs]'],
	['SFsawn', 0.58, 'Substitutionsfaktor  för långlivade produkter, excl end-of-life energiåtervinning'],
	['SFpp', '=0.87*B4', 'Substitutionsfaktor för massa och papper'],
	['SFfuel', '=0.4*B4', 'Substitutionsfaktor för biomassa som avänds som bränsle '],
	['p', 0.233, 'Andelen av skördat timmer som blir sågade produkter'],
	['pp', 0.301, 'Andelen av skördad massaved som blir papper'],
	['pf', 0.159, 'Proportion av skörden som går direkt till energiproduktion'],
	['psfuel', 0.8, 'Proportion av förluster från poolen sågade trävaror som ger end of life subst'],
	['k', 0.21, 'Omvandlingsfaktor m3sk till ton C']
])

monetization_calculation_part1_header = pd.DataFrame([
	['Sheet updated', datetime.date.today().isoformat(), '', 'Productive, active area, ha', '', 0.0, '', 'Base case pools, ton CO2', '', '', '', '', '', '', '', 'Project case pools, ton CO2', '', 'Climate benefit, ton CO2', '', '', '', '', ''],
	['By', 'Python script', '', 'Measure, Carbon or CO2?', '', 'CO2', '=IF(F2="C";1;44/12)', 'Unit area: 1 ha', '', '', '', '=CONCATENATE("Active area: ";$F$1;" ha")', '', '', '', '=H2', 'Active area', '=H2', '', '', '=CONCATENATE("Active area: ";$F$1;" ha")', '', ''],
	['del t', 'year 0', 'Total extracted', 'Total extracted', 'Total extrated to sawn prod', 'Prod pool change', 'Substitution', '=CONCATENATE("Ton ";$F$2;"/ha")', '', '', '', '=CONCATENATE("Total ton ";$F$2;"/ha")', '', '', '', '=CONCATENATE("Ton ";$F$2;"/ha")', '=CONCATENATE("Total ton ";$F$2)', '=CONCATENATE("Ton ";$F$2;"/ha")', '=CONCATENATE("Ton ";$F$2;"/ha/";A5;"yr")', '=CONCATENATE("Ton ";$F$2;"/ha/yr")', '=CONCATENATE("Ton ";$F$2)', '=CONCATENATE("Ton ";$F$2;"/";A5;"yr")', '=CONCATENATE("Ton ";$F$2;"/yr")'],
	[5, 2024, 'vol fub m3/ha', 'C ton/ha', 'C ton/ha', 'C ton/ha/ 5yr', 'C ton/ha', 'Forest', 'Product', 'Substitution', 'Base case', 'Forest', 'Product', 'Substitution', 'Base case', 'Project', 'Project', 'Accumulated', 'Interval', 'Yearly', 'Accumulated', 'Interval', 'Yearly'],
	['t', 'year', 'COPY OVER!', '', '', '', '', 'COPY OVER!', '', '', '', '', '', '', '', 'COPY OVER!', '', '', '', '', '', '', ''],
	['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
])


monetization_calculation_part1 = monetization_calculation_part1_header.append(
	pd.DataFrame([[0,
		'=B$5+A{}'.format(_row),
		'="Rearranged results"!AY4*$G$2',
		'=C{}*k'.format(_row),
		'=D{}*p'.format(_row),
		'=IF(H{}>0;I{}-I{};)'.format(_row, _row, _row - 1),
		'=D{}*(p*SFsawn+pp*SFpp+pf*SFfuel)+(E{}-F{})*psfuel*SFfuel'.format(_row, _row, _row),
		'="Rearranged results"!AX4*$G$2', '=IF(H{}>0;E{}+I{}*0,5^(5/hsawn);)'.format(_row, _row, _row - 1),
		'=IF(H{}>0;SUM($G$8:G{});)'.format(_row, _row),
		'=IF(H{}>0;H{}+I{}+J{};)'.format(_row, _row, _row - 1, _row - 1),
		'=H{}*$F$1'.format(_row),
		'=I{}*$F$1'.format(_row),
		'=J{}*$F$1'.format(_row),
		'=K{}*$F$1'.format(_row),
		'="Rearranged results"!BI4*$G$2',
		'=P{}*$F$1'.format(_row),
		'=P{}-K{}'.format(_row, _row),
		'=IF(H{}>0;R{}-R{};)'.format(_row, _row, _row - 1),
		'=S{}/$A$5'.format(_row),
		'=Q{}-O{}'.format(_row, _row),
		'=IF(H{}>0;U{}-U{};)'.format(_row, _row, _row - 1),
		'=V{}/$A$5'.format(_row)] for _row in np.arange(8, 49)]))


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
