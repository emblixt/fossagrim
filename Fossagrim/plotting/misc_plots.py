import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import Fossagrim.utils.projects as fup
import Fossagrim.io.fossagrim_io as fio

markers = ['o', 'v', '^', '<', '>', 's', 'p', 'P', '*', 'X', 'D']

def find_first_index(table, key, value):
	found = False
	i = None
	if isinstance(key, list) and isinstance(value, list):
		for i, this_value in enumerate(table[key[0]]):
			if (value[0] in this_value) and (value[1] in table[key[1]][i]):
				found = True
				break
	else:
		for i, this_value in enumerate(table[key]):
			if this_value == value:
				found = True
				break
	if found:
		return i
	else:
		return None


def plot_plant_density():
	projects = ['FHF23-0{}'.format(_x) for _x in ['03', '04', '05', '06', '07', '08', '09', '10', '12']]
	spruce_sis = []; pine_sis = []; spruce_plant_den = []; pine_plant_den = []

	for project_tag in projects:
		print(project_tag)
		_, _, _, _, _, _, _, csv_stand_file, csv_treatment_file = fup.project_settings(project_tag)
		this_table = pd.read_csv(csv_stand_file, sep=';', header=1, encoding="ISO-8859-1")
		i_spruce = find_first_index(this_table, 'SiteIndexSpecies', 'G')
		i_pine = find_first_index(this_table, 'SiteIndexSpecies', 'T')
		if i_spruce is None:
			spruce_sis.append(np.nan)
		else:
			spruce_sis.append(this_table['SIS'][i_spruce])

		if i_pine is None:
			pine_sis.append(np.nan)
		else:
			pine_sis.append(this_table['SIS'][i_pine])

		this_table = pd.read_csv(csv_treatment_file, sep=';', header=1, encoding="ISO-8859-1")
		i_spruce = find_first_index(this_table, ['StandId', 'Treatment'], ['Spruce', 'Planting'])
		i_pine = find_first_index(this_table, ['StandId', 'Treatment'], ['Pine', 'Planting'])
		if i_spruce is None:
			spruce_plant_den.append(np.nan)
		else:
			spruce_plant_den.append(this_table['PlantDensity'][i_spruce])
		if i_pine is None:
			pine_plant_den.append(np.nan)
		else:
			pine_plant_den.append(this_table['PlantDensity'][i_pine])

	fig, ax = plt.subplots()
	for i, project_tag in enumerate(projects):
		ax.scatter(spruce_sis[i], spruce_plant_den[i], c='b', marker=markers[i], label=project_tag)
		ax.scatter(pine_sis[i], pine_plant_den[i], c='g', marker=markers[i], label='_nolegend_')

	ax.set_xlabel('SIS')
	ax.set_ylabel('Plant density [plants/ha]')
	ax.set_title('Plant density vs Bonity (blue = Spruce)')
	ax.legend()
	plt.show()

	return spruce_sis, pine_sis, spruce_plant_den, pine_plant_den
