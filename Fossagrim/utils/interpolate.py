import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import scipy.interpolate as intrp
from copy import deepcopy

def five_to_one(years, data, verbose=True):
    """
    Tries to interpolate data from Heureka, which has a time step of five years, to use a one year time step

    :param years:
    :param data:
    :param verbose:
    :return:
    """
    pass


def rotation_period_interpolation():
    """
    Reads transposed Heureka data from an excel sheet, with 5 year period increments, and tries to interpolate the data
    with one year increment according to the rotations ('FinalFelling' is the Treatment column), so that the quadratic
    spline interpolation doesn't go haywire when the Total Carbon suddenly drops

    :return:
    """
    verbose = True
    # filename = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\Pål Bjørnstad\\PålBjørnstad Heureka results Spruce and Pine.xlsx"
    filename = "C:\\Users\\marten\\Downloads\\PålBjørnstad Heureka results Spruce and Pine.xlsx"

    bau_spruce = {'usecols': 'D:G', 'skiprows': 3, 'nrows': 40}
    bau_pine = {'usecols': 'D:G', 'skiprows': 49, 'nrows': 40}
    table = pd.read_excel(filename, sheet_name='Heureka results', **bau_pine,
                          names=['Age', 'Total carbon', 'Extracted biomass', 'Treatment'],
                          engine='openpyxl')

    n = len(table['Age'])
    yrs_5 = np.linspace(0, (n - 1) * 5, n)
    yrs_1 = np.linspace(0, int(yrs_5[-1]), int(yrs_5[-1]) + 1)
    # plot original carbon content
    if verbose:
        fig_c, ax_c = plt.subplots()
        ax_c.plot(yrs_5, table['Total carbon'], 'o')
        ax_c.set_xlabel('Years')
        ax_c.set_ylabel('Total carbon')
        fig_a, ax_a = plt.subplots()
        ax_a.plot(yrs_5, table['Age'], 'o')
        ax_a.set_xlabel('Years')
        ax_a.set_ylabel('Mean age of trees')
        fig_b, ax_b = plt.subplots()
        ax_b.plot(yrs_5, table['Extracted biomass'], 'o')
        ax_b.set_xlabel('Years')
        ax_b.set_ylabel('Extracted biomass')
    else:
        ax_c = None
        ax_a = None
        ax_b = None

    # Find number of Final Fellings, and their indexes
    felling_indexes = []
    for i, treatment in enumerate(table['Treatment']):
        if 'Final' in treatment:
            felling_indexes.append(i)
    felling_indexes.append(n-1)

    output_age = np.array([])
    output_carbon = np.array([])
    output_biomass = np.array([])
    output_treatment = []

    last_index = 0
    counter = 0
    last_max_age = None
    # Split the data in to rotation periods
    for this_i in felling_indexes:
        if this_i == 0:
            continue
        counter += 1

        # the "this_i+1" ensures that the first value of the next rotation period is included in the
        # interpolation. The append function below avoids that value to the appended twice
        this_age = deepcopy(table['Age'][last_index:this_i+1])
        this_age.iloc[0] = 0.
        this_carbon = table['Total carbon'][last_index:this_i+1]
        this_biomass = table['Extracted biomass'][last_index:this_i+1]
        this_treatment = table['Treatment'][last_index:this_i+1]

        n_periods = len(this_age)
        five_years = np.linspace(0, (n_periods-1) * 5, n_periods)
        one_years = np.linspace(0, int(five_years[-1]), int(five_years[-1]) + 1)

        resampled_biomass = np.zeros(len(one_years))
        resampled_treatment = [''] * len(one_years)
        # find out which years we have extracted biomass, and insert those into the resampled biomass
        for k, bio in enumerate(this_biomass):
            if bio != 0.:
                year = five_years[k]
                resampled_biomass[np.argmin(np.sqrt((one_years-year)**2))] = bio
        # find out which years we have done some treatment, and insert those into the resampled treatment
        for k, treat in enumerate(this_treatment):
            if treat != 'None':
                year = five_years[k]
                # print(year, treat)
                resampled_treatment[np.argmin(np.sqrt((one_years-year)**2))] = treat

        if counter == 1:
            output_carbon = np.append(output_carbon, intrp.interp1d(five_years, this_carbon, kind='cubic')(one_years))
            new_age = intrp.interp1d(five_years, this_age)(one_years)
            new_age[0] = table['Age'].iloc[0]  # inherent max age from last rotation period
            output_age = np.append(output_age, new_age)
            output_biomass = np.append(output_biomass, resampled_biomass)
            output_treatment += resampled_treatment
        else:
            output_carbon = np.append(output_carbon[:-1],
                                      intrp.interp1d(five_years, this_carbon, kind='cubic')(one_years))
            new_age = intrp.interp1d(five_years, this_age)(one_years)
            new_age[0] = last_max_age  # inherent max age from last rotation period
            output_age = np.append(output_age[:-1], new_age)
            output_biomass = np.append(output_biomass[:-1], resampled_biomass)
            output_treatment += resampled_treatment[1:]

        last_max_age = output_age[-1]
        last_index = this_i

    if verbose:
        ax_c.plot(yrs_1, output_carbon, 'r-')
        ax_a.plot(yrs_1, output_age, 'r-')
        ax_b.plot(yrs_1, output_biomass, 'r-')
        plt.show()
        for i, treat in enumerate(output_treatment):
            if treat != '':
                print(yrs_1[i], treat)

    output = pd.DataFrame(data={
        'Age [yrs]': output_age,
        'Total carbon [ton C/ha]': output_carbon,
        'Extracted biomass [m3 fub/ha]': output_biomass,
        'Treatment': output_treatment
    })

    with pd.ExcelWriter(filename, mode='a') as writer:
        output.to_excel(writer, sheet_name='BAU, Pine, resampled', engine='openpyxl')
    return output




