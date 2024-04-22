import pandas as pd
import datetime
import numpy as np

_sep = ','  # ';'
_dec = '.'  # ','
# TODO
# These parameters should be placed in the project_settings_file
# e.g. 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings.xlsx'
# instead of hardcoded here OR? Currently they can be modified in the Monetization file to test various
# scenarios. If we write them in the project_settings_file then they seem more "eternal"
parameters = pd.DataFrame([
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

variables_used_in_monetization = [
    'Soil Carbon Stock',
    'Total Carbon Deadwood',
    'Total Carbon Living Stumps and Roots',
    'Total Carbon Living Trees (excl. stump and roots)',
    'Total Carbon Stock (dead wood, soil, trees, stumps and roots)',
    'Total Extracted Volume Fub (m³fub)',
    'Mean Age (all trees, always basal area weighted) Before',
    'Treatment',
    'Year'
]

#  Calculation part 1 of Monatization file, Columns A to (including) W
calculation_part1_header = pd.DataFrame([
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


calculation_part1 = pd.concat([
    calculation_part1_header,
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
                   '=V{}/$A$5'.format(_row)] for _row in np.arange(8, 49)])
 ], ignore_index=True)


#  Resampled section of Monatization file, Columns Y to (including) AD
resampled_section = pd.concat(
[
        pd.DataFrame([
            ['Resampled Climate benefit, ton CO2', '', '', '', 'Climate benefit', ''],
            ['', '', '', '', '', ''],
            ['', '', '', '', '', ''],
            ['', '', '=CONCATENATE(\"Ton \"{}$F$2{}\"/yr")'.format(_sep, _sep),
                 '=CONCATENATE(\"Ton \"{}$F$2{}\"/yr")'.format(_sep, _sep), 'Ton CO2/yr', 'Ton CO2'],
            ['', '', 'Linear intpol / 5yr', 'Running average', 'Annual Climate benefit', 'Accumulated Climate benefit'],
            ['t', 'year', '', '', '', ''],
            ['', '', '', '', '', '']]),
        pd.DataFrame([
            [float('{}'.format(_row)),
            '=$B$5+Y{}'.format(_row + 8),
            '=$W{}'.format(9 + int((_row - np.mod(_row, 5))/5)),  # +1 in every 5 iteration
            '=AVERAGE(AA{}:AA{})'.format(_row + 8 - 2,  _row + 8 + 2),
            '=AB{}*IF($F$2=\"C\"{}44/12{}1)'.format(_row + 8, _sep, _sep),
            '=SUM($AC$8:AC{})'.format(_row + 8)
            ] for _row in np.arange(103)])
], ignore_index=True)

resampled_section.iloc[7, 3] = '{} + (AB10-AB12)'.format(resampled_section.iloc[7, 3])
resampled_section.iloc[8, 3] = '{} + (AB10-AB11)'.format(resampled_section.iloc[8, 3])


# Monetary value section, Columns AE to (including) AI
money_value = pd.DataFrame([
    ['', 'Flow1', 'Flow2', 'Total Flow', 'Max potential'],
    ['Volume, m3', 0, 0, '=AF3+AG3',  ''],
    ['Volume % total',  '=AF3/$AI$3', '=AG3/$AI$3','=AH3/$AH$3','=AI3/$AH$3'],
    ['Harvest time', '', '=AF5+AG6*365{}25'.format(_dec), '', ''],
    ['Harvest delay yr', '-', '', '', ''],
    ['', '', '', '', ''],
    ['Root net, kr / m3', '', '=$AF$8', '=$AF$8', '=$AF$8'],
    ['Root net total, kr', '=AF3*AF8', '=AG3*AG8', '=AH3*AH8', '=AI3*AI8']
])


#  Divide climate benefits into flows, columns AJ to (including) AM
cbo_flow = pd.concat([
    pd.DataFrame([
        ['', 'CBO Flow 1', 'CBO Flow 2', 'Project flow'],
        ['Share', '=AF3/$AI$3', '=AG3/$AI$3', '=SUM(AK3:AL3)'],
        ['Harvest year', '=AF5', '=AG5', ''],
        ['', '', '', ''],
        ['t', '', '', ''],
        ['', '', '', '']]),
    pd.DataFrame([
        [float('{}'.format(_row)),
         '=IF($AJ{}<$AN$4{}$AC${}*AK$3{})'.format(_row+8, _sep, _row+8, _sep),
         '=IF($AJ{}<$AN$4{}$AC${}*AL$3{})'.format(_row+8, _sep, _row+8, _sep),
         '=SUM(AK{}:AL{})'.format(_row+8, _row+8)
         ] for _row in np.arange(103)])
], ignore_index=True)


# Project benefits, columns AN to (including) AW
project_benefits = pd.concat([
    pd.DataFrame([
        ['Project Benefits - rental contracts', '', '', '', '', '', '', '', '', ''],
        ['Contract', 'Rent', 'Price growth', 'Lambda', '', '', 'Contract', '', '', ''],
        ['', '', '', '=AO4-AP4', '', '', 100., '', '', ''],
        ['Offset tons', '', 'ton/yr', 'ton', 'Annual', 'Accum', '', '', 'Annual', 'Accum'],
        ['Tau', 'N', 'Farmed offsets', 'Accum offsets',
         '=CONCATENATE($AN$4{}\" yr contract\")'.format(_sep),
         '=CONCATENATE($AN$4{}\" yr contract\")'.format(_sep),
         'Tau', 'N',
         '=CONCATENATE($AT$4{}\" yr contract\")'.format(_sep),
         '=CONCATENATE($AT$4{}\" yr contract\")'.format(_sep)],
        ['', '', '', '', '', '', '', '', '', '']
    ]),
    pd.DataFrame([
        [
            '=IF((AN$4-$Y{})>0{}AN$4-$Y{}{})'.format(_row+8, _sep, _row+8, _sep),
            '=IF(AN{}>0{}1/(1-EXP(-AQ$4*AN{})){})'.format(_row+8, _sep, _row+8, _sep),
            '=IF(AO{}{}$AM{}/AO{}{})'.format(_row+8, _sep, _row+8,  _row+8, _sep),
            '=IF(AN{}>0{}SUM(AP$8:AP{}){}0)'.format(_row+8, _sep, _row+8, _sep),
            '=IF(AO{}{}$AC{}/AO{}{})'.format(_row + 8, _sep, _row + 8, _row + 8, _sep),
            '=IF(AN{}>0{}SUM(AR$8:AR{}){}0)'.format(_row + 8, _sep, _row + 8, _sep),
            '=IF((AT$4-$Y{})>0{}AT$4-$Y{}{})'.format(_row + 8, _sep, _row + 8, _sep),
            '=IF(AT{}>0{}1/(1-EXP(-AQ$4*AT{})){})'.format(_row + 8, _sep, _row + 8, _sep),
            '=IF(AU{}{}$AC{}/AU{}{})'.format(_row + 8, _sep, _row + 8, _row + 8, _sep),
            '=SUM(AV$8:AV{})'.format(_row+8)
        ] for _row in np.arange(103)])
], ignore_index=True)


# Project buffer, columns AX to (including) BC
buffer = pd.concat([
    pd.DataFrame([
        ['Buffer', 'Reserve years', 'Release years', 'Total release + new', 'Tonyears', 'Gradient'],
        ['', '', '=AN4-AY4-1',
         '=SUM(AX8:AX78)+SUMIF(Y8:Y78{}\">=\"&(AY4+1){}AP8:AP78)'.format(_sep, _sep),
         '=SUMIF(AN8:AN78{}\"<=\"&$AZ$4{}AN8:AN78)'.format(_sep, _sep), '=BA4/BB4'],
        ['ton/yr', 'ton', 'ton/yr', 'ton/yr', 'ton', 'ton'],
        ['Buffer reserved', 'Accum buffer', 'Net farmed offsets / yr', 'Buffer released',
         'Accum net farmed offsets', 'Accum buffer released'],
        ['', '', '', '', '', ''],
    ]),
    pd.DataFrame([
        [
            '=IF(Y{}<$AY$4{}AP{}*$AX$4{})'.format(_row+8, _sep, _row+8, _sep),
            '=IF(AX{}>0{}SUM(AX$8:AX{}){}$AF$153)'.format(_row+8, _sep, _row+8, _sep),
            '=AP{}-AX{}'.format(_row+8, _row+8),
            '=IF(AN{}<=$AZ$4{}AN{}*$BC$4-AZ{}{}0)'.format(_row+8, _sep, _row+8, _row+8, _sep),
            '=IF(AN{}>0{}SUM(AZ$8:AZ{}){}0)'.format(_row+8, _sep, _row+8, _sep),
            '=IF(BA{}>0{}SUM($BA$8:BA{}){})'.format(_row+8, _sep, _row+8, _sep),
        ] for _row in np.arange(103)])
], ignore_index=True)


# Section for Fossagrim cash flow, columns BD to (including) BN
fossagrim_values = pd.concat([
    pd.DataFrame([
        ['', '', '', 'Ref price', '=AH9/SUM(BG8:BG78)', '', '', '', 'Mark-up / handling fee', '=(BI8-BH8)/BH8', ''],
        ['', '', '', '', '', '', '', '', '', '', ''],
        ['', '', '', 'Sales', '', '', '', '', '', '', ''],
        ['', '', '', '', 'Forest Owner', 'Total', 'Fossagrim', '', '',
         '=CONCATENATE(\"Annual revenue at \"{}BI8{}\"kr/ton\")'.format(_sep, _sep),
         'Total revenue'],
        ['Year', 'Project Benefit', 'Buffer', 'Offsets', 'Net price', 'Gross Price', 'Cut', 'Sales margin',
         'Gross sales', 'Fossagrim', 'Accum Fossagrim'],
        ['', '', '', '', '', '', '', '', '', '', '']
    ]),
    pd.DataFrame([
        [
            '=Y{}+1'.format(_row+8),
            '=AP{}'.format(_row + 8),
            '=IF(AX{}{}-AX{}{}BA{})'.format(_row+8, _sep, _row+8, _sep, _row+8),
            '=AP{}-AX{}+BA{}'.format(_row+8, _row+8, _row+8),
            '=IF(BG{}>0{}BH{}{})'.format(_row+8, _sep, _row+7, _sep),
            '=BH{}*(1+$BM$2)'.format(_row+8),
            '=BI{}-BH{}'.format(_row+8, _row+8),
            '=IF(BI{}{}BJ{}/BI{}{})'.format(_row+8, _sep, _row+8, _row+8, _sep),
            '=BI{}*$BG{}'.format(_row+8, _row+8),
            '=BG{}*$BJ{}'.format(_row+8, _row+8),
            '=IF(AN{}>0{}SUM(BM$8:BM{}){})'.format(_row+8, _sep, _row+8, _sep)
        ] for _row in np.arange(103)])
], ignore_index=True)

forest_owner_values = pd.concat([
    pd.DataFrame([
        # ['', '', '', '', '', ''],
        ['Forest owner', '', '', '', '', ''],
        ['10 year interest rate', '', 'Base case', 'Project case', '', ''],
        ['Min. net price [kr/ton CO2]',
         '=$BQ$7/($BG$8 + NPV($BP$3{} $BG$9:$BG$50))'.format(_sep),
         '', '', '', ''],
        ['Net price [kr/ton CO2]', '=$BH$8', '', '', '', ''],
        ['Year', '', 'Est annual rev', 'Est annual rev', 'Est total rev', 'Sales completion'],
        ['', 'NPV of total:', '=BQ8+NPV($BP$3{} BQ9:BQ37)'.format(_sep), '=BR8+NPV($BP$3{} BR9:BR37)'.format(_sep), '', '']
        ]),
    pd.DataFrame([
        [
            '=Y{}+1'.format(_row+8),
            '',
            '=IF(ROW()=8{} $AF$3*$AF$8{} IF(ROW()-7=$AG$6{} $AG$3*$AG$8{} 0))'.format(_sep,_sep,_sep,_sep),
            '=BL{}-BM{}'.format(_row + 8, _row + 8),
            '=IF(AO{}>0{}SUM(BR$8:BR{}){})'.format(_row + 8, _sep, _row + 8, _sep),
            '=SUM(BR$8:BR{})/SUM(BR$8:BR$78)'.format(_row+8)
        ] for _row in np.arange(103)])
], ignore_index=True)
