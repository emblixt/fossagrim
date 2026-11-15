import pytest

from Fossagrim.io.common import read_excel, read_csv_file
from Fossagrim.io.stands import (arrange_export_of_one_stand, average_over_stands, export_fossagrim_stand_to_heureka,
                                 get_active_forests_from_stand, matched_average_groups, EXTRA_EXPORT_KEYS)
from Fossagrim.unit_tests.conftest import DEFAULT_STANDS, stand_row, write_forvaltningsplan

ID = 'Fossagrim ID'


@pytest.fixture
def stand_table(forvaltningsplan_file):
    return read_excel(forvaltningsplan_file, 7, 'data')


def test_average_over_stands(stand_table):
    avg = average_over_stands({'Spruce': ['1-1', '2-1'], 'Pine': ['3-1']}, stand_table, ID, 'Avg ')
    assert len(avg) == 2
    spruce = avg.iloc[0]
    assert spruce[ID] == 'Avg Spruce'
    assert spruce['Bestand'] == 'Avg Spruce'
    # sums
    assert spruce['Prod.areal'] == pytest.approx(40.)
    assert spruce['Gran'] == pytest.approx(680.)
    assert spruce['Total'] == pytest.approx(800.)
    # area weighted averages: (10 * 60 + 30 * 80) / 40
    assert spruce['År'] == pytest.approx(75.)
    assert spruce['Altitude'] == pytest.approx(300.)
    # 'most of': area weighted averages 13.75 and 2.75 -> closest original values
    assert spruce['Markslag'] == 14
    assert spruce['SoilMoistureCode'] == 3
    # copied from the first stand
    assert spruce['Tetthet'] == 'Middels'
    assert spruce['Bonitering\ntreslag'] == 'Gran'

    pine = avg.iloc[1]
    assert pine[ID] == 'Avg Pine'
    assert pine['År'] == pytest.approx(100.)


def test_average_over_stands_accepts_decimal_comma(tmp_path):
    stands = [stand_row(1, '1-1', 'Gran', 1, prod_areal='10,0', gran=100., furu=0., lauv=0., alder='60,0'),
              stand_row(2, '2-1', 'Gran', 1, prod_areal=30., gran=100., furu=0., lauv=0., alder=80.)]
    table = read_excel(write_forvaltningsplan(str(tmp_path / 'f.xlsx'), stands), 7, 'data')
    avg = average_over_stands({'Spruce': ['1-1', '2-1']}, table, ID, '')
    assert avg.iloc[0]['Prod.areal'] == pytest.approx(40.)
    assert avg.iloc[0]['År'] == pytest.approx(75.)


def test_average_over_stands_skips_groups_without_match(stand_table, capsys):
    average_over = {'Spruce': ['1-1'], 'Birch': ['does not exist'], 'Pine': ['3-1']}
    avg = average_over_stands(average_over, stand_table, ID, '')
    assert list(avg[ID]) == ['Spruce', 'Pine']
    assert 'no match for average group Birch' in capsys.readouterr().out
    assert matched_average_groups(average_over, stand_table, ID) == ['Spruce', 'Pine']


def test_arrange_export_of_one_stand_converts_units(stand_table):
    avg = average_over_stands({'Spruce': ['1-1', '2-1']}, stand_table, ID, 'Avg ')
    kwargs = arrange_export_of_one_stand(0, avg, EXTRA_EXPORT_KEYS + ['UniqueID'], unique_id='v7')
    assert kwargs['StandId'] == 'Avg Spruce'
    assert kwargs['CountyCode'] == 13
    assert kwargs['SiteIndexSpecies'] == 'G'
    assert kwargs['InventoryYear'] == 2023
    assert kwargs['ProdArea'] == pytest.approx(4.0)  # daa -> ha
    assert kwargs['V'] == pytest.approx(20. * 10 * 1.19)  # m3fub/daa -> m3sk/ha
    assert kwargs['N'] == pytest.approx(800.)  # trees/daa -> trees/ha
    assert kwargs['PlantDensity'] == pytest.approx(2000.)
    assert kwargs['PropSpruce'] == pytest.approx(680. / 800.)
    assert kwargs['PropPine'] == pytest.approx(70. / 800.)
    assert kwargs['PropBirch'] == pytest.approx(50. / 800.)
    assert kwargs['MeanAge'] == pytest.approx(75.)
    assert kwargs['UserDefinedVariable2_RotationPeriod'] == pytest.approx(70.)
    assert kwargs['UniqueID'] == 'v7'
    assert kwargs['PropOak'] == ''


def test_arrange_export_rejects_unknown_species(stand_table):
    stand_table.loc[0, 'Bonitering\ntreslag'] = 'Eik'
    with pytest.raises(NotImplementedError):
        arrange_export_of_one_stand(0, stand_table)


def test_export_averaged_stands(forvaltningsplan_file, tmp_path):
    csv_file = str(tmp_path / 'Averaged stand data.csv')
    # Birch has no matching stands and sits between the other groups, so the notes must not be shifted
    average_over = {'Spruce': ['1-1', '2-1'], 'Birch': ['none'], 'Pine': ['3-1']}
    n = export_fossagrim_stand_to_heureka(forvaltningsplan_file, csv_file, average_over=average_over,
                                          stand_id_key=ID, sheet_name='data', average_name='P1 ', unique_id=42)
    assert n == 2
    data, _ = read_csv_file(csv_file)
    assert data['StandId'] == ['P1 Spruce', 'P1 Pine']
    assert data['SiteIndexSpecies'] == ['G', 'T']
    assert data['Note'] == ['Area weighted average for Spruce of stands 1-1, 2-1',
                            'Area weighted average for Pine of stands 3-1']
    assert data['UniqueID'] == [42, 42]
    assert data['ProdArea'] == [4, 2]


def test_export_single_stand_skips_unproductive(forvaltningsplan_file, tmp_path):
    csv_file = str(tmp_path / 'stands.csv')
    n = export_fossagrim_stand_to_heureka(forvaltningsplan_file, csv_file, stand_id_key=ID, sheet_name='data')
    assert n == 4  # the unproductive stand is skipped
    data, _ = read_csv_file(csv_file)
    assert '5-1' not in data['StandId']
    assert 'UniqueID' in data and data['UniqueID'] == [''] * 4

    n = export_fossagrim_stand_to_heureka(forvaltningsplan_file, csv_file, stand_id_key=ID, sheet_name='data',
                                          this_stand_only='3-1')
    assert n == 1


def test_get_active_forests_from_stand(forvaltningsplan_file):
    active, areas, fractions = get_active_forests_from_stand(forvaltningsplan_file, stand_id_key=ID,
                                                             sheet_name='data')
    assert active == {'Spruce': ['1-1', '2-1'], 'Pine': ['3-1']}  # inactive stands are left out, no Birch
    assert areas == {'Spruce': [10., 30.], 'Pine': [20.]}
    assert fractions['Spruce'] == pytest.approx(40. / 60.)
    assert fractions['Pine'] == pytest.approx(20. / 60.)


def test_get_active_forests_unknown_species(tmp_path):
    stands = DEFAULT_STANDS + [stand_row(6, '6-1', 'Eik', 1, prod_areal=5., gran=0., furu=0., lauv=10.)]
    f = write_forvaltningsplan(str(tmp_path / 'f.xlsx'), stands)
    with pytest.raises(ValueError, match='6-1'):
        get_active_forests_from_stand(f, stand_id_key=ID, sheet_name='data')


def test_get_active_forests_without_active_stands(tmp_path):
    f = write_forvaltningsplan(str(tmp_path / 'f.xlsx'), [DEFAULT_STANDS[3]])
    with pytest.raises(ValueError, match='No active stands'):
        get_active_forests_from_stand(f, stand_id_key=ID, sheet_name='data')
