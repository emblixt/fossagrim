from Fossagrim.utils import definitions as d


def test_key_definitions_are_consistent():
    d.validate_key_definitions()


def test_translation_targets_exist():
    translation = d.translate_keys_from_fossagrim_to_heureka()
    for heureka_key, fossagrim_key in translation.items():
        assert heureka_key in d.heureka_standdata_keys
        assert fossagrim_key in d.fossagrim_standdata_keys


def test_user_defined_columns_have_expected_positions():
    """Other tools (and old scripts) rely on the position of these columns in the stand csv files"""
    assert d.heureka_standdata_keys.index('UserDefinedVariable2_RotationPeriod') == 98
    assert d.heureka_standdata_keys.index('UserDefinedVariable5_ClimateEffect') == 101
    assert d.heureka_standdata_keys.index('UniqueID') == 104


def test_species_mappings():
    for norwegian_name in d.HEUREKA_SITE_INDEX_SPECIES:
        assert any(norwegian_name in names for names in d.SPECIES_NAMES.values())
