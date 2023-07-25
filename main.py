import os
import Fossagrim.io.fossagrim_io as fio


def FHF23_003(arrange='import'):
    project_tag = 'FHF23-003'
    bau_tag = '{} Business as usual'.format(project_tag)
    preserv_tag = '{} Preservation'.format(project_tag)
    project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra"
    stand_file = os.path.join(project_folder, '{} Bestandsutvalg.xlsx'.format(project_tag))
    result_file = os.path.join(project_folder, '{} Heureka results.xlsx'.format(project_tag))

    def arrange_import():
        fio.export_fossagrim_stand_to_heureka(
            stand_file,
            os.path.join(project_folder, '{} Averaged stand data.csv'.format(project_tag)),
            average_over=[67, 68, 70, 72, 75, 79, 80, 85, 88, 90, 92, 93, 95, 97, 98, 99, 104, 105, 109]
        )

        fio.export_fossagrim_treatment(
            stand_file,
            os.path.join(project_folder, '{} Averaged treatment.csv'.format(project_tag)),
            average_over=[67, 68, 70, 72, 75, 79, 80, 85, 88, 90, 92, 93, 95, 97, 98, 99, 104, 105, 109]
        )

    def arrange_results():
        fio.rearrange_raw_heureka_results(result_file, [bau_tag, preserv_tag])

    if arrange == 'import':
        arrange_import()
    elif arrange == 'results':
        arrange_results()


if __name__ == '__main__':
    FHF23_003('results')
