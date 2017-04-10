from win32com.client.gencache import EnsureDispatch
from win32com.client import constants
from os import path
from os import getcwd
import os
import re
import shutil


def run():
    # get current working directory of the py file
    print('Begin Processing...')
    cwd = getcwd()
    source_file_dir = path.join(cwd, 'original reports')
    handled_file_dir = path.join(cwd, 'handled reports')

    excel = EnsureDispatch('Excel.Application')
    excel.Visible = False

    pattern = r'^日期：(\d{4})年(\d{2})月(\d{2})日6:00 至'
    # 日期：2017年03月19日6:00 至

    data = {}
    unhandled_files = {}
    try:
        for root, dirs, files in os.walk(source_file_dir):
            for file in files:
                file_path = path.join(root, file)
                print('  Reading file: ' + file + '...\t', end='\t')
                wkb_src = excel.Workbooks.Open(file_path)
                wks_src = wkb_src.Worksheets('迎峰度夏每日汇报表')
                wks_src.Activate()

                # get report date
                report_date_string = wks_src.Cells(2,1).Value
                m = re.match(pattern, report_date_string)
                report_date = '-'.join(m.groups())

                data_range = get_data_range(excel, wks_src)

                # if content is not found, skip this file
                if  data_range is not None:
                    entry = []
                    for rng in wks_src.Range(data_range).SpecialCells(constants.xlCellTypeVisible).Rows:
                        line = [report_date]
                        for col in [2, 3, 5, 6, 11, 12, 14, 16]:
                            line.append(rng.Cells(1, col).Value)
                        entry.append(line)

                    data[file] = [file_path, entry]
                    print('Done!')

                else:
                    unhandled_files[file] = file_path
                    print("Skipped!")

                wkb_src.Close(SaveChanges=False)

        # copy data to excel
        total = len(data)
        print('\n  Copy data to destination excel file, {} files in total'.format(total))

        wkb_dest = excel.Workbooks.Open(path.join(cwd, '调度日报汇总.xlsx'))
        wks_dest = wkb_dest.Worksheets('Summary')

        if wks_dest.Range('A2').Value == '':
            dest_start_row = 2
        else:
            dest_start_row = wks_dest.UsedRange.Rows.Count + 1

        for file_name, file_info in data.items():
            file_path = file_info[0]
            file_contents = file_info[1]

            for entry in file_contents:
                for cnt, cell_data in enumerate(entry, 2):
                    wks_dest.Cells(dest_start_row, cnt).Value = cell_data
                dest_start_row += 1

            # move file to handled folder
            shutil.move(file_path, path.join(handled_file_dir, file_name))
            print('  Moving {} to restore folder...\tDone!'.format(file_name))

        if len(unhandled_files) != 0:
            print()
            print('  Please check the files failed to be processed:')
            for entry in unhandled_files:
                print(' '*4 + entry)

        print('All Processes Finished!')
        excel.Visible = True

    except Exception as e:
        print(e)

    finally:
        pass
        # excel.Quit()


def get_data_range(excel, wks_src):

    max_row_count = wks_src.UsedRange.Rows.Count
    find_range = 'A1:R' + str(max_row_count)
    wks_src.Range('A1').Select()
    rng = wks_src.Range(find_range).Find(What='今日计划停役', After=excel.ActiveCell,
                                         LookIn=constants.xlValues, LookAt=constants.xlPart,
                                         SearchOrder=constants.xlByColumns, SearchDirection=constants.xlNext,
                                         MatchCase=False, MatchByte=False, SearchFormat=False)

    if rng == None:
        return None

    data_start_row = rng.Row + 2

    rng = wks_src.Range(find_range).Find(What='今日保电任务', After=excel.ActiveCell,
                                         LookIn=constants.xlValues, LookAt=constants.xlPart,
                                         SearchOrder=constants.xlByColumns, SearchDirection=constants.xlNext,
                                         MatchCase=False, MatchByte=False, SearchFormat=False)
    if rng == None:
        return None

    data_stop_row = rng.Row - 1
    data_range = 'A' + str(data_start_row) + ':R' + str(data_stop_row)
    return data_range


if __name__ == '__main__':
    run()
    os.system('pause')
