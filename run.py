from excel_service import create_workbook, check_toloka, split_workbook


def run_toloka_checking():
    workbooks = split_workbook(check_toloka(create_workbook()))
    for filename in workbooks.keys():
        workbooks[filename].save(filename+'.xlsx')


if __name__ == '__main__':
    run_toloka_checking()
