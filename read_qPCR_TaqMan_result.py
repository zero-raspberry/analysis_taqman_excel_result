#! /usr/local/bin/python3
# _*_ coding:utf-8 _*_


import sys
import xlrd
import xlwt


def xls2ct(fnames):
    """
    get ct information from given excel
    output:
        (samplename(str), target(str), ct(float OR str))
    """
    for fname in fnames:
        try:
            wb = xlrd.open_workbook(fname)
        except:
            print('error when loading excel file')
            raise

        result_sheet = wb.sheet_by_name('Results')

        found = False
        for row in result_sheet.get_rows():
            if row[0].value == 'Well':
                found = True
                yield row
                continue
            if found and len(row) > 1:
                yield row


def main():
    filenames = sys.argv[1:]
    savefilename = 'result.xls'
    wbk = xlwt.Workbook(encoding='utf-8')
    sheet = wbk.add_sheet('results')

    for row_idx, row_line in enumerate(xls2ct(filenames)):
        for column_idx, cell in enumerate(row_line):
            sheet.write(row_idx, column_idx, cell.value)

    wbk.save(savefilename)
    print('data saved in file: {}'.format(savefilename))


if __name__ == '__main__':
    main()
    print('done!')
