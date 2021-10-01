import xlsxwriter

workbook = xlsxwriter.Workbook('write_dict.xlsx')
worksheet = workbook.add_worksheet()

my_dict = {'Bob': [10, 11, 12],
           'Ann': [20, 21, 22],
           'May': [30, 31, 32]}

col_num = 0
for key, value in my_dict.items():
    worksheet.write(0, col_num, key)
    worksheet.write_column(1, col_num, value)
    col_num += 1

workbook.close()


