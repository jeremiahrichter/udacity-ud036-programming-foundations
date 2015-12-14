#!/usr/bin/python

import sys

import openpyxl as op

from string_functions import display_menu
from string_functions import display_rows

#   if len(sys.argv) < 2:
if False:
    print('Please specify a spreadsheet file to open.')
else:
    try:
        # wb = op.load_workbook(sys.argv[1])
        wb = op.load_workbook('test.xlsx')  # load a workbook using openpyxl

        # print('Opened file %s' % wb.name)

        if wb is not None:
            exitChosen = False
            row_start = 1
            row_end = 5
            ws_num = 0
            col_start = 1
            col_end = 5

            display_rows(wb, ws_num, row_start, row_end, col_start, col_end)

            while not exitChosen:

                choice = input('choice: ')
                choice = choice if len(choice) == 1 else choice[0]

                if choice == 'n':
                    ws_num += 1
                elif choice == 'r':
                    col_start += 5
                    col_end += 5
                elif choice == 'l' and col_end > 5:
                    col_start -= 5
                    col_end -= 5
                elif choice == 'p' and ws_num > 0:
                    ws_num -= 1
                elif choice == 'd':
                    row_start += 5
                    row_end += 5
                elif choice == 'u' and row_end > 5:
                    row_start -= 5
                    row_end -= 5
                elif choice == 'x':
                    exitChosen = True
                else:
                    print('invalid choice, please choose one of the options provided\n')
                    display_menu(wb, ws_num, row_start, row_end, col_start, col_end)
                    continue
                if not exitChosen:
                    display_rows(wb, ws_num, row_start, row_end, col_start, col_end)

            sys.exit(0)
    except Exception as e:
        sys.exit(1)
