#!/usr/bin/python

import sys

import openpyxl as op

from string_functions import display_menu
from string_functions import display_rows


def main():
    if len(sys.argv) < 2:
        print('Please enter a filename')
        exit(1)
    else:
        try:
            wb = op.load_workbook(sys.argv[1])

            if wb is not None:
                exit_chosen = False
                row_start = 1
                row_end = 5
                ws_num = 0
                col_start = 1
                col_end = 5

                display_rows(wb, ws_num, row_start, row_end, col_start, col_end)

                while not exit_chosen:

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
                        exit_chosen = True
                    else:
                        print('invalid choice, please choose one of the options provided\n')
                        display_menu(wb, ws_num, row_start, row_end, col_start, col_end)
                        continue
                    if not exit_chosen:
                        display_rows(wb, ws_num, row_start, row_end, col_start, col_end)

                sys.exit(0)
        except (NameError, TypeError) as e:
            print(str(e))
            exit(1)
