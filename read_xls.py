#!/usr/bin/python

import sys
import openpyxl as op

if len(sys.argv) < 2:
  print('Please specify a spreadsheet file to open.')
else:
  try:
#     wb = op.load_workbook(sys.argv[1])
    wb = op.load_workbook('test.xlsx')

    print('Opened file %s' % sys.argv[1])
  except:
    print('unable to open file %s' % sys.argv[1])


if wb:
  exit = False
