# -*- coding:utf-8 -*-
from googletrans import Translator
import sys

translator = Translator()

## this is a test
#print(translator.translate('你在干嘛').text)

#use py to operate excel
import xlrd
from xlutils.copy import copy

'''
	#open excel to read
	data = xlrd.open_workbook('excelFile.xls')
	
	tabel = data.sheets()[0]
	
	#get the raw/col number
	#nrows = table.nrows
	ncols = table.ncols
	
	#get the raw value
	#table.row_values(i)
	#get the col value
	col = table.col_values(i)
	'''

#def open(file, mode='w+', buffering=None, encoding=None, errors=None, newline=None, closefd=True)


after_translator = {}

def write_excel(path, start, nrows, num_sheet):
	
	workbook = xlrd.open_workbook(path)
	newb = copy(workbook)
	sheet = newb.get_sheet(num_sheet)
	sheet.write(0, 20, '翻译')
	
	for i in range(start, nrows):
		sheet.write(i, 20, after_translator[i])
		print(i)
	
	newb.save(path)

def c_2_e(path = '../resource/r1copy.xls', num_sheet = 1):
	
	workbook = xlrd.open_workbook(path)
	sheets = workbook.sheets()
	
	sheet = sheets[num_sheet]
	
	if 'r1' in path:
		if num_sheet == 1:
			col_val = sheet.col_values(5)
		elif num_sheet == 2:
			col_val = sheet.col_values(6)
	elif 'r2' in path:
		col_val = sheet.col_values(9)
	else:
		col_val = sheet.col_values(10)
	
	error = []

	nrows = len(col_val)
	#nrows = 100
	times = nrows // 1000
	#times = 1   #debug
	for index_out in range(times+1):
		start = index_out * 1000
		if start == 0:
			start = 1
		end = (index_out + 1) * 1000
		if end > nrows:
			end = nrows

		for i in range(start, end):
			try:
				text = translator.translate(col_val[i]).text
				after_translator[i] = text
			except:
				error.append(i)
			print(i)

		'''
			while(len(error) != 0):
			index = error.pop(0)
			try:
			text = translator.translate(col_val[index]).text
			after_translator.append(text)
			except:
			error.append(i)
		'''

		print('开始错误修正')
		while len(error) != 0:
			index_error = error.pop(0)
			print(index_error)
			try:
				text = translator.translate(col_val[index_error]).text
				after_translator[index_error] = text
			except:
				error.append(index_error)
		print('完成错误修正')

		print('开始翻译')
		write_excel(path, start, end, num_sheet)
		print('翻译成功')







