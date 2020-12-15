import os

import xlrd
import xlwt


def welcome():
	'''
	引导欢迎 程序，必须输入作者姓名才能继续。
	:return:
	'''
	print('欢迎使用excel合并程序。')
	input_str = ''
	while input_str != '吴刚':
		input_str = input('请输入作者姓名继续：')


def choice_exception():
	'''
	获取标题行数
	:return: 返回标题行数变量
	'''
	rows = input('请输入标题行数：')
	return rows


def read_list():
	'''
	读取当前目录及子目录，判断是否excel文件
	:return: excel文件列表
	'''
	all_list = os.listdir('.')
	result = []
	for i in all_list:
		print('正在检查%s' % i)
		if os.path.isdir(i):
			temp_list = os.listdir('./%s' % i)
			for j in temp_list:
				print('正在检查%s' % j)
				if j.endswith('.xls') or j.endswith('xlsx'):
					result.append(os.getcwd() + '\\' + i + '\\' + j)
		else:
			if i.endswith('.xls') or i.endswith('xlsx'):
				result.append(os.getcwd() + '\\' + i)
	return result


def merge_excel(rows, files):
	'''
	合并所有excel文件
	:param rows: 标题行数
	:param files: 需要合并的文件列表
	:return: 无
	'''
	rows = int(rows)
	file_count = 0
	new_rows = -1
	w_book = xlwt.Workbook()
	w_sheet = w_book.add_sheet('合并')
	for i in files:
		print('正在 处理', i)
		if os.path.basename(i)[:2] != '~$':

			file_count = file_count + 1
			r_book = xlrd.open_workbook(i)
			r_sheet = r_book.sheet_by_index(0)
			this_rows = r_sheet.nrows
			this_cols = r_sheet.ncols
			if file_count == 1:
				# 如果是第一个文件则从头开始
				start = 0
			else:
				# 如果不是第一个文件，则从标题行后开始。
				start = rows
			for j in range(start, this_rows):
				new_rows = new_rows + 1
				for k in range(this_cols):
					# print(new_rows, ' ', j, ' ', k, ' ', r_sheet.cell(j, k).value)
					w_sheet.write(new_rows, k, r_sheet.cell(j, k).value)
				print('%s 文件共%s行，已完成合并' % (i, this_rows - rows))

	w_book.save('合并.xls')


def rm_result():
	'''
	删除输出文件防止打开输出文件异常。
	:return: 无
	'''
	if os.path.isfile('合并.xls'):
		os.remove('合并.xls')


def bye():
	print('感谢您使用excel自动合并程序。')
	input('按任意键退出程序')
	os._exit()


def main():
	'''
	主运行函数
	:return: 无
	'''
	welcome()
	rm_result()
	rows = choice_exception()
	file_list = read_list()
	merge_excel(rows, file_list)
	bye()


if __name__ == '__main__':
	main()
