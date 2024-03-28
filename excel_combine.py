import os
import pandas as pd
from openFolder import scanDir


remove_header = ['번호', '상태', '본부', '반', '교육생', '교육상태', '검토담당자', '검토일시']


def openExcel(func):
	def wrapper(*arg, **kwargs):
		kwargs['data'] = scanDir()
		result = func(**kwargs)
		return result
	return wrapper


@openExcel
def readExcel(**kwargs):
	data = None
	print(kwargs['data'])
	for excel in kwargs['data']:
		excel_read = pd.read_excel(excel, header=1, sheet_name=None)
		excel_keys = excel_read.keys()
		value = excel.replace(f"{os.getcwd()}/excel/", '').replace('.xlsx', '')
		print(value)
		for key in excel_keys:
			excel_read[key] = excel_read[key].drop(columns=remove_header, axis=1)
			excel_read[key].rename(columns={'답변': '튜터 답변'})
			excel_read[key].insert(
				loc=4,
				column='과목명',
				value=value,
			)
			excel_read[key]['챗GPT 문의내용'] = ""
			excel_read[key]['챗GPT답변'] = ""
			excel_read[key]['답변 사용 가능 여부'] = ""
			excel_read[key]['비고'] = ""
			if data is None:
				data = excel_read[key]
			else:
				data = pd.concat([data, excel_read[key]])
	data.to_excel(
		os.getcwd() + '/dist/combined.xlsx',
		engine="openpyxl",
	)


def main():
	readExcel()


main()
