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
	ai_data = None
	dx_data = None
	for excel in kwargs['data']:
		excel_read = pd.read_excel(excel, header=1, sheet_name=None)
		excel_keys = excel_read.keys()
		tmpvalue = excel.replace(f"{os.getcwd()}/excel/", '').replace('.xlsx', '')
		value = tmpvalue.split('_')
		print(value)
		for key in excel_keys:
			excel_read[key] = excel_read[key].drop(columns=remove_header, axis=1)
			excel_read[key].rename(columns={'답변': '튜터 답변'}, inplace=True)
			excel_read[key].insert(
				loc=4,
				column='과목명',
				value=value[1],
			)
			excel_read[key]['챗GPT 문의내용'] = ""
			excel_read[key]['챗GPT답변'] = ""
			excel_read[key]['답변 사용 가능 여부'] = ""
			excel_read[key]['비고'] = ""
			if value[0] == "AI":
				if ai_data is None:
					ai_data = excel_read[key]
				else:
					ai_data = pd.concat([ai_data, excel_read[key]])
			elif value[0] == "DX":
				if dx_data is None:
					dx_data = excel_read[key]
				else:
					dx_data = pd.concat([dx_data, excel_read[key]])

	with pd.ExcelWriter('dist/combine.xlsx') as writer:
		ai_data.reset_index(drop=True, inplace=True)
		ai_data.drop(index=0, axis=1)
		ai_data.insert(
			loc=0,
			column='번호',
			value=ai_data.index,
		)
		ai_data.to_excel(writer, sheet_name='AI', index=False)
		dx_data.reset_index(drop=True, inplace=True)
		dx_data.drop(index=0, axis=1)
		dx_data.insert(
			loc=0,
			column='번호',
			value=ai_data.index,
		)
		dx_data.to_excel(writer, sheet_name='DX', index=False)
	# data.to_excel(
	# 	os.getcwd() + '/dist/combined.xlsx',
	# 	engine="openpyxl",
	# 	index=False,
	# )


def main():
	readExcel()


main()
