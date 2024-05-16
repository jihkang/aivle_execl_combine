import os
import pandas as pd
import warnings
from openFolder import scanDir
from datetime import datetime
from openpyxl.styles import Alignment, Font
import openpyxl
import openai


warnings.filterwarnings('ignore')
remove_header = ['번호', '상태', '본부', '반', '교육생', '교육상태', '검토담당자', '검토일시']


def get_course(x):
	ind = pd.read_excel('3기 과정 목록.xlsx', header=1, sheet_name=None)
	for key in ind.keys():
		ind[key] = ind[key][['시작일자', '종료일자', '과정명']][:24]
	data = ind[(ind['시작일자'] <= x)]['과정명'].tolist()
	print(f"x : {x} 과목명")
	if data:
		print(f"data: {data[len(data) - 1]}")
		return data[len(data) - 1]
	else:
 	   return "없음"


def openExcel(func):
	def wrapper():
		result = func(scanDir())
		return result
	return wrapper


def writeExcel(data):
	with pd.ExcelWriter('dist/result.xlsx') as writer:
		for key in data.keys():
			if data[key] is None:
				continue
			data[key].reset_index(drop=True, inplace=True)
			data[key].drop(index=0, axis=1)
			data[key].insert(
				loc=0,
				column='번호',
				value=data[key].index,
			)
			data[key].to_excel(writer, sheet_name=key, index=False)



@openExcel
def readExcel(rdata):
	data = {
		'AI': None,
		'DX': None,
	}
	for excel in rdata:
		excel_read = pd.read_excel(
			excel,
			header=1,
			sheet_name=None,
			engine="openpyxl"
		)
		excel_keys = excel_read.keys()
		value = excel.replace(f"{os.getcwd()}/excel/", '').replace('.xlsx', '').split('_')
		for key in excel_keys:
			excel_read[key] = excel_read[key].drop(
				columns=remove_header,
				axis=1
			)
			excel_read[key].rename(columns={'답변': '튜터 답변'}, inplace=True)
			excel_read[key].insert(
				loc=4,
				column='과목명',
				value=value[1],
			)
			excel_read[key]['챗GPT 문의내용'] = excel_read[key]["문의내용"]
			excel_read[key]['챗GPT답변'] = excel_read[key]['답변 사용 가능 여부'] = excel_read[key]['비고'] = ""
			response = openai.Completion.create(
                    engine="gpt-3.5-turbo-instruct",
                    prompt=f"please answer each question: {excel_read[key]['챗GPT 문의내용']}",
                    max_tokens=1000
                )
			exce_read[key]['챗GPT 답변'] = response.choices[0].text.strip()
			if data[value[0]] is None:
				data[value[0]] = excel_read[key]
			else:
				data[value[0]] = pd.concat([data[value[0]], excel_read[key]])
	writeExcel(data)


def main():
	test = openpyxl.load_workbook('dist/result.xlsx')
	for i in test.sheetnames:
		for row in test[i]:
			for cell in row:      
				cell.alignment = Alignment(wrap_text=True)
	test.save('dist/result.xlsx')	

main()
