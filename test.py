import pandas as pd

data = pd.read_excel('DX_데이터 다루기.xlsx', sheet_name=None)
result = pd.DataFrame()
keys = data.keys()

for key in keys:
	"""안쓰는 데이터 의 열 0 3 6 7 8 10
	
	번호	문의유형	답변자	상태	기수	트랙/코스	본부	반	교육생	ID	교육상태	문의내용	문의일시	답변일시	답변	검토담당자	검토일시	
	["번호", "상태", "본부", "반", "교육생", "검토담당자", "검토일시"],
	"""
	print(data[key].drop(
		axis=1
	))
	print(data[key])
	# excel = data[key].drop([data[key][]], axis=1)