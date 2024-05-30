import openpyxl
from openpyxl import load_workbook
from function import func_excel, func_pdf

import re


# file : 그래프를 그리는 파일, data_set : 정보를 가져오는 파일
def execute(file, data_set):
    # data set 가져오기
    data = func_excel.get_history_data_set_info(data_set)

    # 그래프의 업체명 모두 가져오기
    institutions = func_excel.get_column_data(file, 'B')

    # data_set에서 뽑아온 결의 마다 루프 (resolution : 결의)
    for resolution in data:
        res_title = resolution[4]  # dataset의 업체명
        insert_result = "" # 이걸 그래프의 n.n입 으로쓸거임

        row = None
        for idx, institution in enumerate(institutions):
            if institution in res_title:
                row = idx
                break

        func_excel.insert_value_to_first_empty_cell(file, row, )


def extract_and_compute_difference(text):
    # 정규 표현식 패턴
    patterns = [
        r'(\d{1,2}월분)',  # n월분 형태
        r'(\d{1,2}~\d{1,2}월)'  # n~m월 형태
    ]

    numbers = []

    for pattern in patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            # n~m월분 형태인 경우, 시작과 끝 숫자를 분리
            if '~' in match:
                range_parts = re.findall(r'\d+', match)
                numbers.extend(map(int, range_parts))
            # n월분과 n월 형태인 경우, 숫자만 추출
            else:
                month_parts = re.findall(r'\d+', match)
                numbers.extend(map(int, month_parts))

    # 숫자가 여러 개일 경우, 차이를 계산
    if len(numbers) > 1:
        difference = max(numbers) - min(numbers)
    else:
        difference = 0  # 숫자가 하나거나 없으면 차이는 0으로 설정

    return numbers, difference


def insert_value_to_first_empty_cell(filename, row, text):
    # 엑셀 파일 열기
    workbook = load_workbook(filename)
    sheet = workbook.active  # 활성 시트 선택

    # extract_and_compute_difference 함수 호출
    _, difference = extract_and_compute_difference(text)

    # 차이에 따른 칸 수와 반환 값 결정
    if difference == 0:
        cells_to_merge = 1
        return_value = "월"
    elif difference == 3:
        cells_to_merge = 3
        return_value = "분기"
    elif difference == 6:
        cells_to_merge = 6
        return_value = "반기"
    elif difference == 12:
        cells_to_merge = 12
        return_value = "연"
    else:
        cells_to_merge = 1
        return_value = "월"

    # 주어진 행(row)에서 가장 처음으로 비어있는 열 찾기
    column = 1  # 열 인덱스
    while True:
        cell = sheet.cell(row, column)
        if cell.value is None:
            # 필요한 칸 수만큼 셀 병합
            sheet.merge_cells(row, column, row, column + cells_to_merge - 1)
            sheet.cell(row, column).value = text
            break
        column += 1

    # 변경 내용을 저장하고 파일 닫기
    workbook.save(filename)
    workbook.close()

    return return_value