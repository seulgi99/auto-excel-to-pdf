from datetime import datetime

import openpyxl
from openpyxl import load_workbook
from function import func_excel, func_pdf

import re


# file : 그래프를 그리는 파일, data_set : 정보를 가져오는 파일
def execute(file, data_set):
    # data set 가져오기 (엑셀 파일 형식의 2차원배열)
    data = func_excel.get_history_data_set_info(data_set)

    # 그래프의 검색어 모두 가져오기
    keywords = func_excel.get_column_data(file, 'C')
    filtered_keywords = [x for x in keywords if x is not None]
    for keyword in filtered_keywords:
        print(keyword)

    # 그래프의 금액 모두 가져오기
    money_values = func_excel.get_column_data(file, 'F')

    # data_set에서 뽑아온 결의 한 행 마다 루프 (resolution : 결의)
    for resolution in data:
        res_title = resolution[4]  # dataset의 결의서 제목
        res_money_value = resolution[7]  # dataset의 차변

        print("res_title: ", res_title)
        print("res_money_value: ", res_money_value)

        row = None
        for idx, keyword in enumerate(filtered_keywords):  # 검색어 루프
            for money_value in money_values:  # 금액 루프
                if keyword in res_title and money_value == res_money_value:  # 검색어가 data_set의 결의서 제목에 포함되고, 금액까지 같다면?
                    row = idx
                    print("전표일자(resolution[2]) : ", resolution[2])
                    insert_result = format_date(resolution[2])  # resolution[2] : dataset의 전표일자, insert_result : n.n입
                    print("insert result : ", insert_result)
                    insert_value_to_merged_cell(file, row, res_title, resolution[2], insert_result)


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
            # n월분 형태인 경우, 숫자만 추출
            else:
                month_parts = re.findall(r'\d+', match)
                numbers.extend(map(int, month_parts))

    # 숫자가 여러 개일 경우, 차이를 계산
    if len(numbers) > 1:
        difference = max(numbers) - min(numbers)
    else:
        difference = 0  # 숫자가 하나거나 없으면 차이는 0으로 설정

    return numbers, difference


# insert_result: n.n입, resolution_date: dataset의 전표일자, text: dataset의 결의서제목
def insert_value_to_merged_cell(filename, row, text, resolution_date, insert_result):
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
        return_value = "년"
    else:
        cells_to_merge = 1
        return_value = "월"

    print("cells_to_merge", cells_to_merge)

    # 날짜에 맞는 열 찾기
    date_str = resolution_date.strftime('%Y-%m-%d')
    print("date_str : ", date_str)

    column = func_excel.get_column_from_date(sheet, date_str)

    print("cell : ", cell_address(row, column))

    cell = sheet.cell(row=row, column=column)

    print("insert_result : ", insert_result)

    # 지정된 행(row)과 찾은 열(column)에 값 입력
    if cell.value is None:  # 셀이 비어있다면

        # 필요한 칸 수만큼 셀 병합
        if cells_to_merge > 1:
            sheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + cells_to_merge - 1)

        sheet.cell(row=row, column=column).value = insert_result

        # 변경 내용을 저장
        workbook.save(filename)

    # 파일 닫기
    workbook.close()

    return return_value


def cell_address(row, column):
    # 열 번호를 열 문자로 변환
    column_letter = openpyxl.utils.get_column_letter(column)
    # 셀 주소 반환
    return f"{column_letter}{row}"


def format_date(date_str):
    print("date_str : ", date_str)
    # 문자열을 datetime 객체로 변환
    # date_obj = datetime.strptime(date_str, '%Y-%m-%d')

    # 연, 월, 일을 각각 추출
    year = date_str.year
    month = date_str.strftime('%m')
    day = date_str.strftime('%d')

    # 변환된 문자열을 반환
    formatted_date = f"{month}.{day}입"
    return formatted_date
