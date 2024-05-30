from openpyxl.styles import PatternFill
from openpyxl import load_workbook


def check_line(sheet):
    max_row = sheet.max_row
    max_row = int(max_row)
    print(f"데이터가 존재하는 행의 수: {max_row}")
    return max_row


# 거래내역 조회서 자동화에서만 쓰는 것임
def check_max_index(sheet, column, start_row, end_row):
    max_index = None

    # 데이터 탐색 및 최대값 찾기
    for ro in range(start_row, end_row + 1):
        cell_value = sheet[f'{column}{row}'].value
        if cell_value is not None:
            if max_index is None or cell_value > max_index:
                max_index = cell_value
        else:
            raise Exception(f'{column}{row}이 비었거나 숫자가 아닙니다.')

    print(f'최종 분할되어 나올 pdf의 개수: {max_index}')
    return max_index


# 거래내역 조회서 자동화에서만 쓰는 것임
def get_target_row_list(sheet, index, column, start_row, end_row):
    target_row_list = []

    # 데이터 탐색 및 해당 인덱스가 들어가있는 열 찾기
    for row in range(start_row, end_row + 1):
        cell_value = sheet[f'{column}{row}'].value
        cell_value = int(cell_value)
        if cell_value == index:
            target_row_list.append(row)
    if not target_row_list:
        raise Exception(f'{index}번이 들어있지 않습니다.')
    return target_row_list


def color_cell_yellow(sheet, row, column):
    cell = sheet[column + str(row)]
    fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # 노란색 배경
    # 셀에 배경 색상 적용
    cell.fill = fill


def delete_column(sheet, column, start_row, end_row):
    for row in range(start_row, end_row + 1):
        cell = column + str(row)
        sheet[cell].value = None


def get_history_data_set_info(data_set):
    # 엑셀 파일 열기
    workbook = load_workbook(data_set)
    sheet = workbook.active  # 활성 시트 선택

    # 데이터 저장을 위한 리스트
    data = []

    # 모든 행을 반복하면서 A부터 AG열의 값을 가져오기
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=33):  # A부터 AG까지 33개의 열
        row_data = []
        for cell in row:
            row_data.append(cell.value)
        data.append(row_data)

    return data


def get_column_data(filename, column):
    # 엑셀 파일 열기
    workbook = load_workbook(filename)
    sheet = workbook.active  # 활성 시트 선택

    # 데이터 저장을 위한 리스트
    data = []

    # 특정 열의 데이터 가져오기
    for row in range(2, 69):  # 2행부터 68행까지
        cell_value = sheet[f'{column}{row}'].value
        data.append(cell_value)

    return data


def insert_value_to_first_empty_cell(filename, row, value):
    # 엑셀 파일 열기
    workbook = load_workbook(filename)
    sheet = workbook.active  # 활성 시트 선택

    # 주어진 행(row)에서 가장 처음으로 비어있는 열 찾기
    column = 1  # 열 인덱스
    while True:
        cell = sheet.cell(row, column)
        if cell.value is None:
            cell.value = value
            break
        column += 1

    # 변경 내용을 저장하고 파일 닫기
    workbook.save(filename)
    workbook.close()
