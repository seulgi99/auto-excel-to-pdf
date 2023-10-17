from openpyxl.styles import PatternFill


def check_line(sheet):
    max_row = sheet.max_row
    max_row = int(max_row)
    print(f"데이터가 존재하는 행의 수: {max_row}")
    return max_row

# 거래내역 조회서 자동화에서만 쓰는 것임
def check_max_index(sheet, column, start_row, end_row):
    max_index = None

    # 데이터 탐색 및 최대값 찾기
    for row in range(start_row, end_row + 1):
        cell_value = sheet[f'{column}{row}'].value
        if cell_value is not None:
            if max_index is None or cell_value > max_index:
                max_index = cell_value
        else:
            print(f'{column}{row}이 비었거나 숫자가 아닙니다.')
            input('엔터를 입력하면 종료합니다.')
            exit()

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

