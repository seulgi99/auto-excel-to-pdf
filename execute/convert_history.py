import openpyxl
from openpyxl import Workbook
from function import func_excel, func_pdf

def execute(file) :
    workbook = openpyxl.load_workbook(file)
    source_sheet = workbook.active

    end_row = func_excel.check_line(source_sheet) #데이터가 존재하는 마지막 열

    # I열에서 가장 큰수 찾기
    max_index = func_excel.check_max_index(source_sheet,'J', 8, end_row)
    workbook.close()
    for i in range(1, max_index + 1):
        print(f'{i}번째 파일 작업')
        # 원본 엑셀 파일 열기
        workbook = openpyxl.load_workbook(file)
        source_sheet = workbook.active

        # 색칠해야하는 열 타게팅
        target_row = func_excel.get_target_row_list(source_sheet, i, 'J', 8, end_row)

        # 새로운 엑셀에 색칠하기(노란색)
        for row in target_row:
            func_excel.color_cell_yellow(source_sheet, row, 'A')
            func_excel.color_cell_yellow(source_sheet, row, 'E')
            func_excel.color_cell_yellow(source_sheet, row, 'F')

        # 파일 이름 지정
        name_cell = 'K' + str(target_row[0])
        new_excel_name = './excel/' + source_sheet[name_cell].value + '.xlsx'
        new_pdf_name = './pdf/' + source_sheet[name_cell].value + '.pdf'
        # 필요없는 데이터 지우기 ( index 및 파일명)
        func_excel.delete_column(source_sheet, 'J', 8, end_row)
        func_excel.delete_column(source_sheet, 'K', 8, end_row)


        # 새로운 엑셀 파일 저장
        workbook.save(new_excel_name)
        print(new_excel_name + '파일 생성 완료')
        # 엑셀 파일 닫기
        workbook.close()

        func_pdf.excel_to_pdf(new_excel_name, new_pdf_name)