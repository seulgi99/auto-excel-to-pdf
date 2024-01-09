import openpyxl
from openpyxl import Workbook
from function import func_excel, func_pdf



def execute(file, data_set) :
    # data set 가져오기
    history_data_set_info = get_history_data_set_info(data_set)

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

        # 이름 지정
        name_cell = 'K' + str(target_row[0])
        # 월 수 지정
        month_cell = 'L' + str(target_row[0])
        # 구분(입금/출금) 지정
        type_cell = 'M' + str(target_row[0])
        # 이름 + 월 + 구분으로 파일 이름 지정
        file_name = source_sheet[name_cell].value + " " +source_sheet[month_cell].value + " " + source_sheet[type_cell].value

        # 금액모아서 data set 과 비교
        total_money = 0
        for row in target_row:
            total_money += int(source_sheet["E" + row].value)

        for data in history_data_set_info:
            if source_sheet[name_cell].value == data[0] and total_money == int(data[1]):
                break
            if data == history_data_set_info[-1]:
                file_name +=" - 불일치"


        new_excel_name = './excel/' + file_name + '.xlsx'
        new_pdf_name = './pdf/' + file_name + '.pdf'
        # 필요없는 데이터 지우기 ( index 및 파일명,월,구분)
        func_excel.delete_column(source_sheet, 'J', 8, end_row)
        func_excel.delete_column(source_sheet, 'K', 8, end_row)
        func_excel.delete_column(source_sheet, 'L', 8, end_row)
        func_excel.delete_column(source_sheet, 'M', 8, end_row)


        # 새로운 엑셀 파일 저장
        workbook.save(new_excel_name)
        print(new_excel_name + '파일 생성 완료')
        # 엑셀 파일 닫기
        workbook.close()

        func_pdf.excel_to_pdf(new_excel_name, new_pdf_name)

    func_pdf.merge_pdf()
    input('작업이 완료되었습니다. 엔터를 누르면 종료합니다.')
    exit(0)