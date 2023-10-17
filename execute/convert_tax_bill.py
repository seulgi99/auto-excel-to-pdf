import win32com.client as win32
import openpyxl
from openpyxl import Workbook
from function import print_error, func_excel
import os,time

def execute(file, data_set):
    # data set 가져오기
    set_list = get_data_set_info(data_set)

    data_set_info = set_list[0]
    industry_academic_set_info = set_list[1]

    # 시트별 pdf변환
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        work_book = excel.Workbooks.Open(file)
        # 시트별로 루프
        for sheet in work_book.Sheets:
            print(f'{sheet.Name}시트 작업')

            # 상호명 뽑아오기 (Y4 위치)
            business_name =sheet.Range("Y4").Value
            print(f'상호명 : {business_name}')
            file_name = ""
            if business_name != "홍익대학교 산학협력단":
                for data in data_set_info:
                    if business_name == data[0]:
                        file_name = data[1] + " " + data[2] + "월"
            else:
                # 금액 뽑아오기 (B17 위치)
                price = sheet.Range("B17").Value
                for data in industry_academic_set_info:
                    if price == data[0]:
                        file_name = data[1] + " " + data[2] + "월"

            if file_name != "":
                # PDF로 저장
                new_pdf_name = './pdf/' + file_name + '.pdf'
                while(True):
                    if(os.path.exists(new_pdf_name)):
                        new_pdf_name = new_pdf_name[:-4] + "-중복.pdf"
                    else:
                        break

                sheet.ExportAsFixedFormat(0, os.path.abspath(new_pdf_name))
                print(f'{new_pdf_name} 저장 완료')
            else:
                work_book.Close(SaveChanges=False)
                excel.Quit()
                print_error.execute("data set에 매칭되는 내용이 없습니다.")

        work_book.Close(SaveChanges=False)
        excel.Quit()
        print('작업이 완료되었습니다.')

    except Exception as e:
        work_book.Close(SaveChanges=False)
        excel.Quit()
        print_error.execute("에러 발생: " + str(e))


def get_data_set_info(data_set):
    # data set 가져오기
    workbook = openpyxl.load_workbook(data_set)

    data_set_sheet = workbook["일반 데이터"]
    industry_academic_set_sheet = workbook["산학협력단"]

    end_row = func_excel.check_line(data_set_sheet)  # 데이터가 존재하는 마지막 열
    data_set_info = []
    for i in range(end_row - 2):
        line = []
        row = i + 3
        line.append(str(data_set_sheet[f'A{row}'].value))
        line.append(str(data_set_sheet[f'B{row}'].value))
        line.append(str(data_set_sheet[f'C{row}'].value))
        data_set_info.append(line)

    end_row = func_excel.check_line(industry_academic_set_sheet)  # 데이터가 존재하는 마지막 열
    industry_academic_set_info = []
    for i in range(end_row - 2):
        line = []
        row = i + 3
        line.append(str(industry_academic_set_sheet[f'A{row}'].value))
        line.append(str(industry_academic_set_sheet[f'B{row}'].value))
        line.append(str(industry_academic_set_sheet[f'C{row}'].value))
        industry_academic_set_info.append(line)
    workbook.close()

    set_list = [data_set_info, industry_academic_set_info]
    return set_list