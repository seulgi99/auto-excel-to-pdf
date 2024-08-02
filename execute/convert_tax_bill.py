import win32com.client as win32
import openpyxl
from openpyxl import Workbook
from function import print_error, func_excel
import os,time

def execute(file, data_set):
    # data set 가져오기
    data_set_info = get_data_set_info(data_set)

    # 시트별 pdf변환
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True


    try:
        work_book = excel.Workbooks.Open(file)

        # 시트별로 루프
        for sheet in work_book.Sheets:
            print(f'{sheet.Name}시트 작업')
            new_workbook = excel.Workbooks.Add()
            new_sheet = new_workbook.Sheets(1)

            # 시트 복사
            sheet.Copy(Before=new_sheet)
            new_sheet = new_workbook.Sheets(1)
            # 새로운 워크북에서 "Sheet1" 삭제
            for ws in new_workbook.Sheets:
                if ws.Name == "Sheet1":
                    ws.Delete()

            # 상호명 뽑아오기 (Y4 위치)
            business_name =new_sheet.Range("X4").Value
            print(f'상호명 : {business_name}')
            # 금액 뽑아오기 (B17 위치)
            price = str(int(new_sheet.Range("B17").Value))
            # 이메일 뽑아오기 (Y17 위치)
            email = str(int(new_sheet.Range("Y17").Value))

            file_name = ""
            # 산학협력일 시 email 비교
            if business_name == "홍익대학교 산학협력단":
                for data in data_set_info:
                    if email == data[2]:
                        file_name = data[3] + " " + data[4] + "월"

            else:
                for data in data_set_info:
                    if business_name == data[0] and price == data[1]:
                        file_name = data[3] + " " + data[4] + "월"


            if file_name == "":
                file_name = "기타"
            # excel로 저장할 이름
            new_excel_name = './excel/' + file_name + '.xlsx'
            while (True):
                if (os.path.exists(new_excel_name)):
                    file_name = file_name + '-중복'
                    new_excel_name = './excel/' + file_name + '.xlsx'
                else:
                    break
            # PDF로 저장할 이름
            new_pdf_name = './pdf/' + file_name + '.pdf'
            new_sheet.SaveAs(os.path.abspath(new_excel_name))
            print(f'{new_excel_name} 저장 완료')
            new_workbook.ExportAsFixedFormat(0, os.path.abspath(new_pdf_name))
            print(f'{new_pdf_name} 저장 완료')
            print()
            new_workbook.Close(SaveChanges=False)


        work_book.Close(SaveChanges=False)
        excel.Quit()
        input('작업이 완료되었습니다. 엔터를 누르면 메뉴로 돌아갑니다.')

    except Exception as e:
        new_workbook.Close(SaveChanges=False)
        work_book.Close(SaveChanges=False)
        excel.Quit()
        print_error.execute("에러 발생: " + str(e))


def get_data_set_info(data_set):
    # data set 가져오기
    workbook = openpyxl.load_workbook(data_set)

    data_set_sheet = workbook["일반 데이터"]

    end_row = func_excel.check_line(data_set_sheet)  # 데이터가 존재하는 마지막 열
    data_set_info = []
    for i in range(end_row - 2):
        line = []
        row = i + 3
        line.append(str(data_set_sheet[f'A{row}'].value))  # 상호명 data[0]
        line.append(str(data_set_sheet[f'B{row}'].value))  # 금액   data[1]
        line.append(str(data_set_sheet[f'C{row}'].value))  # 이메일 data[2]
        line.append(str(data_set_sheet[f'D{row}'].value))  # 생성할 파일 명 data[3]
        line.append(str(data_set_sheet[f'E{row}'].value))  # 월    data[4]
        data_set_info.append(line)

    workbook.close()
    return data_set_info