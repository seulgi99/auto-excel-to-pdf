import os
import win32com.client as win32

def excel_to_pdf(excel_file, pdf_file):
    # Excel 인스턴스 생성
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # Excel 창을 표시하지 않음

    try:
        # Excel 파일 열기
        workbook = excel.Workbooks.Open(os.path.abspath(excel_file))

        # PDF로 저장
        workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_file))

        # Excel 파일 닫기
        workbook.Close()

    except Exception as e:
        print(f"에러 발생: {e}")
    finally:
        # Excel 인스턴스 종료
        excel.Quit()