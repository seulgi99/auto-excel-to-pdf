import os
import glob
import win32com.client as win32
from PyPDF2 import PdfMerger
from glob import glob

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
        
        
def merge_pdf():
    try:
        merger = PdfMerger()
        current_directory = os.getcwd()
        # .pdf 확장자를 가진 모든 파일을 찾습니다.

        pdf_files = glob(os.path.join(current_directory, "pdf/*.pdf"))

        for file in pdf_files:
            merger.append(file)

        merger.write('./병합본.pdf')
        merger.close()

    except Exception as e:
        print(2)
        merger.close()
        print(f"에러 발생: {e}")