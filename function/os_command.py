import os
import glob

def make_directory(directory):
    if os.path.exists('./' + directory):
        print(directory + ' 폴더가 있으므로 생성하지 않습니다.')
    else :
        print(directory + ' 폴더를 생성합니다.')
        os.mkdir(directory)
        print(directory + '폴더 생성 완료.')

def check_file_name(file_name):
    file = file_name + '.xlsx'
    if os.path.exists(file):
        return file
    else:
        print(file + '이 존재하지 않습니다.')

def get_xlsx_file():
    # 현재 디렉토리의 모든 파일을 가져옵니다.
    current_directory = os.getcwd()

    # .xlsx 확장자를 가진 모든 파일을 찾습니다.
    xlsx_files = glob.glob(os.path.join(current_directory, "*.xlsx"))
    if len(xlsx_files) == 1:
        return xlsx_files[0]
    else:
        print("xlsx파일이 한개가 아닙니다. 폴더에 xlsx파일이 하나만 존재하도록 하고 다시 실행해주세요.")
