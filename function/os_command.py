import os

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