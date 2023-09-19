from function import os_command
from execute import  transaction_history

if __name__ == '__main__':
    os_command.make_directory('excel')
    os_command.make_directory('pdf')
    while True:
        print('작업종류')
        print('1: 거래내역조회서 pdf변환 자동화')
        execute_type = input('작업 종류를 입력해주세요: ')

        if execute_type == '1':
            pass
        else:
            print('잘못된 입력 입니다. 다시 입력해주세요.')
            continue

        print('이 프로그램은 xlsx파일에만 사용이 가능합니다. xls파일이라면 변환해주세요.')
        # 엑셀 파일의 이름을 가져옵니다.
        file = os_command.get_xlsx_file()

        if file == "":
            continue

        if execute_type == '1':
            transaction_history.execute(file)


