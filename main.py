from function import os_command
from execute import convert_history, convert_tax_bill, update_payment_graph

if __name__ == '__main__':
    os_command.make_directory('excel')
    os_command.make_directory('pdf')
    while True:
        print('이 프로그램은 xlsx파일에만 사용이 가능합니다. xls파일이라면 변환해주세요.')
        print('작업종류')
        print('1: 거래내역조회서 pdf변환 자동화')
        print('2: 전자세금계산서 pdf변환 자동화')
        execute_type = input('작업 종류를 입력해주세요: ')

        if execute_type == '1':
            file = os_command.get_xlsx_file()
            data_set = os_command.get_data_set_xlsx_file("dataset")
            if file != "" and data_set != "":
                convert_history.execute(file, data_set)

        elif execute_type == '2':
            file = os_command.get_xlsx_file()
            data_set = os_command.get_data_set_xlsx_file("dataset")
            if file != "" and data_set != "":
                convert_tax_bill.execute(file, data_set)

        # elif execute_type == '3':
        #     file = os_command.get_xlsx_file()
        #     data_set = os_command.get_data_set_xlsx_file("dataset")
        #     update_payment_graph.execute(file, data_set)
        else:
            print('해당하는 작업이 없습니다. 다시 입력해주세요.')
            continue
