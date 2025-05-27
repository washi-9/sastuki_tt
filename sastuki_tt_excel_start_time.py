import openpyxl
import os

# 出発時間を設定する

def time_check(time):
    time = str(time).zfill(6)
    if not time.isdigit() or len(time) != 6:
        print("ERROR")
        workbook.save(path)
        exit(1)

    hours= int(time[-6] + time[-5])
    minuites = int(time[-4] + time[-3])
    seconds = int(time[-2] + time[-1])

    # seconds, minuites, hoursが正しい範囲内にあるか確認
    if (seconds >= 60):
        minuites += 1
        seconds -= 60
    
    if (minuites >= 60):
        hours += 1
        minuites -= 60

    if (hours >= 24):
        hours -= 24
    
    result = hours*10**4 + minuites*10**2 + seconds

    if (len(str(result)) == 6):
        return result
    
    else: # hoursが一桁の場合
        return "0" + str(result)

file_name = 'data.xlsx' # 編集するファイル名
script_dir = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(script_dir, file_name)

try:
    workbook = openpyxl.load_workbook(path)
    worksheet = workbook.active

    # はじめと終わりを設定
    row_start = 2 
    row_end = 100

    # パラメータ
    time_start = 110000
    dist_time = 60
    people_number = 2

    for row in range(row_start, row_end, people_number):
        if worksheet[f"A{row}"].value == None:
            break
        for i in range(0, people_number):
            if worksheet[f"A{row}"].value == None:
                break

            # 編集する列を選択
            c_cell = worksheet[f"C{row}"]
            time_start = time_check(time_start)
            c_cell.value = time_start
            row += 1

        time_start += dist_time

    workbook.save(path)

except FileNotFoundError:
    print("指定したファイルが見つかりません。ファイル名とパスを確認してください。")    
