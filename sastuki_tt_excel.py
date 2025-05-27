import openpyxl
import os
import pandas as pd

# 計測結果からタイムを計算する

def calculate_time(time):
    # HHMMSS形式の時間を秒に変換する

    time = str(time).zfill(6)
    if not time.isdigit() or len(time) != 6:
        print("ERROR")
        workbook.save(path)
        exit(1)

    hours= int(time[-6] + time[-5])
    minuites = int(time[-4] + time[-3])
    seconds = int(time[-2] + time[-1])
    # print(f'{hours}, {minuites}, {seconds}')
    return (hours*3600 + minuites * 60 + seconds)

def second_to_minuite(time):
    minuites = time // 60
    seconds = time % 60
    return [minuites, seconds]

file_name = 'data.xlsx' # 編集するファイル名 
script_dir = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(script_dir, file_name)

try:
    # Excelファイルを読み込む
    workbook = openpyxl.load_workbook(path)

    # ワークシートを選択する
    worksheet = workbook.active

    # はじめと終わりを設定
    row_start = 2 
    row_end = 69 # 自動で判定してくれる

    for row in range(row_start, row_end):
        if worksheet[f"A{row}"].value == None:
            break

        b_cell = worksheet[f"B{row}"]
        c_cell = worksheet[f"C{row}"]
        d_cell = worksheet[f"D{row}"]
        e_cell = worksheet[f"E{row}"]
        f_cell = worksheet[f"F{row}"]

        # ゴールタイムがない場合の処理
        if d_cell.value == None:
            d_cell.value = 245959 # 一日の最後にゴールしたとする

        # 合計タイム(秒)を計算
        sum_time = calculate_time(d_cell.value) - calculate_time(c_cell.value)

        # セルEにタイム(秒)を書き込み
        e_cell.value = sum_time

        # タイム(分秒を計算)
        result = second_to_minuite(sum_time)
        result_str = f"{result[0]}分{result[1]}秒"

        # セルFにタイム(分秒)を書き込み
        f_cell.value = result_str

    # ファイルを保存
    workbook.save(path)

    # 結果を昇順にソートする    
    # pandasに切り替える
    df = pd.read_excel(path, sheet_name = 'Sheet1')
    df
    df.sort_values(by='time(second)', ascending=True, inplace=True)
    # インデックスをリセット
    df.reset_index(drop=True, inplace=True)
    df.index = df.index + 1

    # スタートとゴールの時間とタイム(秒)の列を削除
    df.drop(['start(HHMMSS)','goal(HHMMSS)', 'time(second)'],axis=1, inplace=True)

    print(df) 

    result_file_name = 'result.xlsx'
    path = os.path.join(script_dir, result_file_name)
    df.to_excel(path)

    
except FileNotFoundError:
    print("指定したファイルが見つかりません。ファイル名とパスを確認してください。")