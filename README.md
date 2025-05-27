# sastuki_tt_excelの使い方

　このプログラムはOUCC(大阪大学サイクリング部ツアー班)の活動の1つである「五月TT」の結果を計算するために作成された．
sastuki_tt_excel.pyはExcel上のスタートとゴール時刻から所要時間を計算し，ソートした結果のExcelファイルを出力する．
sastuki_tt_excel_start_input.pyはExcelのデータに開始時刻を一括設定するためのもの補助プログラムである．

# 環境
- Python 3.x
- openpyxl
- pandas

openpyxlはExcelの読み書き，pandasはソート・出力に利用している．

# 入力ファイルの準備
スクリプトと同じフォルダにdata.xlsxという名前で以下のような列構成のシートを用意する．
時間の入力補助プログラムについては下部の「データの入力補助プログラム(sastuki_tt_excel_start_input.py)について」で説明する．

A列 | B列 | C 列(start) | D列(goal) | E列 |F列
--- | --- | --- | --- | --- | ---
回生 | 名前| HHMMSS | HHMMSS | 所要時間(秒) | 所要時間(分)

C列に開始時刻，D列にゴール時刻をHHMMSS形式で入力する．
(例:113042=11時30分42秒)
D列が空の場合は「24:59:59」として扱われる．
E，F列は自動で計算されるので空欄にしておく．

プログラム(sastuki__tt_excel.py)の46行目のrow_startでデータの始まる行数を設定する．
デフォルトでは2行目から100行目までを読み込むようにしている．
途中で空白行があった場合はその行まで読み込むため，行数は多めに設定している．
100人以上のデータがある場合は，40行目のrow_endを変更すること．
```py:sastuki_tt_excel.py
38    # はじめと終わりを設定
39    row_start = 2 
40    row_end = 100 # 自動で判定してくれる
```

また，37行目の'file_name'の右辺値(デフォルトでは'data.xlsx)を変更することで，読み込むファイル名を変更できる．

# 実行方法
コマンドラインから以下のように実行する．
```bash
python sastuki_tt_excel.py
```
実行中にエラーがなければ，まずdata.xlsxの同じシートに以下がかき込まれる．
-E列: 所要時間(秒)
-F列: 所要時間(分秒表示，例12分34秒)

# 出力
コマンドライン上にタイムが短い順にソートされた結果が表示される．
また，同じディレクトリにresult.xlsxが作成される．
このファイルは先ほどコマンドラインに出力された結果と同じ内容である．
このExcelファイルをエクスポートすることで，結果のpdfを作成することができる．

# データの入力補助プログラム(sastuki_tt_excel_start_input.py)について

このプログラムはExcelのデータに開始時刻を一括設定するためのものである．

# 設定可能なパラメータ

設定可能なパラメータは以下の通りである．
- row_start: データの始まる行数
- row_end: データの終わる行数
- time_start: 開始時刻の初期値(例: 110000=11時00分00秒)
- dist_time: 各グループ間の出発時間の間隔(秒)
- people_number: 各グループの人数

~~~py:sastuki_tt_excel_start_input.py
45    # はじめと終わりを設定
46    row_start = 2 
47    row_end = 100
48
49    # パラメータ
50    time_start = 110000
51    dist_time = 60
52    people_number = 2
~~~

# 高度な設定

62行目のf"C{row}"の"C"の部分を変更することで，開始時刻を入力する列を変更できる．
(例:worksheet[f"D{row}"]とするとD列に時刻が入力される)
プログラムのシミュレーションを行う場合などに，ゴール時刻を入力することができる．
~~~py:sastuki_tt_excel_start_input.py
61            # 編集する列を選択
62            c_cell = worksheet[f"C{row}"]
63            time_start = time_check(time_start)
64            c_cell.value = time_start
65            row += 1
~~~

# 実行方法
コマンドラインから以下のように実行する．
```bash
python sastuki_tt_excel.py
```
実行中にエラーがなければ，まずdata.xlsxの同じシートのC列に時刻が書き込まれている．