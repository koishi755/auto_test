import functions
import generate_sm37_htm
import pyautogui
import time

generate_sm37_htm.generate_job_list()

# ファイルを読み込み
t_code = 'sm37'
file_name = f'{t_code}_canceled_jobs'
data = functions.open_htm(file_name)

#　データを成型
def data_organize(data):
    tmp_data = data.split(r'<tbody>')
    tmp_data = tmp_data[2].split(r'</tr>')
    tmp_data.pop(-1)
    return tmp_data

# dataからjobnameを取り出す
def get_data(data):
    data_list = []
    for item in data:
        item = item.split(r'>')
        item = item[5].split(r'</')
        item = item[0].replace(r'&nbsp;', '')
        data_list.append(item)
    return data_list

data = data_organize(data)
job_name_list = get_data(data)

# cancelされたジョブのステータスを取得
def get_status(data):
    status_list = []
    for item in data:
        item = item.split(r'>')
        item = item[27].split(r'</')
        item = item[0].replace(r'&nbsp;', '')
        status_list.append(item)
    return status_list

# 重複した値を削除
job_name_list = list(dict.fromkeys(job_name_list))

# /nsm37を実行
def execute_sm37():
    # functions.get_window('Job Overview')
    functions.execute_t_code(t_code)
    functions.get_window('Simple Job Selection')

# Simple_Job_Selectionでの入力
def put_job_name_to_Simple_Job_Selection(job_name):
    pyautogui.typewrite(job_name)
    time.sleep(3)
    pyautogui.press('tab')
    pyautogui.typewrite(r'*')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('space')
    pyautogui.press('tab')
    pyautogui.press('space')
    pyautogui.press('tab')
    pyautogui.press('space')
    time.sleep(2)
    pyautogui.press('f8')

# Simple Job Selection画面に戻る
functions.get_window('Simple Job Selection')
execute_sm37()

# job_name_listの数だけhtmファイルを作成
for job_name in job_name_list:
    tmp_job_name = job_name.replace(r'/', ',')
    put_job_name_to_Simple_Job_Selection(job_name)
    functions.get_window('Job Overview')
    functions.move_to_localfile()
    functions.get_window('Save list in file...')
    functions.click_save_list_in_file()
    functions.get_window('Job Overview')
    file_name = f'{t_code}_{tmp_job_name}'
    functions.click_to_genarate(file_name)
    functions.get_window('Job Overview')
    execute_sm37()

# 最初の画面に戻る
functions.get_window('Simple Job Selection')
pyautogui.press(['f3'])

# job_name_listの数だけhtmファイルを取得
put_job = []
for job_name in job_name_list:
    tmp_job_name = job_name.replace(r'/', ',')
    file_name = f'{t_code}_{tmp_job_name}'
    canceled_job = functions.open_htm(file_name)
    canceled_job = data_organize(canceled_job)
    status_list = get_status(canceled_job)
    if status_list[0] == 'Canceled':
        put_job.append(job_name)

# キャンセルされたジョブをエクセルに書き込む
wb = functions.read_report()
sheet = wb['sheet1']
row = 26
for item in put_job:
    sheet.cell(row, column=3).value = item
    row += 1

functions.save_wb(wb)