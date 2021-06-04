import os
import pyautogui
import time
import win32gui
import openpyxl


current_directory = os.getcwd()
htm_directory = (f'{current_directory}\\tmp\\')

# htm ファイルを取得
def open_htm(file_name):
    file = open(f'{htm_directory}{file_name}.htm', 'r', encoding='UTF-8')
    data = file.read()
    file.close()
    return data

# ウィンドウがアクティブになったら前面に表示
def get_window(window_name):
    win = win32gui.FindWindow(None, window_name)
    while True:
        if not win:
            time.sleep(1)
            win = win32gui.FindWindow(None, window_name)
        else:
            break
    win32gui.SetForegroundWindow(win)
    del win

# トランザクションコードを実行
def execute_t_code(t_code):
    time.sleep(1)
    pyautogui.hotkey('ctrl', '/')
    pyautogui.typewrite(f'/n{t_code}')
    time.sleep(1)
    pyautogui.press('enter')

# ファイルを取得するまでのマウスの動き
def move_to_localfile():
    time.sleep(3)
    pyautogui.click(325, 23)
    time.sleep(2)
    pyautogui.moveTo(325, 155, 0.5)
    pyautogui.moveTo(545, 155, 0.5)
    pyautogui.moveTo(545, 200, 0.5)
    pyautogui.moveTo(680, 200, 0.5)
    pyautogui.moveTo(680, 250, 0.5)
    pyautogui.click(680, 250)

# save list in file 画面でのキー操作
def click_save_list_in_file():
    time.sleep(2)
    pyautogui.press(['down', 'down', 'down', 'tab'])
    time.sleep(3)
    pyautogui.press(['enter'])

# ファイルを作成するまでのキー操作
def click_to_genarate(file_name):
    time.sleep(2)
    pyautogui.typewrite(f'{file_name}.htm')
    time.sleep(3)
    pyautogui.hotkey('shift', 'tab')
    time.sleep(1)
    pyautogui.typewrite(htm_directory)
    time.sleep(1)
    pyautogui.press(['tab', 'tab'])
    time.sleep(3)
    pyautogui.press(['enter'])

# エクセル を取得
def read_report():
    wb = openpyxl.load_workbook(f'{htm_directory}DD-MM-YYYY-R11_monitoring_report.xlsm', keep_vba=True)
    return wb

# エクセル を保存
def save_wb(wb):
    wb.save(f'{htm_directory}DD-MM-YYYY-R11_monitoring_report.xlsm')