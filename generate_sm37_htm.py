import functions
import pyautogui
import time

def generate_job_list():
    # sm37を実行
    t_code = 'sm37'
    functions.get_window('SAP Easy Access')
    functions.execute_t_code(t_code)
    functions.get_window('Simple Job Selection')

    # Simple Job Selectionでキャンセルされたジョブのみ選択
    pyautogui.press('tab')
    pyautogui.typewrite(r'*')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('space')
    pyautogui.press('tab')
    pyautogui.press('space')
    pyautogui.press('tab')
    pyautogui.press('space')
    pyautogui.press('tab')
    pyautogui.press('space')
    time.sleep(2)
    pyautogui.press('f8')
    functions.get_window('Job Overview')


    # htmファイルを作成
    functions.move_to_localfile()
    functions.click_save_list_in_file()
    file_name = f'{t_code}_canceled_jobs'
    functions.click_to_genarate(file_name)

    # 画面を戻る
    functions.get_window('Job Overview')
    time.sleep(2)
    pyautogui.click(93, 53)
    pyautogui.press('f3')
