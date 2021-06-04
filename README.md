# 業務の自動化

### 準備
以下をインストール:

    pip install pywin32
    pip install pyautogui


### ~\Lib\site-packages\pyautogui\_pyautogui_win.py に以下を追記

追記箇所:

    def _keyDown(key):

        # 略
        needsShift = pyautogui.isShiftCharacter(key)

        #needsShiftの下の行に以下の3行を追加
        if key == '@': needsShift = False
        if key == '^': needsShift = False
        if key == ':': needsShift = False

ああ
