# -*- coding:utf-8 -*-
import pyautogui

pag = pyautogui

print(pag.size())

#### main monitor width ####
# 맥북 모니터가 메인 모니터이기때문에 값 안줌
mmw = 0

# 작업 속도 (0.5 이하로 내려가면 로딩속도로 인해 씹힘)
intv = 0.5
# 반복횟수
# count = 2

# 엑셀시트 페이지
sheetPosW = mmw + 500
sheetPosH = 540

# vrew 페이지
vrewPosW = mmw + 1620
vrewPosH = 560

# 동영상 시트 영어단어 위치
engVideoTxtPosW = mmw + 1150
engVideoTxtPosH = 580

# 동영상 시트 한글단어 위치
korVideoTxtPosW = mmw + 1150
korVideoTxtPosH = 633

# ai 목소리 수정 위치
editVoiceW = mmw + 1400
editVoiceH = 1100

# ai 목소리 수정 confirm 위치
editVoiceConfirmW = mmw + 1490
editVoiceConfirmH = 920


#### 함수 ####
# 엑셀 시트위치로 이동
def moveToSheet():
    pag.moveTo(sheetPosW, sheetPosH)
    pag.click(interval=intv)


# vrew 창으로 이동
def moveToVrew():
    pag.moveTo(vrewPosW, vrewPosH)
    pag.click(clicks=2, interval=intv)


# 영문복사
def engCopy():
    pag.press("down", interval=intv)
    pag.hotkey("command", "c", interval=intv)


# 한글복사
def korCopy():
    pag.press("right", interval=intv)
    pag.hotkey("command", "c", interval=intv)
    pag.press("left", interval=intv)


# vrew 다음 컨텐츠 선택
def selectNextContent():
    pag.click(clicks=1, interval=intv)
    pag.hotkey("command", "down", interval=intv)


# 동영상 영어 단어 수정
def engEdit():
    pag.moveTo(engVideoTxtPosW, engVideoTxtPosH)
    pag.click(clicks=1, interval=intv)
    pag.hotkey("command", "a", interval=intv)
    pag.hotkey("command", "v", interval=intv)
    moveToVrew()


# 동영상 한글 단어 수정
def korEdit():
    pag.moveTo(korVideoTxtPosW, korVideoTxtPosH)
    pag.click(clicks=1, interval=intv)
    pag.hotkey("command", "a", interval=intv)
    pag.hotkey("command", "v", interval=intv)
    moveToVrew()


# ai 목소리 수정
def voiceEdit():
    # pag.moveTo(editVoiceW, editVoiceH)
    # pag.click(interval=intv)
    pag.hotkey("F8", interval=intv)
    pag.hotkey("command", "v", interval=intv)
    pag.hotkey("enter", interval=intv)
    # confirmVoiceEdit()


# ai 수정본 적용
def confirmVoiceEdit():
    pag.moveTo(editVoiceConfirmW, editVoiceConfirmH)
    pag.click(interval=intv)


# 마우스 포인터 위치값 받기
def getPointerPos():
    print("Press Ctrl-C to quit.")
    try:
        while True:
            x, y = pyautogui.position()
            positionStr = "X: " + str(x - mmw).rjust(4) + " Y: " + str(y).rjust(4)
            print(positionStr, end="")
            print("\b" * len(positionStr), end="", flush=True)
    except KeyboardInterrupt:
        print("\n")


# getPointerPos()

#### Run From Here ####

count = pag.prompt(title="putCount", text="단어를 몇개 입력하실건가요?(숫자로만 입력)")

for i in range(int(count)):
    #### 영문수정 ####

    moveToSheet()
    engCopy()

    moveToVrew()
    selectNextContent()

    engEdit()
    voiceEdit()

    # #### 한글수정

    moveToSheet()
    korCopy()

    moveToVrew()
    selectNextContent()

    korEdit()
    voiceEdit()

print("TaskDone")
print("\a")
print("\a")
print("\a")
