# -*- coding:utf-8 -*-
import pyautogui

pag = pyautogui

print(pag.size())

#### main monitor width ####
# 듀얼 모니터 중 서브 모니터라면 메인모니터 가로 길이를 입력하면 되지만 될수 있으면 메인 모니터에서 하는게 좋음
mmw = 0

# 작업 속도 (ai 보이스 변경을 단축키로 대체 해서 작업속도 올림 0.5 -> 0.3)
intv = 0.3

# 엑셀시트 페이지 (mac 과 같은 값임 수정 불필요)
sheetPosW = mmw + 500
sheetPosH = 540

# vrew 페이지
vrewPosW = mmw + 1700
vrewPosH = 560

# 동영상 시트 영어단어 위치
engVideoTxtPosW = mmw + 1440
engVideoTxtPosH = 504

# 동영상 시트 한글단어 위치
korVideoTxtPosW = mmw + 1440
korVideoTxtPosH = 554

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

# count = pag.prompt(title="putCount", text="단어를 몇개 입력하실건가요?(숫자로만 입력)")
count = 18

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
