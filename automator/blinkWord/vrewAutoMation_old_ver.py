import pyautogui

pag = pyautogui

print(pag.size())
#### main monitor width ####
# 만약 메인모니터가 있다면 메인모니터의 가로 값을 입력
# 메인모니터이거나 없으면 0
mmw = 3440
# 작업 속도 (0.3 이하로 내려가면 속도를 못따라감)
intv = 0.3
# 반복횟수
count = 2

# 엑셀시트 페이지
sheetPosW = mmw + 500
sheetPosH = 540

# vrew 페이지
vrewPosW = mmw + 1000
vrewPosH = 200

# 동영상 시트 영어단어 위치
engVideoTxtPosW = mmw + 1437
engVideoTxtPosH = 513

# 동영상 시트 한글단어 위치
korVideoTxtPosW = mmw + 1437
korVideoTxtPosH = 560

# ai 목소리 수정 위치
editVoiceW = mmw + 1550
editVoiceH = 990

# ai 목소리 수정 confirm 위치
editVoiceConfirmW = mmw + 1820
editVoiceConfirmH = 900

#### 함수 ####
# 엑셀 시트위치로 이동
def moveToSheet():
  pag.moveTo(sheetPosW, sheetPosH)
  pag.click(interval=intv)
  
# 영문복사
def engCopy():
  pag.press('down', interval=intv)
  pag.hotkey('command', 'c', interval=intv)

# 한글복사
def korCopy():
  pag.press('right',interval=intv)
  pag.hotkey('command', 'c', interval=intv)
  pag.press('left',interval=intv)
  
# vrew 창에서 가장 첫번째 컨텐츠로 이동
def moveToVrewFirstRun():
  pag.moveTo(vrewPosW, vrewPosH)
  pag.click(clicks=2, interval=intv)
  pag.hotkey('command', 'home', interval=intv)

# vrew 창으로 이동
def moveToVrew():
  pag.moveTo(vrewPosW, vrewPosH)
  pag.click(clicks=2, interval=intv)

# vrew 다음 컨텐츠 선택
def selectNextContent() :
  pag.press('esc',interval=intv)
  pag.press('esc',interval=intv)
  pag.hotkey('command', 'down', interval=intv)

# 동영상 영어 단어 수정
def engEdit():
  pag.moveTo(engVideoTxtPosW, engVideoTxtPosH)
  pag.click(interval=intv)
  pag.hotkey('command', 'a', interval=intv)
  pag.hotkey('command', 'v', interval=intv)
  
# 동영상 한글 단어 수정
def korEdit():
  pag.moveTo(korVideoTxtPosW, korVideoTxtPosH)
  pag.click(interval=intv)
  pag.hotkey('command', 'a', interval=intv)
  pag.hotkey('command', 'v', interval=intv)

# ai 목소리 수정
def voiceEdit(val): 
  pag.moveTo(editVoiceW, editVoiceH + val)
  pag.click(interval=intv)
  pag.hotkey('command', 'v', interval=intv)
  confirmVoiceEdit()

# ai 수정본 적용
def confirmVoiceEdit():
  pag.moveTo(editVoiceConfirmW, editVoiceConfirmH)
  pag.click(interval=intv)

# 마우스 포인터 위치값 받기
def getPointerPos():
  print('Press Ctrl-C to quit.')
  try:
      while True:
          x, y = pyautogui.position()
          positionStr = 'X: ' + str(x - mmw).rjust(4) + ' Y: ' + str(y).rjust(4)
          print(positionStr, end='')
          print('\b' * len(positionStr), end='', flush=True)
  except KeyboardInterrupt:
      print('\n')
  

#### Run From Here ####
moveToVrewFirstRun()

for i in range(21):
  #### 영문수정 ####
  moveToSheet()
  engCopy()
  
  moveToVrew()
  selectNextContent()
  
  engEdit()
  voiceEdit(0)
  
  #### 한글수정
  moveToSheet()
  korCopy()

  moveToVrew()
  selectNextContent()

  korEdit()
  voiceEdit(0)

print('TaskDone')

