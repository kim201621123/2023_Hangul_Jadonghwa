#2023 10 25 kim
#결과서 만들기 - 결과서를 쪼갠다. 1 재발급 있는거/없는거, 표를 만들 때 총 수가 6개 이하/ 14 /초대 20까지 로 쪼갠다.
#총 양식은 5장정도???
#완성이긴 한데 돌발상황(수정발급, 공동제조자 등)에는 대응 할 수 없음 - 수기로 추가하시오

import win32com.client as win32  # 모듈 임포트
import pandas as pd
import datetime
import locale
import os
import re
import time
import pyperclip as cb
import han_proceeding_auto
import openpyxl
from win32 import win32api
import win32security

def make_result_paper(read_file_route,save_file_route,meeting_times,indivisual_participants,AGSnum,ReAGSnum,GaNaDaAGSnum):
    locale.setlocale(locale.LC_ALL, "ko_KR.utf-8")
    
    this_file_name = "제"+str(meeting_times)+"차 결과서"
    
    # 인수대로 서식 파일을 그냥 만들어둬야겠다,.....
    # 만약 ReAGS번호 없으면  NOGeAGS.hwp 열기
    if len(ReAGSnum) > 0:
        read_file_name = "제차 결과서_"+"ReAGS"+".hwp"
    else :
        read_file_name = "제차 결과서_"+"NoReAGS"+".hwp"
    
    # 날짜를 'yyyy년 mm월 dd일' 형식으로 변환
    today = datetime.datetime.today()
    # 일,월이 한 자리수일 경우 '0'을 제거합니다.
    formatted_date = today.strftime("%Y년   %#m월   %#d일")
    formatted_date_2 = today.strftime("%Y 년   %#m 월   %#d 일")
    
    #한글파일 열기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
    hwp.XHwpWindows.Item(0).Visible = True
    
    #한글 문서안에 텍스트 넣기 함수
    def insert_text(text,hwp):
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        #time.sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)

    #hwp 파일 오픈
    #혹은 \를 쓰고싶으면 r을 붙이기
    hwp.Open(read_file_route+read_file_name)
   
   
    #참석위원을 한줄로 만들기
    real_participants = ""
    for i in range(len(indivisual_participants)):
        if i != 0:
            real_participants +=", "
        real_participants += indivisual_participants[i]
    
    hwp.PutFieldText(f"차시",meeting_times)
    hwp.PutFieldText(f"날짜",formatted_date)
    hwp.PutFieldText(f"참석위원",real_participants)
    hwp.PutFieldText(f"건수",len(AGSnum))
    hwp.PutFieldText(f"재발급건수",len(ReAGSnum))
    
    #가.AGS번호 넣기
    for i in range(len(GaNaDaAGSnum)):
        insert_text(GaNaDaAGSnum[i],hwp)
        hwp.HAction.Run("BreakPara")
    
    
    #제품정보 엑셀 읽어오기
    #여기서부터는 재발급이 있을 때만 해야 함
    if len(ReAGSnum) != 0 :
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        
        sheet = han_proceeding_auto.access_to_shared_excel(meeting_times)
        
        for i in range(len(ReAGSnum)):
            for j in range(4):
                sheet_value = sheet.cell(i+4,10+j).value
                if sheet_value:
                    lines = sheet_value.split('\n')
                    for k, line in enumerate(lines):
                        insert_text(line,hwp)
                        if k != (len(lines) - 1):
                            hwp.HAction.Run("BreakPara")
                hwp.HAction.Run("MoveRight")
            if i != (len(ReAGSnum)-1):
                hwp.HAction.Run('TableAppendRow')
                hwp.HAction.Run("MoveLeft")
                hwp.HAction.Run("MoveLeft")
                hwp.HAction.Run("MoveLeft")
            
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveLeft")
        hwp.HAction.Run("MoveLeft")
    else: #재발급 없으면
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")


    all_AGSNum = AGSnum + ReAGSnum
    
    for i in range(0, len(all_AGSNum)):
        #재발급 번호면
        if i >= len(AGSnum):
            insert_text(all_AGSNum[i][0:7],hwp)
            hwp.HAction.Run("BreakPara")
            insert_text(all_AGSNum[i][7:],hwp)
        else:
            insert_text(all_AGSNum[i],hwp)
        time.sleep(0.4)
        hwp.HAction.Run("MoveRight")
        insert_text(" 인증 심의 : 가결(可決)  부결(否決)  보류(保留)",hwp)
        hwp.HAction.Run("BreakPara")
        insert_text(" 의견 : ",hwp)
        
        time.sleep(0.4)
        
        if i != (len(all_AGSNum)-1):
            hwp.HAction.Run('TableAppendRow')
            hwp.HAction.Run("MoveLeft")
    
    
    hwp.PutFieldText(f"날 짜",formatted_date_2)
    
    #hwp_2.XHwpDocuments.Item(0).Close(isDirty=False)
    #hwp_2.Quit()

        
    hwp.SaveAs(os.path.join(save_file_route, this_file_name + ".hwp"))  # hwp 로 저장
    #hwp_2.HAction.Run("CopyPage")
    #hwp.HAction.Run("PastePage")
    