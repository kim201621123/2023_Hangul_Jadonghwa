AAGS#2023 11 08 kim
#의사록 만들기


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

def make_minutes_paper(read_file_route,save_file_route,meeting_times,indivisual_participants,AGSnum,ReAGSnum,GaNaDaAGSnum,Section_Chief):
    locale.setlocale(locale.LC_ALL, "ko_KR.utf-8")


    this_file_name = "제"+str(meeting_times)+"차 의사록"

    read_file_name = "제차 의사록"+".hwp"

    # 인수대로 서식 파일을 그냥 만들어둬야겠다,.....
    # 만약 ReAGS번호 없으면  NOGeAGS.hwp 열기
    if len(ReAGSnum) > 0:
        read_file_name = "제차 의사록_"+"ReAGS"+".hwp"
    else :
        read_file_name = "제차 의사록_"+"NoReAGS"+".hwp"


    
    # 날짜를 'yyyy년 mm월 dd일' 형식으로 변환
    today = datetime.datetime.today()
    # 일,월이 한 자리수일 경우 '0'을 제거합니다.
    formatted_date = today.strftime("%Y년   %#m월   %#d일")
    formatted_date_2 = today.strftime("%Y 년   %#m 월   %#d 일")
    
    #한글파일 열기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "AutomationModule") #보안모듈 설치 후 open확인 메시지 삭제함 굿!
    hwp.XHwpWindows.Item(0).Visible = True
    
    #한글 문서안에 텍스트 넣기 함수
    def insert_text(text,hwp):
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        #time.sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)

    #hwp 파일 오픈
    #혹은 \를 쓰고싶으면 r을 붙이기
    hwp.Open(read_file_route+read_file_name,"HWP","forceopen:true")
   
    #참석위원을 한줄로 만들기
    #여기서 상수로 사용한 '6'은 사용했던 서식에서 참석자가 6명이 넘으면 한 줄을 내려야 해서 저렇게 적은것으로 보입니다.
    real_participants = ""
    for i in range(len(indivisual_participants)):
        if i != 0:
            real_participants +=","
            if i == 6:
                real_participants +="ㅤ"
            real_participants +=" "
        real_participants += indivisual_participants[i]
    
    hwp.PutFieldText(f"차시",meeting_times)
    hwp.PutFieldText(f"날짜",formatted_date)
    hwp.PutFieldText(f"참석위원",real_participants)
    hwp.PutFieldText(f"건수",len(AGSnum))
    hwp.PutFieldText(f"재발급 건수",len(ReAGSnum))
    participants_people = han_proceeding_auto.find_participants_people()
    hwp.PutFieldText(f"참관인",participants_people)
    
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
        
        sheet = han_proceeding_auto.excel_file_root(meeting_times)
        
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
        hwp.HAction.Run("MoveDown")
        
    else:
        hwp.HAction.Run("MoveDown") #몇번?
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
    
    match_plus_AGS = ""    
    match_minus_AGS = ""
    match_plus_AGS, match_minus_AGS = han_proceeding_auto.find_PlusMinusAGS()
    
    if (match_plus_AGS.count("AGS") != 0) or (match_minus_AGS.count("AGS") != 0):
        insert_text("안건 "+str(len(AGSnum)+match_minus_AGS.count("AGS")-match_plus_AGS.count("AGS"))+"건에서 ",hwp)
    if (match_minus_AGS.count("AGS") != 0):
        insert_text(str(match_minus_AGS.count("AGS"))+"건이 취소("+match_minus_AGS+")되",hwp)
    if (match_plus_AGS.count("AGS") != 0) and (match_minus_AGS.count("AGS") != 0):
        insert_text("고 ",hwp)
    elif (match_plus_AGS.count("AGS") == 0) and (match_minus_AGS.count("AGS") != 0):
        insert_text("어 ",hwp)
    if (match_plus_AGS.count("AGS") != 0):
        insert_text(str(match_plus_AGS.count("AGS"))+"건이 추가("+match_plus_AGS+")되어 ",hwp)
    insert_text("제품 "+str(len(AGSnum))+"건 심의하고 승인함",hwp)
    
    
    hwp.HAction.Run("MoveDown")
    if len(ReAGSnum) > 0:
        insert_text("2) AGS인증서(1등급) 재발급 "+str(len(ReAGSnum))+"건 심의하고 승인함",hwp)
    
    hwp.PutFieldText(f"날 짜",formatted_date_2)

    hwp.SaveAs(os.path.join(save_file_route, this_file_name + ".hwp"))  # hwp 로 저장
    #hwp_2.HAction.Run("CopyPage")
    #hwp.HAction.Run("PastePage")
    