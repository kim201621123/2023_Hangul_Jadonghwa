#2023 10 11 kim
#결과표 만들기
import win32com.client as win32  # 모듈 임포트
import pandas as pd
import datetime
import locale
import os
import re
import time
import pyperclip as cb
import han_proceeding_auto



def make_result_list(read_file_route,save_file_route,meeting_times,indivisual_participants):
    locale.setlocale(locale.LC_ALL, "ko_KR.utf-8")
    
    #여기서 쓸 파일 이름을 저장
    this_file_name = "제"+str(meeting_times)+"차 결과표"
    
    if (len(indivisual_participants)) <= 4:
        print("참석자 수가 5인을 넘지 못해 서식이 없습니다. 손으로 만드세용")
        return 0
    
    # 인수대로 서식 파일을 그냥 만들어둬야겠다,.....
    read_file_name = "제차 결과표_"+str(len(indivisual_participants))+".hwp"
    
    #participants.xlsx 파일 읽어오기
    df = pd.read_excel(read_file_route + r"\participants.xlsx")
    hwp_name = df["이름"]
    
    
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
    hwp.XHwpWindows.Item(0).Visible = True
    
    #한글 문서안에 텍스트 넣기 함수
    def insert_text(text):
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        time.sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)

    # 날짜를 'yyyy년 mm월 dd일' 형식으로 변환
    today = datetime.datetime.today()
    # 일,월이 한 자리수일 경우 '0'을 제거합니다.
    formatted_date = today.strftime("%Y 년   %#m 월   %#d 일")
    
    #AGS번호 받아오기
    AGSnum = []
    AGSnum = han_proceeding_auto.find_AGSnum_from_txt()
    #재발급 AGS번호 받아오기
    ReAGSnum = []
    ReAGSnum = han_proceeding_auto.find_ReAGSnum_from_txt()
    
    
    #hwp 파일 오픈
    #혹은 \를 쓰고싶으면 r을 붙이기
    hwp.Open(read_file_route+read_file_name)
    
    
    #신규 AGS 시험번호 넣기
    for i in range(len(AGSnum)):
        insert_text(AGSnum[i])
        hwp.HAction.Run("MoveRight")
        for j in range(len(indivisual_participants)+1):
            insert_text("가, 부, 보류\n")
            hwp.HAction.Run("MoveRight")
        
        insert_text("1등급")
        if i != (len(AGSnum)-1):
            hwp.HAction.Run('TableAppendRow')
        for k in range(len(indivisual_participants)+2):
            hwp.HAction.Run("MoveLeft")
    
    #재발급이 있으면 줄 추가
    if len(ReAGSnum) != 0:
        hwp.HAction.Run('TableAppendRow')
        for k in range(len(indivisual_participants)+1):
            hwp.HAction.Run("MoveLeft")
    
    
    
    # 정규 표현식을 생성합니다.
    pattern = re.compile(r"^(.{7})(.*)")    #앞 7자리
    # 텍스트에서 정규 표현식에 일치하는 텍스트를 찾습니다.
    results = []    #앞 인증번호
    rests = []       #뒤 인증날짜
    for text_item in ReAGSnum:
        matches = pattern.finditer(text_item)
        for match in matches:
            results.append(match.group(1))
            rests.append(match.group(2))
    
        
    #재발급 AGS 시험번호 넣기
    for i in range(len(ReAGSnum)):
        
        insert_text(results[i])
        hwp.HAction.Run("BreakPara")
        insert_text(rests[i])
        hwp.HAction.Run("MoveRight")
        
        for j in range(len(indivisual_participants)+1):
            insert_text("가, 부, 보류")
            hwp.HAction.Run("MoveRight")
        
        insert_text("1등급")
        hwp.HAction.Run("BreakPara")
        insert_text("(재발급)")
        
        if i != (len(ReAGSnum)-1):
            hwp.HAction.Run('TableAppendRow')
            
        for k in range(len(indivisual_participants)+2):
            hwp.HAction.Run("MoveLeft")
    
    #차시, 날짜, 이름 등 넣기
    hwp.PutFieldText(f"차시",meeting_times)

    for i in range(len(indivisual_participants)):
        hwp.PutFieldText(f"위원"+str(i+1),indivisual_participants[i])
        
    hwp.PutFieldText(f"날짜",formatted_date)
    
    #문서 끝으로 이동
    hwp.MovePos(3)
    for i in range(10):
        hwp.HAction.Run("MoveUp")
        
    hwp.HAction.Run("BreakPara")
    
    #심사한 제품+재발급 건수가 5이상이면 줄 한칸 줄이기
    if (len(AGSnum) + len(ReAGSnum)) > 12:
        hwp.HAction.Run("DeleteBack")
    
    #값 다 넣었고 페이지 저장
    hwp.SaveAs(os.path.join(save_file_route, this_file_name + ".hwp"))  # hwp 로 저장
    #hwp.Quit()
    
    

