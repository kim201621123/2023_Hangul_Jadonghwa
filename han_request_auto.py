#2023 10 10 김현성
#요청서 만들기
import win32com.client as win32  # 모듈 임포트
import pandas as pd
import datetime
import locale
import os
import time
import pyperclip as cb
import han_proceeding_auto



def make_request_list(read_file_route,save_file_route,meeting_times,Section_Chief):
    locale.setlocale(locale.LC_ALL, "ko_KR.utf-8")

    #여기서 쓸 파일 이름을 저장
    this_file_name = "제"+str(meeting_times)+"차 인증심의요청서"
    read_file_name = "제차 인증심의요청서.hwp"
    
    #participants.xlsx 파일 읽어오기
    df = pd.read_excel(read_file_route + r"\participants.xlsx")
    hwp_name = df["이름"]
    print(hwp_name)
    
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "AutomationModule")
    hwp.XHwpWindows.Item(0).Visible = True
    
    #한글 문서안에 텍스트 넣기 함수
    def insert_text(text):
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        #time.sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)

    # 날짜를 'yyyy년 mm월 dd일' 형식으로 변환
    today = datetime.datetime.today()
    # 일,월이 한 자리수일 경우 '0'을 제거합니다.
    formatted_date = today.strftime("%Y년   %#m월   %#d일")
    formatted_date_2 = today.strftime("%Y 년   %#m 월   %#d 일")
    
    #AGS번호 받아오기
    AGSnum = []
    AGSnum = han_proceeding_auto.find_AGSnum_from_txt()
    
    #hwp 파일 오픈
    #혹은 \를 쓰고싶으면 r을 붙이기
    hwp.Open(read_file_route+read_file_name)
    for i in range(len(AGSnum)):
        insert_text(i+1)
        hwp.HAction.Run("MoveRight")
        insert_text(AGSnum[i]) #번호
        hwp.HAction.Run("MoveRight")
        insert_text("AGS인증 1등급 심의")
        hwp.HAction.Run("MoveRight")
        #time.sleep(1.0)  # 1초 쉬어줌(꼭 필요)
        if i+1 != len(AGSnum):
            hwp.HAction.Run('TableAppendRow')
            hwp.HAction.Run("MoveLeft")
            hwp.HAction.Run("MoveLeft")
            hwp.HAction.Run("MoveLeft")
    hwp.HAction.Run("MoveDown")
    hwp.HAction.Run("MoveDown")
    
    
    
    #이제 누름틀로 이름넣고 날짜넣고 등등

    
    for i in range(len(hwp_name)):
        hwp.PutFieldText(f"차시",meeting_times)
        hwp.PutFieldText(f"AGS건수",len(AGSnum))
        hwp.PutFieldText(f"날짜",formatted_date)
        hwp.PutFieldText(f"날 짜",formatted_date_2)
    
    hwp.MovePos(3) #문서 끝으로 이동    
    hwp.HAction.Run("MoveUp")
    hwp.HAction.Run("MoveUp")
    
    #time.sleep(1.0)  # 1초 쉬어줌(꼭 필요)
    #AGS 숫자에 따라서 enter를 넣자 13-len(AGSnum)
    for i in range(13-len(AGSnum)):
        hwp.HAction.Run("BreakPara")

    
    #이름 넣어줘야지
    for i in range(len(hwp_name)):
        
        hwp.MovePos(0) #문서 처음으로 이동
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        
        name_i = df["동의인"][i]
        time.sleep(1.0)
        for i in range(15):
            hwp.HAction.Run("MoveLeft")
        
        insert_text(name_i)
        
        hwp.HAction.Run("CopyPage")
        hwp.HAction.Run("PastePage")
        
                
        time.sleep(0.6)
        hwp.MovePos(0) #문서 처음으로 이동
        time.sleep(0.6)

        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        hwp.HAction.Run("MoveDown")
        
        for i in range(11):
            hwp.HAction.Run("MoveLeft")
        for i in range(7):
            hwp.HAction.Run("DeleteBack")
        time.sleep(1.0)
    
    
    
    insert_text('  '.join(Section_Chief[:3]))    #엑셀에 단장님 이름 없음
    """
    hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = len(han_proceeding_auto.find_AGSnum_from_txt())
    hwp.HParameterSet.HTableCreation.Cols = 4
    hwp.HParameterSet.HTableCreation.WidthType = 2
    hwp.HParameterSet.HTableCreation.HeightType = 1
    hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(20.0)
    hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(10)
    hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 4)
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0,hwp.MiliToHwpUnit(33.4))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1,hwp.MiliToHwpUnit(33.4))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2,hwp.MiliToHwpUnit(33.4))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(3,hwp.MiliToHwpUnit(33.4))
    hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 3)
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0,hwp.MiliToHwpUnit(40.0))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1,hwp.MiliToHwpUnit(40.0))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(2,hwp.MiliToHwpUnit(20.0))
    hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1 #글자처럼 취급
    hwp.HParameterSet.HTableCreation.TableProperties.Width = hwp.MiliToHwpUnit(148)
    
    #바깥여백 0000으로 맞추기
    hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    hwp.HParameterSet.HShapeObject.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HShapeObject.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HShapeObject.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HShapeObject.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HShapeObject.HSet.SetItem("ShapeType", 6)
    hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
    
    hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    """
    
    #값 다 넣었고 페이지 저장
    hwp.SaveAs(os.path.join(save_file_route, this_file_name + ".hwp"))  # hwp 로 저장
    
    

