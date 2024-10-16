#2023 10 06 kim
#참석자 list를 전달받아 만드는 코드
import win32com.client as win32  # 모듈 임포트
import pandas as pd
import datetime
import locale
import os
import time

def make_participants_list(read_file_route,save_file_route,meeting_times,indivisual_participants):
    locale.setlocale(locale.LC_ALL, "ko_KR.utf-8")

    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.RegisterModule("FilePathCheckDLL", "AutomationModule") #보안모듈 설치 후 open확인 메시지 삭제함 굿!
    
    # 날짜를 'yyyy년 mm월 dd일' 형식으로 변환
    today = datetime.datetime.today()
    # 일,월이 한 자리수일 경우 '0'을 제거합니다.
    formatted_date = today.strftime("%Y년   %#m월   %#d일")

    #participants.xlsx 파일 읽어오기
    df = pd.read_excel(read_file_route + r"\participants.xlsx")


    #이름 따로 분리하고
    hwp_name = df["이름"]

    #hwp 파일 오픈
    #혹은 \를 쓰고싶으면 r을 붙이기
    
    hwp.Open(read_file_route+"제차_위원회-참석명단 -  위원님.hwp")

    print(indivisual_participants)
    print("-------개별싸인------")
    #페이지 만들기 
    for i in range(len(hwp_name)):
        for j in range(len(indivisual_participants)):
            if indivisual_participants[j][:3] in df["이름"][i][:3]:
                #값 넣기
                hwp.PutFieldText(f"차시", meeting_times)
                hwp.PutFieldText(f"소속/부서명",df["소속/부서명"][i])
                hwp.PutFieldText(f"이름",df["이름"][i])
                hwp.PutFieldText(f"은행명",df["은행명"][i])
                hwp.PutFieldText(f"계좌번호",df["계좌번호"][i])
                hwp.PutFieldText(f"주소",df["주소"][i])
                hwp.PutFieldText(f"날짜",formatted_date)
                hwp.PutFieldText(f"동의인",df["동의인"][i])

                #값 다 넣었고 페이지 저장
                hwp_i= "제"+ str(meeting_times) + "차_위원회-참석명단 - "+ hwp_name[i]+ "위원님"
                hwp.SaveAs(os.path.join(save_file_route, hwp_i + ".hwp"))  # hwp 로 저장
                time.sleep(0.2)  # 0.1초 쉬어줌(꼭 필요)
    