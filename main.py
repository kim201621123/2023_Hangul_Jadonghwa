import datetime
import os
import han_proceeding_auto
import han_participants_auto
import han_request_auto
import han_result_table_1
import han_result_table
import han_result_paper
import han_minutes_paper
import excel_to_excel_copy
import xlwinAGS as xw


participants = []

Section_Chief = "AAA단장"
Team_1_Manager = "BBB팀장"
Team_2_Manager = "CCC팀장"


# 소스파일이 있는 디렉토리에서 .txt 확장자를 가진 파일만 선택합니다.
files = os.listdir()
txt_files = [f for f in files if f.endswith(".txt")]

#연도 끝 2자리 자르고 파일 차시랑 더해서 '연도 끝 두자리-**'으로 만들기
now = datetime.datetime.now()
year = now.year
year_str = str(year)
year_end_two_digits = year_str[-2:]
meeting_times = txt_files[0][0:2]
meeting_times = year_end_two_digits+"-"+meeting_times

list_participants=han_proceeding_auto.find_participants_from_txt(Section_Chief) #메모장의 참석자 명단 받아오기
AGSnum = han_proceeding_auto.find_AGSnum_from_txt()
ReAGSnum = han_proceeding_auto.find_ReAGSnum_from_txt()
GaNaDaAGSnum = han_proceeding_auto.find_GaNaDaAGSnum_from_txt()



read_file_route = 'C:/Users/DDA/Desktop/hangule_Automated/forms/'
save_file_route = 'C:/Users/DDA/Desktop/hangule_Automated/saves/' + meeting_times + '/'
if not os.path.exists(save_file_route):
    os.makedirs(save_file_route)
else:
    pass

#-----------------------------------------------------------------------------------------


#요청서는 participants.xlsx 기반으로 만든다
han_request_auto.make_request_list(read_file_route,save_file_route,meeting_times,Section_Chief) #요청서


#참석명단 파일은 participants.xlsx과 현재 디렉토리의 txt 파일로 생성한다
#han_participants_auto.make_participants_list(read_file_route,save_file_route,meeting_times,list_participants) #참석명단 개별 싸인받는거 만들기


#참석명단 아래에 있어야 함
# 이전까지는 단장님이 안들어가있어서 따로 넣어줘야 함 / 항상 활성화 추천
if input("혹시 단장님이 부재신가요? y/n : ") == "y":
    if input("혹시 2팀 팀장님이 부재신가요? y/n : ") == "y":
        list_participants.append(Team_1_Manager)
    else:
        list_participants.append(Team_2_Manager)
else:
    list_participants.append(Section_Chief)
print(list_participants)


#결과표는 participants.xlsxrhk txt의 참석명단으로 만든다
# 참석인원 수에 맞는 한글파일을 연다
#han_result_table.make_result_list(read_file_route,save_file_route,meeting_times,list_participants)


#의사록 만들기
#han_minutes_paper.make_minutes_paper(read_file_route,save_file_route,meeting_times,list_participants,AGSnum,ReAGSnum,GaNaDaAGSnum,Section_Chief)


#결과서 만들기
#han_result_paper.make_result_paper(read_file_route,save_file_route,meeting_times,list_participants,AGSnum,ReAGSnum,GaNaDaAGSnum)



#--------------------------문서와는 관련 없음------------------------------


#제품정보 저장 페이지에서 정보 받아와서 65-ip 폴더-파일에 값 넣고 save폴더에 저장
#실행 전에 파일 정리하고 실행하기
#excel_to_excel_copy.excel_copy(year_str,meeting_times,AGSnum,save_file_route)

#가나다 AGS번호 대로 폴더 만드는 거
#han_proceeding_auto.make_AGS_folder(GaNaDaAGSnum,meeting_times,han_proceeding_auto.The_Day_from_txt(meeting_times))


