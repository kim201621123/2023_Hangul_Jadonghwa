import os
import datetime
import xlwinAGS as xw

def excel_copy(year_str,meeting_times,AGSnum,save_file_route):
    now_time = datetime.datetime.now()
    source_file_route = 'C:/Users/DDA/Desktop/부업무/제품요청/'+year_str+'/'+meeting_times[-2:]+'차 제품정보/'
    shared_path = "\\\\-공유저장소ip주소-\\DDA_심의\\"+str(now_time.year)+"년 위원회"
    
    
    files_of_savefolder = os.listdir(save_file_route)
    xlsm_route = []
    
    # 이번 차시 번호로 시작하는 xlsm파일 찾기
    for file in files_of_savefolder:
        if file.startswith("제"+meeting_times) and file.endswith(".xlsm"):
            xlsm_route = os.path.join(save_file_route, file)
    
    app = xw.App()
    
    if xlsm_route != [] :
        shared_path = xlsm_route
    else :
    # 저장할 .xlsx 파일 열기
        

        files = os.listdir(shared_path)
        file_list = os.listdir(shared_path)
        # 이번 차시로 시작하는 xlsm파일 찾기
        for file in file_list:
            if file.startswith(meeting_times[-2:]) and os.path.isdir(os.path.join(shared_path, file)):
                shared_path = os.path.join(shared_path, file)

        file_list = os.listdir(shared_path)

        # 이번 차시 번호로 시작하는 xlsm파일 찾기
        for file in file_list:
            if file.startswith("제"+meeting_times) and file.endswith(".xlsm"):
                shared_path = os.path.join(shared_path, file)  

    # b.xlsm 파일 열기
    wb_b = app.books.open(shared_path)

    # xlsx파일들 찾기
    files = os.listdir(source_file_route)
    xlsx_files = [f for f in files if f.endswith(".xlsx")]

    print(shared_path)
    for xlsx_file in xlsx_files:
        wb_a = app.books.open(source_file_route+xlsx_file)
        
        # a.xlsx 파일의 D5(AGS번호)셀 값 가져오기
        value_a = wb_a.sheets["제품 정보 요청"].range("D5").value 
        for i in range(len(AGSnum)):
            #AGS번호 순서대로, 회사명이 비어있으면 읽어 넣기
            if (value_a == AGSnum[i]) & (wb_b.sheets["위원회 목록"].range("C"+str(4+i)).value == None):
            #if value_a == AGSnum[i]:
                value_a = wb_a.sheets["제품 정보 요청"].range("B5").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("C"+str(4+i)).value = value_a
                
                value_a = wb_a.sheets["제품 정보 요청"].range("C5:L5").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("E"+str(4+i)+":"+"N"+str(4+i)).value = value_a           

                value_a = wb_a.sheets["제품 정보 요청"].range("B7:N7").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("O"+str(4+i)+":"+"AA"+str(4+i)).value = value_a

                value_a = wb_a.sheets["제품 정보 요청"].range("M5:N5").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("AB"+str(4+i)+":"+"AC"+str(4+i)).value = value_a
                
                value_a = wb_a.sheets["제품 정보 요청"].range("B9").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("AD"+str(4+i)).value = value_a

                value_a = wb_a.sheets["제품 정보 요청"].range("D9").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("AE"+str(4+i)).value = value_a

                value_a = wb_a.sheets["제품 정보 요청"].range("F9").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("AF"+str(4+i)).value = value_a
                
                value_a = wb_a.sheets["제품 정보 요청"].range("G9").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("AG"+str(4+i)).value = value_a

                value_a = wb_a.sheets["제품 정보 요청"].range("K9").value #원하는 범위값
                wb_b.sheets["위원회 목록"].range("AH"+str(4+i)).value = value_a

    # b.xlsm 파일 저장
    wb_b.save(save_file_route+"제"+meeting_times+"차.xlsm")

    # 엑셀 애플리케이션 종료
    app.quit()