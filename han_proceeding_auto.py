#2023 10 10 kim with Bard, chatGPT

import os
import re
import datetime
import sys
import openpyxl
#with open(txt_file, "r") as f:
#이상하게도 한번 파일 열고 찾으면 커서가 맨 밑으로 가버려서 그 안에서 뭔가 하기 힘들 것 같다. 그래서 뭔가 할 때마다 다시 여는중...

#매개변수로 현재 디렉토리랑 차시를 입력받으면 좋지 않을까?
#결론 : 메모장의 참석자 명단만 list로 전달한다.  
def find_participants_from_txt(Section_Chief):
    # 현재 디렉토리의 파일 목록을 가져옵니다.
    files = os.listdir()

    participants = []
    
    # .txt 확장자를 가진 파일만 선택합니다.
    txt_files = [f for f in files if f.endswith(".txt")]

    # 찾고자 하는 텍스트를 찾습니다.
    for txt_file in txt_files:

        # 메모장 파일을 열고, 읽기 모드로 엽니다.
        with open(txt_file, "r") as f:

            # match_start가 None이면, 에러 메시지를 출력하고 코드를 계속 실행합니다.
            match_start = re.search("<참석가능>", f.read())
            if match_start is None:
                print(f"'{txt_file}' 파일에서 '<참석가능>' 텍스트를 찾을 수 없습니다.")
                continue
        
        with open(txt_file, "r") as f:
            # 찾은 텍스트의 끝 위치를 가져옵니다.
            match_end = re.search("<참석불가능>", f.read())
            if match_end is None:
                print(f"'{txt_file}' 파일에서 '<참석불가능>' 텍스트를 찾을 수 없습니다.")
                continue
        
        with open(txt_file, "r") as f:    
            # 찾은 텍스트 사이의 문자열을 한 줄씩 읽어옵니다.
            for line in f.read()[match_start.end():match_end.start()].splitlines():

                # "(" 뒤 3자리 텍스트를 가져옵니다.
                match = re.search("\((.*?)\)", line)
                # 찾은 경우, "(" 뒤 3자리 텍스트를 출력합니다.
                if match is not None:
                    participants.append(match.group(1)[0:3])
    
    # 각 문자열의 앞글자끼리 비교합니다.
    for i in range(len(participants) - 1):
        for j in range(i + 1, len(participants)):
            if participants[i][0] < participants[j][0]:
                participants[i], participants[j] = participants[j], participants[i]
    #한글 오름차순으로 정렬
    participants = participants[::-1]
    if "자차카" in participants:
        participants = [x for x in participants if x not in ["자차카"]]
        participants.insert(0, "자차카")
    participants = [x + "위원" for x in participants]
    #단장님은 넣지 않는다
    participants = [x for x in participants if x not in [Section_Chief[:3]+"위원"]]
    print(participants)
    print("------메모장에서 단장님 이름 뻄-----")
    return participants

#신규 AGS 번호 찾아서 list를 return
def find_AGSnum_from_txt():
    # 현재 디렉토리의 파일 목록을 가져옵니다.
    files = os.listdir()
    AGSnum = []
    # .txt 확장자를 가진 파일만 선택합니다.
    txt_files = [f for f in files if f.endswith(".txt")]
    
    # 파일 열기 및 처리
    try:        
        with open(txt_files[0], "r", encoding="cp949") as file:
            lines = file.readlines()
            for line in lines:
                # "AGS-"로 시작하는 텍스트인 경우에만 리스트에 추가
                if line.startswith("AGS-"):
                    AGSnum.append(line.strip())  # 줄 바꿈 문자 제거 후 추가
        #print(AGSnum)
    #파일이 없어요?
    except FileNotFoundError:
        print(f"'{txt_files[0]}' 파일을 찾을 수 없습니다. '**차.txt' 파일을 현 위치에 위치시켜주세요")
    except Exception as e:
        print(f"오류 발생: {e}")
        
    return AGSnum
 
 
 # 가나다 AGS 번호 찾아서 list를 return
def find_GaNaDaAGSnum_from_txt():
    # 현재 디렉토리의 파일 목록을 가져옵니다.
    files = os.listdir()
    AGSnum = []
    # .txt 확장자를 가진 파일만 선택합니다.
    txt_files = [f for f in files if f.endswith(".txt")]
    
    # 파일 열기 및 처리
    try:
        with open(txt_files[0], "r", encoding="cp949") as file:
            lines = file.readlines()
            # 정규 표현식 사용
            pattern = re.compile("^(가|나|다|라|마|바|사|아|자|차|카|타|파). AGS-")
            # 리스트 생성
            result = []
            # 텍스트 순회
            for line in lines:
                # 정규 표현식에 일치하는지 확인
                if pattern.match(line):
                    # 리스트 추가
                    result.append(line.rstrip("\n"))

            # 출력
            #print(result)
    #파일이 없어요?
    except FileNotFoundError:
        print(f"'{txt_files[0]}' 파일을 찾을 수 없습니다. '**차.txt' 파일을 현 위치에 위치시켜주세요")
    except Exception as e:
        print(f"오류 발생: {e}")
        
    return result
     
       
#재발급 AGS번호를 가져오기       
def find_ReAGSnum_from_txt():
    # 현재 디렉토리의 파일 목록을 가져옵니다.
    files = os.listdir()
    # .txt 확장자를 가진 파일만 선택합니다.
    txt_files = [f for f in files if f.endswith(".txt")]
    # 메모장 파일을 읽습니다.
    with open(txt_files[0], "r") as f:
        text = f.read()
    # 정규 표현식을 생성합니다.
    pattern = re.compile(r"(\d{2})-(\d{4})\((\d{4}.\d{2}.\d{2})\)")
    # 텍스트에서 정규 표현식에 일치하는 텍스트를 찾습니다.
    matches = pattern.finditer(text)
    # 찾은 텍스트를 리스트에 저장합니다.
    ReAGSnum = []
    for match in matches:
        ReAGSnum.append(match.group(0))
    
    return ReAGSnum

#공유폴더에서 엑셀 찾기
def access_to_shared_excel(meeting_times):
    now_time = datetime.datetime.now()
    
    shared_path = "\\\\공유폴더 ip주소\\DDA_인증심의\\"+str(now_time.year)+"년 위원회"
    file_list = os.listdir(shared_path)
    # 이번 인증위 차시로 시작하는 xlsm파일 찾기
    for file in file_list:
        if file.startswith(meeting_times[-2:]) and os.path.isdir(os.path.join(shared_path, file)):
            shared_path = os.path.join(shared_path, file)
    print(shared_path)
    file_list = os.listdir(shared_path)
    print(file_list)
    # '58'로 시작하는 xlsx파일 찾기
    for file in file_list:
        if file.startswith("제"+meeting_times) and file.endswith(".xlsm"):
            shared_path = os.path.join(shared_path, file)  
    print(shared_path)
    # 파일을 읽어 들입니다.
    workbook = openpyxl.load_workbook(shared_path, data_only=True)
    # 시트 이름을 입력합니다.
    sheet_name = "제품설명(보고용)"

    # 시트를 가져옵니다.
    sheet = workbook[sheet_name]
            
    return sheet


def excel_file_root(meeting_times):
    # xlsm 파일을 찾습니다.
    files = os.listdir()
    for file in files:
        if file.startswith("제"+meeting_times) and file.endswith(".xlsm"):
            # 파일 이름을 입력합니다.
            file_name = file

    # 파일을 읽어 들입니다.
    workbook = openpyxl.load_workbook(file_name, data_only=True)
    
    # 시트 이름을 입력합니다.
    sheet_name = "제품설명(보고용)"

    # 시트를 가져옵니다.
    sheet = workbook[sheet_name]
    
    return sheet

#의사록에 들어가는 ~건 취소, ~건 추가 문구 만드려고
def find_PlusMinusAGS():
    # 현재 디렉토리의 파일 목록을 가져옵니다.
    files = os.listdir()
    # .txt 확장자를 가진 파일만 선택합니다.
    txt_files = [f for f in files if f.endswith(".txt")]
    # 메모장 파일을 읽습니다.
    with open(txt_files[0], "r") as f:
        text = f.read()
    # '추가'라는 텍스트 뒤에 '('와 ')' 사이의 텍스트를 찾음
    a=""
    b=""
    match_plus=""
    match_minus=""
    if re.search(r"추가\s*\((.*?)\)", text) :
        a = re.search(r"추가\s*\((.*?)\)", text)
        match_plus = a.group(1)
    if re.search(r"취소\s*\((.*?)\)", text):
        b = re.search(r"취소\s*\((.*?)\)", text)
        match_minus = b.group(1)
    # 텍스트가 존재하는 경우
    return str(match_plus), str(match_minus)

#참관인
def find_participants_people():
    # 현재 디렉토리의 파일 목록을 가져옵니다.
    files = os.listdir()
    # .txt 확장자를 가진 파일만 선택합니다.
    txt_files = [f for f in files if f.endswith(".txt")]
    # 파일을 열고 내용을 읽습니다.
    with open(txt_files[0], "r") as file:
        lines = file.readlines()        
    # "참관인" 텍스트를 찾습니다.
    for i, line in enumerate(lines):
        if "<참관인>" in line:
            # 해당 텍스트를 찾았으므로 다음 4줄을 가져옵니다.
            result_text_1 = "".join(lines[i+1:i+2])
            result_text_1 = result_text_1[0:3]
            result_text_2 = "".join(lines[i+4:i+5])
            result_text_1 = result_text_1+", "+ result_text_2
            return result_text_1
        
def make_AGS_folder(GaNaDaAGSnum,meeting_times,numbers_only_date):
    now_time = datetime.datetime.now()
    shared_path = "\\\\공유폴더 ip\\DDA_인증심의\\"+str(now_time.year)+"년 위원회"
    
    file_list = os.listdir(shared_path)
    # 이번 인증위 차시로 시작하는 xlsm파일 찾기
    for file in file_list:
        if file.startswith(meeting_times[-2:]) and os.path.isdir(os.path.join(shared_path, file)):
            shared_path = os.path.join(shared_path, file)

    file_list = os.listdir(shared_path)
    for file in file_list:
        if file.startswith("9. edms"):
            shared_path = os.path.join(shared_path, file)
    print("여기까지"+shared_path)
            
    file_list = os.listdir(shared_path)
    print(file_list)
    #9.edms폴더에 이번차시 파일이 없으면
    if file_list == []:
        file_name = meeting_times[-2:]+" "+meeting_times[-2:]+"차 위원회("+numbers_only_date+")"
        print("없다")
        os.makedirs(shared_path+"\\"+file_name)
        shared_path = os.path.join(shared_path, file_name)
        print(shared_path)
    else:
        for file in file_list:
            if file.startswith(meeting_times[-2:]):
                shared_path = os.path.join(shared_path, file)
                print(shared_path)
    print(shared_path)

    for i in range(len(GaNaDaAGSnum)):
        if not os.path.exists(shared_path+"\\"+GaNaDaAGSnum[i]):
            os.makedirs(shared_path+"\\"+GaNaDaAGSnum[i])
        else:
            pass
        
def The_Day_from_txt(meeting_times):
    now_time = datetime.datetime.now()
    shared_path = "\\\\공유폴더 ip\\DDA_인증심의\\"+str(now_time.year)+"년 위원회"
    
    print(shared_path)
    file_list = os.listdir(shared_path)
    # 이번 인증위 차시로 시작하는 xlsm파일 찾기
    for file in file_list:
        if file.startswith(meeting_times[-2:]) and os.path.isdir(os.path.join(shared_path, file)):
            shared_path = os.path.join(shared_path, file)    
    print(file_list)        
    file_list = os.listdir(shared_path)
    txt_files = [f for f in file_list if f.endswith(".txt")]
    print(file_list)
    
    # 파일을 열고 내용을 읽습니다.
    with open(txt_files[0], "r") as file:
        lines = file.readlines()
    
    print(lines[0])
    
    def convert_text_to_date(text):
        # 텍스트를 "."로 구분하여 리스트로 변환합니다.
        parts = text.split(".")

        for i in range(1, 3):
            if len(parts[i]) == 1:
                parts[i] = "0" + parts[i]

        return parts[0]+parts[1]+parts[2]
    
    converted_date = convert_text_to_date(lines[0])
    print(converted_date)
    return converted_date