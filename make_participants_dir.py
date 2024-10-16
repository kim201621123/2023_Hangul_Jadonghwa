import os
import datetime
import sys

now_time = datetime.datetime.now()
    
shared_path = "\\\\공유폴더 주소 ip\\DDA 심의\\"+str(now_time.year)+"년 위원회"

# 각 문자열의 앞글자끼리 비교합니다.
for i in range(len(indivisual_participants) - 1):
    for j in range(i + 1, len(indivisual_participants)):
        if indivisual_participants[i][0] < indivisual_participants[j][0]:
            indivisual_participants[i], indivisual_participants[j] = indivisual_participants[j], indivisual_participants[i]
#한글 오름차순으로 정렬
indivisual_participants = indivisual_participants[::-1]
if "자차카" in indivisual_participants:
    indivisual_participants = [x for x in indivisual_participants if x not in ["자차카"]]
    indivisual_participants.insert(0, "자차카")
indivisual_participants = [x + "위원" for x in indivisual_participants]
indivisual_participants = [x for x in indivisual_participants if x not in ["구구구위원"]]
indivisual_participants.append("구구구단장")