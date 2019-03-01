# 채점할 때 도움이 되면 좋겠습니다:)
# 2019.03.02. Cheonuk
import os
import shutil
import openpyxl         #데이터를 엑셀파일로 받았으므로, openpyxl 외부 패키지를 설치해 사용합니다.

#분반-조교별 채점할 학생들 학번 모음
TA = {'5김예원': 0, '5김경혜': 0, '5김경렬': 1,'5홍예빈': 1,'5성제현': 2,'5김기태': 2,
      '14김금비': 3, '14박상준': 3, '14문태근': 4,'14김민경': 4,'14성제현': 5,'14이은정': 5,
      '17김금비': 6, '17우경찬': 6, '17김기태': 7, '17백도협': 7, '17성제현': 8, '17이은정': 8}

def student_list():
    '''
    본인이 담당하는 학생 명단을 리스트화 하여 과제 파일의 이름과 비교할 수 있게 하는
    전처리 함수입니다.
    '''
    wb = openpyxl.load_workbook('student_id.xlsx')  # 엑셀파일 열기
    ws = wb.active          #파일내 Active Sheet로 접근
    for r in ws.rows:       #엑셀을 한줄씩 읽습니다
        if r[TA[TA_Name]].value != None:
            STUDENT_ID.append(str(r[TA[TA_Name]].value))
    STUDENT_ID.pop(0)


if __name__ == "__main__":
    print('과제를 분류해 자신의 학생들 과제파일만 저장해 줍니다.')
    #print('과제 미제출자들 학번을 알려줍니다.')
    TA_Name = input('\n나눌 과제의 분반과 조교이름을 입력하세요.(띄어쓰기 ㄴㄴ)\n(예시 : 만약 17분반 이라면 -> 17박천욱 ) :')
    DIR_FROM = input('과제들이 들어있는 파일 절대경로를 입력하세요 :')
    DIR_TO = input('자신의 학생들의 과제파일만을 저장할 폴더에 대한 절대 경로를 입력하세요 :')

    STUDENT_ID = []
    student_list()
    STUDENT_COUNT=len(STUDENT_ID)

    for (path, dirt, files) in os.walk(DIR_FROM):
        for filename in files:
            if filename[:10] in STUDENT_ID:
                shutil.copy(os.path.join(path, filename), DIR_TO)
                STUDENT_ID.remove(filename[:10])

    print('\n총원 [',STUDENT_COUNT, '] 명 중 [', STUDENT_COUNT-len(STUDENT_ID), '] 명이 과제를 제출했습니다.' )
    print('미제출자는 [', len(STUDENT_ID), '] 명 입니다.')
    print('\n#############미제출자 학번###################')
    for N_STUDENT in STUDENT_ID:
        print(N_STUDENT)