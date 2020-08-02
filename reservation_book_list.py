'''
requests : 웹페이지 객체를 가져오는 라이브러리
BeautifulSoup : 웹페이지 파싱 라이브러리
pandas : 데이터 관리 라이브러리 (엑셀과 비슷)

'''
 
import requests
from bs4 import BeautifulSoup
import os
import pandas as pd 
# import numpy
import progress
import time
import pkg_resources.py2_warn
import sys

progress.printProgress(0,100)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 스마트도서대출반납기 리스트를 가져온다
try :
    res = requests.get('http://163.152.81.120/wssl/WSSLLOG.asp?QUERY=Y')
    soup = BeautifulSoup(res.content, 'html.parser')
    tables = soup.select('table')
except Exception as e :
    # print(e)
    print("에러. WSS1에 연결할 수 없습니다. 엔터 키를 누르면 종료합니다.")
    a = input()
    if (a or a =='') :
        sys.exit()


# try :
#     with open(BASE_DIR + os.path.sep + "WSSL1.html") as file :
#         soup = BeautifulSoup(file, "html.parser");
#         tables = soup.select('table');
# except Exception as e :
#     # print(e)
#     print("에러. 파일을 불러올 수 없습니다. 엔터 키를 누르면 종료합니다.")
#     a = input()
#     if (a or a =='') :
#         quit()


progress.printProgress(10,100)


# 테이블 객체를 문자열로 변환
wssl_html = str(tables)
wssl_html_list = pd.read_html(wssl_html)

# 데이터프레임 선택하기
wssl_df = wssl_html_list[2]

progress.printProgress(30,100)

# 0, 1 row를 삭제한다 (필요없는 정보)
for i in range (0,2) :
    wssl_df = wssl_df.drop(i)

wssl_df.to_excel('wssl1_list.xlsx', sheet_name = 'sheet1')

progress.printProgress(60,100)

try:
    # Tulip의 중앙도서관 '도착통보' 리스트를 불러온다
    tulip_df = pd.read_excel(BASE_DIR + os.path.sep + 'a.xls', skiprows=[0,1,2])
except Exception as e:
    print("에러\na.xls를 불러오는 과정에 오류가 발생했습니다")
    print("a.xls를 다시 생성해주세요")
    print('엔터키를 누르면 종료합니다')
    a = input()
    if (a or a =='') :
        quit()

progress.printProgress(75,100)

# tulip_df - WSSL 스대기 리스트
wssl_list = wssl_df[1]
for wssl_num in wssl_list :

    if (wssl_num == '등록번호') :
        continue

    wssl_num = int(wssl_num)

    for idx in tulip_df.index :
        tulip_num = tulip_df.loc[idx, '반납도서등록번호']
        if (tulip_num == wssl_num) :
            tulip_df = tulip_df.drop(idx)


progress.printProgress(95,100)

# 인포 소관 자료실이 아닌 경우 삭제한다
# lib_list = ['제2자료실(3층)', '제3자료실(4층)', '서고6층', '서고7층', '서고2층(대형)']

# for idx in tulip_df.index :
#     lib_data = tulip_df.loc[idx, '자료실']

#     if lib_data not in lib_list :
#         tulip_df = tulip_df.drop(idx)


# print(tulip_df)

# 최종 결과를 엑셀로 출력한다
tulip_df.to_excel('예약도서_리스트.xlsx', sheet_name = 'sheet1')
progress.printProgress(100,100)

print("출력완료 \n예약도서_리스트.xlsx를 확인하세요")
print('엔터키를 누르면 종료합니다')
a = input()
if (a or a =='') :
    sys.exit()

# 'pkg_resources.py2_warn'
