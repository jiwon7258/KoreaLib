'''
requests : 웹페이지 객체를 가져오는 라이브러리
BeautifulSoup : 웹페이지 파싱 라이브러리
pandas : 데이터 관리 라이브러리 (엑셀과 비슷)

'''
 
import requests
from bs4 import BeautifulSoup
import os
import pandas as pd 
import numpy

def main() :
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # 스마트도서대출반납기 리스트를 가져온다

    # res = requests.get('http://163.152.81.120/wssl/WSSLLOG.asp?QUERY=Y')
    # soup = BeautifulSoup(res.content, 'html.parser')
    # tables = soup.select('table')

    with open(BASE_DIR + os.path.sep + "WSSL1.html") as file :
        soup = BeautifulSoup(file, "html.parser");
        tables = soup.select('table');

    # 테이블 객체를 문자열로 변환
    wssl_html = str(tables)
    wssl_html_list = pd.read_html(wssl_html)

    # 데이터프레임 선택하기
    wssl_df = wssl_html_list[2]

    # 0, 1 row를 삭제한다 (필요없는 정보)
    for i in range (0,2) :
        wssl_df = wssl_df.drop(i)

    wssl_df.to_excel('wssl1_list.xlsx', sheet_name = 'sheet1')


    # Tulip의 중앙도서관 '도착통보' 리스트를 불러온다
    tulip_df = pd.read_excel(BASE_DIR + os.path.sep + 'a.xlsx', skiprows=[0,1,2])

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

    # 인포 소관 자료실이 아닌 경우 삭제한다
    lib_list = ['제2자료실(3층)', '제3자료실(4층)', '서고6층', '서고7층', '서고2층(대형)']

    for idx in tulip_df.index :
        lib_data = tulip_df.loc[idx, '자료실']

        if lib_data not in lib_list :
            tulip_df = tulip_df.drop(idx)


    print(tulip_df)

    # 최종 결과를 엑셀로 출력한다
    tulip_df.to_excel('예약도서_리스트.xlsx', sheet_name = 'sheet1')
