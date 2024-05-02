import os
import sys
import urllib.request
import pandas as pd
import json
import re

# 네이버 API 클라이언트 ID와 시크릿
client_id = 'uxBMxwJzw_zJHEfayokT'
client_secret = 'Naxf078JUD'

# 사용자로부터 검색어와 엑셀 파일명, 가져올 결과 수를 입력받음
print("\n\n-------------------------------------------------------------")
query = input("검색어 입력  :  ")
query_encoded = urllib.parse.quote(query)
my_xlsx = input("엑셀 파일로 저장할 파일 이름 입력   :  ")
end = int(input("총 몇개의 결과를 가져오겠습니까?   :  "))
print("\n\n-------------------------------------------------------------")

# 한 번에 가져올 결과 수
display = 100

# 결과를 저장할 리스트
results = []

# 요청 헤더에 클라이언트 ID와 시크릿 추가
headers = {
    "X-Naver-Client-Id": client_id,
    "X-Naver-Client-Secret": client_secret
}

# 결과를 가져오기 위해 반복문 사용
for start in range(1, end + 1, display):
    # 네이버 검색 API 호출을 위한 URL 생성
    url = f"https://openapi.naver.com/v1/search/blog.json?query={query_encoded}&display={display}&start={start}"

    # HTTP 요청을 보내고 응답을 받음
    req = urllib.request.Request(url, headers=headers)
    response = urllib.request.urlopen(req)

    # 응답 코드 확인
    if response.getcode() == 200:
        # 응답 본문을 읽어옴
        response_dict = json.loads(response.read().decode("utf-8"))
        # 검색 결과 중 내용만 추출
        items = response_dict.get("items", [])

        # 각 항목에 대해 결과 리스트에 추가
        for item in items:
            result = {
                "title": re.sub('<.+?>', '', item["title"]),
                "link": item["link"],
                "description": re.sub('<.+?>', '', item["description"]),
                "bloggername": item["bloggername"],
                "bloggerlink": item["bloggerlink"],
                "postdate": item["postdate"]
            }
            results.append(result)

# 결과를 데이터프레임으로 변환
result_df = pd.DataFrame(results)

# 데이터프레임을 엑셀 파일로 저장
output_path = f"./../TestData/{my_xlsx}.xlsx"
result_df.to_excel(output_path, index=False)

print("저장이 완료되었습니다.")
