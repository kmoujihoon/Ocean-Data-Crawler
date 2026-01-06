import requests
import pandas as pd
import math
import os
import time

# ==========================================
# 1. 설정 (URL 및 인증키)
# ==========================================
url = "https://api.odcloud.kr/api/3056346/v1/uddi:3f9e30f3-0056-4a93-9f1c-df725b5200f5"
service_key = "91f9446255cc0889b1c7527d5342a2ae5e2c9c6a1565ef9ffb427dc388834ec1"

print("데이터 수집 프로그램을 시작합니다...")

# ==========================================
# 2. 데이터 전체 개수 파악
# ==========================================
try:
    params = {
        "serviceKey": service_key,
        "page": "1",
        "perPage": "1",
        "returnType": "JSON"
    }
    response = requests.get(url, params=params)
    total_count = response.json()['totalCount']
    print(f"전체 데이터 개수 확인됨: {total_count}개")
except Exception as e:
    print("초기 연결 실패! 인터넷 연결이나 키를 확인하세요.")
    print(f"에러 내용: {e}")
    total_count = 0

# ==========================================
# 3. 전체 데이터 수집 (반복문)
# ==========================================
if total_count > 0:
    all_data = []
    chunk_size = 2000
    total_pages = math.ceil(total_count / chunk_size)

    print(f"총 {total_pages}번 나누어 다운로드를 시작합니다.")

    for page in range(1, total_pages + 1):
        # 진행상황 표시
        print(f"[ {page} / {total_pages} ] 페이지 다운로드 중...", end="\r")
        
        params = {
            "serviceKey": service_key,
            "page": page,
            "perPage": chunk_size,
            "returnType": "JSON"
        }
        
        try:
            res = requests.get(url, params=params)
            data = res.json()['data']
            all_data.extend(data)
            time.sleep(0.1) # 서버 과부하 방지
        except Exception as e:
            print(f"\n{page}페이지 다운로드 실패: {e}")

    print(f"\n수집 완료! 총 {len(all_data)}개의 데이터를 확보했습니다.")

    # ==========================================
    # 4. 엑셀 저장 (경로: 바탕 화면)
    # ==========================================
    df = pd.DataFrame(all_data)

    # 사용자 홈 폴더 + "바탕 화면" 경로 지정
    home_dir = os.path.expanduser("~")
    save_dir = os.path.join(home_dir, "바탕 화면")
    
    # 만약 "바탕 화면" 폴더가 없으면 에러가 날 수 있으니 체크 후 생성 시도
    if not os.path.exists(save_dir):
        # 바탕 화면 폴더가 없으면 그냥 홈 폴더에 저장
        save_dir = home_dir

    file_name = "해양사고통계_전체데이터.xlsx"
    full_path = os.path.join(save_dir, file_name)

    print(f"저장을 시도합니다: {full_path}")
    
    try:
        df.to_excel(full_path, index=False)
        print("[성공] 엑셀 파일 저장이 완료되었습니다!")
        os.startfile(save_dir) # 저장된 폴더 열기
    except Exception as e:
        print(f"저장 중 에러 발생: {e}")
        print("엑셀 파일을 열어두셨다면 끄고 다시 실행해주세요.")

else:
    print("데이터를 찾을 수 없어 종료합니다.")