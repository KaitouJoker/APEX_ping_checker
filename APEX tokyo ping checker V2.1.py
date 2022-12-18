from os import system
from math import ceil
from tqdm import tqdm
from pythonping import ping
from styleframe import StyleFrame
from numpy import array, delete, mean
from warnings import simplefilter as s_f
from openpyxl import load_workbook as l_w
from pandas import DataFrame, isnull, read_excel
from multiprocessing import Pool, cpu_count, freeze_support

# 초기 변수 설정
t_o = 'Time out'
t_o_ten = 'Time out 10 times'

# cmd의 cls
cls = lambda : system('cls')

# cmd의 pause
pause = lambda : system('pause')

# 사용자 입력 받아오는 함수
def user_input():
    
    # 올바른 입력을 받을때 까지 무한 반복
    while True:
        cls()
        # 선택지를 '\n'단위로 출력
        print('\n Made by MEZZ0\n'     ,
              ' Please select a region',
              ' 1. tokyo'              ,
              ' 2. taiwan'             ,
              ' 3. singapore\n'        ,
              sep = '\n'                )
        
        # 입력 받아 저장
        user_enter = input(' region: ')
        
        # 선택지 안에 입력이 있으면
        if user_enter in ['1', '2', '3']:
            
            # 입력 숫자로 변환 후 저장
            reg = int(user_enter)
            break
        else:
            print('잘못된 선택지 입니다. 다시 입력해 주세요.')
            pause()

    # 시트 이름 리스트
    region_list = ['tokyo', 'taiwan', 'singapore']
    
    # 시트 리스트에서 이름 값 가져오기
    sel_region = region_list[reg - 1]
    print(f'{sel_region} selected.')
    
    return sel_region

# 엑셀 시트에서 데이터 가져오기
def pd_read_excel(sel):
    # 선택한 시트의 데이터 읽어오기
    ip_pd = read_excel('./에펙 도쿄_타이완_싱가폴_홍콩.xlsx', sheet_name = sel, header = None, usecols=[*range(1, 9)])
    
    # Numpy 어레이로 변환
    ip_np = array(ip_pd)
    
    # 라우터 부분만 추출
    router_list        = ip_np[:, 0]
    
    delete_router_list = delete(ip_np, 0, 1).tolist()
    
    # 공백 아닌칸만 추출
    ip_list            = array([[j for j in i if not isnull(j)] for i in delete_router_list], dtype = list).tolist()
    
    return router_list, ip_list
    
# CPU 코어 세주는 함수
def count_cpu():
    n_cpu = cpu_count()
    # print(f'n_cpu = {n_cpu}')
    return n_cpu - 1

# ping 결과가 시간초과인지 지연시간값 인지 구별해 반환
def change_s(not_str):
    
    # 문자열 비교를 위한 변환
    str_s = str(not_str)
    
    # 결과가 'Request timed out'이면 'Time out' 반환
    if str_s == 'Request timed out':
        return t_o
    
    # 결과 문자열에서 지연시간 부분만 잘라 반환
    else:
        n_plus_index = str_s.index       ('n ') + 2
        ms_index     = str_s.index       ('ms')
        
        # 문자열 자르기
        split_str    = str_s[n_plus_index : ms_index]
        
        # float 형태로 저장
        to_int       = float(split_str)
        
        return to_int

# 핑 10번 보내고 지연시간을 리스트로 반환
def send_ping(ip):
    # tqdm.write(f'sending ping to {ip}')
    
    # 핑 10번 보내고 리스트로 만들기 (타임아웃 >= 400ms) (500ms 초 간격으로 보냄)
    ping_l1 = list(ping(ip, count = 10, size = 32, timeout = 0.4, interval = 0.5))
    
    # 시간초과 및 지연시간 리스트
    result_l1 = [change_s(data) for data in ping_l1]
    
    # 'Time out'이 하나라도 포함되어 있으면
    if t_o in result_l1:
        
        # 시간초과 뿐이면 'Time out 10 times'반환
        if len([i for i in result_l1 if type(i) == type(int())]) == 0:
            return t_o_ten
        
        # 아니면 유효한 부분만 연산
        else:
            
            # 유효한 지연시간만 추출
            digit_l1 = [i for i in result_l1 if i.isdigit()]
            
            # 최소값, 평균값 연산 후 저장
            min_p = min(digit_l1)
            avg_p = mean(digit_l1)
            
            # [최소값, 최대값, 평균값] 리스트 반환
            result = list(map(ceil, [min_p, t_o, avg_p]))

    # 지연시간이 모두 유효하면
    else:
        
        # 최소값, 평균값, 최대값 연산 후 저장
        min_p = min     (result_l1)
        avg_p = mean    (result_l1)
        max_p = max     (result_l1)
        
        # [최소값, 최대값, 평균값] 순으로된 리스트 반환
        result = list(map(ceil, [min_p, max_p, avg_p]))
    
    return result

# 리스트 별 요소 개수 세서 리스트로 반환
def get_index(ip_list):
    idx_list = [len(i) for i in ip_list]
    return idx_list

# 다중 코어로 핑 날리고 리스트로 반환
def send_ping_row(ips):
    
    # cpu 전체 쓰레드 수 - 1 구해 저장
    cores = count_cpu()
    
    # 코어 작업장 설정
    p = Pool(processes = cores)
    
    # 저장할 리스트 생성
    results = []
    
    # 진행바 설정
    with tqdm(total = len(ips)) as pbar:
        
        # 다중 코어로 핑 날리기
        for k in p.imap(send_ping, ips):
            
            # 반환값 추가
            results.append(k)
            
            # 진행바 업데이트
            pbar.update()
        
    return results

# 합쳐진 리스트 다시 원래 형태로 잘라 리스트로 반환
def slice_pings(origin_indexs, sum_list):
    sl_list = []
    
    # 초기 시작 인덱스 지정
    start_index = 0
    for k in origin_indexs:
        # 끝 인덱스 지정
        end_index = start_index + k
        
        # 리스트 자르기
        temp = sum_list[start_index : end_index]
        
        # 자른 리스트 추가
        sl_list.append(temp)
        
        # 향후 시작 인덱스 재지정
        start_index += k
    
    return sl_list

# ip_list 를 받아 각각의 핑 반환
def ip_list_to_ping_list(ip_list):
    
    # ip_list의 각각의 리스트 길이 기록
    idx_list = get_index(ip_list)
    
    # ip_list 하나로 합치기
    sum_list = sum(ip_list, [])
    
    # 핑 결과 전체 리스트
    ping_result = send_ping_row(sum_list)
    
    # 다시 본래의 형태로 자르기
    slice_ping_list = slice_pings(idx_list, ping_result)
    
    return slice_ping_list

# 행의 평균값 찾기
def find_avg_number(lists):
    digit = []
    
    # 한 행을 가져와 list_x 에 하나씩 넣기
    for list_x in lists:
        
        # list_x 가 타임아웃이면 넘기기
        if list_x == t_o_ten:
            continue
        
        else:
            
            # 아니면 평균값을 리스트에 추가
            digit.append(list_x[2])
    
    # 추가된게 아무것도 없으면
    if digit == []:
        
        # 타임아웃 반환
        return t_o
    
    else:
        
        # 아니면 평균값 연산 후 반환
        avg = ceil(mean(digit))
        return avg

# 한 칸에 넣을 ip와 지연시간 서식 만들기
def ip_and_pings(ip, digit):
    
    # 지연시간이 타임아웃이면
    if digit == t_o_ten:
        
        # ip 와 타임아웃으로 된 문자열 만들어 반환
        return f'{ip}\n{digit}'
    
    # 지연시간이 있으면
    else:
        # ip 와 최소, 최대, 평균 순으로 문자열로 만들어 반환
        return f"{ip}\n{', '.join(map(str, digit))}"

# 리스트를 스타일대로 가공
def temp3(router_list, ip_list, ping_list):
    
    # 넘파이 어레이로 변환
    ping_np = array(ping_list, dtype = list)
    
    # 아래의 과정은 본인이 만들었지만 과정이 기억이 안남
    temp2   = [[ip_and_pings(ip, digit) for ip, digit in zip(ip_x, list_x)] for ip_x, list_x in zip(ip_list, ping_np)]
    temp11  = [find_avg_number(x) for x in ping_np]
    temp12  = [ceil(mean(y)) if type(y) == type([]) else y for y in temp11]
    temp3   = [[ping, rtr] + np_list for ping, rtr, np_list in zip(temp12, router_list, temp2)]
    
    return temp3

# 데이터 프레임 엑셀에 저장
def data_frame(list, user_data):
    
    # 데이터 프레임에서 인덱스 제거
    df = DataFrame(list, index = None)
    
    # 가공된 데이터 프레임을 인덱스 & 헤더 제거후 엑셀에 저장
    StyleFrame(df).to_excel('./ping_result.xlsx', sheet_name = user_data, index = False, header = False).save()

# 엑셀 2차 가공
def wb():
    
    # 엑셀파일 로드
    wb = l_w('./ping_result.xlsx')
    ws = wb.active
    
    # A열 너비 설정
    ws.column_dimensions['A'].width = 8.63
    
    # 나머지 열 너비 설정
    for column in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
        ws.column_dimensions[column].width = 17
        
    # 저장
    wb.save('./ping_result.xlsx')
    print('all ping data save done. to "\\ping_result.xlsx"')
    
if __name__ == "__main__":
    # FutureWarning 제거
    s_f(action = 'ignore', category = FutureWarning)
    
    # Windows 에서 multiprocess의 메인 프로세스가 꺼지는 현상 방지
    freeze_support()
    
    # 사용자 입력 받기
    user_data = user_input()
    
    # 콘솔창 지우고 출력
    cls()
    print(f'\n {user_data} selected.', ' reading excel file...', sep = '\n')
    
    # 엑셀 파일 읽어와 가공
    router_list, ip_list = pd_read_excel(user_data)
    print(' excel file read completed.')
    
    # 각 ip별로 ping 보내고 저장
    print(' Start sending pings with multiple cores...')
    ping_list = ip_list_to_ping_list(ip_list)
    
    # 엑셀 1차 가공
    processed_data = temp3(router_list, ip_list, ping_list)
    data_frame(processed_data, user_data)
    
    # 엑셀 2차 가공
    wb()
    
    # 완료
    pause()
