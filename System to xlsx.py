'''
1. 모니터가 1대면 오류가 발생하나 정보 추출은 정상적임
2. 모니터 정보는 2대까지만 저장
3. 저장한 자료는 upload_path에 저장됨 [권한필요]
4. 저장 데이터는 컴퓨터이름/도메인/로그인계정/IP주소/데스크탑SN/모니터1SN/모니터2SN
'''

import pandas as pd
import subprocess, os, socket
# 파이썬 오류 무시하기
import warnings
warnings.filterwarnings('ignore', category=DeprecationWarning)
# 모니터 1 정보
monitor1 = subprocess.run('powershell \n [System.Text.Encoding]::ASCII.GetString($(Get-CimInstance WmiMonitorID -Namespace root\wmi)[0].SerialNumberID -notmatch 0)', stdout=subprocess.PIPE)
monitor1 = monitor1.stdout.decode('utf-8')
monitor1 = monitor1.replace('\n', '') # 엔터제거
monitor1 = monitor1.replace('\r', '') # \r제거
# 모니터 2 정보
monitor2 = subprocess.run('powershell \n [System.Text.Encoding]::ASCII.GetString($(Get-CimInstance WmiMonitorID -Namespace root\wmi)[1].SerialNumberID -notmatch 0)', stdout=subprocess.PIPE)
monitor2 = monitor2.stdout.decode('utf-8')
monitor2 = monitor2.replace('\n', '') # 엔터제거
monitor2 = monitor2.replace('\r', '') # \r제거
# 컴퓨터 이름
hostname = socket.gethostname()
# 도메인 정보
domain = subprocess.run('wmic computersystem get domain /format:list', stdout=subprocess.PIPE)
domain = domain.stdout.decode('utf-8')
domain = domain.replace('\n','') # 엔터제거
domain = domain.replace('\r', '') # \r 불필요 항목 제거
domain = domain[7:] # 앞자리 제거
# 사용자 정보
id = subprocess.run('whoami', stdout=subprocess.PIPE)
id = id.stdout.decode('utf-8')
id = id.replace('\r', '') # \r 불필요 항목 제거
id = id[id.find('\\')+1:].replace('\n','')
# IP 정보
ip = (socket.gethostbyname(socket.gethostname()))
# 데스크탑 정보
desktop = subprocess.run('WMIC CSPRODUCT GET IDENTIFYINGNUMBER', stdout=subprocess.PIPE)
desktop = desktop.stdout.decode('utf-8')
desktop = desktop.replace('\n', '') # 엔터제거
desktop = desktop.replace('\r', '') # \r 불필요 항목 제거
desktop = desktop.replace(' ','') # 공백제거
desktop = desktop[17:] # 슬라이싱

name = ['PC NAME',
        'DOMAIN',
        'ID',
        'IP ADDRESS',
        'DESKTOP',
        'MONITOR 1',
        'MONITOR 2']
info = [[hostname, 
        domain,
        id, 
        ip, 
        desktop, 
        monitor1, 
        monitor2]]
df = pd.DataFrame(info, columns = name)
print(df)
# ======================================================
# 저장경로

directory = os.path.join(os.environ['USERPROFILE'], 'Desktop') # 바탕화면
save_path = directory+'\\'+id+'.xlsx'
# DataFrame을 저장
df.to_excel(save_path)
# ======================================================


