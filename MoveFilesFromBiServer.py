"""
Power BI 서버에서 파일을 다운로드하여 로컬에 저장하는 스크립트입니다.

작성자: 김세희
작성일: 2024-06-04
최종 수정일: 2024-06-04

실행 방법:
1. 스크립트를 실행하기 전에 requests_ntlm, requests, urllib, datetime 패키지를 설치해야 합니다. (아래 '참고' 참조)
2. cmd 창에서 스크립트가 있는 디렉토리로 이동한 후 'python MoveFilesFromBiServer.py' 명령을 실행합니다.
3. 프롬프트에 요청된 정보를 입력합니다.

사용 방법:
1. 설정 정보
    - BASE_URL: Power BI 서버 URL
    - TIMEOUT: 요청 제한 시간 (초 단위)

2. extractDateFromFilename를 사용하여 파일명에서 날짜 값을 추출합니다.
    - 파일명 예시: 'SBL_P5_보성엔지니어링_배관_240306_Ver.1.0.2.xlsx'
    - 파일명을 '_'로 분리하여 5번째 요소를 추출합니다.
3. main 함수를 실행하여 스크립트를 실행합니다.

참고:
- requests_ntlm 패키지 설치 필요: pip install requests_ntlm
- requests 패키지 설치 필요: pip install requests
- urllib 패키지 설치 필요: pip install urllib
- datetime 패키지 설치 필요: pip install datetime

* 주의: 이 스크립트는 Power BI 서버에 대한 권한이 필요합니다.
"""


import requests
from requests_ntlm import HttpNtlmAuth
import os
from urllib.parse import quote
from datetime import datetime

# 서버 설정
BASE_URL = "https://pbi.secl.co.kr"

# 타임아웃 설정 (초 단위)
TIMEOUT = 30

# URL 인코딩 함수
def encodeUrl(path):
    return quote(path, safe='/:?=&')

# 파일명에서 날짜 값 추출 함수
def extractDateFromFilename(fileName):
    # 파일명 예시: 'SBL_P5_보성엔지니어링_배관_240306_Ver.1.0.2.xlsx'
    try:
        strDate = fileName.split('_')[4]  # 파일명을 '_'로 분리하여 5번째 요소 추출
        strDate = strDate[:6]  # 앞에서부터 6자리만 추출 (시간 정보 제외)
        objDate = datetime.strptime(strDate, "%y%m%d")  # 날짜 문자열을 날짜 객체로 변환
        return objDate
    except (IndexError) as e:
        print(f"Error extracting date from filename: {fileName}")
        return None
    except ValueError as e:
        print(f"Incorrect date format in filename: {fileName}")
        return None

# 폴더 내 항목 목록 가져오기
def getFolderItems(folder_path, auth):
    encodedPath = encodeUrl(folder_path)
    try:
        response = requests.get(f"{BASE_URL}{encodedPath}", auth=auth, timeout=TIMEOUT)
        if response.status_code == 200:
            print("Successfully retrieved folder items.")
            return response.json().get('value', [])  # JSON 응답의 'value' 키에서 항목 목록 추출
        else:
            print(f"Failed to retrieve folder items from {BASE_URL}{encodedPath}: {response.status_code}")
            return []
    except requests.exceptions.RequestException as e:
        print("------Error------")
        print(f"Error retrieving folder items from {BASE_URL}{encodedPath}: {e}")
        return []

# 항목 처리 함수
def processItems(items, auth, currentFolderPath="", startDate=None, endDate=None):
    for item in items:
        if isinstance(item, dict) and 'Id' in item and 'Name' in item and 'Type' in item:
            itemId = item['Id']
            itemName = item['Name']
            itemType = item['Type']
            print(f"Processing item: {itemName} (Type: {itemType})")

            if itemType == 'Folder':
                # 폴더일 경우 해당 폴더를 로컬에 생성하고 하위 파일 이동
                folderPath = os.path.join(currentFolderPath, itemName)
                os.makedirs(folderPath, exist_ok=True)  # 폴더 생성 (이미 존재할 경우 무시)
                print(f"Created folder: {folderPath}")

                # 하위 항목 가져오기
                subItems = getFolderItems(f"/reports/api/v2.0/Folders({itemId})/CatalogItems", auth)
                processItems(subItems, auth, folderPath, startDate, endDate)
            elif itemType == 'ExcelWorkbook':
                # 엑셀 워크북일 경우
                downloadExcelContent(itemId, itemName + ".xlsx", currentFolderPath, auth, startDate, endDate)
            elif itemType in ['Resource', 'Pdf', 'Png', 'Jpeg']:
                # 리소스 파일일 경우
                downloadFile(f"{BASE_URL}/reports/api/v2.0/CatalogItems({itemId})/Content/$value", itemName, currentFolderPath, auth, startDate, endDate)
            else:
                print(f"Unknown item type: {itemType}")
        else:
            print(f"Unexpected item format: {item}")

# 엑셀 파일 다운로드 및 저장 함수
def downloadExcelContent(workbookId, fileName, folderPath, auth, startDate=None, endDate=None):
    contentUrl = f"{BASE_URL}/reports/api/v2.0/CatalogItems({workbookId})/Content/$value"
    try:
        response = requests.get(contentUrl, auth=auth, timeout=TIMEOUT)
        if response.status_code == 200:
            # 파일명에서 날짜 값 추출
            fileDate = extractDateFromFilename(fileName)
            # 시작일과 종료일 사이에 있는 파일만 다운로드
            if (fileDate and (not startDate or startDate <= fileDate) and (not endDate or fileDate <= endDate)):
                localFilePath = os.path.join(folderPath, fileName)
                with open(localFilePath, 'wb') as file:
                    file.write(response.content)
                print(f"Downloaded and saved Excel content to '{localFilePath}'")
            else:
                print(f"Skipping file '{fileName}' outside of date range.")
        else:
            print(f"Failed to download content from '{contentUrl}': {response.status_code}")
    except requests.exceptions.RequestException as e:
        print("------Error------")
        print(f"Error downloading content from '{contentUrl}': {e}")

# 리소스 파일 다운로드 함수
def downloadFile(fileUrl, fileName, folderPath, auth, startDate=None, endDate=None):
    try:
        response = requests.get(fileUrl, auth=auth, timeout=TIMEOUT)
        if response.status_code == 200:
            # 파일명에서 날짜 값 추출
            fileDate = extractDateFromFilename(fileName)
            # 시작일과 종료일 사이에 있는 파일만 다운로드
            if (fileDate and (not startDate or startDate <= fileDate) and (not endDate or fileDate <= endDate)):
                localFilePath = os.path.join(folderPath, fileName)
                with open(localFilePath, 'wb') as file:
                    file.write(response.content)
                print(f"Downloaded '{fileUrl}' to '{localFilePath}'")
            else:
                print(f"Skipping file '{fileName}' outside of date range.")
        else:
            print(f"Failed to download '{fileUrl}': {response.status_code}")
    except requests.exceptions.RequestException as e:
        print("------Error------")
        print(f"Error downloading '{fileUrl}': {e}")

# 사용자 정보 입력 및 처리 함수
def getUserInput():
    print("----------------------") 
    username = input("Enter your username: ")
    password = input("Enter your password: ")
    print("----------------------")
    return username, password

# 접속할 폴더 경로 및 로컬 저장 경로 입력 및 처리 함수
def getFolderPaths():
    print("----------------------")
    remotePath = input("Enter the remote folder path(ex: /Bio/RPA%20TEST/Test/Test2): ")
    localPath = input("Enter the local folder path(ex: C:/Users/user/Desktop/01 송도): ")
    print("----------------------")
    return remotePath, localPath

# 입력받은 날짜를 datetime 객체로 변환하는 함수
def getDateRange():
    print("----------------------")
    strStartDate = input("Enter the start date (YYYY-MM-DD): ")
    strEndDate = input("Enter the end date (YYYY-MM-DD): ")
    print("----------------------")
    try:
        startDate = datetime.strptime(strStartDate, "%Y-%m-%d")
        endDate = datetime.strptime(strEndDate, "%Y-%m-%d")
        return startDate, endDate
    except ValueError as e:
        print(f"Invalid date format. Please enter dates in the format 'YYYY-MM-DD'.")
        return None, None

def main():
    # 조회하는 경로 안내(https://pbi.secl.co.kr) 및 사용자 정보 입력 안내
    print("----------------------")
    print("")
    print("✼ ҉ ✼ Welcome to the Power BI Server File Downloader! ✼ ҉ ✼")
    print("This script will download files from the Power BI server(https://pbi.secl.co.kr)." )
    print("Please enter your username and password to authenticate.")
    print("")

    # 사용자 정보 입력
    userName, password = getUserInput()

    try:
        auth = HttpNtlmAuth(userName, password)
        print("Successfully authenticated!")

        # 폴더 정보 요청
        remotePath, localPath = getFolderPaths()

        # 넘겨 받은 localPath가 존재하지 않을 경우 프로그램 종료
        if not os.path.exists(localPath):
            print("Local path does not exist. Exiting program...")
            return
        
        fullRemotePath = f"/reports/api/v2.0/Folders(Path='{remotePath}')/CatalogItems"
        print("Requesting folder information from", BASE_URL + fullRemotePath)

        folderItems = getFolderItems(fullRemotePath, auth)
        if folderItems:
            # 날짜 범위 입력
            startDate, endDate = getDateRange()

            if startDate and endDate:
                # 진행 시작 구분용
                print("----------------------")
                print("|                    |")
                print("| **Process Start**  |")
                print("|                    |")
                print("----------------------")

                processItems(folderItems, auth, localPath, startDate, endDate)
                print("----------------------")
                print("☆♬○♩●♪✧♩((ヽ( ᐛ )ﾉ))♩✧♪●♩○♬☆")
                print("Successfully processed all items! Bye!")
                print("----------------------")

            else:
                print("Date range not specified. Exiting program...")
        else:
            print("No items retrieved.")
    except requests.exceptions.RequestException as e:
        print("------Error------")
        print(f"Error authenticating: {e}")

if __name__ == "__main__":
    main()
