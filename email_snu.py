import os
import csv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def getEmails(service, userId='me', maxResults=1000):
    emails = []
    nextPageToken = None
    while len(emails) < maxResults:
        # 메시지 리스트 요청
        response = service.users().messages().list(
            userId=userId,
            maxResults=min(maxResults - len(emails), 500),  # 최대 500개의 메시지씩 가져옴
            pageToken=nextPageToken
        ).execute()

        messages = response.get('messages', [])
        emails.extend(messages)

        nextPageToken = response.get('nextPageToken')
        if not nextPageToken:
            break

    return emails

def saveEmailsToCsv(emails, service, csvFilename="emails.csv"):
    with open(csvFilename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["From", "Subject"])  # CSV 파일의 헤더 설정

        for email in emails:
            # 메일 데이터 가져오기
            msg = service.users().messages().get(userId='me', id=email['id']).execute()

            headers = msg['payload']['headers']
            msgFrom = next((header['value'] for header in headers if header['name'] == 'From'), 'N/A')
            msgSubject = next((header['value'] for header in headers if header['name'] == 'Subject'), 'N/A')

            # CSV 파일에 쓰기
            writer.writerow([msgFrom, msgSubject])

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    # 이메일 가져오기
    emails = getEmails(service, maxResults=6500)
    print(f"Total emails fetched: {len(emails)}")

    # 이메일을 CSV 파일로 저장
    saveEmailsToCsv(emails, service)

if __name__ == '__main__':
    main()
