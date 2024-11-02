import os
import csv
import pandas as pd
import joblib
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import MultinomialNB
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import smtplib
from email.mime.text import MIMEText
import schedule
import time

def loadData(filePath='email_snu.csv'):
    dataFrame = pd.read_csv(filePath)
    dataFrame['Sender'].fillna('', inplace=True)
    dataFrame['Subject'].fillna('', inplace=True)
    dataFrame['text'] = dataFrame['Sender'] + " " + dataFrame['Subject']
    X = dataFrame['text']
    y = dataFrame['Label']
    return X, y

def trainModel(X, y):
    tfidfVectorizer = TfidfVectorizer(stop_words='english', max_df=0.95)
    XTfidf = tfidfVectorizer.fit_transform(X)
    XTrain, XTest, yTrain, yTest = train_test_split(XTfidf, y, test_size=0.2, random_state=42)
    
    naiveBayesModel = MultinomialNB()
    naiveBayesModel.fit(XTrain, yTrain)
    
    joblib.dump(naiveBayesModel, 'emailClassifierModel.joblib')
    joblib.dump(tfidfVectorizer, 'tfidfVectorizer.joblib')
    
    yPred = naiveBayesModel.predict(XTest)
    print("Accuracy:", accuracy_score(yTest, yPred))
    print("\nClassification Report:\n", classification_report(yTest, yPred))
    print("\nConfusion Matrix:\n", confusion_matrix(yTest, yPred))

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
def gmailAuthenticate():
    credentials = None
    if os.path.exists('token.json'):
        credentials = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(credentials.to_json())
    return build('gmail', 'v1', credentials=credentials)

def checkEmail(service):
    try:
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="is:unread").execute()
        messages = results.get('messages', [])
        
        if not messages:
            print("No new emails.")
        else:
            for message in messages:
                messageId = message['id']
                messageData = service.users().messages().get(userId='me', id=messageId).execute()
                payload = messageData['payload']
                headers = payload.get('headers', [])

                sender = next((header['value'] for header in headers if header['name'] == 'From'), None)
                subject = next((header['value'] for header in headers if header['name'] == 'Subject'), None)
                
                if sender and subject:
                    label = classifyEmail(sender, subject)
                    print(f"Email from {sender} with subject '{subject}' classified as: {label}")

                    if label != 'spam':
                        sendNotification(sender, subject, label)

                    service.users().messages().modify(userId='me', id=messageId, body={'removeLabelIds': ['UNREAD']}).execute()
                else:
                    print("Email missing sender or subject.")
    except Exception as e:
        print(f"An error occurred while checking email: {e}")

def loadModel():
    naiveBayesModel = joblib.load('emailClassifierModel.joblib')
    tfidfVectorizer = joblib.load('tfidfVectorizer.joblib')
    return naiveBayesModel, tfidfVectorizer

naiveBayesModel, tfidfVectorizer = loadModel()

def classifyEmail(sender, subject):
    emailContent = f"{sender} {subject}"
    emailTfidf = tfidfVectorizer.transform([emailContent])
    label = naiveBayesModel.predict(emailTfidf)[0]
    return label

def sendNotification(sender, subject, label):
    smtpServer = 'smtp.gmail.com'
    smtpPort = 587
    fromEmail = 'wkdwlgh03@gmail.com'
    fromPassword = os.getenv('EMAIL_PASSWORD')

    message = MIMEText(f"New email received.\n\nFrom: {sender}\nSubject: {subject}\nClassification: {label}")
    message['Subject'] = "New Email Notification"
    message['From'] = fromEmail
    message['To'] = fromEmail

    try:
        with smtplib.SMTP(smtpServer, smtpPort) as server:
            server.starttls()
            server.login(fromEmail, fromPassword)
            server.sendmail(fromEmail, fromEmail, message.as_string())
    except Exception as e:
        print(f"Failed to send notification email: {e}")

if not os.path.exists('emailClassifierModel.joblib') or not os.path.exists('tfidfVectorizer.joblib'):
    X, y = loadData()
    trainModel(X, y)

service = gmailAuthenticate()
schedule.every(1).minutes.do(checkEmail, service)

while True:
    schedule.run_pending()
    time.sleep(1)
