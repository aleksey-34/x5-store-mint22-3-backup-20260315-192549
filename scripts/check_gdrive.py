"""Check what files are accessible via Google Drive service account."""
import sys
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

KEY = r'C:\Users\Aleksei\Downloads\frobloc-proj-4060d4d1867c.json'
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

try:
    creds = Credentials.from_service_account_file(KEY, scopes=SCOPES)
    service = build('drive', 'v3', credentials=creds)

    # List all files the SA has access to
    results = service.files().list(
        pageSize=100,
        fields='nextPageToken, files(id, name, mimeType, modifiedTime, size)',
        orderBy='modifiedTime desc'
    ).execute()

    files = results.get('files', [])
    if not files:
        print('=== НЕТ ДОСТУПНЫХ ФАЙЛОВ ===')
        print('Сервисный аккаунт не имеет доступа ни к каким файлам.')
        print('Нужно расшарить нужные файлы/папки на: x5-store@frobloc-proj.iam.gserviceaccount.com')
    else:
        print(f'=== НАЙДЕНО ФАЙЛОВ: {len(files)} ===\n')
        for f in files:
            mime = f.get('mimeType', '')
            ftype = mime.split('.')[-1] if '.' in mime else mime
            size = f.get('size', '—')
            mod = f.get('modifiedTime', '—')[:10]
            print(f"  [{ftype}] {f['name']}")
            print(f"     id={f['id']}  modified={mod}  size={size}")
        if results.get('nextPageToken'):
            print(f'\n(есть ещё файлы, показаны первые 100)')

except HttpError as e:
    print(f'HTTP ошибка: {e}')
    sys.exit(1)
except Exception as e:
    print(f'Ошибка: {e}')
    sys.exit(1)
