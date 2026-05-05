import os
import re
import base64
import datetime

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'


def _check_google_libs():
    try:
        import google.auth.transport.requests  # noqa: F401
        import google.oauth2.credentials        # noqa: F401
        import google_auth_oauthlib.flow        # noqa: F401
        import googleapiclient.discovery        # noqa: F401
        return True
    except ImportError:
        return False


def get_gmail_service():
    if not _check_google_libs():
        raise ImportError(
            "Google APIライブラリが未インストールです。\n"
            "コマンドプロンプトで以下を実行してください:\n\n"
            "pip install google-auth-oauthlib google-api-python-client"
        )

    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"{CREDENTIALS_FILE} が見つかりません。\n\n"
                    "【Gmail API 初回設定手順】\n"
                    "1. https://console.cloud.google.com/ にアクセス\n"
                    "2. 新しいプロジェクトを作成\n"
                    "3. 「APIとサービス」→「ライブラリ」で Gmail API を有効化\n"
                    "4. 「APIとサービス」→「認証情報」→「OAuth 2.0 クライアントID」を作成\n"
                    "   アプリケーションの種類：デスクトップアプリ\n"
                    "5. JSONをダウンロードして credentials.json という名前で\n"
                    "   このツールと同じフォルダに保存\n"
                    "6. 再度「Gmail取得」ボタンを押すとブラウザが開くので認証してください"
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'w', encoding='utf-8') as f:
            f.write(creds.to_json())

    from googleapiclient.discovery import build
    return build('gmail', 'v1', credentials=creds)


def _decode_body(data):
    if not data:
        return ''
    padding = 4 - len(data) % 4
    if padding != 4:
        data += '=' * padding
    try:
        return base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
    except Exception:
        return ''


def _extract_text_body(payload):
    mime = payload.get('mimeType', '')

    if mime == 'text/plain':
        return _decode_body(payload.get('body', {}).get('data', ''))

    for part in payload.get('parts', []):
        if part.get('mimeType') == 'text/plain':
            text = _decode_body(part.get('body', {}).get('data', ''))
            if text:
                return text
        elif part.get('mimeType', '').startswith('multipart/'):
            text = _extract_text_body(part)
            if text:
                return text

    return ''


def parse_amazon_email(body):
    """
    Amazonの注文確定メールから情報を抽出する

    Returns:
        dict:
            product_name  : str  商品名
            selling_price : int  販売価格（円）
            amazon_fee    : int  Amazon手数料（円）
    """
    result = {'product_name': '', 'selling_price': 0, 'amazon_fee': 0}

    # 商品名
    for pat in [r'商品名[：:]\s*(.+)', r'タイトル[：:]\s*(.+)', r'アイテム[：:]\s*(.+)']:
        m = re.search(pat, body)
        if m:
            name = m.group(1).strip()
            if name and len(name) < 200:
                result['product_name'] = name
                break

    # 販売価格
    for pat in [
        r'販売価格[：:]\s*[￥¥]([\d,]+)',
        r'商品の小計[：:]\s*[￥¥]([\d,]+)',
        r'商品代金[：:]\s*[￥¥]([\d,]+)',
        r'注文金額[：:]\s*[￥¥]([\d,]+)',
    ]:
        m = re.search(pat, body)
        if m:
            result['selling_price'] = int(m.group(1).replace(',', ''))
            break

    # Amazon手数料  ← ユーザー確認済み形式: "Amazon手数料： (￥○○○ )"
    for pat in [
        r'Amazon手数料[：:]\s*\(\s*[￥¥]([\d,]+)\s*\)',
        r'Amazon手数料[：:]\s*-\s*[￥¥]([\d,]+)',
        r'Amazon手数料[：:]\s*[￥¥]([\d,]+)',
        r'販売手数料[：:]\s*\(\s*[￥¥]([\d,]+)\s*\)',
        r'販売手数料[：:]\s*[￥¥]([\d,]+)',
    ]:
        m = re.search(pat, body)
        if m:
            result['amazon_fee'] = int(m.group(1).replace(',', ''))
            break

    return result


def fetch_amazon_orders(max_results=5):
    """
    最近の Amazon 注文確定メールを取得して解析する

    Returns:
        list of dict:
            subject       : str  メール件名
            date          : str  受信日時
            product_name  : str
            selling_price : int
            amazon_fee    : int
    """
    service = get_gmail_service()

    query = 'from:seller-notification@amazon.co.jp subject:注文確定'
    resp = service.users().messages().list(
        userId='me', q=query, maxResults=max_results
    ).execute()

    messages = resp.get('messages', [])
    if not messages:
        return []

    orders = []
    for msg_info in messages:
        msg = service.users().messages().get(
            userId='me', id=msg_info['id'], format='full'
        ).execute()

        headers = {h['name']: h['value']
                   for h in msg['payload'].get('headers', [])}

        body = _extract_text_body(msg['payload'])
        parsed = parse_amazon_email(body)
        parsed['subject'] = headers.get('Subject', '(件名なし)')
        parsed['date'] = headers.get('Date', '')
        orders.append(parsed)

    return orders
