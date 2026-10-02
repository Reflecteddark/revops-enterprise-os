import os
import sys
import time
import json
import ssl
import logging
import urllib.request
import urllib.parse
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime

# Windows console encoding fix
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Ensure log directory exists
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "max_bot.log")

# Setup logging
logger = logging.getLogger("MAX_BOT")
logger.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

BOT_TOKEN = os.environ.get("MAX_BOT_TOKEN", "f9LHodD0cOJTnBPV68BrzK-RA3ZhV6cmttvWtFpfM_fJH8mCqoAKSL96wERe7SnOq4v_2fsgeesQI4-I8Q6c")
BASE_URL = "https://platform-api2.max.ru"
SSL_CTX = ssl._create_unverified_context()

GREETING_TEXT = (
    "Здравствуйте! Я ИИ-супервайзер платформы AI-ROP (ai-rop.ru).\n\n"
    "1. Экспресс-аудит (0 руб): пришлите сюда 3 аудиозаписи звонков — мы бесплатно разберем их по 13 бизнес-критериям за 20 минут.\n\n"
    "2. Калькулятор потерь: откройте мини-приложение кнопкой ниже, чтобы посчитать упущенную выручку вашего отдела.\n\n"
    "3. Связь с основателем: @dm1918 в Telegram или info@ai-rop.ru."
)

ATTACHMENT_RECEIVED_TEXT = (
    "Аудиозапись звонка получена! Ваша запись поставлена в очередь на экспресс-аудит Whisper Pro + RevOps OS.\n\n"
    "Если у вас есть еще записи (до 3 звонков), присылайте сюда. "
    "Наш специалист или основатель свяжется с вами для передачи итогового отчета."
)

KEYBOARD_BUTTONS = [
    [
        {"type": "link", "text": "Открыть сайт ai-rop.ru", "url": "https://ai-rop.ru"},
        {"type": "link", "text": "Основатель в TG (@dm1918)", "url": "https://t.me/dm1918"}
    ]
]

class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(b'{"status": "ok", "service": "AI-ROP MAX Bot", "bot": "@se14526668_bot"}')

    def log_message(self, format, *args):
        return

def start_health_server(port):
    server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
    logger.info(f"Health check HTTP server listening on 0.0.0.0:{port}")
    server.serve_forever()

def api_request(endpoint, method="GET", data=None, params=None):
    url = f"{BASE_URL}{endpoint}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    
    headers = {
        "Authorization": BOT_TOKEN,
        "User-Agent": "AI-ROP-Bot/1.0"
    }
    
    body = None
    if data is not None:
        headers["Content-Type"] = "application/json"
        body = json.dumps(data).encode("utf-8")
        
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, context=SSL_CTX, timeout=30) as resp:
            content = resp.read().decode("utf-8")
            return json.loads(content) if content else {}
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        logger.error(f"HTTP {e.code} on {url}: {err_body}")
        return None
    except Exception as e:
        logger.error(f"Request error on {url}: {e}")
        return None

def send_notification_email(user_info, text_content, attachments_count=0):
    """Sends lead notification to info@ai-rop.ru via FormSubmit"""
    try:
        now_str = datetime.now().strftime("%d.%m.%Y, %H:%M:%S")
        user_name = user_info.get("name") or "Не указано"
        user_id = str(user_info.get("id") or "Unknown")
        user_handle = user_info.get("username") or ""
        
        contact_str = f"{user_name} (ID: {user_id})"
        if user_handle:
            contact_str += f", @{user_handle}"
            
        payload = {
            "_subject": f"[MAX] Новое обращение от {user_name}",
            "Источник": "Мессенджер MAX (@se14526668_bot)",
            "Пользователь": contact_str,
            "Сообщение": text_content or "(содержимое без текста)",
            "Файлов прикреплено": str(attachments_count),
            "Дата": now_str,
            "_template": "table",
            "_captcha": "false"
        }
        
        req = urllib.request.Request(
            "https://formsubmit.co/ajax/info@ai-rop.ru",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Origin": "https://ai-rop.ru",
                "Referer": "https://ai-rop.ru/"
            }
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            logger.info(f"FormSubmit notification sent for MAX user {user_id}")
    except Exception as e:
        logger.warning(f"Failed to dispatch email notification: {e}")

def send_bot_message(chat_id=None, user_id=None, text="", with_keyboard=True):
    params = {}
    if chat_id:
        params["chat_id"] = chat_id
    elif user_id:
        params["user_id"] = user_id
    else:
        return False

    payload = {"text": text}
    if with_keyboard:
        payload["attachments"] = [
            {
                "type": "inline_keyboard",
                "payload": {"buttons": KEYBOARD_BUTTONS}
            }
        ]

    res = api_request("/messages", method="POST", data=payload, params=params)
    return res is not None

def run_bot():
    logger.info("Initializing AI-ROP MAX Bot service...")
    
    # 1. Check Bot Identity
    me = api_request("/me")
    if not me:
        logger.error("Failed to authenticate bot token. Please check MAX_BOT_TOKEN.")
        return
    
    logger.info(f"Connected to MAX as @{me.get('username')} (ID: {me.get('user_id')}, Name: {me.get('first_name')})")

    # Start Health Check HTTP server for Render/Railway/Cloud Run if PORT is set
    port_env = os.environ.get("PORT")
    if port_env:
        try:
            port = int(port_env)
            t = threading.Thread(target=start_health_server, args=(port,), daemon=True)
            t.start()
        except Exception as e:
            logger.error(f"Failed to start health check server on port {port_env}: {e}")

    marker = None
    processed_msg_ids = set()

    logger.info("Entering polling loop (long-polling timeout=20s)...")
    while True:
        try:
            params = {"timeout": 20}
            if marker:
                params["marker"] = marker
                
            updates_data = api_request("/updates", params=params)
            if not updates_data:
                time.sleep(2)
                continue
                
            updates = updates_data.get("updates", [])
            marker = updates_data.get("marker", marker)
            
            for upd in updates:
                upd_type = upd.get("update_type") or upd.get("type")
                
                # Check for message update
                msg = upd.get("message")
                if not msg and upd_type in ["message_created", "message"]:
                    msg = upd
                    
                if not msg:
                    continue

                msg_body = msg.get("body", {})
                msg_id = msg_body.get("mid") or msg.get("mid") or msg.get("id")
                if msg_id and msg_id in processed_msg_ids:
                    continue
                if msg_id:
                    processed_msg_ids.add(msg_id)
                    # Limit size of processed cache
                    if len(processed_msg_ids) > 10000:
                        processed_msg_ids.clear()
                    
                sender = msg.get("sender", {})
                chat_id = msg.get("recipient", {}).get("chat_id") or msg.get("chat_id")
                user_id = sender.get("user_id") or sender.get("id")
                user_name = sender.get("name") or sender.get("first_name") or f"User_{user_id}"
                user_handle = sender.get("username") or ""
                
                # Ignore messages sent by the bot itself
                if user_id == me.get("user_id"):
                    continue

                text = msg_body.get("text") or msg.get("text") or ""
                attachments = msg_body.get("attachments") or msg.get("attachments") or []
                att_count = len(attachments)
                
                logger.info(f"Incoming message from {user_name} (chat={chat_id}, user={user_id}, text='{text}', attachments={att_count})")
                
                # Choose reply message
                reply_text = ATTACHMENT_RECEIVED_TEXT if att_count > 0 else GREETING_TEXT
                
                # Send response in MAX chat
                send_bot_message(chat_id=chat_id, user_id=user_id, text=reply_text, with_keyboard=True)
                
                # Forward notification to email
                user_dict = {"id": user_id, "name": user_name, "username": user_handle}
                send_notification_email(user_dict, text or "(файловое вложение)", attachments_count=att_count)
                
        except KeyboardInterrupt:
            logger.info("Bot service stopped by user.")
            break
        except Exception as e:
            logger.error(f"Error in polling loop: {e}", exc_info=True)
            time.sleep(5)

if __name__ == "__main__":
    run_bot()
