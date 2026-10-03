# -*- coding: utf-8 -*-
"""
Upgraded MAX Bot with automated file downloads to Desktop and 3-call progress tracker.
"""

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

# Ensure log and download directory exists
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "max_bot.log")

# Desktop incoming folder for Dmitry
DESKTOP_INCOMING = r"C:\Users\strel\Desktop\RevOps Platform\Входящие_Звонки_MAX"
LOCAL_INCOMING = os.path.join(BASE_DIR, "incoming_audios", "max")
os.makedirs(DESKTOP_INCOMING, exist_ok=True)
os.makedirs(LOCAL_INCOMING, exist_ok=True)

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
    "👋 Здравствуйте! Я ИИ-супервайзер сервиса RevOps OS (ai-rop.ru).\n\n"
    "🎯 БЕСПЛАТНЫЙ АУДИТ 3 ЗВОНКОВ (0 ₽):\n"
    "Прикрепите сюда 3 аудиозаписи звонков ваших менеджеров (файлами mp3/wav/m4a или перешлите аудио/войс из чата).\n\n"
    "⏱️ За 20 минут мы расшифруем диалоги через Whisper Pro, оцифруем скрытые сливы выручки и пришлем вам PDF-карту потерь!\n\n"
    "👤 Основатель сервиса Дмитрий Федотов: @dm1918 в Telegram."
)

KEYBOARD_BUTTONS = [
    [
        {"type": "link", "text": "🌐 Сайт ai-rop.ru", "url": "https://ai-rop.ru"},
        {"type": "link", "text": "📄 Образец PDF-отчета", "url": "https://ai-rop.ru/sample-audit-report.html"}
    ],
    [
        {"type": "link", "text": "💬 Связь с Дмитрием в TG", "url": "https://t.me/dm1918"},
        {"type": "link", "text": "🤖 Telegram-бот", "url": "https://t.me/RevOps_Super_Audit_Bot"}
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

def download_attachment(att, user_name, user_id, index):
    """Downloads audio attachment and saves to Desktop and local backup."""
    payload = att.get("payload", {})
    file_url = payload.get("url") or payload.get("link") or att.get("url")
    file_name = payload.get("name") or f"call_{index}.mp3"
    
    # Clean folder name
    safe_name = "".join(c for c in f"{user_name}_{user_id}" if c.isalnum() or c in " _-").strip()
    target_desktop_folder = os.path.join(DESKTOP_INCOMING, safe_name)
    target_local_folder = os.path.join(LOCAL_INCOMING, safe_name)
    os.makedirs(target_desktop_folder, exist_ok=True)
    os.makedirs(target_local_folder, exist_ok=True)
    
    dest_file = os.path.join(target_desktop_folder, f"{index}_{file_name}")
    dest_local = os.path.join(target_local_folder, f"{index}_{file_name}")
    
    if file_url:
        try:
            req = urllib.request.Request(file_url, headers={"User-Agent": "AI-ROP-Bot/1.0", "Authorization": BOT_TOKEN})
            with urllib.request.urlopen(req, context=SSL_CTX, timeout=30) as resp:
                data = resp.read()
                with open(dest_file, "wb") as f:
                    f.write(data)
                with open(dest_local, "wb") as f:
                    f.write(data)
            logger.info(f"✓ Saved audio ({len(data)} bytes) to: {dest_file}")
            return dest_file
        except Exception as e:
            logger.error(f"Failed to download attachment {file_url}: {e}")
    return None

def send_notification_email(user_info, text_content, attachments_count=0, local_files=None):
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
            "_subject": f"[MAX] 🎙️ {attachments_count} звонков получено от {user_name}!",
            "Источник": "Мессенджер MAX (@se14526668_bot)",
            "Клиент": contact_str,
            "Сообщение": text_content or "(аудиозаписи звонков)",
            "Количество файлов": str(attachments_count),
            "Сохранено в папку": DESKTOP_INCOMING,
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
            logger.info(f"FormSubmit email notification sent for MAX user {user_id}")
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

    # Start Health Check HTTP server if PORT set
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
    user_audio_counts = {}

    logger.info(f"Incoming audios will be saved to: {DESKTOP_INCOMING}")
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
                
                if att_count > 0:
                    current_count = user_audio_counts.get(user_id, 0) + att_count
                    user_audio_counts[user_id] = current_count
                    
                    # Download files
                    saved_files = []
                    for idx, att in enumerate(attachments, 1):
                        sf = download_attachment(att, user_name, user_id, current_count - att_count + idx)
                        if sf: saved_files.append(sf)
                    
                    if current_count == 1:
                        reply_text = (
                            "✅ <b>Аудиозапись 1 из 3 успешно принята!</b>\n\n"
                            "Присылайте еще 2 записи звонков ваших менеджеров для полного сравнительного анализа сливов."
                        )
                    elif current_count == 2:
                        reply_text = (
                            "✅ <b>Аудиозапись 2 из 3 успешно принята!</b>\n\n"
                            "Отлично! Прикрепите последнюю (3-ю) запись звонка, и ИИ запустит расшифровку."
                        )
                    else:
                        reply_text = (
                            f"🚀 <b>Все {current_count} звонка успешно получены!</b>\n\n"
                            "Нейросеть Whisper Pro и модель RevOps AI приступили к анализу по 13 критериям качества.\n\n"
                            "Основатель сервиса Дмитрий Федотов уже подключился к задаче и пришлет вам оцифрованную PDF-карту сливов в этот чат в течение 20 минут!"
                        )
                        user_audio_counts[user_id] = 0 # reset cycle
                        
                    user_dict = {"id": user_id, "name": user_name, "username": user_handle}
                    send_notification_email(user_dict, text or "(файловые вложения звонков)", attachments_count=current_count, local_files=saved_files)
                else:
                    reply_text = GREETING_TEXT
                    user_dict = {"id": user_id, "name": user_name, "username": user_handle}
                    send_notification_email(user_dict, text, attachments_count=0)
                
                # Send response in MAX chat
                send_bot_message(chat_id=chat_id, user_id=user_id, text=reply_text, with_keyboard=True)
                
        except KeyboardInterrupt:
            logger.info("Bot service stopped by user.")
            break
        except Exception as e:
            logger.error(f"Error in polling loop: {e}", exc_info=True)
            time.sleep(5)

if __name__ == "__main__":
    run_bot()
