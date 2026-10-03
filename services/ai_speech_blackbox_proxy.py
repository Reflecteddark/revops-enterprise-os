"""
=============================================================================
🤖 REVOPS ENTERPRISE — PROPRIETARY AI BLACK BOX SPEECH PROXY
=============================================================================
Architecture: Secure SaaS Backend (Self-Hosted on Yandex Cloud / Selectel)

PURPOSE:
Protects Intellectual Property (IP):
1. Whisper Large V3 transcription & 152-FZ PII redaction pipeline.
2. Proprietary 13-criteria B2B sales scoring system prompt.
3. Next Step & Objection handling detectors.
4. Auto-injection of scores directly into client's Google Sheets raw_calls.

THE CLIENT NEVER HAS ACCESS TO SYSTEM PROMPTS OR AI ORCHESTRATION.
If the subscription is cancelled, this service rejects webhooks, and the client
spreadsheet immediately stops receiving speech analytics and AI coaching tips.
=============================================================================
"""

import os
import sys
import json
import logging
from typing import Dict, Any, List
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("AIBlackBoxProxy")

# 1. PROPRIETARY SYSTEM PROMPT (STRICT TRADE SECRET — NEVER STORED IN SHEETS)
PROPRIETARY_13_CRITERIA_PROMPT = """
Ты — RevOps Enterprise AI-Супервайзер звонков топ-уровня. 
Твоя задача — объективно оценить стенограмму телефонного B2B-диалога менеджера по 13 жестким стандартам продаж (0 или 1 по каждому).

СТАНДАРТЫ ОЦЕНКИ:
1. К1_NEXT_STEP: Назначена ли ТОЧНАЯ дата и время следующего шага? (0 если "я наберу на следующей неделе")
2. К2_INITIATIVE: Удерживал ли менеджер инициативу вопросами?
3. К3_DECISION_MAKER: Квалифицирован ли ЛПР и структура принятия решений?
4. К4_PAIN_DISCOVERY: Выявлена ли конкретная финансовая боль клиента?
5. К5_PRICE_DEFENSE: Защищена ли ценность до озвучивания цены?
6. К6_OBJECTION_HANDLING: Отработано ли возражение "Дорого" через окупаемость?
7. К7_CASE_STUDY: Приведен ли релевантный кейс с твердыми цифрами?
8. К8_CLOSING_ATTEMPT: Была ли попытка закрытия на целевое действие?
9. К9_GREETING_REGULATION: Корректное приветствие по регламенту компании?
10. К10_SPEECH_PURITY: Отсутствие слов-паразитов, пауз > 5 сек, уверенность?
11. К11_ACTIVE_LISTENING: Не перебивал ли клиента, перефразировал мысли?
12. К12_CALL_SUMMARY: Подведены ли итоги встречи в конце разговора?
13. К13_MARGIN_SECURITY: Не предложена ли необоснованная скидка без согласования?

ТРЕБОВАНИЯ К ВЫХОДУ (ТОЛЬКО ЧИСТЫЙ JSON):
{
  "scores": [0 или 1 для каждого из 13 критериев],
  "total_score": сумма баллов от 0 до 13,
  "next_step_flag": true/false,
  "critical_defect": "название критического срыва если total_score < 9",
  "quote": "дословная цитата из диалога где менеджер ошибся",
  "ai_recommendation": "конкретное действие для РОПа по спасению сделки",
  "urgency": "🔴 КРИТИЧЕСКИЙ / 🟡 ВНИМАНИЕ / 🟢 В НОРМЕ"
}
"""

class TenantLicenseManager:
    """Verifies tenant active license status before processing AI requests."""
    
    @staticmethod
    def verify_tenant(tenant_id: str) -> bool:
        # In production: check against PostgreSQL / Redis subscription DB
        # If tenant is expired or unpaid, immediately block processing
        logger.info(f"Checking license entitlement for tenant: '{tenant_id}'")
        # Simulating active tenant check
        return True

class SpeechAnalyticsPipeline:
    """Enterprise AI pipeline for processing calls and updating client sheets."""
    
    def __init__(self, tenant_id: str, spreadsheet_id: str):
        self.tenant_id = tenant_id
        self.spreadsheet_id = spreadsheet_id

    def process_incoming_call(self, call_payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Full lifecycle:
        Audio -> Whisper Large V3 -> 152-FZ Redaction -> DeepSeek V3 Scoring -> Sheets Injection
        """
        call_id = call_payload.get("call_id")
        deal_id = call_payload.get("deal_id")
        manager_id = call_payload.get("manager_id")
        duration_sec = call_payload.get("duration_sec", 0)

        # 1. License Gate
        if not TenantLicenseManager.verify_tenant(self.tenant_id):
            logger.error(f"Tenant '{self.tenant_id}' license expired! Rejecting call processing.")
            return {"status": "REJECTED_LICENSE_EXPIRED"}

        logger.info(f"Processing Call {call_id} for Deal {deal_id} (Duration: {duration_sec}s)...")
        
        # 2. Whisper Large V3 Transcribe + PII Masking (Simulated)
        # In production: whisper.transcribe(audio_file)
        simulated_transcript = (
            "Менеджер: Добрый день, ООО Альфа, меня зовут Мельников. "
            "Клиент: Здравствуйте, получили ваше КП на 850 000 рублей, это дороговато. "
            "Менеджер: Ну... мы можем сделать скидочку 10%. "
            "Клиент: Хорошо, подумаем, перезвоните как-нибудь. "
            "Менеджер: Договорились, наберу на следующей неделе, до свидания."
        )

        # 3. LLM Scoring using PROPRIETARY_13_CRITERIA_PROMPT (Black Box)
        # In production: deepseek_client.chat.completions.create(model="deepseek-chat", messages=[...])
        ai_result = {
            "call_id": call_id,
            "deal_id": deal_id,
            "manager_id": manager_id,
            "call_date": datetime.now().strftime("%Y-%m-%d"),
            "duration_sec": duration_sec,
            "call_type": "Исходящий",
            "criteria_scores": [0, 1, 1, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0], # 5 of 13
            "total_score": 5,
            "next_step_flag": False,
            "error_summary": "Слив сделки: скидка без защиты цены + нет фиксации даты Next Step",
            "ai_recommendation": (
                "🚨 Экстренный звонок РОПа ЛПР до 12:00: принести извинения за консультацию, "
                "аннулировать скидку в обмен на спец-гарантию 24 мес и зафиксировать встречу на вторник 11:00."
            ),
            "quote": "«Ну... мы можем сделать скидочку 10%. Договорились, наберу на следующей неделе»"
        }

        # 4. Formats row directly for client's raw_calls sheet
        raw_calls_row = [
            ai_result["call_id"],
            ai_result["deal_id"],
            ai_result["manager_id"],
            ai_result["call_date"],
            ai_result["duration_sec"],
            ai_result["call_type"],
            *ai_result["criteria_scores"],
            ai_result["total_score"],
            ai_result["next_step_flag"],
            ai_result["error_summary"],
            ai_result["ai_recommendation"],
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]

        logger.info(f"Call {call_id} evaluated with score {ai_result['total_score']}/13. Pushing to client sheet...")
        return {
            "status": "SUCCESS",
            "result": ai_result,
            "sheets_row_payload": raw_calls_row
        }

if __name__ == "__main__":
    print("Testing AI Speech Black Box Proxy...")
    pipeline = SpeechAnalyticsPipeline(
        tenant_id="CLIENT-TENANT-001",
        spreadsheet_id="1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A"
    )
    test_call = {
        "call_id": "C-901",
        "deal_id": "D-201",
        "manager_id": 101,
        "duration_sec": 240
    }
    output = pipeline.process_incoming_call(test_call)
    print("Pipeline Output Status:", output["status"])
    print("Evaluated Total Score:", output["result"]["total_score"], "/ 13")
    print("AI Recommendation:", output["result"]["ai_recommendation"][:70], "...")
    print("✅ AI Black Box Proxy verification complete!")
