import os
import sys
import json
import logging
import asyncio
import random
from typing import Any, Dict, Optional, List
from copy import deepcopy

import aiohttp
from aiohttp import web

# Force unbuffered output
sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

# -----------------------------------------------------------------------------
# Environment Variables
# -----------------------------------------------------------------------------
TOKEN = os.getenv("TOKEN", "")
CONFIRMATION_TOKEN = os.getenv("CONFIRMATION_TOKEN", "")
ADMIN_CHAT = os.getenv("ADMIN_CHAT", "")
PORT = int(os.getenv("PORT", "8080"))

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("vk_bot")

# Print startup info
print("=" * 50, flush=True)
print("VK BOT STARTING", flush=True)
print(f"TOKEN: {'SET' if TOKEN else 'NOT SET'}", flush=True)
print(f"CONFIRMATION_TOKEN: {'SET' if CONFIRMATION_TOKEN else 'NOT SET'}", flush=True)
print(f"ADMIN_CHAT: {ADMIN_CHAT}", flush=True)
print(f"PORT: {PORT}", flush=True)
print("=" * 50, flush=True)

# -----------------------------------------------------------------------------
# Load Data Files
# -----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def load_json_file(filename: str) -> Dict:
    """Load JSON file from disk."""
    filepath = os.path.join(BASE_DIR, filename)
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"Loaded {filename} successfully", flush=True)
        return data
    except Exception as e:
        print(f"Error loading {filename}: {e}", flush=True)
        return {}

# Load tests and texts
TESTS_DATA = load_json_file("tests.json")
TEXTS_DATA = load_json_file("texts.json")

# -----------------------------------------------------------------------------
# User Sessions Storage (in-memory)
# Format: {user_id: {"variant": N, "question": N, "score": N, "shuffled_answers": [...]}}
# -----------------------------------------------------------------------------
USER_SESSIONS: Dict[int, Dict] = {}

# -----------------------------------------------------------------------------
# VK API Helper
# -----------------------------------------------------------------------------
class VKAPI:
    API_URL = "https://api.vk.com/method/"
    API_VERSION = "5.199"
    
    def __init__(self, token: str):
        self.token = token
        self.session: Optional[aiohttp.ClientSession] = None
    
    async def init(self):
        if self.session is None:
            self.session = aiohttp.ClientSession()
    
    async def close(self):
        if self.session:
            await self.session.close()
            self.session = None
    
    async def call(self, method: str, params: Dict[str, Any]) -> Dict[str, Any]:
        if not self.session:
            await self.init()
        
        params["access_token"] = self.token
        params["v"] = self.API_VERSION
        url = f"{self.API_URL}{method}"
        
        try:
            async with self.session.post(url, data=params) as resp:
                result = await resp.json()
                if "error" in result:
                    logger.error(f"VK API error: {result['error']}")
                return result
        except Exception as e:
            logger.exception(f"VK API call error: {e}")
            return {"error": str(e)}
    
    async def send_message(
        self, 
        user_id: int, 
        message: str, 
        keyboard: Optional[Dict] = None,
        peer_id: Optional[int] = None
    ) -> Dict[str, Any]:
        params = {
            "message": message,
            "random_id": random.randint(0, 2**31 - 1)
        }
        
        if peer_id:
            params["peer_id"] = peer_id
        else:
            params["user_id"] = user_id
        
        if keyboard:
            params["keyboard"] = json.dumps(keyboard, ensure_ascii=False)
        
        return await self.call("messages.send", params)
    
    async def get_user_info(self, user_id: int) -> Dict[str, Any]:
        """Get user info by ID."""
        params = {"user_ids": user_id}
        return await self.call("users.get", params)


# -----------------------------------------------------------------------------
# Keyboard Builders
# -----------------------------------------------------------------------------
def create_main_menu_keyboard() -> Dict:
    """Main keyboard with only Menu button. Always visible."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {"type": "text", "label": "Меню"},
                    "color": "primary"
                }
            ]
        ]
    }

def create_menu_keyboard() -> Dict:
    """Menu with Testing button."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {"type": "text", "label": "Тестирование"},
                    "color": "positive"
                }
            ]
        ]
    }

def create_answer_keyboard(shuffled_answers: List[Dict]) -> Dict:
    """Keyboard with answer buttons 1, 2, 3."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {"type": "text", "label": "1"},
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": "2"},
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": "3"},
                    "color": "primary"
                }
            ]
        ]
    }

def create_retry_keyboard() -> Dict:
    """Keyboard with Menu and Retry buttons for failed test."""
    restart_text = TEXTS_DATA.get("restart_button", "Пройти заново")
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {"type": "text", "label": "Меню"},
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": restart_text},
                    "color": "positive"
                }
            ]
        ]
    }


# -----------------------------------------------------------------------------
# Test Logic
# -----------------------------------------------------------------------------
def shuffle_answers(answers: List[Dict]) -> List[Dict]:
    """Shuffle answers and return new list with button numbers."""
    shuffled = deepcopy(answers)
    random.shuffle(shuffled)
    return shuffled

def get_random_variant() -> int:
    """Get random variant number."""
    variants = TESTS_DATA.get("variants", [])
    if variants:
        return random.randint(0, len(variants) - 1)
    return 0

def get_question(variant_idx: int, question_idx: int) -> Optional[Dict]:
    """Get question from variant."""
    variants = TESTS_DATA.get("variants", [])
    if variant_idx < len(variants):
        questions = variants[variant_idx].get("questions", [])
        if question_idx < len(questions):
            return questions[question_idx]
    return None

def format_question_message(question: Dict, question_num: int, total: int) -> str:
    """Format question text for sending."""
    return f"Вопрос {question_num}/{total}:\n\n{question['question']}"

def get_correct_answer_text(question: Dict) -> str:
    """Get text of correct answer."""
    for answer in question.get("answers", []):
        if answer.get("is_correct", False):
            return answer.get("text", "")
    return ""


# -----------------------------------------------------------------------------
# Web Server
# -----------------------------------------------------------------------------
class WebServer:
    def __init__(self):
        self.app = web.Application()
        self.vk_api = VKAPI(TOKEN) if TOKEN else None
        self._setup_routes()
        print("WebServer initialized", flush=True)

    def _setup_routes(self) -> None:
        self.app.router.add_post("/", self.vk_webhook)
        self.app.router.add_post("/webhook", self.vk_webhook)
        self.app.router.add_get("/", self.health)
        self.app.router.add_get("/health", self.health)
        print("Routes setup complete", flush=True)

    async def health(self, request: web.Request) -> web.Response:
        print("Health check requested", flush=True)
        return web.json_response({
            "status": "ok",
            "token_configured": bool(TOKEN),
            "confirmation_token_configured": bool(CONFIRMATION_TOKEN),
            "admin_chat": ADMIN_CHAT,
            "tests_loaded": bool(TESTS_DATA),
            "texts_loaded": bool(TEXTS_DATA)
        })

    async def vk_webhook(self, request: web.Request) -> web.Response:
        try:
            body = await request.text()
            print(f"POST received: {body[:500]}", flush=True)
            
            try:
                data = json.loads(body)
            except json.JSONDecodeError as e:
                print(f"JSON parse error: {e}", flush=True)
                return web.Response(text="invalid json", status=400)
            
            event_type = data.get("type", "")
            group_id = data.get("group_id", 0)
            
            print(f"Event: {event_type}, group: {group_id}", flush=True)
            
            if event_type == "confirmation":
                print(f"CONFIRMATION REQUEST - returning: {CONFIRMATION_TOKEN}", flush=True)
                return web.Response(text=CONFIRMATION_TOKEN)
            
            if event_type == "message_new":
                print("Handling message_new", flush=True)
                await self._handle_message_new(data)
                return web.Response(text="ok")
            
            print(f"Unknown event: {event_type}", flush=True)
            return web.Response(text="ok")
            
        except Exception as e:
            print(f"ERROR: {e}", flush=True)
            return web.Response(text="ok")

    async def _handle_message_new(self, data: Dict) -> None:
        try:
            if not self.vk_api:
                print("No VK API", flush=True)
                return
            
            await self.vk_api.init()
            
            obj = data.get("object", {})
            message = obj.get("message", obj)
            
            user_id = message.get("from_id", message.get("user_id", 0))
            peer_id = message.get("peer_id", user_id)
            text = message.get("text", "").strip()
            
            print(f"User {user_id}: {text}", flush=True)
            
            # Handle "Start" button
            if text.lower() in ["начать", "start", "/start"]:
                # Clear any existing session
                if user_id in USER_SESSIONS:
                    del USER_SESSIONS[user_id]
                
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Добро пожаловать!\n\nБот не реагирует на текстовые сообщения. Используйте кнопки клавиатуры.",
                    peer_id=peer_id,
                    keyboard=create_main_menu_keyboard()
                )
                return
            
            # Handle "Меню" button
            if text.lower() == "меню":
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Выберите действие:",
                    peer_id=peer_id,
                    keyboard=create_menu_keyboard()
                )
                return
            
            # Handle "Тестирование" button
            if text.lower() == "тестирование":
                await self._start_test(user_id, peer_id)
                return
            
            # Handle "Пройти заново" button
            restart_text = TEXTS_DATA.get("restart_button", "Пройти заново")
            if text.lower() == restart_text.lower():
                await self._start_test(user_id, peer_id)
                return
            
            # Handle answer buttons (1, 2, 3)
            if text in ["1", "2", "3"]:
                await self._handle_answer(user_id, peer_id, int(text))
                return
            
            # Ignore all other text
            print(f"Ignoring text message from user {user_id}: {text}", flush=True)
            
        except Exception as e:
            print(f"Message handling error: {e}", flush=True)

    async def _start_test(self, user_id: int, peer_id: int) -> None:
        """Start new test for user."""
        # Send intro text
        intro_text = TEXTS_DATA.get("test_intro", "Начинаем тестирование!")
        await self.vk_api.send_message(
            user_id=user_id,
            message=intro_text,
            peer_id=peer_id
        )
        
        # Initialize session
        variant_idx = get_random_variant()
        USER_SESSIONS[user_id] = {
            "variant": variant_idx,
            "question": 0,
            "score": 0,
            "shuffled_answers": None
        }
        
        # Send first question
        await self._send_question(user_id, peer_id)

    async def _send_question(self, user_id: int, peer_id: int) -> None:
        """Send current question to user."""
        session = USER_SESSIONS.get(user_id)
        if not session:
            print(f"No session for user {user_id}", flush=True)
            return
        
        variant_idx = session["variant"]
        question_idx = session["question"]
        
        question = get_question(variant_idx, question_idx)
        if not question:
            print(f"No question found: variant={variant_idx}, question={question_idx}", flush=True)
            return
        
        # Shuffle answers
        shuffled = shuffle_answers(question.get("answers", []))
        session["shuffled_answers"] = shuffled
        
        # Get total questions
        total = len(TESTS_DATA.get("variants", [{}])[variant_idx].get("questions", []))
        
        # Format and send question
        question_text = format_question_message(question, question_idx + 1, total)
        
        # Add answer options
        answer_text = "\n\n1️⃣ " + shuffled[0]["text"]
        answer_text += "\n2️⃣ " + shuffled[1]["text"]
        answer_text += "\n3️⃣ " + shuffled[2]["text"]
        
        await self.vk_api.send_message(
            user_id=user_id,
            message=question_text + answer_text,
            peer_id=peer_id,
            keyboard=create_answer_keyboard(shuffled)
        )

    async def _handle_answer(self, user_id: int, peer_id: int, answer_num: int) -> None:
        """Handle user's answer."""
        session = USER_SESSIONS.get(user_id)
        if not session:
            print(f"No session for user {user_id}", flush=True)
            return
        
        shuffled = session.get("shuffled_answers", [])
        if answer_num < 1 or answer_num > len(shuffled):
            return
        
        # Check answer (answer_num is 1-indexed)
        selected_answer = shuffled[answer_num - 1]
        is_correct = selected_answer.get("is_correct", False)
        
        variant_idx = session["variant"]
        question_idx = session["question"]
        question = get_question(variant_idx, question_idx)
        
        if is_correct:
            session["score"] += 1
            correct_text = TEXTS_DATA.get("correct_answer", "Верно!")
            await self.vk_api.send_message(
                user_id=user_id,
                message=correct_text,
                peer_id=peer_id
            )
        else:
            wrong_text = TEXTS_DATA.get("wrong_answer", "Неверно!")
            correct_answer = get_correct_answer_text(question)
            message = f"{wrong_text}\n\nПравильный ответ: {correct_answer}"
            await self.vk_api.send_message(
                user_id=user_id,
                message=message,
                peer_id=peer_id
            )
        
        # Move to next question
        session["question"] += 1
        
        # Check if test is complete
        total = len(TESTS_DATA.get("variants", [{}])[variant_idx].get("questions", []))
        
        if session["question"] >= total:
            await self._finish_test(user_id, peer_id)
        else:
            await self._send_question(user_id, peer_id)

    async def _finish_test(self, user_id: int, peer_id: int) -> None:
        """Finish test and show results."""
        session = USER_SESSIONS.get(user_id)
        if not session:
            return
        
        score = session["score"]
        total = len(TESTS_DATA.get("variants", [{}])[session["variant"]].get("questions", []))
        passing_score = TESTS_DATA.get("test_info", {}).get("passing_score", 18)
        passed = score >= passing_score
        
        # Get user name for admin notification
        user_name = f"ID{user_id}"
        try:
            user_info = await self.vk_api.get_user_info(user_id)
            if "response" in user_info and user_info["response"]:
                first_name = user_info["response"][0].get("first_name", "")
                last_name = user_info["response"][0].get("last_name", "")
                user_name = f"{first_name} {last_name} (ID: {user_id})"
        except Exception as e:
            print(f"Error getting user info: {e}", flush=True)
        
        # Send notification to admin chat
        if ADMIN_CHAT:
            try:
                status = "СДАЛ" if passed else "НЕ СДАЛ"
                admin_message = f"Пользователь {user_name} {status} тест!\nРезультат: {score}/{total}"
                await self.vk_api.send_message(
                    user_id=0,
                    message=admin_message,
                    peer_id=int(ADMIN_CHAT)
                )
                print(f"Admin notification sent to {ADMIN_CHAT}", flush=True)
            except Exception as e:
                print(f"Failed to send admin notification: {e}", flush=True)
        
        # Send result to user
        if passed:
            passed_text = TEXTS_DATA.get("test_passed", "🎉 Поздравляем! Вы сдали тест!")
            message = f"{passed_text}\n\nВаш результат: {score}/{total}"
            await self.vk_api.send_message(
                user_id=user_id,
                message=message,
                peer_id=peer_id,
                keyboard=create_main_menu_keyboard()
            )
        else:
            failed_text = TEXTS_DATA.get("test_failed", "К сожалению, вы не набрали нужное количество баллов.")
            message = f"{failed_text}\n\nВаш результат: {score}/{total}\nНеобходимо: {passing_score}/{total}"
            await self.vk_api.send_message(
                user_id=user_id,
                message=message,
                peer_id=peer_id,
                keyboard=create_retry_keyboard()
            )
        
        # Clear session
        del USER_SESSIONS[user_id]


# -----------------------------------------------------------------------------
# Entrypoint
# -----------------------------------------------------------------------------
def main_sync():
    print("Creating web server...", flush=True)
    server = WebServer()
    
    print(f"Starting on port {PORT}...", flush=True)
    web.run_app(server.app, host="0.0.0.0", port=PORT, print=lambda x: print(x, flush=True))


if __name__ == "__main__":
    main_sync()
