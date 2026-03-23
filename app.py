import os
import sys
import json
import logging
import asyncio
import random
import io
from typing import Any, Dict, Optional, List
from copy import deepcopy
from datetime import datetime

import aiohttp
from aiohttp import web
import asyncpg
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Force unbuffered output
sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

# -----------------------------------------------------------------------------
# Environment Variables
# -----------------------------------------------------------------------------
TOKEN = os.getenv("TOKEN", "")
CONFIRMATION_TOKEN = os.getenv("CONFIRMATION_TOKEN", "")
CHECK_TEST_STR = os.getenv("CHECK_TEST", "")  # Admin IDs for test notifications
USER_ADMIN = os.getenv("USER_ADMIN", "")  # Super admin ID for database export
PORT = int(os.getenv("PORT", "8080"))

# Database configuration
DATABASE_URL = os.getenv("BD", "")  # Full database URL
DB_USER = os.getenv("BD_USER_NAME", "")
DB_PASSWORD = os.getenv("BD_USER_PASSWORD", "")

# Parse admin IDs from comma-separated string
CHECK_TEST_IDS = [int(id.strip()) for id in CHECK_TEST_STR.split(",") if id.strip()]

# Parse super admin ID
USER_ADMIN_ID = int(USER_ADMIN) if USER_ADMIN else 0

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
print(f"CHECK_TEST_IDS: {CHECK_TEST_IDS}", flush=True)
print(f"USER_ADMIN_ID: {USER_ADMIN_ID}", flush=True)
print(f"DATABASE_URL: {'SET' if DATABASE_URL else 'NOT SET'}", flush=True)
print(f"DB_USER: {DB_USER}", flush=True)
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
# Database Helper
# -----------------------------------------------------------------------------
class Database:
    """PostgreSQL database helper for user data."""
    
    def __init__(self):
        self.pool: Optional[asyncpg.Pool] = None
    
    async def init(self):
        """Initialize database connection pool and create table if not exists."""
        try:
            # Build connection string
            if DATABASE_URL:
                conn_string = DATABASE_URL
            elif DB_USER and DB_PASSWORD:
                conn_string = f"postgresql://{DB_USER}:{DB_PASSWORD}@localhost/bvc_bot"
            else:
                print("No database configuration found!", flush=True)
                return False
            
            print(f"Connecting to database...", flush=True)
            self.pool = await asyncpg.create_pool(conn_string, min_size=2, max_size=10)
            
            # Create table if not exists
            async with self.pool.acquire() as conn:
                await conn.execute('''
                    CREATE TABLE IF NOT EXISTS users (
                        user_id BIGINT PRIMARY KEY,
                        user_name TEXT,
                        form BOOLEAN DEFAULT FALSE,
                        form_answer TEXT DEFAULT '',
                        test_book_1 BOOLEAN DEFAULT FALSE,
                        test_book_2 BOOLEAN DEFAULT FALSE,
                        test_book_3 BOOLEAN DEFAULT FALSE,
                        test_book_4 BOOLEAN DEFAULT FALSE,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                print("Database table verified/created", flush=True)
            
            print("Database connection established", flush=True)
            return True
            
        except Exception as e:
            print(f"Database initialization error: {e}", flush=True)
            return False
    
    async def close(self):
        """Close database connection pool."""
        if self.pool:
            await self.pool.close()
            print("Database connection closed", flush=True)
    
    async def get_user(self, user_id: int) -> Optional[Dict]:
        """Get user by ID."""
        if not self.pool:
            return None
        
        try:
            async with self.pool.acquire() as conn:
                row = await conn.fetchrow(
                    "SELECT * FROM users WHERE user_id = $1", user_id
                )
                if row:
                    return dict(row)
                return None
        except Exception as e:
            print(f"Error getting user {user_id}: {e}", flush=True)
            return None
    
    async def get_all_users(self) -> List[Dict]:
        """Get all users from database."""
        if not self.pool:
            return []
        
        try:
            async with self.pool.acquire() as conn:
                rows = await conn.fetch("SELECT * FROM users ORDER BY created_at DESC")
                return [dict(row) for row in rows]
        except Exception as e:
            print(f"Error getting all users: {e}", flush=True)
            return []
    
    async def create_user(self, user_id: int, user_name: str) -> bool:
        """Create new user."""
        if not self.pool:
            return False
        
        try:
            async with self.pool.acquire() as conn:
                await conn.execute('''
                    INSERT INTO users (user_id, user_name, form, form_answer, 
                                       test_book_1, test_book_2, test_book_3, test_book_4)
                    VALUES ($1, $2, FALSE, '', FALSE, FALSE, FALSE, FALSE)
                    ON CONFLICT (user_id) DO UPDATE SET user_name = $2, updated_at = CURRENT_TIMESTAMP
                ''', user_id, user_name)
                print(f"User {user_id} created/updated in database", flush=True)
                return True
        except Exception as e:
            print(f"Error creating user {user_id}: {e}", flush=True)
            return False
    
    async def update_user_field(self, user_id: int, field: str, value: Any) -> bool:
        """Update a specific field for user."""
        if not self.pool:
            return False
        
        allowed_fields = ['form', 'form_answer', 'test_book_1', 'test_book_2', 'test_book_3', 'test_book_4', 'user_name']
        if field not in allowed_fields:
            return False
        
        try:
            async with self.pool.acquire() as conn:
                await conn.execute(
                    f"UPDATE users SET {field} = $1, updated_at = CURRENT_TIMESTAMP WHERE user_id = $2",
                    value, user_id
                )
                return True
        except Exception as e:
            print(f"Error updating user {user_id} field {field}: {e}", flush=True)
            return False

# Global database instance
db = Database()

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
        peer_id: Optional[int] = None,
        attachment: Optional[str] = None
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
        
        if attachment:
            params["attachment"] = attachment
        
        return await self.call("messages.send", params)
    
    async def get_user_info(self, user_id: int) -> Dict[str, Any]:
        """Get user info by ID."""
        params = {"user_ids": user_id}
        return await self.call("users.get", params)
    
    async def get_upload_server(self, peer_id: int) -> Optional[str]:
        """Get upload server URL for documents."""
        result = await self.call("docs.getMessagesUploadServer", {"peer_id": peer_id, "type": "doc"})
        if "response" in result and "upload_url" in result["response"]:
            return result["response"]["upload_url"]
        print(f"Error getting upload server: {result}", flush=True)
        return None
    
    async def upload_document(self, upload_url: str, file_data: bytes, filename: str) -> Optional[Dict]:
        """Upload document to VK."""
        if not self.session:
            await self.init()
        
        try:
            form = aiohttp.FormData()
            form.add_field('file', file_data, filename=filename, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            
            async with self.session.post(upload_url, data=form) as resp:
                result = await resp.json()
                return result
        except Exception as e:
            print(f"Error uploading document: {e}", flush=True)
            return None
    
    async def save_document(self, file: str, title: str) -> Optional[Dict]:
        """Save uploaded document."""
        result = await self.call("docs.save", {"file": file, "title": title})
        return result
    
    async def send_document(self, peer_id: int, file_data: bytes, filename: str, message: str = "") -> bool:
        """Upload and send document to user."""
        try:
            # Get upload server
            upload_url = await self.get_upload_server(peer_id)
            if not upload_url:
                return False
            
            # Upload file
            upload_result = await self.upload_document(upload_url, file_data, filename)
            if not upload_result or "file" not in upload_result:
                print(f"Upload failed: {upload_result}", flush=True)
                return False
            
            # Save document
            save_result = await self.save_document(upload_result["file"], filename)
            if not save_result or "response" not in save_result:
                print(f"Save failed: {save_result}", flush=True)
                return False
            
            # Get document attachment string
            doc = save_result["response"].get("doc", save_result["response"].get("docs", [{}])[0] if "docs" in save_result["response"] else {})
            if not doc:
                print(f"No doc in response: {save_result}", flush=True)
                return False
            
            owner_id = doc.get("owner_id", doc.get("doc", {}).get("owner_id", ""))
            doc_id = doc.get("id", doc.get("doc", {}).get("id", ""))
            attachment = f"doc{owner_id}_{doc_id}"
            
            # Send message with attachment
            await self.send_message(
                user_id=0,
                message=message,
                peer_id=peer_id,
                attachment=attachment
            )
            
            return True
            
        except Exception as e:
            print(f"Error sending document: {e}", flush=True)
            return False


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

def create_menu_keyboard_with_form(is_admin: bool = False) -> Dict:
    """Menu with Анкета button (for users who haven't completed form)."""
    buttons = [
        [
            {
                "action": {"type": "text", "label": "Меню"},
                "color": "primary"
            },
            {
                "action": {"type": "text", "label": "Анкета"},
                "color": "positive"
            }
        ]
    ]
    
    if is_admin:
        buttons.append([
            {
                "action": {"type": "text", "label": "АДМИН"},
                "color": "negative"
            }
        ])
    
    return {
        "one_time": False,
        "inline": False,
        "buttons": buttons
    }

def create_menu_keyboard_with_test(is_admin: bool = False) -> Dict:
    """Menu with Testing button (for users who completed form)."""
    buttons = [
        [
            {
                "action": {"type": "text", "label": "Меню"},
                "color": "primary"
            },
            {
                "action": {"type": "text", "label": "Тестирование"},
                "color": "positive"
            }
        ]
    ]
    
    if is_admin:
        buttons.append([
            {
                "action": {"type": "text", "label": "АДМИН"},
                "color": "negative"
            }
        ])
    
    return {
        "one_time": False,
        "inline": False,
        "buttons": buttons
    }

def create_admin_keyboard() -> Dict:
    """Admin panel keyboard."""
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
                    "action": {"type": "text", "label": "АДМИН"},
                    "color": "negative"
                }
            ]
        ]
    }

def create_answer_keyboard(shuffled_answers: List[Dict]) -> Dict:
    """Keyboard with Menu button and answer buttons 1, 2, 3."""
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
                    "action": {"type": "text", "label": "1"},
                    "color": "primary"
                },
                {
                    "action": {"type": "text", "label": "2"},
                    "color": "primary"
                },
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
# Database Export Helper
# -----------------------------------------------------------------------------
def create_users_xlsx(users: List[Dict]) -> bytes:
    """Create XLSX file with users data."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Пользователи"
    
    # Define headers
    headers = [
        "ID пользователя",
        "Имя",
        "Анкета заполнена",
        "Ответ анкеты",
        "Тест 1",
        "Тест 2",
        "Тест 3",
        "Тест 4",
        "Дата создания",
        "Дата обновления"
    ]
    
    # Style for headers
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Write headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Write data
    for row_num, user in enumerate(users, 2):
        # Convert boolean values to Russian
        def bool_ru(val):
            return "Да" if val else "Нет"
        
        def format_datetime(dt):
            if dt:
                if isinstance(dt, str):
                    return dt[:19] if len(dt) > 19 else dt
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            return ""
        
        row_data = [
            user.get("user_id", ""),
            user.get("user_name", ""),
            bool_ru(user.get("form", False)),
            user.get("form_answer", ""),
            bool_ru(user.get("test_book_1", False)),
            bool_ru(user.get("test_book_2", False)),
            bool_ru(user.get("test_book_3", False)),
            bool_ru(user.get("test_book_4", False)),
            format_datetime(user.get("created_at", "")),
            format_datetime(user.get("updated_at", ""))
        ]
        
        for col, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col, value=value)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    
    # Adjust column widths
    column_widths = [15, 25, 18, 40, 10, 10, 10, 10, 20, 20]
    for col, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64 + col) if col <= 26 else f"A{chr(64 + col - 26)}"].width = width
    
    # Freeze first row
    ws.freeze_panes = "A2"
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


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
            "check_test_ids": CHECK_TEST_IDS,
            "user_admin_id": USER_ADMIN_ID,
            "tests_loaded": bool(TESTS_DATA),
            "texts_loaded": bool(TEXTS_DATA),
            "database_connected": db.pool is not None
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
            
            # Check if user is admin
            is_admin = (user_id == USER_ADMIN_ID)
            
            # Handle "Start" button
            if text.lower() in ["начать", "start", "/start"]:
                # Clear any existing session
                if user_id in USER_SESSIONS:
                    del USER_SESSIONS[user_id]
                
                # Get user name from VK
                user_name = f"ID{user_id}"
                try:
                    user_info = await self.vk_api.get_user_info(user_id)
                    if "response" in user_info and user_info["response"]:
                        first_name = user_info["response"][0].get("first_name", "")
                        last_name = user_info["response"][0].get("last_name", "")
                        user_name = f"{first_name} {last_name}"
                except Exception as e:
                    print(f"Error getting user info: {e}", flush=True)
                
                # Check/create user in database
                user_data = await db.get_user(user_id)
                if not user_data:
                    # Create new user
                    await db.create_user(user_id, user_name)
                    user_data = await db.get_user(user_id)
                
                # Check FORM status
                form_completed = user_data.get("form", False) if user_data else False
                
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Добро пожаловать!\n\nБот не реагирует на текстовые сообщения. Используйте кнопки клавиатуры.",
                    peer_id=peer_id,
                    keyboard=create_menu_keyboard_with_form(is_admin) if not form_completed else create_menu_keyboard_with_test(is_admin)
                )
                return
            
            # Handle "Меню" button - always available
            if text.lower() == "меню":
                # Clear session if user is in the middle of a test
                if user_id in USER_SESSIONS:
                    del USER_SESSIONS[user_id]
                    print(f"Session cleared for user {user_id} - returned to menu", flush=True)
                
                # Check user FORM status
                user_data = await db.get_user(user_id)
                form_completed = user_data.get("form", False) if user_data else False
                
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Выберите действие:",
                    peer_id=peer_id,
                    keyboard=create_menu_keyboard_with_form(is_admin) if not form_completed else create_menu_keyboard_with_test(is_admin)
                )
                return
            
            # Handle "АДМИН" button - only for USER_ADMIN
            if text.lower() == "админ" and is_admin:
                await self._handle_admin_button(user_id, peer_id)
                return
            
            # Handle "Анкета" button (stub)
            if text.lower() == "анкета":
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Анкета находится в разработке...",
                    peer_id=peer_id,
                    keyboard=create_menu_keyboard_with_form(is_admin)
                )
                return
            
            # Handle "Тестирование" button
            if text.lower() == "тестирование":
                # Check if user completed form
                user_data = await db.get_user(user_id)
                form_completed = user_data.get("form", False) if user_data else False
                
                if not form_completed:
                    await self.vk_api.send_message(
                        user_id=user_id,
                        message="Сначала необходимо заполнить анкету!",
                        peer_id=peer_id,
                        keyboard=create_menu_keyboard_with_form(is_admin)
                    )
                    return
                
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

    async def _handle_admin_button(self, user_id: int, peer_id: int) -> None:
        """Handle admin button - export database to xlsx."""
        print(f"Admin button pressed by user {user_id}", flush=True)
        
        try:
            # Get all users from database
            users = await db.get_all_users()
            
            if not users:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="База данных пуста.",
                    peer_id=peer_id,
                    keyboard=create_admin_keyboard()
                )
                return
            
            # Create XLSX file
            xlsx_data = create_users_xlsx(users)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"users_{timestamp}.xlsx"
            
            # Send document
            success = await self.vk_api.send_document(
                peer_id=peer_id,
                file_data=xlsx_data,
                filename=filename,
                message=f"📊 База данных: {len(users)} пользователей"
            )
            
            if not success:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Ошибка при отправке файла.",
                    peer_id=peer_id,
                    keyboard=create_admin_keyboard()
                )
            
        except Exception as e:
            print(f"Error handling admin button: {e}", flush=True)
            await self.vk_api.send_message(
                user_id=user_id,
                message=f"Ошибка: {e}",
                peer_id=peer_id,
                keyboard=create_admin_keyboard()
            )

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
        
        # Check if user is admin
        is_admin = (user_id == USER_ADMIN_ID)
        
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
        
        # Send notification to all admins
        if CHECK_TEST_IDS:
            status = "СДАЛ" if passed else "НЕ СДАЛ"
            # Create clickable link to user profile
            user_link = f"[id{user_id}|{user_name}]"
            admin_message = f"Пользователь {user_link} {status} тест!\nРезультат: {score}/{total}"
            for admin_id in CHECK_TEST_IDS:
                try:
                    await self.vk_api.send_message(
                        user_id=admin_id,
                        message=admin_message
                    )
                    print(f"Admin notification sent to {admin_id}", flush=True)
                except Exception as e:
                    print(f"Failed to send admin notification to {admin_id}: {e}", flush=True)
        
        # Get user form status for correct keyboard
        user_data = await db.get_user(user_id)
        form_completed = user_data.get("form", False) if user_data else False
        
        # Send result to user
        if passed:
            passed_text = TEXTS_DATA.get("test_passed", "🎉 Поздравляем! Вы сдали тест!")
            message = f"{passed_text}\n\nВаш результат: {score}/{total}"
            await self.vk_api.send_message(
                user_id=user_id,
                message=message,
                peer_id=peer_id,
                keyboard=create_menu_keyboard_with_test(is_admin)
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
async def async_main():
    """Async main entrypoint."""
    print("Creating web server...", flush=True)
    server = WebServer()
    
    # Initialize database
    await db.init()
    
    runner = web.AppRunner(server.app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    
    print(f"Server started on port {PORT}", flush=True)
    
    try:
        await asyncio.Event().wait()
    except KeyboardInterrupt:
        print("Shutting down...", flush=True)
    finally:
        await db.close()
        await runner.cleanup()
        print("Application shutdown complete", flush=True)


if __name__ == "__main__":
    asyncio.run(async_main())
