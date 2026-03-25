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

# Database configuration (Amvera PostgreSQL)
DB_HOST = os.getenv("DB_HOST", "")
DB_NAME = os.getenv("DB_NAME", "")
DB_USER = os.getenv("DB_USER", "")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")

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
print(f"DB_HOST: {'SET' if DB_HOST else 'NOT SET'}", flush=True)
print(f"DB_NAME: {'SET' if DB_NAME else 'NOT SET'}", flush=True)
print(f"DB_USER: {'SET' if DB_USER else 'NOT SET'}", flush=True)
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

# Load tests, texts and form
TESTS_DATA = load_json_file("tests.json")
TEXTS_DATA = load_json_file("texts.json")
FORM_DATA = load_json_file("form.json")

# -----------------------------------------------------------------------------
# User Sessions Storage (in-memory)
# Format: {user_id: {"variant": N, "question": N, "score": N, "shuffled_answers": [...]}}
# -----------------------------------------------------------------------------
USER_SESSIONS: Dict[int, Dict] = {}

# -----------------------------------------------------------------------------
# Form Sessions Storage (in-memory)
# Format: {user_id: {"question": N, "answers": [str, str, ...]}}
# -----------------------------------------------------------------------------
FORM_SESSIONS: Dict[int, Dict] = {}

# -----------------------------------------------------------------------------
# Admin Search Sessions (in-memory)
# Format: {admin_id: {"step": str, "search_text": str, "results": [Dict], "page": int, "selected_user_id": int}}
# Steps: "search", "select_user", "select_course", "confirm_action"
# -----------------------------------------------------------------------------
ADMIN_SEARCH_SESSIONS: Dict[int, Dict] = {}

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
            # Build connection string from Amvera variables
            # Format: postgresql://user:password@host:5432/database
            if DB_HOST and DB_NAME and DB_USER and DB_PASSWORD:
                conn_string = f"postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:5432/{DB_NAME}"
            else:
                print("No database configuration found!", flush=True)
                print(f"DB_HOST: {'SET' if DB_HOST else 'NOT SET'}", flush=True)
                print(f"DB_NAME: {'SET' if DB_NAME else 'NOT SET'}", flush=True)
                print(f"DB_USER: {'SET' if DB_USER else 'NOT SET'}", flush=True)
                print(f"DB_PASSWORD: {'SET' if DB_PASSWORD else 'NOT SET'}", flush=True)
                return False
            
            print(f"Connecting to database {DB_NAME} at {DB_HOST}...", flush=True)
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
                        practice_1 BOOLEAN DEFAULT FALSE,
                        practice_2 BOOLEAN DEFAULT FALSE,
                        practice_3 BOOLEAN DEFAULT FALSE,
                        practice_4 BOOLEAN DEFAULT FALSE,
                        course_1 BOOLEAN DEFAULT FALSE,
                        course_2 BOOLEAN DEFAULT FALSE,
                        course_3 BOOLEAN DEFAULT FALSE,
                        course_4 BOOLEAN DEFAULT FALSE,
                        access_survey_1 BOOLEAN DEFAULT FALSE,
                        access_survey_2 BOOLEAN DEFAULT FALSE,
                        access_survey_3 BOOLEAN DEFAULT FALSE,
                        access_survey_4 BOOLEAN DEFAULT FALSE,
                        fortune_wheel_1 BOOLEAN DEFAULT FALSE,
                        fortune_wheel_2 BOOLEAN DEFAULT FALSE,
                        fortune_wheel_3 BOOLEAN DEFAULT FALSE,
                        fortune_wheel_4 BOOLEAN DEFAULT FALSE,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                print("Database table verified/created", flush=True)
            
            # Add new columns if they don't exist (for existing tables)
            async with self.pool.acquire() as conn:
                new_columns = [
                    ("practice_1", "BOOLEAN DEFAULT FALSE"),
                    ("practice_2", "BOOLEAN DEFAULT FALSE"),
                    ("practice_3", "BOOLEAN DEFAULT FALSE"),
                    ("practice_4", "BOOLEAN DEFAULT FALSE"),
                    ("course_1", "BOOLEAN DEFAULT FALSE"),
                    ("course_2", "BOOLEAN DEFAULT FALSE"),
                    ("course_3", "BOOLEAN DEFAULT FALSE"),
                    ("course_4", "BOOLEAN DEFAULT FALSE"),
                    ("access_survey_1", "BOOLEAN DEFAULT FALSE"),
                    ("access_survey_2", "BOOLEAN DEFAULT FALSE"),
                    ("access_survey_3", "BOOLEAN DEFAULT FALSE"),
                    ("access_survey_4", "BOOLEAN DEFAULT FALSE"),
                    ("fortune_wheel_1", "BOOLEAN DEFAULT FALSE"),
                    ("fortune_wheel_2", "BOOLEAN DEFAULT FALSE"),
                    ("fortune_wheel_3", "BOOLEAN DEFAULT FALSE"),
                    ("fortune_wheel_4", "BOOLEAN DEFAULT FALSE"),
                ]
                for col_name, col_type in new_columns:
                    try:
                        await conn.execute(f"ALTER TABLE users ADD COLUMN IF NOT EXISTS {col_name} {col_type}")
                    except Exception as e:
                        print(f"Column {col_name} might already exist: {e}", flush=True)
                print("Database columns verified/added", flush=True)
            
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
        
        allowed_fields = [
            'form', 'form_answer', 'test_book_1', 'test_book_2', 'test_book_3', 'test_book_4', 'user_name',
            'practice_1', 'practice_2', 'practice_3', 'practice_4',
            'course_1', 'course_2', 'course_3', 'course_4',
            'access_survey_1', 'access_survey_2', 'access_survey_3', 'access_survey_4',
            'fortune_wheel_1', 'fortune_wheel_2', 'fortune_wheel_3', 'fortune_wheel_4'
        ]
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
    
    async def search_users_by_name(self, search_text: str) -> List[Dict]:
        """Search users by name (case-insensitive partial match)."""
        if not self.pool:
            return []
        
        try:
            async with self.pool.acquire() as conn:
                # Case-insensitive search with ILIKE
                rows = await conn.fetch(
                    "SELECT user_id, user_name, form FROM users WHERE user_name ILIKE $1 ORDER BY user_name",
                    f"%{search_text}%"
                )
                return [dict(row) for row in rows]
        except Exception as e:
            print(f"Error searching users by name '{search_text}': {e}", flush=True)
            return []

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
    
    async def get_conversations(self, offset: int = 0, count: int = 200) -> Dict[str, Any]:
        """Get conversations list with pagination."""
        params = {
            "offset": offset,
            "count": count,
            "filter": "all"
        }
        return await self.call("messages.getConversations", params)
    
    async def get_all_conversations(self) -> List[int]:
        """Get all user IDs from conversations with pagination."""
        all_user_ids = []
        offset = 0
        count = 200  # Max per request
        
        while True:
            result = await self.get_conversations(offset=offset, count=count)
            
            if "error" in result:
                print(f"Error getting conversations: {result['error']}", flush=True)
                break
            
            if "response" not in result:
                print(f"No response in conversations result: {result}", flush=True)
                break
            
            items = result["response"].get("items", [])
            if not items:
                break
            
            for item in items:
                conversation = item.get("conversation", {})
                peer = conversation.get("peer", {})
                peer_type = peer.get("type", "")
                peer_id = peer.get("id", 0)
                
                # Only user conversations (not groups, not chats)
                if peer_type == "user" and peer_id > 0:
                    all_user_ids.append(peer_id)
            
            total_count = result["response"].get("count", 0)
            offset += count
            
            print(f"Fetched {len(all_user_ids)} conversations, total available: {total_count}", flush=True)
            
            # Check if we got all
            if offset >= total_count:
                break
        
        return all_user_ids
    
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
    """Admin panel keyboard with submenu."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {"type": "text", "label": "Скачать базу"},
                    "color": "positive"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": "Обновление базы"},
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": "Открыть доступ к Анкете"},
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": "Меню"},
                    "color": "secondary"
                }
            ]
        ]
    }

def create_user_search_keyboard(users: List[Dict], page: int = 0, per_page: int = 6) -> Dict:
    """Create keyboard with found users (6 per page with pagination)."""
    start_idx = page * per_page
    end_idx = start_idx + per_page
    page_users = users[start_idx:end_idx]
    total_pages = (len(users) + per_page - 1) // per_page
    
    buttons = []
    
    # Add user buttons (2 per row, 3 rows = 6 users)
    for i in range(0, len(page_users), 2):
        row = []
        for j in range(2):
            if i + j < len(page_users):
                user = page_users[i + j]
                user_name = user.get("user_name", "Unknown")[:20]  # Limit button text
                user_id = user.get("user_id")
                row.append({
                    "action": {"type": "text", "label": f"👤{user_name}"},
                    "color": "primary"
                })
        if row:
            buttons.append(row)
    
    # Add navigation buttons
    nav_row = [
        {
            "action": {"type": "text", "label": "Меню"},
            "color": "secondary"
        }
    ]
    
    if page > 0:
        nav_row.append({
            "action": {"type": "text", "label": "◀️ Назад"},
            "color": "primary"
        })
    
    if page < total_pages - 1:
        nav_row.append({
            "action": {"type": "text", "label": "Далее ▶️"},
            "color": "primary"
        })
    
    nav_row.append({
        "action": {"type": "text", "label": "🔄 Заново"},
        "color": "primary"
    })
    
    buttons.append(nav_row)
    
    return {
        "one_time": False,
        "inline": False,
        "buttons": buttons
    }

def create_course_selection_keyboard() -> Dict:
    """Keyboard for selecting course number."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
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
                },
                {
                    "action": {"type": "text", "label": "4"},
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": "Меню"},
                    "color": "secondary"
                }
            ]
        ]
    }

def create_access_action_keyboard() -> Dict:
    """Keyboard for opening/closing access."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {"type": "text", "label": "🔓 Открыть"},
                    "color": "positive"
                },
                {
                    "action": {"type": "text", "label": "🔒 Закрыть"},
                    "color": "negative"
                }
            ],
            [
                {
                    "action": {"type": "text", "label": "Меню"},
                    "color": "secondary"
                }
            ]
        ]
    }

def create_menu_keyboard_with_final_survey(is_admin: bool = False) -> Dict:
    """Menu with Testing and Final Survey buttons (for users who have access to final survey)."""
    buttons = [
        [
            {
                "action": {"type": "text", "label": "Меню"},
                "color": "primary"
            },
            {
                "action": {"type": "text", "label": "Финальное анкетирование"},
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

def create_menu_keyboard_after_test_passed(is_admin: bool = False) -> Dict:
    """Menu for users who passed test but haven't done practice yet."""
    buttons = [
        [
            {
                "action": {"type": "text", "label": "Меню"},
                "color": "primary"
            },
            {
                "action": {"type": "text", "label": "Сдал(а) практику"},
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

def create_menu_keyboard_with_practice_and_survey(is_admin: bool = False) -> Dict:
    """Menu for users who passed test and have access to final survey (also show practice button)."""
    buttons = [
        [
            {
                "action": {"type": "text", "label": "Меню"},
                "color": "primary"
            },
            {
                "action": {"type": "text", "label": "Сдал(а) практику"},
                "color": "positive"
            }
        ],
        [
            {
                "action": {"type": "text", "label": "Финальное анкетирование"},
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

def create_form_keyboard() -> Dict:
    """Keyboard with only Menu button for form questions."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {"type": "text", "label": "Меню"},
                    "color": "secondary"
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
            "form_loaded": bool(FORM_DATA),
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
                # Clear any existing sessions
                if user_id in USER_SESSIONS:
                    del USER_SESSIONS[user_id]
                if user_id in FORM_SESSIONS:
                    del FORM_SESSIONS[user_id]
                if user_id in ADMIN_SEARCH_SESSIONS:
                    del ADMIN_SEARCH_SESSIONS[user_id]
                
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
                
                # Check FORM status, test status, and access to final survey
                form_completed = user_data.get("form", False) if user_data else False
                test_passed = user_data.get("test_book_1", False) if user_data else False
                has_final_survey_access = any([
                    user_data.get("access_survey_1", False),
                    user_data.get("access_survey_2", False),
                    user_data.get("access_survey_3", False),
                    user_data.get("access_survey_4", False)
                ]) if user_data else False
                
                # Choose appropriate keyboard
                if has_final_survey_access and test_passed:
                    keyboard = create_menu_keyboard_with_practice_and_survey(is_admin)
                elif test_passed:
                    keyboard = create_menu_keyboard_after_test_passed(is_admin)
                elif form_completed:
                    keyboard = create_menu_keyboard_with_test(is_admin)
                else:
                    keyboard = create_menu_keyboard_with_form(is_admin)
                
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("welcome_message", "Добро пожаловать! Используйте кнопки клавиатуры."),
                    peer_id=peer_id,
                    keyboard=keyboard
                )
                return
            
            # Handle "Меню" button - always available
            if text.lower() == "меню":
                # Clear test session if user is in the middle of a test
                if user_id in USER_SESSIONS:
                    del USER_SESSIONS[user_id]
                    print(f"Test session cleared for user {user_id} - returned to menu", flush=True)
                
                # Clear form session if user is in the middle of form
                if user_id in FORM_SESSIONS:
                    del FORM_SESSIONS[user_id]
                    print(f"Form session cleared for user {user_id} - returned to menu", flush=True)
                
                # Clear admin search session
                if user_id in ADMIN_SEARCH_SESSIONS:
                    del ADMIN_SEARCH_SESSIONS[user_id]
                    print(f"Admin search session cleared for user {user_id} - returned to menu", flush=True)
                
                # Check/create user by ID
                user_data = await db.get_user(user_id)
                if not user_data:
                    # Get user name from VK for new user
                    user_name = f"ID{user_id}"
                    try:
                        user_info = await self.vk_api.get_user_info(user_id)
                        if "response" in user_info and user_info["response"]:
                            first_name = user_info["response"][0].get("first_name", "")
                            last_name = user_info["response"][0].get("last_name", "")
                            user_name = f"{first_name} {last_name}"
                    except Exception as e:
                        print(f"Error getting user info: {e}", flush=True)
                    
                    # Create new user
                    await db.create_user(user_id, user_name)
                    print(f"New user {user_id} ({user_name}) created via Menu button", flush=True)
                    user_data = await db.get_user(user_id)
                
                # Get FORM status, test status, and access to final survey
                form_completed = user_data.get("form", False) if user_data else False
                test_passed = user_data.get("test_book_1", False) if user_data else False
                has_final_survey_access = any([
                    user_data.get("access_survey_1", False),
                    user_data.get("access_survey_2", False),
                    user_data.get("access_survey_3", False),
                    user_data.get("access_survey_4", False)
                ]) if user_data else False
                
                # Choose appropriate keyboard
                if has_final_survey_access and test_passed:
                    keyboard = create_menu_keyboard_with_practice_and_survey(is_admin)
                elif test_passed:
                    keyboard = create_menu_keyboard_after_test_passed(is_admin)
                elif form_completed:
                    keyboard = create_menu_keyboard_with_test(is_admin)
                else:
                    keyboard = create_menu_keyboard_with_form(is_admin)
                
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("menu_select_action", "Выберите действие:"),
                    peer_id=peer_id,
                    keyboard=keyboard
                )
                return
            
            # Handle "АДМИН" button - only for USER_ADMIN (show admin menu)
            if text.lower() == "админ" and is_admin:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("admin_panel_title", "🔧 Админ-панель:\n\nВыберите действие:"),
                    peer_id=peer_id,
                    keyboard=create_admin_keyboard()
                )
                return
            
            # Handle "Скачать базу" button
            if text.lower() == "скачать базу" and is_admin:
                await self._handle_download_db(user_id, peer_id)
                return
            
            # Handle "Обновление базы" button
            if text.lower() == "обновление базы" and is_admin:
                await self._handle_sync_users(user_id, peer_id)
                return
            
            # Handle "Открыть доступ к Анкете" button
            if text.lower() == "открыть доступ к анкете" and is_admin:
                # Clear any existing admin search session
                if user_id in ADMIN_SEARCH_SESSIONS:
                    del ADMIN_SEARCH_SESSIONS[user_id]
                
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("admin_search_prompt", "🔍 Поиск пользователя:\n\nВведите имя или часть имени для поиска:"),
                    peer_id=peer_id,
                    keyboard=create_main_menu_keyboard()
                )
                # Start admin search session
                ADMIN_SEARCH_SESSIONS[user_id] = {
                    "step": "search",
                    "search_text": "",
                    "results": [],
                    "page": 0,
                    "selected_user_id": None,
                    "selected_course": None
                }
                return
            
            # Handle admin search session
            if is_admin and user_id in ADMIN_SEARCH_SESSIONS:
                session = ADMIN_SEARCH_SESSIONS[user_id]
                
                # Handle pagination - "Далее"
                if text.startswith("Далее"):
                    session["page"] += 1
                    await self._show_search_results(user_id, peer_id)
                    return
                
                # Handle pagination - "Назад"
                if text.startswith("◀️") or text == "Назад":
                    session["page"] = max(0, session["page"] - 1)
                    await self._show_search_results(user_id, peer_id)
                    return
                
                # Handle "Заново" - restart search
                if text.startswith("🔄") or text == "Заново":
                    del ADMIN_SEARCH_SESSIONS[user_id]
                    ADMIN_SEARCH_SESSIONS[user_id] = {
                        "step": "search",
                        "search_text": "",
                        "results": [],
                        "page": 0,
                        "selected_user_id": None,
                        "selected_course": None
                    }
                    await self.vk_api.send_message(
                        user_id=user_id,
                        message=TEXTS_DATA.get("admin_search_prompt_retry", "🔍 Введите имя или часть имени для поиска:"),
                        peer_id=peer_id,
                        keyboard=create_main_menu_keyboard()
                    )
                    return
                
                # Handle user selection (username button)
                if text.startswith("👤"):
                    # Extract user name from button and find user in results
                    selected_name = text[1:].strip()  # Remove 👤 prefix
                    # Find user in results by name
                    for user in session["results"]:
                        if user.get("user_name", "").startswith(selected_name):
                            session["selected_user_id"] = user.get("user_id")
                            session["step"] = "select_course"
                            break
                    
                    if session["selected_user_id"]:
                        msg_template = TEXTS_DATA.get("admin_select_course", "Выберите курс для пользователя:\n{user_name}")
                        await self.vk_api.send_message(
                            user_id=user_id,
                            message=msg_template.format(user_name=selected_name),
                            peer_id=peer_id,
                            keyboard=create_course_selection_keyboard()
                        )
                    return
                
                # Handle course selection (1-4) when selecting user
                if session["step"] == "select_course" and text in ["1", "2", "3", "4"]:
                    session["selected_course"] = int(text)
                    
                    # Get selected user info
                    selected_user = await db.get_user(session["selected_user_id"])
                    user_name = selected_user.get("user_name", "Unknown") if selected_user else "Unknown"
                    
                    # Check if practice is completed for this course
                    practice_field = f"practice_{text}"
                    practice_completed = selected_user.get(practice_field, False) if selected_user else False
                    
                    if not practice_completed:
                        # Block access: practice not completed
                        session["step"] = "confirm_action"
                        msg_template = TEXTS_DATA.get("admin_practice_not_completed", "❌ Нельзя открыть доступ к финальной анкете!\n\nПользователь: {user_name}\nКурс: {course}\n\nПричина: Практика не сдана (Практика {course} = Нет)")
                        await self.vk_api.send_message(
                            user_id=user_id,
                            message=msg_template.format(user_name=user_name, course=text),
                            peer_id=peer_id,
                            keyboard=create_admin_keyboard()
                        )
                        del ADMIN_SEARCH_SESSIONS[user_id]
                        return
                    
                    session["step"] = "confirm_action"
                    
                    msg_template = TEXTS_DATA.get("admin_access_action", "Выберите действие с доступом к финальной анкете после прохождения курса {course}:\n\nПользователь: {user_name}")
                    await self.vk_api.send_message(
                        user_id=user_id,
                        message=msg_template.format(course=text, user_name=user_name),
                        peer_id=peer_id,
                        keyboard=create_access_action_keyboard()
                    )
                    return
                
                # Handle "Открыть" - open access
                if session["step"] == "confirm_action" and text.startswith("🔓"):
                    course = session["selected_course"]
                    target_user_id = session["selected_user_id"]
                    field_name = f"access_survey_{course}"
                    
                    # Update database
                    await db.update_user_field(target_user_id, field_name, True)
                    
                    # Get user info
                    target_user = await db.get_user(target_user_id)
                    user_name = target_user.get("user_name", "Unknown") if target_user else "Unknown"
                    
                    # Notify admin
                    msg_template = TEXTS_DATA.get("admin_access_opened", "✅ Доступ к финальной анкете курса {course} ОТКРЫТ для:\n{user_name}")
                    await self.vk_api.send_message(
                        user_id=user_id,
                        message=msg_template.format(course=course, user_name=user_name),
                        peer_id=peer_id,
                        keyboard=create_admin_keyboard()
                    )
                    
                    # Notify user
                    try:
                        msg_template = TEXTS_DATA.get("user_access_opened", "Вам открыт доступ к прохождению анкетирования после прохождения курса {course} - используйте кнопки меню")
                        await self.vk_api.send_message(
                            user_id=target_user_id,
                            message=msg_template.format(course=course),
                            peer_id=target_user_id
                        )
                    except Exception as e:
                        print(f"Failed to notify user {target_user_id}: {e}", flush=True)
                    
                    del ADMIN_SEARCH_SESSIONS[user_id]
                    return
                
                # Handle "Закрыть" - close access
                if session["step"] == "confirm_action" and text.startswith("🔒"):
                    course = session["selected_course"]
                    target_user_id = session["selected_user_id"]
                    field_name = f"access_survey_{course}"
                    
                    # Update database
                    await db.update_user_field(target_user_id, field_name, False)
                    
                    # Get user info
                    target_user = await db.get_user(target_user_id)
                    user_name = target_user.get("user_name", "Unknown") if target_user else "Unknown"
                    
                    # Notify admin
                    msg_template = TEXTS_DATA.get("admin_access_closed", "🔒 Доступ к финальной анкете курса {course} ЗАКРЫТ для:\n{user_name}")
                    await self.vk_api.send_message(
                        user_id=user_id,
                        message=msg_template.format(course=course, user_name=user_name),
                        peer_id=peer_id,
                        keyboard=create_admin_keyboard()
                    )
                    
                    del ADMIN_SEARCH_SESSIONS[user_id]
                    return
                
                # Handle search text input
                if session["step"] == "search" and text:
                    # Search users by name
                    results = await db.search_users_by_name(text)
                    
                    if not results:
                        msg_template = TEXTS_DATA.get("admin_search_not_found", "По запросу \"{query}\" ничего не найдено.\n\nПопробуйте другой поиск:")
                        await self.vk_api.send_message(
                            user_id=user_id,
                            message=msg_template.format(query=text),
                            peer_id=peer_id,
                            keyboard=create_main_menu_keyboard()
                        )
                        return
                    
                    session["search_text"] = text
                    session["results"] = results
                    session["page"] = 0
                    session["step"] = "select_user"
                    
                    await self._show_search_results(user_id, peer_id)
                    return
            
            # Handle "Финальное анкетирование" button
            if text.lower() == "финальное анкетирование":
                # Stub for now
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("final_survey_stub", "Тут будет анкетирование после прохождения курса"),
                    peer_id=peer_id,
                    keyboard=create_menu_keyboard_with_final_survey(is_admin)
                )
                return
            
            # Handle "Сдал(а) практику" button
            if text.lower() == "сдал(а) практику" or text.lower() == "сдала практику" or text.lower() == "сдал практику":
                # Get user info for notification
                user_name = f"ID{user_id}"
                try:
                    user_info = await self.vk_api.get_user_info(user_id)
                    if "response" in user_info and user_info["response"]:
                        first_name = user_info["response"][0].get("first_name", "")
                        last_name = user_info["response"][0].get("last_name", "")
                        user_name = f"{first_name} {last_name}"
                except Exception as e:
                    print(f"Error getting user info: {e}", flush=True)
                
                # Update practice_1 in database
                await db.update_user_field(user_id, "practice_1", True)
                print(f"Practice 1 marked as completed for user {user_id}", flush=True)
                
                # Notify all CHECK_TEST admins
                if CHECK_TEST_IDS:
                    user_link = f"[id{user_id}|{user_name}]"
                    msg_template = TEXTS_DATA.get("practice_notification", "Пользователь {user_link} сдал(а) Практику в Курс 1 - используйте меню, чтобы открыть ему доступ к финальному анкетированию")
                    admin_message = msg_template.format(user_link=user_link)
                    for admin_id in CHECK_TEST_IDS:
                        try:
                            await self.vk_api.send_message(
                                user_id=admin_id,
                                message=admin_message
                            )
                            print(f"Practice notification sent to admin {admin_id}", flush=True)
                        except Exception as e:
                            print(f"Failed to send practice notification to {admin_id}: {e}", flush=True)
                
                # Confirm to user
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("practice_confirmed", "✅ Уведомление о сдаче практики отправлено менеджеру.\n\nОжидайте, вам откроют доступ к финальному анкетированию."),
                    peer_id=peer_id,
                    keyboard=create_menu_keyboard_after_test_passed(is_admin)
                )
                return
            
            # Handle "Анкета" button
            if text.lower() == "анкета":
                # Check if user already completed form
                user_data = await db.get_user(user_id)
                form_completed = user_data.get("form", False) if user_data else False
                
                if form_completed:
                    await self.vk_api.send_message(
                        user_id=user_id,
                        message=TEXTS_DATA.get("form_already_completed", "Вы уже заполнили анкету!"),
                        peer_id=peer_id,
                        keyboard=create_menu_keyboard_with_test(is_admin)
                    )
                    return
                
                # Send warning message
                warning_text = TEXTS_DATA.get("form_warning", "Если оборвать прохождение анкетирования, данные не сохранятся.")
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=warning_text,
                    peer_id=peer_id,
                    keyboard=create_form_keyboard()
                )
                
                # Start form session
                FORM_SESSIONS[user_id] = {
                    "question": 0,
                    "answers": []
                }
                
                # Send start message
                start_text = TEXTS_DATA.get("form_start", "Начинаем заполнение анкеты.")
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=start_text,
                    peer_id=peer_id
                )
                
                # Send first question
                await self._send_form_question(user_id, peer_id)
                return
            
            # Handle "Тестирование" button
            if text.lower() == "тестирование":
                # Check if user completed form
                user_data = await db.get_user(user_id)
                form_completed = user_data.get("form", False) if user_data else False
                
                if not form_completed:
                    await self.vk_api.send_message(
                        user_id=user_id,
                        message=TEXTS_DATA.get("test_not_available", "Сначала необходимо заполнить анкету!"),
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
            
            # Handle form answer (if user is in form session)
            if user_id in FORM_SESSIONS:
                await self._handle_form_answer(user_id, peer_id, text)
                return
            
            # Ignore all other text
            print(f"Ignoring text message from user {user_id}: {text}", flush=True)
            
        except Exception as e:
            print(f"Message handling error: {e}", flush=True)

    async def _handle_download_db(self, user_id: int, peer_id: int) -> None:
        """Handle 'Скачать базу' button - export database to xlsx."""
        print(f"Download DB button pressed by user {user_id}", flush=True)
        
        try:
            # Get all users from database
            users = await db.get_all_users()
            
            if not users:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("db_empty", "База данных пуста."),
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
            msg_template = TEXTS_DATA.get("db_export_success", "📊 База данных: {count} пользователей")
            success = await self.vk_api.send_document(
                peer_id=peer_id,
                file_data=xlsx_data,
                filename=filename,
                message=msg_template.format(count=len(users))
            )
            
            if not success:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("db_export_error", "Ошибка при отправке файла."),
                    peer_id=peer_id,
                    keyboard=create_admin_keyboard()
                )
            
        except Exception as e:
            print(f"Error handling download DB: {e}", flush=True)
            msg_template = TEXTS_DATA.get("db_export_error_detail", "Ошибка: {error}")
            await self.vk_api.send_message(
                user_id=user_id,
                message=msg_template.format(error=e),
                peer_id=peer_id,
                keyboard=create_admin_keyboard()
            )

    async def _handle_sync_users(self, user_id: int, peer_id: int) -> None:
        """Handle 'Обновление базы' button - sync all VK conversations to database."""
        print(f"Sync users button pressed by user {user_id}", flush=True)
        
        try:
            # Notify admin that sync started
            await self.vk_api.send_message(
                user_id=user_id,
                message=TEXTS_DATA.get("db_sync_start", "🔄 Начинаю синхронизацию пользователей из чатов..."),
                peer_id=peer_id,
                keyboard=create_admin_keyboard()
            )
            
            # Get all conversations
            all_user_ids = await self.vk_api.get_all_conversations()
            
            if not all_user_ids:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=TEXTS_DATA.get("db_sync_no_chats", "Чаты не найдены."),
                    peer_id=peer_id,
                    keyboard=create_admin_keyboard()
                )
                return
            
            print(f"Found {len(all_user_ids)} user conversations", flush=True)
            
            # Get existing users from database
            existing_users = await db.get_all_users()
            existing_ids = set(u.get("user_id") for u in existing_users) if existing_users else set()
            
            # Find new users
            new_user_ids = [uid for uid in all_user_ids if uid not in existing_ids]
            
            if not new_user_ids:
                msg_template = TEXTS_DATA.get("db_sync_no_new", "✅ Синхронизация завершена.\n\nВсего чатов: {total}\nНовых пользователей: 0\nВсе уже есть в базе.")
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=msg_template.format(total=len(all_user_ids)),
                    peer_id=peer_id,
                    keyboard=create_admin_keyboard()
                )
                return
            
            # Create new users
            created_count = 0
            for new_user_id in new_user_ids:
                # Get user name from VK
                user_name = f"ID{new_user_id}"
                try:
                    user_info = await self.vk_api.get_user_info(new_user_id)
                    if "response" in user_info and user_info["response"]:
                        first_name = user_info["response"][0].get("first_name", "")
                        last_name = user_info["response"][0].get("last_name", "")
                        user_name = f"{first_name} {last_name}"
                except Exception as e:
                    print(f"Error getting user info for {new_user_id}: {e}", flush=True)
                
                # Create user in database
                success = await db.create_user(new_user_id, user_name)
                if success:
                    created_count += 1
                    print(f"Created user {new_user_id} ({user_name})", flush=True)
            
            # Send result
            msg_template = TEXTS_DATA.get("db_sync_success", "✅ Синхронизация завершена.\n\nВсего чатов: {total}\nУже в базе: {existing}\nДобавлено новых: {created}")
            await self.vk_api.send_message(
                user_id=user_id,
                message=msg_template.format(total=len(all_user_ids), existing=len(existing_ids), created=created_count),
                peer_id=peer_id,
                keyboard=create_admin_keyboard()
            )
            
        except Exception as e:
            print(f"Error handling sync users: {e}", flush=True)
            msg_template = TEXTS_DATA.get("db_sync_error", "❌ Ошибка синхронизации: {error}")
            await self.vk_api.send_message(
                user_id=user_id,
                message=msg_template.format(error=e),
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
        has_final_survey_access = any([
            user_data.get("access_survey_1", False),
            user_data.get("access_survey_2", False),
            user_data.get("access_survey_3", False),
            user_data.get("access_survey_4", False)
        ]) if user_data else False
        
        # Send result to user
        if passed:
            # Update test_book_1 in database
            await db.update_user_field(user_id, "test_book_1", True)
            
            passed_text = TEXTS_DATA.get("test_passed", "🎉 Поздравляем! Вы сдали тест!")
            practice_info = "\n\nЕсли вы сдали практику - используйте кнопки Меню для уведомления менеджера. Он откроет вам доступ к финальному анкетированию и получению диплома о прохождении курса."
            message = f"{passed_text}\n\nВаш результат: {score}/{total}{practice_info}"
            
            # Use keyboard with practice button
            if has_final_survey_access:
                keyboard = create_menu_keyboard_with_practice_and_survey(is_admin)
            else:
                keyboard = create_menu_keyboard_after_test_passed(is_admin)
            
            await self.vk_api.send_message(
                user_id=user_id,
                message=message,
                peer_id=peer_id,
                keyboard=keyboard
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

    async def _send_form_question(self, user_id: int, peer_id: int) -> None:
        """Send current form question to user."""
        session = FORM_SESSIONS.get(user_id)
        if not session:
            print(f"No form session for user {user_id}", flush=True)
            return
        
        questions = FORM_DATA.get("questions", [])
        question_idx = session["question"]
        
        if question_idx >= len(questions):
            print(f"Question index out of range: {question_idx}", flush=True)
            return
        
        question = questions[question_idx]
        total = len(questions)
        
        # Format question message
        prefix_template = TEXTS_DATA.get("form_question_prefix", "Вопрос {current}/{total}:")
        prefix = prefix_template.format(current=question_idx + 1, total=total)
        message = f"{prefix}\n\n{question['question']}"
        
        await self.vk_api.send_message(
            user_id=user_id,
            message=message,
            peer_id=peer_id,
            keyboard=create_form_keyboard()
        )

    async def _handle_form_answer(self, user_id: int, peer_id: int, answer: str) -> None:
        """Handle user's answer to form question."""
        session = FORM_SESSIONS.get(user_id)
        if not session:
            print(f"No form session for user {user_id}", flush=True)
            return
        
        # Save answer
        session["answers"].append(answer)
        
        # Move to next question
        session["question"] += 1
        
        questions = FORM_DATA.get("questions", [])
        
        # Check if form is complete
        if session["question"] >= len(questions):
            await self._finish_form(user_id, peer_id)
        else:
            await self._send_form_question(user_id, peer_id)

    async def _finish_form(self, user_id: int, peer_id: int) -> None:
        """Finish form and save results."""
        session = FORM_SESSIONS.get(user_id)
        if not session:
            return
        
        answers = session["answers"]
        questions = FORM_DATA.get("questions", [])
        
        # Get user name from first answer (trimmed)
        user_name = answers[0].strip() if answers else f"ID{user_id}"
        
        # Format form answers as readable text
        form_answer_lines = []
        for i, (q, a) in enumerate(zip(questions, answers), 1):
            form_answer_lines.append(f"Вопрос {i}: {q['question']}")
            form_answer_lines.append(f"Ответ {i}: {a}")
            if i < len(questions):
                form_answer_lines.append("")  # Empty line between Q&A
        form_answer = "\n".join(form_answer_lines)
        
        # Update user in database: name, answers, form completed
        await db.update_user_field(user_id, "user_name", user_name)
        await db.update_user_field(user_id, "form_answer", form_answer)
        await db.update_user_field(user_id, "form", True)
        
        print(f"Form completed for user {user_id}, name: {user_name}", flush=True)
        
        # Check if user is admin
        is_admin = (user_id == USER_ADMIN_ID)
        
        # Send completion message
        completed_text = TEXTS_DATA.get("form_completed", "✅ Анкета успешно заполнена!")
        await self.vk_api.send_message(
            user_id=user_id,
            message=completed_text,
            peer_id=peer_id,
            keyboard=create_menu_keyboard_with_test(is_admin)
        )
        
        # Send instructions
        instructions_text = TEXTS_DATA.get("instructions", "Инструкция")
        await self.vk_api.send_message(
            user_id=user_id,
            message=instructions_text,
            peer_id=peer_id,
            keyboard=create_menu_keyboard_with_test(is_admin)
        )
        
        # Notify admin about form completion with full answers
        if USER_ADMIN_ID:
            # Create clickable link to user profile
            user_link = f"[id{user_id}|{user_name}]"
            admin_message = f"Пользователь {user_link}, прошел анкетирование!\n\n{form_answer}"
            
            try:
                await self.vk_api.send_message(
                    user_id=USER_ADMIN_ID,
                    message=admin_message
                )
                print(f"Admin notification sent about form completion by user {user_id}", flush=True)
            except Exception as e:
                print(f"Failed to send admin notification: {e}", flush=True)
        
        # Clear form session
        del FORM_SESSIONS[user_id]

    async def _show_search_results(self, user_id: int, peer_id: int) -> None:
        """Show search results with pagination."""
        session = ADMIN_SEARCH_SESSIONS.get(user_id)
        if not session:
            return
        
        results = session["results"]
        page = session["page"]
        per_page = 6
        total_pages = (len(results) + per_page - 1) // per_page
        
        # Create keyboard with users
        keyboard = create_user_search_keyboard(results, page, per_page)
        
        # Create message
        search_text = session["search_text"]
        total = len(results)
        start_idx = page * per_page + 1
        end_idx = min((page + 1) * per_page, total)
        
        message = f"🔍 Результаты поиска \"{search_text}\":\n\n"
        message += f"Найдено: {total} пользователей\n"
        message += f"Страница {page + 1}/{total_pages}\n\n"
        message += "Выберите пользователя:"
        
        await self.vk_api.send_message(
            user_id=user_id,
            message=message,
            peer_id=peer_id,
            keyboard=keyboard
        )


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
