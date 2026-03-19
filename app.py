import os
import sys
import json
import logging
import asyncio
import signal
import random
from typing import Any, Dict, Optional

import aiohttp
from aiohttp import web

# Force unbuffered output
sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

# -----------------------------------------------------------------------------
# Environment Variables
# -----------------------------------------------------------------------------
TOKEN = os.getenv("TOKEN", "")  # VK Access Token
CONFIRMATION_TOKEN = os.getenv("CONFIRMATION_TOKEN", "")  # VK Confirmation Token
ADMIN_IDS_STR = os.getenv("ADMIN_ID", "535097409")  # Admin IDs separated by comma
ADMIN_CHAT = os.getenv("ADMIN_CHAT", "")  # Peer ID of admin chat (e.g., 2000000002)
PORT = int(os.getenv("PORT", "8080"))

# Parse admin IDs from comma-separated string
ADMIN_IDS = [int(id.strip()) for id in ADMIN_IDS_STR.split(",") if id.strip()]

# Setup logging with forced stdout
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger("vk_bot")

# Print immediately
print("=" * 50, flush=True)
print("VK BOT STARTING", flush=True)
print(f"TOKEN: {'SET' if TOKEN else 'NOT SET'}", flush=True)
print(f"CONFIRMATION_TOKEN: {'SET' if CONFIRMATION_TOKEN else 'NOT SET'}", flush=True)
print(f"ADMIN_IDS: {ADMIN_IDS}", flush=True)
print(f"ADMIN_CHAT: {ADMIN_CHAT}", flush=True)
print(f"PORT: {PORT}", flush=True)
print("=" * 50, flush=True)

# -----------------------------------------------------------------------------
# VK API Helper
# -----------------------------------------------------------------------------
class VKAPI:
    """Helper class for VK API calls."""
    
    API_URL = "https://api.vk.com/method/"
    API_VERSION = "5.199"
    
    def __init__(self, token: str):
        self.token = token
        self.session: Optional[aiohttp.ClientSession] = None
    
    async def init(self):
        """Initialize aiohttp session."""
        if self.session is None:
            self.session = aiohttp.ClientSession()
    
    async def close(self):
        """Close aiohttp session."""
        if self.session:
            await self.session.close()
            self.session = None
    
    async def call(self, method: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """Call VK API method."""
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
        """Send message to user."""
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
        params = {
            "user_ids": user_id
        }
        result = await self.call("users.get", params)
        return result


# -----------------------------------------------------------------------------
# Keyboard Builders
# -----------------------------------------------------------------------------
def create_main_keyboard() -> Dict:
    """Create main keyboard with Menu, Yes, No buttons. Always visible."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {
                        "type": "text",
                        "label": "Меню"
                    },
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {
                        "type": "text",
                        "label": "Да"
                    },
                    "color": "positive"
                },
                {
                    "action": {
                        "type": "text",
                        "label": "Нет"
                    },
                    "color": "negative"
                }
            ]
        ]
    }


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
        """Setup routes for web server."""
        # VK sends POST to root URL for confirmation
        self.app.router.add_post("/", self.vk_webhook)
        self.app.router.add_post("/webhook", self.vk_webhook)
        # Health check
        self.app.router.add_get("/", self.health)
        self.app.router.add_get("/health", self.health)
        print("Routes setup complete", flush=True)

    async def health(self, request: web.Request) -> web.Response:
        """Health check endpoint."""
        print("Health check requested", flush=True)
        return web.json_response({
            "status": "ok",
            "token_configured": bool(TOKEN),
            "confirmation_token_configured": bool(CONFIRMATION_TOKEN),
            "admin_ids": ADMIN_IDS,
            "admin_chat": ADMIN_CHAT
        })

    async def vk_webhook(self, request: web.Request) -> web.Response:
        """Handle VK Callback API webhook."""
        try:
            # Read request body
            body = await request.text()
            print(f"POST received: {body[:500]}", flush=True)
            
            # Parse JSON
            try:
                data = json.loads(body)
            except json.JSONDecodeError as e:
                print(f"JSON parse error: {e}", flush=True)
                return web.Response(text="invalid json", status=400)
            
            event_type = data.get("type", "")
            group_id = data.get("group_id", 0)
            
            print(f"Event: {event_type}, group: {group_id}", flush=True)
            
            # Handle confirmation
            if event_type == "confirmation":
                print(f"CONFIRMATION REQUEST - returning: {CONFIRMATION_TOKEN}", flush=True)
                return web.Response(text=CONFIRMATION_TOKEN)
            
            # Handle message_new
            if event_type == "message_new":
                print("Handling message_new", flush=True)
                await self._handle_message_new(data)
                return web.Response(text="ok")
            
            # Unknown event
            print(f"Unknown event: {event_type}", flush=True)
            return web.Response(text="ok")
            
        except Exception as e:
            print(f"ERROR: {e}", flush=True)
            return web.Response(text="ok")

    async def _handle_message_new(self, data: Dict) -> None:
        """Handle message_new event."""
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
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Добро пожаловать!\n\nБот не реагирует на текстовые сообщения. Используйте кнопки клавиатуры.",
                    peer_id=peer_id,
                    keyboard=create_main_keyboard()
                )
                return
            
            # Handle "Меню" button (no functionality yet)
            if text.lower() == "меню":
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Меню в разработке...",
                    peer_id=peer_id,
                    keyboard=create_main_keyboard()
                )
                return
            
            # Handle "Да" button
            if text.lower() == "да":
                # Get user info
                user_info = await self.vk_api.get_user_info(user_id)
                user_name = "Пользователь"
                if "response" in user_info and user_info["response"]:
                    first_name = user_info["response"][0].get("first_name", "")
                    last_name = user_info["response"][0].get("last_name", "")
                    user_name = f"{first_name} {last_name}"
                
                admin_message = f"Пользователь {user_name} (ID: {user_id}) нажал 'Да'!"
                
                # Send notification to admin chat (if configured)
                if ADMIN_CHAT:
                    try:
                        await self.vk_api.send_message(
                            user_id=0,
                            message=admin_message,
                            peer_id=int(ADMIN_CHAT)
                        )
                        print(f"Notification sent to admin chat {ADMIN_CHAT}", flush=True)
                    except Exception as e:
                        print(f"Failed to send to admin chat: {e}", flush=True)
                else:
                    print("ADMIN_CHAT not configured", flush=True)
                
                # Confirm to user
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Спасибо за ваш ответ! Администратор получит уведомление.",
                    peer_id=peer_id,
                    keyboard=create_main_keyboard()
                )
                return
            
            # Handle "Нет" button - do nothing
            if text.lower() == "нет":
                print(f"User {user_id} pressed 'Нет' - no action", flush=True)
                return
            
            # Ignore all other text input
            print(f"Ignoring text message from user {user_id}: {text}", flush=True)
            # No response - just ignore
            
        except Exception as e:
            print(f"Message handling error: {e}", flush=True)


# -----------------------------------------------------------------------------
# Entrypoint
# -----------------------------------------------------------------------------
def main_sync():
    """Synchronous entry point."""
    print("Creating web server...", flush=True)
    server = WebServer()
    
    print(f"Starting on port {PORT}...", flush=True)
    web.run_app(server.app, host="0.0.0.0", port=PORT, print=lambda x: print(x, flush=True))


if __name__ == "__main__":
    main_sync()
