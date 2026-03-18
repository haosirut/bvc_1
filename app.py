import os
import json
import logging
import asyncio
import random
from typing import Any, Dict, Optional

import aiohttp
from aiohttp import web

# -----------------------------------------------------------------------------
# Environment Variables
# -----------------------------------------------------------------------------
TOKEN = os.getenv("TOKEN", "")  # VK Access Token
CONFIRMATION_TOKEN = os.getenv("CONFIRMATION_TOKEN", "")  # VK Confirmation Token
PORT = int(os.getenv("PORT", "8080"))

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("vk_bot")

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


# -----------------------------------------------------------------------------
# Keyboard Builder
# -----------------------------------------------------------------------------
def create_main_keyboard() -> Dict:
    """Create keyboard with buttons 1, 2, 3."""
    return {
        "one_time": False,
        "inline": False,
        "buttons": [
            [
                {
                    "action": {
                        "type": "text",
                        "label": "1"
                    },
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {
                        "type": "text",
                        "label": "2"
                    },
                    "color": "primary"
                }
            ],
            [
                {
                    "action": {
                        "type": "text",
                        "label": "3"
                    },
                    "color": "primary"
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

    def _setup_routes(self) -> None:
        """Setup routes for web server."""
        # VK sends POST to root URL for confirmation
        self.app.router.add_post("/", self.vk_webhook)
        self.app.router.add_post("/webhook", self.vk_webhook)
        # Health check
        self.app.router.add_get("/", self.health)
        self.app.router.add_get("/health", self.health)

    async def health(self, request: web.Request) -> web.Response:
        """Health check endpoint."""
        logger.info("Health check requested")
        return web.json_response({
            "status": "ok",
            "token_configured": bool(TOKEN),
            "confirmation_token_configured": bool(CONFIRMATION_TOKEN)
        })

    async def vk_webhook(self, request: web.Request) -> web.Response:
        """Handle VK Callback API webhook."""
        try:
            # Read request body
            body = await request.text()
            logger.info(f"Received POST request: {body[:500]}")
            
            # Parse JSON
            try:
                data = json.loads(body)
            except json.JSONDecodeError as e:
                logger.error(f"Failed to parse JSON: {e}")
                return web.Response(text="invalid json", status=400)
            
            event_type = data.get("type", "")
            group_id = data.get("group_id", 0)
            
            logger.info(f"Event type: {event_type}, group_id: {group_id}")
            
            # Handle confirmation
            if event_type == "confirmation":
                logger.info(f"Confirmation request received")
                if CONFIRMATION_TOKEN:
                    logger.info(f"Returning confirmation token")
                    return web.Response(text=CONFIRMATION_TOKEN)
                else:
                    logger.error("CONFIRMATION_TOKEN not configured!")
                    return web.Response(text="error", status=500)
            
            # Handle message_new
            if event_type == "message_new":
                logger.info("Handling message_new")
                await self._handle_message_new(data)
                return web.Response(text="ok")
            
            # Unknown event type - still return ok
            logger.info(f"Unknown event type: {event_type}, returning ok")
            return web.Response(text="ok")
            
        except Exception as e:
            logger.exception(f"Webhook error: {e}")
            return web.Response(text="ok")

    async def _handle_message_new(self, data: Dict) -> None:
        """Handle message_new event."""
        try:
            if not self.vk_api:
                logger.error("VK API not initialized (TOKEN missing)")
                return
            
            await self.vk_api.init()
            
            obj = data.get("object", {})
            message = obj.get("message", obj)
            
            user_id = message.get("from_id", message.get("user_id", 0))
            peer_id = message.get("peer_id", user_id)
            text = message.get("text", "").strip()
            
            logger.info(f"Message from user {user_id}: {text}")
            
            # Handle "Start" button
            if text.lower() in ["начать", "start", "/start"]:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message="Добро пожаловать! Выберите один из вариантов:",
                    peer_id=peer_id,
                    keyboard=create_main_keyboard()
                )
                return
            
            # Handle button presses
            if text in ["1", "2", "3"]:
                await self.vk_api.send_message(
                    user_id=user_id,
                    message=f"Вы выбрали вариант: {text}",
                    peer_id=peer_id,
                    keyboard=create_main_keyboard()
                )
                return
            
            # Default response
            await self.vk_api.send_message(
                user_id=user_id,
                message="Выберите один из вариантов:",
                peer_id=peer_id,
                keyboard=create_main_keyboard()
            )
            
        except Exception as e:
            logger.exception(f"Error handling message: {e}")


# -----------------------------------------------------------------------------
# Entrypoint
# -----------------------------------------------------------------------------
async def main(port: int = 8080) -> None:
    server = WebServer()
    
    runner = web.AppRunner(server.app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    
    logger.info(f"VK Bot server started on 0.0.0.0:{port}")
    logger.info(f"TOKEN configured: {bool(TOKEN)}")
    logger.info(f"CONFIRMATION_TOKEN configured: {bool(CONFIRMATION_TOKEN)}")
    if CONFIRMATION_TOKEN:
        logger.info(f"CONFIRMATION_TOKEN value: {CONFIRMATION_TOKEN}")
    
    try:
        await asyncio.Event().wait()
    except KeyboardInterrupt:
        logger.info("Shutting down...")
    finally:
        if server.vk_api:
            await server.vk_api.close()
        await runner.cleanup()
        logger.info("Application shutdown complete")


if __name__ == "__main__":
    logger.info("Starting VK Bot...")
    logger.info(f"TOKEN: {'configured' if TOKEN else 'NOT SET'}")
    logger.info(f"CONFIRMATION_TOKEN: {'configured' if CONFIRMATION_TOKEN else 'NOT SET'}")
    asyncio.run(main(PORT))
