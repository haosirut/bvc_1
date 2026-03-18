import os
import json
import logging
import asyncio
import hashlib
from typing import Any, Dict, Optional

import aiohttp
from aiohttp import web

# -----------------------------------------------------------------------------
# Environment Variables
# -----------------------------------------------------------------------------
TOKEN = os.getenv("TOKEN", "")  # VK Access Token
CONFIRMATION_TOKEN = os.getenv("CONFIRMATION_TOKEN", "")  # VK Confirmation Token
PORT = int(os.getenv("PORT", "8080"))
LOGGING = os.getenv("LOGGING", "True").lower() in ("true", "1", "yes", "on")

logger = logging.getLogger("vk_bot")

# -----------------------------------------------------------------------------
# Logging Setup
# -----------------------------------------------------------------------------
def _setup_logging():
    """Setup logging based on LOGGING environment variable."""
    level = logging.DEBUG if LOGGING else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(),
        ]
    )

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
            "user_id": user_id,
            "message": message,
            "random_id": int(asyncio.get_event_loop().time() * 1000000)
        }
        
        if peer_id:
            params["peer_id"] = peer_id
            if "user_id" in params:
                del params["user_id"]
        
        if keyboard:
            params["keyboard"] = json.dumps(keyboard, ensure_ascii=False)
        
        return await self.call("messages.send", params)
    
    async def send_message_event_answer(
        self, 
        event_id: str, 
        user_id: int, 
        peer_id: int,
        event_data: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Answer to message event (callback button)."""
        params = {
            "event_id": event_id,
            "user_id": user_id,
            "peer_id": peer_id
        }
        
        if event_data:
            params["event_data"] = json.dumps(event_data, ensure_ascii=False)
        
        return await self.call("messages.sendMessageEventAnswer", params)


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
        return web.json_response({
            "status": "ok",
            "token_configured": bool(TOKEN),
            "confirmation_token_configured": bool(CONFIRMATION_TOKEN)
        })

    async def vk_webhook(self, request: web.Request) -> web.Response:
        """
        Handle VK Callback API webhook.
        
        VK sends events to this endpoint.
        Types:
        - confirmation: Server verification (return confirmation token)
        - message_new: New message
        - message_event: Callback button pressed
        """
        try:
            # Log raw request
            logger.info(f"Received request: method={request.method}, path={request.path}")
            
            # Read request body
            body = await request.text()
            logger.info(f"Request body: {body}")
            
            try:
                data = json.loads(body)
            except json.JSONDecodeError as e:
                logger.error(f"Failed to parse JSON: {e}")
                return web.Response(text="invalid json", status=400)
            
            event_type = data.get("type", "")
            group_id = data.get("group_id", 0)
            
            logger.info(f"VK webhook received: type={event_type}, group_id={group_id}")
            
            # Handle confirmation
            if event_type == "confirmation":
                if not CONFIRMATION_TOKEN:
                    logger.error("CONFIRMATION_TOKEN not configured")
                    return web.Response(text="error", status=500)
                logger.info(f"Returning confirmation token: {CONFIRMATION_TOKEN}")
                return web.Response(text=CONFIRMATION_TOKEN)
            
            # Handle message_new
            if event_type == "message_new":
                await self._handle_message_new(data)
                return web.Response(text="ok")
            
            # Handle message_event (callback buttons)
            if event_type == "message_event":
                await self._handle_message_event(data)
                return web.Response(text="ok")
            
            # Unknown event type
            logger.warning(f"Unknown event type: {event_type}")
            return web.Response(text="ok")
            
        except Exception as e:
            logger.exception(f"Webhook error: {e}")
            return web.Response(text="ok")

    async def _handle_message_new(self, data: Dict) -> None:
        """Handle message_new event."""
        if not self.vk_api:
            logger.error("VK API not initialized (TOKEN missing)")
            return
        
        await self.vk_api.init()
        
        obj = data.get("object", {})
        message = obj.get("message", obj)  # Different VK API versions
        client_info = obj.get("client_info", {})
        
        user_id = message.get("from_id", message.get("user_id", 0))
        peer_id = message.get("peer_id", user_id)
        text = message.get("text", "").strip()
        payload = message.get("payload", "")
        
        logger.info(f"Message from user_id={user_id}: text='{text}', payload='{payload}'")
        
        # Parse payload if it's a JSON string
        if payload:
            try:
                payload_obj = json.loads(payload)
                payload = payload_obj
            except (json.JSONDecodeError, TypeError):
                pass
        
        # Handle "Start" button (payload = "start" or text = "Начать" or first message)
        if payload == "start" or payload == {"command": "start"} or text.lower() in ["начать", "start", "/start"]:
            await self._send_welcome_keyboard(user_id, peer_id)
            return
        
        # Handle button presses
        if text in ["1", "2", "3"]:
            response_text = f"Вы выбрали вариант: {text}"
            await self.vk_api.send_message(
                user_id=user_id,
                message=response_text,
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

    async def _handle_message_event(self, data: Dict) -> None:
        """Handle message_event (callback button press)."""
        if not self.vk_api:
            logger.error("VK API not initialized (TOKEN missing)")
            return
        
        await self.vk_api.init()
        
        obj = data.get("object", {})
        user_id = obj.get("user_id", 0)
        peer_id = obj.get("peer_id", user_id)
        event_id = obj.get("event_id", "")
        payload = obj.get("payload", {})
        
        logger.info(f"Message event from user_id={user_id}: payload={payload}")
        
        # Handle start button
        if payload == "start" or (isinstance(payload, dict) and payload.get("command") == "start"):
            # Send welcome message with keyboard
            await self._send_welcome_keyboard(user_id, peer_id)
            
            # Answer the event (show snackbar or just acknowledge)
            await self.vk_api.send_message_event_answer(
                event_id=event_id,
                user_id=user_id,
                peer_id=peer_id
            )
            return
        
        # Handle button choices
        if isinstance(payload, dict) and "button" in payload:
            button = payload["button"]
            await self.vk_api.send_message(
                user_id=user_id,
                message=f"Вы выбрали вариант: {button}",
                peer_id=peer_id,
                keyboard=create_main_keyboard()
            )
            
            # Answer the event
            await self.vk_api.send_message_event_answer(
                event_id=event_id,
                user_id=user_id,
                peer_id=peer_id
            )
            return
        
        # Default: just acknowledge
        await self.vk_api.send_message_event_answer(
            event_id=event_id,
            user_id=user_id,
            peer_id=peer_id
        )

    async def _send_welcome_keyboard(self, user_id: int, peer_id: int) -> None:
        """Send welcome message with main keyboard."""
        if not self.vk_api:
            return
        
        await self.vk_api.send_message(
            user_id=user_id,
            message="Добро пожаловать! Выберите один из вариантов:",
            peer_id=peer_id,
            keyboard=create_main_keyboard()
        )


# -----------------------------------------------------------------------------
# Entrypoint
# -----------------------------------------------------------------------------
async def main(port: int = 8080) -> None:
    server = WebServer()
    
    runner = web.AppRunner(server.app, shutdown_timeout=600)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    logger.info(f"VK Bot server started on 0.0.0.0:{port}")
    logger.info(f"TOKEN configured: {bool(TOKEN)}")
    logger.info(f"CONFIRMATION_TOKEN configured: {bool(CONFIRMATION_TOKEN)}")
    
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
    _setup_logging()
    port = PORT
    asyncio.run(main(port))
