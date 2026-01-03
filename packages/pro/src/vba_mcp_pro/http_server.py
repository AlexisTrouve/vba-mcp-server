#!/usr/bin/env python3
"""
VBA MCP Server Pro - HTTP/SSE Transport

OS-agnostic HTTP/SSE server that allows remote connections to the VBA MCP server.
Perfect for connecting from WSL, macOS, or Linux to a Windows machine running the server.

Usage:
    # On Windows (where COM works):
    python -m vba_mcp_pro.http_server --host 0.0.0.0 --port 8000

    # From WSL/Mac/Linux, configure Claude Code:
    {
      "mcpServers": {
        "vba-pro": {
          "url": "http://localhost:8000/sse"
        }
      }
    }

Version: 0.1.0
Author: Alexis Trouve
"""

import asyncio
import argparse
import logging
from pathlib import Path

from starlette.applications import Starlette
from starlette.routing import Route, Mount
from starlette.responses import Response
from starlette.requests import Request
import uvicorn

from mcp.server.sse import SseServerTransport

# Import the existing MCP app
from .server import app, _check_wsl_environment
from .session_manager import OfficeSessionManager

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


async def sse_handler(scope, receive, send):
    """ASGI app for handling SSE connections."""
    if scope["type"] == "http" and scope["method"] == "GET":
        async with sse.connect_sse(scope, receive, send) as streams:
            await app.run(
                streams[0],
                streams[1],
                app.create_initialization_options()
            )
    else:
        # Method not allowed
        await send({
            "type": "http.response.start",
            "status": 405,
            "headers": [[b"content-type", b"text/plain"]],
        })
        await send({
            "type": "http.response.body",
            "body": b"Method Not Allowed",
        })


def create_app(sse_transport: SseServerTransport) -> Starlette:
    """Create the Starlette application with SSE routes."""
    routes = [
        Mount("/sse", app=sse_handler),
        Mount("/messages", app=sse_transport.handle_post_message),
    ]

    return Starlette(routes=routes)


# Global SSE transport
sse = SseServerTransport("/messages/")


async def startup():
    """Initialize resources on startup."""
    logger.info("Starting VBA MCP Server Pro (HTTP/SSE)")

    # Check for WSL environment
    _check_wsl_environment()

    # Initialize session manager
    manager = OfficeSessionManager.get_instance()
    manager.start_cleanup_task()

    logger.info("Server ready to accept connections")


async def shutdown():
    """Cleanup resources on shutdown."""
    logger.info("Shutting down VBA MCP Server Pro")

    # Cleanup all sessions
    manager = OfficeSessionManager.get_instance()
    await manager.stop_cleanup_task()
    await manager.close_all_sessions(save=True)

    logger.info("Server shutdown complete")


def main():
    """Main entry point for HTTP/SSE server."""
    parser = argparse.ArgumentParser(
        description="VBA MCP Server Pro - HTTP/SSE Transport"
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host to bind to (default: 127.0.0.1, use 0.0.0.0 for all interfaces)"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port to listen on (default: 8000)"
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level (default: INFO)"
    )

    args = parser.parse_args()

    # Set logging level
    logging.getLogger().setLevel(getattr(logging, args.log_level))

    # Create Starlette app
    starlette_app = create_app(sse)

    # Add lifecycle handlers
    starlette_app.add_event_handler("startup", startup)
    starlette_app.add_event_handler("shutdown", shutdown)

    # Run server
    logger.info(f"Starting server on {args.host}:{args.port}")
    logger.info(f"SSE endpoint: http://{args.host}:{args.port}/sse")
    logger.info(f"Messages endpoint: http://{args.host}:{args.port}/messages/")

    if args.host == "0.0.0.0":
        logger.warning(
            "WARNING: Server is listening on all interfaces (0.0.0.0). "
            "This allows connections from other machines on your network. "
            "Make sure you trust your network environment."
        )

    uvicorn.run(
        starlette_app,
        host=args.host,
        port=args.port,
        log_level=args.log_level.lower()
    )


if __name__ == "__main__":
    main()
