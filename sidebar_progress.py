"""Persisted questionnaire progress for Chainlit's chat history sidebar."""

import asyncio
import json
from typing import Annotated

from chainlit.auth import get_current_user
from chainlit.config import config
from chainlit.user import User
from fastapi import Depends, HTTPException
from pydantic import BaseModel, Field
from sqlalchemy import select
from starlette.responses import StreamingResponse

from database import get_db
from models import OnboardingChatCheckpoint
from interview_state import interview_title

_subscribers: dict[str, set[asyncio.Queue]] = {}


def build_chat_title(state):
    """Use the entered business names instead of Chainlit's first action name."""
    return interview_title(state)


def sidebar_details(state, progress):
    return {"progress": progress, "name": build_chat_title(state), "lang": state.get("lang_code")}


def broadcast_sidebar_progress(owner, payload):
    """Deliver updates directly to this user's open browsers, including iframes."""
    for queue in tuple(_subscribers.get(owner, ())):
        if queue.full():
            queue.get_nowait()
        queue.put_nowait(payload)


async def sidebar_events(owner):
    queue = asyncio.Queue(maxsize=100)
    listeners = _subscribers.setdefault(owner, set())
    listeners.add(queue)
    try:
        yield ": connected\n\n"
        while True:
            try:
                payload = await asyncio.wait_for(queue.get(), timeout=15)
                yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"
            except asyncio.TimeoutError:
                yield ": keep-alive\n\n"
    finally:
        listeners.discard(queue)
        if not listeners:
            _subscribers.pop(owner, None)


async def stream_sidebar_progress(
    current_user: Annotated[User | None, Depends(get_current_user)],
):
    if current_user is None:
        raise HTTPException(status_code=401, detail="Unauthorized")
    return StreamingResponse(
        sidebar_events(current_user.identifier),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


class ProgressRequest(BaseModel):
    thread_ids: list[str] = Field(max_length=1000)


def get_sidebar_progress(
    payload: ProgressRequest,
    current_user: Annotated[User | None, Depends(get_current_user)],
):
    """Read only the requested checkpoints belonging to the logged-in user."""
    if current_user is None:
        raise HTTPException(status_code=401, detail="Unauthorized")

    with get_db() as db:
        rows = db.execute(
            select(
                OnboardingChatCheckpoint.thread_id,
                OnboardingChatCheckpoint.progress,
                OnboardingChatCheckpoint.state_json,
            )
            .where(
                OnboardingChatCheckpoint.owner == current_user.identifier,
                OnboardingChatCheckpoint.thread_id.in_(payload.thread_ids),
            )
        ).all()
    return {
        thread_id: sidebar_details(json.loads(state_json), progress)
        for thread_id, progress, state_json in rows
    }


def install_sidebar_progress():
    # Keep the assets configured in tracked code: .chainlit/ is gitignored.
    from chainlit.server import app

    config.ui.custom_js = "/public/sidebar-progress.js?v=3"
    config.ui.custom_css = "/public/sidebar-progress.css?v=3"
    if not any(route.name == "onboarding_sidebar_progress" for route in app.routes):
        app.add_api_route(
            f"{config.run.root_path}/onboarding/progress",
            get_sidebar_progress,
            methods=["POST"],
            name="onboarding_sidebar_progress",
        )
    if not any(route.name == "onboarding_sidebar_events" for route in app.routes):
        app.add_api_route(
            f"{config.run.root_path}/onboarding/progress/events",
            stream_sidebar_progress,
            methods=["GET"],
            name="onboarding_sidebar_events",
        )
        # GET endpoints must precede Chainlit's catch-all route serving the SPA.
        app.router.routes.insert(0, app.router.routes.pop())
