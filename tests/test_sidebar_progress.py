"""Verify progress ownership and notifications at checkpoint boundaries."""

import ast
import asyncio
import copy
import logging
import json
from contextlib import contextmanager
from functools import wraps
from pathlib import Path
from types import SimpleNamespace
from unittest import IsolatedAsyncioTestCase, TestCase
from unittest.mock import AsyncMock, patch

from fastapi import FastAPI
from fastapi.testclient import TestClient
from sqlalchemy import create_engine, insert, select, text, update
from sqlalchemy.orm import Session
from sqlalchemy.pool import StaticPool

from models import OnboardingChatCheckpoint, Thread
from interview_state import interview_started
from sidebar_progress import (
    _subscribers, broadcast_sidebar_progress, build_chat_title, get_current_user,
    get_sidebar_progress, sidebar_details, sidebar_events, stream_sidebar_progress,
)


class SidebarProgressAPITests(TestCase):
    def setUp(self):
        self.engine = create_engine(
            "sqlite://", connect_args={"check_same_thread": False}, poolclass=StaticPool,
        )
        OnboardingChatCheckpoint.__table__.create(self.engine)
        with Session(self.engine) as db:
            db.add_all([
                OnboardingChatCheckpoint(thread_id="mine", owner="alice", state_json='{"company":"ACME","system":"HR","system_type":"LDAP","lang_code":"it"}', progress=42),
                OnboardingChatCheckpoint(thread_id="other", owner="bob", state_json="{}", progress=80),
                OnboardingChatCheckpoint(thread_id="not-requested", owner="alice", state_json="{}", progress=100),
            ])
            db.commit()

        @contextmanager
        def test_db():
            with Session(self.engine) as db:
                yield db

        self.db_patch = patch("sidebar_progress.get_db", test_db)
        self.db_patch.start()
        self.addCleanup(self.db_patch.stop)
        self.addCleanup(self.engine.dispose)
        self.app = FastAPI()
        self.app.add_api_route("/onboarding/progress", get_sidebar_progress, methods=["POST"])
        self.app.add_api_route("/onboarding/progress/events", stream_sidebar_progress, methods=["GET"])
        self.app.dependency_overrides[get_current_user] = lambda: SimpleNamespace(identifier="alice")
        self.client = TestClient(self.app)
        self.addCleanup(self.client.close)

    def test_reads_saved_progress_without_exposing_other_users(self):
        response = self.client.post("/onboarding/progress", json={"thread_ids": ["mine", "other", "missing"]})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {"mine": {"progress": 42, "name": "ACME · HR · LDAP", "lang": "it"}})

    def test_requires_authenticated_user(self):
        self.app.dependency_overrides[get_current_user] = lambda: None
        response = self.client.post("/onboarding/progress", json={"thread_ids": ["mine"]})
        self.assertEqual(response.status_code, 401)
        self.assertEqual(self.client.get("/onboarding/progress/events").status_code, 401)

    def test_empty_and_oversized_requests(self):
        self.assertEqual(self.client.post("/onboarding/progress", json={"thread_ids": []}).json(), {})
        response = self.client.post("/onboarding/progress", json={"thread_ids": ["mine"] * 1001})
        self.assertEqual(response.status_code, 422)


class ChatTitleTests(TestCase):
    def test_title_uses_entered_names_and_normalizes_whitespace(self):
        self.assertEqual(build_chat_title({"company": " ACME  Italia ", "system": " HR\nAPI ", "system_type": "LDAP"}), "ACME Italia · HR API · LDAP")
        self.assertIsNone(build_chat_title({"company": "ACME"}))
        self.assertIsNone(build_chat_title({}))


class SidebarStreamTests(IsolatedAsyncioTestCase):
    async def test_updates_reach_all_owner_tabs_and_cleanup_on_disconnect(self):
        alice_tabs = [sidebar_events("alice"), sidebar_events("alice")]
        bob = sidebar_events("bob")
        streams = [*alice_tabs, bob]
        try:
            for stream in streams:
                self.assertEqual(await anext(stream), ": connected\n\n")
            payload = {"thread_id": "mine", "progress": 50, "name": "ACME · HR"}
            broadcast_sidebar_progress("alice", payload)
            for stream in alice_tabs:
                event = await asyncio.wait_for(anext(stream), timeout=1)
                self.assertEqual(json.loads(event.removeprefix("data: ")), payload)
            self.assertTrue(all(queue.empty() for queue in _subscribers["bob"]))
        finally:
            for stream in streams:
                await stream.aclose()
        self.assertNotIn("alice", _subscribers)
        self.assertNotIn("bob", _subscribers)


class CheckpointPersistenceTests(TestCase):
    def test_checkpoint_never_locks_chainlit_threads_table(self):
        engine = create_engine("sqlite://")
        self.addCleanup(engine.dispose)
        OnboardingChatCheckpoint.__table__.create(engine)
        with engine.begin() as connection:
            connection.execute(text('CREATE TABLE threads (id TEXT PRIMARY KEY, name TEXT, "userIdentifier" TEXT)'))
            connection.execute(text("INSERT INTO threads VALUES ('mine', 'choose_language', 'alice')"))
        state = {"company": "ACME", "system": "HR", "system_type": "LDAP", "questions": ["Q1", "Q2"], "answers": {"Q1": "Yes"}}
        namespace = {
            "json": json, "insert": insert, "select": select, "update": update,
            "OnboardingChatCheckpoint": OnboardingChatCheckpoint, "Thread": Thread,
            "session_identity": lambda: ("mine", "alice"), "build_chat_title": build_chat_title, "interview_started": interview_started,
            "cl": SimpleNamespace(user_session=SimpleNamespace(set=state.__setitem__)),
        }
        tree = ast.parse(Path("app.py").read_text(encoding="utf-8"))
        functions = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in {
            "write_checkpoint", "synchronize_state_progress", "calculate_questionnaire_progress", "is_non_empty_answer",
        }]
        exec(compile(ast.Module(body=functions, type_ignores=[]), "app.py", "exec"), namespace)
        with Session(engine) as db:
            namespace["write_checkpoint"](db, {"company": "ACME", "system": "HR"})
            self.assertIsNone(db.get(OnboardingChatCheckpoint, "mine"))
            namespace["write_checkpoint"](db, state)
            db.commit()
            self.assertEqual(db.get(OnboardingChatCheckpoint, "mine").progress, 50)
            self.assertEqual(db.execute(select(Thread.name)).scalar_one(), "choose_language")
            state["system"] = "LDAP"
            namespace["write_checkpoint"](db, state)
            db.rollback()
            self.assertEqual(db.execute(select(Thread.name)).scalar_one(), "choose_language")
            self.assertEqual(json.loads(db.get(OnboardingChatCheckpoint, "mine").state_json)["system"], "HR")


class CheckpointProgressTests(IsolatedAsyncioTestCase):
    def setUp(self):
        self.state = {"step": "conversational_chat", "progress": 25}
        self.saved = []
        self.notifications = []
        self.runtime = {"lock": asyncio.Lock(), "task": None}
        self.layer = SimpleNamespace(update_thread=AsyncMock())

        def notify(owner, payload):
            self.assertEqual(owner, "alice")
            self.notifications.append(payload.copy())

        class Message:
            def __init__(self, **kwargs):
                pass

            async def send(self):
                pass

        namespace = {
            "asyncio": asyncio, "wraps": wraps,
            "runtime_state": lambda: self.runtime,
            "snapshot_state": lambda: copy.deepcopy(self.state),
            "restore_state": lambda snapshot: self.state.update(snapshot),
            "save_checkpoint": lambda: self.saved.append(copy.deepcopy(self.state)),
            "session_identity": lambda: ("mine", "alice"),
            "broadcast_sidebar_progress": notify, "sidebar_details": sidebar_details,
            "build_chat_title": build_chat_title, "get_data_layer": lambda: self.layer,
            "refresh_session_progress": lambda: (self.state["progress"], 1, 2),
            "format_progress": lambda *args: "Completion",
            "LOGGER": logging.getLogger("test.sidebar_progress"),
            "cl": SimpleNamespace(user_session=SimpleNamespace(get=self.state.get), Message=Message),
        }
        tree = ast.parse(Path("app.py").read_text(encoding="utf-8"))
        functions = [node for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
                     and node.name in {"checkpointed", "publish_sidebar_progress", "show_progress", "persist_chat_title"}]
        exec(compile(ast.Module(body=functions, type_ignores=[]), "app.py", "exec"), namespace)
        self.checkpointed = namespace["checkpointed"]
        self.show_progress = namespace["show_progress"]
        self.persist_title = namespace["persist_chat_title"]

    async def test_title_uses_async_data_layer_and_only_updates_when_changed(self):
        self.state.update(company="ACME", system="HR", system_type="LDAP")
        await self.persist_title()
        await self.persist_title()
        self.layer.update_thread.assert_awaited_once_with(thread_id="mine", name="ACME · HR · LDAP", metadata={})

    async def test_title_failure_does_not_break_message_or_rollback_answer(self):
        self.state.update(company="ACME", system="HR", system_type="LDAP")
        self.layer.update_thread.side_effect = RuntimeError("title storage unavailable")

        @self.checkpointed
        async def answer():
            self.state["progress"] = 50
            return "sent"

        with self.assertLogs("test.sidebar_progress", level="WARNING"):
            self.assertEqual(await answer(), "sent")
        self.assertEqual(self.saved[-1]["progress"], 50)
        self.assertEqual(self.state["progress"], 50)

    async def test_each_answer_notifies_after_save_including_completion(self):
        @self.checkpointed
        async def answer(progress):
            self.state["progress"] = progress
            return "answered"

        for progress in (50, 75, 100):
            self.assertEqual(await answer(progress), "answered")
        self.assertEqual([item["progress"] for item in self.notifications], [50, 75, 100])
        self.assertTrue(all(item["thread_id"] == "mine" for item in self.notifications))
        self.assertEqual([item["progress"] for item in self.saved], [50, 75, 100])

    async def test_progress_changes_before_the_next_question_finishes(self):
        @self.checkpointed
        async def answer():
            self.state["progress"] = 50
            await self.show_progress()
            self.assertEqual(self.notifications[-1]["progress"], 50)
            self.assertEqual(self.saved, [])
            # The subsequent LLM call has not completed or committed yet.
            await asyncio.sleep(0)

        await answer()
        self.assertEqual(self.saved[-1]["progress"], 50)

    async def test_failed_or_cancelled_answer_notifies_restored_progress(self):
        for error in (ValueError, asyncio.CancelledError):
            @self.checkpointed
            async def interrupted_answer():
                self.state["progress"] = 75
                raise error()

            with self.assertRaises(error):
                await interrupted_answer()
            self.assertEqual(self.notifications[-1]["progress"], 25)
            self.assertIsNone(self.runtime["task"])

