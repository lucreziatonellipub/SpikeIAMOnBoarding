"""Exercise the actual app transition without database/API connections."""
import ast
import logging
from pathlib import Path
from types import SimpleNamespace
from unittest import IsolatedAsyncioTestCase
from unittest.mock import AsyncMock


class DocumentFlowTests(IsolatedAsyncioTestCase):
    async def test_import_preserves_user_answers_and_continues_with_partial_context(self):
        state = {"questions": ["Protocol?", "Environment?"], "document_text": "LDAP",
                 "answers": {"Protocol?": "User correction"}}
        session = SimpleNamespace(get=state.get, set=state.__setitem__)

        class Step:
            def __init__(self, **kwargs):
                pass

            async def __aenter__(self):
                return self

            async def __aexit__(self, *args):
                pass

        class Message:
            def __init__(self, **kwargs):
                pass

            async def send(self):
                pass

        def make_async(function):
            async def run(*args):
                return function(*args)
            return run

        next_question = AsyncMock()
        namespace = {
            "cl": SimpleNamespace(user_session=session, Step=Step, Message=Message, make_async=make_async),
            "LOGGER": logging.getLogger("test.document_flow"),
            "t": lambda it, en: en, "step_name": lambda key: key, "call_azure_llm": None,
            "understand_document": lambda *args: (
                {"Protocol?": "LDAP"}, {"Environment?": {"answer": "Production", "missing": "Test"}},
                {"Protocol?": "LDAP"}),
            "ask_next_question": next_question,
        }
        tree = ast.parse(Path("app.py").read_text(encoding="utf-8"))
        function = next(node for node in tree.body if isinstance(node, ast.AsyncFunctionDef)
                        and node.name == "apply_document_answers")
        exec(compile(ast.Module(body=[function], type_ignores=[]), "app.py", "exec"), namespace)
        await namespace["apply_document_answers"]()
        self.assertEqual(state["answers"], {"Protocol?": "User correction"})
        self.assertIn("Environment?", state["document_partial"])
        self.assertEqual(state["step"], "conversational_chat")
        next_question.assert_awaited_once()

    async def test_document_context_skips_setup_fields_and_loads_questions(self):
        state = {
            "step": "language", "lang_code": None, "document_text": None,
            "document_names": [], "answers": {}, "questions": None,
        }
        session = SimpleNamespace(get=state.get, set=state.__setitem__)

        class Step:
            def __init__(self, **kwargs):
                pass

            async def __aenter__(self):
                return self

            async def __aexit__(self, *args):
                pass

        class Message:
            def __init__(self, **kwargs):
                self.kwargs = kwargs

            async def send(self):
                pass

        def make_async(function):
            async def run(*args):
                return function(*args)
            return run

        apply_answers = AsyncMock()
        namespace = {
            "cl": SimpleNamespace(user_session=session, Step=Step, Message=Message, make_async=make_async),
            "LOGGER": logging.getLogger("test.document_flow"),
            "t": lambda it, en: it if state.get("lang_code") == "it" else en,
            "step_name": lambda key: key,
            "read_document": lambda path, name: "Documento IAM",
            "understand_document_context": lambda text, llm: ({
                "language": "it", "company": "ACME", "target_system": "HR API",
                "system_type": "Generic",
            }, {"language": "Documento IAM"}),
            "call_azure_llm": None, "MAX_CHARS": 120_000,
            "set_language": lambda code: (state.__setitem__("lang_code", code), state.__setitem__("lang_name", "Italiano")),
            "load_questions_from_DB": lambda system_type: ["Protocollo?"],
            "apply_document_answers": apply_answers,
            "language_actions": lambda: [], "system_type_actions": lambda: [],
        }
        tree = ast.parse(Path("app.py").read_text(encoding="utf-8"))
        functions = [node for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
                     and node.name in {"continue_after_document_analysis", "process_document_upload"}]
        exec(compile(ast.Module(body=functions, type_ignores=[]), "app.py", "exec"), namespace)
        await namespace["process_document_upload"](SimpleNamespace(path="source", name="source.pdf"))

        self.assertEqual(state["lang_code"], "it")
        self.assertEqual(state["company"], "ACME")
        self.assertEqual(state["system"], "HR API")
        self.assertEqual(state["system_type"], "Generic")
        self.assertEqual(state["questions"], ["Protocollo?"])
        apply_answers.assert_awaited_once()
