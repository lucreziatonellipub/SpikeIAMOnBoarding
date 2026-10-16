import json
import tempfile
import unittest
from pathlib import Path

from document_agent import read_document, understand_document, understand_document_context


class DocumentAgentTests(unittest.TestCase):
    def test_answer_contains_only_facts_without_document_framing(self):
        question = "Are there any exceptions or non-standard scenarios?"
        source = (
            "Validare porte definitive ed endpoint; eventuale uso di Global Catalog, "
            "Kerberos, SMB, RPC, WinRM ed egress 443."
        )
        response = {"items": [{
            "question_id": "0",
            "answer": (
                "Sì: il documento segnala come eccezioni o scenari non standard "
                "la necessità di validare porte definitive ed endpoint; eventuale uso "
                "di Global Catalog, Kerberos, SMB, RPC, WinRM ed egress 443. &#x20;"
            ),
            "status": "complete",
            "evidence": source,
            "missing": "",
        }]}
        complete, _, _ = understand_document(
            source, [question], lambda **kwargs: json.dumps(response)
        )
        self.assertEqual(
            complete[question],
            "La necessità di validare porte definitive ed endpoint; eventuale uso "
            "di Global Catalog, Kerberos, SMB, RPC, WinRM ed egress 443.",
        )

    def test_question_ids_map_answers_without_exact_question_echo(self):
        response = {"items": [{
            "question_id": "0",
            "question": "A rephrased question that must be ignored",
            "answer": "REST API",
            "status": "complete",
            "evidence": "Integration through REST API",
            "missing": "",
        }]}
        complete, partial, evidence = understand_document(
            "Integration through REST API",
            ["Which integration interface is available?"],
            lambda **kwargs: json.dumps(response),
        )
        self.assertEqual(complete, {"Which integration interface is available?": "REST API"})
        self.assertEqual(partial, {})
        self.assertIn("Which integration interface is available?", evidence)

    def test_evidence_wrapped_in_quotes_is_still_grounded(self):
        response = {"items": [{
            "question_id": "0", "answer": "REST API", "status": "complete",
            "evidence": '"The preferred interface is a REST API."', "missing": "",
        }]}
        complete, partial, _ = understand_document(
            "The preferred interface is a REST API.", ["Which interface is preferred?"],
            lambda **kwargs: json.dumps(response),
        )
        self.assertEqual(complete, {"Which interface is preferred?": "REST API"})
        self.assertEqual(partial, {})

    def test_multiple_evidence_excerpts_are_supported(self):
        response = {"items": [{
            "question_id": "0", "answer": "DEV and PROD are in scope",
            "status": "complete", "evidence": ["DEV exists.", "PROD is in scope."],
            "missing": "",
        }]}
        complete, _, _ = understand_document(
            "DEV exists. Additional text. PROD is in scope.", ["Which environments are in scope?"],
            lambda **kwargs: json.dumps(response),
        )
        self.assertIn("Which environments are in scope?", complete)

    def test_context_extracts_only_allowed_evidence_grounded_values(self):
        source = "Società: ACME. Sistema: People API. Integrazione via REST API."
        response = {
            "language": {"value": "it", "evidence": "Società"},
            "company": {"value": "ACME", "evidence": "Società: ACME"},
            "target_system": {"value": "People API", "evidence": "Sistema: People API"},
            "system_type": {"value": "Generic", "evidence": "via REST API"},
        }
        context, evidence = understand_document_context(
            source, lambda **kwargs: json.dumps(response)
        )
        self.assertEqual(context, {
            "language": "it", "company": "ACME",
            "target_system": "People API", "system_type": "Generic",
        })
        self.assertEqual(set(evidence), set(context))

    def test_context_rejects_unsupported_type_and_invented_evidence(self):
        response = {
            "language": {"value": "fr", "evidence": "bonjour"},
            "company": {"value": "Invented", "evidence": "not in source"},
            "target_system": {"value": None, "evidence": ""},
            "system_type": {"value": "Web App", "evidence": "API"},
        }
        context, evidence = understand_document_context(
            "bonjour API", lambda **kwargs: json.dumps(response)
        )
        self.assertEqual(context, {})
        self.assertEqual(evidence, {})

    def test_only_supported_complete_answers_close_questions(self):
        items = [
            {"question": "Protocol?", "answer": "LDAP", "status": "complete", "evidence": "LDAP"},
            {"question": "Environments?", "answer": "Production", "status": "partial", "evidence": "Production", "missing": "Test environment"},
            {"question": "Unknown?", "answer": "Yes", "status": "complete", "evidence": "LDAP"},
            {"question": "Internet?", "answer": "Yes", "status": "complete", "evidence": "Invented"},
        ]
        complete, partial, evidence = understand_document(
            "LDAP in Production", ["Protocol?", "Environments?", "Internet?"],
            lambda **kwargs: json.dumps({"items": items}),
        )
        self.assertEqual(complete, {"Protocol?": "LDAP"})
        self.assertEqual(partial["Environments?"]["missing"], "Test environment")
        self.assertEqual(set(evidence), {"Protocol?", "Environments?"})

    def test_agent_errors_are_not_treated_as_empty_success(self):
        for raw in ['{"status":"error"}', 'invalid', '{"items":{}}']:
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                understand_document("data", ["Question"], lambda **kwargs: raw)

    def test_read_text_and_reject_empty_unsupported_and_oversized(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.txt"
            path.write_text("LDAP", encoding="utf-8")
            self.assertEqual(read_document(str(path), "source.txt"), "LDAP")
            with self.assertRaises(ValueError):
                read_document(str(path), "source.exe")
            path.write_text("", encoding="utf-8")
            with self.assertRaises(ValueError):
                read_document(str(path), "source.txt")
            path.write_text("x" * 120_001, encoding="utf-8")
            with self.assertRaises(ValueError):
                read_document(str(path), "source.txt")


if __name__ == "__main__":
    unittest.main()
