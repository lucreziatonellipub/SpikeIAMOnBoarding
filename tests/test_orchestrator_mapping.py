import unittest

from app import normalize_orchestrator_answers


class OrchestratorMappingTests(unittest.TestCase):
    def setUp(self):
        self.questions = ["Protocol?", "Environments?"]
        self.catalog = {"P0": "Protocol?", "P1": "Environments?"}

    def test_current_id_schema_maps_answers(self):
        mapped, other = normalize_orchestrator_answers(
            [{"question_id": "P0", "answer": "LDAPS"}],
            self.catalog,
            self.questions,
            ["Port 636"],
        )
        self.assertEqual(mapped, {"Protocol?": "LDAPS"})
        self.assertEqual(other, ["Port 636"])

    def test_legacy_question_answer_schema_still_maps(self):
        mapped, other = normalize_orchestrator_answers(
            {
                "question": "Environments?",
                "answer": "DEV and PROD",
                "other_infos": ["Two instances"],
            },
            self.catalog,
            self.questions,
        )
        self.assertEqual(mapped, {"Environments?": "DEV and PROD"})
        self.assertEqual(other, ["Two instances"])

    def test_legacy_question_keyed_schema_still_maps(self):
        mapped, _ = normalize_orchestrator_answers(
            {"Protocol?": "LDAPv3"}, self.catalog, self.questions
        )
        self.assertEqual(mapped, {"Protocol?": "LDAPv3"})

    def test_unknown_ids_are_not_mapped(self):
        mapped, _ = normalize_orchestrator_answers(
            [{"question_id": "P99", "answer": "value"}],
            self.catalog,
            self.questions,
        )
        self.assertEqual(mapped, {})


if __name__ == "__main__":
    unittest.main()
