"""Admin interview lists use persisted checkpoint progress and phase."""
import json
from contextlib import contextmanager
from unittest import TestCase
from unittest.mock import patch

from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy.pool import StaticPool

import admin_dashboard as dashboard
from models import OnboardingChatCheckpoint, OnboardingSession


class AdminInterviewTests(TestCase):
    def setUp(self):
        self.engine = create_engine('sqlite://', connect_args={'check_same_thread': False}, poolclass=StaticPool)
        self.addCleanup(self.engine.dispose)
        for table in (OnboardingChatCheckpoint.__table__, OnboardingSession.__table__):
            table.create(self.engine)
        with Session(self.engine) as db:
            for thread, owner, step, progress in (
                ('active', 'alice', 'conversational_chat', 37),
                ('saving', 'bob', 'finalizing', 100),
                ('done', 'alice', 'completed', 100),
                ('new', 'carol', 'language', 0),
            ):
                state = {'company': 'ACME', 'system': 'HR', 'system_type': 'LDAP' if thread != 'new' else None, 'step': step,
                         'progress': 99, 'answers': {'Protocol?': 'LDAP'}}
                db.add(OnboardingChatCheckpoint(thread_id=thread, owner=owner, progress=progress, state_json=json.dumps(state)))
            db.add(OnboardingSession(company='ACME', target_system='HR', system_type='LDAP'))
            db.commit()

        @contextmanager
        def test_db():
            with Session(self.engine) as db:
                yield db

        patcher = patch.object(dashboard, 'get_db', test_db)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.client = dashboard.app.test_client()
        with self.client.session_transaction() as session:
            session['authenticated'] = True

    def test_lists_active_checkpoints_from_all_owners_with_database_progress(self):
        response = self.client.get('/api/interviews/in-progress')
        self.assertEqual(response.status_code, 200)
        rows = {row['thread_id']: row for row in response.json}
        self.assertEqual(set(rows), {'active', 'saving'})
        self.assertEqual(rows['active']['progress'], 37)
        self.assertEqual(rows['active']['status_label'], 'Questionnaire in progress')
        self.assertEqual(rows['saving']['step'], 'finalizing')
        self.assertEqual(rows['saving']['progress'], 100)
        self.assertNotIn('new', rows)
        self.assertNotIn('collected_data_original', rows['active'])

    def test_completed_results_remain_separate_without_duplicate_checkpoints(self):
        response = self.client.get('/api/sessions')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(len(response.json), 1)
        self.assertEqual(response.json[0]['company'], 'ACME')

    def test_active_details_show_saved_answers_status_and_progress(self):
        with Session(self.engine) as db:
            checkpoint = db.get(OnboardingChatCheckpoint, 'active')
            state = json.loads(checkpoint.state_json)
            state.update(questions=['Protocol?', 'Environments?'], current_asked_question='Environments?',
                         document_partial={'Environments?': {'answer': 'PROD', 'missing': 'Test environment'}},
                         document_names=['requirements.pdf'])
            checkpoint.state_json = json.dumps(state)
            db.commit()
        response = self.client.get('/api/interviews/checkpoints/active')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json['progress'], 37)
        self.assertEqual(response.json['collected_data_original'], {'Protocol?': 'LDAP'})
        self.assertEqual(response.json['questions'], ['Protocol?', 'Environments?'])
        self.assertEqual(response.json['current_question'], 'Environments?')
        self.assertEqual(response.json['document_partial']['Environments?']['missing'], 'Test environment')
        self.assertEqual(response.json['document_names'], ['requirements.pdf'])
        self.assertEqual(self.client.get('/api/interviews/checkpoints/missing').status_code, 404)

    def test_page_has_both_sections_and_admin_authentication_is_required(self):
        response = self.client.get('/interviews')
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'Completed interviews', response.data)
        self.assertIn(b'id="inProgressTable"', response.data)
        anonymous = dashboard.app.test_client()
        for url in ('/api/interviews/in-progress', '/api/interviews/checkpoints/active'):
            self.assertEqual(anonymous.get(url).status_code, 302)

    def test_database_progress_is_refreshed_on_the_next_request(self):
        with Session(self.engine) as db:
            checkpoint = db.get(OnboardingChatCheckpoint, 'active')
            checkpoint.progress = 64
            db.commit()
        rows = self.client.get('/api/interviews/in-progress').json
        self.assertEqual(next(row for row in rows if row['thread_id'] == 'active')['progress'], 64)

    def test_pagination_counts_and_page_changes_for_both_sections(self):
        with Session(self.engine) as db:
            for index in range(25):
                db.add(OnboardingSession(company=f'Company {index}', target_system='HR', system_type='LDAP'))
                db.add(OnboardingChatCheckpoint(
                    thread_id=f'page-{index:02}', owner='alice', progress=0,
                    state_json=json.dumps(dict(company='ACME', system='HR', system_type='LDAP', step='choose_method')),
                ))
            db.commit()
        for endpoint, total in (('/api/sessions', 26), ('/api/interviews/in-progress', 27)):
            first = self.client.get(endpoint + '?page=1&page_size=10').json
            second = self.client.get(endpoint + '?page=2&page_size=10').json
            self.assertEqual(first['total'], total)
            self.assertEqual(first['pages'], 3)
            self.assertEqual(len(first['items']), 10)
            self.assertEqual(len(second['items']), 10)
            self.assertNotEqual(first['items'], second['items'])
            last = self.client.get(endpoint + '?page=99&page_size=25').json
            self.assertEqual(last['page'], 2)
            self.assertEqual(len(last['items']), total - 25)
            self.assertEqual(self.client.get(endpoint + '?page=0&page_size=10').status_code, 400)
            self.assertEqual(self.client.get(endpoint + '?page=1&page_size=999').status_code, 400)
