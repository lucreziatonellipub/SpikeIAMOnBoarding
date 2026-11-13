import copy
from types import SimpleNamespace
from unittest import IsolatedAsyncioTestCase, TestCase
from unittest.mock import AsyncMock, patch

from chainlit.data.sql_alchemy import SQLAlchemyDataLayer
from chainlit.context import context_var
from cleanup_unstarted_interviews import empty_draft
from interview_state import interview_started, interview_title
from onboarding_data_layer import OnboardingDataLayer


class InterviewStartTests(TestCase):
    def test_all_setup_fields_are_required_and_used_in_title(self):
        state = dict(company=' ACME ', system='HR', system_type=' LDAP ')
        self.assertEqual(interview_title(state), 'ACME · HR · LDAP')
        for key in state:
            draft = {**state, key: ' '}
            self.assertFalse(interview_started(draft))
            self.assertIsNone(interview_title(draft))

    def test_cleanup_preserves_started_zero_percent_and_partial_documents(self):
        checkpoint = SimpleNamespace(progress=0)
        state = dict(company='ACME', system='HR', system_type='LDAP', step='choose_method')
        self.assertFalse(empty_draft(checkpoint, state))
        self.assertTrue(empty_draft(checkpoint, {'step': 'language'}))
        for key in ('document_text', 'answers', 'document_names', 'other_infos'):
            self.assertFalse(empty_draft(checkpoint, {'step': 'language', key: 'saved content'}))


class DraftDataLayerTests(IsolatedAsyncioTestCase):
    def setUp(self):
        self.layer = OnboardingDataLayer.__new__(OnboardingDataLayer)
        self.layer._active = set()
        self.layer._drafts = {}
        self.layer._discarded = set()

    async def test_no_database_writes_before_activation_then_setup_history_is_flushed(self):
        with patch.object(SQLAlchemyDataLayer, 'update_thread', new_callable=AsyncMock) as thread_write, \
             patch.object(SQLAlchemyDataLayer, 'create_step', new_callable=AsyncMock) as step_write:
            step = {'id': 'message', 'threadId': 'draft', 'output': 'Company?'}
            await self.layer.update_thread('draft', name='choose_language', user_id='user')
            await self.layer.create_step(step)
            thread_write.assert_not_awaited()
            step_write.assert_not_awaited()
            step['output'] = 'mutated'
            self.assertTrue(await self.layer.activate_thread('draft', 'ACME · HR · LDAP', 'user'))
            step_write.assert_awaited_once_with({'id': 'message', 'threadId': 'draft', 'output': 'Company?'})
            self.assertTrue(all(call.kwargs.get('name') == 'ACME · HR · LDAP' for call in thread_write.await_args_list))
            self.assertNotIn('draft', self.layer._drafts)
            await self.layer.create_step({'id': 'new', 'threadId': 'draft'})
            self.assertEqual(step_write.await_count, 2)

    async def test_abandoning_setup_discards_it_including_disconnect_metadata(self):
        with patch.object(SQLAlchemyDataLayer, 'update_thread', new_callable=AsyncMock) as write:
            await self.layer.update_thread('draft', name='choose_language')
            self.layer.discard_draft('draft')
            await self.layer.update_thread('draft', metadata={'step': 'language'})
            write.assert_not_awaited()
            self.assertNotIn('draft', self.layer._drafts)

    async def test_resuming_existing_chat_allows_normal_persistence(self):
        with patch.object(SQLAlchemyDataLayer, 'get_thread', new_callable=AsyncMock, return_value={'id': 'existing'}), \
             patch.object(SQLAlchemyDataLayer, 'update_thread', new_callable=AsyncMock) as write:
            await self.layer.get_thread('existing')
            await self.layer.update_thread('existing', metadata={'step': 'conversational_chat'})
            write.assert_awaited_once()

    async def test_activation_runs_native_sql_methods_without_recursion_or_draft_writes(self):
        self.layer.show_logger = False
        async def execute(query, parameters):
            if query.startswith('SELECT identifier'):
                return [{'identifier': 'alice'}]
            if query.lstrip().startswith('SELECT'):
                return []
            return 1
        self.layer.execute_sql = AsyncMock(side_effect=execute)
        token = context_var.set(SimpleNamespace(session=SimpleNamespace(has_first_interaction=True)))
        try:
            await self.layer.update_thread('new', name='choose_language', user_id='user-id')
            await self.layer.create_step({'id': 'message', 'threadId': 'new', 'name': 'User', 'type': 'user_message', 'output': 'HR'})
            self.layer.execute_sql.assert_not_awaited()
            await self.layer.activate_thread('new', 'ACME · HR · LDAP', 'user-id')
            writes = [call.kwargs['query'] for call in self.layer.execute_sql.await_args_list
                      if call.kwargs.get('query', '').lstrip().startswith('INSERT')]
            self.assertTrue(any('INSERT INTO threads' in sql for sql in writes))
            self.assertTrue(any('INSERT INTO steps' in sql for sql in writes))
        finally:
            context_var.reset(token)
