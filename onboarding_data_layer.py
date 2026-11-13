"""Keep setup messages in memory until the questionnaire is actually started."""
from copy import deepcopy

from chainlit.data.sql_alchemy import SQLAlchemyDataLayer


class OnboardingDataLayer(SQLAlchemyDataLayer):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._active = set()
        self._drafts = {}
        self._discarded = set()

    async def update_thread(self, thread_id, name=None, user_id=None, metadata=None, tags=None):
        if thread_id in self._discarded:
            return
        if thread_id not in self._active:
            self._drafts.setdefault(thread_id, []).append(("update_thread", (), dict(
                thread_id=thread_id, name=name, user_id=user_id, metadata=deepcopy(metadata), tags=tags)))
            return
        return await super().update_thread(thread_id, name, user_id, metadata, tags)

    async def create_step(self, step_dict):
        if step_dict['threadId'] in self._discarded:
            return
        if step_dict['threadId'] not in self._active:
            self._drafts.setdefault(step_dict['threadId'], []).append(("create_step", (deepcopy(step_dict),), {}))
            return
        return await super().create_step(step_dict)

    async def update_step(self, step_dict):
        return await self.create_step(step_dict)

    async def create_element(self, element):
        if element.thread_id in self._discarded:
            return
        if element.thread_id not in self._active:
            self._drafts.setdefault(element.thread_id, []).append(("create_element", (element,), {}))
            return
        return await super().create_element(element)

    async def get_thread(self, thread_id):
        thread = await super().get_thread(thread_id)
        if thread:
            self._active.add(thread_id)
        return thread

    async def activate_thread(self, thread_id, title, user_id):
        if thread_id in self._active:
            await super().update_thread(thread_id, name=title, metadata={})
            return False
        # The parent row must exist before any buffered messages are written.
        await super().update_thread(thread_id, name=title, user_id=user_id)
        self._active.add(thread_id)
        try:
            while self._drafts.get(thread_id):
                method, args, kwargs = self._drafts[thread_id].pop(0)
                if method == "update_thread":
                    kwargs['name'] = title
                await getattr(super(), method)(*args, **kwargs)
        finally:
            if not self._drafts.get(thread_id):
                self._drafts.pop(thread_id, None)
        return True

    def discard_draft(self, thread_id):
        if thread_id not in self._active:
            self._drafts.pop(thread_id, None)
            self._discarded.add(thread_id)
