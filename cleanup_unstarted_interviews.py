"""Preview or remove abandoned setup records; back up every affected row first."""
import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from sqlalchemy import MetaData, Table, delete, inspect, select, update

from database import get_db
from interview_state import interview_started, interview_title
from models import OnboardingChatCheckpoint


def empty_draft(checkpoint, state):
    return (
        not interview_started(state)
        and checkpoint.progress == 0
        and state.get('step') in ('language', 'company', 'system', 'system_type')
        and not any(state.get(key) for key in (
            'answers', 'document_text', 'document_names', 'document_evidence',
            'document_partial', 'document_context_evidence', 'other_infos', 'other_conversation',
        ))
    )


def cleanup(apply=False):
    with get_db() as db:
        query = select(OnboardingChatCheckpoint)
        if apply:
            query = query.with_for_update()
        checkpoints = db.execute(query).scalars().all()
        drafts, titles = [], {}
        for checkpoint in checkpoints:
            try:
                state = json.loads(checkpoint.state_json)
                if not isinstance(state, dict):
                    continue
            except (ValueError, TypeError):
                continue
            if empty_draft(checkpoint, state):
                drafts.append(checkpoint)
            elif interview_started(state):
                titles[checkpoint.thread_id] = (checkpoint.owner, interview_title(state))

        metadata = MetaData()
        names = set(inspect(db.connection()).get_table_names())
        tables = {name: Table(name, metadata, autoload_with=db.connection())
                  for name in ('threads', 'steps', 'elements', 'feedbacks') if name in names}
        threads = tables.get('threads')
        thread_rows = {}
        if threads is not None:
            thread_rows = {row['id']: dict(row) for row in db.execute(select(threads)).mappings()}
        drafts = [cp for cp in drafts if cp.thread_id not in thread_rows
                  or thread_rows[cp.thread_id]['userIdentifier'] == cp.owner]
        titles = {tid: (owner, title) for tid, (owner, title) in titles.items()
                  if tid in thread_rows and thread_rows[tid]['userIdentifier'] == owner
                  and thread_rows[tid]['name'] != title}
        print(f"Empty setup checkpoints: {len(drafts)}; existing titles to update: {len(titles)}")
        if not apply:
            return

        ids = [cp.thread_id for cp in drafts]
        backup = {'onboarding_chat_checkpoint': [
            {column.name: getattr(cp, column.name) for column in cp.__table__.columns} for cp in drafts
        ]}
        for name, table in tables.items():
            column = table.c.id if name == 'threads' else table.c.threadId
            affected_ids = ids + list(titles) if name == 'threads' else ids
            backup[name] = [dict(row) for row in db.execute(select(table).where(column.in_(affected_ids))).mappings()]
        directory = Path(__file__).resolve().parent / 'backups'
        directory.mkdir(exist_ok=True)
        path = directory / (datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ') + '-unstarted-interviews.json')
        with path.open('x', encoding='utf-8') as file:
            json.dump(backup, file, ensure_ascii=False, indent=2, default=str)

        for name in ('feedbacks', 'elements', 'steps', 'threads'):
            if name in tables:
                table = tables[name]
                column = table.c.id if name == 'threads' else table.c.threadId
                db.execute(delete(table).where(column.in_(ids)))
        db.execute(delete(OnboardingChatCheckpoint).where(OnboardingChatCheckpoint.thread_id.in_(ids)))
        for tid, (owner, title) in titles.items():
            db.execute(update(threads).where(threads.c.id == tid, threads.c.userIdentifier == owner).values(name=title))
        db.commit()
        print(f"Removed {len(ids)} empty checkpoints and their chat records; updated {len(titles)} titles.")
        print(f"Backup: {path}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--apply', action='store_true', help='Apply changes after writing a local backup')
    cleanup(parser.parse_args().apply)
