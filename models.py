# SQLAlchemy models for onboarding data and PostgreSQL-backed Chainlit entities.

from datetime import datetime

from sqlalchemy import (
    ARRAY,
    JSON,
    Boolean,
    Column,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    String,
    Text,
    text,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import declarative_base


Base = declarative_base()


class OnboardingSession(Base):
    __tablename__ = "onboarding_sessions"

    id = Column(Integer, primary_key=True, index=True)
    company = Column(String, index=True)
    target_system = Column(String)
    system_type = Column(String)
    collected_data_original = Column(JSON)
    collected_data_english = Column(JSON)
    created_at = Column(DateTime, default=datetime.utcnow)

class OnboardingChatCheckpoint(Base):
    __tablename__ = "onboarding_chat_checkpoint"

    thread_id = Column(String(255), primary_key=True)
    owner = Column(String(255), nullable=False)
    # Contiene il JSON serializzato con json.dumps (come prima: colonna TEXT)
    state_json = Column(Text, nullable=False)
    progress = Column(
        Integer,
        nullable=False,
        default=0,
        server_default=text("0"),
    )


class Question(Base):
    __tablename__ = "question"

    id = Column(Integer, primary_key=True, index=True)
    question = Column(String, index=True)
    system_type = Column(String)


class User(Base):
    __tablename__ = "users"

    id = Column(Text, primary_key=True)
    identifier = Column(Text, nullable=False, unique=True)
    metadata_ = Column(
        "metadata",
        JSONB,
        nullable=False,
        server_default=text("'{}'::jsonb"),
    )
    created_at = Column("createdAt", Text, nullable=False)


class Thread(Base):
    __tablename__ = "threads"
    __table_args__ = (
        Index("chainlit_threads_user_idx", "userId"),
    )

    id = Column(Text, primary_key=True)
    created_at = Column("createdAt", Text, nullable=True)
    name = Column(Text, nullable=True)
    user_id = Column(
        "userId",
        Text,
        ForeignKey("users.id", ondelete="CASCADE"),
        nullable=True,
    )
    user_identifier = Column("userIdentifier", Text, nullable=True)
    tags = Column(ARRAY(Text), nullable=True)
    metadata_ = Column(
        "metadata",
        JSONB,
        nullable=False,
        server_default=text("'{}'::jsonb"),
    )


class Step(Base):
    __tablename__ = "steps"
    __table_args__ = (
        Index(
            "chainlit_steps_thread_created_idx",
            "threadId",
            "createdAt",
        ),
    )

    id = Column(Text, primary_key=True)
    name = Column(Text, nullable=False)
    type = Column(Text, nullable=False)
    thread_id = Column(
        "threadId",
        Text,
        ForeignKey("threads.id", ondelete="CASCADE"),
        nullable=False,
    )
    parent_id = Column("parentId", Text, nullable=True)
    streaming = Column(Boolean, nullable=True)
    wait_for_answer = Column("waitForAnswer", Boolean, nullable=True)
    is_error = Column("isError", Boolean, nullable=True)
    metadata_ = Column(
        "metadata",
        Text,
        nullable=True,
        server_default=text("'{}'"),
    )
    tags = Column(ARRAY(Text), nullable=True)
    input = Column(Text, nullable=True)
    output = Column(Text, nullable=True)
    created_at = Column("createdAt", Text, nullable=True)
    start = Column(Text, nullable=True)
    end = Column(Text, nullable=True)
    generation = Column(
        Text,
        nullable=True,
        server_default=text("'{}'"),
    )
    show_input = Column("showInput", Text, nullable=True)
    language = Column(Text, nullable=True)
    default_open = Column("defaultOpen", Boolean, nullable=True)


class Element(Base):
    __tablename__ = "elements"
    __table_args__ = (
        Index("chainlit_elements_thread_idx", "threadId"),
        Index("chainlit_elements_for_idx", "forId"),
    )

    id = Column(Text, primary_key=True)
    thread_id = Column(
        "threadId",
        Text,
        ForeignKey("threads.id", ondelete="CASCADE"),
        nullable=True,
    )
    type = Column(Text, nullable=True)
    chainlit_key = Column("chainlitKey", Text, nullable=True)
    url = Column(Text, nullable=True)
    object_key = Column("objectKey", Text, nullable=True)
    name = Column(Text, nullable=False)
    display = Column(Text, nullable=True)
    size = Column(Text, nullable=True)
    language = Column(Text, nullable=True)
    page = Column(Integer, nullable=True)
    auto_play = Column("autoPlay", Boolean, nullable=True)
    player_config = Column("playerConfig", JSONB, nullable=True)
    for_id = Column("forId", Text, nullable=True)
    mime = Column(Text, nullable=True)
    props = Column(
        Text,
        nullable=True,
        server_default=text("'{}'"),
    )


class Feedback(Base):
    __tablename__ = "feedbacks"
    __table_args__ = (
        Index("chainlit_feedbacks_for_idx", "forId"),
        Index("chainlit_feedbacks_thread_idx", "threadId"),
    )

    id = Column(Text, primary_key=True)
    for_id = Column("forId", Text, nullable=False)
    thread_id = Column(
        "threadId",
        Text,
        ForeignKey("threads.id", ondelete="CASCADE"),
        nullable=False,
    )
    value = Column(Integer, nullable=False)
    comment = Column(Text, nullable=True)