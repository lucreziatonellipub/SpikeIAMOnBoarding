from sqlalchemy import Column, Integer, String, JSON, DateTime
from sqlalchemy.ext.declarative import declarative_base
from datetime import datetime

Base = declarative_base()

class OnboardingSession(Base):
    __tablename__ = "onboarding_session"

    id = Column(Integer, primary_key=True, index=True)
    company = Column(String, index=True)
    target_system = Column(String)
    system_type = Column(String)
    collected_data_original = Column(JSON) # Salva dizionari Python come JSON
    collected_data_english = Column(JSON)
    created_at = Column(DateTime, default=datetime.utcnow)

class Question(Base):
    __tablename__ = "question"

    id = Column(Integer, primary_key=True, index=True)
    question = Column(String, index=True)
    system_type = Column(String)