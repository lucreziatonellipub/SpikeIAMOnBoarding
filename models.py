from sqlalchemy.orm import declarative_base
from sqlalchemy import Column, Integer, String

Base = declarative_base()

class OnboardingAnswer(Base):
    __tablename__ = "OnboardingAnswers"

    id = Column(Integer, primary_key=True, autoincrement=True)
    customer = Column(String)
    target_system = Column(String)
    question = Column(String)
    answer = Column(String)