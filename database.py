
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

password = "NZDHJpm9EufinDF2A0Yp"

DATABASE_URL = "postgresql://postgres:" + password + "@localhost:5432/iam_onboarding_db"

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)