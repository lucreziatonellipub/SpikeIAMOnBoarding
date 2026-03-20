from database import engine, SessionLocal
from models import Base, OnboardingAnswer

Base.metadata.create_all(bind=engine)

def main():
    db = SessionLocal()

    nuovo = OnboardingAnswer(customer="testCustomer", target_system="AD", question="testQuestion")
    db.add(nuovo)
    db.commit()
    db.close()

    risultati = db.query(OnboardingAnswer).all()

    for r in risultati:
        print(r.id, r.customer, r.target_system, r.question)

    db.close()

if __name__ == "__main__":
    main()