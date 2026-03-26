from database import SessionLocal, engine
from models import Base, Risultato

# Crea le tabelle se non esistono
Base.metadata.create_all(bind=engine)

db = SessionLocal()

# Inserimento record iniziali
# TODO: da fare
db.add_all([
    Risultato(valore="Primo valore"),
    Risultato(valore="Secondo valore")
])

db.commit()
db.close()

print("Record iniziali inseriti.")