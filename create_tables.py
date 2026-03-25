from database import engine
from models import Base

# Crea le tabelle nel DB basandosi sui modelli definiti
Base.metadata.create_all(bind=engine)

print("✅ Creation of tables successful.")