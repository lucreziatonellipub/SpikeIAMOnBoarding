"""import bcrypt
import json
import os

USERS_FILE = "users.json"

def hash_password(plain_password: str) -> str:
    return bcrypt.hashpw(plain_password.encode(), bcrypt.gensalt()).decode()

def verify_password(plain_password: str, hashed: str) -> bool:
    return bcrypt.checkpw(plain_password.encode(), hashed.encode())

def load_users() -> dict:
    if not os.path.exists(USERS_FILE):
        return {}
    with open(USERS_FILE, "r") as f:
        return json.load(f)

def create_user(username: str, plain_password: str):
    users = load_users()
    users[username] = hash_password(plain_password)
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)"""