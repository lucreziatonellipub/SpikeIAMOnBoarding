import json
import os
import bcrypt

USERS_FILE = os.path.join(os.path.dirname(__file__), "users.json")

def load_users() -> dict:
    if not os.path.exists(USERS_FILE):
        return {}
    with open(USERS_FILE, "r") as f:
        return json.load(f)

def verify_password(plain_password: str, hashed: str) -> bool:
 return bcrypt.checkpw(
 plain_password.encode("utf-8"),
 hashed.encode("utf-8")
 )

def create_user(username: str, plain_password: str):
    users = load_users()
    users[username] = bcrypt.hashpw(plain_password.encode(), bcrypt.gensalt()).decode()
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)