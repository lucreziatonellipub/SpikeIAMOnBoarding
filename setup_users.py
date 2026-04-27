from auth import create_user

create_user("admin", "admin123")
create_user("user1", "password1")


# python setup_users.py
"""

check


python -c "
from auth import verify_password, load_users
users = load_users()
print(verify_password('admin123', users['admin']))
"
"""