# Python small image
FROM python:3.11-slim

# Set default folder for the app
WORKDIR /app

# Copying dependencies file
COPY requirements.txt .

# Installing dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copying the rest of the app
COPY . .

# expose the port that the app will run on
EXPOSE 8000

# The command to start the app in "watch" mode and listen on all interfaces
CMD ["chainlit", "run", "app.py", "-w", "--host", "0.0.0.0", "--port", "8000"]