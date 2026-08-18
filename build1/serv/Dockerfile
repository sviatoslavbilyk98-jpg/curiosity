FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN touch todos.db

EXPOSE 5000

CMD ["python", "app.py"]
