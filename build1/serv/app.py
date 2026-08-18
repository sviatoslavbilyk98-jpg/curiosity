import os
import time
import psycopg2
from flask import Flask, flash, redirect, render_template, request, session, url_for
from flask_session import Session
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)

app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)

DB_HOST = os.environ.get("DB_HOST", "database")
DB_NAME = os.environ.get("DB_NAME", "todos")
DB_USER = os.environ.get("DB_USER", "postgres")
DB_PASS = os.environ.get("DB_PASS", "postgres")


def get_db():
    for i in range(10):
        try:
            conn = psycopg2.connect(
                host=DB_HOST, dbname=DB_NAME, user=DB_USER, password=DB_PASS
            )
            return conn
        except psycopg2.OperationalError:
            time.sleep(2)
    raise Exception("Could not connect to database")


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            hash TEXT NOT NULL
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS todos (
            id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id),
            task TEXT NOT NULL,
            completed INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    cur.close()
    conn.close()


init_db()


def db_execute(query, params=None):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(query, params)
    if query.strip().upper().startswith("SELECT"):
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
        cur.close()
        conn.close()
        return [dict(zip(cols, row)) for row in rows]
    else:
        conn.commit()
        cur.close()
        conn.close()
        return None


def apology(message, code=400):
    return render_template("apology.html", message=message, code=code), code


def login_required(f):
    from functools import wraps

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get("user_id") is None:
            return redirect("/login")
        return f(*args, **kwargs)

    return decorated_function


@app.route("/")
@login_required
def index():
    todos = db_execute(
        "SELECT * FROM todos WHERE user_id = %s ORDER BY created_at DESC",
        (session["user_id"],),
    )
    return render_template("index.html", todos=todos)


@app.route("/add", methods=["POST"])
@login_required
def add():
    task = request.form.get("task")
    if not task:
        flash("Task cannot be empty", "danger")
        return redirect("/")
    db_execute(
        "INSERT INTO todos (user_id, task) VALUES (%s, %s)",
        (session["user_id"], task),
    )
    flash("Task added!", "success")
    return redirect("/")


@app.route("/toggle/<int:todo_id>", methods=["POST"])
@login_required
def toggle(todo_id):
    todo = db_execute(
        "SELECT * FROM todos WHERE id = %s AND user_id = %s",
        (todo_id, session["user_id"]),
    )
    if not todo:
        flash("Task not found", "danger")
        return redirect("/")
    new_status = 0 if todo[0]["completed"] else 1
    db_execute("UPDATE todos SET completed = %s WHERE id = %s", (new_status, todo_id))
    return redirect("/")


@app.route("/delete/<int:todo_id>", methods=["POST"])
@login_required
def delete(todo_id):
    db_execute(
        "DELETE FROM todos WHERE id = %s AND user_id = %s",
        (todo_id, session["user_id"]),
    )
    flash("Task deleted", "success")
    return redirect("/")


@app.route("/login", methods=["GET", "POST"])
def login():
    session.clear()
    if request.method == "POST":
        if not request.form.get("username"):
            return apology("Username is required", 400)
        if not request.form.get("password"):
            return apology("Password is required", 400)

        rows = db_execute(
            "SELECT * FROM users WHERE username = %s",
            (request.form.get("username"),),
        )
        if len(rows) != 1 or not check_password_hash(
            rows[0]["hash"], request.form.get("password")
        ):
            return apology("Invalid username or password", 403)

        session["user_id"] = rows[0]["id"]
        return redirect("/")
    else:
        return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    session.clear()
    if request.method == "POST":
        if not request.form.get("username"):
            return apology("Username is required", 400)
        if not request.form.get("password"):
            return apology("Password is required", 400)
        if not request.form.get("confirmation"):
            return apology("You must confirm your password", 400)
        if request.form.get("password") != request.form.get("confirmation"):
            return apology("Passwords do not match", 400)

        rows = db_execute(
            "SELECT * FROM users WHERE username = %s",
            (request.form.get("username"),),
        )
        if len(rows) != 0:
            return apology("Username already exists", 400)

        db_execute(
            "INSERT INTO users (username, hash) VALUES (%s, %s)",
            (
                request.form.get("username"),
                generate_password_hash(request.form.get("password")),
            ),
        )

        rows = db_execute(
            "SELECT * FROM users WHERE username = %s",
            (request.form.get("username"),),
        )
        session["user_id"] = rows[0]["id"]
        return redirect("/")
    else:
        return render_template("register.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
