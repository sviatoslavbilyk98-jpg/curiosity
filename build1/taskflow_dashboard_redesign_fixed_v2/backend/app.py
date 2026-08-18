import os
import json
import time
import random
import psycopg2
from datetime import date, datetime
from flask import Flask, flash, redirect, render_template, request, session, url_for
from flask_session import Session
from werkzeug.security import check_password_hash, generate_password_hash

TRANSLATIONS_DIR = os.path.join(os.path.dirname(__file__), 'translations')
TRANSLATIONS = {}
for lang_code in ['uk', 'ru', 'en']:
    path = os.path.join(TRANSLATIONS_DIR, f'{lang_code}.json')
    with open(path, 'r', encoding='utf-8') as f:
        TRANSLATIONS[lang_code] = json.load(f)

SUPPORTED_LANGS = ['uk', 'ru', 'en']
DEFAULT_LANG = 'uk'


def get_lang():
    return session.get('lang', DEFAULT_LANG)


def t(key):
    lang = get_lang()
    return TRANSLATIONS.get(lang, TRANSLATIONS[DEFAULT_LANG]).get(key, key)

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)

@app.context_processor
def inject_translations():
    return dict(t=t, current_lang=get_lang(), supported_langs=SUPPORTED_LANGS)

DB_HOST = os.environ.get("DB_HOST", "database")
DB_NAME = os.environ.get("DB_NAME", "todos")
DB_USER = os.environ.get("DB_USER", "postgres")
DB_PASS = os.environ.get("DB_PASS", "postgres")

AVATAR_COLORS = [
    '#7c3aed', '#3b82f6', '#06b6d4', '#10b981', '#f59e0b',
    '#ef4444', '#ec4899', '#8b5cf6', '#14b8a6', '#f97316'
]

COMMUNITY_COLORS = ['#7c3aed', '#3b82f6', '#06b6d4', '#10b981', '#ef4444', '#f59e0b', '#ec4899']

COMMUNITY_NAME_KEYS = {
    'Розробники': 'community_devs',
    'Дизайнери': 'community_designers',
    'Маркетологи': 'community_marketers',
}


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
    cur.execute("""
        CREATE TABLE IF NOT EXISTS communities (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            color TEXT DEFAULT '#7c3aed'
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS community_members (
            community_id INTEGER REFERENCES communities(id),
            user_id INTEGER REFERENCES users(id),
            PRIMARY KEY (community_id, user_id)
        )
    """)
    cur.execute("SELECT COUNT(*) FROM communities")
    if cur.fetchone()[0] == 0:
        sample_communities = [
            ('Розробники', '#3b82f6'),
            ('Дизайнери', '#ec4899'),
            ('Маркетологи', '#10b981'),
        ]
        for name, color in sample_communities:
            cur.execute("INSERT INTO communities (name, color) VALUES (%s, %s)", (name, color))
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


def get_user_color(user_id):
    return AVATAR_COLORS[user_id % len(AVATAR_COLORS)]


def get_base_context():
    ctx = {}
    if session.get("user_id"):
        user = db_execute("SELECT * FROM users WHERE id = %s", (session["user_id"],))
        if user:
            u = user[0]
            u["color"] = get_user_color(u["id"])
            ctx["current_user"] = u

        all_users = db_execute("SELECT * FROM users ORDER BY username")
        for u in all_users:
            u["color"] = get_user_color(u["id"])
        ctx["all_users"] = all_users

        communities = db_execute("SELECT * FROM communities ORDER BY name")
        for c in communities:
            if not c.get("color"):
                c["color"] = COMMUNITY_COLORS[c["id"] % len(COMMUNITY_COLORS)]
            name_key = COMMUNITY_NAME_KEYS.get(c["name"])
            if name_key:
                c["display_name"] = t(name_key)
            else:
                c["display_name"] = c["name"]
        ctx["communities"] = communities

        today = date.today()
        today_start = datetime.combine(today, datetime.min.time())
        today_end = datetime.combine(today, datetime.max.time())
        today_todos = db_execute(
            "SELECT * FROM todos WHERE user_id = %s AND created_at BETWEEN %s AND %s ORDER BY created_at DESC",
            (session["user_id"], today_start, today_end),
        )
        ctx["today_todos"] = today_todos

        recent_todos = db_execute(
            "SELECT * FROM todos WHERE user_id = %s ORDER BY created_at DESC LIMIT 20",
            (session["user_id"],),
        )
        activity_icons = ['fa-circle-plus', 'fa-circle-check', 'fa-trash', 'fa-pen']
        activity_texts_templates = [
            t('activity_added'), t('activity_completed'), t('activity_deleted'), t('activity_updated')
        ]
        activities = []
        for i, todo_item in enumerate(recent_todos[:6]):
            idx = i % 4
            activities.append({
                'text': f'{activity_texts_templates[idx]}: {todo_item["task"][:30]}',
                'time': todo_item["created_at"].strftime("%H:%M") if todo_item["created_at"] else "",
                'icon': activity_icons[idx],
                'color': get_user_color(session["user_id"]) if idx in [0, 2] else '#10b981',
            })
        ctx["activities"] = activities

        task_dates_raw = db_execute(
            "SELECT DISTINCT DATE(created_at) as d FROM todos WHERE user_id = %s",
            (session["user_id"],),
        )
        ctx["task_dates"] = [r["d"].isoformat() for r in task_dates_raw if r["d"]]
    else:
        ctx["current_user"] = {"username": "", "color": "#7c3aed"}
        ctx["all_users"] = []
        ctx["communities"] = []
        ctx["activities"] = []
        ctx["today_todos"] = []
        ctx["task_dates"] = []
    return ctx


@app.route("/")
@login_required
def index():
    todos = db_execute(
        "SELECT * FROM todos WHERE user_id = %s ORDER BY created_at DESC",
        (session["user_id"],),
    )
    ctx = get_base_context()
    return render_template("index.html", todos=todos, **ctx)


@app.route("/add", methods=["POST"])
@login_required
def add():
    task = request.form.get("task")
    if not task:
        flash(t('flash_task_empty'), "danger")
        return redirect("/")
    db_execute(
        "INSERT INTO todos (user_id, task) VALUES (%s, %s)",
        (session["user_id"], task),
    )
    flash(t('flash_task_added'), "success")
    return redirect("/")


@app.route("/toggle/<int:todo_id>", methods=["POST"])
@login_required
def toggle(todo_id):
    todo = db_execute(
        "SELECT * FROM todos WHERE id = %s AND user_id = %s",
        (todo_id, session["user_id"]),
    )
    if not todo:
        flash(t('flash_task_not_found'), "danger")
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
    flash(t('flash_task_deleted'), "success")
    return redirect("/")


@app.route("/login", methods=["GET", "POST"])
def login():
    session.clear()
    if request.method == "POST":
        if not request.form.get("username"):
            return apology(t('flash_username_required'), 400)
        if not request.form.get("password"):
            return apology(t('flash_password_required'), 400)

        rows = db_execute(
            "SELECT * FROM users WHERE username = %s",
            (request.form.get("username"),),
        )
        if len(rows) != 1 or not check_password_hash(
            rows[0]["hash"], request.form.get("password")
        ):
            return apology(t('flash_invalid_credentials'), 403)

        session["user_id"] = rows[0]["id"]
        return redirect("/")
    else:
        ctx = get_base_context()
        return render_template("login.html", **ctx)


@app.route("/register", methods=["GET", "POST"])
def register():
    session.clear()
    if request.method == "POST":
        if not request.form.get("username"):
            return apology(t('flash_username_required'), 400)
        if not request.form.get("password"):
            return apology(t('flash_password_required'), 400)
        if not request.form.get("confirmation"):
            return apology(t('flash_confirm_required'), 400)
        if request.form.get("password") != request.form.get("confirmation"):
            return apology(t('flash_passwords_mismatch'), 400)

        rows = db_execute(
            "SELECT * FROM users WHERE username = %s",
            (request.form.get("username"),),
        )
        if len(rows) != 0:
            return apology(t('flash_username_exists'), 400)

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
        ctx = get_base_context()
        return render_template("register.html", **ctx)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


@app.route("/lang/<lang_code>")
def set_lang(lang_code):
    if lang_code in SUPPORTED_LANGS:
        session['lang'] = lang_code
    return redirect(request.referrer or '/')


@app.route("/activity")
@login_required
def activity():
    ctx = get_base_context()
    return render_template("activity.html", **ctx)


@app.route("/messages")
@login_required
def messages():
    ctx = get_base_context()
    return render_template("messages.html", **ctx)


@app.route("/communities")
@login_required
def communities():
    ctx = get_base_context()
    return render_template("communities.html", **ctx)


@app.route("/communities/<int:community_id>")
@login_required
def community_detail(community_id):
    ctx = get_base_context()
    community = db_execute("SELECT * FROM communities WHERE id = %s", (community_id,))
    if community:
        c = community[0]
        name_key = COMMUNITY_NAME_KEYS.get(c["name"])
        c["display_name"] = t(name_key) if name_key else c["name"]
        ctx["community"] = c
    else:
        ctx["community"] = None
    return render_template("community_detail.html", **ctx)


@app.route("/settings")
@login_required
def settings():
    ctx = get_base_context()
    return render_template("settings.html", **ctx)


@app.route("/cabinet")
@login_required
def cabinet():
    ctx = get_base_context()
    stats = {
        'total': db_execute(
            "SELECT COUNT(*) as c FROM todos WHERE user_id = %s", (session["user_id"],)
        )[0]['c'],
        'done': db_execute(
            "SELECT COUNT(*) as c FROM todos WHERE user_id = %s AND completed = 1", (session["user_id"],)
        )[0]['c'],
    }
    stats['active'] = stats['total'] - stats['done']
    ctx["cabinet_stats"] = stats
    return render_template("cabinet.html", **ctx)


@app.route("/plans")
@login_required
def plans():
    ctx = get_base_context()
    return render_template("plans.html", **ctx)


@app.route("/pay/<plan>", methods=["GET", "POST"])
@login_required
def pay(plan):
    ctx = get_base_context()
    ctx["plan"] = plan
    if request.method == "POST":
        flash(t('flash_payment_success'), "success")
        return redirect("/plans")
    return render_template("pay.html", **ctx)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
