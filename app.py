"""Step 1 - QR registration.

Participants scan one QR code, tap "Join", and receive a random unique number
from 1001-1100. Registration is open for WINDOW_SECONDS (default 300 = 5 min)
after the organizer presses Start on /admin, then closes automatically.
The organizer downloads the registered numbers as Participants.xlsx.

Run:  ADMIN_PASSWORD=yourpassword PUBLIC_URL=https://your-site.example.com python app.py
"""
import io
import os
import random
import secrets
import sqlite3
import time

import openpyxl
from flask import (Flask, abort, make_response, redirect, render_template_string,
                   request, send_file, session, url_for)

LOW, HIGH = 1001, 1100
WINDOW_SECONDS = int(os.environ.get("WINDOW_SECONDS", 300))
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "changeme")
PUBLIC_URL = os.environ.get("PUBLIC_URL", "").rstrip("/")
DB_PATH = os.environ.get("DB_PATH", os.path.join(os.path.dirname(os.path.abspath(__file__)), "registrations.db"))

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(16))


# ---------- storage ----------
def db():
    con = sqlite3.connect(DB_PATH, timeout=10, isolation_level=None)
    con.execute("CREATE TABLE IF NOT EXISTS participants "
                "(number INTEGER PRIMARY KEY, token TEXT UNIQUE NOT NULL, ts REAL NOT NULL)")
    con.execute("CREATE TABLE IF NOT EXISTS state (k TEXT PRIMARY KEY, v REAL)")
    return con


def get_window(con):
    rows = dict(con.execute("SELECT k, v FROM state").fetchall())
    return rows.get("opens_at"), rows.get("closes_at")


def status(con):
    """Returns (state, seconds_left): state is 'waiting', 'open' or 'closed'."""
    opens, closes = get_window(con)
    if opens is None:
        return "waiting", 0
    now = time.time()
    if now < closes:
        return "open", int(closes - now) + 1
    return "closed", 0


def count(con):
    return con.execute("SELECT COUNT(*) FROM participants").fetchone()[0]


def join(token):
    """Give this token a random free number. Returns (number|None, reason)."""
    con = db()
    try:
        con.execute("BEGIN IMMEDIATE")
        row = con.execute("SELECT number FROM participants WHERE token=?", (token,)).fetchone()
        if row:
            con.execute("COMMIT")
            return row[0], "existing"
        state, _ = status(con)
        if state != "open":
            con.execute("ROLLBACK")
            return None, state
        taken = {r[0] for r in con.execute("SELECT number FROM participants")}
        free = [n for n in range(LOW, HIGH + 1) if n not in taken]
        if not free:
            con.execute("ROLLBACK")
            return None, "full"
        number = random.choice(free)
        con.execute("INSERT INTO participants VALUES (?,?,?)", (number, token, time.time()))
        con.execute("COMMIT")
        return number, "new"
    finally:
        con.close()


# ---------- pages ----------
STYLE = """
<meta name=viewport content="width=device-width,initial-scale=1">
<style>
 body{font-family:system-ui,sans-serif;margin:0;background:#0f172a;color:#f8fafc;display:grid;place-items:center;min-height:100vh;text-align:center}
 main{padding:24px;max-width:420px;width:100%}
 h1{font-size:1.5rem;margin:0 0 8px}
 p{color:#cbd5e1}
 .num{font-size:5rem;font-weight:800;letter-spacing:2px;margin:12px 0;color:#fbbf24}
 button,.btn{background:#fbbf24;color:#0f172a;border:0;border-radius:12px;padding:14px 28px;font-size:1.1rem;font-weight:700;cursor:pointer;text-decoration:none;display:inline-block}
 .btn.alt{background:#334155;color:#f8fafc}
 .card{background:#1e293b;border-radius:16px;padding:20px;margin:12px 0}
 input{padding:12px;border-radius:10px;border:0;width:100%;box-sizing:border-box;font-size:1rem;margin-bottom:12px}
</style>
"""

REGISTER = STYLE + """
<title>Join</title>
<main>
{% if number %}
  <h1>You're in!</h1>
  <p>Your participant number</p>
  <div class=num>{{ number }}</div>
  <p>Keep this screen. Your number is saved to this device.</p>
{% elif state == 'waiting' %}
  <h1>Registration hasn't opened yet</h1><p>Please wait for the organizer, then refresh this page.</p>
{% elif state == 'closed' %}
  <h1>Registration is closed</h1><p>Sorry, the sign-up window has ended.</p>
{% elif state == 'full' %}
  <h1>All numbers are taken</h1><p>Sorry, all {{ total }} spots are filled.</p>
{% else %}
  <h1>Join the draw</h1>
  <p>Tap the button to get your random participant number.</p>
  <p>Closes in <b id=t>{{ left }}</b> s</p>
  <form method=post><button>Get my number</button></form>
  <script>let s={{ left }};setInterval(()=>{s=Math.max(0,s-1);document.getElementById('t').textContent=s;if(!s)location.reload()},1000)</script>
{% endif %}
</main>
"""

LOGIN = STYLE + """
<title>Organizer</title><main><h1>Organizer login</h1>
<form method=post><input type=password name=password placeholder=Password autofocus><button>Log in</button></form>
{% if error %}<p>Wrong password.</p>{% endif %}</main>
"""

ADMIN = STYLE + """
<title>Organizer</title>
<main>
 <h1>Organizer</h1>
 <div class=card>
  <p>Status: <b>{{ state|upper }}</b>{% if state == 'open' %} - closes in <b id=t>{{ left }}</b> s{% endif %}</p>
  <p>Registered: <b>{{ n }}</b> / {{ total }}</p>
 </div>
 <form method=post action="{{ url_for('admin_start') }}"><button>Start {{ minutes }}-minute registration</button></form><br>
 <form method=post action="{{ url_for('admin_close') }}"><button class="btn alt">Close now</button></form><br>
 <a class=btn href="{{ url_for('download') }}">Download Participants.xlsx</a>
 <div class=card>
  <p>QR code for the poster<br><small>{{ reg_url }}</small></p>
  <div id=qr style="background:#fff;padding:12px;display:inline-block;border-radius:8px"></div><br><br>
  <a class="btn alt" id=dl download="registration-qr.png" href="#">Download QR image</a>
 </div>
 <form method=post action="{{ url_for('admin_reset') }}" onsubmit="return confirm('Delete all registrations and start over?')">
  <button class="btn alt">Reset everything</button></form>
</main>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<script>
 new QRCode(document.getElementById('qr'),{text:{{ reg_url|tojson }},width:256,height:256,correctLevel:QRCode.CorrectLevel.H});
 setTimeout(()=>{const c=document.querySelector('#qr canvas');if(c)document.getElementById('dl').href=c.toDataURL('image/png')},300);
 setInterval(()=>location.reload(),5000);
</script>
"""


@app.route("/")
def index():
    return redirect(url_for("register"))


@app.route("/register", methods=["GET", "POST"])
def register():
    token = request.cookies.get("pid") or secrets.token_hex(16)
    number = None
    if request.method == "POST":
        number, _ = join(token)
    else:  # returning visitor: show their existing number
        con = db()
        row = con.execute("SELECT number FROM participants WHERE token=?", (token,)).fetchone()
        con.close()
        number = row[0] if row else None
    con = db()
    state, left = status(con)
    con.close()
    if not number and request.method == "POST":
        state = state if state != "open" else "full"
    resp = make_response(render_template_string(
        REGISTER, number=number, state=state, left=left, total=HIGH - LOW + 1))
    resp.set_cookie("pid", token, max_age=86400, httponly=True, samesite="Lax")
    return resp


def require_admin():
    if not session.get("admin"):
        abort(redirect(url_for("admin_login")))


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if secrets.compare_digest(request.form.get("password", ""), ADMIN_PASSWORD):
            session["admin"] = True
            return redirect(url_for("admin"))
        return render_template_string(LOGIN, error=True)
    return render_template_string(LOGIN, error=False)


@app.route("/admin")
def admin():
    require_admin()
    con = db()
    state, left = status(con)
    n = count(con)
    con.close()
    base = PUBLIC_URL or request.host_url.rstrip("/")
    return render_template_string(ADMIN, state=state, left=left, n=n, total=HIGH - LOW + 1,
                                  minutes=WINDOW_SECONDS // 60, reg_url=base + "/register")


@app.route("/admin/start", methods=["POST"])
def admin_start():
    require_admin()
    con = db()
    now = time.time()
    con.execute("INSERT OR REPLACE INTO state VALUES ('opens_at', ?)", (now,))
    con.execute("INSERT OR REPLACE INTO state VALUES ('closes_at', ?)", (now + WINDOW_SECONDS,))
    con.close()
    return redirect(url_for("admin"))


@app.route("/admin/close", methods=["POST"])
def admin_close():
    require_admin()
    con = db()
    now = time.time()
    if get_window(con)[0] is not None:
        con.execute("INSERT OR REPLACE INTO state VALUES ('closes_at', ?)", (now,))
    con.close()
    return redirect(url_for("admin"))


@app.route("/admin/reset", methods=["POST"])
def admin_reset():
    require_admin()
    con = db()
    con.execute("DELETE FROM participants")
    con.execute("DELETE FROM state")
    con.close()
    return redirect(url_for("admin"))


@app.route("/admin/download")
def download():
    require_admin()
    con = db()
    numbers = [r[0] for r in con.execute("SELECT number FROM participants ORDER BY number")]
    con.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Participants"
    ws.append(["Participant Number"])
    for n in numbers:
        ws.append([n])
    ws.column_dimensions["A"].width = 22
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Participants.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
