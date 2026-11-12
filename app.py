"""كاشف: أداة محلية لمراجعة التقارير والبحوث."""
import hmac
import io
import json
import os
import re
import secrets
import sqlite3
import sys
from datetime import datetime
from functools import wraps
from pathlib import Path

BASE = Path(__file__).resolve().parent


def load_env():
    env = BASE / ".env"
    if not env.exists():
        return
    for line in env.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


load_env()

from flask import Flask, g, jsonify, redirect, render_template, request, send_file, session, url_for  # noqa: E402

import services as sv  # noqa: E402

DB = Path(os.getenv("KASHEF_DB", BASE / "kashef.db"))
SECRET_FILE = BASE / ".secret"
if not SECRET_FILE.exists():
    SECRET_FILE.write_text(secrets.token_hex(32))

app = Flask(__name__)
app.secret_key = SECRET_FILE.read_text()
app.config.update(
    MAX_CONTENT_LENGTH=25 * 1024 * 1024,
    SESSION_COOKIE_SAMESITE="Strict",
    SESSION_COOKIE_HTTPONLY=True,
)
detector = sv.Copyleaks()

SCHEMA = """
CREATE TABLE IF NOT EXISTS reviews(
  id INTEGER PRIMARY KEY,
  employee TEXT NOT NULL,
  title TEXT NOT NULL,
  original_text TEXT NOT NULL,
  language TEXT,
  ai_score REAL NOT NULL,
  ranges TEXT NOT NULL,
  demo INTEGER NOT NULL DEFAULT 0,
  reasons TEXT NOT NULL DEFAULT '[]',
  new_score REAL,
  rewritten_text TEXT,
  final_text TEXT,
  status TEXT NOT NULL DEFAULT 'review',
  created_at TEXT NOT NULL,
  approved_at TEXT
);
CREATE TABLE IF NOT EXISTS paragraph_scores(
  key TEXT PRIMARY KEY,
  score INTEGER NOT NULL,
  reason TEXT
);
CREATE TABLE IF NOT EXISTS issues(
  id INTEGER PRIMARY KEY,
  review_id INTEGER NOT NULL REFERENCES reviews(id) ON DELETE CASCADE,
  kind TEXT NOT NULL,
  excerpt TEXT,
  note TEXT
);
"""


def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys=ON")
        g.db.executescript(SCHEMA)
        columns = {c["name"] for c in g.db.execute("PRAGMA table_info(reviews)")}
        if "reasons" not in columns:  # قاعدة من الإصدار السابق
            g.db.execute("ALTER TABLE reviews ADD COLUMN reasons TEXT NOT NULL DEFAULT '[]'")
            g.db.execute("ALTER TABLE reviews ADD COLUMN new_score REAL")
    return g.db


@app.teardown_appcontext
def close_db(_):
    conn = g.pop("db", None)
    if conn:
        conn.close()


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def login_required(view):
    @wraps(view)
    def wrapper(*a, **k):
        if not session.get("ok"):
            if request.path.startswith("/api/"):
                return jsonify(error="انتهت الجلسة. سجّل الدخول من جديد."), 401
            return redirect(url_for("login"))
        return view(*a, **k)
    return wrapper


@app.errorhandler(sv.ServiceError)
def service_error(err):
    return jsonify(error=str(err)), 400


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="حجم الملف أكبر من 25 ميجابايت."), 413


@app.errorhandler(sv.requests.RequestException)
def network_error(_):
    return jsonify(error="تعذّر الوصول إلى الخدمة الخارجية. تحقق من اتصال الإنترنت ثم أعد المحاولة."), 502


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        expected = os.getenv("APP_PASSWORD", "")
        # Never accept an empty password, even if APP_PASSWORD is missing
        # (e.g. the app was started without run.bat / __main__ checks).
        if expected and hmac.compare_digest(request.form.get("password", "").encode(), expected.encode()):
            session["ok"] = True
            return redirect(url_for("index"))
        error = "كلمة الدخول غير صحيحة."
    return render_template("login.html", error=error)


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.get("/")
@login_required
def index():
    return render_template("index.html", demo=not live_scan())


class ScoreCache:
    """نفس الفقرة تأخذ دائمًا نفس التقييم."""

    def get(self, key):
        row = db().execute("SELECT score,reason FROM paragraph_scores WHERE key=?", (key,)).fetchone()
        return {"ai_likelihood": row["score"], "reason": row["reason"]} if row else None

    def put(self, items):
        db().executemany(
            "INSERT OR REPLACE INTO paragraph_scores(key,score,reason) VALUES(?,?,?)",
            [(k, int(v["ai_likelihood"]), str(v.get("reason", ""))) for k, v in items.items()],
        )
        db().commit()


def live_scan():
    return detector.configured or sv.gemini_configured()


def scan(text):
    if detector.configured:
        return detector.scan(text, sv.detect_language(text))
    if sv.gemini_configured():
        return sv.gemini_scan(text, cache=ScoreCache())
    return sv.demo_scan(text)


def review_json(row):
    data = dict(row)
    data["ranges"] = json.loads(data["ranges"])
    data["reasons"] = json.loads(data["reasons"])
    data["demo"] = bool(data["demo"])
    data["issues"] = [dict(i) for i in db().execute("SELECT * FROM issues WHERE review_id=? ORDER BY id", (row["id"],))]
    data["missing"] = sv.missing_tokens(data["original_text"], data["rewritten_text"]) if data["rewritten_text"] else []
    return data


def get_review(rid):
    row = db().execute("SELECT * FROM reviews WHERE id=?", (rid,)).fetchone()
    if not row:
        raise sv.ServiceError("هذا السجل غير موجود.")
    return row


@app.post("/api/extract")
@login_required
def api_extract():
    upload = request.files.get("file")
    if not upload:
        raise sv.ServiceError("اختر ملفًا أولًا.")
    text, warning = sv.extract_text(upload.filename, upload.read())
    title = re.sub(r"\.[^.]+$", "", upload.filename)
    return jsonify(text=text, warning=warning, title=title)


@app.post("/api/reviews")
@login_required
def api_create():
    body = request.get_json(force=True)
    employee = (body.get("employee") or "").strip()
    title = (body.get("title") or "").strip()
    text = (body.get("text") or "").strip()
    if not employee or not title:
        raise sv.ServiceError("اكتب اسم الموظف وعنوان المستند.")
    minimum = sv.MIN_CHARS if detector.configured else sv.MIN_CHARS_GEMINI
    if len(text) < minimum:
        raise sv.ServiceError(f"النص قصير ({len(text)} حرفًا). الفحص يحتاج {minimum} حرفًا على الأقل.")
    if len(text) > sv.MAX_CHARS:
        raise sv.ServiceError("النص أطول من 100,000 حرف. قسّمه إلى أجزاء.")
    language = sv.detect_language(text)
    result = scan(text)
    cur = db().execute(
        "INSERT INTO reviews(employee,title,original_text,language,ai_score,ranges,demo,reasons,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
        (employee, title, text, language, result["ai"], json.dumps(result["ranges"]), 0 if live_scan() else 1,
         json.dumps(result.get("reasons", []), ensure_ascii=False), now()),
    )
    db().commit()
    return jsonify(review_json(get_review(cur.lastrowid)))


@app.get("/api/reviews")
@login_required
def api_list():
    q = f"%{request.args.get('q', '').strip()}%"
    rows = db().execute(
        """SELECT r.id,r.employee,r.title,r.ai_score,r.status,r.created_at,r.demo,
                  (SELECT COUNT(*) FROM issues i WHERE i.review_id=r.id) AS issue_count
           FROM reviews r WHERE r.employee LIKE ? OR r.title LIKE ? ORDER BY r.id DESC""",
        (q, q),
    )
    return jsonify([dict(r) for r in rows])


@app.get("/api/reviews/<int:rid>")
@login_required
def api_get(rid):
    return jsonify(review_json(get_review(rid)))


@app.post("/api/reviews/<int:rid>/issues")
@login_required
def api_issue(rid):
    get_review(rid)
    body = request.get_json(force=True)
    kind = (body.get("kind") or "").strip()
    note = (body.get("note") or "").strip()
    if not kind or not note:
        raise sv.ServiceError("اختر نوع الخطأ واكتب وصفه.")
    db().execute("INSERT INTO issues(review_id,kind,excerpt,note) VALUES(?,?,?,?)", (rid, kind, (body.get("excerpt") or "").strip()[:500], note))
    db().commit()
    return jsonify(review_json(get_review(rid)))


@app.delete("/api/reviews/<int:rid>/issues/<int:iid>")
@login_required
def api_issue_delete(rid, iid):
    db().execute("DELETE FROM issues WHERE id=? AND review_id=?", (iid, rid))
    db().commit()
    return jsonify(review_json(get_review(rid)))


@app.post("/api/reviews/<int:rid>/humanize")
@login_required
def api_humanize(rid):
    row = get_review(rid)
    scope = request.get_json(force=True).get("scope", "all")
    if not sv.gemini_configured():
        text, count = row["original_text"], 0
    else:
        text, count = sv.humanize(row["original_text"], json.loads(row["ranges"]), scope)
    db().execute("UPDATE reviews SET rewritten_text=? WHERE id=?", (text, rid))
    db().commit()
    data = review_json(get_review(rid))
    data["rewritten_paragraphs"] = count
    return jsonify(data)


@app.post("/api/reviews/<int:rid>/rescan")
@login_required
def api_rescan(rid):
    get_review(rid)
    text = (request.get_json(force=True).get("text") or "").strip()
    if len(text) < sv.MIN_CHARS_GEMINI:
        raise sv.ServiceError("النص قصير جدًا على الفحص.")
    result = scan(text)
    db().execute("UPDATE reviews SET new_score=? WHERE id=?", (result["ai"], rid))
    db().commit()
    data = review_json(get_review(rid))
    data["new_flagged"] = len(result["ranges"])
    return jsonify(data)


@app.put("/api/reviews/<int:rid>/final")
@login_required
def api_final(rid):
    get_review(rid)
    text = (request.get_json(force=True).get("text") or "").strip()
    if not text:
        raise sv.ServiceError("النص النهائي فارغ.")
    db().execute("UPDATE reviews SET final_text=?, status='approved', approved_at=? WHERE id=?", (text, now(), rid))
    db().commit()
    return jsonify(review_json(get_review(rid)))


def word(data, name):
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=re.sub(r'[\\/:*?"<>|]', "-", name) + ".docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.get("/api/reviews/<int:rid>/final.docx")
@login_required
def api_final_docx(rid):
    row = get_review(rid)
    if not row["final_text"]:
        raise sv.ServiceError("اعتمد النص أولًا.")
    return word(sv.final_docx(row["title"], row["final_text"]), f"{row['title']} - النسخة المعتمدة")


@app.get("/api/reviews/<int:rid>/report.docx")
@login_required
def api_report_docx(rid):
    row = get_review(rid)
    issues = db().execute("SELECT * FROM issues WHERE review_id=? ORDER BY id", (rid,)).fetchall()
    return word(sv.report_docx(row, issues), f"تقرير {row['employee']} - {row['title']}")


if __name__ == "__main__":
    password = os.getenv("APP_PASSWORD", "")
    if len(password) < 8:
        sys.exit("اكتب كلمة دخول من 8 أحرف على الأقل في APP_PASSWORD داخل ملف .env ثم أعد التشغيل.")
    host, port = os.getenv("HOST", "127.0.0.1"), int(os.getenv("PORT", "8765"))
    print(f"افتح المتصفح على: http://{host}:{port}")
    app.run(host=host, port=port, threaded=True, debug=False)
