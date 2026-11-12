import io, json, os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
os.environ.update(KASHEF_DB=tempfile.mktemp(suffix=".db"), APP_PASSWORD="test-pass-1", COPYLEAKS_EMAIL="", COPYLEAKS_API_KEY="", GEMINI_API_KEY="")
import app as A
import services as sv
from docx import Document

AR = "تعد الجهة المختصة بالتخطيط الجهة المسؤولة عن إعداد المخططات في ليبيا. بلغ عدد المخططات 120 مخططًا في عام 2024.\n\n" \
     "وفي هذا السياق، تجدر الإشارة إلى أن التخطيط العمراني يلعب دورًا محوريًا في تحقيق التنمية المستدامة وتعزيز جودة الحياة للمواطنين في مختلف المدن.\n\n" \
     "كتب الموظف هذه الفقرة بنفسه بعد زيارة ميدانية إلى موقع المشروع، وذكر أن نسبة الإنجاز وصلت إلى ٤٥٪ حتى نهاية الشهر الماضي."

def test_parse_documented_response():
    r = sv.parse_copyleaks({"modelVersion":"v11.0","results":[
        {"classification":1,"probability":.9,"matches":[{"text":{"chars":{"starts":[0],"lengths":[10]}}}]},
        {"classification":2,"probability":1,"matches":[{"text":{"chars":{"starts":[10,50],"lengths":[20,5]}}}]}],
        "summary":{"human":.4,"ai":.6}})
    assert r["ai"] == .6 and r["ranges"] == [[10,30],[50,55]]

def test_copyleaks_request(monkeypatch):
    calls = []
    class R:
        def __init__(s, code, data): s.status_code, s._d = code, data
        def json(s): return s._d
    def fake(url, json=None, headers=None, timeout=None):
        calls.append((url, json, headers))
        if "login" in url: return R(200, {"access_token":"tok"})
        return R(200, {"results":[], "summary":{"ai":.1,"human":.9}})
    monkeypatch.setattr(sv.requests, "post", fake)
    monkeypatch.setenv("COPYLEAKS_EMAIL","a@b.c"); monkeypatch.setenv("COPYLEAKS_API_KEY","k")
    assert sv.Copyleaks().scan(AR, "ar")["ai"] == .1
    url, body, headers = calls[1]
    assert "/v2/writer-detector/" in url and len(url.split("/")[-2]) <= 36
    assert body["language"] == "ar" and body["sensitivity"] == 2 and headers["Authorization"] == "Bearer tok"

def test_humanize_flagged_only_keeps_others():
    spans = sv.split_paragraphs(AR)
    assert len(spans) == 3
    out, n = sv.humanize(AR, [[spans[1][0]+3, spans[1][0]+9]], "flagged", rewrite=lambda t: "نص جديد")
    assert n == 1
    assert out.split("\n\n") == [AR.split("\n\n")[0], "نص جديد", AR.split("\n\n")[2]]

def test_humanize_all_and_count_mismatch():
    seen = []
    def rw(t):
        seen.append(t)
        return "فقرة واحدة فقط" if "\n\n" in t else "م"
    out, n = sv.humanize(AR, [], "all", rewrite=rw)
    assert n == 3 and out.split("\n\n") == ["م","م","م"] and len(seen) == 4

def test_long_text_batches():
    paras = "\n\n".join(("ف%d " % i) + "كلمة " * 399 + "كلمة" for i in range(10))
    sizes = []
    def rw(t): sizes.append(len(t)); return t
    out, n = sv.humanize(paras, [], "all", rewrite=rw)
    assert out == paras and n == 10 and len(sizes) > 1 and max(sizes) < 8000

def test_missing_tokens():
    assert sv.missing_tokens("نسبة ٤٥٪ من 120 https://a.ly", "نسبة 45٪ من 130") == ["120", "https://a.ly"]
    assert sv.missing_tokens("بلغ ١٢٠", "بلغ 120") == []

def test_docx_rtl():
    doc = Document(io.BytesIO(sv.final_docx("تقرير", AR)))
    assert doc.paragraphs[1]._p.pPr.find(sv.qn("w:bidi")) is not None
    en = Document(io.BytesIO(sv.final_docx("Report", "Plain English paragraph.")))
    assert en.paragraphs[1]._p.pPr is None or en.paragraphs[1]._p.pPr.find(sv.qn("w:bidi")) is None

def test_flow():
    c = A.app.test_client()
    assert c.get("/").status_code == 302
    assert c.post("/login", data={"password":"no"}).status_code == 200
    c.post("/login", data={"password":"test-pass-1"})
    assert c.post("/api/reviews", json={"employee":"a","title":"b","text":"قصير"}).status_code == 400
    r = c.post("/api/reviews", json={"employee":"علي","title":"تقرير يناير","text":AR}).get_json()
    assert r["demo"] is True
    rid = r["id"]
    r = c.post(f"/api/reviews/{rid}/issues", json={"kind":"معلومة خاطئة","excerpt":"120 مخططًا","note":"العدد 100"}).get_json()
    assert len(r["issues"]) == 1
    assert c.post(f"/api/reviews/{rid}/humanize", json={"scope":"all"}).status_code == 200
    r = c.put(f"/api/reviews/{rid}/final", json={"text":"النص النهائي المعتمد"}).get_json()
    assert r["status"] == "approved" and r["original_text"] == AR
    d = c.get(f"/api/reviews/{rid}/final.docx")
    assert "النص النهائي المعتمد" in [p.text for p in Document(io.BytesIO(d.data)).paragraphs]
    rep = Document(io.BytesIO(c.get(f"/api/reviews/{rid}/report.docx").data))
    assert any("العدد 100" in p.text for p in rep.paragraphs)
    assert c.get("/api/reviews?q=علي").get_json()[0]["issue_count"] == 1


class R:
    def __init__(s, code, data): s.status_code, s._d = code, data
    def json(s): return s._d

def reply(text, finish="STOP"):
    return R(200, {"candidates":[{"finishReason":finish,"content":{"parts":[{"text":text}]}}]})

def test_gemini_request_and_scan(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY","k")
    monkeypatch.setenv("GEMINI_SCAN_RUNS","3")
    calls = []
    def fake(url, headers=None, json=None, timeout=None):
        import json as J
        calls.append((url, headers, json))
        ids = [i["id"] for i in J.loads(json["contents"][0]["parts"][0]["text"])]
        return reply(J.dumps({"items":[{"id":i,"ai_likelihood":[20,90,10][i],"reason":"ع"} for i in ids]}))
    monkeypatch.setattr(sv.requests, "post", fake)
    r = sv.gemini_scan(AR)
    url, headers, body = calls[0]
    assert url.endswith("/models/gemini-3.8-flash:generateContent") and headers["x-goog-api-key"] == "k"
    assert body["generationConfig"]["responseFormat"]["text"]["mimeType"] == "APPLICATION_JSON"
    spans = sv.split_paragraphs(AR)
    assert r["ranges"] == [list(spans[1])] and len(r["reasons"]) == 1 and r["reasons"][0]["score"] == 90
    assert .2 < r["ai"] < .7

def test_scan_retries_missing_paragraph(monkeypatch):
    monkeypatch.setenv("GEMINI_SCAN_RUNS","1")
    seen = []
    def call(system, prompt, schema):
        ids = [i["id"] for i in json.loads(prompt)]
        seen.append(ids)
        return json.dumps({"items":[{"id":i,"ai_likelihood":10,"reason":""} for i in ids[:-1 if len(seen)==1 else None]]})
    sv.gemini_scan(AR, call=call)
    assert seen == [[0,1,2],[2]]

class Cache(dict):
    def put(self, items): self.update(items)

def test_median_ignores_outlier(monkeypatch):
    monkeypatch.setenv("GEMINI_SCAN_RUNS","3")
    values = iter([30, 95, 35])
    def call(system, prompt, schema):
        v = next(values)
        return json.dumps({"items":[{"id":i["id"],"ai_likelihood":v,"reason":""} for i in json.loads(prompt)]})
    assert round(sv.gemini_scan(AR, call=call)["ai"], 2) == .35

def test_same_text_same_result_without_new_calls(monkeypatch):
    monkeypatch.setenv("GEMINI_SCAN_RUNS","3")
    import random
    calls = []
    def call(system, prompt, schema):
        calls.append(prompt)
        return json.dumps({"items":[{"id":i["id"],"ai_likelihood":random.randint(0,100),"reason":"س"} for i in json.loads(prompt)]})
    cache = Cache()
    first = sv.gemini_scan(AR, call=call, cache=cache)
    used = len(calls)
    assert used == 3
    for _ in range(5):
        assert sv.gemini_scan(AR, call=call, cache=cache) == first
    sv.gemini_scan(AR.replace("ليبيا.", "ليبيا.  "), call=call, cache=cache)
    assert len(calls) == used

def test_edit_rescores_only_changed_paragraph(monkeypatch):
    monkeypatch.setenv("GEMINI_SCAN_RUNS","1")
    calls = []
    def call(system, prompt, schema):
        items = json.loads(prompt); calls.append(len(items))
        return json.dumps({"items":[{"id":i["id"],"ai_likelihood":50,"reason":""} for i in items]})
    cache = Cache()
    sv.gemini_scan(AR, call=call, cache=cache)
    sv.gemini_scan(AR + " وأضاف الموظف جملة جديدة.", call=call, cache=cache)
    assert calls == [3, 1]

def test_cache_persists_through_app(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY","k")
    monkeypatch.setattr(A.sv, "call_gemini", lambda *a: "")
    import random
    def fake(text, cache=None, call=None):
        real = sv_scan(text, cache=cache, call=lambda s,p,c: json.dumps({"items":[{"id":i["id"],"ai_likelihood":random.randint(0,100),"reason":"س"} for i in json.loads(p)]}))
        return real
    sv_scan = sv.gemini_scan
    monkeypatch.setattr(A.sv, "gemini_scan", fake)
    c = A.app.test_client()
    c.post("/login", data={"password":"test-pass-1"})
    one = c.post("/api/reviews", json={"employee":"ع","title":"ت","text":AR}).get_json()
    two = c.post("/api/reviews", json={"employee":"ع","title":"ت","text":AR}).get_json()
    assert one["ai_score"] == two["ai_score"] and one["ranges"] == two["ranges"]

def test_rate_limit_falls_back(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY","k")
    monkeypatch.setattr(sv.time, "sleep", lambda s: None)
    urls = []
    def fake(url, **k):
        urls.append(url)
        return R(429, {"error":{"message":"quota"}}) if "3.8" in url else reply("نص")
    monkeypatch.setattr(sv.requests, "post", fake)
    assert sv.call_gemini("s","p") == "نص"
    assert len(urls) == 4 and "3.5-flash-lite" in urls[-1]

def test_all_limits_gone_message(monkeypatch):
    monkeypatch.setattr(sv.time, "sleep", lambda s: None)
    monkeypatch.setattr(sv.requests, "post", lambda *a, **k: R(429, {"error":{"message":"quota"}}))
    import pytest
    with pytest.raises(sv.ServiceError, match="الحد المجاني"):
        sv.call_gemini("s","p")

def test_rewrite_retries_lost_number():
    prompts = []
    def call(system, prompt):
        prompts.append(prompt)
        return "بلغ 120 مخططًا" if len(prompts) == 2 else "بلغ مخططات كثيرة"
    assert "120" in sv.rewrite_text("بلغ 120 مخططًا", call=call)
    assert "120" in prompts[1].split("\n")[0]

def test_old_database_upgrades():
    import sqlite3
    path = tempfile.mktemp(suffix=".db")
    c = sqlite3.connect(path)
    c.execute("CREATE TABLE reviews(id INTEGER PRIMARY KEY, employee TEXT, title TEXT, original_text TEXT, language TEXT, ai_score REAL, ranges TEXT, demo INTEGER, rewritten_text TEXT, final_text TEXT, status TEXT, created_at TEXT, approved_at TEXT)")
    c.commit(); c.close()
    old = A.DB; A.DB = path
    try:
        client = A.app.test_client()
        client.post("/login", data={"password":"test-pass-1"})
        assert client.get("/api/reviews").status_code == 200
    finally:
        A.DB = old
