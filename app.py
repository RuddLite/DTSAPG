# app.py  –  thin Flask wrapper for the Arbitration Packet Generator (APG)

from functools import wraps
from flask import Flask, render_template, request, send_file, Response, abort, redirect, url_for
from arb_packet_generator import make_arb_packet, states_with_case_law
from field_schema import FIELDS, TOP_FIELDS, BOTTOM_FIELDS, STATE_OPTIONS
import tempfile, pathlib, os, re, uuid, shutil, zipfile, time

app = Flask(__name__)

# ─── simple Basic-Auth creds (override in Render dashboard) ────────────────
USER = os.getenv("APP_USER", "cash")
PW   = os.getenv("APP_PASS", "staff")

def protect(view):
    @wraps(view)
    def _wrap(*a, **kw):
        auth = request.authorization
        ok = auth and auth.username == USER and auth.password == PW
        if not ok:
            return Response(
                "Login required", 401,
                {"WWW-Authenticate": 'Basic realm="ArbPacketGenerator"'}
            )
        return view(*a, **kw)
    return _wrap

# ─── helpers ───────────────────────────────────────────────────────────────
_safe_re = re.compile(r"[^A-Za-z0-9\-_]+")

def safe_filename(name: str) -> str:
    cleaned = _safe_re.sub("", (name or "").strip())
    return cleaned or "arb_packet"

# In-memory registry of generated packets, keyed by a one-time token.
# Good enough for a small internal tool; swap for S3/Render disk + a real
# TTL cache if this needs to survive dyno restarts / scale past one worker.
_PACKETS: dict[str, dict] = {}
_PACKET_TTL_SECONDS = 60 * 60  # 1 hour

def _cleanup_expired():
    now = time.time()
    expired = [t for t, v in _PACKETS.items() if now - v["created"] > _PACKET_TTL_SECONDS]
    for t in expired:
        shutil.rmtree(_PACKETS[t]["dir"], ignore_errors=True)
        _PACKETS.pop(t, None)

# ─── main route: intake dashboard ───────────────────────────────────────────
@app.route("/", methods=["GET"])
@protect
def form():
    return render_template(
        "arb_form.html",
        top_fields=TOP_FIELDS,
        bottom_fields=BOTTOM_FIELDS,
        states=STATE_OPTIONS,
        states_with_case_law=set(states_with_case_law()),
    )

# ─── generate both documents ─────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
@protect
def generate():
    _cleanup_expired()
    data = {f["id"]: request.form.get(f["id"], "") for f in FIELDS}
    data["accident_state"] = request.form.get("accident_state", "").upper().strip()

    # basic required-field validation (conditional fields are validated
    # client-side against their depends_on state; server just checks the
    # always-required ones so a broken/disabled JS doesn't silently drop data)
    missing = [
        f["label"] for f in FIELDS
        if f.get("required") and not f.get("depends_on") and not data.get(f["id"], "").strip()
    ]
    if missing:
        abort(400, "Missing required fields: " + ", ".join(missing))

    try:
        numeric_ids = [f["id"] for f in FIELDS if f["type"] == "number"]
        for numeric_id in numeric_ids:
            if data.get(numeric_id):
                float(data[numeric_id])
    except ValueError:
        abort(400, "One or more numeric fields contain a non-numeric value. "
                    "Please remove any \"$\" or \"%\" symbols and resubmit.")

    token = uuid.uuid4().hex
    out_dir = pathlib.Path(tempfile.mkdtemp(prefix=f"apg_{token}_"))
    prefix = safe_filename(request.form.get("our_file") or data.get("accident_state") or "arb_packet")

    try:
        top_path, bottom_path = make_arb_packet(data, out_dir=out_dir, file_prefix=prefix)
    except ValueError as e:
        shutil.rmtree(out_dir, ignore_errors=True)
        abort(400, str(e))

    _PACKETS[token] = {
        "dir": str(out_dir),
        "top": str(top_path),
        "bottom": str(bottom_path),
        "created": time.time(),
    }
    return redirect(url_for("result", token=token))

# ─── results page: two downloads + a convenience zip ────────────────────────
@app.route("/result/<token>", methods=["GET"])
@protect
def result(token):
    packet = _PACKETS.get(token)
    if not packet:
        abort(404, "This packet has expired or was already downloaded. Please regenerate it.")
    return render_template(
        "result.html",
        token=token,
        top_name=pathlib.Path(packet["top"]).name,
        bottom_name=pathlib.Path(packet["bottom"]).name,
    )

_DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

@app.route("/download/<token>/<which>", methods=["GET"])
@protect
def download(token, which):
    packet = _PACKETS.get(token)
    if not packet:
        abort(404, "This packet has expired. Please regenerate it.")

    if which == "top":
        path = packet["top"]
    elif which == "bottom":
        path = packet["bottom"]
    elif which == "zip":
        zip_path = pathlib.Path(packet["dir"]) / "arb_packet.zip"
        if not zip_path.exists():
            with zipfile.ZipFile(zip_path, "w") as zf:
                zf.write(packet["top"], pathlib.Path(packet["top"]).name)
                zf.write(packet["bottom"], pathlib.Path(packet["bottom"]).name)
        return send_file(zip_path, as_attachment=True,
                          download_name="arb_packet.zip", mimetype="application/zip")
    else:
        abort(404)

    return send_file(path, as_attachment=True,
                      download_name=pathlib.Path(path).name, mimetype=_DOCX_MIME)

# ─── CLI entrypoint (optional local dev) ───────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)
