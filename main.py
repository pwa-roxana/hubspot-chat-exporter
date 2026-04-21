from flask import Flask, request, jsonify, send_file, render_template_string
import requests
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io
import os

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html>
<head>
  <title>PWA Live Chat Exporter</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f4f6f9; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
    .card { background: white; border-radius: 12px; padding: 2rem; width: 100%; max-width: 520px; box-shadow: 0 2px 16px rgba(0,0,0,0.08); }
    .logo { font-size: 13px; font-weight: 600; color: #f97316; letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 0.5rem; }
    h1 { font-size: 22px; font-weight: 600; color: #1a1a2e; margin-bottom: 0.25rem; }
    p.sub { font-size: 14px; color: #6b7280; margin-bottom: 1.75rem; }
    label { display: block; font-size: 13px; font-weight: 500; color: #374151; margin-bottom: 5px; }
    input, select { width: 100%; padding: 10px 12px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 14px; color: #111827; margin-bottom: 1rem; outline: none; transition: border 0.2s; }
    input:focus, select:focus { border-color: #1F3864; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    button { width: 100%; padding: 12px; background: #1F3864; color: white; border: none; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; transition: background 0.2s; margin-top: 0.25rem; }
    button:hover { background: #162a4a; }
    button:disabled { background: #9ca3af; cursor: not-allowed; }
    .status { margin-top: 1.25rem; padding: 12px 14px; border-radius: 8px; font-size: 14px; display: none; }
    .status.info { background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe; }
    .status.success { background: #f0fdf4; color: #166534; border: 1px solid #bbf7d0; }
    .status.error { background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; }
    .progress { height: 4px; background: #e5e7eb; border-radius: 4px; margin-top: 10px; overflow: hidden; display: none; }
    .progress-bar { height: 100%; background: #1F3864; width: 0%; transition: width 0.4s; border-radius: 4px; }
    a.dl { display: block; text-align: center; margin-top: 1rem; padding: 11px; background: #166534; color: white; border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 14px; }
  </style>
</head>
<body>
<div class="card">
  <div class="logo">Pacific West Academy</div>
  <h1>Live Chat Exporter</h1>
  <p class="sub">Export all HubSpot conversations to Excel</p>

  <label>HubSpot Private App Token</label>
  <input type="password" id="token" placeholder="pat-na2-xxxxxxxx..." />

  <div class="row">
    <div>
      <label>From Date (optional)</label>
      <input type="date" id="from_date" />
    </div>
    <div>
      <label>To Date (optional)</label>
      <input type="date" id="to_date" />
    </div>
  </div>

  <button id="btn" onclick="runExport()">Export to Excel</button>

  <div class="progress" id="progress"><div class="progress-bar" id="bar"></div></div>
  <div class="status" id="status"></div>
  <div id="dl-area"></div>
</div>

<script>
async function runExport() {
  const token = document.getElementById('token').value.trim();
  if (!token) { showStatus('Please enter your HubSpot token.', 'error'); return; }

  const btn = document.getElementById('btn');
  btn.disabled = true;
  btn.textContent = 'Exporting...';
  document.getElementById('dl-area').innerHTML = '';
  document.getElementById('progress').style.display = 'block';
  setBar(10);
  showStatus('Connecting to HubSpot...', 'info');

  try {
    const body = {
      token,
      from_date: document.getElementById('from_date').value || null,
      to_date: document.getElementById('to_date').value || null
    };

    setBar(30);
    showStatus('Fetching conversations and messages... this may take a minute.', 'info');

    const resp = await fetch('/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    setBar(90);

    if (!resp.ok) {
      const err = await resp.json();
      throw new Error(err.error || 'Export failed');
    }

    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const date = new Date().toISOString().slice(0,10);
    const filename = `hubspot_live_chats_${date}.xlsx`;

    setBar(100);
    showStatus('Export complete!', 'success');
    document.getElementById('dl-area').innerHTML = `<a class="dl" href="${url}" download="${filename}">Download Excel File</a>`;
  } catch(e) {
    showStatus('Error: ' + e.message, 'error');
    document.getElementById('progress').style.display = 'none';
  } finally {
    btn.disabled = false;
    btn.textContent = 'Export to Excel';
  }
}

function showStatus(msg, type) {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status ' + type;
  el.style.display = 'block';
}

function setBar(pct) {
  document.getElementById('bar').style.width = pct + '%';
}
</script>
</body>
</html>
"""

BASE = "https://api.hubapi.com"
agent_cache = {}


def fmt_dt(date_str):
    if not date_str:
        return ""
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        return dt.strftime("%m/%d/%Y %I:%M %p")
    except:
        return date_str


def parse_ts(date_str):
    if not date_str:
        return None
    return datetime.fromisoformat(date_str.replace("Z", "+00:00")).timestamp() * 1000


def fetch_all_conversations(headers, from_ts, to_ts):
    conversations = []
    after = None
    seen = set()
    while True:
        params = {"limit": 50}
        if after:
            params["after"] = after
        resp = requests.get(f"{BASE}/conversations/v3/conversations/threads", headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        results = data.get("results", [])
        if not results:
            break
        for conv in results:
            created = conv.get("createdAt")
            if created:
                ts = parse_ts(created)
                if from_ts and ts < from_ts:
                    continue
                if to_ts and ts > to_ts:
                    continue
            conversations.append(conv)
        next_after = data.get("paging", {}).get("next", {}).get("after")
        if not next_after or next_after in seen:
            break
        seen.add(next_after)
        after = next_after
    return conversations


def fetch_contact(headers, contact_id):
    if not contact_id:
        return "", "", ""
    try:
        resp = requests.get(
            f"{BASE}/crm/v3/objects/contacts/{contact_id}",
            headers=headers,
            params={"properties": "firstname,lastname,email,phone"}
        )
        if resp.ok:
            props = resp.json().get("properties", {})
            name = f"{props.get('firstname') or ''} {props.get('lastname') or ''}".strip()
            return name, props.get("email", ""), props.get("phone", "")
    except:
        pass
    return "", "", ""


def resolve_agent(headers, actor_id):
    if not actor_id:
        return ""
    if actor_id in agent_cache:
        return agent_cache[actor_id]
    user_id = actor_id.replace("A-", "") if actor_id.startswith("A-") else actor_id
    name = actor_id
    try:
        resp = requests.get(f"{BASE}/crm/v3/owners/{user_id}", headers=headers)
        if resp.ok:
            data = resp.json()
            first = data.get("firstName", "") or ""
            last = data.get("lastName", "") or ""
            email = data.get("email", "") or ""
            full = f"{first} {last}".strip()
            name = full if full else email if email else actor_id
    except:
        pass
    agent_cache[actor_id] = name
    return name


def fetch_messages(headers, thread_id):
    try:
        resp = requests.get(
            f"{BASE}/conversations/v3/conversations/threads/{thread_id}/messages",
            headers=headers,
            params={"limit": 100}
        )
        resp.raise_for_status()
        return resp.json().get("results", [])
    except:
        return []


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/export", methods=["POST"])
def export():
    body = request.get_json()
    token = body.get("token", "").strip()
    from_date = body.get("from_date")
    to_date = body.get("to_date")

    if not token:
        return jsonify({"error": "Token is required"}), 400

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    from_ts = parse_ts(from_date + "T00:00:00Z") if from_date else None
    to_ts = parse_ts(to_date + "T23:59:59Z") if to_date else None

    try:
        conversations = fetch_all_conversations(headers, from_ts, to_ts)
    except Exception as e:
        return jsonify({"error": f"Failed to fetch conversations: {str(e)}"}), 400

    wb = Workbook()
    ws = wb.active
    ws.title = "Live Chat Conversations"

    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    col_headers = ["Thread ID", "Date", "Status", "Visitor Name", "Visitor Email",
                   "Visitor Phone", "Assigned To", "Sender", "Message Time", "Message"]
    col_widths = [15, 20, 12, 22, 32, 18, 22, 16, 20, 80]

    for col, (h, w) in enumerate(zip(col_headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 30

    fill_a = PatternFill("solid", fgColor="EEF2F7")
    fill_b = PatternFill("solid", fgColor="FFFFFF")
    row_num = 2

    for i, conv in enumerate(conversations):
        tid = conv.get("id", "")
        contact_id = conv.get("contactId") or conv.get("associatedContactId")
        contact_name, contact_email, contact_phone = fetch_contact(headers, contact_id)
        assigned_raw = conv.get("assignedTo") or conv.get("assignedActorId") or ""
        assigned = resolve_agent(headers, assigned_raw) if assigned_raw else ""
        status = (conv.get("status") or "").replace("_", " ").title()
        created = fmt_dt(conv.get("createdAt", ""))
        messages = fetch_messages(headers, tid)
        fill = fill_a if i % 2 == 0 else fill_b

        if not messages:
            row_data = [tid, created, status, contact_name, contact_email,
                        contact_phone, assigned, "", "", ""]
            for col, val in enumerate(row_data, 1):
                cell = ws.cell(row=row_num, column=col, value=val)
                cell.fill = fill
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            row_num += 1
        else:
            for msg in messages:
                sender_type = msg.get("senderType") or msg.get("type") or ""
                if sender_type == "WELCOME_MESSAGE":
                    continue
                msg_text = (msg.get("text") or msg.get("body") or msg.get("richText") or "").replace("\n", " ").strip()
                if not msg_text:
                    continue
                sender_raw = msg.get("sender", {})
                if isinstance(sender_raw, dict):
                    sender_actor = sender_raw.get("actorId", "") or ""
                    sender = resolve_agent(headers, sender_actor) if sender_actor.startswith("A-") else sender_type
                else:
                    sender = sender_type
                msg_time = fmt_dt(msg.get("createdAt", ""))
                row_data = [tid, created, status, contact_name, contact_email,
                            contact_phone, assigned, sender, msg_time, msg_text]
                for col, val in enumerate(row_data, 1):
                    cell = ws.cell(row=row_num, column=col, value=val)
                    cell.fill = fill
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                row_num += 1

    ws.freeze_panes = "A2"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"hubspot_live_chats_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
