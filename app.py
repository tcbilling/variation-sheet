from flask import Flask, request, send_file
import xlsxwriter, io, os, json, csv
from datetime import datetime, date, timezone

app = Flask(__name__)

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
LOG_FILE  = os.path.join(BASE_DIR, 'download_log.csv')
HTML_PAGE = open(os.path.join(BASE_DIR, 'index.html'), encoding='utf-8').read()

# ── Logging setup ─────────────────────────────────────────────────────────────
LOG_HEADERS = [
    'timestamp_ist', 'ip', 'user_agent',
    'total_clients', 'total_size', 'total_amount', 'avg_yield',
    'clients_json', 'filename'
]

def ensure_log():
    """Create log file with headers if it doesn't exist."""
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(LOG_HEADERS)

def write_log(parsed, total_size, total_amnt, avg_yield, filename):
    ensure_log()
    # IST = UTC+5:30
    now_utc = datetime.now(timezone.utc)
    ist_offset = 5 * 60 + 30  # minutes
    from datetime import timedelta
    now_ist = now_utc + timedelta(minutes=ist_offset)
    timestamp = now_ist.strftime('%Y-%m-%d %H:%M:%S IST')

    ip = (request.headers.get('X-Forwarded-For', '') or '').split(',')[0].strip() \
         or request.remote_addr or 'unknown'
    ua = request.headers.get('User-Agent', 'unknown')[:120]

    clients_summary = [
        {'name': r['name'], 'size': r['size'], 'amnt': r['amnt'], 'reason': r['reason']}
        for r in parsed
    ]

    with open(LOG_FILE, 'a', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow([
            timestamp,
            ip,
            ua,
            len(parsed),
            total_size,
            total_amnt,
            round(avg_yield, 2),
            json.dumps(clients_summary, ensure_ascii=False),
            filename
        ])

# ── Helpers ───────────────────────────────────────────────────────────────────
def safe_num(v):
    try:
        f = float(v)
        return f if f == f else 0
    except (TypeError, ValueError):
        return 0

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return HTML_PAGE

@app.route('/generate', methods=['POST'])
def generate():
    data      = request.get_json(force=True)
    rows_data = data.get('rows', [])

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Variation Sheet')

    # ── Colours ──────────────────────────────────────
    DARK_BLUE = '#1F3864';  MED_BLUE  = '#4472C4'
    WHITE     = '#FFFFFF';  TOTAL_BG  = '#BDD7EE'
    GREEN_BG  = '#C6EFCE';  GREEN_FG  = '#276221'
    YELLOW_BG = '#FFEB9C';  YELLOW_FG = '#9C6500'
    RED_BG    = '#FFC7CE';  RED_FG    = '#9C0006'
    BLACK     = '#000000'

    def f(**kw):
        base = {'font_name': 'Arial', 'font_size': 11, 'font_color': BLACK}
        base.update(kw)
        return wb.add_format(base)

    hdr_dk  = f(bold=True, font_color=WHITE, bg_color=DARK_BLUE, border=1, align='center', valign='vcenter', text_wrap=True)
    hdr_dk_L= f(bold=True, font_color=WHITE, bg_color=DARK_BLUE, border=1, align='left',   valign='vcenter')
    hdr_bl  = f(bold=True, font_color=WHITE, bg_color=MED_BLUE,  border=1, align='center', valign='vcenter', text_wrap=True)
    d_ctr   = f(border=1, align='center', valign='vcenter')
    d_lft   = f(border=1, align='left',   valign='vcenter')
    v_ok    = f(font_color=GREEN_FG,  bg_color=GREEN_BG,  border=1, align='center', valign='vcenter', num_format='0%')
    v_warn  = f(font_color=YELLOW_FG, bg_color=YELLOW_BG, border=1, align='center', valign='vcenter', num_format='0%')
    v_bad   = f(font_color=RED_FG,    bg_color=RED_BG,    border=1, align='center', valign='vcenter', num_format='0%')
    s_ok    = f(font_color=GREEN_FG,  bg_color=GREEN_BG,  border=1, align='center', valign='vcenter', num_format='0%')
    s_warn  = f(font_color=YELLOW_FG, bg_color=YELLOW_BG, border=1, align='center', valign='vcenter', num_format='0%')
    s_bad   = f(font_color=RED_FG,    bg_color=RED_BG,    border=1, align='center', valign='vcenter', num_format='0%')
    tot     = f(bold=True, bg_color=TOTAL_BG, border=1, align='center', valign='vcenter')
    note_f  = f(bold=True, italic=True)

    # ── Column widths ─────────────────────────────────
    ws.set_column('A:A', 4.1);  ws.set_column('B:B', 5.6)
    ws.set_column('C:C', 32.1); ws.set_column('D:D', 11.4)
    ws.set_column('E:E', 10.2); ws.set_column('F:F', 13.2)
    ws.set_column('G:G', 13.2); ws.set_column('H:H', 13.2)
    ws.set_column('I:I', 13.2); ws.set_column('J:J', 36.3)

    # ── Header row ───────────────────────────────────
    ws.set_row(3, 29)
    ws.write(3, 1, 'S.N',                                     hdr_dk)
    ws.write(3, 2, 'Client Name',                              hdr_dk_L)
    ws.write(3, 3, 'Size in Sqcm',                             hdr_dk)
    ws.write(3, 4, 'Amnt',                                     hdr_dk)
    ws.write(3, 5, 'Net Yeild',                                hdr_bl)
    ws.write(3, 6, 'Avg Yield',                                hdr_bl)
    ws.write(3, 7, 'Variation',                                hdr_bl)
    ws.write(3, 8, 'Space Allocation',                         hdr_bl)
    ws.write(3, 9, 'Reason where Variation is more than 30%',  hdr_dk_L)

    # ── Parse: skip fully blank rows ─────────────────
    all_parsed = [{
        'name':   str(r.get('name', '') or '').strip(),
        'size':   safe_num(r.get('size', 0)),
        'amnt':   safe_num(r.get('amnt', 0)),
        'reason': str(r.get('reason', '') or ''),
    } for r in rows_data]

    parsed     = [r for r in all_parsed if r['name']]          # named rows only → Excel
    active     = [r for r in parsed    if r['size'] > 0]       # named + sized  → averages
    total_size = sum(r['size'] for r in active)
    total_amnt = sum(r['amnt'] for r in active)
    avg_yield  = total_amnt / total_size if total_size > 0 else 0

    # ── Data rows ─────────────────────────────────────
    for i, row in enumerate(parsed):
        xr   = 4 + i
        sz   = row['size'];  am = row['amnt'];  name = row['name']
        net  = am / sz if sz > 0 else 0
        vari = (net - avg_yield) / avg_yield if avg_yield > 0 and sz > 0 else 0
        sa   = sz / total_size if total_size > 0 and sz > 0 else 0
        vp   = vari * 100;  sp = sa * 100

        ws.write(xr, 1, i + 1,  d_ctr)
        ws.write(xr, 2, name,   d_lft)
        ws.write(xr, 3, sz,     d_ctr)
        ws.write(xr, 4, am,     d_ctr)
        ws.write(xr, 5, round(net, 0) if sz > 0 else 0,      d_ctr)
        ws.write(xr, 6, round(avg_yield, 0) if sz > 0 else 0, d_ctr)

        vfmt = v_ok if -20<=vp<=20 else (v_warn if abs(vp)<=100 else v_bad)
        if sz == 0 or not name: vfmt = v_bad
        ws.write(xr, 7, vari if sz > 0 else 0, vfmt)

        sfmt = s_bad if sp > 25 else (s_warn if sp >= 20 else s_ok)
        ws.write(xr, 8, sa if sz > 0 else 0, sfmt)
        ws.write(xr, 9, row['reason'], d_lft)

    # ── Totals row ────────────────────────────────────
    tr = 4 + len(parsed)
    for col in range(1, 10): ws.write(tr, col, '', tot)
    ws.write(tr, 3, total_size,           tot)
    ws.write(tr, 4, total_amnt,           tot)
    ws.write(tr, 5, round(avg_yield, 0),  tot)
    ws.write(tr, 6, round(avg_yield, 0),  tot)

    # ── Note ──────────────────────────────────────────
    ws.write(tr + 1, 2,
        'Note: If space allocation is more than 25% for any client, '
        'then approval from Sameer Sir is required', note_f)

    wb.close()
    output.seek(0)

    filename = f'Variation_Sheet_{date.today()}.xlsx'

    # ── Write log entry ───────────────────────────────
    try:
        write_log(parsed, total_size, total_amnt, avg_yield, filename)
    except Exception as log_err:
        print(f'[LOG ERROR] {log_err}')   # never crash the download for a log failure

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

# ── Log viewer ────────────────────────────────────────────────────────────────
LOG_VIEWER_KEY = os.environ.get('LOG_KEY', 'admin123')   # set LOG_KEY env var on your server

@app.route('/logs')
def view_logs():
    key = request.args.get('key', '')
    if key != LOG_VIEWER_KEY:
        return '<h2>401 - Invalid key. Add ?key=YOUR_LOG_KEY to the URL.</h2>', 401

    ensure_log()
    try:
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            rows_raw = list(csv.DictReader(f))
    except Exception as e:
        return f'<p>Error reading log: {e}</p>', 500

    total = len(rows_raw)
    rows_raw.reverse()   # newest first

    # Build HTML table
    rows_html = ''
    for r in rows_raw:
        try:
            clients = json.loads(r.get('clients_json', '[]'))
            client_list = '<br>'.join(
                f"• {c['name']} | Sz:{int(c['size'])} | Amt:{int(c['amnt'])}"
                + (f" | {c['reason']}" if c.get('reason') else '')
                for c in clients
            ) or '—'
        except Exception:
            client_list = r.get('clients_json', '—')

        rows_html += f"""
        <tr>
          <td>{r.get('timestamp_ist','')}</td>
          <td>{r.get('ip','')}</td>
          <td>{r.get('total_clients','')}</td>
          <td>{r.get('total_size','')}</td>
          <td>{r.get('total_amount','')}</td>
          <td>{r.get('avg_yield','')}</td>
          <td class="clients">{client_list}</td>
          <td>{r.get('filename','')}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Download Log — Variation Sheet</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:Arial,sans-serif;background:#f0f4f8;padding:24px}}
  h1{{color:#1F3864;margin-bottom:4px;font-size:1.3rem}}
  .sub{{color:#666;font-size:.85rem;margin-bottom:20px}}
  .stats{{display:flex;gap:16px;margin-bottom:20px;flex-wrap:wrap}}
  .stat{{background:#1F3864;color:#fff;border-radius:8px;padding:12px 20px;text-align:center}}
  .stat .n{{font-size:1.6rem;font-weight:700}}
  .stat .l{{font-size:.75rem;opacity:.8}}
  table{{border-collapse:collapse;width:100%;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.08);font-size:.83rem}}
  thead th{{background:#1F3864;color:#fff;padding:10px 12px;text-align:left;white-space:nowrap}}
  tbody tr:nth-child(even){{background:#f5f8ff}}
  tbody tr:hover{{background:#e8eef9}}
  td{{padding:8px 12px;border-bottom:1px solid #e0e0e0;vertical-align:top}}
  .clients{{font-size:.78rem;color:#444;max-width:320px}}
  .dl-btn{{display:inline-block;margin-bottom:16px;padding:8px 20px;background:#1F5C1F;color:#fff;border-radius:6px;text-decoration:none;font-size:.85rem;font-weight:700}}
  .dl-btn:hover{{opacity:.85}}
</style>
</head>
<body>
  <h1>&#x1F4CB; Download Log — Variation Sheet</h1>
  <div class="sub">Showing {total} download(s) — newest first</div>
  <div class="stats">
    <div class="stat"><div class="n">{total}</div><div class="l">Total Downloads</div></div>
  </div>
  <a class="dl-btn" href="/logs/download?key={key}">&#x2B07; Download Log as CSV</a>
  <table>
    <thead>
      <tr>
        <th>Timestamp (IST)</th>
        <th>IP Address</th>
        <th>Clients</th>
        <th>Total Size</th>
        <th>Total Amount</th>
        <th>Avg Yield</th>
        <th>Client Details</th>
        <th>File</th>
      </tr>
    </thead>
    <tbody>{rows_html}</tbody>
  </table>
</body></html>"""

@app.route('/logs/download')
def download_log():
    key = request.args.get('key', '')
    if key != LOG_VIEWER_KEY:
        return '401 Invalid key', 401
    ensure_log()
    return send_file(LOG_FILE, mimetype='text/csv', as_attachment=True,
                     download_name=f'variation_log_{date.today()}.csv')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
