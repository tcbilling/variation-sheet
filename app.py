from flask import Flask, request, send_file
import xlsxwriter, io, os, json, csv
from datetime import datetime, date, timezone, timedelta

app     = Flask(__name__)
BASE    = os.path.dirname(os.path.abspath(__file__))
LOG     = os.path.join(BASE, 'download_log.csv')
HTML    = open(os.path.join(BASE, 'index.html'), encoding='utf-8').read()
LOG_KEY = os.environ.get('LOG_KEY', 'admin123')

EXEMPT_KW = ['generic','filler','research','survey']
isExempt  = lambda n: any(k in (n or '').lower() for k in EXEMPT_KW)
safeN     = lambda v: float(v) if v not in (None, '') else 0.0

def ensure_log():
    if not os.path.exists(LOG):
        with open(LOG,'w',newline='',encoding='utf-8') as f:
            csv.writer(f).writerow(['timestamp_ist','ip','ua','mfn','gfrs_size',
                'actual_sold','mfn_benchmark','clients','total_size','total_amnt',
                'avg_yield','clients_json','file'])

def write_log(parsed, tSz, tAm, avg, mfn, gfrs_size, actual_sold, mfn_rate, fname):
    ensure_log()
    now = datetime.now(timezone.utc) + timedelta(minutes=330)
    ip  = (request.headers.get('X-Forwarded-For','') or '').split(',')[0].strip() or request.remote_addr
    ua  = request.headers.get('User-Agent','')[:100]
    cl  = json.dumps([{'name':r['name'],'size':r['size'],'amnt':r['amnt']} for r in parsed], ensure_ascii=False)
    with open(LOG,'a',newline='',encoding='utf-8') as f:
        csv.writer(f).writerow([now.strftime('%Y-%m-%d %H:%M:%S IST'), ip, ua,
            mfn, gfrs_size, round(actual_sold,0), round(mfn_rate,0),
            len(parsed), tSz, tAm, round(avg,2), cl, fname])

@app.route('/')
def index(): return HTML

@app.route('/generate', methods=['POST'])
def generate():
    data      = request.get_json(force=True)
    rows_data = data.get('rows', [])
    mfn       = safeN(data.get('mfn', 0))
    gfrs_size = safeN(data.get('gfrs_size', 0))
    gfrs_amnt = safeN(data.get('gfrs_amnt', 0))

    out = io.BytesIO()
    wb  = xlsxwriter.Workbook(out, {'in_memory': True})
    ws  = wb.add_worksheet('Variation Sheet')

    # ── Colours ──────────────────────────────────────────────
    DARK  = '#1F3864'; BLUE  = '#4472C4'; WHITE = '#FFFFFF'
    TOTBG = '#BDD7EE'; YELBG = '#FFFF00'; LTBLU = '#DEEAF1'
    GRN_B = '#C6EFCE'; GRN_F = '#276221'
    AMB_B = '#FFEB9C'; AMB_F = '#9C6500'
    RED_B = '#FFC7CE'; RED_F = '#9C0006'
    BLK   = '#000000'

    def f(**kw):
        base = {'font_name':'Arial','font_size':11,'font_color':BLK}
        base.update(kw); return wb.add_format(base)

    hDk  = f(bold=True,font_color=WHITE,bg_color=DARK,border=1,align='center',valign='vcenter',text_wrap=True)
    hDkL = f(bold=True,font_color=WHITE,bg_color=DARK,border=1,align='left',valign='vcenter')
    hBl  = f(bold=True,font_color=WHITE,bg_color=BLUE,border=1,align='center',valign='vcenter',text_wrap=True)
    dCtr = f(border=1,align='center',valign='vcenter')
    dLft = f(border=1,align='left',valign='vcenter')
    pct  = f(border=1,align='center',valign='vcenter',num_format='0%')
    vOk  = f(border=1,align='center',valign='vcenter',num_format='0%',bg_color=GRN_B,font_color=GRN_F)
    vWn  = f(border=1,align='center',valign='vcenter',num_format='0%',bg_color=AMB_B,font_color=AMB_F)
    vBd  = f(border=1,align='center',valign='vcenter',num_format='0%',bg_color=RED_B,font_color=RED_F)
    saOk = f(border=1,align='center',valign='vcenter',num_format='0%',bg_color=GRN_B,font_color=GRN_F)
    saWn = f(border=1,align='center',valign='vcenter',num_format='0%',bg_color=AMB_B,font_color=AMB_F)
    saBd = f(border=1,align='center',valign='vcenter',num_format='0%',bg_color=RED_B,font_color=RED_F)
    mOk  = f(border=1,align='center',valign='vcenter',num_format='0%',font_color=GRN_F)
    mBd  = f(border=1,align='center',valign='vcenter',num_format='0%',font_color=RED_F)
    exC  = f(bold=True,bg_color=YELBG,border=1,align='center',valign='vcenter')
    exCp = f(bold=True,bg_color=YELBG,border=1,align='center',valign='vcenter',num_format='0%')
    exL  = f(bold=True,bg_color=YELBG,border=1,align='left',valign='vcenter')
    tot  = f(bold=True,bg_color=TOTBG,border=1,align='center',valign='vcenter')
    totP = f(bold=True,bg_color=TOTBG,border=1,align='center',valign='vcenter',num_format='0%')
    cfgF = f(bold=True,bg_color=LTBLU,border=1,align='center',valign='vcenter')
    lblF = f(bold=True,align='right',valign='vcenter')
    noteF= f(bold=True,italic=True)

    # ── Column widths (matching template) ────────────────────
    ws.set_column('A:A', 3);   ws.set_column('B:B', 5.5)
    ws.set_column('C:C', 30);  ws.set_column('D:D', 11)
    ws.set_column('E:E', 10);  ws.set_column('F:F', 12)
    ws.set_column('G:G', 12);  ws.set_column('H:H', 12)
    ws.set_column('I:I', 13);  ws.set_column('J:J', 14)
    ws.set_column('K:K', 36)

    # ── Row 2: MFN config (matching template layout) ─────────
    ws.set_row(1, 20)
    ws.write(1,3,'MFA',       lblF); ws.write(1,4, mfn,       cfgF)
    ws.write(1,5,'GFRS Size', lblF); ws.write(1,6, gfrs_size, cfgF)

    # ── Header row 4 (index 3) ───────────────────────────────
    ws.set_row(3, 29)
    ws.write(3,1, 'S.N',                                     hDk)
    ws.write(3,2, 'Client Name',                              hDkL)
    ws.write(3,3, 'Size in Sqcm',                             hDk)
    ws.write(3,4, 'Amnt',                                     hDk)
    ws.write(3,5, 'Net Yeild',                                hBl)
    ws.write(3,6, 'Avg Yield',                                hBl)
    ws.write(3,7, 'Variation',                                hBl)
    ws.write(3,8, 'Space Allocation',                         hBl)
    ws.write(3,9, 'MFN Yeild Variation',                      hBl)
    ws.write(3,10,'Reason where Variation is more than 30%',  hDkL)

    # ── Parse client rows (skip blank, skip exempt) ──────────
    def pr(r):
        return {'name': str(r.get('name','') or '').strip(),
                'size': safeN(r.get('size',0)),
                'amnt': safeN(r.get('amnt',0)),
                'reason': str(r.get('reason','') or '')}

    all_parsed = [pr(r) for r in rows_data if str(r.get('name','') or '').strip()]

    # Client rows (non-exempt) for avg yield calc
    client_rows = [r for r in all_parsed if not isExempt(r['name']) and r['size']>0]
    tSz_clients = sum(r['size'] for r in client_rows)
    tAm_clients = sum(r['amnt'] for r in client_rows)

    # Total size = all named rows + GFRS row
    tSz_all = tSz_clients + (gfrs_size if gfrs_size > 0 else 0)
    tAm_all = tAm_clients + (gfrs_amnt if gfrs_amnt > 0 else 0)

    # Avg Yield = Total Amount / Total Size (all rows incl GFRS)
    avg = tAm_all / tSz_all if tSz_all > 0 else 0

    # MFN formula (from template):
    # Actual Sold Volume (G2) = Total Size (D20) - GFRS Size (D19)
    # MFN Benchmark (J20)     = MFN Value (E2) / Actual Sold Volume (G2)
    actual_sold = tSz_all - gfrs_size if tSz_all > gfrs_size else 0
    mfn_rate    = mfn / actual_sold if actual_sold > 0 else 0

    # ── Write client data rows ───────────────────────────────
    for i, row in enumerate(all_parsed):
        xr  = 4 + i
        sz  = row['size']; am = row['amnt']; name = row['name']
        ex  = isExempt(name)
        net = am / sz if sz > 0 else 0
        # Variation = Net Yield / Avg Yield
        vari = net / avg if avg > 0 and sz > 0 and not ex else 0
        sa   = sz / tSz_all if tSz_all > 0 and sz > 0 else 0
        # MFN Yield Variation = Net Yield / MFN Benchmark  (formula: =F/J20)
        mfnV = net / mfn_rate if mfn_rate > 0 and sz > 0 and not ex else 0
        vp   = vari * 100; sp = sa * 100; mp = mfnV * 100

        if ex:
            ws.write(xr,1,i+1,  exC)
            ws.write(xr,2,name, exL)
            ws.write(xr,3,sz,   exC)
            ws.write(xr,4,am,   exC)
            ws.write(xr,5,net if sz>0 else 0, exC)
            ws.write(xr,6,avg if sz>0 else 0, exC)
            ws.write(xr,7,0,    exC)
            ws.write(xr,8,sa if sz>0 else 0,  exCp)
            ws.write(xr,9,0,    exC)
            ws.write(xr,10,row['reason'], dLft)
            continue

        ws.write(xr,1,i+1,  dCtr)
        ws.write(xr,2,name, dLft)
        ws.write(xr,3,sz,   dCtr)
        ws.write(xr,4,am,   dCtr)
        ws.write(xr,5,round(net,0) if sz>0 else 0, dCtr)
        ws.write(xr,6,round(avg,0) if sz>0 else 0, dCtr)

        vfmt = vOk if vp>=100 else vBd
        if sz==0: vfmt=vBd
        ws.write(xr,7, vari if sz>0 else 0, vfmt)

        sfmt = saBd if sp>25 else (saWn if sp>=20 else saOk)
        ws.write(xr,8, sa if sz>0 else 0, sfmt)

        mfmt = mOk if mp>=100 else mBd
        ws.write(xr,9, mfnV if sz>0 else 0, mfmt)
        ws.write(xr,10,row['reason'], dLft)

    # ── GFRS row at bottom (if provided) ─────────────────────
    gfrs_row_idx = 4 + len(all_parsed)
    if gfrs_size > 0:
        gNet = gfrs_amnt / gfrs_size if gfrs_size > 0 else 0
        gSa  = gfrs_size / tSz_all if tSz_all > 0 else 0
        gSp  = gSa * 100
        sfmt = saBd if gSp>25 else (saWn if gSp>=20 else saOk)
        # Use yellow (exempt) formatting
        ws.write(gfrs_row_idx,1, len(all_parsed)+1, exC)
        ws.write(gfrs_row_idx,2, 'Generic/Filler/Research/Survey', exL)
        ws.write(gfrs_row_idx,3, gfrs_size, exC)
        ws.write(gfrs_row_idx,4, gfrs_amnt, exC)
        ws.write(gfrs_row_idx,5, round(gNet,0), exC)
        ws.write(gfrs_row_idx,6, round(avg,0),  exC)
        ws.write(gfrs_row_idx,7, 0,   exC)
        ws.write(gfrs_row_idx,8, gSa, exCp)
        ws.write(gfrs_row_idx,9, 0,   exC)
        ws.write(gfrs_row_idx,10,'',  dLft)
        totals_row = gfrs_row_idx + 1
    else:
        totals_row = gfrs_row_idx

    # ── Totals row ───────────────────────────────────────────
    tr = totals_row
    for c in range(1,11): ws.write(tr,c,'',tot)
    ws.write(tr,3, tSz_all,         tot)
    ws.write(tr,4, tAm_all,         tot)
    ws.write(tr,5, round(avg,0),    tot)
    ws.write(tr,6, round(avg,0),    tot)
    ws.write(tr,9, round(mfn_rate,0) if mfn_rate>0 else 0, tot)  # J20 = MFN Benchmark

    # ── Notes ────────────────────────────────────────────────
    ws.write(tr+1,2,'Note 1: All Variation requires approval from National Sales Head.', noteF)
    ws.write(tr+2,2,'Note 2: Space Allocation above 25% requires Sameer Sir approval.', noteF)

    wb.close(); out.seek(0)
    fname = f'Variation_Sheet_{date.today()}.xlsx'
    try: write_log(all_parsed, tSz_all, tAm_all, avg, mfn, gfrs_size, actual_sold, mfn_rate, fname)
    except Exception as e: print(f'[LOG] {e}')

    return send_file(out,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=fname)

# ── Log viewer ───────────────────────────────────────────────
@app.route('/logs')
def logs():
    if request.args.get('key','') != LOG_KEY: return '<h2>401</h2>',401
    ensure_log()
    with open(LOG,'r',encoding='utf-8') as f: rows_raw=list(csv.DictReader(f))
    total=len(rows_raw); rows_raw.reverse()
    key=request.args.get('key','')
    tbody=''
    for r in rows_raw:
        try: cl='; '.join(f"{c['name']} sz:{int(float(c['size']))} am:{int(float(c['amnt']))}"
                          for c in json.loads(r.get('clients_json','[]')))
        except: cl='—'
        tbody+=f"<tr><td>{r.get('timestamp_ist','')}</td><td>{r.get('ip','')}</td><td>{r.get('mfn','')}</td><td>{r.get('gfrs_size','')}</td><td>{r.get('actual_sold','')}</td><td>{r.get('mfn_benchmark','')}</td><td>{r.get('clients','')}</td><td>{r.get('total_size','')}</td><td>{r.get('total_amnt','')}</td><td>{r.get('avg_yield','')}</td><td style='font-size:.75rem'>{cl}</td></tr>"
    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Download Log</title>
<style>body{{font-family:Arial;background:#f0f4f8;padding:24px}}h1{{color:#1F3864;margin-bottom:16px}}
table{{border-collapse:collapse;width:100%;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1);font-size:.8rem}}
th{{background:#1F3864;color:#fff;padding:8px 10px;text-align:left}}
td{{padding:6px 10px;border-bottom:1px solid #eee}}tr:nth-child(even){{background:#f5f8ff}}
a{{display:inline-block;margin-bottom:16px;padding:8px 18px;background:#1F5C1F;color:#fff;border-radius:6px;text-decoration:none;font-weight:700}}</style>
</head><body><h1>&#x1F4CB; Download Log — {total} entries</h1>
<a href="/logs/download?key={key}">&#x2B07; Download CSV</a>
<table><thead><tr><th>Timestamp</th><th>IP</th><th>MFN</th><th>GFRS Size</th><th>Actual Sold</th><th>MFN Benchmark</th><th>Clients</th><th>Total Size</th><th>Total Amt</th><th>Avg Yield</th><th>Details</th></tr></thead>
<tbody>{tbody}</tbody></table></body></html>"""

@app.route('/logs/download')
def logs_dl():
    if request.args.get('key','') != LOG_KEY: return '401',401
    ensure_log()
    return send_file(LOG, mimetype='text/csv', as_attachment=True, download_name=f'log_{date.today()}.csv')

if __name__ == '__main__':
    port = int(os.environ.get('PORT',5000))
    app.run(host='0.0.0.0', port=port, debug=False)
