# Variation Sheet Web App

## Run Locally
```bash
pip install -r requirements.txt
python app.py
```
Open: http://localhost:5000

---

## View Download Logs

Every download is recorded automatically. To view logs:

**Web viewer:**
```
http://localhost:5000/logs?key=admin123
```

**Download log as CSV:**
```
http://localhost:5000/logs/download?key=admin123
```

Each log entry records:
- Timestamp (IST)
- IP address of downloader
- Number of clients, total size, total amount, avg yield
- Full client-level detail (name, size, amount, remarks)
- Filename downloaded

**Change the default log key** (recommended before hosting):
Set the `LOG_KEY` environment variable on your server:
```
LOG_KEY=mysecretkey python app.py
```
Then access logs at: `http://yoursite.com/logs?key=mysecretkey`

---

## Deploy FREE on Render.com (shareable link)
1. Push this folder to a GitHub repository
2. Go to https://render.com → New → Web Service → connect repo
3. Add environment variable: `LOG_KEY` = your secret key
4. Deploy → get link like: https://variation-sheet.onrender.com
5. View logs at: https://variation-sheet.onrender.com/logs?key=yourkey

## Deploy FREE on Railway.app
1. Go to https://railway.app → New Project → Deploy from GitHub
2. Add env var: `LOG_KEY` = your secret key
3. Get your link instantly

## Log file location
`download_log.csv` — created automatically in the same folder as app.py.
