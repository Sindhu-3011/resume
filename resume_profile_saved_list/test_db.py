import psycopg2
from psycopg2.extras import RealDictCursor
import os
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

url = os.environ.get("DATABASE_URL", "postgresql://postgres:postgres@localhost:5432/resume_profiles")
conn = psycopg2.connect(url)
cur = conn.cursor(cursor_factory=RealDictCursor)
cur.execute("SELECT id, full_name, title, email FROM resume ORDER BY id")
rows = cur.fetchall()
print(f"Connected to PostgreSQL OK. Found {len(rows)} resume(s):")
for r in rows:
    print(f"  [{r['id']}] {r['full_name']} | {r['title']} | {r['email']}")
cur.execute("SELECT COUNT(*) AS cnt FROM resume_skill")
skill_count = cur.fetchone()["cnt"]
print(f"  resume_skill rows: {skill_count}")
conn.close()
