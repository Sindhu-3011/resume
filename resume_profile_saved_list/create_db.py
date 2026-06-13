import psycopg2

conn = psycopg2.connect("postgresql://postgres:postgres@localhost:5432/postgres")
conn.autocommit = True
cur = conn.cursor()
cur.execute("SELECT 1 FROM pg_database WHERE datname = 'resume_profiles'")
exists = cur.fetchone()
print("DB exists:", bool(exists))
if not exists:
    cur.execute("CREATE DATABASE resume_profiles")
    print("Created resume_profiles database")
else:
    print("Database already exists, nothing to do")
conn.close()
