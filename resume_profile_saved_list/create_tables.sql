-- Create all tables for resume_profiles database
-- Run this in DBeaver connected to resume_profiles

CREATE TABLE IF NOT EXISTS schema_version (
    v INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS resume (
    id              SERIAL PRIMARY KEY,
    full_name       TEXT,
    title           TEXT,
    email           TEXT,
    phone           TEXT,
    linkedin        TEXT DEFAULT '',
    location        TEXT,
    summary         TEXT,
    skills          TEXT,
    experience      TEXT,
    education       TEXT,
    certifications  TEXT,
    projects        TEXT,
    slug            TEXT,
    resume_file     TEXT,
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    updated_at      TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS resume_skill (
    id          SERIAL PRIMARY KEY,
    resume_id   INTEGER NOT NULL REFERENCES resume(id) ON DELETE CASCADE,
    skill       TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_resume_skill_resume_id ON resume_skill(resume_id);

INSERT INTO schema_version (v) VALUES (4);

-- Verify tables were created
SELECT table_name FROM information_schema.tables
WHERE table_schema = 'public'
ORDER BY table_name;
