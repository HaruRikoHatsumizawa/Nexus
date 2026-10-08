DROP TABLE IF EXISTS energy_readings;

CREATE TABLE IF NOT EXISTS uploaded_spreadsheets (
    id TEXT PRIMARY KEY,
    filename TEXT NOT NULL,
    size_bytes INTEGER NOT NULL,
    sha256 TEXT NOT NULL,
    uploaded_at TEXT NOT NULL,
    chunk_count INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_uploaded_spreadsheets_uploaded_at
ON uploaded_spreadsheets(uploaded_at);

CREATE TABLE IF NOT EXISTS uploaded_spreadsheet_chunks (
    spreadsheet_id TEXT NOT NULL,
    chunk_index INTEGER NOT NULL,
    data_b64 TEXT NOT NULL,
    PRIMARY KEY (spreadsheet_id, chunk_index),
    FOREIGN KEY (spreadsheet_id) REFERENCES uploaded_spreadsheets(id) ON DELETE CASCADE
);