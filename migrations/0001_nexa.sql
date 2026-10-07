CREATE TABLE IF NOT EXISTS sessoes (
    usuario TEXT PRIMARY KEY,
    token TEXT NOT NULL,
    atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS energia_leituras (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    dispositivo TEXT NOT NULL,
    energia_kwh REAL NOT NULL,
    custo_brl REAL NOT NULL DEFAULT 0,
    potencia_w REAL NOT NULL DEFAULT 0,
    tensao_v REAL NOT NULL DEFAULT 0,
    corrente_a REAL NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_energia_timestamp
ON energia_leituras(timestamp);

CREATE INDEX IF NOT EXISTS idx_energia_dispositivo
ON energia_leituras(dispositivo);
