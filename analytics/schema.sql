-- Banco D1 "autoria-analytics": um registro por evento da página.
-- Sem IP, sem user agent, sem identificador de pessoa (LGPD: só o necessário para contar).
CREATE TABLE IF NOT EXISTS eventos (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL,            -- horário UTC, ISO 8601
  dia TEXT NOT NULL,           -- data em Brasília (AAAA-MM-DD)
  tipo TEXT NOT NULL,          -- visita | viu_precos | calc_inicio | calc_enviar | whats_direto | copiar_numero | exemplo
  origem TEXT NOT NULL,        -- whatsapp | google | indicacao-* | direto | ... (ver classificar() no worker)
  ref_host TEXT,               -- domínio de quem mandou (sem caminho), quando houver
  detalhe TEXT,                -- só em calc_enviar: serviço|faculdade|faixa de páginas
  dispositivo TEXT,            -- celular | computador
  pais TEXT                    -- código do país da Cloudflare
);
CREATE INDEX IF NOT EXISTS idx_eventos_dia ON eventos (dia);
