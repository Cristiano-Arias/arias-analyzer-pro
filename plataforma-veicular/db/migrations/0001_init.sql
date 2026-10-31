-- =============================================================================
-- Garagem360 — Migração 0001: esquema inicial do MVP
-- PostgreSQL 16+
--
-- Princípios (ver docs/arquitetura/modelo-de-dados.md):
--  * Todo dado de negócio tem account_id (conta pessoal ou organização) -> RLS.
--  * Tabelas de referência (categorias, propulsões, energias) são LINHAS, não
--    enums: nova categoria = INSERT, sem alterar estrutura.
--  * Tabelas sincronizáveis têm: version, updated_at, deleted_at (exclusão
--    lógica) e server_seq (cursor monotônico global para /sync/pull).
--  * Dinheiro em centavos (bigint). Quantidades físicas em numeric.
--  * Medições têm procedência: source + confidence.
-- =============================================================================

BEGIN;

CREATE EXTENSION IF NOT EXISTS citext;
CREATE EXTENSION IF NOT EXISTS pgcrypto; -- gen_random_uuid()

CREATE SCHEMA IF NOT EXISTS app;

-- -----------------------------------------------------------------------------
-- Contexto de segurança da requisição (definido pela API em cada transação):
--   SET LOCAL app.user_id = '<uuid>';
--   SET LOCAL app.account_ids = '{<uuid>,<uuid>}';
--   SET LOCAL app.shared_vehicle_ids = '{<uuid>}';
-- Valores ausentes => nenhum acesso (fail-closed).
-- -----------------------------------------------------------------------------
CREATE FUNCTION app.current_user_id() RETURNS uuid
LANGUAGE sql STABLE AS $$
  SELECT NULLIF(current_setting('app.user_id', true), '')::uuid
$$;

CREATE FUNCTION app.current_account_ids() RETURNS uuid[]
LANGUAGE sql STABLE AS $$
  SELECT COALESCE(NULLIF(current_setting('app.account_ids', true), '')::uuid[], '{}')
$$;

CREATE FUNCTION app.current_shared_vehicle_ids() RETURNS uuid[]
LANGUAGE sql STABLE AS $$
  SELECT COALESCE(NULLIF(current_setting('app.shared_vehicle_ids', true), '')::uuid[], '{}')
$$;

-- Sequência global usada como cursor de sincronização.
CREATE SEQUENCE app.sync_seq AS bigint;

-- Trigger: mantém version/updated_at/server_seq em toda escrita.
CREATE FUNCTION app.touch_sync() RETURNS trigger
LANGUAGE plpgsql AS $$
BEGIN
  IF TG_OP = 'INSERT' THEN
    NEW.version := 1;
    NEW.created_at := COALESCE(NEW.created_at, now());
  ELSE
    NEW.version := OLD.version + 1;
    NEW.created_at := OLD.created_at;          -- imutável
  END IF;
  NEW.updated_at := now();
  NEW.server_seq := nextval('app.sync_seq');
  RETURN NEW;
END $$;

-- =============================================================================
-- 1. Identidade, contas e acesso
-- =============================================================================
CREATE TABLE users (
  id                 uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  email              citext NOT NULL,
  password_hash      text,                     -- NULL se apenas login social
  full_name          text NOT NULL CHECK (length(full_name) BETWEEN 1 AND 160),
  email_verified_at  timestamptz,
  locale             text NOT NULL DEFAULT 'pt-BR',
  timezone           text NOT NULL DEFAULT 'America/Sao_Paulo',
  created_at         timestamptz NOT NULL DEFAULT now(),
  updated_at         timestamptz NOT NULL DEFAULT now(),
  deleted_at         timestamptz
);
CREATE UNIQUE INDEX users_email_active_uq ON users (email) WHERE deleted_at IS NULL;

CREATE TABLE user_identities (               -- login social (OIDC)
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id     uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  provider    text NOT NULL CHECK (provider IN ('google','apple')),
  subject     text NOT NULL,
  created_at  timestamptz NOT NULL DEFAULT now(),
  UNIQUE (provider, subject)
);

CREATE TABLE auth_sessions (                 -- refresh tokens rotativos
  id               uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id          uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  family_id        uuid NOT NULL,              -- reuso detectado => revoga a família
  token_hash       bytea NOT NULL UNIQUE,      -- SHA-256 do token; nunca o token puro
  device_label     text,
  user_agent       text,
  ip               inet,
  created_at       timestamptz NOT NULL DEFAULT now(),
  expires_at       timestamptz NOT NULL,
  rotated_at       timestamptz,
  revoked_at       timestamptz
);
CREATE INDEX auth_sessions_user_idx ON auth_sessions (user_id) WHERE revoked_at IS NULL;
CREATE INDEX auth_sessions_family_idx ON auth_sessions (family_id);

CREATE TABLE one_time_tokens (               -- verificação de e-mail / reset de senha
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id     uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  purpose     text NOT NULL CHECK (purpose IN ('verify_email','reset_password')),
  token_hash  bytea NOT NULL UNIQUE,
  expires_at  timestamptz NOT NULL,
  used_at     timestamptz,
  created_at  timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE consents (                      -- LGPD: consentimento versionado
  id             uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id        uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  purpose        text NOT NULL,              -- 'terms','privacy','marketing','ai_processing'
  document_version text NOT NULL,
  granted_at     timestamptz NOT NULL DEFAULT now(),
  revoked_at     timestamptz,
  ip             inet
);
CREATE INDEX consents_user_idx ON consents (user_id, purpose);

CREATE TABLE plans (                         -- limites por plano; SEM preços (D-02)
  code        text PRIMARY KEY,
  name        text NOT NULL,
  limits      jsonb NOT NULL DEFAULT '{}'::jsonb
);

CREATE TABLE accounts (
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  kind        text NOT NULL CHECK (kind IN ('personal','organization')),
  name        text NOT NULL CHECK (length(name) BETWEEN 1 AND 160),
  cnpj        text CHECK (cnpj IS NULL OR cnpj ~ '^\d{14}$'),
  plan_code   text NOT NULL DEFAULT 'free' REFERENCES plans(code),
  created_at  timestamptz NOT NULL DEFAULT now(),
  updated_at  timestamptz NOT NULL DEFAULT now(),
  deleted_at  timestamptz,
  CHECK (kind = 'organization' OR cnpj IS NULL)
);

CREATE TABLE account_members (
  account_id  uuid NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
  user_id     uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  role        text NOT NULL CHECK (role IN ('owner','admin','manager','driver','viewer')),
  created_at  timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (account_id, user_id)
);
CREATE INDEX account_members_user_idx ON account_members (user_id);

CREATE TABLE subscriptions (
  id                  uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id          uuid NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
  plan_code           text NOT NULL REFERENCES plans(code),
  status              text NOT NULL CHECK (status IN ('trialing','active','past_due','canceled')),
  provider            text,                   -- ex.: 'app_store','play','stripe' (futuro)
  provider_ref        text,
  current_period_end  timestamptz,
  created_at          timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE devices (                       -- push
  id           uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id      uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  platform     text NOT NULL CHECK (platform IN ('ios','android','web')),
  push_token   text NOT NULL UNIQUE,
  last_seen_at timestamptz NOT NULL DEFAULT now()
);

-- =============================================================================
-- 2. Tabelas de referência (extensíveis por INSERT)
-- =============================================================================
CREATE TABLE ref_vehicle_categories (
  code         text PRIMARY KEY,
  name_pt      text NOT NULL,
  is_powered   boolean NOT NULL DEFAULT true,   -- reboque/semirreboque = false
  is_heavy     boolean NOT NULL DEFAULT false,
  active       boolean NOT NULL DEFAULT true
);

CREATE TABLE ref_propulsion_types (
  code              text PRIMARY KEY,
  name_pt           text NOT NULL,
  has_combustion    boolean NOT NULL,
  has_traction_battery boolean NOT NULL,
  is_plug_in        boolean NOT NULL,
  active            boolean NOT NULL DEFAULT true
);

CREATE TABLE ref_energy_sources (
  code        text PRIMARY KEY,
  name_pt     text NOT NULL,
  unit        text NOT NULL CHECK (unit IN ('L','m3','kg','kWh'))
);

CREATE TABLE ref_expense_categories (
  code     text PRIMARY KEY,
  name_pt  text NOT NULL
);

-- Catálogo opcional (origem a definir — risco R-02). Texto livre é sempre aceito.
CREATE TABLE catalog_makes (
  id    uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name  citext NOT NULL UNIQUE
);
CREATE TABLE catalog_models (
  id       uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  make_id  uuid NOT NULL REFERENCES catalog_makes(id),
  name     citext NOT NULL,
  category_code text REFERENCES ref_vehicle_categories(code),
  UNIQUE (make_id, name)
);
CREATE TABLE catalog_versions (
  id               uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  model_id         uuid NOT NULL REFERENCES catalog_models(id),
  name             citext NOT NULL,
  year_from        smallint,
  year_to          smallint,
  propulsion_code  text REFERENCES ref_propulsion_types(code),
  specs            jsonb NOT NULL DEFAULT '{}'::jsonb,
  data_source      text NOT NULL,             -- procedência obrigatória
  UNIQUE (model_id, name, year_from)
);

-- =============================================================================
-- 3. Garagem
-- =============================================================================
CREATE TABLE vehicles (
  id                   uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id           uuid NOT NULL REFERENCES accounts(id),
  category_code        text NOT NULL REFERENCES ref_vehicle_categories(code),
  body_type            text,
  propulsion_code      text REFERENCES ref_propulsion_types(code), -- NULL p/ não motorizado
  make_name            text NOT NULL,
  model_name           text NOT NULL,
  version_name         text,
  catalog_version_id   uuid REFERENCES catalog_versions(id),
  year_manufacture     smallint CHECK (year_manufacture BETWEEN 1900 AND 2100),
  year_model           smallint CHECK (year_model BETWEEN 1900 AND 2101),
  plate                text CHECK (plate IS NULL OR plate ~ '^[A-Z]{3}[0-9][A-Z0-9][0-9]{2}$'),
  vin                  text CHECK (vin IS NULL OR vin ~ '^[A-HJ-NPR-Z0-9]{17}$'),
  renavam              text CHECK (renavam IS NULL OR renavam ~ '^\d{11}$'),
  nickname             text,
  color                text,
  transmission         text CHECK (transmission IS NULL OR transmission IN ('manual','automatic','cvt','dct','amt','single_speed','other')),
  engine_description   text,
  fuel_tank_liters     numeric(7,2) CHECK (fuel_tank_liters IS NULL OR fuel_tank_liters > 0),
  battery_gross_kwh    numeric(7,2) CHECK (battery_gross_kwh IS NULL OR battery_gross_kwh > 0),
  battery_usable_kwh   numeric(7,2) CHECK (battery_usable_kwh IS NULL OR battery_usable_kwh > 0),
  odometer_km          integer NOT NULL DEFAULT 0 CHECK (odometer_km >= 0),
  acquired_on          date,
  purchase_price_cents bigint CHECK (purchase_price_cents IS NULL OR purchase_price_cents >= 0),
  status               text NOT NULL DEFAULT 'active' CHECK (status IN ('active','sold','archived')),
  attributes           jsonb NOT NULL DEFAULT '{}'::jsonb,  -- campos específicos por categoria (ex.: eixos)
  version              integer NOT NULL DEFAULT 1,
  server_seq           bigint NOT NULL DEFAULT 0,
  created_at           timestamptz NOT NULL DEFAULT now(),
  updated_at           timestamptz NOT NULL DEFAULT now(),
  deleted_at           timestamptz,
  CHECK (battery_usable_kwh IS NULL OR battery_gross_kwh IS NULL OR battery_usable_kwh <= battery_gross_kwh),
  CHECK (year_model IS NULL OR year_manufacture IS NULL OR year_model BETWEEN year_manufacture AND year_manufacture + 1)
);
CREATE INDEX vehicles_account_idx ON vehicles (account_id) WHERE deleted_at IS NULL;
CREATE INDEX vehicles_sync_idx ON vehicles (account_id, server_seq);
CREATE UNIQUE INDEX vehicles_plate_account_uq ON vehicles (account_id, plate) WHERE plate IS NOT NULL AND deleted_at IS NULL;

-- Combustíveis/energias aceitos pelo veículo (flex = gasolina + etanol; kit GNV etc.)
CREATE TABLE vehicle_energy_sources (
  vehicle_id   uuid NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE,
  energy_code  text NOT NULL REFERENCES ref_energy_sources(code),
  PRIMARY KEY (vehicle_id, energy_code)
);

CREATE TABLE vehicle_shares (                -- compartilhamento com familiares
  vehicle_id  uuid NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE,
  user_id     uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  role        text NOT NULL CHECK (role IN ('editor','viewer')),
  granted_by  uuid NOT NULL REFERENCES users(id),
  created_at  timestamptz NOT NULL DEFAULT now(),
  revoked_at  timestamptz,
  PRIMARY KEY (vehicle_id, user_id)
);

-- Macro reutilizada: colunas de sincronização em tabelas de negócio.
-- (Repetidas explicitamente por clareza e compatibilidade com ferramentas.)

CREATE TABLE odometer_readings (
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id  uuid NOT NULL REFERENCES accounts(id),
  vehicle_id  uuid NOT NULL REFERENCES vehicles(id),
  reading_km  integer NOT NULL CHECK (reading_km >= 0),
  read_at     timestamptz NOT NULL,
  source      text NOT NULL DEFAULT 'user' CHECK (source IN ('user','ocr','obd','manufacturer_api','telematics','derived')),
  confidence  text NOT NULL DEFAULT 'confirmed' CHECK (confidence IN ('confirmed','estimated')),
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX odometer_vehicle_idx ON odometer_readings (vehicle_id, read_at);
CREATE INDEX odometer_sync_idx ON odometer_readings (account_id, server_seq);

-- =============================================================================
-- 4. Arquivos (bucket privado; metadados aqui)
-- =============================================================================
CREATE TABLE files (
  id             uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id     uuid NOT NULL REFERENCES accounts(id),
  storage_key    text NOT NULL UNIQUE,         -- nunca exposto como URL pública
  original_name  text NOT NULL,
  mime_type      text NOT NULL CHECK (mime_type IN ('application/pdf','image/jpeg','image/png','image/heic','image/webp')),
  size_bytes     bigint NOT NULL CHECK (size_bytes > 0 AND size_bytes <= 26214400), -- 25 MiB
  sha256         bytea,
  status         text NOT NULL DEFAULT 'pending' CHECK (status IN ('pending','ready','rejected')),
  uploaded_by    uuid NOT NULL REFERENCES users(id),
  -- vínculo polimórfico simples (validado na API)
  entity_type    text CHECK (entity_type IN ('vehicle','maintenance','document','expense','fuel_log','charging')),
  entity_id      uuid,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX files_entity_idx ON files (entity_type, entity_id);
CREATE INDEX files_sync_idx ON files (account_id, server_seq);

-- =============================================================================
-- 5. Manutenção
-- =============================================================================
CREATE TABLE workshops (
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id  uuid NOT NULL REFERENCES accounts(id),
  name        text NOT NULL,
  cnpj        text CHECK (cnpj IS NULL OR cnpj ~ '^\d{14}$'),
  phone       text,
  address     text,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX workshops_sync_idx ON workshops (account_id, server_seq);

CREATE TABLE maintenance_plan_items (
  id               uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id       uuid NOT NULL REFERENCES accounts(id),
  vehicle_id       uuid NOT NULL REFERENCES vehicles(id),
  category         text NOT NULL,             -- ver domain/maintenance.ts
  title            text NOT NULL,
  interval_km      integer CHECK (interval_km IS NULL OR interval_km > 0),
  interval_months  smallint CHECK (interval_months IS NULL OR interval_months > 0),
  interval_source  text NOT NULL CHECK (interval_source IN ('manufacturer_manual','user','generic_template')),
  source_note      text,                      -- ex.: "Manual do proprietário, p. 212"
  last_done_on     date,
  last_done_km     integer CHECK (last_done_km IS NULL OR last_done_km >= 0),
  remind_before_km integer NOT NULL DEFAULT 1000 CHECK (remind_before_km >= 0),
  remind_before_days smallint NOT NULL DEFAULT 15 CHECK (remind_before_days >= 0),
  active           boolean NOT NULL DEFAULT true,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz,
  CHECK (interval_km IS NOT NULL OR interval_months IS NOT NULL)
);
CREATE INDEX mpi_vehicle_idx ON maintenance_plan_items (vehicle_id) WHERE deleted_at IS NULL;
CREATE INDEX mpi_sync_idx ON maintenance_plan_items (account_id, server_seq);

CREATE TABLE maintenance_records (
  id               uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id       uuid NOT NULL REFERENCES accounts(id),
  vehicle_id       uuid NOT NULL REFERENCES vehicles(id),
  kind             text NOT NULL CHECK (kind IN ('preventive','corrective','inspection','recall','other')),
  category         text NOT NULL,
  description      text NOT NULL,
  performed_on     date NOT NULL,
  odometer_km      integer CHECK (odometer_km IS NULL OR odometer_km >= 0),
  workshop_id      uuid REFERENCES workshops(id),
  technician_name  text,
  parts_cents      bigint NOT NULL DEFAULT 0 CHECK (parts_cents >= 0),
  labor_cents      bigint NOT NULL DEFAULT 0 CHECK (labor_cents >= 0),
  total_cents      bigint NOT NULL CHECK (total_cents >= 0),
  warranty_until   date,
  warranty_km      integer CHECK (warranty_km IS NULL OR warranty_km > 0),
  notes            text,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX mr_vehicle_idx ON maintenance_records (vehicle_id, performed_on DESC);
CREATE INDEX mr_sync_idx ON maintenance_records (account_id, server_seq);

CREATE TABLE maintenance_items (
  id               uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id       uuid NOT NULL REFERENCES accounts(id),
  record_id        uuid NOT NULL REFERENCES maintenance_records(id) ON DELETE CASCADE,
  plan_item_id     uuid REFERENCES maintenance_plan_items(id),  -- atualiza "último feito"
  item_type        text NOT NULL CHECK (item_type IN ('part','service','fluid')),
  description      text NOT NULL,
  part_number      text,
  brand            text,
  quantity         numeric(10,3) NOT NULL DEFAULT 1 CHECK (quantity > 0),
  unit_price_cents bigint NOT NULL DEFAULT 0 CHECK (unit_price_cents >= 0),
  total_cents      bigint NOT NULL CHECK (total_cents >= 0),
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX mi_record_idx ON maintenance_items (record_id);
CREATE INDEX mi_sync_idx ON maintenance_items (account_id, server_seq);

-- =============================================================================
-- 6. Documentos
-- =============================================================================
CREATE TABLE documents (
  id               uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id       uuid NOT NULL REFERENCES accounts(id),
  vehicle_id       uuid REFERENCES vehicles(id),      -- NULL p/ CNH (pessoa)
  doc_type         text NOT NULL CHECK (doc_type IN (
                     'crlv','ipva','licensing','insurance','inspection','fine',
                     'authorization','transport','driver_license','certificate',
                     'warranty','purchase_sale','other')),
  title            text NOT NULL,
  number           text,
  holder_name      text,
  issued_on        date,
  expires_on       date,
  amount_cents     bigint CHECK (amount_cents IS NULL OR amount_cents >= 0),
  payment_status   text NOT NULL DEFAULT 'not_applicable' CHECK (payment_status IN ('not_applicable','pending','paid','overdue')),
  paid_on          date,
  reminder_days    smallint[] NOT NULL DEFAULT '{30,7,1}',
  details          jsonb NOT NULL DEFAULT '{}'::jsonb,  -- ex.: seguradora, apólice, cobertura
  notes            text,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz,
  CHECK (expires_on IS NULL OR issued_on IS NULL OR expires_on >= issued_on),
  CHECK (payment_status <> 'paid' OR paid_on IS NOT NULL)
);
CREATE INDEX documents_expiry_idx ON documents (expires_on) WHERE deleted_at IS NULL AND expires_on IS NOT NULL;
CREATE INDEX documents_sync_idx ON documents (account_id, server_seq);

-- =============================================================================
-- 7. Abastecimento e recarga
-- =============================================================================
CREATE TABLE fuel_logs (
  id                uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id        uuid NOT NULL REFERENCES accounts(id),
  vehicle_id        uuid NOT NULL REFERENCES vehicles(id),
  filled_at         timestamptz NOT NULL,
  odometer_km       integer NOT NULL CHECK (odometer_km >= 0),
  energy_code       text NOT NULL REFERENCES ref_energy_sources(code),
  quantity          numeric(9,3) NOT NULL CHECK (quantity > 0),   -- L, m³ ou kg conforme energia
  unit_price_milli  integer CHECK (unit_price_milli IS NULL OR unit_price_milli > 0), -- R$/unid. em milésimos (5,879 => 5879)
  total_cents       bigint NOT NULL CHECK (total_cents >= 0),
  is_full           boolean NOT NULL DEFAULT true,
  missed_previous   boolean NOT NULL DEFAULT false,  -- usuário esqueceu de registrar o anterior
  station_name      text,
  notes             text,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX fuel_vehicle_idx ON fuel_logs (vehicle_id, odometer_km);
CREATE INDEX fuel_sync_idx ON fuel_logs (account_id, server_seq);

CREATE TABLE charging_sessions (
  id               uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id       uuid NOT NULL REFERENCES accounts(id),
  vehicle_id       uuid NOT NULL REFERENCES vehicles(id),
  started_at       timestamptz NOT NULL,
  ended_at         timestamptz,
  odometer_km      integer CHECK (odometer_km IS NULL OR odometer_km >= 0),
  energy_kwh       numeric(8,3) NOT NULL CHECK (energy_kwh > 0),
  energy_measured_at text NOT NULL DEFAULT 'unknown' CHECK (energy_measured_at IN ('charger','vehicle','unknown')),
  location_type    text NOT NULL CHECK (location_type IN ('home','public','work','other')),
  connector        text CHECK (connector IS NULL OR connector IN ('type2','ccs2','chademo','gbt','tesla','schuko','nbr14136','other')),
  power_kw         numeric(6,1) CHECK (power_kw IS NULL OR power_kw > 0),
  soc_start_pct    smallint CHECK (soc_start_pct BETWEEN 0 AND 100),
  soc_end_pct      smallint CHECK (soc_end_pct BETWEEN 0 AND 100),
  cost_cents       bigint NOT NULL DEFAULT 0 CHECK (cost_cents >= 0),
  provider_name    text,
  source           text NOT NULL DEFAULT 'user' CHECK (source IN ('user','ocr','charger_api','manufacturer_api','telematics')),
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz,
  CHECK (ended_at IS NULL OR ended_at >= started_at)
);
CREATE INDEX charging_vehicle_idx ON charging_sessions (vehicle_id, started_at);
CREATE INDEX charging_sync_idx ON charging_sessions (account_id, server_seq);

CREATE TABLE battery_health_readings (       -- preparado p/ Fase 5 (SOH); sempre com procedência
  id           uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id   uuid NOT NULL REFERENCES accounts(id),
  vehicle_id   uuid NOT NULL REFERENCES vehicles(id),
  measured_at  timestamptz NOT NULL,
  soh_pct      numeric(5,2) CHECK (soh_pct BETWEEN 0 AND 100),
  usable_kwh   numeric(7,2),
  source       text NOT NULL CHECK (source IN ('user','service_report','obd','manufacturer_api','telematics')),
  confidence   text NOT NULL CHECK (confidence IN ('confirmed','estimated')),
  notes        text,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX bhr_sync_idx ON battery_health_readings (account_id, server_seq);

-- =============================================================================
-- 8. Financeiro
-- =============================================================================
CREATE TABLE expenses (
  id              uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id      uuid NOT NULL REFERENCES accounts(id),
  vehicle_id      uuid REFERENCES vehicles(id),
  category_code   text NOT NULL REFERENCES ref_expense_categories(code),
  occurred_on     date NOT NULL,
  amount_cents    bigint NOT NULL CHECK (amount_cents >= 0),
  description     text,
  payment_method  text CHECK (payment_method IS NULL OR payment_method IN ('cash','debit','credit','pix','boleto','fleet_card','other')),
  odometer_km     integer CHECK (odometer_km IS NULL OR odometer_km >= 0),
  -- Despesa derivada de outro registro: evita lançamento duplicado (RF-04-2)
  origin_type     text NOT NULL DEFAULT 'manual' CHECK (origin_type IN ('manual','maintenance','fuel','charging','document')),
  origin_id       uuid,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz,
  CHECK ((origin_type = 'manual') = (origin_id IS NULL))
);
CREATE UNIQUE INDEX expenses_origin_uq ON expenses (origin_type, origin_id) WHERE origin_id IS NOT NULL AND deleted_at IS NULL;
CREATE INDEX expenses_report_idx ON expenses (account_id, vehicle_id, occurred_on) WHERE deleted_at IS NULL;
CREATE INDEX expenses_sync_idx ON expenses (account_id, server_seq);

CREATE TABLE incomes (
  id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id    uuid NOT NULL REFERENCES accounts(id),
  vehicle_id    uuid REFERENCES vehicles(id),
  category      text NOT NULL,                -- 'ride','freight','rental','other'
  occurred_on   date NOT NULL,
  amount_cents  bigint NOT NULL CHECK (amount_cents >= 0),
  description   text,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz
);
CREATE INDEX incomes_sync_idx ON incomes (account_id, server_seq);

-- =============================================================================
-- 9. Notificações e auditoria
-- =============================================================================
CREATE TABLE notifications (
  id             uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id     uuid NOT NULL REFERENCES accounts(id),
  user_id        uuid NOT NULL REFERENCES users(id),
  kind           text NOT NULL,               -- 'document_expiry','maintenance_due',...
  title          text NOT NULL,
  body           text NOT NULL,
  entity_type    text,
  entity_id      uuid,
  dedupe_key     text NOT NULL,               -- evita alertas repetidos (ex.: doc:<id>:7d)
  scheduled_for  timestamptz NOT NULL,
  sent_at        timestamptz,
  read_at        timestamptz,
  version integer NOT NULL DEFAULT 1, server_seq bigint NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(), updated_at timestamptz NOT NULL DEFAULT now(), deleted_at timestamptz,
  UNIQUE (user_id, dedupe_key)
);
CREATE INDEX notifications_pending_idx ON notifications (scheduled_for) WHERE sent_at IS NULL;
CREATE INDEX notifications_sync_idx ON notifications (account_id, server_seq);

CREATE TABLE audit_log (                     -- somente inserção
  id           bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  occurred_at  timestamptz NOT NULL DEFAULT now(),
  account_id   uuid,
  user_id      uuid,
  action       text NOT NULL,                 -- 'insert','update','delete','login','export',...
  entity_type  text NOT NULL,
  entity_id    uuid,
  changes      jsonb,
  request_id   text,
  ip           inet
);
CREATE INDEX audit_account_idx ON audit_log (account_id, occurred_at DESC);

CREATE FUNCTION app.audit_row() RETURNS trigger
LANGUAGE plpgsql SECURITY DEFINER SET search_path = public, app AS $$
DECLARE
  rec jsonb := to_jsonb(COALESCE(NEW, OLD));
BEGIN
  INSERT INTO audit_log (account_id, user_id, action, entity_type, entity_id, changes, request_id)
  VALUES (
    (rec->>'account_id')::uuid,
    app.current_user_id(),
    CASE WHEN TG_OP = 'UPDATE' AND NEW.deleted_at IS NOT NULL AND OLD.deleted_at IS NULL
         THEN 'delete' ELSE lower(TG_OP) END,
    TG_TABLE_NAME,
    (rec->>'id')::uuid,
    CASE WHEN TG_OP = 'UPDATE' THEN
      (SELECT jsonb_object_agg(n.key, n.value)
         FROM jsonb_each(to_jsonb(NEW)) n
         JOIN jsonb_each(to_jsonb(OLD)) o USING (key)
        WHERE n.value IS DISTINCT FROM o.value
          AND n.key NOT IN ('version','server_seq','updated_at'))
    ELSE NULL END,
    NULLIF(current_setting('app.request_id', true), '')
  );
  RETURN NULL;
END $$;

CREATE TABLE ai_conversations (              -- estrutura p/ Fase 5
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  account_id  uuid NOT NULL REFERENCES accounts(id),
  user_id     uuid NOT NULL REFERENCES users(id),
  vehicle_id  uuid REFERENCES vehicles(id),
  title       text,
  created_at  timestamptz NOT NULL DEFAULT now(),
  deleted_at  timestamptz
);

-- =============================================================================
-- 10. Triggers de sincronização + auditoria + RLS nas tabelas de negócio
-- =============================================================================
DO $$
DECLARE
  t text;
  synced text[] := ARRAY[
    'vehicles','odometer_readings','files','workshops','maintenance_plan_items',
    'maintenance_records','maintenance_items','documents','fuel_logs',
    'charging_sessions','battery_health_readings','expenses','incomes','notifications'];
BEGIN
  FOREACH t IN ARRAY synced LOOP
    EXECUTE format('CREATE TRIGGER %I BEFORE INSERT OR UPDATE ON %I
                    FOR EACH ROW EXECUTE FUNCTION app.touch_sync()', t || '_touch', t);
    EXECUTE format('CREATE TRIGGER %I AFTER INSERT OR UPDATE OR DELETE ON %I
                    FOR EACH ROW EXECUTE FUNCTION app.audit_row()', t || '_audit', t);
    EXECUTE format('ALTER TABLE %I ENABLE ROW LEVEL SECURITY', t);
    EXECUTE format('ALTER TABLE %I FORCE ROW LEVEL SECURITY', t);
  END LOOP;
END $$;

-- Políticas: acesso pela conta OU por compartilhamento do veículo.
-- Leitura: conta ou compartilhamento (qualquer papel).
-- Escrita: somente conta (MVP). Compartilhamento 'editor' é tratado na API e
-- incluído aqui na Fase de compartilhamento avançado.
CREATE POLICY vehicles_rw ON vehicles
  USING (account_id = ANY (app.current_account_ids()) OR id = ANY (app.current_shared_vehicle_ids()))
  WITH CHECK (account_id = ANY (app.current_account_ids()));

DO $$
DECLARE
  t text;
  by_vehicle text[] := ARRAY['odometer_readings','maintenance_plan_items','maintenance_records',
                             'fuel_logs','charging_sessions','battery_health_readings'];
  by_account text[] := ARRAY['files','workshops','maintenance_items','documents',
                             'expenses','incomes','notifications'];
BEGIN
  FOREACH t IN ARRAY by_vehicle LOOP
    EXECUTE format($p$CREATE POLICY %I ON %I
      USING (account_id = ANY (app.current_account_ids()) OR vehicle_id = ANY (app.current_shared_vehicle_ids()))
      WITH CHECK (account_id = ANY (app.current_account_ids()))$p$, t || '_rw', t);
  END LOOP;
  FOREACH t IN ARRAY by_account LOOP
    EXECUTE format($p$CREATE POLICY %I ON %I
      USING (account_id = ANY (app.current_account_ids()))
      WITH CHECK (account_id = ANY (app.current_account_ids()))$p$, t || '_rw', t);
  END LOOP;
END $$;

-- Notificações: além da conta, só o próprio destinatário.
DROP POLICY notifications_rw ON notifications;
CREATE POLICY notifications_rw ON notifications
  USING (account_id = ANY (app.current_account_ids()) AND user_id = app.current_user_id())
  WITH CHECK (account_id = ANY (app.current_account_ids()));

-- =============================================================================
-- 11. Dados de referência (não fictícios: taxonomias do domínio)
-- =============================================================================
INSERT INTO plans (code, name, limits) VALUES
  ('free',         'Gratuito',     '{"max_vehicles": 1, "ai": false}'),
  ('premium',      'Premium',      '{"max_vehicles": 5, "ai": true}'),
  ('professional', 'Profissional', '{"max_vehicles": 10, "ai": true}'),
  ('enterprise',   'Empresarial',  '{"max_vehicles": null, "ai": true}');
-- Limites provisórios; preços NÃO definidos (decisão D-02).

INSERT INTO ref_vehicle_categories (code, name_pt, is_powered, is_heavy) VALUES
  ('car',            'Automóvel de passeio', true,  false),
  ('motorcycle',     'Motocicleta',          true,  false),
  ('scooter',        'Scooter / ciclomotor', true,  false),
  ('pickup',         'Picape',               true,  false),
  ('van',            'Van / utilitário',     true,  false),
  ('light_commercial','Veículo comercial leve', true, false),
  ('truck_light',    'Caminhão leve',        true,  true),
  ('truck_medium',   'Caminhão médio',       true,  true),
  ('truck_heavy',    'Caminhão pesado',      true,  true),
  ('truck_extra_heavy','Caminhão extrapesado', true, true),
  ('tractor_unit',   'Cavalo mecânico',      true,  true),
  ('bus',            'Ônibus',               true,  true),
  ('minibus',        'Micro-ônibus',         true,  true),
  ('semi_trailer',   'Semirreboque',         false, true),
  ('trailer',        'Reboque',              false, true);

INSERT INTO ref_propulsion_types (code, name_pt, has_combustion, has_traction_battery, is_plug_in) VALUES
  ('ice',   'Combustão interna',            true,  false, false),
  ('mhev',  'Híbrido leve (MHEV)',          true,  true,  false),
  ('hev',   'Híbrido (HEV)',                true,  true,  false),
  ('phev',  'Híbrido plug-in (PHEV)',       true,  true,  true),
  ('erev',  'Elétrico c/ extensor (EREV)',  true,  true,  true),
  ('bev',   'Elétrico a bateria (BEV)',     false, true,  true);

INSERT INTO ref_energy_sources (code, name_pt, unit) VALUES
  ('gasoline',     'Gasolina',          'L'),
  ('gasoline_premium', 'Gasolina aditivada/premium', 'L'),
  ('ethanol',      'Etanol',            'L'),
  ('diesel_s10',   'Diesel S10',        'L'),
  ('diesel_s500',  'Diesel S500',       'L'),
  ('cng',          'GNV',               'm3'),
  ('biomethane',   'Biometano',         'm3'),
  ('arla32',       'Arla 32',           'L'),
  ('electricity',  'Energia elétrica',  'kWh');

INSERT INTO ref_expense_categories (code, name_pt) VALUES
  ('fuel','Combustível'), ('electricity','Energia elétrica'), ('maintenance','Manutenção'),
  ('repair','Reparos'), ('parts','Peças'), ('taxes','Impostos'), ('insurance','Seguro'),
  ('financing','Financiamento'), ('tolls','Pedágios'), ('parking','Estacionamento'),
  ('washing','Lavagem'), ('fines','Multas'), ('tires','Pneus'), ('accessories','Acessórios'),
  ('services','Serviços'), ('other','Outros');

COMMIT;
