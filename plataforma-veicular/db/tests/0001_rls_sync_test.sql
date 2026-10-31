-- Testes da migração 0001: isolamento (RLS), sincronização, auditoria e restrições.
-- Execução: ver db/run-tests.sh. Qualquer falha aborta com ON_ERROR_STOP.
-- Dados abaixo são FICTÍCIOS e existem apenas neste banco de teste efêmero.

\set ON_ERROR_STOP on

-- Papel de aplicação SEM privilégio de dono (RLS se aplica a ele).
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'garagem_app') THEN
    CREATE ROLE garagem_app NOLOGIN;
  END IF;
END $$;
GRANT USAGE ON SCHEMA app TO garagem_app;
GRANT USAGE ON SEQUENCE app.sync_seq TO garagem_app;
GRANT SELECT ON ALL TABLES IN SCHEMA public TO garagem_app;
GRANT INSERT, UPDATE ON vehicles, fuel_logs, documents, expenses, maintenance_records TO garagem_app;

-- Fixtures (como superusuário)
INSERT INTO users (id, email, full_name) VALUES
  ('00000000-0000-0000-0000-0000000000a1', 'ana@teste.invalid',   'Usuária Teste A'),
  ('00000000-0000-0000-0000-0000000000b1', 'bruno@teste.invalid', 'Usuário Teste B');
INSERT INTO accounts (id, kind, name) VALUES
  ('00000000-0000-0000-0000-00000000aaaa', 'personal', 'Conta A'),
  ('00000000-0000-0000-0000-00000000bbbb', 'personal', 'Conta B');
INSERT INTO account_members VALUES
  ('00000000-0000-0000-0000-00000000aaaa', '00000000-0000-0000-0000-0000000000a1', 'owner', now()),
  ('00000000-0000-0000-0000-00000000bbbb', '00000000-0000-0000-0000-0000000000b1', 'owner', now());

-- ---------------------------------------------------------------------------
-- 1. Usuário A cria um veículo e um abastecimento
-- ---------------------------------------------------------------------------
BEGIN;
SET LOCAL ROLE garagem_app;
SET LOCAL app.user_id = '00000000-0000-0000-0000-0000000000a1';
SET LOCAL app.account_ids = '{00000000-0000-0000-0000-00000000aaaa}';
INSERT INTO vehicles (id, account_id, category_code, propulsion_code, make_name, model_name, plate, vin, year_manufacture, year_model)
VALUES ('00000000-0000-0000-0000-0000000000c1', '00000000-0000-0000-0000-00000000aaaa',
        'car', 'ice', 'MarcaTeste', 'ModeloTeste', 'ABC1D23', '9BWZZZ377VT004251', 2022, 2023);
INSERT INTO fuel_logs (account_id, vehicle_id, filled_at, odometer_km, energy_code, quantity, unit_price_milli, total_cents)
VALUES ('00000000-0000-0000-0000-00000000aaaa', '00000000-0000-0000-0000-0000000000c1', now(), 1000, 'gasoline', 40, 5879, 23516);
COMMIT;

-- ---------------------------------------------------------------------------
-- 2. Usuário B NÃO enxerga dados de A
-- ---------------------------------------------------------------------------
BEGIN;
SET LOCAL ROLE garagem_app;
SET LOCAL app.user_id = '00000000-0000-0000-0000-0000000000b1';
SET LOCAL app.account_ids = '{00000000-0000-0000-0000-00000000bbbb}';
DO $$ BEGIN
  IF (SELECT count(*) FROM vehicles) <> 0 THEN RAISE EXCEPTION 'FALHA: B viu veículos de A'; END IF;
  IF (SELECT count(*) FROM fuel_logs) <> 0 THEN RAISE EXCEPTION 'FALHA: B viu abastecimentos de A'; END IF;
  RAISE NOTICE 'OK: isolamento de leitura entre contas';
END $$;
COMMIT;

-- 3. B não consegue inserir dados na conta de A
BEGIN;
SET LOCAL ROLE garagem_app;
SET LOCAL app.account_ids = '{00000000-0000-0000-0000-00000000bbbb}';
DO $$ BEGIN
  BEGIN
    INSERT INTO vehicles (account_id, category_code, propulsion_code, make_name, model_name)
    VALUES ('00000000-0000-0000-0000-00000000aaaa', 'car', 'ice', 'X', 'Y');
    RAISE EXCEPTION 'FALHA: B inseriu na conta de A';
  EXCEPTION WHEN insufficient_privilege THEN
    RAISE NOTICE 'OK: escrita cruzada bloqueada pela RLS';
  END;
END $$;
COMMIT;

-- 4. B não consegue alterar dados de A (UPDATE afeta 0 linhas)
BEGIN;
SET LOCAL ROLE garagem_app;
SET LOCAL app.account_ids = '{00000000-0000-0000-0000-00000000bbbb}';
DO $$ DECLARE n int; BEGIN
  UPDATE vehicles SET nickname = 'invadido' WHERE id = '00000000-0000-0000-0000-0000000000c1';
  GET DIAGNOSTICS n = ROW_COUNT;
  IF n <> 0 THEN RAISE EXCEPTION 'FALHA: B alterou veículo de A'; END IF;
  RAISE NOTICE 'OK: update cruzado afeta 0 linhas';
END $$;
COMMIT;

-- 5. Sem contexto => nada visível (fail-closed)
BEGIN;
SET LOCAL ROLE garagem_app;
DO $$ BEGIN
  IF (SELECT count(*) FROM vehicles) <> 0 THEN RAISE EXCEPTION 'FALHA: acesso sem contexto'; END IF;
  RAISE NOTICE 'OK: sem contexto de segurança, nenhuma linha visível';
END $$;
COMMIT;

-- 6. Compartilhamento: B lê o veículo compartilhado, mas não escreve
BEGIN;
SET LOCAL ROLE garagem_app;
SET LOCAL app.account_ids = '{00000000-0000-0000-0000-00000000bbbb}';
SET LOCAL app.shared_vehicle_ids = '{00000000-0000-0000-0000-0000000000c1}';
DO $$ DECLARE n int; BEGIN
  IF (SELECT count(*) FROM vehicles) <> 1 THEN RAISE EXCEPTION 'FALHA: compartilhado não visível'; END IF;
  IF (SELECT count(*) FROM fuel_logs) <> 1 THEN RAISE EXCEPTION 'FALHA: histórico compartilhado não visível'; END IF;
  BEGIN
    UPDATE vehicles SET nickname = 'x' WHERE id = '00000000-0000-0000-0000-0000000000c1';
    RAISE EXCEPTION 'FALHA: leitor compartilhado conseguiu escrever';
  EXCEPTION WHEN insufficient_privilege THEN
    RAISE NOTICE 'OK: compartilhamento dá leitura, não escrita';
  END;
END $$;
COMMIT;

-- ---------------------------------------------------------------------------
-- 7. Sincronização: version e server_seq avançam; created_at é imutável
-- ---------------------------------------------------------------------------
BEGIN;
SET LOCAL ROLE garagem_app;
SET LOCAL app.user_id = '00000000-0000-0000-0000-0000000000a1';
SET LOCAL app.account_ids = '{00000000-0000-0000-0000-00000000aaaa}';
DO $$ DECLARE v0 int; s0 bigint; c0 timestamptz; v1 int; s1 bigint; c1 timestamptz; BEGIN
  SELECT version, server_seq, created_at INTO v0, s0, c0 FROM vehicles WHERE id = '00000000-0000-0000-0000-0000000000c1';
  UPDATE vehicles SET nickname = 'Carro da família', created_at = '2000-01-01'
   WHERE id = '00000000-0000-0000-0000-0000000000c1';
  SELECT version, server_seq, created_at INTO v1, s1, c1 FROM vehicles WHERE id = '00000000-0000-0000-0000-0000000000c1';
  IF v1 <> v0 + 1 THEN RAISE EXCEPTION 'FALHA: version % -> %', v0, v1; END IF;
  IF s1 <= s0 THEN RAISE EXCEPTION 'FALHA: server_seq não avançou'; END IF;
  IF c1 <> c0 THEN RAISE EXCEPTION 'FALHA: created_at foi alterado'; END IF;
  RAISE NOTICE 'OK: version/server_seq/created_at corretos';
END $$;
-- exclusão lógica
UPDATE fuel_logs SET deleted_at = now() WHERE vehicle_id = '00000000-0000-0000-0000-0000000000c1';
COMMIT;

-- ---------------------------------------------------------------------------
-- 8. Auditoria registrou insert, update (com diff) e delete lógico
-- ---------------------------------------------------------------------------
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM audit_log WHERE entity_type='vehicles' AND action='insert') THEN
    RAISE EXCEPTION 'FALHA: auditoria de insert'; END IF;
  IF NOT EXISTS (SELECT 1 FROM audit_log WHERE entity_type='vehicles' AND action='update'
                 AND changes ? 'nickname' AND user_id = '00000000-0000-0000-0000-0000000000a1') THEN
    RAISE EXCEPTION 'FALHA: auditoria de update/diff/usuário'; END IF;
  IF NOT EXISTS (SELECT 1 FROM audit_log WHERE entity_type='fuel_logs' AND action='delete') THEN
    RAISE EXCEPTION 'FALHA: auditoria de exclusão lógica'; END IF;
  RAISE NOTICE 'OK: auditoria';
END $$;

-- O papel da aplicação não pode apagar/alterar a auditoria
BEGIN;
SET LOCAL ROLE garagem_app;
DO $$ BEGIN
  BEGIN
    DELETE FROM audit_log;
    RAISE EXCEPTION 'FALHA: app apagou auditoria';
  EXCEPTION WHEN insufficient_privilege THEN
    RAISE NOTICE 'OK: auditoria protegida contra a aplicação';
  END;
END $$;
COMMIT;

-- ---------------------------------------------------------------------------
-- 9. Restrições de domínio
-- ---------------------------------------------------------------------------
DO $$
DECLARE
  bad text[] := ARRAY[
    $q$INSERT INTO vehicles (account_id, category_code, make_name, model_name, plate) VALUES ('00000000-0000-0000-0000-00000000aaaa','car','X','Y','ABC-1234')$q$,
    $q$INSERT INTO vehicles (account_id, category_code, make_name, model_name, vin) VALUES ('00000000-0000-0000-0000-00000000aaaa','car','X','Y','9BWZZZ377VT00425I')$q$,
    $q$INSERT INTO vehicles (account_id, category_code, make_name, model_name, battery_gross_kwh, battery_usable_kwh) VALUES ('00000000-0000-0000-0000-00000000aaaa','car','X','Y',50,60)$q$,
    $q$INSERT INTO vehicles (account_id, category_code, make_name, model_name, plate) VALUES ('00000000-0000-0000-0000-00000000aaaa','car','X','Y','ABC1D23')$q$,
    $q$INSERT INTO expenses (account_id, category_code, occurred_on, amount_cents, origin_type) VALUES ('00000000-0000-0000-0000-00000000aaaa','fuel','2026-01-01',100,'fuel')$q$,
    $q$INSERT INTO documents (account_id, doc_type, title, payment_status) VALUES ('00000000-0000-0000-0000-00000000aaaa','ipva','IPVA','paid')$q$,
    $q$INSERT INTO maintenance_plan_items (account_id, vehicle_id, category, title, interval_source) VALUES ('00000000-0000-0000-0000-00000000aaaa','00000000-0000-0000-0000-0000000000c1','oil','Óleo','user')$q$
  ];
  q text; ok int := 0;
BEGIN
  FOREACH q IN ARRAY bad LOOP
    BEGIN
      EXECUTE q;
      RAISE EXCEPTION 'FALHA: restrição não aplicada: %', q;
    EXCEPTION WHEN check_violation OR unique_violation THEN ok := ok + 1;
    END;
  END LOOP;
  RAISE NOTICE 'OK: % restrições de domínio rejeitaram dados inválidos', ok;
END $$;

-- 10. Extensibilidade: nova categoria = INSERT, sem alterar estrutura
INSERT INTO ref_vehicle_categories (code, name_pt, is_powered, is_heavy) VALUES ('agri_machine','Máquina agrícola', true, true);
INSERT INTO vehicles (account_id, category_code, propulsion_code, make_name, model_name)
VALUES ('00000000-0000-0000-0000-00000000aaaa', 'agri_machine', 'ice', 'X', 'Y');
\echo 'OK: nova categoria adicionada sem migração estrutural'

\echo '== TODOS OS TESTES SQL PASSARAM =='
