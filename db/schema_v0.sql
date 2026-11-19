-- TQN Register Universal — schema v0 (G2 POC, chưa phải migration production)
-- Mục đích: chứng minh mô hình ADR-002/003/004 bằng PostgreSQL thật trước khi viết C#.
-- Quy ước: mọi bảng nghiệp vụ có tenant_id + project_id; thời gian lưu UTC (timestamptz).
BEGIN;
CREATE EXTENSION IF NOT EXISTS pgcrypto;
CREATE SCHEMA IF NOT EXISTS tqn;
SET search_path = tqn;

CREATE TABLE tenant (
  tenant_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text NOT NULL UNIQUE,
  name text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE project (
  project_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id uuid NOT NULL REFERENCES tenant,
  code text NOT NULL,
  name text NOT NULL,
  timezone text NOT NULL DEFAULT 'Asia/Ho_Chi_Minh',
  lat numeric(9,6), lon numeric(9,6),
  gps_confirmed boolean NOT NULL DEFAULT false,
  weather_enabled boolean NOT NULL DEFAULT false,
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (tenant_id, code),
  UNIQUE (tenant_id, project_id),
  -- Thời tiết chỉ bật khi GPS đã xác nhận (chỉ thị VI: không suy tọa độ âm thầm)
  CHECK (NOT weather_enabled OR (gps_confirmed AND lat IS NOT NULL AND lon IS NOT NULL))
);

-- Cấu hình dự án có phiên bản, không ghi đè (profile COASTAL_2026_V1 là một bản ghi, không phải code)
CREATE TABLE project_config_version (
  project_id uuid NOT NULL REFERENCES project,
  version int NOT NULL,
  code_profile text NOT NULL,
  config jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  created_by text NOT NULL,
  PRIMARY KEY (project_id, version)
);

CREATE TABLE organization (
  org_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id uuid NOT NULL REFERENCES tenant,
  project_id uuid NOT NULL,
  code text NOT NULL,          -- TBCONS / TVGS / CĐT / TVTK ...
  role text NOT NULL,          -- CONTRACTOR / SUPERVISOR / OWNER / DESIGNER
  UNIQUE (project_id, code),
  FOREIGN KEY (tenant_id, project_id) REFERENCES project (tenant_id, project_id)
);

-- Dãy số: phát sinh trong transaction, khóa dòng (SELECT ... FOR UPDATE / UPDATE ... RETURNING)
CREATE TABLE number_sequence (
  project_id uuid NOT NULL REFERENCES project,
  doc_type text NOT NULL,
  scope text NOT NULL DEFAULT '',      -- ví dụ 'NT', 'NP', 'HTKT'
  last_value int NOT NULL DEFAULT 0,
  PRIMARY KEY (project_id, doc_type, scope)
);

-- Hồ sơ logic
CREATE TABLE document (
  document_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id uuid NOT NULL,
  project_id uuid NOT NULL,
  doc_type text NOT NULL,                 -- RFI/RFA/RFC/SDS/MSS/MAS/OTS/SCH/RFIs...
  seq_no int NOT NULL,
  doc_key text NOT NULL,                  -- mã không gồm tiền tố ngày và đuôi rev
  subject text NOT NULL,
  owner_org_id uuid REFERENCES organization,
  created_at timestamptz NOT NULL DEFAULT now(),
  FOREIGN KEY (tenant_id, project_id) REFERENCES project (tenant_id, project_id),
  UNIQUE (project_id, doc_type, seq_no),
  UNIQUE (project_id, doc_key)
);

-- Phiên bản phát hành: chỉ thêm, không sửa, không xóa
CREATE TABLE document_revision (
  revision_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  document_id uuid NOT NULL REFERENCES document,
  rev_label text NOT NULL,                -- 'Rev00' / 'Rev01' / 'A' / 'B' — theo profile, không ép
  rev_order int NOT NULL,
  raw_document_number text NOT NULL,      -- số hiệu thô đúng như đã phát hành
  source_sheet text, source_row int, source_hash text, mapper_version text,
  issued_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (document_id, rev_label),
  UNIQUE (document_id, rev_order)
);

-- Lần trình
CREATE TABLE submission (
  submission_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  revision_id uuid NOT NULL REFERENCES document_revision,
  submitted_by_org uuid REFERENCES organization,
  submitted_at timestamptz NOT NULL,
  due_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now()
);

-- Quyết định của từng tổ chức cho từng lượt — có actor, căn cứ, thời điểm
CREATE TABLE review_decision (
  decision_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  submission_id uuid NOT NULL REFERENCES submission,
  org_id uuid NOT NULL REFERENCES organization,
  actor text NOT NULL,
  decision text NOT NULL CHECK (decision IN
    ('DA_DUYET','DUYET_CO_Y_KIEN','KHONG_DUYET_TRINH_LAI','DA_PHAN_HOI')),
  comment text,
  basis_file text,
  decided_at timestamptz NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (submission_id, org_id)
);

CREATE TABLE attachment (
  attachment_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  revision_id uuid NOT NULL REFERENCES document_revision,
  file_name text NOT NULL,
  sha256 text NOT NULL,
  provider text NOT NULL DEFAULT 'GOOGLE_DRIVE',
  provider_file_id text,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE audit_event (
  audit_id bigserial PRIMARY KEY,
  tenant_id uuid, project_id uuid,
  actor text NOT NULL,
  action text NOT NULL,
  entity text NOT NULL,
  entity_id text NOT NULL,
  payload jsonb,
  at timestamptz NOT NULL DEFAULT now()
);

-- Import từ Register Google Sheets/xlsx: chỉ một chiều, idempotent theo source_hash
CREATE TABLE import_batch (
  batch_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  project_id uuid NOT NULL REFERENCES project,
  source_name text NOT NULL,
  source_sha256 text NOT NULL,
  mapper_version text NOT NULL,
  started_at timestamptz NOT NULL DEFAULT now(),
  status text NOT NULL DEFAULT 'PREVIEW' CHECK (status IN ('PREVIEW','APPLIED','ROLLED_BACK','FAILED'))
);
CREATE TABLE import_row_result (
  batch_id uuid NOT NULL REFERENCES import_batch,
  source_sheet text NOT NULL,
  source_row int NOT NULL,
  row_hash text NOT NULL,
  result text NOT NULL CHECK (result IN ('NEW','UNCHANGED','CONFLICT','UNMAPPED','ERROR')),
  message text,
  PRIMARY KEY (batch_id, source_sheet, source_row)
);

-- Thời tiết: khóa chống chạy trùng theo (project, ngày địa phương, loại, slot)
CREATE TABLE weather_job (
  job_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  project_id uuid NOT NULL REFERENCES project,
  local_date date NOT NULL,
  job_type text NOT NULL DEFAULT 'FORECAST',
  slot text NOT NULL DEFAULT '06:00',
  state text NOT NULL DEFAULT 'RUNNING' CHECK (state IN ('RUNNING','COMPLETE','PARTIAL','FAILED')),
  attempts int NOT NULL DEFAULT 1,
  started_at timestamptz NOT NULL DEFAULT now(),
  finished_at timestamptz,
  error text,
  UNIQUE (project_id, local_date, job_type, slot)
);
CREATE TABLE weather_snapshot (
  snapshot_id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  job_id uuid NOT NULL REFERENCES weather_job,
  kind text NOT NULL CHECK (kind IN ('FORECAST','OBSERVATION')),
  provider text NOT NULL, model text, issued_at timestamptz NOT NULL,
  raw jsonb NOT NULL, raw_sha256 text NOT NULL
);

-- ===== Bất biến: không UPDATE/DELETE revision và audit =====
CREATE FUNCTION forbid_change() RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN RAISE EXCEPTION 'TQN: % trên % bị cấm (bản ghi bất biến)', TG_OP, TG_TABLE_NAME; END $$;
CREATE TRIGGER rev_immutable BEFORE UPDATE OR DELETE ON document_revision
  FOR EACH ROW EXECUTE FUNCTION forbid_change();
CREATE TRIGGER audit_immutable BEFORE UPDATE OR DELETE ON audit_event
  FOR EACH ROW EXECUTE FUNCTION forbid_change();

-- ===== Cấp số an toàn đa người: một câu UPDATE ... RETURNING trong transaction =====
CREATE FUNCTION next_number(p_project uuid, p_type text, p_scope text DEFAULT '')
RETURNS int LANGUAGE plpgsql AS $$
DECLARE v int;
BEGIN
  INSERT INTO number_sequence(project_id, doc_type, scope) VALUES (p_project, p_type, p_scope)
  ON CONFLICT DO NOTHING;
  UPDATE number_sequence SET last_value = last_value + 1
   WHERE project_id = p_project AND doc_type = p_type AND scope = p_scope
   RETURNING last_value INTO v;
  RETURN v;
END $$;

-- LatestRevision là projection tính được, không lưu đè
CREATE VIEW latest_revision AS
SELECT DISTINCT ON (r.document_id) r.*
FROM document_revision r ORDER BY r.document_id, r.rev_order DESC;

COMMIT;
