#!/usr/bin/env bash
# Kiểm thử schema v0 trên PostgreSQL thật. Dùng: PGHOST=/tmp/pg PGPORT=55432 PGUSER=tqn bash run_db_tests.sh
set -u
DB=tqn_test_$$
HERE=$(cd "$(dirname "$0")" && pwd)
q()  { psql -X -q -v ON_ERROR_STOP=1 -d $DB -At -c "SET search_path=tqn; $1" 2>&1; }
pass=0; fail=0
ok() { if [ "$2" = "$3" ]; then echo "PASS $1 (= $3)"; pass=$((pass+1)); else echo "FAIL $1: được '$2', cần '$3'"; fail=$((fail+1)); fi; }
err(){ out=$(q "$2"); if echo "$out" | grep -q "$3"; then echo "PASS $1 (bị chặn: $3)"; pass=$((pass+1)); else echo "FAIL $1: không bị chặn -> $out"; fail=$((fail+1)); fi; }

createdb $DB || exit 1
psql -X -q -v ON_ERROR_STOP=1 -d $DB -f "$HERE/../schema_v0.sql" >/dev/null && echo "PASS T01 áp schema" && pass=$((pass+1))

# Dữ liệu 2 dự án demo (không phải dữ liệu thật)
q "INSERT INTO tenant(code,name) VALUES ('TQN','TQN DEMO');
   INSERT INTO project(tenant_id,code,name) SELECT tenant_id,'DEMO_A','Du an A (DEMO)' FROM tenant;
   INSERT INTO project(tenant_id,code,name) SELECT tenant_id,'DEMO_B','Du an B (DEMO)' FROM tenant;" >/dev/null
A=$(q "SELECT project_id FROM project WHERE code='DEMO_A'"); B=$(q "SELECT project_id FROM project WHERE code='DEMO_B'")
T=$(q "SELECT tenant_id FROM tenant")

# T02 cùng mã ở 2 dự án không va chạm; trùng trong 1 dự án bị chặn
q "INSERT INTO document(tenant_id,project_id,doc_type,seq_no,doc_key,subject) VALUES
   ('$T','$A','RFI',1,'COA-NT-TBC-RFI-001','A1'),('$T','$B','RFI',1,'COA-NT-TBC-RFI-001','B1');" >/dev/null
ok T02a "$(q "SELECT count(*) FROM document WHERE doc_key='COA-NT-TBC-RFI-001'")" 2
err T02b "INSERT INTO document(tenant_id,project_id,doc_type,seq_no,doc_key,subject) VALUES ('$T','$A','RFI',2,'COA-NT-TBC-RFI-001','trung')" "duplicate key"

# T03 100 yêu cầu cấp số đồng thời (20 tiến trình x 5) không trùng
for i in $(seq 1 20); do ( for j in 1 2 3 4 5; do q "SELECT next_number('$A','RFA','NT')"; done ) > /tmp/seq_$$_$i.txt & done; wait
ok T03a "$(cat /tmp/seq_$$_*.txt | sort -n | uniq | wc -l | tr -d ' ')" 100
ok T03b "$(cat /tmp/seq_$$_*.txt | sort -n | tail -1)" 100
ok T03c "$(q "SELECT next_number('$B','RFA','NT')")" 1
rm -f /tmp/seq_$$_*.txt

# T04/T05 vòng đời: 3 Rev, 2 lượt trình, 2 bên ý kiến khác nhau
D=$(q "SELECT document_id FROM document WHERE project_id='$A'")
q "INSERT INTO organization(tenant_id,project_id,code,role) VALUES ('$T','$A','TVGS','SUPERVISOR'),('$T','$A','CDT','OWNER');
   INSERT INTO document_revision(document_id,rev_label,rev_order,raw_document_number) VALUES
   ('$D','Rev00',0,'261101-COA-NT-TBC-RFI-001 Rev00'),('$D','Rev01',1,'261105-COA-NT-TBC-RFI-001 Rev01'),('$D','Rev02',2,'261110-COA-NT-TBC-RFI-001 Rev02');
   INSERT INTO submission(revision_id,submitted_at) SELECT revision_id, now() FROM document_revision WHERE rev_label IN ('Rev01','Rev02');
   INSERT INTO review_decision(submission_id,org_id,actor,decision,decided_at)
     SELECT s.submission_id,o.org_id,'demo',CASE o.code WHEN 'TVGS' THEN 'DUYET_CO_Y_KIEN' ELSE 'KHONG_DUYET_TRINH_LAI' END, now()
     FROM submission s JOIN document_revision r USING(revision_id), organization o WHERE r.rev_label='Rev02';" >/dev/null
err T04 "UPDATE document_revision SET raw_document_number='sua' WHERE rev_label='Rev00'" "bất biến"
err T04b "DELETE FROM document_revision WHERE rev_label='Rev00'" "bất biến"
ok T05a "$(q "SELECT rev_label FROM latest_revision WHERE document_id='$D'")" Rev02
ok T05b "$(q "SELECT count(*) FROM document_revision WHERE document_id='$D'")" 3
ok T05c "$(q "SELECT string_agg(decision,',' ORDER BY decision) FROM review_decision")" "DUYET_CO_Y_KIEN,KHONG_DUYET_TRINH_LAI"
err T05d "INSERT INTO review_decision(submission_id,org_id,actor,decision,decided_at) SELECT submission_id,org_id,'x','DA_DUYET',now() FROM review_decision LIMIT 1" "duplicate key"

# T06 import 2 lần không nhân đôi (upsert theo doc_key)
IMP="INSERT INTO document(tenant_id,project_id,doc_type,seq_no,doc_key,subject) VALUES ('$T','$B','SDS',25,'COA-NT-TBC-SDS-025','imp') ON CONFLICT (project_id,doc_key) DO NOTHING"
q "$IMP" >/dev/null; q "$IMP" >/dev/null
ok T06 "$(q "SELECT count(*) FROM document WHERE project_id='$B'")" 2

# T07 khóa thời tiết: 2 lần gọi cùng slot -> 1 job
J="INSERT INTO weather_job(project_id,local_date) VALUES ('$A','2026-11-20') ON CONFLICT DO NOTHING RETURNING 1"
r1=$(q "$J"); r2=$(q "$J")
ok T07 "${r1}|${r2}|$(q "SELECT count(*) FROM weather_job")" "1||1"

# T08 không bật thời tiết khi chưa xác nhận GPS
err T08 "UPDATE project SET weather_enabled=true WHERE project_id='$B'" "check constraint"

# T09 không gắn hồ sơ vào dự án của tenant khác
q "INSERT INTO tenant(code,name) VALUES ('KHAC','Tenant khac')" >/dev/null
T2=$(q "SELECT tenant_id FROM tenant WHERE code='KHAC'")
err T09 "INSERT INTO document(tenant_id,project_id,doc_type,seq_no,doc_key,subject) VALUES ('$T2','$A','RFI',9,'X','x')" "foreign key"

# T10 audit bất biến
q "INSERT INTO audit_event(actor,action,entity,entity_id) VALUES ('demo','CREATE','document','$D')" >/dev/null
err T10 "DELETE FROM audit_event" "bất biến"

dropdb $DB
echo "TONG: PASS=$pass FAIL=$fail"
[ $fail -eq 0 ]
