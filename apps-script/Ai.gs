/**
 * Ai.gs – TRỢ LÝ AI ĐỌC NỘI DUNG hồ sơ (S2).
 * Đọc nội dung + ý kiến TVGS/TVTK/CĐT của 1 hồ sơ → tóm tắt, mức độ, bên phải xử lý, việc tiếp theo,
 * danh sách cần bổ sung, nháp trả lời. KẾT QUẢ CHỈ LÀ BẢN NHÁP: ghi vào tab TQN_AI, không ghi vào sổ gốc,
 * người phụ trách đánh dấu "Đã duyệt" sau khi kiểm.
 * Khóa API để ở Script Properties (ANTHROPIC_API_KEY) – không để trong ô bảng tính, không gửi qua chat.
 */
var TQN_AI_DEFAULTS = [
  ['ai_enabled', 'Không', 'Có = bật trợ lý AI đọc nội dung (cần khóa ANTHROPIC_API_KEY trong Script Properties)'],
  ['ai_model', 'claude-sonnet-5-5', 'Mô hình AI'],
  ['ai_batch_max', '10', 'Số hồ sơ tối đa mỗi lần bấm "AI rà hồ sơ cần xử lý"']
];
var TQN_AI_HEADERS = ['Thời điểm', 'Vị trí', 'Số hiệu', 'Loại', 'Mức độ', 'Bên cần xử lý', 'Tóm tắt', 'Việc tiếp theo', 'Cần bổ sung',
  'Nháp trả lời', 'Mô hình', 'Đã duyệt bởi (người ghi)', 'Ghi chú'];
var TQN_AI_LEVELS = ['CAO', 'TRUNG_BINH', 'THAP'];

function tqnEnsureAiConfig_() {
  tqnEnsureConfig_();
  var sh = tqnSs_().getSheetByName(TQN.SH_CONFIG), have = {};
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = 1; });
  TQN_AI_DEFAULTS.forEach(function (r) { if (!have[r[0]]) sh.appendRow(r); });
}
function tqnAiCfg_() {
  var cfg = tqnConfig_();
  TQN_AI_DEFAULTS.forEach(function (r) { if (cfg[r[0]] === undefined || cfg[r[0]] === '') cfg[r[0]] = r[1]; });
  cfg.aiOn = /^c|^y|^1/i.test(cfg.ai_enabled);
  cfg.aiMax = Math.max(1, Math.min(30, parseInt(cfg.ai_batch_max, 10) || 10));
  return cfg;
}
function tqnAiBlocked_(cfg) {
  if (!cfg.aiOn) return 'ai_enabled chưa bật (Có) trong TQN_CONFIG';
  var k = PropertiesService.getScriptProperties ? PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') : null;
  if (!k) return 'Chưa có khóa ANTHROPIC_API_KEY trong Script Properties (Apps Script → Cài đặt dự án → Thuộc tính tập lệnh)';
  return '';
}

/** Lấy 1 hồ sơ từ TQN_DATA theo vị trí "Tab!dòng". */
function tqnAiRecord_(at) {
  var sh = tqnSs_().getSheetByName(TQN.SH_DATA);
  if (!sh || sh.getLastRow() < 2) throw new Error('Chưa có TQN_DATA – bấm "Nhập & kiểm tra" trước.');
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, TQN_DATA_HEADERS.length).getValues(), C = TQN_COL;
  var key = String(at).split(',')[0].trim();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][C['Nguồn']]) === key) {
    var r = vals[i], o = {};
    TQN_DATA_HEADERS.forEach(function (h, j) { o[h] = r[j] instanceof Date ? Utilities.formatDate(r[j], 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : r[j]; });
    return o;
  }
  throw new Error('Không thấy hồ sơ ở ' + key);
}

/** Câu lệnh gửi AI: chỉ dữ liệu của đúng 1 hồ sơ, không gửi email/SĐT. Dùng chung cho Apps Script và bản xem thử. */
function tqnAiPrompt_(rec) {
  var keep = ['Loại', 'Số hiệu (nguyên văn)', 'Rev', 'Nội dung', 'Tình trạng (sổ chi tiết)', 'Ngày trình', 'TVGS ngày', 'TVGS ý kiến',
    'TVTK ngày', 'TVTK ý kiến', 'CĐT ngày', 'CĐT ý kiến', 'Trình đến', 'Hạn phản hồi', 'Ngày phản hồi', 'Kết quả', 'Tình trạng (Sổ trình duyệt)'];
  var data = keep.map(function (k) { return k + ': ' + String(rec[k] === undefined || rec[k] === '' ? '(trống)' : rec[k]).replace(/[\w.+-]+@[\w-]+\.[\w.]+/g, '[email]').slice(0, 1500); }).join('\n');
  return 'Bạn là thư ký kiểm soát hồ sơ trình duyệt của công trường xây dựng tại Việt Nam (dự án Coastal Quảng Ngãi; các bên: Nhà thầu TBCONS, '
    + 'Tư vấn giám sát TVGS, Tư vấn thiết kế TVTK, Chủ đầu tư CĐT). Đọc dữ liệu 1 hồ sơ dưới đây và trả lời ĐÚNG MỘT đối tượng JSON, không thêm chữ nào khác, theo mẫu:\n'
    + '{"tom_tat":"≤2 câu","muc_do":"CAO|TRUNG_BINH|THAP","ben_xu_ly":"Nhà thầu|TVGS|TVTK|CĐT|Không ai","viec_tiep_theo":"1 câu, động từ đầu câu",'
    + '"can_bo_sung":["mục cần bổ sung/sửa theo ý kiến, nếu có"],"nhap_tra_loi":"nháp văn bản trả lời/đệ trình lại ngắn gọn, lịch sự, tiếng Việt"}\n'
    + 'Quy tắc: chỉ dựa vào dữ liệu đã cho, không bịa số liệu, ngày, tên người. Thiếu thông tin thì ghi rõ "chưa có thông tin". '
    + 'Mức độ CAO nếu quá hạn, bị trả về/không duyệt, hoặc ý kiến yêu cầu sửa quan trọng; THAP nếu đã duyệt xong.\n\n--- DỮ LIỆU HỒ SƠ ---\n' + data;
}

/** Đọc JSON từ câu trả lời AI, kiểm khuôn; lỗi thì ném ra (không ghi bậy). */
function tqnAiParse_(text) {
  var s = String(text || ''), a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a < 0 || b <= a) throw new Error('AI không trả về JSON');
  var o = JSON.parse(s.slice(a, b + 1));
  if (!o.tom_tat || !o.viec_tiep_theo) throw new Error('AI trả về thiếu trường tom_tat / viec_tiep_theo');
  o.muc_do = TQN_AI_LEVELS.indexOf(String(o.muc_do).toUpperCase()) >= 0 ? String(o.muc_do).toUpperCase() : 'TRUNG_BINH';
  o.can_bo_sung = Array.isArray(o.can_bo_sung) ? o.can_bo_sung.map(String).slice(0, 10) : [];
  o.nhap_tra_loi = String(o.nhap_tra_loi || '');
  o.ben_xu_ly = String(o.ben_xu_ly || '');
  return o;
}

function tqnAiSave_(rec, o, model) {
  var sh = tqnOwnSheet_(TQN.PREFIX + 'AI', TQN_AI_HEADERS);
  sh.appendRow([new Date(), rec['Nguồn'], rec['Số hiệu (nguyên văn)'], rec['Loại'], o.muc_do, o.ben_xu_ly, o.tom_tat, o.viec_tiep_theo,
    o.can_bo_sung.join('; '), o.nhap_tra_loi, model, '', 'Bản nháp do AI soạn – cần người kiểm tra trước khi dùng']);
  tqnAudit_('AI_DOC_HO_SO', rec['Nguồn'] + ' ' + rec['Số hiệu (nguyên văn)'], { muc_do: o.muc_do, model: model });
}

/** Gọi Claude qua API Anthropic (thử lại khi 429/5xx). */
function tqnAiCall_(cfg, prompt) {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'), last = '';
  for (var i = 0; i < 3; i++) {
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: cfg.ai_model, max_tokens: 1200, messages: [{ role: 'user', content: prompt }] })   // prompt: chữ, hoặc mảng [ảnh/PDF, chữ] khi đọc phiếu
    });
    var code = res.getResponseCode(), body = res.getContentText();
    if (code === 200) {
      var j = JSON.parse(body);
      return (j.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
    }
    last = 'HTTP ' + code + ' ' + body.slice(0, 200);
    if (code !== 429 && code < 500) break;
    Utilities.sleep(Math.pow(2, i) * 1500);
  }
  throw new Error('Gọi AI lỗi: ' + last);
}

/** Phân tích 1 hồ sơ (gọi từ bảng quy trình). */
function tqnAiAnalyze(at) {
  var cfg = tqnAiCfg_(), why = tqnAiBlocked_(cfg);
  if (why) throw new Error('Chưa bật AI: ' + why);
  var rec = tqnAiRecord_(at), o = tqnAiParse_(tqnAiCall_(cfg, tqnAiPrompt_(rec)));
  tqnAiSave_(rec, o, cfg.ai_model);
  o.so = rec['Số hiệu (nguyên văn)']; o.at = rec['Nguồn']; o.model = cfg.ai_model;
  return o;
}

/** AI rà các hồ sơ cần xử lý trước (quá hạn / trả về / sắp hạn), tối đa ai_batch_max. */
function tqnAiBatch() {
  var cfg = tqnAiCfg_(), why = tqnAiBlocked_(cfg);
  if (why) throw new Error('Chưa bật AI: ' + why);
  var sh = tqnSs_().getSheetByName(TQN.PREFIX + 'QUY_TRINH');
  if (!sh || sh.getLastRow() < 2) throw new Error('Chưa có TQN_QUY_TRINH – bấm "Nhập & kiểm tra" trước.');
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, TQN_WF_HEADERS.length).getValues()
    .filter(function (r) { return r[11] || /^4/.test(r[0]); }).slice(0, cfg.aiMax);
  var ok = 0, err = [];
  rows.forEach(function (r) { try { tqnAiAnalyze(r[13]); ok++; } catch (e) { err.push(r[3] + ': ' + e.message); } });
  return { ok: ok, err: err, total: rows.length };
}
function tqnMenuAiBatch() {
  tqnEnsureAiConfig_();
  var r = tqnAiBatch();
  SpreadsheetApp.getUi().alert('AI rà hồ sơ', 'Đã đọc ' + r.ok + '/' + r.total + ' hồ sơ → tab TQN_AI (bản nháp, cần người kiểm).' + (r.err.length ? '\nLỗi: ' + r.err.join('\n') : ''), SpreadsheetApp.getUi().ButtonSet.OK);
}
