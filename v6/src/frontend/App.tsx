import { useEffect, useMemo, useRef, useState } from 'react';
import { median, p95 } from './stats';
import { api, isDemo, newIdem, setDemoUser } from './api';

type Doc = any;
const ST: Record<string, [string, string]> = {
  DRAFT: ['Đang soạn', 'gray'], SUBMITTED: ['Chờ CĐT', 'amber'], NEED_INFO: ['Bổ sung', 'red'], RESUBMIT_REQUIRED: ['Trình lại', 'red'],
  FORWARDED: ['Đã chuyển tiếp', 'blue'], CLOSED: ['Đã đóng', 'green'], REJECTED: ['Từ chối', 'gray'],
};
const fmt = (s?: string | null) => (s ? new Date(s).toLocaleString('vi-VN', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }) : '—');
type Pend = { label: string; op: string; idem: string; a: any; file?: File; fileName?: string; at: number; err?: string; code?: string; job?: any };
type Perf = { label: string; ui: number; call: number; server?: Record<string, any> };
const RUNS: Perf[] = [];
function savePend(p: Pend | null) { try { if (p) localStorage.setItem('tqn6_pend', JSON.stringify({ ...p, file: undefined })); else localStorage.removeItem('tqn6_pend'); } catch { /* trình duyệt chặn lưu cục bộ */ } }
function loadPend(): Pend | null { try { const v = localStorage.getItem('tqn6_pend'); if (!v) return null; const p = JSON.parse(v); return { ...p, err: p.err || 'Thao tác chưa được máy chủ xác nhận trước khi trang tải lại – bấm Thử lại (an toàn, không ghi trùng)' }; } catch { return null; } }
const ACT: Record<string, string> = { CREATE: 'Tạo hồ sơ', ATTACH: 'Gắn file', SUBMIT: 'Trình', REPLY: 'CĐT trả lời', CLOSE_FORWARDED: 'Đóng sau chuyển tiếp' };

export default function App() {
  const [who, setWho] = useState<{ email: string; roles: string[]; build: any } | null>(null);
  const [docs, setDocs] = useState<Doc[]>([]);
  const [sel, setSel] = useState<string | null>(null);
  const [v360, setV360] = useState<any>(null);
  const [tab, setTab] = useState<'mine' | 'late' | 'all'>('mine');
  const [q, setQ] = useState('');
  const [msg, setMsg] = useState<[string, boolean] | null>(null);
  const [newOpen, setNewOpen] = useState(false), [bench, setBench] = useState(false), [rep, setRep] = useState(false);

  const isRev = who?.roles.includes('REVIEWER') || who?.roles.includes('ADMIN');
  const isNt = who?.roles.some(r => ['CONTRACTOR', 'DC', 'ADMIN'].includes(r));

  // ---- bộ đệm phía trình duyệt: theo dự án + người dùng + phiên bản dữ liệu (dv). dv đổi → bỏ bộ đệm, lấy bản mới ----
  const [cache] = useState(() => new Map<string, any>());
  const dvRef = useRef(''), selRef = useRef<string | null>(null);
  const [pend, setPend] = useState<Pend | null>(() => loadPend());
  const [perf, setPerf] = useState<Perf | null>(null);
  const [fresh, setFresh] = useState('');
  selRef.current = sel;
  function seeDv(v?: string) { if (v && v !== dvRef.current) { cache.clear(); dvRef.current = v; } }
  function take(r: any) {
    seeDv(r.dv); if (r.list) setDocs(r.list);
    if (r.v360) { cache.set(r.v360.doc.documentId, r.v360); if (!selRef.current || selRef.current === r.v360.doc.documentId) { setSel(r.v360.doc.documentId); setV360(r.v360); } }
  }
  async function load(keep?: string) { const b = await api.boot(keep); setWho(b.who); take(b); setFresh(new Date().toLocaleTimeString('vi-VN')); }
  useEffect(() => { load().catch(e => setMsg([e.message, false])); }, []);
  // phát hiện thay đổi của người khác: hỏi nhẹ mỗi 20 s và khi quay lại tab (không mở Sheets)
  useEffect(() => {
    const check = () => { if (document.hidden) return; api.peek().then(p => { if (p.dv !== dvRef.current && !pendRef.current) load(selRef.current ?? undefined).then(() => setMsg(['Có cập nhật mới từ người khác – đã tải bản mới', true])); }).catch(() => {}); };
    const t = setInterval(check, 20000); document.addEventListener('visibilitychange', check);
    return () => { clearInterval(t); document.removeEventListener('visibilitychange', check); };
  }, []);
  const pendRef = useRef<Pend | null>(null); pendRef.current = pend && !pend.err ? pend : null;
  function pick(id: string) {
    setSel(id); const c = cache.get(id); if (c) { setV360(c); return; }   // có trong bộ đệm cùng dv → hiện ngay, không gọi máy chủ
    api.get360(id).then(v => { seeDv(v.dv); cache.set(id, v); if (selRef.current === id) setV360(v); }).catch(e => setMsg([e.message, false]));
  }

  /** SAFE WRITE: bấm → hiện ĐANG ĐỒNG BỘ ngay (không báo thành công giả) → chỉ báo "Đã lưu" khi máy chủ xác nhận.
   *  Lỗi: giữ nguyên dữ liệu đã nhập + cùng khóa chống trùng để Thử lại; không hoàn lại mù – lấy bản thật từ máy chủ. */
  async function act(label: string, op: string, idem: string, a: any, file?: File) {
    if (pendRef.current) return;
    if (a.id && a.ev === undefined && v360?.doc.documentId === a.id && op !== 'attach') a = { ...a, ev: v360.doc.version };
    const t0 = performance.now(), p: Pend = { label, op, idem, a, file, fileName: file?.name, at: Date.now() };
    setPend(p); savePend(p); setMsg(null);
    let ui = 0; requestAnimationFrame(() => requestAnimationFrame(() => { ui = performance.now() - t0; }));
    try {
      const r = file ? await api.upload(idem, a.id, file) : await api.act(op, idem, a);
      take(r); setPend(null); savePend(null);
      const pf: Perf = { label, ui, call: performance.now() - t0, server: r.ms }; setPerf(pf); RUNS.push(pf);
      const so = r.soRFI || '', soBad = /^Lỗi|^Không/.test(so);
      setMsg([`Đã lưu: ${label}${r.job?.note ? ' – ' + r.job.note : ''} (${(pf.call / 1000).toFixed(1)} s)${so ? (soBad ? ' · ⚠ ' : ' · sổ ') + so : ''}`, !soBad]);
    } catch (e: any) {
      const conflict = e.code === 'CONFLICT';
      const np = conflict ? null : { ...p, err: e.message, code: e.code, job: e.data?.job };
      setPend(np); savePend(np);
      if (conflict) setMsg([e.message, false]);
      // lấy trạng thái thật (người khác có thể vừa đổi) – không tự đoán, không hoàn lại mù
      load(a.id || selRef.current || undefined).catch(() => {});
    }
  }
  const retry = () => { if (!pend) return; if (pend.op === 'upload' && !pend.file) { setMsg(['Trang đã tải lại nên mất file đang chọn – chọn lại file (cùng job, không sinh bản trùng)', false]); return; }
    const p = pend; setPend(null); setTimeout(() => act(p.label, p.op, p.idem, { ...p.a, ev: undefined }, p.file)); };
  const busy = !!pendRef.current;

  const mine = (d: Doc) => (isRev && d.status === 'SUBMITTED') || (isNt && ['DRAFT', 'NEED_INFO', 'RESUBMIT_REQUIRED'].includes(d.status));
  const shown = useMemo(() => docs.filter(d => (tab === 'all' || (tab === 'late' ? d.overdueDays > 0 : mine(d))) &&
    (!q || (d.docNo + ' ' + d.title).toLowerCase().includes(q.toLowerCase()))), [docs, tab, q, who]);
  const cnt = { mine: docs.filter(mine).length, late: docs.filter(d => d.overdueDays > 0).length, all: docs.length };

  return (
    <div className="app">
      <header>
        <div className="brand"><b>TQN</b> Register <span className="muted">· RFI</span></div>
        <div className="sp" />
        {isDemo && <select value={who?.email} onChange={e => { setDemoUser(e.target.value); cache.clear(); load(sel ?? undefined); }} title="Chỉ có ở bản xem trước">
          <option value="nhathau@demo">Xem như: Nhà thầu (DEMO)</option><option value="cdt@demo">Xem như: CĐT (DEMO)</option></select>}
        <span className="who">{who?.email} · {who?.roles.join(', ') || 'chưa có quyền'}</span>
        <button onClick={() => setRep(true)}>⤓ Xuất báo cáo</button>
        {isNt && <button className="pri" onClick={() => setNewOpen(true)}>+ RFI mới</button>}
      </header>
      {pend && !pend.err && <div className="sync">⟳ ĐANG ĐỒNG BỘ: {pend.label}… <Since t={pend.at} /></div>}
      {pend?.err && <div className="sync err">✗ CHƯA LƯU: {pend.label} – {pend.err}{pend.job?.state && <> · job {pend.job.jobId?.slice(-6)} dừng ở <b>{pend.job.failedAt || pend.job.state}</b></>}
        <button onClick={retry}>Thử lại</button><button onClick={() => { setPend(null); savePend(null); }}>Bỏ</button></div>}
      {!pend && msg && <div className={'toast ' + (msg[1] ? 'ok' : 'err')} onClick={() => setMsg(null)}>{msg[0]}</div>}
      <main>
        <aside>
          <div className="tabs">
            {(['mine', 'late', 'all'] as const).map(t => <button key={t} className={tab === t ? 'on' : ''} onClick={() => setTab(t)}>
              {t === 'mine' ? 'Việc của tôi' : t === 'late' ? 'Quá hạn' : 'Tất cả'} <i>{cnt[t]}</i></button>)}
          </div>
          <input className="q" placeholder="Tìm mã hoặc nội dung…" value={q} onChange={e => setQ(e.target.value)} />
          <ul>{shown.map(d => <li key={d.documentId} className={sel === d.documentId ? 'on' : ''} onClick={() => pick(d.documentId)}>
            <div className="row1"><code>{d.docNo}</code>{pend && pend.a?.id === d.documentId && <span className={'pill ' + (pend.err ? 'red' : 'amber')}>{pend.err ? '✗ chưa lưu' : '⟳ đồng bộ'}</span>}<span className={'pill ' + ST[d.status][1]}>{ST[d.status][0]}</span></div>
            <div className="ttl">{d.title}</div>
            <div className="row2">rev{String(d.currentRev).padStart(2, '0')} {d.dueAt && <>· hạn {fmt(d.dueAt)}</>} {d.overdueDays > 0 && <b className="late">· trễ {d.overdueDays} ngày</b>}</div>
          </li>)}{!shown.length && <li className="empty">Không có hồ sơ trong mục này.</li>}</ul>
        </aside>
        <section>{v360 ? <Detail key={v360.doc.documentId + v360.doc.status + v360.doc.currentRev} v={v360} isRev={!!isRev} isNt={!!isNt} busy={busy} act={act} /> : <div className="empty">Chọn một hồ sơ bên trái.</div>}</section>
      </main>
      <footer>Bản {who?.build?.version} · {who?.build?.env} · commit {who?.build?.commit} · dữ liệu lúc {fresh} {isDemo && '· XEM TRƯỚC – dữ liệu DEMO, không ghi vào Register'}
        {perf && <span className="perf"> · ⏱ {perf.label}: giao diện {Math.round(perf.ui)} ms · cả vòng {(perf.call / 1000).toFixed(2)} s{perf.server && <> · máy chủ {Object.entries(perf.server).filter(([k]) => k !== 'calls').map(([k, v]) => k + ' ' + v).join(', ')} ms</>}</span>}
        {who?.build?.env !== 'PROD' && who?.roles.includes('ADMIN') && <button className="mini" onClick={() => setBench(true)}>Đo tốc độ</button>}</footer>
      {rep && <Report docs={docs} close={() => setRep(false)} who={who} />}
      {bench && <Bench close={() => setBench(false)} docs={docs} />}
      {newOpen && <NewRfi close={() => setNewOpen(false)} act={act} />}
    </div>
  );
}

function Detail({ v, isRev, isNt, busy, act }: any) {
  const d = v.doc, opts = api.options().RFI!.options;
  const [note, setNote] = useState(''), [file, setFile] = useState(''), [idem] = useState(() => ({ s: newIdem(), r: newIdem(), a: newIdem(), c: newIdem() }));
  const canReply = isRev && d.status === 'SUBMITTED', canEdit = isNt && ['DRAFT', 'NEED_INFO', 'RESUBMIT_REQUIRED'].includes(d.status);
  const draftRev = d.status === 'DRAFT' ? d.currentRev : d.currentRev + 1;
  const draft = v.revisions.find((r: any) => r.revNo === draftRev);
  return <div className="detail">
    <div className="dh"><code>{d.docNo}</code><span className={'pill ' + ST[d.status][1]}>{ST[d.status][0]}</span>{v.overdueDays > 0 && <span className="pill red">Trễ {v.overdueDays} ngày</span>}</div>
    <h2>{d.title}</h2>
    <div className="meta">Đang giữ: <b>{d.holder || '—'}</b> · Hạn: <b>{fmt(d.dueAt)}</b> · Người phụ trách: {d.owner} · ID: <code>{d.documentId}</code></div>

    {canEdit && <div className="box">
      <b>Bước tiếp theo: {d.status === 'DRAFT' ? 'gắn file rồi trình' : 'gắn file bổ sung cho rev' + String(draftRev).padStart(2, '0') + ' rồi trình lại'}</b>
      <div className="line"><input placeholder="Dán link file trên OneDrive / Drive (https://…)" value={file} onChange={e => setFile(e.target.value)} />
        <button disabled={busy || !file} onClick={() => (act('Gắn file', 'attach', idem.a + file.length, { id: d.documentId, file: { name: file.split('/').pop() || 'file', url: file.trim(), source: /1drv|sharepoint|onedrive/.test(file) ? 'ONEDRIVE' : /drive\.google/.test(file) ? 'DRIVE' : 'LINK' } }), setFile(''))}>Gắn</button></div>
      <label className="up">⬆ Tải file từ máy (≤25 MB)<input type="file" disabled={busy} onChange={e => { const f = e.target.files?.[0]; if (f) act('Tải ' + f.name, 'upload', 'job-' + newIdem(), { id: d.documentId }, f); e.target.value = ''; }} /></label>
      <div className="files">{(draft?.files || []).map((f: any) => <a key={f.url} href={f.url} target="_blank">📎 {f.name}</a>)}{!draft?.files?.length && <span className="muted">Chưa có file</span>}</div>
      <button className="pri" disabled={busy || !draft?.files?.length} onClick={() => act('Trình', 'submit', idem.s + '-' + draftRev, { id: d.documentId })}>{d.status === 'DRAFT' ? 'Trình CĐT' : 'Trình lại'}</button>
    </div>}

    {canReply && <div className="box">
      <b>CĐT trả lời – chọn đúng ô trên phiếu</b>
      <textarea placeholder="Ý kiến (bắt buộc với ô 2 và 3)" value={note} onChange={e => setNote(e.target.value)} />
      <div className="line">{opts.map((o: any) => <button key={o.code} disabled={busy} className={o.result === 'CLOSE' ? 'pri' : ''} onClick={() => act('Trả lời ô ' + o.code, 'reply', idem.r + o.code + '-' + d.currentRev, { id: d.documentId, code: o.code, note })}>{o.code} · {o.text}</button>)}</div>
      <div className="muted">Chỉ người có quyền CĐT mới bấm được. AI không trả lời thay.</div>
    </div>}
    {isRev && d.status === 'FORWARDED' && <div className="box"><b>Đã chuyển tiếp – ghi câu trả lời nhận được để đóng</b>
      <textarea value={note} onChange={e => setNote(e.target.value)} />
      <button className="pri" disabled={busy} onClick={() => act('Đóng', 'closeForwarded', idem.c, { id: d.documentId, note })}>Đóng hồ sơ</button></div>}

    <h3>Các lần trình (revision)</h3>
    <table><thead><tr><th>Rev</th><th>Trình</th><th>Trả lời</th><th>Kết quả</th><th>File</th></tr></thead><tbody>
      {v.revisions.map((r: any) => <tr key={r.revNo}><td>rev{String(r.revNo).padStart(2, '0')}</td><td>{fmt(r.submittedAt)}</td><td>{fmt(r.repliedAt)}</td>
        <td>{r.replyCode ? <><b>{r.replyCode}</b> · {r.replyText}{r.replyNote && <div className="muted">“{r.replyNote}”</div>}</> : '—'}</td>
        <td>{r.files.map((f: any) => <a key={f.url} href={f.url} target="_blank">📎 {f.name} </a>)}</td></tr>)}
    </tbody></table>
    <h3>Nhật ký</h3>
    <ol className="tl">{[...v.events].reverse().map((e: any) => <li key={e.eventId} className={e.ok ? '' : 'bad'}><span>{fmt(e.at)}</span> <b>{ACT[e.action] || e.action}</b> · {e.actor}{!e.ok && ' · LỖI ' + JSON.parse(e.detail).error}</li>)}</ol>
  </div>;
}

function NewRfi({ close, act }: any) {
  const [title, setTitle] = useState(''), [file, setFile] = useState(''), [sug, setSug] = useState<any>(null), [idem] = useState(newIdem);
  useEffect(() => { if (file) api.classify(file).then(setSug); else setSug(null); }, [file]);
  return <div className="modal" onClick={close}><div className="card" onClick={e => e.stopPropagation()}>
    <h2>RFI mới</h2>
    <label>Nội dung yêu cầu thông tin<input autoFocus value={title} onChange={e => setTitle(e.target.value)} placeholder="vd: Yêu cầu thông tin gạch porcelain khu WC" /></label>
    <label>Link file (OneDrive / Drive) – có thể gắn sau<input value={file} onChange={e => setFile(e.target.value)} placeholder="https://…" /></label>
    {sug && <div className={'muted ' + (sug.type && sug.type !== 'RFI' ? 'warn' : '')}>Máy đoán loại: <b>{sug.type || 'không rõ'}</b> ({Math.round(sug.confidence * 100)}%) – {sug.reason}{sug.type && sug.type !== 'RFI' && ' · kiểm lại, file có thể không phải RFI'}</div>}
    <div className="muted">Mã hồ sơ do hệ thống cấp khi lưu, không cần gõ.</div>
    <div className="line end"><button onClick={close}>Hủy</button>
      <button className="pri" disabled={title.trim().length < 5} onClick={() => (act('Tạo RFI', 'create', idem, { type: 'RFI', title, files: file ? [{ name: file.split('/').pop() || 'file', url: file.trim(), source: 'LINK' }] : [] }), close())}>Tạo &amp; cấp mã</button></div>
  </div></div>;
}

function Since({ t }: { t: number }) { const [, f] = useState(0); useEffect(() => { const i = setInterval(() => f(x => x + 1), 250); return () => clearInterval(i); }, []); return <span className="muted">{((Date.now() - t) / 1000).toFixed(1)} s</span>; }

/** Đo tốc độ trên DEV: N lần mỗi thao tác, median + p95 (cả vòng ở trình duyệt và từng chặng máy chủ). Hồ sơ đo ghi rõ [ĐO TỐC ĐỘ]. */
function Bench({ close }: any) {
  const [rows, setRows] = useState<any[]>([]), [run, setRun] = useState(''), [n, setN] = useState(10);
  async function go() {
    const res: Record<string, { call: number[]; srv: number[]; parts: Record<string, number[]> }> = {};
    const rec = (k: string, t0: number, r: any) => { const x = (res[k] ||= { call: [], srv: [], parts: {} }); x.call.push(performance.now() - t0); if (r?.ms?.total != null) { x.srv.push(r.ms.total); for (const [p, v] of Object.entries(r.ms)) if (typeof v === 'number' && p !== 'total') (x.parts[p] ||= []).push(v as number); } };
    const show = () => setRows(Object.entries(res).map(([k, v]) => ({ k, n: v.call.length, med: median(v.call), p95: p95(v.call), smed: median(v.srv), sp95: p95(v.srv), parts: Object.entries(v.parts).map(([p, a]) => p + ' ' + Math.round(median(a))).join(', ') })));
    for (let i = 0; i < n; i++) { setRun('Hỏi nhẹ (peek) ' + (i + 1)); let t = performance.now(); rec('Hỏi có thay đổi (peek)', t, await api.peek()); show(); }
    for (let i = 0; i < n; i++) { setRun('Mở web (boot) ' + (i + 1)); let t = performance.now(); rec('Mở web / làm mới (boot)', t, await api.boot()); show(); }
    for (let i = 0; i < n; i++) {
      setRun('Tạo + gắn + trình ' + (i + 1) + '/' + n);
      let t = performance.now(); const c = await api.act('create', newIdem(), { type: 'RFI', title: '[ĐO TỐC ĐỘ] lần ' + (i + 1) + ' – xóa được' }); rec('Tạo hồ sơ + cấp mã', t, c);
      const id = c.result.documentId;
      t = performance.now(); rec('Xem hồ sơ (get360)', t, await api.get360(id));
      t = performance.now(); rec('Gắn file (link)', t, await api.act('attach', newIdem(), { id, file: { name: 'do-toc-do.pdf', url: 'https://example.com/do.pdf', source: 'LINK' } }));
      t = performance.now(); rec('Trình', t, await api.act('submit', newIdem(), { id })); show();
    }
    setRun('Xong');
  }
  return <div className="modal" onClick={close}><div className="card wide" onClick={e => e.stopPropagation()}>
    <h2>Đo tốc độ (median / p95)</h2>
    <div className="muted">Chạy thật trên máy chủ môi trường này. Hồ sơ đo có tiêu đề [ĐO TỐC ĐỘ]. Không chạy trên PROD.</div>
    <div className="line"><label>Số lần <input type="number" min={3} max={20} value={n} onChange={e => setN(+e.target.value)} /></label><button className="pri" disabled={!!run && run !== 'Xong'} onClick={go}>Bắt đầu</button><span className="muted">{run}</span></div>
    <table><thead><tr><th>Thao tác</th><th>n</th><th>Cả vòng median</th><th>p95</th><th>Máy chủ median</th><th>p95</th><th>Chặng máy chủ (median ms)</th></tr></thead>
      <tbody>{rows.map(r => <tr key={r.k}><td>{r.k}</td><td>{r.n}</td><td>{Math.round(r.med)} ms</td><td>{Math.round(r.p95)} ms</td><td>{r.smed ? Math.round(r.smed) + ' ms' : '—'}</td><td>{r.sp95 ? Math.round(r.sp95) + ' ms' : '—'}</td><td className="muted">{r.parts}</td></tr>)}</tbody></table>
    <div className="line end"><button onClick={() => navigator.clipboard?.writeText(JSON.stringify(rows))}>Chép kết quả</button><button onClick={close}>Đóng</button></div>
  </div></div>;
}

/** Báo cáo gửi sếp: tóm tắt + bảng. In / Lưu PDF bằng trình duyệt; tải CSV mở bằng Excel (UTF-8 có BOM). */
function Report({ docs, close, who }: any) {
  const today = new Date(), d7 = new Date(today.getTime() - 7 * 864e5);
  const n = (f: (d: any) => boolean) => docs.filter(f).length;
  const sum = [['Tổng RFI', docs.length], ['Đang chờ CĐT', n((d: any) => d.status === 'SUBMITTED')], ['Quá hạn trả lời', n((d: any) => d.overdueDays > 0)],
    ['Cần bổ sung / trình lại', n((d: any) => ['NEED_INFO', 'RESUBMIT_REQUIRED'].includes(d.status))], ['Đang soạn', n((d: any) => d.status === 'DRAFT')],
    ['Đã đóng', n((d: any) => d.status === 'CLOSED')], ['Mới trong 7 ngày', n((d: any) => new Date(d.createdAt) >= d7)]];
  const ord: Record<string, number> = { SUBMITTED: 0, NEED_INFO: 1, RESUBMIT_REQUIRED: 1, FORWARDED: 2, DRAFT: 3, CLOSED: 4, REJECTED: 5 };
  const rows = [...docs].sort((a, b) => (b.overdueDays - a.overdueDays) || (ord[a.status] - ord[b.status]) || a.docNo.localeCompare(b.docNo));
  function csv() {
    const head = ['Mã hồ sơ', 'Nội dung', 'Tình trạng', 'Rev', 'Bên đang giữ', 'Hạn', 'Trễ (ngày)', 'Người phụ trách', 'Ngày tạo'];
    const q = (v: any) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
    const body = rows.map(d => [d.docNo, d.title, ST[d.status][0], 'rev' + String(d.currentRev).padStart(2, '0'), d.holder, fmt(d.dueAt), d.overdueDays || '', d.owner, fmt(d.createdAt)].map(q).join(','));
    const blob = new Blob(['﻿' + [head.map(q).join(','), ...body].join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'Bao_cao_RFI_' + today.toISOString().slice(0, 10) + '.csv'; a.click();
  }
  return <div className="modal rep-wrap" onClick={close}><div className="card wide report" onClick={e => e.stopPropagation()}>
    <div className="rep-h"><div><b>BÁO CÁO TÌNH HÌNH RFI</b><div className="muted">Lập lúc {today.toLocaleString('vi-VN')} · người lập {who?.email}</div></div>
      <div className="line noprint"><button className="pri" onClick={() => window.print()}>🖨 In / Lưu PDF</button><button onClick={csv}>⤓ Tải Excel (CSV)</button><button onClick={close}>Đóng</button></div></div>
    <div className="kpis">{sum.map(([k, v]) => <div key={k as string} className={'kpi' + ((k === 'Quá hạn trả lời' && (v as number) > 0) ? ' bad' : '')}><b>{v}</b><span>{k}</span></div>)}</div>
    <table><thead><tr><th>Mã hồ sơ</th><th>Nội dung</th><th>Tình trạng</th><th>Rev</th><th>Đang giữ</th><th>Hạn</th><th>Trễ</th></tr></thead>
      <tbody>{rows.map(d => <tr key={d.documentId} className={d.overdueDays > 0 ? 'late-row' : ''}><td><code>{d.docNo}</code></td><td>{d.title}</td><td>{ST[d.status][0]}</td><td>rev{String(d.currentRev).padStart(2, '0')}</td><td>{d.holder || '—'}</td><td>{fmt(d.dueAt)}</td><td>{d.overdueDays > 0 ? d.overdueDays + ' ngày' : ''}</td></tr>)}</tbody></table>
  </div></div>;
}
