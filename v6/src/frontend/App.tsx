import { useEffect, useMemo, useState } from 'react';
import { api, isDemo, newIdem, setDemoUser } from './api';

type Doc = any;
const ST: Record<string, [string, string]> = {
  DRAFT: ['Đang soạn', 'gray'], SUBMITTED: ['Chờ CĐT', 'amber'], NEED_INFO: ['Bổ sung', 'red'], RESUBMIT_REQUIRED: ['Trình lại', 'red'],
  FORWARDED: ['Đã chuyển tiếp', 'blue'], CLOSED: ['Đã đóng', 'green'], REJECTED: ['Từ chối', 'gray'],
};
const fmt = (s?: string | null) => (s ? new Date(s).toLocaleString('vi-VN', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }) : '—');
const ACT: Record<string, string> = { CREATE: 'Tạo hồ sơ', ATTACH: 'Gắn file', SUBMIT: 'Trình', REPLY: 'CĐT trả lời', CLOSE_FORWARDED: 'Đóng sau chuyển tiếp' };

export default function App() {
  const [who, setWho] = useState<{ email: string; roles: string[]; build: any } | null>(null);
  const [docs, setDocs] = useState<Doc[]>([]);
  const [sel, setSel] = useState<string | null>(null);
  const [v360, setV360] = useState<any>(null);
  const [tab, setTab] = useState<'mine' | 'late' | 'all'>('mine');
  const [q, setQ] = useState('');
  const [msg, setMsg] = useState<[string, boolean] | null>(null);
  const [busy, setBusy] = useState(false);
  const [newOpen, setNewOpen] = useState(false);

  const isRev = who?.roles.includes('REVIEWER') || who?.roles.includes('ADMIN');
  const isNt = who?.roles.some(r => ['CONTRACTOR', 'DC', 'ADMIN'].includes(r));

  const [cache] = useState(() => new Map<string, any>());       // hồ sơ đã xem: chuyển qua lại không chờ máy chủ
  const [pending, setPending] = useState('');
  async function load(keep?: string) {
    const b = await api.boot(keep);
    setWho(b.who); setDocs(b.list);
    if (b.v360) { cache.set(b.v360.doc.documentId, b.v360); setSel(b.v360.doc.documentId); setV360(b.v360); }
  }
  useEffect(() => { load().catch(e => setMsg([e.message, false])); }, []);
  function pick(id: string) {
    setSel(id); const c = cache.get(id); if (c) setV360(c);        // hiện ngay bản đã có
    api.get360(id).then(v => { cache.set(id, v); setSel(s => { if (s === id) setV360(v); return s; }); });
  }

  /** Bấm là phản hồi ngay (trạng thái "đang lưu"), máy chủ trả kết quả + dữ liệu mới trong 1 lần gọi. */
  // Trạng thái đoán trước: đổi ngay trên màn hình, máy chủ trả về thì thay bằng số thật; lỗi thì hoàn lại.
  const GUESS: Record<string, string> = { submit: 'SUBMITTED', closeForwarded: 'CLOSED' };
  async function act(label: string, op: string, idem: string, a: any, file?: File) {
    if (busy) return; setBusy(true); setPending(label + '…'); const t0 = performance.now();
    const oldDocs = docs, oldV = v360;
    const st = op === 'reply' ? (a.code === '1' ? 'CLOSED' : a.code === '3' ? 'FORWARDED' : 'NEED_INFO') : GUESS[op];
    if (st && a.id) { setDocs(ds => ds.map(d => d.documentId === a.id ? { ...d, status: st } : d)); if (v360?.doc.documentId === a.id) setV360({ ...v360, doc: { ...v360.doc, status: st }, guess: true }); }
    try {
      const r = file ? await api.upload(idem, a.id, file) : await api.act(op, idem, a);
      setDocs(r.list); if (r.v360) { cache.set(r.v360.doc.documentId, r.v360); setSel(r.v360.doc.documentId); setV360(r.v360); }
      setMsg([label + ' – xong (' + ((performance.now() - t0) / 1000).toFixed(1) + ' giây)' + (r.soRFI ? ' · sổ ' + r.soRFI : ''), !/^Lỗi|Không/.test(r.soRFI || '')]);
    } catch (e: any) { setDocs(oldDocs); setV360(oldV); setMsg([e.message + ' – đã hoàn lại', false]); }
    finally { setBusy(false); setPending(''); }
  }

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
        {isNt && <button className="pri" onClick={() => setNewOpen(true)}>+ RFI mới</button>}
      </header>
      {pending ? <div className="toast wait">{pending}</div> : msg && <div className={'toast ' + (msg[1] ? 'ok' : 'err')} onClick={() => setMsg(null)}>{msg[0]}</div>}
      <main>
        <aside>
          <div className="tabs">
            {(['mine', 'late', 'all'] as const).map(t => <button key={t} className={tab === t ? 'on' : ''} onClick={() => setTab(t)}>
              {t === 'mine' ? 'Việc của tôi' : t === 'late' ? 'Quá hạn' : 'Tất cả'} <i>{cnt[t]}</i></button>)}
          </div>
          <input className="q" placeholder="Tìm mã hoặc nội dung…" value={q} onChange={e => setQ(e.target.value)} />
          <ul>{shown.map(d => <li key={d.documentId} className={sel === d.documentId ? 'on' : ''} onClick={() => pick(d.documentId)}>
            <div className="row1"><code>{d.docNo}</code><span className={'pill ' + ST[d.status][1]}>{ST[d.status][0]}</span></div>
            <div className="ttl">{d.title}</div>
            <div className="row2">rev{String(d.currentRev).padStart(2, '0')} {d.dueAt && <>· hạn {fmt(d.dueAt)}</>} {d.overdueDays > 0 && <b className="late">· trễ {d.overdueDays} ngày</b>}</div>
          </li>)}{!shown.length && <li className="empty">Không có hồ sơ trong mục này.</li>}</ul>
        </aside>
        <section>{v360 ? <Detail key={v360.doc.documentId + v360.doc.status + v360.doc.currentRev} v={v360} isRev={!!isRev} isNt={!!isNt} busy={busy} act={act} /> : <div className="empty">Chọn một hồ sơ bên trái.</div>}</section>
      </main>
      <footer>Bản {who?.build?.version} · {who?.build?.env} · commit {who?.build?.commit} {isDemo && '· XEM TRƯỚC – dữ liệu DEMO, không ghi vào Register'}</footer>
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
      <label className="up">⬆ Tải file từ máy (≤25 MB)<input type="file" disabled={busy} onChange={e => { const f = e.target.files?.[0]; if (f) act('Tải ' + f.name, 'attach', idem.a + '-up-' + f.name + f.size, { id: d.documentId }, f); e.target.value = ''; }} /></label>
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
      <button className="pri" disabled={title.trim().length < 5} onClick={() => (close(), act('Tạo RFI', 'create', idem, { type: 'RFI', title, files: file ? [{ name: file.split('/').pop() || 'file', url: file.trim(), source: 'LINK' }] : [] }))}>Tạo &amp; cấp mã</button></div>
  </div></div>;
}
