"""Gộp index.template.html + data + code + sidebar thành 1 file site_single.html (không fetch)."""
import json, os
H = os.path.dirname(os.path.abspath(__file__))
s = open(f'{H}/index.template.html', encoding='utf-8').read()
data = open(f'{H}/site/data.json', encoding='utf-8').read(); code = open(f'{H}/site/tqn.js', encoding='utf-8').read(); side = open(f'{H}/site/sidebar.html', encoding='utf-8').read()
safe = lambda t: t.replace('</script', '<\\/script')
s = s.replace('<script src="tqn.js"></script>', '<script>\n' + safe(code) + '\n</script>')
s = s.replace("const data=await (await fetch('data.json')).json();", "const data=JSON.parse(document.getElementById('tqn-data').textContent);")
s = s.replace('<script>\n/* ===== MD5', '<script>window.addEventListener("error",e=>{const b=document.getElementById("err");if(b){b.style.display="block";b.textContent="Lỗi trang: "+e.message}});</script>\n<script type="application/json" id="tqn-data">' + safe(data) + '</script>\n<script>\n/* ===== MD5', 1)
s = s.replace('<div class="menus" id="menus">', '<div id="err" style="display:none;background:#fde8e8;color:#9b1c1c;padding:6px 12px;font-size:12px"></div>\n<div class="menus" id="menus">')
s = s.replace(" window.__ready=true})();", " window.__ready=true})().catch(e=>{const b=document.getElementById('err');b.style.display='block';b.textContent='Lỗi khởi động: '+(e.message||e)});")
assert "fetch(" not in s.split('/* ===== Giao diện')[1]
open(f'{H}/site_single.html', 'w', encoding='utf-8').write(s); print(len(s) // 1024, 'KB')
