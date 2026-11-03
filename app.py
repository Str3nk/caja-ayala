import pandas as pd
from flask import Flask, request, render_template_string, session, redirect, jsonify
import os, json
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'lionfire-secret-2025'
EXCEL_FILE = 'EXCEL CAJA A 2025.xlsx'
USERS_FILE = 'users.json'

def load_users():
    if not os.path.exists(USERS_FILE):
        default = {
            "ayala": {"pass": "lionfire", "role": "admin"},
            "Edward": {"pass": "Edward1", "role": "admin"}
        }
        with open(USERS_FILE, 'w') as f:
            json.dump(default, f)
        return default
    
    with open(USERS_FILE, 'r') as f:
        users = json.load(f)
    
    if "Edward" not in users:
        users["Edward"] = {"pass": "Edward1", "role": "admin"}
        with open(USERS_FILE, 'w') as f:
            json.dump(users, f)
    
    return users

def save_users(users):
    with open(USERS_FILE, 'w') as f:
        json.dump(users, f, indent=2)

def load_data():
    df = pd.read_excel(EXCEL_FILE, dtype=str, sheet_name='DATA')
    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df.fillna('')
    print(f"Excel leído: {len(df)} expedientes")
    return df.to_dict(orient='records')

DATA = load_data()
DATA_VERSION = datetime.now().strftime("%Y-%m-%d %H:%M")
print(f"Cargados {len(DATA)} expedientes - Usuarios: {list(load_users().keys())} - Versión: {DATA_VERSION}")

LOADER = """
<div id="loader-caja"><div class="caja-box"><div class="caja-icon">💰</div><div class="spinner-caja"></div><h2>Caja Ayala</h2><p>Cargando sistema...</p></div></div>
<style>#loader-caja{position:fixed;inset:0;background:#0f172a;display:flex;align-items:center;justify-content:center;z-index:99999;transition:opacity .6s ease}#loader-caja.hide{opacity:0;pointer-events:none}.caja-box{text-align:center;color:#fff;font-family:system-ui}.caja-icon{font-size:52px;animation:float 2s ease-in-out infinite}.spinner-caja{width:44px;height:44px;border:3px solid #1e293b;border-top-color:#10b981;border-radius:50%;margin:16px auto;animation:spin .8s linear infinite}.caja-box h2{margin:0;font-size:22px;font-weight:700}.caja-box p{opacity:.6;font-size:13px;margin-top:6px}@keyframes spin{to{transform:rotate(360deg)}}@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}</style>
<script>window.addEventListener('load',()=>{setTimeout(()=>{const l=document.getElementById('loader-caja');l.classList.add('hide');setTimeout(()=>l.remove(),600)},400)});</script>
"""

HTML_LOGIN = f"""
<html><head><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="manifest" href="/manifest.json">
<meta name="theme-color" content="#0f172a"><style>
body{{font-family:Arial;background:#0f172a;color:white;display:flex;justify-content:center;align-items:center;height:100vh;margin:0}}
.card{{background:#1e293b;padding:30px;border-radius:12px;width:90%;max-width:350px}}
input{{width:100%;padding:12px;margin:8px 0;border-radius:8px;border:none;box-sizing:border-box}}
button{{width:100%;padding:12px;background:#3b82f6;color:white;border:none;border-radius:8px;font-weight:bold;cursor:pointer}}
</style></head><body>{LOADER}<div class="card">
<h2>🌙 Login</h2><form method="post">
<input name="user" placeholder="Usuario" required>
<input name="pass" type="password" placeholder="Clave" required>
<button>Entrar</button>{{% if error %}}<p style="color:#f87171">{{{{error}}}}</p>{{% endif %}}</form>
<p style="font-size:12px;opacity:.7;margin-top:15px;text-align:center">Instala la app: menú ⋮ → Agregar a pantalla de inicio</p>
</div></body></html>
"""

HTML_APP = f"""
<html><head><meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<link rel="manifest" href="/manifest.json"><meta name="theme-color" content="#0f172a"><style>
body{{font-family:Arial;background:#f1f5f9;margin:0;padding:10px}}
.header{{background:#0f172a;color:white;padding:15px;border-radius:12px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}}
.search{{background:white;padding:15px;border-radius:12px;margin-top:15px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}}
input{{width:100%;padding:16px;border:2px solid #e2e8f0;border-radius:12px;font-size:18px;box-sizing:border-box}}
table{{width:100%;border-collapse:collapse;margin-top:15px;background:white;border-radius:12px;overflow:hidden;display:block;overflow-x:auto}}
th{{background:#0f172a;color:white;padding:12px;text-align:left;font-size:12px}} td{{padding:12px;border-bottom:1px solid #e2e8f0;font-size:14px}}
.badge{{padding:6px 10px;border-radius:8px;background:#dbeafe;color:#1e40af;font-weight:bold}}
.btn-admin{{background:#10b981;color:white;padding:8px 14px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:bold}}
.btn-logout{{color:#93c5fd;text-decoration:none;font-size:13px}}
.btn-sync{{background:#3b82f6;color:white;padding:8px 14px;border-radius:8px;border:none;font-size:13px;font-weight:bold;cursor:pointer}}
#result{{font-weight:bold;margin-top:10px}}
#offline-status{{font-size:12px;padding:8px;border-radius:8px;margin-top:10px;display:none}}
.offline{{background:#fef3c7;color:#92400e;display:block !important}}
.online{{background:#d1fae5;color:#065f46;display:block !important}}
tr.row-vigente td{{background:#22c55e !important;color:#ffffff !important}}
tr.row-tramite td{{background:#eab308 !important;color:#422006 !important}}
tr.row-vencida td{{background:#ef4444 !important;color:#ffffff !important}}
tr.row-inactiva td{{background:#94a3b8 !important;color:#1e293b !important}}
tr.row-vigente td .badge, tr.row-tramite td .badge, tr.row-vencida td .badge, tr.row-inactiva td .badge{{background:rgba(255,255,255,0.9) !important;color:#1e293b !important}}
</style></head><body>
{LOADER}
<div class="header">
<div><b>EXPEDIENTES ({{{{count}}}})</b> - Hola {{{{username}}}} {{% if role=='admin' %}}(ADMIN){{% endif %}}<br><small style="opacity:.8">Datos: <span id="fecha-datos">{{{{version}}}}</span> | <span id="modo">Modo: Online</span></small></div>
<div style="display:flex;gap:8px;align-items:center">
<button class="btn-sync" onclick="sincronizar()">🔄 Actualizar</button>
{{% if role=='admin' %}}<a href="/admin" class="btn-admin">👥 Usuarios</a>{{% endif %}}
<a href="/logout" class="btn-logout">Salir</a>
</div>
</div>
<div id="offline-status"></div>
<div class="search"><input id="q" placeholder="🔍 Buscar por RIF, código, razón social, sector, archivo, gaveta o último pago..." onkeyup="filtrar()" autofocus>
<div id="result">Escribe para buscar</div>
<table><thead><tr><th>CODIGO EXPEDIENTE</th><th>RIF</th><th>RAZON SOCIAL</th><th>SECTOR</th><th>ESTATUS</th><th>ARCHIVO</th><th>GAVETA</th><th>ULTIMO PAGO</th><th>RECIBO</th></tr></thead><tbody id="tbody"></tbody></table></div>
<script>
const SERVER_DATA = {{{{data_json|safe}}}};
const SERVER_VERSION = "{{{{version}}}}";
let data = [];

// Cargar datos: primero local, luego servidor
function cargarDatos() {{
  const localData = localStorage.getItem('caja_datos');
  const localVersion = localStorage.getItem('caja_version');
  
  if (localData) {{
    data = JSON.parse(localData);
    document.getElementById('fecha-datos').innerText = localVersion || 'Sin fecha';
    document.getElementById('modo').innerText = 'Modo: Offline (datos guardados)';
    mostrarEstado('offline', '📴 Modo offline: usando datos guardados del ' + (localVersion || ''));
  }} else {{
    data = SERVER_DATA;
    localStorage.setItem('caja_datos', JSON.stringify(data));
    localStorage.setItem('caja_version', SERVER_VERSION);
    document.getElementById('fecha-datos').innerText = SERVER_VERSION;
  }}
}}

function mostrarEstado(tipo, texto) {{
  const el = document.getElementById('offline-status');
  el.className = tipo;
  el.innerText = texto;
  el.style.display = 'block';
}}

function sincronizar() {{
  mostrarEstado('online', '🔄 Buscando actualización...');
  fetch('/datos.json')
    .then(r => r.json())
    .then(res => {{
      const localVersion = localStorage.getItem('caja_version');
      if (res.version !== localVersion) {{
        localStorage.setItem('caja_datos', JSON.stringify(res.data));
        localStorage.setItem('caja_version', res.version);
        data = res.data;
        document.getElementById('fecha-datos').innerText = res.version;
        mostrarEstado('online', '✅ Datos actualizados: ' + res.version + ' (' + res.data.length + ' expedientes)');
        document.getElementById('modo').innerText = 'Modo: Online';
      }} else {{
        mostrarEstado('online', '✅ Ya estás actualizado: ' + res.version);
      }}
    }})
    .catch(() => {{
      mostrarEstado('offline', '📴 Sin internet. Usando datos guardados del ' + (localStorage.getItem('caja_version') || ''));
    }});
}}

// Auto-sincronizar al abrir si hay internet
cargarDatos();
if (navigator.onLine) {{
  setTimeout(sincronizar, 2000);
}}

// Detectar cambio de conexión
window.addEventListener('online', () => {{
  document.getElementById('modo').innerText = 'Modo: Online';
  mostrarEstado('online', '🟢 Conectado. Sincronizando...');
  sincronizar();
}});
window.addEventListener('offline', () => {{
  document.getElementById('modo').innerText = 'Modo: Offline';
  mostrarEstado('offline', '📴 Sin conexión. Modo offline activado.');
}});

function getRowClass(est){{
  const e = String(est||'').toLowerCase().trim();
  if(e.includes('vigente')) return 'row-vigente';
  if(e.includes('trámite') || e.includes('tramite')) return 'row-tramite';
  if(e.includes('vencida')) return 'row-vencida';
  if(e.includes('inactiva') || e.includes('cese')) return 'row-inactiva';
  return 'row-inactiva';
}}
function filtrar(){{
  const q=document.getElementById('q').value.toLowerCase().trim();
  const tbody=document.getElementById('tbody');
  const res=document.getElementById('result');
  tbody.innerHTML='';
  if(!q){{res.innerText='Escribe para buscar';return}}
  const f=data.filter(r=>
    String(r['RIF']).toLowerCase().includes(q)||
    String(r['RAZON SOCIAL']).toLowerCase().includes(q)||
    String(r['CODIGO EXPEDIENTE']).toLowerCase().includes(q)||
    String(r['SECTOR']).toLowerCase().includes(q)||
    String(r['ARCHIVO']).toLowerCase().includes(q)||
    String(r['GAVETA']).toLowerCase().includes(q)||
    String(r['ULTIMO PAGO']).toLowerCase().includes(q)||
    String(r['NUMERO DE RECIBO']).toLowerCase().includes(q)||
    String(r['ESTATUS DE LICENCIA']).toLowerCase().includes(q)
  );
  res.innerText=f.length+' resultados';
  f.forEach(r=>{{
    const tr=document.createElement('tr');
    const est = r['ESTATUS DE LICENCIA'] || '';
    tr.className = getRowClass(est);
    tr.innerHTML=`<td><span class=badge>${{r['CODIGO EXPEDIENTE']}}</span></td><td>${{r['RIF']}}</td><td>${{r['RAZON SOCIAL']}}</td><td>${{r['SECTOR']}}</td><td>${{est}}</td><td>${{r['ARCHIVO']}}</td><td>${{r['GAVETA']}}</td><td>${{r['ULTIMO PAGO']}}</td><td>${{r['NUMERO DE RECIBO']}}</td>`;
    tbody.appendChild(tr);
  }});
}}
if('serviceWorker' in navigator){{navigator.serviceWorker.register('/sw.js')}}
</script></body></html>
"""

HTML_ADMIN = """
<html><head><meta name="viewport" content="width=device-width, initial-scale=1">
<style>body{font-family:Arial;background:#f1f5f9;margin:0;padding:15px}.card{background:white;padding:20px;border-radius:12px;margin-bottom:15px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}input{width:100%;padding:12px;margin:6px 0;border:1px solid #ccc;border-radius:8px;box-sizing:border-box}button{padding:12px 18px;background:#10b981;color:white;border:none;border-radius:8px;font-weight:bold;cursor:pointer}.danger{background:#ef4444;padding:6px 12px;border-radius:6px;color:white;text-decoration:none;font-size:12px}table{width:100%;border-collapse:collapse}th,td{padding:10px;text-align:left;border-bottom:1px solid #eee}a.back{color:#0f172a;text-decoration:none;font-weight:bold}</style></head><body>
<a href="/buscar" class="back">← Volver a Caja</a>
<div class="card"><h2>👥 Administrar Usuarios</h2><p>Solo visible para <b>ayala y Edward</b> (ADMIN)</p>
<form method="post"><h3>Agregar Nuevo Usuario</h3>
<input name="new_user" placeholder="Nuevo usuario ej: Cajero1" required>
<input name="new_pass" placeholder="Clave ej: 1234" required>
<select name="new_role" style="width:100%;padding:12px;margin:6px 0;border-radius:8px"><option value="user">Usuario Normal (solo ver caja)</option><option value="admin">Administrador (puede crear usuarios)</option></select>
<button>Crear Usuario</button>{% if msg %}<p style="color:green;font-weight:bold">{{msg}}</p>{% endif %}</form></div>
<div class="card"><h3>Usuarios Actuales ({{users|length}})</h3>
<table><tr><th>Usuario</th><th>Rol</th><th>Acción</th></tr>
{% for u, info in users.items() %}<tr><td>{{u}}</td><td>{{info.role}}</td><td>{% if u not in ['ayala', 'Edward'] %}<a href="/admin/delete/{{u}}" class="danger" onclick="return confirm('¿Borrar {{u}}?')">Borrar</a>{% else %}Principal{% endif %}</td></tr>{% endfor %}
</table></div></body></html>
"""

@app.route('/manifest.json')
def manifest():
    return {"name":"Caja Ayala","short_name":"Caja","start_url":"/buscar","display":"standalone","background_color":"#0f172a","theme_color":"#0f172a","icons":[{"src":"https://cdn-icons-png.flaticon.com/512/2830/2830284.png","sizes":"512x512","type":"image/png"}]}

@app.route('/sw.js')
def sw():
    sw_code = """
const CACHE_NAME = 'caja-ayala-v2';
const urlsToCache = ['/buscar', '/datos.json'];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(urlsToCache).catch(()=>{});
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
});

self.addEventListener('fetch', event => {
  // Estrategia: Cache primero, luego red
  event.respondWith(
    caches.match(event.request).then(cachedResponse => {
      const fetchPromise = fetch(event.request).then(networkResponse => {
        // Actualizar caché si la respuesta es válida
        if (networkResponse && networkResponse.status === 200) {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch(() => {
        // Si no hay internet, devolver caché
        return cachedResponse;
      });
      
      return cachedResponse || fetchPromise;
    })
  );
});
    """
    return sw_code, 200, {'Content-Type':'application/javascript'}

@app.route('/datos.json')
def datos():
    if not session.get('logged'):
        return redirect('/')
    return jsonify({"data": DATA, "version": DATA_VERSION, "total": len(DATA)})

@app.route('/ping')
@app.route('/health')
def health():
    return {"status": "ok", "expedientes": len(DATA), "version": DATA_VERSION}, 200

@app.route('/', methods=['GET','POST'])
def login():
    if session.get('logged'): return redirect('/buscar')
    error=None
    if request.method=='POST':
        users=load_users()
        u=request.form.get('user','').strip()
        p=request.form.get('pass','').strip()
        if u in users and users[u]['pass']==p:
            session['logged']=True
            session['username']=u
            session['role']=users[u]['role']
            return redirect('/buscar')
        else: error='Usuario o clave incorrecta'
    return render_template_string(HTML_LOGIN, error=error)

@app.route('/buscar')
def buscar():
    if not session.get('logged'): return redirect('/')
    return render_template_string(HTML_APP, data_json=json.dumps(DATA, ensure_ascii=False), count=len(DATA), version=DATA_VERSION, username=session.get('username'), role=session.get('role'))

@app.route('/admin', methods=['GET','POST'])
def admin():
    if not session.get('logged') or session.get('role')!='admin': return redirect('/buscar')
    users=load_users()
    msg=None
    if request.method=='POST':
        nu=request.form.get('new_user','').strip()
        np=request.form.get('new_pass','').strip()
        nr=request.form.get('new_role','user')
        if nu and np:
            if nu in users: msg=f'Usuario {nu} ya existe'
            else:
                users[nu]={"pass":np,"role":nr}
                save_users(users)
                msg=f'Usuario {nu} creado con éxito!'
    users=load_users()
    return render_template_string(HTML_ADMIN, users=users, msg=msg)

@app.route('/admin/delete/<user>')
def delete_user(user):
    if not session.get('logged') or session.get('role')!='admin': return redirect('/buscar')
    users=load_users()
    if user not in ['ayala', 'Edward'] and user in users:
        del users[user]
        save_users(users)
    return redirect('/admin')

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

if __name__=='__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
