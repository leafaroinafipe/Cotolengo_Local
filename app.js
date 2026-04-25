// ============================================================
//  NurseShift Pro — app.js v3.0
//  Regras de negócio: algoritmo de escala em 4 fases
//  Todas as solicitações são aprovadas automaticamente
// ============================================================

// ── CONFIG ────────────────────────────────────────────────────
const COORD_PASS  = 'coord2026';
const NURSE_PASS  = 'enfermeira123';
const GOOGLE_API_URL = 'https://script.google.com/macros/s/AKfycbwsfN_I_dP8H6C7odQDPeppoecyiUPAtdo6_P3bBgIj_vfMULKX6Qm5XyZB4P2zmYWiqQ/exec';
const API_KEY = 'cotolengo_2026_secure_key';

// ── UTILS: ID único ──────────────────────────────────────────
function generateId() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
    return Date.now() + '-' + Math.random().toString(36).substring(2, 9);
}

// ── CACHE TTL (evita cold start a cada abertura do app) ──────
const CACHE_PREFIX = 'nurseshift_cache_';
const CACHE_TTL = {
    Funcionarios: 60 * 60 * 1000,
    Escala:       10 * 60 * 1000,
    Solicitacoes: 30 * 1000,
    _default:      5 * 60 * 1000
};
function cacheKey(name) { return CACHE_PREFIX + name; }
function cacheGet(name) {
    try {
        const raw = localStorage.getItem(cacheKey(name));
        if (!raw) return null;
        const obj = JSON.parse(raw);
        if (!obj || !obj.t) return null;
        const ttl = CACHE_TTL[name] || CACHE_TTL._default;
        const age = Date.now() - obj.t;
        return { data: obj.d, stale: age > ttl, age };
    } catch (e) { return null; }
}
function cacheSet(name, data) {
    try { localStorage.setItem(cacheKey(name), JSON.stringify({ t: Date.now(), d: data })); }
    catch (e) { console.warn('cacheSet failed', e); }
}
function cacheInvalidate(name) {
    try { localStorage.removeItem(cacheKey(name)); } catch (e) {}
}

// ── UTILS: Sanitiza datas vindas do cloud (remove parte de horário) ──────────
// Corrige bug de datas no formato "2026-05-23T03:00:00Z" retornadas pelo Google Sheets
// que quebrava tanto o display ("23T03:00:00Z/05/2026") quanto o parser da geração de escala.
function sanitizeDate(d) {
    if (!d || typeof d !== 'string') return '';
    return d.split('T')[0]; // "2026-05-23T03:00:00Z" → "2026-05-23"
}

const SHIFTS = {
    'M1': { name: 'Mattina 1', h: 7.0, color: '#f59e0b', text: '#1a1a00', period:'morning' },
    'M2': { name: 'Mattina 2', h: 4.5, color: '#fcd34d', text: '#1a1a00', period:'morning' },
    'MF': { name: 'Mattina Festivo', h: 7.5, color: '#f97316', text: '#fff', period:'morning' },
    'G':  { name: 'Giornata Intera', h: 8, color: '#0ea5e9', text: '#fff', period:'morning' },
    'P':  { name: 'Pomeriggio', h: 8.5, color: '#8b5cf6', text: '#fff', period:'afternoon' },
    'PF': { name: 'Pomeriggio Festivo', h: 10, color: '#a78bfa', text: '#fff', period:'afternoon' },
    'N':  { name: 'Notte', h: 9, color: '#1e1b4b', text: '#fff', period:'night' },
    'OFF':{ name: 'Riposo', h: 0, color: 'rgba(255,255,255,0.03)', text: 'rgba(255,255,255,0.2)', period:'off' },
    'FE': { name: 'Ferie', h: 0, color: '#10b981', text: '#fff', period:'off' },
    'AT': { name: 'Certificato/Licenza', h: 0, color: '#ef4444', text: '#fff', period:'off' },
};

// Hora de início de cada turno (para regra de progressão)
const SHIFT_START = { M1:7, M2:7.5, MF:7, G:7.5, P:14, PF:14.5, N:22, OFF:0, FE:0, AT:0 };

let NURSES = [
    { id:'n1', name:'Balla Sabina',        initials:'BS', nightQuota:5 },
    { id:'n2', name:'Batista Bianca',      initials:'BB', nightQuota:5 },
    { id:'n3', name:'De Carvalho Eduarda', initials:'CE', nightQuota:5 },
    { id:'n4', name:'Festa Alves Melissa', initials:'FM', nightQuota:5 },
    { id:'n5', name:'Delizzeti Sirlene',   initials:'DS', nightQuota:5 },
    { id:'n6', name:'Moslih Miriam',       initials:'MM', nightQuota:5 },
    { id:'n7', name:'Kocevska Kristina',   initials:'KK', nightQuota:5 }
];

// ── STATE ─────────────────────────────────────────────────────
let currentUser   = null;
let isCoordinator = false;
let schedule      = {};   // key: `${nurseId}_${month}_${year}_${day}` → shiftCode
let requests      = [];
let currentMonth  = new Date();
currentMonth.setDate(1); // Força dia 1 para evitar pular mês em dias 31
let selectedCell  = null;
let occurrences = []; // stores { id, nurseId, type, start, end, desc }
let monthlyOrder  = {}; // key: `${month}_${year}` → array of nurseIds


// ── UTILS ─────────────────────────────────────────────────────
function key(nurseId, day) { return `${nurseId}_${currentMonth.getMonth()}_${currentMonth.getFullYear()}_${day}`; }
function daysInMonth(m) { return new Date(m.getFullYear(), m.getMonth()+1, 0).getDate(); }
function isWeekend(m, day) { const d = new Date(m.getFullYear(), m.getMonth(), day); return d.getDay()===0||d.getDay()===6; }
function getShift(nurseId, day) { return schedule[key(nurseId,day)] || 'OFF'; }

function shiftCount(nurseId, codes) {
    if (!Array.isArray(codes)) codes = [codes];
    let c = 0;
    for (let d=1; d<=daysInMonth(currentMonth); d++) {
        if (codes.includes(schedule[key(nurseId, d)])) c++;
    }
    return c;
}

function getMonthlyNurses() {
    let k = `${currentMonth.getMonth()}_${currentMonth.getFullYear()}`;
    if (!monthlyOrder[k]) {
        let bestOrder = null;
        let targetVal = currentMonth.getFullYear() * 12 + currentMonth.getMonth();
        let maxPast = -1;
        for (let mo in monthlyOrder) {
            let [mStr, yStr] = mo.split('_');
            let val = parseInt(yStr) * 12 + parseInt(mStr);
            if (val < targetVal && val > maxPast) {
                maxPast = val;
                bestOrder = monthlyOrder[mo];
            }
        }
        monthlyOrder[k] = bestOrder ? [...bestOrder] : NURSES.map(n => n.id);
        saveData();
    }
    return monthlyOrder[k].map(id => NURSES.find(n => n.id === id)).filter(Boolean);
}

// ── PERSISTENCE ───────────────────────────────────────────────
let _saveTimer = null;
function saveData() {
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(_doSave, 500);
}
function _doSave() {
    localStorage.setItem('escala_nurses', JSON.stringify(NURSES));
    localStorage.setItem('escala_schedule', JSON.stringify(schedule));
    localStorage.setItem('escala_occurrences', JSON.stringify(occurrences));
    localStorage.setItem('escala_monthlyOrder', JSON.stringify(monthlyOrder));
    localStorage.setItem('escala_requests', JSON.stringify(requests));
}

function loadData() {
    try {
        const nr = localStorage.getItem('escala_nurses');
        const s = localStorage.getItem('escala_schedule');
        const o = localStorage.getItem('escala_occurrences');
        const m = localStorage.getItem('escala_monthlyOrder');
        const r = localStorage.getItem('escala_requests');
        if (nr) NURSES = JSON.parse(nr);
        if (s) schedule = JSON.parse(s);
        if (o) occurrences = JSON.parse(o);
        if (m) monthlyOrder = JSON.parse(m);
        if (r) requests = JSON.parse(r);
    } catch(e) { console.error('Erro ao carregar dados locais', e); }
}

// ── GOOGLE SHEETS API (CLOUD DB) ──────────────────────────────
async function fetchGoogleDB(action, sheetName, dataObject = null, opts = {}) {
    if (!GOOGLE_API_URL) {
        console.info('Aviso: Operando apenas no Banco Local (Localstorage). URL da Nuvem não fornecida no app.js.');
        return null;
    }
    // Cache hit para reads, a menos que forceFresh
    if (action === 'read' && opts.useCache !== false && !opts.forceFresh) {
        const cached = cacheGet(sheetName);
        if (cached && !cached.stale) {
            return { status: 'success', data: cached.data, _cached: true };
        }
    }
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 12000); // 12s timeout
    try {
        if (action === 'read') {
            const url = `${GOOGLE_API_URL}?action=${action}&sheetName=${sheetName}&apiKey=${API_KEY}`;
            const response = await fetch(url, { signal: controller.signal });
            clearTimeout(timeoutId);
            const json = await response.json();
            if (json && json.status === 'success' && Array.isArray(json.data)) cacheSet(sheetName, json.data);
            return json;
        } else if (action === 'write') {
            const url = `${GOOGLE_API_URL}?action=${action}&sheetName=${sheetName}&apiKey=${API_KEY}`;
            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(dataObject || {}),
                signal: controller.signal
            });
            clearTimeout(timeoutId);
            cacheInvalidate(sheetName);
            return await response.json();
        }
    } catch (error) {
        clearTimeout(timeoutId);
        // Stale-while-error fallback
        if (action === 'read') {
            const cached = cacheGet(sheetName);
            if (cached) {
                console.warn(`[CACHE] Usando dados antigos de ${sheetName} (erro de rede)`);
                return { status: 'success', data: cached.data, _cached: true, _stale: true };
            }
        }
        if (error.name === 'AbortError') {
            console.warn("Timeout na comunicação com a Base Google — modo offline ativado.");
            toast('Tempo limite excedito. Modo offline ativo.', 'warning');
        } else {
            console.error("Erro na comunicação com a Base Google:", error);
            toast('Conexão instável com Banco Nuvem.', 'warning');
        }
        return null;
    }
}

// ── PUBLICAÇÃO NA NUVEM (para o App Mobile) ──────────────────
// ⚠️ Nunca usar clearAll:true em Solicitacoes (perderia requests do mobile).
//    1) Pull antes de push (merge não-destrutivo)
//    2) Escala: clearFilter por mês (não afeta outros meses)
//    3) Solicitações: update por id (insert se ausente)
async function publishToCloud() {
    const btn = document.getElementById('publishBtn');
    const originalText = btn.innerHTML;
    btn.innerHTML = '<span class="spinner" style="width:14px;height:14px;margin:0;border-width:2px;display:inline-block;"></span> Publicando...';
    btn.style.pointerEvents = 'none';

    let ok = { nurses: false, escala: false, requests: false };

    try {
        // Pré-sync: puxa solicitações do mobile
        await syncRequestsFromCloud();

        const m = currentMonth.getMonth();
        const y = currentMonth.getFullYear();
        const days = daysInMonth(currentMonth);
        const displayNurses = getMonthlyNurses();

        // 0. Setup headers para Escala
        try {
            const setupEscala = `${GOOGLE_API_URL}?action=setupHeaders&sheetName=Escala&apiKey=${API_KEY}`;
            await fetch(setupEscala, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ headers: ['nurseId','month','year','d1','d2','d3','d4','d5','d6','d7','d8','d9','d10','d11','d12','d13','d14','d15','d16','d17','d18','d19','d20','d21','d22','d23','d24','d25','d26','d27','d28','d29','d30','d31'] })
            });
        } catch(e) { console.warn('setupHeaders falhou', e); }

        // 1. Publicar Funcionários
        try {
            const nursesRows = NURSES.map(n => ({
                ID_Funcionario: n.id,
                Nome: n.name,
                Turno_Padrao: '',
                Carga_Horaria_Mensal: n.nightQuota || ''
            }));
            const funcUrl = `${GOOGLE_API_URL}?action=bulkWrite&sheetName=Funcionarios&apiKey=${API_KEY}`;
            await fetch(funcUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ clearAll: true, rows: nursesRows })
            });
            cacheInvalidate('Funcionarios');
            ok.nurses = true;
        } catch(e) { console.error('Publicar Funcionarios falhou', e); }

        // 2. Publicar Escala do mês atual
        try {
            const escalaRows = displayNurses.map(nurse => {
                const row = { nurseId: nurse.id, month: String(m + 1), year: String(y) };
                for (let d = 1; d <= 31; d++) {
                    row['d' + d] = d <= days ? getShift(nurse.id, d) : '';
                }
                return row;
            });
            const escalaUrl = `${GOOGLE_API_URL}?action=bulkWrite&sheetName=Escala&apiKey=${API_KEY}`;
            await fetch(escalaUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({
                    clearFilter: { column: 'month', value: String(m + 1) },
                    rows: escalaRows
                })
            });
            cacheInvalidate('Escala');
            ok.escala = true;
        } catch(e) { console.error('Publicar Escala falhou', e); }

        // 3. Publicar Solicitações — update por id (nunca clearAll)
        try {
            const updateUrl = `${GOOGLE_API_URL}?action=update&sheetName=Solicitacoes&apiKey=${API_KEY}`;
            const jobs = requests.map(r => {
                const row = {
                    _keyColumn: 'id',
                    _keyValue: String(r.id),
                    id: String(r.id),
                    type: r.type,
                    status: r.status,
                    nurseId: r.nurseId || r.fromNurseId || '',
                    nurseName: r.nurseName || r.fromNurseName || '',
                    startDate: r.startDate || r.date || '',
                    endDate: r.endDate || r.startDate || r.date || '',
                    desc: r.desc || r.reason || '',
                    createdAt: r.createdAt || '',
                    approvedAt: r.approvedAt || '',
                    approvedBy: r.approvedBy || ''
                };
                return fetch(updateUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify(row)
                });
            });
            const results = await Promise.allSettled(jobs);
            const failed = results.filter(r => r.status === 'rejected').length;
            cacheInvalidate('Solicitacoes');
            ok.requests = failed === 0;
            if (failed > 0) console.warn(`[SYNC] ${failed}/${jobs.length} solicitações falharam`);
        } catch(e) { console.error('Publicar Solicitacoes falhou', e); }

        if (ok.nurses && ok.escala && ok.requests) {
            toast('☁️ Turni pubblicati in Cloud! I dipendenti possono già vederli nell\'App.', 'success', 5000);
        } else {
            const parts = [];
            if (!ok.nurses)   parts.push('Infermiere');
            if (!ok.escala)   parts.push('Scala');
            if (!ok.requests) parts.push('Richieste');
            toast(`⚠️ Pubblicazione parziale: falliti ${parts.join(', ')}. Riprova.`, 'warning', 6000);
        }
    } catch (error) {
        console.error('Erro ao publicar na nuvem:', error);
        toast('Errore nella pubblicazione nel cloud. Verifica la connessione.', 'error');
    }

    btn.innerHTML = originalText;
    btn.style.pointerEvents = '';
}

// Toast
function toast(msg, type='success', dur=3000) {
    const icons = { success:'✅', error:'❌', warning:'⚠️', info:'ℹ️' };
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.innerHTML = `<span class="toast-icon">${icons[type]||'•'}</span><span>${msg}</span>`;
    document.getElementById('toastContainer').appendChild(el);
    setTimeout(()=>{ el.classList.add('hiding'); setTimeout(()=>el.remove(), 350); }, dur);
}

// ── LOGIN / AUTH ───────────────────────────────────────────────
// ── SPLASH SCREEN / AUTH ──────────────────────────────────────
function initApp() {
    isCoordinator = true;
    currentUser   = { id:'coordinator', name:'Coordinatrice', initials:'C' };
    
    setTimeout(()=>{
        document.getElementById('splashScreen').style.opacity = '0';
        setTimeout(()=>{
            document.getElementById('splashScreen').classList.remove('active');
            document.getElementById('mainApp').classList.add('active');
            
            document.getElementById('userInfo').textContent = currentUser.name;
            document.getElementById('userAvatar').textContent = currentUser.initials;
            document.getElementById('generateBtn').style.display = 'flex';
            document.getElementById('clearBtn').style.display = 'flex';
            const manageBtn = document.getElementById('manageNursesBtn');
            if(manageBtn) manageBtn.style.display = 'flex';
            
            buildLegend();
            updateMonthDisplay();
            renderCalendar();
            populateOccNurses();
            renderOccurrences();
            renderRequests();
            
            // TESTE DE CONEXÃO COM A NUVEM E SINCPRONIZAÇÃO INICIAL
            setTimeout(async () => {
                const dbTest = await fetchGoogleDB('read', 'Funcionarios');
                const pDot = document.getElementById('cloudStatusDot');
                const pTxt = document.getElementById('cloudStatusText');
                
                if (dbTest && dbTest.status === 'success') {
                    if (dbTest.data && dbTest.data.length > 0) {
                        try {
                            let mergedNurses = false;
                            dbTest.data.forEach(cn => {
                                const id = String(cn.ID_Funcionario || cn.id || '');
                                const name = String(cn.Nome || cn.name || '');
                                const quota = parseInt(cn.Carga_Horaria_Mensal || cn.nightQuota) || 5;
                                if (!id || id === 'undefined' || !name || name === 'undefined') return;
                                
                                const localNurse = NURSES.find(n => String(n.id) === id);
                                if (!localNurse) {
                                    NURSES.push({
                                        id: id,
                                        name: name,
                                        initials: name.split(' ').map(w => w[0] || '').join('').substring(0,2).toUpperCase(),
                                        nightQuota: quota
                                    });
                                    let targetVal = currentMonth.getFullYear() * 12 + currentMonth.getMonth();
                                    for (let mo in monthlyOrder) {
                                        let [mStr, yStr] = mo.split('_');
                                        let val = parseInt(yStr) * 12 + parseInt(mStr);
                                        if (val >= targetVal) {
                                            if (!monthlyOrder[mo].includes(id)) {
                                                monthlyOrder[mo].push(id);
                                            }
                                        }
                                    }
                                    let currK = `${currentMonth.getMonth()}_${currentMonth.getFullYear()}`;
                                    if (!monthlyOrder[currK]) {
                                        monthlyOrder[currK] = NURSES.map(n => n.id);
                                    } else if (!monthlyOrder[currK].includes(id)) {
                                        monthlyOrder[currK].push(id);
                                    }
                                    mergedNurses = true;
                                } else {
                                    // Atualiza dados alterados
                                    if(localNurse.name !== name || localNurse.nightQuota !== quota){
                                        localNurse.name = name;
                                        localNurse.initials = name.split(' ').map(w => w[0] || '').join('').substring(0,2).toUpperCase();
                                        localNurse.nightQuota = quota;
                                        mergedNurses = true;
                                    }
                                }
                            });
                            if (mergedNurses) {
                                saveData();
                                renderCalendar();
                                populateOccNurses();
                            }
                        } catch(e) { console.error('Erro na sincronização de funcionários:', e); }
                    }

                    if (pDot) pDot.style.background = 'var(--success)';
                    if (pTxt) { pTxt.textContent = 'App Sincronizzato'; pTxt.style.color = 'var(--text)'; }
                    
                    toast('🟢 Sistema Online connesso al Database cloud!', 'success', 3500);
                    
                    // Dispara a sincronização de turnos pendentes e requisições
                    await syncScheduleFromCloud();
                    await syncRequestsFromCloud();
                    renderRequests();
                    renderOccurrences();
                } else {
                    if (pDot) pDot.style.background = 'var(--danger)';
                    if (pTxt) { pTxt.textContent = 'Offline (Errore)'; pTxt.style.color = '#fff'; }
                    console.warn("[CLOUD DB] Falha no teste inicial:", dbTest);
                    toast('🔴 Modalità offline attiva. Connessione persa.', 'warning', 5000);
                }
            }, 600);

        }, 500); 
    }, 2200); 
}

function logout() {
    location.reload();
}

// ── LEGEND ────────────────────────────────────────────────────
function buildLegend() {
    const codes = ['M1','M2','MF','G','P','PF','N','OFF','FE','AT'];
    document.getElementById('shiftLegend').innerHTML = codes.map(c => {
        const s = SHIFTS[c];
        return `<div class="legend-item"><div class="legend-dot" style="background:${s.color};border:1px solid rgba(0,0,0,.1)"></div>${s.name}</div>`;
    }).join('');
}

// ── NAVIGATION ────────────────────────────────────────────────
let _lastReqSync = 0;
let _activeTab = 'schedule';
let _requestsPollTimer = null;

function startRequestsPolling() {
    stopRequestsPolling();
    // Polling leve (30s) apenas quando a aba "requests" está ativa e a página
    // visível. Se o tab ficar escondido (usuário minimizou/abriu outra aba do
    // sistema), paramos de poll para poupar quota de API.
    _requestsPollTimer = setInterval(() => {
        if (document.hidden) return;
        if (_activeTab !== 'requests') return;
        syncRequestsFromCloud().then(() => {
            renderRequests();
            renderOccurrences();
        }).catch(()=>{});
    }, 30000);
}
function stopRequestsPolling() {
    if (_requestsPollTimer) { clearInterval(_requestsPollTimer); _requestsPollTimer = null; }
}

// Botão "🔄 Aggiorna" da barra de requests — força sincronização com o cloud
// (ignora cache) e processa a fila de pendingDeletes.
async function manualRetrySync(ev) {
    if (ev && ev.preventDefault) ev.preventDefault();
    try {
        toast('🔄 Sincronizzando...', 'info', 1200);
        await retryPendingDeletes();
        cacheInvalidate('Solicitacoes');
        _lastReqSync = Date.now();
        await syncRequestsFromCloud();
        renderRequests();
        renderOccurrences();
        toast('✅ Richieste aggiornate', 'success', 1500);
    } catch (e) {
        toast('⚠️ Errore: ' + (e.message || e), 'error', 2500);
    }
}

function showTab(tab, btn) {
    document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
    document.getElementById(tab+'Tab').classList.add('active');
    btn.classList.add('active');
    _activeTab = tab;
    if (tab==='reports') renderReports();
    if (tab==='requests') {
        const now = Date.now();
        if (now - _lastReqSync > 60000) { // só sincroniza se passaram mais de 60s
            _lastReqSync = now;
            syncRequestsFromCloud().then(() => {
                renderRequests();
                renderOccurrences();
            });
        } else {
            renderRequests();
            renderOccurrences();
        }
        startRequestsPolling();
    } else {
        stopRequestsPolling();
    }
}

// Sincroniza solicitações criadas no app mobile para o sistema local
async function syncRequestsFromCloud() {
    try {
        const cloudResult = await fetchGoogleDB('read', 'Solicitacoes');
        if (cloudResult && cloudResult.status === 'success' && cloudResult.data) {
            const cloudRequests = cloudResult.data;
            const cloudIds = new Set(cloudRequests.map(cr => String(cr.id)));
            
            let updatedRequests = [];
            let isModified = false;
            
            // 1. Process cloud requests (update existing, add new)
            cloudRequests.forEach(cr => {
                const crId = String(cr.id);
                const local = requests.find(r => String(r.id) === crId);
                
                if (local) {
                    if (local.status !== cr.status) {
                        local.status = cr.status;
                        local.approvedAt = cr.approvedAt || local.approvedAt;
                        local.approvedBy = cr.approvedBy || local.approvedBy;
                        isModified = true;
                    }
                    updatedRequests.push(local);
                } else {
                    // Sanitiza datas vindas do cloud: Google Sheets pode retornar
                    // "2026-05-23T03:00:00Z" — o split('T')[0] garante "2026-05-23"
                    const rawDate = sanitizeDate(cr.startDate || cr.date || '');
                    updatedRequests.push({
                        id: crId,
                        type: cr.type || 'OFF',
                        status: cr.status || 'pending',
                        nurseId: cr.nurseId || '',
                        fromNurseId: cr.nurseId || '',
                        nurseName: cr.nurseName || '',
                        fromNurseName: cr.nurseName || '',
                        startDate: rawDate,
                        date: rawDate,
                        endDate: sanitizeDate(cr.endDate || ''),
                        desc: cr.desc || '',
                        reason: cr.desc || '',
                        createdAt: cr.createdAt || new Date().toISOString(),
                        approvedAt: cr.approvedAt || '',
                        approvedBy: cr.approvedBy || ''
                    });
                    isModified = true;
                }
            });
            
            // 2. Check if we need to drop any local requests that no longer exist in cloud
            if (requests.length !== updatedRequests.length) {
                isModified = true;
            }
            
            if (isModified) {
                requests = updatedRequests; // Replace entirely to drop ghost requests
                saveData();
                updateBadge();
                toast('☁️ Richieste sincronizzate e allineate col cloud', 'info');
            }
        }
    } catch (e) {
        console.warn('Sync cloud requests:', e);
    }
}

// ── MONTH NAV E CLOUD DOWNLOAD AUTOMÁTICO ─────────────────
function changeMonth(d) {
    currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth()+d, 1);
    updateMonthDisplay();
    renderCalendar();

    // Auto-syncroniza quando a coordenadora visualiza um mês:
    syncScheduleFromCloud();

    // Se o dashboard de relatórios estiver aberto, re-renderiza para refletir
    // as métricas do novo mês (carga por enfermeira, cobertura diária etc.).
    if (_activeTab === 'reports') renderReports();
}

async function syncScheduleFromCloud() {
    try {
        const pTxt = document.getElementById('cloudStatusText');
        if(pTxt) pTxt.textContent = 'Sincronizzando...';

        const m = currentMonth.getMonth(); // 0 a 11 base JS
        const y = currentMonth.getFullYear();

        const dbRes = await fetchGoogleDB('read', 'Escala');
        if (dbRes && dbRes.status === 'success' && dbRes.data) {
            // Bug #3 — Sempre que sincronizamos do cloud, o cloud É a fonte da verdade
            // para o mês visualizado. Se um turno foi REMOVIDO no mobile (virando vazio
            // em Sheets), a chave local correspondente precisa sumir também. Antes,
            // apenas re-populávamos por cima (merge), deixando turnos "fantasmas".
            const suffix = `_${m}_${y}_`;
            Object.keys(schedule).forEach(k => {
                // chave: nurseId_month_year_day
                const parts = k.split('_');
                if (parts.length >= 4) {
                    const keyMonth = parts[parts.length - 3];
                    const keyYear  = parts[parts.length - 2];
                    if (keyMonth === String(m) && keyYear === String(y)) {
                        delete schedule[k];
                    }
                }
            });

            let loaded = 0;
            const sheetMonthValue = String(m + 1); // 1 a 12 base Sheets Visual

            dbRes.data.forEach(row => {
                if (String(row.month) === sheetMonthValue && String(row.year) === String(y)) {
                    for (let d = 1; d <= 31; d++) {
                        let shiftCode = row['d' + d];
                        if (shiftCode) {
                            schedule[`${row.nurseId}_${m}_${y}_${d}`] = shiftCode;
                            loaded++;
                        }
                    }
                }
            });
            saveData();
            renderCalendar();
            if (loaded > 0) {
                toast(`☁️ Turni del mese scaricati in base al cloud!`, 'info', 2000);
            }
            if(pTxt) pTxt.textContent = 'App Sincronizzato';
        }
    } catch (e) {
        console.error("Errore download turni", e);
        const pTxt = document.getElementById('cloudStatusText');
        if (pTxt) pTxt.textContent = 'Errore sincronizzazione';
    }
}


function updateMonthDisplay() {
    const months = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
    const lbl = `${months[currentMonth.getMonth()]} ${currentMonth.getFullYear()}`;
    document.getElementById('monthYear').textContent = lbl;
    const rm = document.getElementById('reportMonth');
    if (rm) rm.textContent = lbl;
}

// ── CALENDAR RENDER ───────────────────────────────────────────
function renderCalendar() {
    // saveData movido para pontos de mutação de estado específicos
    
    const days   = daysInMonth(currentMonth);
    const dayNames = ['D','L','M','M','G','V','S'];

    // Header — construído como string para evitar O(n²) de reflow
    const hdRow = document.getElementById('calendarDays');
    let headerHtml = `<th class="nurse-cell">Infermiera</th>`;
    for (let d=1; d<=days; d++) {
        const wk = isWeekend(currentMonth, d);
        const dow = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), d).getDay();
        headerHtml += `<th class="${wk?'wkend':''}">
            <div class="day-num">${d}</div>
            <div class="day-lbl">${dayNames[dow]}</div>
        </th>`;
    }
    hdRow.innerHTML = headerHtml;

    // Body
    const tbody = document.getElementById('calendarBody');
    tbody.innerHTML = '';
    
    const displayNurses = getMonthlyNurses();
    
    displayNurses.forEach((nurse, idx) => {
        const tr = document.createElement('tr');
        tr.draggable = true;
        
        tr.ondragstart = (e) => {
            e.dataTransfer.setData('text/plain', nurse.id);
            e.dataTransfer.effectAllowed = 'move';
            tr.style.opacity = '0.4';
        };
        tr.ondragend = (e) => {
            tr.style.opacity = '1';
        };
        tr.ondragover = (e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            tr.style.outline = '2px dashed var(--primary-color)';
            tr.style.outlineOffset = '-2px';
        };
        tr.ondragleave = (e) => {
            tr.style.outline = '';
        };
        tr.ondrop = (e) => {
            e.preventDefault();
            tr.style.outline = '';
            const draggedId = e.dataTransfer.getData('text/plain');
            if (draggedId && draggedId !== nurse.id) {
                swapNurses(draggedId, nurse.id);
            }
        };
        
        let cellsHtml = '';
        tr.innerHTML = `<td class="nurse-cell" style="cursor:grab; display:flex; align-items:center;" title="Arraste para trocar a escala inteira com outra funcionária">
            <span style="margin-right:8px; opacity:0.3; font-size:16px;">☰</span>
            <span>${nurse.name}</span>
        </td>`;
        for (let d=1; d<=days; d++) {
            const code  = getShift(nurse.id, d);
            const sh    = SHIFTS[code] || SHIFTS['OFF'];
            const clickable = isCoordinator || currentUser?.id===nurse.id;
            cellsHtml += `<td class="shift-cell" 
                style="background:${sh.color};color:${sh.text}"
                title="${sh.name} — ${sh.h}h"
                ${clickable?`onclick="openDayModal('${nurse.id}', ${d})"`:''}>
                ${code === 'OFF' ? '' : code}
            </td>`;
        }
        tr.innerHTML += cellsHtml;
        tbody.appendChild(tr);
    });
    
    // Adicionar rodapé com o Total de Horas do Dia
    const tfoot = document.createElement('tr');
    tfoot.innerHTML = `<td class="nurse-cell" style="font-weight:bold; background:var(--surface-color); color:var(--text); text-align:right; border-top:2px solid var(--border);">Totale O/Giorno</td>`;
    for (let d=1; d<=days; d++) {
        let dailyH = 0;
        displayNurses.forEach(nurse => {
            const sh = SHIFTS[getShift(nurse.id, d)];
            if(sh) dailyH += sh.h;
        });
        tfoot.innerHTML += `<td style="font-weight:bold; font-size:11px; color:var(--text); text-align:center; border-top:2px solid var(--border);">${dailyH}h</td>`;
    }
    tbody.appendChild(tfoot);
    
    renderCalendarSummary();
}

function renderCalendarSummary() {
    const sec = document.getElementById('calSummarySection');
    if (!isCoordinator) { sec.style.display='none'; return; }
    sec.style.display = 'block';
    
    const types = ['G', 'P', 'M1', 'M2', 'N', 'FE', 'OFF', 'AT'];
    const daysInMo = daysInMonth(currentMonth);
    const m = currentMonth.getMonth();
    const y = currentMonth.getFullYear();

    let html = `
    <thead>
        <tr>
            <th rowspan="2" style="min-width: 140px; text-align: left;">Infermiera</th>
            <th colspan="${types.length}" class="th-group" style="text-align:center; border-left: 2px solid var(--border);">Turni Fissi e Variabili</th>
            <th colspan="7" class="th-group" style="text-align:center; border-left: 2px solid var(--border); background: #fdf4ff;">Giorni Lavorati nella Settimana</th>
            <th rowspan="2" style="text-align:center; border-left: 2px solid var(--border);">Totale O.</th>
            <th rowspan="2" style="text-align:center; min-width: 160px; border-left: 2px solid var(--border);">Avvisi UI</th>
        </tr>
        <tr>
    `;
    
    // Header Linha 2 - Turnos
    types.forEach((t, idx) => { html += `<th style="${idx===0?'border-left: 2px solid var(--border);':''}">${t}</th>`; });
    
    // Header Linha 2 - Dias da Semana
    const printDays = [
        {lbl: 'Lun', sysId: 1, c: ''},
        {lbl: 'Mar', sysId: 2, c: ''},
        {lbl: 'Mer', sysId: 3, c: ''},
        {lbl: 'Gio', sysId: 4, c: ''},
        {lbl: 'Ven', sysId: 5, c: ''},
        {lbl: 'Sab', sysId: 6, c: 'td-weekend'},
        {lbl: 'Dom', sysId: 0, c: 'td-weekend'}
    ];
    printDays.forEach((d, idx) => {
        html += `<th class="${d.c}" style="${idx===0?'border-left: 2px solid var(--border);':''}">${d.lbl}</th>`;
    });
    
    html += `</tr></thead><tbody>`;
    
    NURSES.forEach(n => {
        html += `<tr><td style="text-align: left;"><strong>${n.name}</strong></td>`;
        
        // 1. Turnos Count
        types.forEach((code, idx) => {
            let codesToCount = [code];
            if (code === 'M1') codesToCount.push('MF');
            if (code === 'P') codesToCount.push('PF');
            let val = shiftCount(n.id, codesToCount);
            html += `<td style="${idx===0?'border-left: 2px solid var(--border);':''}">${val}</td>`;
        });

        // 2. Dias Trabalhados Count
        let dowCounts = [0,0,0,0,0,0,0]; 
        let weekendsSet = new Set();

        for (let d = 1; d <= daysInMo; d++) {
            let s = getShift(n.id, d);
            if (s && !['OFF', 'FE', 'AT'].includes(s)) {
                let dt = new Date(y, m, d);
                let dow = dt.getDay();
                dowCounts[dow]++;
                
                // Agrupando blocos de finais de semana. Se for sábado ou domingo, amarra ao domingo daquele FDS
                if (dow === 0 || dow === 6) {
                    let sundayDate = d + (dow === 6 ? 1 : 0);
                    weekendsSet.add(sundayDate);
                }
            }
        }
        
        let weekendsWorked = weekendsSet.size;

        printDays.forEach((d, idx) => {
            let v = dowCounts[d.sysId];
            html += `<td class="${d.c}" style="${idx===0?'border-left: 2px solid var(--border);':''}">${v > 0 ? `<span class="day-val">${v}</span>` : '<span style="opacity:0.2">-</span>'}</td>`;
        });

        // 3. Total Horas
        html += `<td style="border-left: 2px solid var(--border); text-align:center;"><strong>${nurseHours(n.id).toFixed(1)}h</strong></td>`;

        // 4. Alertas Inteligentes
        let alerts = [];
        if (weekendsWorked === 0) {
            alerts.push(`<span class="alert-badge st-danger" title="Nessun fine settimana lavorato">🔴 0 Fine Set</span>`);
        } else if (weekendsWorked > 2) {
            alerts.push(`<span class="alert-badge st-warning" title="Ha lavorato ${weekendsWorked} fine settimana nel mese">🟠 ${weekendsWorked} Fine Set</span>`);
        } else {
            alerts.push(`<span class="alert-badge st-ok">✅ Fine Set Ok</span>`);
        }

        let maxDowAmt = 0;
        let maxDowIdx = -1;
        const daysLabelPt = ['Dom','Lun','Mar','Mer','Gio','Ven','Sab'];
        
        for (let i = 0; i < 7; i++) {
            // Verifica dias da semana > 3 vezes
            if (i !== 0 && i !== 6 && dowCounts[i] >= 4) {
                if (dowCounts[i] > maxDowAmt) {
                    maxDowAmt = dowCounts[i];
                    maxDowIdx = i;
                }
            }
        }
        if (maxDowAmt >= 4) {
            alerts.push(`<span class="alert-badge st-info" title="Accumulo in questo giorno - Verificato ${maxDowAmt} volte">🔹 ${maxDowAmt} ${daysLabelPt[maxDowIdx]}</span>`);
        }

        html += `<td style="border-left: 2px solid var(--border); padding: 8px; vertical-align: middle;">
            <div style="display:flex; flex-direction:column; gap:6px; align-items:center; justify-content:center; width:100%;">
                ${alerts.join('')}
            </div>
        </td></tr>`;
    });
    
    html += `</tbody>`;
    document.getElementById('calendarSummaryTable').innerHTML = html;
}

function populateOccNurses() {
    const sel = document.getElementById('occNurse');
    if (sel) sel.innerHTML = '<option value="">Seleziona...</option>' + NURSES.map(n => `<option value="${n.id}">${n.name}</option>`).join('');
}

function renderOccurrences() {
    const tbody = document.getElementById('occTableBody');
    if (!tbody) return;
    if (occurrences.length === 0) {
        tbody.innerHTML = '<div class="empty-state" style="padding:40px;"><div class="empty-icon">📚</div><p>Nessun record ufficiale accumulato in questo mese.</p></div>';
        return;
    }
    
    // Filtro estético para OFF
    const shiftInfoAlias = (t) => t.includes('OFF') ? { color: '#e2e8f0', text: '#64748b', name: t==='OFF_INJ'?'Assenza Ingiust.':'Riposo/Assenza Giust.' } : SHIFTS[t];
    
    tbody.innerHTML = occurrences.sort((a,b)=>b.id - a.id).map(o => {
        const nurse = NURSES.find(n => n.id === o.nurseId);
        const sDate = sanitizeDate(o.start).split('-').reverse().join('/');
        const eDate = sanitizeDate(o.end).split('-').reverse().join('/');
        const shiftInfo = shiftInfoAlias(o.type);
        const color = shiftInfo ? (shiftInfo.color === '#e2e8f0' ? '#64748b' : shiftInfo.color) : '#666';
        
        let attachBadge = o.attachment ? `<div style="font-size:10px; color:var(--text-2); background:var(--bg); padding:2px 6px; border-radius:4px; display:inline-flex; align-items:center; gap:4px; border: 1px solid var(--border);">📎 ${o.attachment}</div>` : '';
        
        return `<div class="occ-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px; display:flex; justify-content:space-between; align-items:center; box-shadow: 0 4px 12px rgba(0,0,0,0.03); margin-bottom: 2px;">
            <div style="display:flex; flex-direction:column; gap:6px; width: 100%;">
                <div style="font-weight:800; color:var(--text); font-size:15px;">${nurse?.name}</div>
                <div style="font-size:13px; color:var(--text-3); font-weight:600; display:flex; gap:6px; align-items:center;">
                    <span>📅 ${sDate === eDate ? sDate : sDate + ' até ' + eDate}</span>
                </div>
                <div style="display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-top:2px;">
                    <span class="shift-badge" style="background:${color}15; color:${color}; border: 1px solid ${color}40">${shiftInfo.name}</span>
                    ${o.desc ? `<span style="font-size:12px; color:var(--text-2);">💬 ${o.desc}</span>` : ''}
                    ${attachBadge}
                </div>
            </div>
            <button onclick="removeOccurrence(${o.id})" class="btn-cls-sm" style="color:var(--danger); font-size:22px; margin-left:14px; padding:8px;" title="Cancella record dal registro">🗑</button>
        </div>`;
    }).join('');
}

let currentAttachedFile = null;

function toggleRhForm() {
    const rf = document.getElementById('rhFormContainer');
    if(rf) {
        if(rf.classList.contains('hidden')) rf.classList.remove('hidden');
        else rf.classList.add('hidden');
    }
}

function handleDragOver(e) { e.preventDefault(); document.getElementById('uploadZone').classList.add('dragover'); }
function handleDragLeave(e) { e.preventDefault(); document.getElementById('uploadZone').classList.remove('dragover'); }
function handleDrop(e) {
    e.preventDefault();
    document.getElementById('uploadZone').classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) processFile(e.dataTransfer.files[0]);
}
function handleFileSelect(e) {
    if (e.target.files && e.target.files.length > 0) processFile(e.target.files[0]);
}
function processFile(file) {
    currentAttachedFile = file.name;
    document.getElementById('uploadZone').classList.add('hidden');
    document.getElementById('filePreviewTemplate').classList.remove('hidden');
    document.getElementById('attachedFileName').textContent = file.name;
}
function removeAttachment(e) {
    if(e) e.stopPropagation();
    resetAttachment();
}
function resetAttachment() {
    currentAttachedFile = null;
    let fileInput = document.getElementById('fileInput');
    if(fileInput) fileInput.value = '';
    const zone = document.getElementById('uploadZone');
    const preview = document.getElementById('filePreviewTemplate');
    if(zone) zone.classList.remove('hidden');
    if(preview) preview.classList.add('hidden');
}

function addOccurrence() {
    const nurseId = document.getElementById('occNurse').value;
    const type = document.getElementById('occReason').value;
    const start = document.getElementById('occStart').value;
    const end = document.getElementById('occEnd').value;
    const desc = document.getElementById('occDesc').value;

    if (!nurseId || !type || !start || !end) {
        toast('Compila le date e l\'infermiera!', 'warning');
        return;
    }
    if (start > end) {
        toast('La data di fine non può essere precedente a quella di inizio.', 'error');
        return;
    }
    if (type === 'AT' && !currentAttachedFile) {
        toast('Per i Certificati/Licenze, inserire l\'allegato comprovante.', 'warning');
        return;
    }

    // Ao invés de jogar direto em occurrences, colocamos em requests pendentes (Workflow B)
    const newReq = { 
        id: generateId(), 
        type: type, // 'FE', 'OFF', 'OFF_INJ', 'AT'
        status: 'pending',
        nurseId, 
        nurseName: NURSES.find(n => n.id === nurseId)?.name,
        startDate: start, 
        endDate: end, 
        desc,
        attachment: currentAttachedFile,
        createdAt: new Date().toISOString()
    };
    requests.push(newReq);
    
    document.getElementById('occNurse').value = '';
    document.getElementById('occStart').value = '';
    document.getElementById('occEnd').value = '';
    document.getElementById('occDesc').value = '';
    resetAttachment();
    
    toggleRhForm();
    saveData();
    renderRequests();
    updateBadge();
    toast('⏳ Richiesta INVIATA. In attesa di Approvazione nella lista!', 'info');
    appendRequestToCloud(newReq);
}

function removeOccurrence(id) {
    if(!confirm('Attenzione: L\'occorrenza verrà rimossa in modo permanente. Confermare?')) return;
    occurrences = occurrences.filter(o => o.id !== id);
    saveData();
    renderOccurrences();
    toast('Registro cancellato. Genera nuovi turni per ricalcolare i riposi!', 'info');
}

// ── GERENCIADOR DE LIMITES (MODAL) E PREPARAÇÃO ─────────────
let tempHourLimits = {}; 

function openHourLimitModal() {
    const list = document.getElementById('hourLimitList');
    list.innerHTML = '';
    
    let savedLimits = {};
    try {
        const raw = localStorage.getItem('nurseHourLimits');
        if (raw && raw !== 'undefined') savedLimits = JSON.parse(raw);
    } catch(e) {
        console.warn('Ripristino limiti personalizzati corrotti.');
    }
    
    NURSES.forEach(n => {
        const defaultVal = savedLimits[n.id] ? savedLimits[n.id] : 130;
        list.innerHTML += `
            <div style="display:flex; justify-content:space-between; align-items:center; background:var(--surface); padding:8px 12px; border-radius:var(--radius-xs); border: 1px solid var(--border);">
                <span style="font-weight:600; color:var(--text); font-size: 14px;">${n.name}</span>
                <div>
                   <span style="font-size:12px; color:var(--text-3); margin-right:6px;">Max</span>
                   <input type="number" id="hl_${n.id}" value="${defaultVal}" step="1" min="10" max="300" class="field-input" style="width: 80px; padding: 6px; text-align: center;">
                </div>
            </div>
        `;
    });
    
    document.getElementById('hourLimitModal').classList.remove('hidden');
}

function closeHourLimitModal() {
    document.getElementById('hourLimitModal').classList.add('hidden');
}

function confirmAndGenerateSchedule() {
    tempHourLimits = {};
    NURSES.forEach(n => {
        const val = parseFloat(document.getElementById(`hl_${n.id}`).value);
        tempHourLimits[n.id] = isNaN(val) ? 130 : val;
    });
    localStorage.setItem('nurseHourLimits', JSON.stringify(tempHourLimits));

    closeHourLimitModal();
    generateSchedule(tempHourLimits, 1);
}

// ── SCHEDULE ALGORITHM — SMART ROSTER (HEURÍSTICA DE PONTUAÇÃO MÚLTIPLA) ────
async function generateSchedule(hourLimits = {}, startDay = 1) {
    document.getElementById('loadingOverlay').classList.remove('hidden');
    
    // Pequeno atraso para a interface conseguir renderizar a tela de carregamento (Spinner)
    await new Promise(r => setTimeout(r, 60));
    
    const m = currentMonth.getMonth();
    const y = currentMonth.getFullYear();
    const prefix = `_${m}_${y}_`;
    
    const days = daysInMonth(currentMonth);
    const simDays = days + 7; // Lookahead (Sobreposição no próximo mês para resolver pontas cegas)
    const hasVacations = occurrences.length > 0;
    const MAX_CONSEC = hasVacations ? 4 : 3;

    // Métodos Contextuais Temporários (Trabalham apenas em um objeto genérico, rápido para simulação RAM)
    function getSh(simObj, nId, d) { return simObj[`${nId}_${m}_${y}_${d}`]; }
    function setSh(simObj, nId, d, code) { simObj[`${nId}_${m}_${y}_${d}`] = code; }

    function nurseHoursTemp(simObj, nId) {
        let h = 0;
        // Importante: Considera-se impacto contratual de Horas APENAS nos dias pertencentes a este mês (<= days)
        for (let d=1; d<=days; d++) h += (SHIFTS[getSh(simObj,nId,d)]?.h||0);
        return h;
    }

    function canWorkConsecTemp(simObj, nId, day) {
        let count = 1;
        let d = day - 1;
        while (d >= 1 && getSh(simObj, nId, d) && !['OFF','FE','AT'].includes(getSh(simObj, nId, d))) { count++; d--; }
        d = day + 1;
        // Look ahead vai até simDays no contínuo dos plantões
        while (d <= simDays && getSh(simObj, nId, d) && !['OFF','FE','AT'].includes(getSh(simObj, nId, d))) { count++; d++; }
        if (count > MAX_CONSEC) return false;

        // Limite restritivo de fadiga - Bloco grande único
        if (count >= 4) {
            let startS = day;
            while(startS > 1 && getSh(simObj, nId, startS-1) && !['OFF','FE','AT'].includes(getSh(simObj, nId, startS-1))) startS--;
            let endS = day;
            while(endS < simDays && getSh(simObj, nId, endS+1) && !['OFF','FE','AT'].includes(getSh(simObj, nId, endS+1))) endS++;
            let otherBlocks = 0;
            let tempStreak = 0;
            for (let i = 1; i <= simDays; i++) {
                if (i >= startS && i <= endS) { tempStreak = 0; continue; }
                const s = getSh(simObj, nId, i);
                if (s && !['OFF','FE','AT'].includes(s)) tempStreak++;
                else { if (tempStreak >= 4) otherBlocks++; tempStreak = 0; }
            }
            if (tempStreak >= 4) otherBlocks++;
            if (otherBlocks > 0) return false;
        }
        return true;
    }

    // Verifica transições impossíveis (hard rules - nunca relaxadas)
    function checkTransitions(simObj, nId, day, code) {
        const morningShifts = ['M1', 'M2', 'MF', 'G'];
        const afternoonShifts = ['P', 'PF'];
        const prev = day > 1 ? getSh(simObj, nId, day - 1) : null;
        const next = day < simDays ? getSh(simObj, nId, day + 1) : null;
        if (morningShifts.includes(code) && afternoonShifts.includes(prev)) return false;
        if (afternoonShifts.includes(code) && morningShifts.includes(next)) return false;
        if (afternoonShifts.includes(code) && (afternoonShifts.includes(prev) || afternoonShifts.includes(next))) return false;
        return true;
    }

    function canAssignTemp(simObj, nId, day, code) {
        if (!code || ['OFF', 'FE', 'AT'].includes(code)) return true;
        if (!checkTransitions(simObj, nId, day, code)) return false;
        
        // Impede ilhas de 1 dia isolado para descanso (EXCETO dia 1 do mês - sem contexto anterior)
        const prev = day > 1 ? getSh(simObj, nId, day - 1) : null;
        if (day > 1 && prev && ['OFF', 'FE', 'AT'].includes(prev)) {
            const prev2 = day > 2 ? getSh(simObj, nId, day - 2) : null;
            if (prev2 && !['OFF', 'FE', 'AT'].includes(prev2)) return false; 
        }
        return true;
    }

    // ── AGREGADOR DE BLOQUEIOS E AUSÊNCIAS (COMPUTADO APENAS 1 VEZ) ──
    const preAssignedShifts = []; 
    
    // 1) Occurrences (Módulo Legado)
    occurrences.forEach(occ => {
        const sDate = new Date(occ.start + 'T00:00:00');
        const eDate = new Date(occ.end + 'T00:00:00');
        for (let d = startDay; d <= simDays; d++) {
            const checkDate = new Date(y, m, d);
            checkDate.setHours(0,0,0,0); sDate.setHours(0,0,0,0); eDate.setHours(0,0,0,0);
            if (checkDate >= sDate && checkDate <= eDate) {
                preAssignedShifts.push({ nurseId: occ.nurseId, day: d, code: occ.type });
            }
        }
    });

    // 2) Requisições Aprovadas (Módulo Principal da Tabela de RH)
    requests.forEach(req => {
        if (req.status !== 'approved') return;
        if (['FE', 'OFF', 'AT', 'OFF_INJ', 'vacation', 'justified', 'unexcused'].includes(req.type)) {
            let startStr = sanitizeDate(req.startDate || req.date || '');
            let endStr   = sanitizeDate(req.endDate || startStr);
            if (!startStr && req.day) {
                const tempD = new Date(y, m, req.day);
                startStr = tempD.toISOString().split('T')[0];
                endStr = startStr;
            }
            if (startStr && endStr) {
                const sDate = new Date(startStr + 'T00:00:00');
                const eDate = new Date(endStr + 'T00:00:00');
                // Traduções de domínios visuais para códigos matriz da engrenagem de geração
                let code = ['AT','FE'].includes(req.type) ? req.type : (req.type === 'vacation' ? 'FE' : 'OFF');
                for (let d = startDay; d <= simDays; d++) {
                    const checkDate = new Date(y, m, d);
                    checkDate.setHours(0,0,0,0); sDate.setHours(0,0,0,0); eDate.setHours(0,0,0,0);
                    if (checkDate >= sDate && checkDate <= eDate) {
                        preAssignedShifts.push({ nurseId: req.nurseId, day: d, code: code });
                    }
                }
            }
        } else if (req.type === 'swap') {
            // Garante que o gerador não sobreescreva trocas manuais que já foram aprovadas neste mês!
            if (req.fromDay >= startDay && req.fromDay <= simDays) {
                preAssignedShifts.push({ nurseId: req.fromNurseId, day: req.fromDay, code: req.toShift });
            }
            if (req.toDay >= startDay && req.toDay <= simDays) {
                preAssignedShifts.push({ nurseId: req.toNurseId, day: req.toDay, code: req.fromShift });
            }
        }
    });

    // SIMULADOR 1 EPOCH DE SOLUÇÃO DA ESCALA
    function simulateOneScale() {
        let tSched = {};
        let emptyShifts = 0; 
        
        // 0. Hardcopy do Passado Congelado (Se startDay > 1)
        if (startDay > 1) {
            NURSES.forEach(n => {
                for (let d = 1; d < startDay; d++) {
                    const existingCode = getShift(n.id, d);
                    // O getShift busca o cache oficial.
                    if (existingCode && existingCode !== 'FO') {
                        setSh(tSched, n.id, d, existingCode);
                    } else {
                        setSh(tSched, n.id, d, 'OFF'); 
                    }
                }
            });
        }

        // 1. Pré-carga (Férias/Ausências e Requisições Aprovadas)
        preAssignedShifts.forEach(item => {
            setSh(tSched, item.nurseId, item.day, item.code);
        });

        const nightCount = {}; NURSES.forEach(n => nightCount[n.id]=0);
        const shiftCountTemp = (nId, type) => {
            // Conta apenas dentro do mês real, não nos dias de lookahead
            let c=0; for(let d=1;d<=days;d++) if(getSh(tSched,nId,d)===type) c++; return c;
        };

        // Fase: Noite
        for (let d = startDay; d <= simDays; d++) {
            const dow = new Date(y, m, d).getDay();
            
            // Reespelhamento de domingo/sábado
            if (dow === 0 && d > 1) { 
                NURSES.forEach(n => {
                    if (getSh(tSched, n.id, d-1) === 'N') {
                        if (!['FE', 'AT'].includes(getSh(tSched, n.id, d))) {
                            setSh(tSched, n.id, d, 'N');
                            nightCount[n.id]++;
                            if (d+1 <= simDays) setSh(tSched, n.id, d+1, 'OFF');
                            if (d+2 <= simDays) setSh(tSched, n.id, d+2, 'OFF');
                        }
                    }
                });
                continue;
            }
            
            if (NURSES.some(n => getSh(tSched, n.id, d)==='N')) continue; // Já ocupado

            let eligible = NURSES.filter(n => {
                if (n.nightQuota === 0) return false; // Enfermeiras sem cota noturna não participam
                if (getSh(tSched, n.id, d)) return false; 
                const p1 = getSh(tSched, n.id, d-1);
                const p2 = getSh(tSched, n.id, d-2);
                
                if (dow === 6 && d+1 <= simDays) {
                    const sunShift = getSh(tSched, n.id, d+1);
                    if (sunShift && ['FE', 'AT'].includes(sunShift)) return false;
                    if (!canWorkConsecTemp(tSched, n.id, d+1)) return false;
                }
                
                if (p1==='N' && p2==='N') return false; 
                if (p1==='N') return dow !== 6; 
                if (p2==='N') return false; 
                
                const individualLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                if (nurseHoursTemp(tSched, n.id) >= individualLimit) return false; 
                
                if (!canWorkConsecTemp(tSched, n.id, d)) return false;
                if (!canAssignTemp(tSched, n.id, d, 'N')) return false; 
                return true;
            });
            
            // Foca nos menos exaustos de noite
            eligible.sort((a,b) => nightCount[a.id] - nightCount[b.id]);
            
            if (eligible.length > 0) {
                let pair = eligible.find(n => getSh(tSched, n.id, d-1)==='N');
                let chosen = pair;
                
                if (!chosen) {
                    let topScore = nightCount[eligible[0].id];
                    let topNurses = eligible.filter(n => nightCount[n.id] === topScore);
                    // Rotação pseudoaleatória vitalícia para a exploração de árvore funcionar
                    chosen = topNurses[Math.floor(Math.random() * topNurses.length)];
                }

                setSh(tSched, chosen.id, d, 'N');
                nightCount[chosen.id]++;
                
                if (!pair) {
                    if (d+1<=simDays && !getSh(tSched, chosen.id, d+1)) setSh(tSched, chosen.id, d+1, 'OFF');
                } else {
                    if (d+1<=simDays) setSh(tSched, chosen.id, d+1, 'OFF');
                    if (d+2<=simDays) setSh(tSched, chosen.id, d+2, 'OFF');
                }
            } else {
                emptyShifts += 1; // "Buraco crítico reportado na Escala"
            }
        }

        // Total per weekday: N(night phase) + M1 + P + (G or M2) = 4 shifts exactly
        // G and M2 alternate to ensure even distribution across the month
        let gCount = 0, m2Count = 0;

        // Fase: Diurna e Plantões Mistos
        for (let d = startDay; d <= simDays; d++) {
            const dow = new Date(y, m, d).getDay();
            const wk = dow === 0 || dow === 6;

            if (dow === 0 && d > 1) {
                NURSES.forEach(n => {
                    const satShift = getSh(tSched, n.id, d-1);
                    if (satShift === 'MF' || satShift === 'PF') {
                        if (!['FE', 'AT'].includes(getSh(tSched, n.id, d))) {
                            setSh(tSched, n.id, d, satShift);
                        }
                    }
                });
                continue; 
            }

            let targets;
            if (wk) {
                targets = ['MF', 'PF'];  // Weekend: 2 day shifts + N = 3
            } else {
                // Weekday: always M1 + P + exactly ONE of G or M2 = 3 day shifts
                // Alternating G/M2 for even monthly distribution
                let thirdShift;
                if (gCount <= m2Count) {
                    thirdShift = 'G';
                    gCount++;
                } else {
                    thirdShift = 'M2';
                    m2Count++;
                }
                targets = ['M1', 'P', thirdShift];  // 3 day shifts + N = 4 total
            }

            // Shuffle target order
            targets.sort(() => Math.random() - 0.5); 

            for (let t of targets) {
                // ── TENTATIVA ESTRITA (todas as regras) ──
                let free = NURSES.filter(n => {
                    if (getSh(tSched, n.id, d)) return false;
                    let p1 = d>1 ? getSh(tSched, n.id, d-1) : null;
                    if (p1 === 'N') return false; 
                    if (!canWorkConsecTemp(tSched, n.id, d)) return false;
                    const individualLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                    if (nurseHoursTemp(tSched, n.id) + SHIFTS[t].h > (individualLimit + 1)) return false;
                    if (!canAssignTemp(tSched, n.id, d, t)) return false;
                    if (dow === 6 && d+1 <= simDays) {
                        const sunShift = getSh(tSched, n.id, d+1);
                        if (sunShift && ['FE', 'AT'].includes(sunShift)) return false;
                    }
                    return true;
                });

                // ── FALLBACK RELAXADO (salvação para não deixar buraco) ──
                if (free.length === 0) {
                    free = NURSES.filter(n => {
                        if (getSh(tSched, n.id, d)) return false;
                        let p1 = d>1 ? getSh(tSched, n.id, d-1) : null;
                        if (p1 === 'N') return false;
                        if (!checkTransitions(tSched, n.id, d, t)) return false;
                        // Relaxamento massivo do individual limit para evitar holes, mas ainda pontuando as menos trabalhadas.
                        const individualLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                        if (nurseHoursTemp(tSched, n.id) + SHIFTS[t].h > (individualLimit + 25)) return false;
                        return true;
                    });
                }

                if (free.length > 0) {
                    free.forEach(n => {
                        let p1 = d>1 ? getSh(tSched, n.id, d-1) : null;
                        let seqPenalty = (p1 === t) ? 200 : 0; // Penaliza brutalmente turnos repetidos exatamente iguais
                        n.tmpScore = (shiftCountTemp(n.id, t) * 8) + nurseHoursTemp(tSched, n.id) + seqPenalty + (Math.random() * 5);
                    });
                    free.sort((a,b) => a.tmpScore - b.tmpScore);
                    setSh(tSched, free[0].id, d, t);
                } else {
                    emptyShifts += 1;
                }
            }
        }

        // Tapa buracos com dias inteiramente de Folga
        NURSES.forEach(n => {
            for (let d = startDay; d <= simDays; d++) {
                if (!getSh(tSched, n.id, d)) setSh(tSched, n.id, d, 'OFF');
            }
        });

        // Computa Variância + penaliza dias incompletos + Repetições Sequenciais
        let maxH = 0, minH = 999;
        let repScore = 0;
        
        NURSES.forEach(n => {
            let h = nurseHoursTemp(tSched, n.id);
            if (h > maxH) maxH = h;
            if (h < minH) minH = h;
            
            // Vasculhando repetições sequenciais de turnos no mês
            let cons = 0;
            for (let dt=2; dt<=simDays; dt++) {
                let curr = getSh(tSched, n.id, dt);
                let prev = getSh(tSched, n.id, dt-1);
                if (curr && curr !== 'OFF' && curr !== 'FE' && curr !== 'AT') {
                    if (curr === prev) cons++;
                }
            }
            repScore += cons;
        });
        
        // Penalizar dias úteis do mês real que ficaram com < 3 turnos diurnos preenchidos
        let incompleteDays = 0;
        for (let dd = 1; dd <= days; dd++) {
            const ddow = new Date(y, m, dd).getDay();
            if (ddow === 0 || ddow === 6) continue;
            let filledCount = 0;
            NURSES.forEach(n => {
                const s = getSh(tSched, n.id, dd);
                if (s && !['OFF','FE','AT','N'].includes(s)) filledCount++;
            });
            // Weekday needs 3 day shifts (M1+P+G/M2) + 1 night = 4 total. Day shifts only here.
            if (filledCount < 3) incompleteDays++;
        }
        
        let fitness = 100000 - (emptyShifts * 25000) - (incompleteDays * 35000) - (repScore * 5000) - ((maxH - minH) * 5);
        return { scheduleModel: tSched, fitness: fitness, emptyShifts: emptyShifts };
    }

    // DISPARO E FILTRAGEM DE MODELOS (MONTAGEM DE MÚLTIPLAS COMBINAÇÕES/SCORING)
    let bestSim = null;
    let validIterCount = 0;
    
    // Tenta montar 800 opções de matriz visual até achar os padrões sem erro. 
    for (let i = 0; i < 800; i++) {
        let sim = simulateOneScale();
        
        if (!bestSim || sim.fitness > bestSim.fitness) {
            bestSim = sim;
        }
        
        // Critério de parada adiantado para economizar RAM do navegador se a escala for impecável:
        if (bestSim.emptyShifts === 0 && bestSim.fitness > 99950) {
            validIterCount++;
            if (validIterCount > 5) break; 
        }
    }

    // LIMPEZA DA TELA E INSERÇÃO DO MODELO VENCEDOR (CORTANDO O EXCESSO FUTURO)
    const prefixRegex = new RegExp(`_${m}_${y}_\\d+$`);
    for (let k in schedule) {
        if (!prefixRegex.test(k)) continue;
        const parts = k.split('_'); 
        const dStr = parts[parts.length - 1]; 
        const dInt = parseInt(dStr, 10);
        
        // Exclui apenas os turnos d >= startDay, preservando os < startDay.
        if (dInt >= startDay) {
            delete schedule[k]; 
        }
    }
    
    // Filtro Lookahead: Só copia para o Schedule local o que estiver dentro do limite restrito (d <= days)
    for (const key in bestSim.scheduleModel) {
        const parts = key.split('_'); 
        // Array => [ nurseId, m, y, d ]
        const dStr = parts[parts.length - 1]; 
        const dInt = parseInt(dStr, 10);
        
        if (dInt >= startDay && dInt <= days) {
            schedule[key] = bestSim.scheduleModel[key];
        }
    }

    // RESTAURAÇÃO DE TELA
    document.getElementById('loadingOverlay').classList.add('hidden');
    renderCalendar();
    
    // Mostra clareza e franqueza à coordenadora sobre resultados complexos:
    if (bestSim.emptyShifts > 0) {
        toast(`⚠️ Turni completati, PERÒ il limite critico ha generato ${bestSim.emptyShifts} assenze o turni vuoti irrimediabili questo mese!`, 'error', 8000);
    } else {
        toast('Turni Smart generati! Il problema matriciale ha avuto il 100% di completamento. ✨', 'success', 5000);
    }
    saveData();
}

function assign(nurseId, day, code) { schedule[key(nurseId,day)] = code; }
function assignIfFree(nurseId, day, code) { if (!schedule[key(nurseId,day)]) assign(nurseId,day,code); }

function nurseHours(nurseId) {
    const days = daysInMonth(currentMonth);
    let h = 0;
    for (let d=1; d<=days; d++) h += (SHIFTS[getShift(nurseId,d)]?.h||0);
    return h;
}

function clearSchedule() {
    if (!confirm('Sei sicuro di voler cancellare tutti i turni costruiti questo mese? Questo rimuoverà tutti i turni attuali del mese dallo schermo.')) return;
    const m = currentMonth.getMonth();
    const y = currentMonth.getFullYear();
    const prefixRegex = new RegExp(`_${m}_${y}_\\d+$`);
    for (let k in schedule) {
        if (prefixRegex.test(k)) {
            delete schedule[k];
        }
    }
    renderCalendar();
    saveData();
    toast('I turni di questo mese sono stati svuotati. Gli altri mesi rimangono intatti.', 'info');
}

function swapNurses(nurseA_id, nurseB_id) {
    if (nurseA_id === nurseB_id) return;
    
    let k = `${currentMonth.getMonth()}_${currentMonth.getFullYear()}`;
    if (!monthlyOrder[k]) monthlyOrder[k] = NURSES.map(n => n.id);
    
    let arr = monthlyOrder[k];
    let idxA = arr.indexOf(nurseA_id);
    let idxB = arr.indexOf(nurseB_id);
    if (idxA === -1 || idxB === -1) return;
    
    // Troca 1: Posição de Renderização
    arr[idxA] = nurseB_id;
    arr[idxB] = nurseA_id;
    
    // Troca 2: A memória da Escala! (Para que as 'escalas da frente' não mudem visualmente)
    const days = daysInMonth(currentMonth);
    for (let d=1; d<=days; d++) {
        let keyA = key(nurseA_id, d);
        let keyB = key(nurseB_id, d);
        let shiftA = schedule[keyA];
        let shiftB = schedule[keyB];
        
        // Férias e Atestados não mudam de pessoa, o resto é transferido linearmente!
        let fixedA = ['FE', 'AT'].includes(shiftA);
        let fixedB = ['FE', 'AT'].includes(shiftB);
        if (fixedA || fixedB) continue; 
        
        if (shiftB) schedule[keyA] = shiftB; else delete schedule[keyA];
        if (shiftA) schedule[keyB] = shiftA; else delete schedule[keyB];
    }
    
    renderCalendar();
    saveData();
}

// ── EXPORT PDF ────────────────────────────────────────────────
function exportSchedule() {
    if (typeof html2pdf === 'undefined') {
        toast('Libreria PDF in caricamento. Riprova tra 1 secondo.', 'warning');
        return;
    }
    const btn = document.getElementById('exportBtn');
    btn.innerHTML = '<span class="spinner" style="width:14px;height:14px;margin:0;border-width:2px;display:inline-block;"></span>...';
    btn.style.pointerEvents = 'none';

    const loading = document.getElementById('loadingOverlay');
    if (loading) {
        const txt = loading.querySelector('.loading-txt');
        const sub = loading.querySelector('.loading-sub');
        if (txt) txt.textContent = "Generazione PDF...";
        if (sub) sub.textContent = "Attendere prego";
        loading.classList.remove('hidden');
    }

    const originalTab = document.getElementById('calendarTab');
    const clone = originalTab.cloneNode(true);
    clone.id = 'calendarTab-pdf';

    // Remove botões/toolbar do clone
    const tbar = clone.querySelector('.cal-toolbar');
    if (tbar) tbar.style.display = 'none';

    // Desfaz overflow e sticky para captura completa
    const cScroll = clone.querySelector('.cal-scroll');
    if (cScroll) { cScroll.style.overflow = 'visible'; cScroll.style.height = 'auto'; cScroll.style.maxHeight = 'none'; }
    clone.querySelectorAll('.cal-table th, .nurse-cell').forEach(el => {
        el.style.position = 'static';
        el.style.transform = 'none';
    });

    // Wrapper off-screen: evita que overflow:hidden do body corte o conteúdo
    // Funciona em Chrome, Firefox e Safari (Mac + Windows)
    const wrapper = document.createElement('div');
    wrapper.style.cssText = [
        'position:fixed',
        'top:0',
        'left:-99999px',
        'width:1400px',
        'background:#ffffff',
        'overflow:visible',
        'z-index:0',
        'padding:20px'
    ].join(';');
    wrapper.appendChild(clone);
    document.body.appendChild(wrapper);

    // Estilos PDF (light theme, legível para impressão)
    const pdfStyle = document.createElement('style');
    pdfStyle.id = 'pdf-temp-style';
    pdfStyle.textContent = `
        #calendarTab-pdf { color:#0f172a !important; background:#ffffff !important; }
        #calendarTab-pdf .shift-legend { background:transparent !important; border:none !important; margin-bottom:16px; }
        #calendarTab-pdf .legend-dot { border:1px solid rgba(0,0,0,.15) !important; }
        #calendarTab-pdf .legend-item { color:#334155 !important; }
        #calendarTab-pdf .cal-table { background:#ffffff !important; border:1px solid #cbd5e1 !important; width:100% !important; border-collapse:collapse !important; }
        #calendarTab-pdf .cal-table th { background:#f8fafc !important; color:#334155 !important; font-weight:700 !important; border:1px solid #cbd5e1 !important; }
        #calendarTab-pdf .cal-table th.wkend { color:#b45309 !important; background:#fffbeb !important; }
        #calendarTab-pdf .cal-table td { border:1px solid #cbd5e1 !important; }
        #calendarTab-pdf .nurse-cell { background:#f1f5f9 !important; color:#0f172a !important; font-weight:700 !important; white-space:nowrap; }
        #calendarTab-pdf .cal-summary-section { display:block !important; }
        #calendarTab-pdf .cal-summary-title { color:#0f172a !important; }
        #calendarTab-pdf .rpt-table th { background:#f8fafc !important; color:#334155 !important; }
        #calendarTab-pdf .rpt-table td { color:#1e293b !important; border-bottom:1px solid #cbd5e1 !important; }
        #calendarTab-pdf * { text-shadow:none !important; box-shadow:none !important; }
    `;
    document.head.appendChild(pdfStyle);

    const months = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
    const monthName = months[currentMonth.getMonth()];
    const year = currentMonth.getFullYear();

    const opt = {
        margin:      [8, 6, 8, 6],
        filename:    `Escala_Cottolengo_${monthName}_${year}.pdf`,
        image:       { type: 'jpeg', quality: 0.95 },
        html2canvas: {
            scale: 1.8,
            useCORS: true,
            allowTaint: true,
            logging: false,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: 0,
            windowWidth: 1440
        },
        jsPDF: { unit: 'mm', format: 'a3', orientation: 'landscape' }
    };

    // 800ms para o browser renderizar completamente o clone antes de capturar
    setTimeout(() => {
        html2pdf().set(opt).from(clone).save()
            .then(() => {
                _cleanupPdf(wrapper, btn, loading);
                toast('PDF scaricato con successo!', 'success');
            })
            .catch(err => {
                _cleanupPdf(wrapper, btn, loading);
                toast('Errore durante l\'esportazione PDF.', 'error');
                console.error('[PDF Export Error]', err);
            });
    }, 800);
}

function _cleanupPdf(wrapper, btn, loading) {
    if (wrapper && wrapper.parentNode) wrapper.remove();
    const tmpStyle = document.getElementById('pdf-temp-style');
    if (tmpStyle) tmpStyle.remove();
    if (btn) { btn.innerHTML = '⬇ Esporta'; btn.style.pointerEvents = ''; }
    if (loading) {
        loading.classList.add('hidden');
        const txt = loading.querySelector('.loading-txt');
        const sub = loading.querySelector('.loading-sub');
        if (txt) txt.textContent = "Generazione turni intelligenti...";
        if (sub) sub.textContent = "Applicazione regole aziendali";
    }
}

// ── DAY MODAL (UNIFIED) ───────────────────────────────────────
function openDayModal(nurseId, day) {
    selectedCell = { nurseId, day };
    const nurse  = NURSES.find(n=>n.id===nurseId);
    const code   = getShift(nurseId, day);
    const sh     = SHIFTS[code] || SHIFTS['OFF'];
    const date   = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), day);

    document.getElementById('modalSubtitle').innerHTML = 
        `<strong>${nurse.name}</strong> · ${date.toLocaleDateString('pt-BR',{weekday:'short',day:'2-digit',month:'short'})}`;
    document.getElementById('modalCurrentShift').innerHTML = `<span class="shift-badge" style="background:${sh.color}20; color:${sh.color}">${sh.name} (${sh.h}h)</span>`;
    
    // Build shift picker grid
    const picker = document.getElementById('dayShiftPicker');
    const codes = ['M1','M2','MF','G','P','PF','N','OFF','FE','AT'];
    let html = `<div style="display:grid; grid-template-columns: repeat(5, 1fr); gap: 8px;">`;
    html += codes.map(c => {
        const s = SHIFTS[c];
        const isActive = c === code;
        return `<button onclick="directChangeShift('${c}')" style="
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            gap: 4px; padding: 14px 6px; border-radius: 12px; border: 2px solid ${isActive ? '#fff' : 'transparent'};
            background: ${s.color}; color: ${s.text}; font-weight: 700; font-size: 13px;
            cursor: pointer; transition: all 200ms ease;
            box-shadow: ${isActive ? '0 0 0 3px var(--primary), 0 6px 20px rgba(0,0,0,.25)' : '0 2px 8px rgba(0,0,0,.1)'};
            transform: ${isActive ? 'scale(1.08)' : 'scale(1)'};
        " onmouseover="this.style.transform='scale(1.1)';this.style.boxShadow='0 6px 18px rgba(0,0,0,.2)'" 
           onmouseout="this.style.transform='${isActive ? 'scale(1.08)' : 'scale(1)'}';this.style.boxShadow='${isActive ? '0 0 0 3px var(--primary), 0 6px 20px rgba(0,0,0,.25)' : '0 2px 8px rgba(0,0,0,.1)'}'"
        >
            <span style="font-size: 15px; font-weight: 800;">${c}</span>
            <span style="font-size: 9px; opacity: .8; font-weight: 500;">${s.name}</span>
        </button>`;
    }).join('');
    html += `</div>`;
    picker.innerHTML = html;

    document.getElementById('dayModal').classList.remove('hidden');
}

function directChangeShift(code) {
    if(!isCoordinator || !selectedCell) return;
    
    // Strict Business Rule check: Only N, MF, PF, OFF, FE, AT on weekends
    const d = selectedCell.day;
    const dow = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), d).getDay();
    if ((dow === 0 || dow === 6) && ['M1','M2','P','G'].includes(code)) {
        toast('Attenzione: Nei fine settimana sono permessi solo turni Festivi (MF, PF) o Notte (N)!', 'error');
        return;
    }
    
    assign(selectedCell.nurseId, selectedCell.day, code);
    renderCalendar();
    saveData();
    closeDayModal();
    toast(`Turno riassegnato a ${code}`, 'success');
}

function closeDayModal(e) {
    if (!e || e.target === document.getElementById('dayModal')) {
        document.getElementById('dayModal').classList.add('hidden');
    }
}

// ── SWAP REQUEST ──────────────────────────────────────────────
function openSwapRequest() {
    closeDayModal();
    if (!selectedCell) return;
    const code = getShift(selectedCell.nurseId, selectedCell.day);
    document.getElementById('currentShiftDisplay').value = `${code} — ${SHIFTS[code].name}`;
    const sel = document.getElementById('swapNurseSelect');
    sel.innerHTML = '<option value="">— Seleziona —</option>' +
        NURSES.filter(n=>n.id!==selectedCell.nurseId)
              .map(n=>`<option value="${n.id}">${n.name}</option>`).join('');
    document.getElementById('swapDateInput').value = '';
    document.getElementById('targetShiftDisplay').value = '';
    document.getElementById('swapModal').classList.remove('hidden');
}

function onSwapNurseChange() { updateTargetShift(); }
function onSwapDateChange()  { updateTargetShift(); }
function updateTargetShift() {
    const nurseId = document.getElementById('swapNurseSelect').value;
    const dateStr = document.getElementById('swapDateInput').value;
    if (!nurseId || !dateStr) return;
    const d = new Date(dateStr+'T00:00:00');
    if (d.getMonth()!==currentMonth.getMonth()||d.getFullYear()!==currentMonth.getFullYear()) {
        document.getElementById('targetShiftDisplay').value = 'Fuori dal mese attuale'; return;
    }
    const day  = d.getDate();
    const code = getShift(nurseId, day);
    document.getElementById('targetShiftDisplay').value = `${code} — ${SHIFTS[code].name}`;
}

function closeSwapModal(e) {
    if (!e || !e.target.closest('.modal-box'))
        document.getElementById('swapModal').classList.add('hidden');
}

function submitSwapRequest() {
    const nurseId = document.getElementById('swapNurseSelect').value;
    const dateStr = document.getElementById('swapDateInput').value;
    if (!nurseId||!dateStr) { toast('Compila tutti i campi','warning'); return; }
    const d = new Date(dateStr+'T00:00:00');
    if (d.getMonth()!==currentMonth.getMonth()) { toast('Data fuori dal mese attuale','error'); return; }
    const toDay    = d.getDate();
    const toShift  = getShift(nurseId, toDay);
    const fromShift= getShift(selectedCell.nurseId, selectedCell.day);
    const toNurse  = NURSES.find(n=>n.id===nurseId);

    const st = 'pending'; // Força o registro como pendente, até para Coordenadores.
    
    // Altera o calendário imediatamente apenas se for Coordenador (autoritário)
    if (st === 'approved') {
        assign(selectedCell.nurseId, selectedCell.day, toShift);
        assign(nurseId, toDay, fromShift);
        renderCalendar();
    }

    requests.push({
        id: generateId(), type:'swap', status: st,
        fromNurseId: selectedCell.nurseId, fromNurseName: currentUser.name,
        fromDay: selectedCell.day, fromShift,
        toNurseId: nurseId, toNurseName: toNurse.name,
        toDay, toShift, createdAt: new Date().toISOString(),
        approvedAt: st === 'approved' ? new Date().toISOString() : null, 
        approvedBy: st === 'approved' ? 'Coordenadora' : null
    });
    closeSwapModal();
    updateBadge();
    saveData();
    renderRequests();
    toast(st === 'approved' ? '✅ Cambio effettuato e approvato!' : '⏳ Richiesta di cambio inviata per approvazione.', 'success');
}

// ── VACATION REQUEST ──────────────────────────────────────────
function openVacationRequest() {
    closeDayModal();
    document.getElementById('vacationStartDate').value = '';
    document.getElementById('vacationEndDate').value   = '';
    document.getElementById('vacationModal').classList.remove('hidden');
}
function closeVacationModal(e) {
    if (!e || !e.target.closest('.modal-box'))
        document.getElementById('vacationModal').classList.add('hidden');
}
function submitVacationRequest() {
    const s = document.getElementById('vacationStartDate').value;
    const e2= document.getElementById('vacationEndDate').value;
    if (!s||!e2) { toast('Compila le date','warning'); return; }
    if (new Date(e2)<new Date(s)) { toast('La data di fine deve essere successiva alla data di inizio','error'); return; }

    const st = 'pending';
    const start = new Date(s+'T00:00:00');
    const end   = new Date(e2+'T00:00:00');

    if (st === 'approved') {
        for (let dt=new Date(start); dt<=end; dt.setDate(dt.getDate()+1)) {
            if (dt.getMonth()===currentMonth.getMonth()&&dt.getFullYear()===currentMonth.getFullYear()) {
                assign(selectedCell?.nurseId||currentUser.id, dt.getDate(), 'FE');
            }
        }
        renderCalendar();
    }

    requests.push({
        id:generateId(), type:'vacation', status: st,
        nurseId: selectedCell?.nurseId||currentUser.id,
        nurseName: currentUser.name, startDate:s, endDate:e2,
        createdAt:new Date().toISOString(), 
        approvedAt: st === 'approved' ? new Date().toISOString() : null, 
        approvedBy: st === 'approved' ? 'Coordinatrice' : null
    });
    closeVacationModal();
    updateBadge();
    saveData();
    renderRequests();
    toast(st === 'approved' ? '✅ Ferie inserite nei turni!' : '⏳ Richiesta di ferie inviata.', 'success');
}

// ── JUSTIFIED ABSENCE ─────────────────────────────────────────
function openJustifiedAbsence() {
    closeDayModal();
    const date = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), selectedCell.day);
    document.getElementById('fjDateDisplay').value = date.toLocaleDateString('pt-BR');
    document.getElementById('fjReason').value = '';
    document.getElementById('justifiedModal').classList.remove('hidden');
}
function closeJustifiedModal(e) {
    if (!e || !e.target.closest('.modal-box'))
        document.getElementById('justifiedModal').classList.add('hidden');
}
function submitJustifiedAbsence() {
    const reason = document.getElementById('fjReason').value.trim();
    if (!reason) { toast('Inserisci il motivo dell\'assenza','warning'); return; }

    const st = 'pending';

    if (st === 'approved') {
        assign(selectedCell.nurseId, selectedCell.day, 'OFF');
        renderCalendar();
    }

    requests.push({
        id:generateId(), type:'justified', status: st,
        nurseId: selectedCell.nurseId, nurseName: currentUser.name,
        day: selectedCell.day, reason,
        date: new Date(currentMonth.getFullYear(), currentMonth.getMonth(), selectedCell.day).toISOString().split('T')[0],
        createdAt:new Date().toISOString(), 
        approvedAt: st === 'approved' ? new Date().toISOString() : null, 
        approvedBy: st === 'approved' ? 'Coordinatrice' : null
    });
    closeJustifiedModal();
    updateBadge();
    saveData();
    renderRequests();
    toast(st === 'approved' ? '✅ Riposo inserito nei turni!' : '⏳ Richiesta di riposo inviata.', 'success');
}

// ── MANUAL REQUESTS (Any role) ──────────────────────────
function openManualRequest() {
    const sel = document.getElementById('mrNurseSelect');
    if (isCoordinator) {
        sel.innerHTML = NURSES.map(n=>`<option value="${n.id}">${n.name}</option>`).join('');
        document.getElementById('mrUnexcusedOpt').style.display = 'block';
        document.getElementById('mrNurseGroup').style.display = 'block';
    } else {
        sel.innerHTML = `<option value="${currentUser.id}">${currentUser.name}</option>`;
        document.getElementById('mrUnexcusedOpt').style.display = 'none';
        document.getElementById('mrNurseGroup').style.display = 'none'; // Nurse asks for themselves
    }
    document.getElementById('mrType').value = 'justified';
    document.getElementById('mrDate').value = '';
    document.getElementById('mrEndDate').value = '';
    onManualRequestTypeChange();
    document.getElementById('manualRequestModal').classList.remove('hidden');
}

function onManualRequestTypeChange() {
    const t = document.getElementById('mrType').value;
    document.getElementById('mrEndDateGroup').style.display = (t==='vacation') ? 'block' : 'none';
    document.getElementById('mrDateLabel').textContent = (t==='vacation') ? 'Data Inizio' : 'Data Unica';
}

function closeManualRequestModal(e) {
    if (!e || !e.target.closest('.modal-box'))
        document.getElementById('manualRequestModal').classList.add('hidden');
}

function submitManualRequest() {
    const nurseId = document.getElementById('mrNurseSelect').value;
    const type = document.getElementById('mrType').value;
    const dateStr = document.getElementById('mrDate').value;
    const endStr = document.getElementById('mrEndDate').value;
    const nurse = NURSES.find(n=>n.id===nurseId);

    const st = 'pending';

    if (!dateStr || (type==='vacation' && !endStr)) { toast('Compila le date richieste','warning'); return; }
    
    if (type === 'vacation') {
        const start = new Date(dateStr+'T00:00:00');
        const end = new Date(endStr+'T00:00:00');
        if (end < start) { toast('La data di fine deve essere successiva alla data di inizio','error'); return; }
        
        if (st === 'approved') {
            for (let dt=new Date(start); dt<=end; dt.setDate(dt.getDate()+1)) {
                if (dt.getMonth()===currentMonth.getMonth()&&dt.getFullYear()===currentMonth.getFullYear()) {
                    assign(nurseId, dt.getDate(), 'FE');
                }
            }
        }
        requests.push({
            id:generateId(), type:'vacation', status: st,
            nurseId, nurseName: nurse.name, startDate:dateStr, endDate:endStr,
            createdAt:new Date().toISOString(), 
            approvedAt: st === 'approved' ? new Date().toISOString() : null, 
            approvedBy: st === 'approved' ? 'Coordinatrice' : null
        });
        toast(st === 'approved' ? 'Ferie registrate con successo!' : 'Richiesta di ferie inviata!', 'success');
    } else {
        const d = new Date(dateStr+'T00:00:00');
        if (st === 'approved' && d.getMonth()===currentMonth.getMonth()&&d.getFullYear()===currentMonth.getFullYear()) {
            assign(nurseId, d.getDate(), 'OFF');
        }
        
        requests.push({
            id:generateId(), type, status: st,
            nurseId, nurseName:nurse.name, 
            day: d.getDate(), date: dateStr, reason: (type==='justified'?'Richiesta di Riposo / Assenza':'Assenza Ingiustificata'),
            createdAt:new Date().toISOString(), 
            approvedAt: st === 'approved' ? new Date().toISOString() : null, 
            approvedBy: st === 'approved' ? 'Coordinatrice' : null
        });
        if (type === 'unexcused') toast('Assenza ingiustificata registrata', 'warning');
        else toast(st === 'approved' ? 'Giorno di riposo approvato!' : 'Richiesta di riposo inviata!', 'success');
    }
    
    renderCalendar();
    saveData();
    closeManualRequestModal();
    updateBadge();
    renderRequests();
}

// ── REQUEST FILTERS ───────────────────────────────────────────
let reqStatusFilter = 'all';
let reqNurseFilter = 'all';

function setReqStatusFilter(status, btn) {
    reqStatusFilter = status;
    document.querySelectorAll('.req-filter-chip').forEach(c => c.classList.remove('active'));
    if (btn) btn.classList.add('active');
    applyReqFilters();
}

function applyReqFilters() {
    reqNurseFilter = document.getElementById('reqFilterNurse')?.value || 'all';
    renderRequests();
}

function populateReqFilterNurse() {
    const sel = document.getElementById('reqFilterNurse');
    if (!sel) return;
    sel.innerHTML = '<option value="all">👤 Tutti i dipendenti</option>' +
        NURSES.map(n => `<option value="${n.id}">${n.name}</option>`).join('');
    sel.value = reqNurseFilter;
}

// ── REQUESTS RENDER ───────────────────────────────────────────
function renderRequests() {
    populateReqFilterNurse();
    const content = document.getElementById('requestsContent');
    if (!requests.length) {
        content.innerHTML = `<div class="empty-state" style="grid-column: 1/-1"><div class="empty-icon">📄</div><p>Nessuna richiesta ancora</p></div>`;
        return;
    }

    // Apply filters
    const dateFrom = document.getElementById('reqFilterFrom')?.value || '';
    const dateTo = document.getElementById('reqFilterTo')?.value || '';

    let filtered = requests.filter(req => {
        // Status filter
        if (reqStatusFilter !== 'all' && req.status !== reqStatusFilter) return false;
        // Nurse filter
        if (reqNurseFilter !== 'all') {
            const nurseId = req.nurseId || req.fromNurseId || '';
            if (nurseId !== reqNurseFilter) return false;
        }
        // Date filter
        const reqDate = req.startDate || req.date || req.createdAt?.split('T')[0] || '';
        if (dateFrom && reqDate < dateFrom) return false;
        if (dateTo && reqDate > dateTo) return false;
        return true;
    });

    if (filtered.length === 0) {
        content.innerHTML = `<div class="empty-state" style="grid-column: 1/-1"><div class="empty-icon">🔍</div><p>Nessuna richiesta corrispondente ai filtri</p></div>`;
        return;
    }

    // Ordenação: Pending primeiro
    const sorted = [...filtered].sort((a,b)=>{
        if (a.status === 'pending' && b.status !== 'pending') return -1;
        if (b.status === 'pending' && a.status !== 'pending') return 1;
        if (a.status === 'pending' && b.status === 'pending') return new Date(a.createdAt) - new Date(b.createdAt);
        return new Date(b.createdAt) - new Date(a.createdAt);
    });
    
    const typeLabels = { swap:'🔄 Cambio Turno', vacation:'🏖️ Ferie', justified:'📋 Riposo', unexcused:'⚠️ Assenza',
                        'FE': '🏖️ Ferie Programmate', 'OFF': '📋 Riposo', 'OFF_INJ': '⚠️ Assenza Ingiustificata', 'AT':'🏥 Certificato/Licenza' };
    const statusLabels = { pending:'⏳ In attesa', approved:'✅ Approvato', rejected:'❌ Rifiutato' };

    const cols = {
        pending: sorted.filter(r => r.status === 'pending'),
        approved: sorted.filter(r => r.status === 'approved'),
        rejected: sorted.filter(r => r.status === 'rejected')
    };

    const mapCard = (req) => {
        const canDelete = isCoordinator || (req.nurseId===currentUser.id||req.fromNurseId===currentUser.id);
        const canApprove = isCoordinator && req.status === 'pending';
        let waitingHtml = '';
        if (req.status === 'pending') {
            const diff = Math.floor((new Date() - new Date(req.createdAt)) / (1000*60*60*24));
            waitingHtml = `<div class="req-waiting" style="margin-bottom:8px">⏱️ Da ${diff === 0 ? 'meno di 1 giorno' : diff + ' giorn' + (diff>1?'i':'o')}</div>`;
        }
        return `<div class="req-card status-${req.status} collapsed" id="req-card-${req.id}">
            <div class="req-card-top">
                <div class="req-card-type">${typeLabels[req.type]||req.type}</div>
                <div class="req-card-status status-pill-${req.status}">${statusLabels[req.status]||req.status}</div>
            </div>
            <div class="req-card-details-wrap">
                <div class="req-card-details">
                    ${getReqDetails(req)}
                </div>
                ${waitingHtml}
                <div class="req-card-actions">
                    ${canApprove ? `<button class="req-action-btn btn-approve" onclick="approveRequest('${req.id}')">✅ Approva</button><button class="req-action-btn btn-reject" onclick="rejectRequest('${req.id}')">❌ Rifiuta</button>` : ''}
                    ${canDelete ? `<button class="req-action-btn" style="background:rgba(239,68,68,0.08); color:#f87171; border:1px solid rgba(239,68,68,0.2);" onclick="deleteRequest('${req.id}')">🗑 Elimina</button>` : ''}
                </div>
            </div>
            <button class="req-action-toggle" onclick="toggleReqCard('${req.id}')"><span>▼</span></button>
        </div>`;
    };

    let html = '';
    if (reqStatusFilter === 'all' || reqStatusFilter === 'pending') {
        html += `<div class="requests-col">
            <div class="requests-col-title"><span style="color:var(--warning)">⏳</span> In Attesa</div>
            ${cols.pending.map(mapCard).join('')}
        </div>`;
    }
    if (reqStatusFilter === 'all' || reqStatusFilter === 'approved') {
        html += `<div class="requests-col">
            <div class="requests-col-title"><span style="color:var(--success)">✅</span> Approvati</div>
            ${cols.approved.map(mapCard).join('')}
        </div>`;
    }
    if (reqStatusFilter === 'all' || reqStatusFilter === 'rejected') {
        html += `<div class="requests-col">
            <div class="requests-col-title"><span style="color:var(--danger)">❌</span> Rifiutati</div>
            ${cols.rejected.map(mapCard).join('')}
        </div>`;
    }

    content.innerHTML = html;
}

function toggleReqCard(id) {
    const card = document.getElementById(`req-card-${id}`);
    if (card) {
        card.classList.toggle('collapsed');
    }
}

// ── AUTO-SYNC (Append Nuova Richiesta al Cloud) ──
async function appendRequestToCloud(req) {
    if (!GOOGLE_API_URL) return;
    try {
        const reqRow = {
            id: String(req.id), type: req.type, status: req.status,
            nurseId: req.nurseId || req.fromNurseId || '', nurseName: req.nurseName || req.fromNurseName || '',
            startDate: req.startDate || req.date || '', endDate: req.endDate || req.startDate || req.date || '',
            desc: req.desc || req.reason || '', createdAt: req.createdAt || '', approvedAt: '', approvedBy: ''
        };
        const updUrl = `${GOOGLE_API_URL}?action=update&sheetName=Solicitacoes&apiKey=${API_KEY}`;
        // Envia as propriedades completas para usar o behavior do Google Sheet AppScript de inserir
        // a linha caso o ID seja ausente. (Polymorfismo de update/insert)
        await fetch(updUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                _keyColumn: 'id', _keyValue: String(req.id), ...reqRow
            })
        });
        console.log('[SYNC] Nuova request pub nell\'oracolo in cloud:', req.id);
    } catch (e) {
        console.warn('[SYNC] Errore publ:', e);
    }
}

function getReqDetails(req) {
    let h = `<div class="req-detail-row"><span class="req-detail-icon">👤</span><strong>${req.nurseName||req.fromNurseName}</strong></div>`;
    if (req.type==='swap') {
        const fromShiftName = SHIFTS[req.fromShift]?.name || req.fromShift || '—';
        const toShiftName = SHIFTS[req.toShift]?.name || req.toShift || '—';
        h += `<div class="req-detail-row"><span class="req-detail-icon">🔄</span><span>${fromShiftName} ➔ ${req.toNurseName || '—'} (${toShiftName})</span></div>`;
    } else if (req.type==='vacation' || req.type==='FE' || req.type==='AT' || req.type==='OFF' || req.type==='OFF_INJ') {
        const rawStart = sanitizeDate(req.startDate || req.date || '');
        const rawEnd   = sanitizeDate(req.endDate || rawStart);
        const dStrStart = rawStart ? rawStart.split('-').reverse().join('/') : '';
        const dStrEnd   = rawEnd   ? rawEnd.split('-').reverse().join('/')   : dStrStart;
        h += `<div class="req-detail-row"><span class="req-detail-icon">📅</span><span>${dStrStart === dStrEnd ? dStrStart : dStrStart + ' a ' + dStrEnd}</span></div>`;
        if(req.desc) h += `<div class="req-detail-row"><span class="req-detail-icon">💬</span><span>${req.desc}</span></div>`;
        if(req.attachment) h += `<div class="req-detail-row"><span class="req-detail-icon">📎</span><span>${req.attachment}</span></div>`;
    } else if (req.type==='justified') {
        const d = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), req.day);
        h += `<div class="req-detail-row"><span class="req-detail-icon">📅</span><span>${d.toLocaleDateString('it-IT')}</span></div><div class="req-detail-row"><span><strong>Motivo:</strong> ${req.reason}</span></div>`;
    } else if (req.type==='unexcused') {
        h += `<div class="req-detail-row"><span class="req-detail-icon">📅</span><span>${new Date(req.date).toLocaleDateString('it-IT')}</span></div>`;
    }
    if (req.approvedAt) h += `<div class="req-detail-row"><span class="req-detail-icon">✍️</span><span style="color:var(--success)">Approvato da: ${req.approvedBy}</span></div>`;
    return h;
}

// ── DELETE INDIVIDUAL DE REQUISIÇÃO NA NUVEM ─────────────────
// Substitui o antigo syncAllRequestsToCloud (que usava clearAll:true e tinha
// risco de sobrescrever solicitações criadas por mobile enquanto o desktop
// estava offline). Agora cada exclusão é enviada por id e, se falhar, é
// enfileirada em pendingDeletes (localStorage) para retry.
const PENDING_DELETES_KEY = 'pendingDeletes';

function loadPendingDeletes() {
    try { return JSON.parse(localStorage.getItem(PENDING_DELETES_KEY) || '[]'); }
    catch(e) { return []; }
}
function savePendingDeletes(arr) {
    try { localStorage.setItem(PENDING_DELETES_KEY, JSON.stringify(arr || [])); } catch(e) {}
}
function enqueuePendingDelete(id) {
    const q = loadPendingDeletes();
    if (!q.includes(String(id))) { q.push(String(id)); savePendingDeletes(q); }
}

async function deleteRequestFromCloud(id) {
    if (!GOOGLE_API_URL) return { ok: false };
    // Backend (Apps Script) espera _keyColumn + _keyValue no body (POST como text/plain)
    const url = `${GOOGLE_API_URL}?action=delete&sheetName=Solicitacoes&apiKey=${API_KEY}`;
    const body = JSON.stringify({ _keyColumn: 'id', _keyValue: String(id) });
    try {
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body
        });
        const json = await res.json().catch(()=>null);
        if (json && json.status === 'success') {
            cacheInvalidate('Solicitacoes');
            console.log(`[DELETE] Richiesta ${id} rimossa dal cloud`);
            return { ok: true };
        }
        enqueuePendingDelete(id);
        return { ok: false, err: json && json.message };
    } catch (e) {
        enqueuePendingDelete(id);
        console.warn('[DELETE] Falha de rede — enfileirado para retry:', id, e);
        return { ok: false, err: e.message };
    }
}

async function retryPendingDeletes() {
    const q = loadPendingDeletes();
    if (!q.length || !GOOGLE_API_URL) return;
    const remaining = [];
    for (const id of q) {
        try {
            const url = `${GOOGLE_API_URL}?action=delete&sheetName=Solicitacoes&apiKey=${API_KEY}`;
            const res = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ _keyColumn: 'id', _keyValue: String(id) })
            });
            const json = await res.json().catch(()=>null);
            if (!json || json.status !== 'success') remaining.push(id);
        } catch (e) {
            remaining.push(id);
        }
    }
    savePendingDeletes(remaining);
    if (remaining.length < q.length) {
        cacheInvalidate('Solicitacoes');
        console.log(`[DELETE-RETRY] ${q.length - remaining.length} delete(s) processados; ${remaining.length} restantes`);
    }
}

// Mantém compatibilidade se houver chamadas residuais — agora apenas processa a fila.
async function syncAllRequestsToCloud() { return retryPendingDeletes(); }

function deleteRequest(id) {
    requests = requests.filter(r=>String(r.id)!==String(id));
    renderRequests();
    updateBadge();
    saveData();
    toast('Richiesta eliminata', 'info');
    // Remove APENAS essa requisição no cloud. Não envia lista completa
    // (evita race condition com solicitações criadas no mobile enquanto offline).
    deleteRequestFromCloud(id);
}

function approveRequest(id) {
    const req = requests.find(r=>String(r.id)===String(id));
    if (!req) return;
    
    // Transferência para a engine de Bloqueios (Ocorrências reais)
    if (['FE', 'OFF', 'AT', 'OFF_INJ'].includes(req.type)) {
        const robType = req.type === 'OFF_INJ' ? 'OFF' : req.type;
        occurrences.push({
            id: generateId(),
            nurseId: req.nurseId,
            type: robType,
            visualType: req.type,
            start: sanitizeDate(req.startDate),
            end: sanitizeDate(req.endDate),
            desc: req.desc,
            attachment: req.attachment
        });
        renderOccurrences();
    } else if (req.type === 'swap') {
        assign(req.fromNurseId, req.fromDay, req.toShift);
        assign(req.toNurseId, req.toDay, req.fromShift);
    } else if (req.type === 'vacation') {
        const start = new Date(req.startDate+'T00:00:00');
        const end = new Date(req.endDate+'T00:00:00');
        for (let dt=new Date(start); dt<=end; dt.setDate(dt.getDate()+1)) {
            if (dt.getMonth()===currentMonth.getMonth()&&dt.getFullYear()===currentMonth.getFullYear()) {
                assign(req.nurseId, dt.getDate(), 'FE');
            }
        }
    } else if (req.type === 'justified') {
        assign(req.nurseId, req.day, 'OFF');
    }
    
    req.status = 'approved';
    req.approvedAt = new Date().toISOString();
    req.approvedBy = currentUser.name;
    
    renderCalendar();
    saveData();
    renderRequests();
    updateBadge();
    toast('Richiesta Approvata', 'success');
    
    // Auto-sync: publicar alteração na nuvem imediatamente
    syncRequestToCloud(req);
}

function rejectRequest(id) {
    const req = requests.find(r=>String(r.id)===String(id));
    if (!req) return;
    
    req.status = 'rejected';
    req.approvedAt = new Date().toISOString();
    req.approvedBy = currentUser.name;
    
    saveData();
    renderRequests();
    updateBadge();
    toast('Richiesta Rifiutata', 'warning');
    
    // Auto-sync: publicar rejeição na nuvem imediatamente
    syncRequestToCloud(req);
}

// ── AUTO-SYNC: Publica status de request na nuvem automaticamente ──
async function syncRequestToCloud(req) {
    if (!GOOGLE_API_URL) return;
    try {
        const reqUrl = `${GOOGLE_API_URL}?action=update&sheetName=Solicitacoes&apiKey=${API_KEY}`;
        await fetch(reqUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                _keyColumn: 'id',
                _keyValue: String(req.id),
                status: req.status,
                approvedAt: req.approvedAt || '',
                approvedBy: req.approvedBy || ''
            })
        });
        console.log('[SYNC] Request sincronizzata nel cloud:', req.id, req.status);
    } catch (e) {
        console.warn('[SYNC] Errore nella sincronizzazione:', e);
        toast('⚠️ Approvazione locale OK, ma sincronizzazione cloud fallita.', 'warning');
    }
}

function updateBadge() {
    const pending = requests.filter(r=>r.status==='pending').length;
    
    const b = document.getElementById('pendingBadge');
    if (b) {
        b.style.display = pending ? 'inline-flex' : 'none';
        b.textContent = pending;
    }
    
    const topCard = document.getElementById('topPendingCard');
    const topText = document.getElementById('topPendingText');
    if (topCard && topText) {
        if (isCoordinator) {
            topCard.classList.remove('hidden-card');
            topCard.style.display = 'flex';
        } else {
            topCard.classList.add('hidden-card');
            topCard.style.display = 'none';
        }
        topText.textContent = pending;
        if (pending === 0) {
            topCard.style.borderLeft = '3px solid var(--success)';
            topText.style.color = 'var(--success)';
        } else {
            topCard.style.borderLeft = '3px solid var(--danger)';
            topText.style.color = 'var(--danger)';
        }
    }
}

// ── REPORTS RENDER ────────────────────────────────────────────
function renderReports() {
    if (!currentUser) {
        document.getElementById('reportsTab').innerHTML =
            `<div class="empty-state" style="padding:80px"><div class="empty-icon">📊</div><p>Accedi per visualizzare il report</p></div>`;
        return;
    }
    if (isCoordinator) {
        renderCoordinatorReports();
    } else {
        renderNurseReports();
    }
}

// ── REPORT INDIVIDUALE (infermiera) ─────────────────────────────
function renderNurseReports() {
    const days = daysInMonth(currentMonth);
    let totalH=0, workDays=0, restDays=0, nightShifts=0;
    const counts = {};
    for (let d=1; d<=days; d++) {
        const code = getShift(currentUser.id, d);
        const sh   = SHIFTS[code];
        totalH    += sh.h;
        counts[code] = (counts[code]||0)+1;
        if (code==='OFF'||code==='FE') restDays++;
        else workDays++;
        if (code==='N') nightShifts++;
    }
    const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    set('totalHoursReport', totalH.toFixed(1)+'h');
    set('workDaysReport',   workDays);
    set('restDaysReport',   restDays);
    set('nightShiftsReport',nightShifts);

    // Shift distribution table
    const rows = Object.entries(counts).filter(([c])=>c!=='OFF').map(([code,cnt])=>{
        const sh = SHIFTS[code];
        return `<tr>
            <td><span class="shift-badge" style="background:${sh.color};color:${sh.text}">${code}</span> ${sh.name}</td>
            <td>${cnt}×</td>
            <td><strong>${(sh.h*cnt).toFixed(1)}h</strong></td>
        </tr>`;
    }).join('');
    const stEl = document.getElementById('shiftsTable');
    if (stEl) stEl.innerHTML =
        `<table class="rpt-table"><thead><tr><th>Turno</th><th>Qta</th><th>Ore</th></tr></thead><tbody>${rows}</tbody></table>`;

    // Night quota
    if (currentUser.nightQuota > 0) {
        const pct = Math.min((nightShifts/currentUser.nightQuota)*100, 100);
        const sec = document.getElementById('nightQuotaSection');
        const rep = document.getElementById('nightQuotaReport');
        if (sec) sec.style.display='block';
        if (rep) rep.innerHTML =
            `<p style="font-size:20px;font-weight:800;margin-bottom:12px">${nightShifts} / ${currentUser.nightQuota}</p>
            <div class="quota-bar"><div class="quota-fill" style="width:${pct}%"></div></div>`;
    }
}

// ── REPORT MANAGERIALE (coordinatore) ───────────────────────────
function renderCoordinatorReports() {
    const days = daysInMonth(currentMonth);
    const m = currentMonth.getMonth();
    const y = currentMonth.getFullYear();
    const team = getMonthlyNurses();

    // ── 1. Cards-riepilogo top-level ────────────────────────────
    let teamHours = 0, teamWorkDays = 0, teamRestDays = 0, teamNights = 0;
    const nursePerStats = {};
    team.forEach(n => {
        let h=0, w=0, r=0, night=0;
        for (let d=1; d<=days; d++) {
            const code = getShift(n.id, d);
            const sh = SHIFTS[code];
            h += sh.h;
            if (code === 'N') night++;
            if (code === 'OFF' || code === 'FE') r++;
            else w++;
        }
        nursePerStats[n.id] = { h, w, r, night, quota: n.nightQuota || 0, name: n.name, initials: n.initials };
        teamHours += h;
        teamWorkDays += w;
        teamRestDays += r;
        teamNights += night;
    });

    const statsGrid = document.getElementById('teamStatsGrid');
    if (statsGrid) {
        statsGrid.innerHTML = `
            <div class="stat-card"><div class="stat-icon">👥</div><div class="stat-val">${team.length}</div><div class="stat-lbl">Personale attivo</div></div>
            <div class="stat-card"><div class="stat-icon">⏱</div><div class="stat-val">${teamHours.toFixed(0)}h</div><div class="stat-lbl">Ore totali</div></div>
            <div class="stat-card"><div class="stat-icon">🌙</div><div class="stat-val">${teamNights}</div><div class="stat-lbl">Notti totali</div></div>
            <div class="stat-card"><div class="stat-icon">🏖</div><div class="stat-val">${teamRestDays}</div><div class="stat-lbl">Giorni riposo</div></div>
        `;
    }

    // ── 2. Tabella carico per infermiera + barra quota notti ───
    const loadTbl = document.getElementById('teamLoadTable');
    if (loadTbl) {
        let html = `<table class="rpt-table"><thead><tr>
            <th>Personale</th><th>Ore</th><th>Lavoro</th><th>Riposo</th><th>Notti / Quota</th>
        </tr></thead><tbody>`;
        team.forEach(n => {
            const st = nursePerStats[n.id];
            const pct = st.quota > 0 ? Math.min((st.night / st.quota) * 100, 100) : 0;
            const barColor = st.quota > 0 && st.night > st.quota ? 'var(--danger,#ef4444)' : 'var(--primary,#8b5cf6)';
            html += `<tr>
                <td><strong>${st.name}</strong> <span style="color:var(--text-3);font-size:12px">${st.initials}</span></td>
                <td>${st.h.toFixed(1)}h</td>
                <td>${st.w}</td>
                <td>${st.r}</td>
                <td>
                    <div style="display:flex;align-items:center;gap:8px">
                        <span style="min-width:48px">${st.night} / ${st.quota || '—'}</span>
                        ${st.quota > 0 ? `<div style="flex:1;background:rgba(0,0,0,.08);border-radius:4px;height:8px;overflow:hidden"><div style="width:${pct}%;height:100%;background:${barColor}"></div></div>` : ''}
                    </div>
                </td>
            </tr>`;
        });
        html += `</tbody></table>`;
        loadTbl.innerHTML = html;
    }

    // ── 3. Cobertura diária (heatmap M/P/N) ─────────────────────
    const covTbl = document.getElementById('dailyCoverageTable');
    if (covTbl) {
        const dayNames = ['D','L','M','M','G','V','S'];
        let head = `<tr><th>Giorno</th><th>M</th><th>P</th><th>N</th><th>Copertura</th></tr>`;
        let body = '';
        for (let d = 1; d <= days; d++) {
            let M=0, P=0, N=0;
            team.forEach(n => {
                const code = getShift(n.id, d);
                if (code === 'M') M++;
                else if (code === 'P') P++;
                else if (code === 'N') N++;
            });
            const dow = new Date(y, m, d).getDay();
            const isWk = dow === 0 || dow === 6;
            const totalShifts = M + P + N;
            // Sinal visual: M+P+N < 3 é baixo; N === 0 é crítico
            let warnBg = '';
            if (N === 0) warnBg = 'background:rgba(239,68,68,.12)';
            else if (totalShifts < 3) warnBg = 'background:rgba(234,179,8,.12)';
            body += `<tr style="${warnBg}">
                <td><strong>${d}</strong> <span style="color:var(--text-3);font-size:11px">${dayNames[dow]}${isWk?' •':''}</span></td>
                <td>${M}</td>
                <td>${P}</td>
                <td>${N === 0 ? '<span style="color:var(--danger,#ef4444)">0 ⚠</span>' : N}</td>
                <td><strong>${totalShifts}</strong></td>
            </tr>`;
        }
        covTbl.innerHTML = `<table class="rpt-table"><thead>${head}</thead><tbody>${body}</tbody></table>`;
    }

    // ── 4. Stats de solicitações (por tipo e status) ───────────
    const reqTbl = document.getElementById('requestsStatsTable');
    if (reqTbl) {
        const monthReqs = requests.filter(r => {
            const d = r.startDate || r.date || '';
            if (!d) return false;
            const parts = String(d).split('-');
            if (parts.length < 2) return false;
            return Number(parts[0]) === y && (Number(parts[1]) - 1) === m;
        });
        const byType = {};
        monthReqs.forEach(r => {
            const t = r.type || 'altro';
            if (!byType[t]) byType[t] = { pending: 0, approved: 0, rejected: 0 };
            const st = r.status || 'pending';
            byType[t][st] = (byType[t][st] || 0) + 1;
        });
        const typeLabels = { OFF:'Riposo', FE:'Ferie', AT:'Permesso', OFF_INJ:'Infortunio', change:'Cambio', unexcused:'Assenza' };
        let html = `<table class="rpt-table"><thead><tr><th>Tipo</th><th>In attesa</th><th>Approvate</th><th>Rifiutate</th></tr></thead><tbody>`;
        if (Object.keys(byType).length === 0) {
            html += `<tr><td colspan="4" style="text-align:center;color:var(--text-3);padding:16px">Nessuna richiesta nel mese</td></tr>`;
        } else {
            Object.entries(byType).forEach(([t, s]) => {
                html += `<tr>
                    <td>${typeLabels[t] || t}</td>
                    <td>${s.pending || 0}</td>
                    <td style="color:var(--success,#16a34a)">${s.approved || 0}</td>
                    <td style="color:var(--danger,#ef4444)">${s.rejected || 0}</td>
                </tr>`;
            });
        }
        html += `</tbody></table>`;
        reqTbl.innerHTML = html;
    }

    // ── 5. Alertas de compliance ────────────────────────────────
    const alerts = document.getElementById('complianceAlerts');
    if (alerts) {
        const msgs = [];
        team.forEach(n => {
            const st = nursePerStats[n.id];
            if (st.quota > 0 && st.night > st.quota) {
                msgs.push(`🌙 <strong>${st.name}</strong> ha ${st.night} notti (quota ${st.quota})`);
            }
            if (st.h > 180) {
                msgs.push(`⏱ <strong>${st.name}</strong> ha ${st.h.toFixed(0)}h — carico elevato`);
            }
        });
        for (let d = 1; d <= days; d++) {
            let N = 0;
            team.forEach(n => { if (getShift(n.id, d) === 'N') N++; });
            if (N === 0) {
                msgs.push(`⚠️ Giorno <strong>${d}</strong>: nessuna copertura notturna`);
            }
        }
        if (msgs.length === 0) {
            alerts.innerHTML = `<div style="padding:12px;background:rgba(22,163,74,.08);border-radius:8px;color:var(--success,#16a34a)">✓ Nessun alert di compliance per questo mese</div>`;
        } else {
            alerts.innerHTML = msgs.slice(0, 12).map(m =>
                `<div style="padding:10px 12px;background:rgba(234,179,8,.08);border-left:3px solid var(--warning,#eab308);border-radius:4px;margin-bottom:6px">${m}</div>`
            ).join('');
        }
    }

    // Campos legacy (cards individuais da antiga view) ficam ocultos via CSS
    // — mas se ainda existirem no DOM, populamos com agregados para evitar "NaN"
    const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    set('totalHoursReport', teamHours.toFixed(0)+'h');
    set('workDaysReport',   teamWorkDays);
    set('restDaysReport',   teamRestDays);
    set('nightShiftsReport',teamNights);
}

// ── INIT ──────────────────────────────────────────────────────
// ── INIT E GERENCIAMENTO DE FUNCIONÁRIOS ──────────────────────
function bootstrap() {
    loadData(); 
    initApp();
}
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootstrap);
} else {
    bootstrap();
}

function openManageNursesModal() {
    renderActiveNursesList();
    document.getElementById('manageNursesModal').classList.remove('hidden');
}
function closeManageNursesModal() {
    document.getElementById('manageNursesModal').classList.add('hidden');
    document.getElementById('newNurseName').value = '';
}

function renderActiveNursesList() {
    const list = document.getElementById('activeNursesList');
    const activeNurses = getMonthlyNurses();
    if (activeNurses.length === 0) {
        list.innerHTML = '<p style="color:var(--text-3); font-size:13px;">Nessun membro del personale nel team di questo mese.</p>';
        return;
    }
    list.innerHTML = activeNurses.map(n => `
        <div style="display:flex; justify-content:space-between; align-items:center; background:var(--bg); padding:10px 14px; border-radius:8px; border:1px solid var(--border);">
            <div><strong style="font-size:14px; color:var(--text);">${n.name}</strong> <span style="font-size:12px; color:var(--text-3); margin-left:6px;">${n.initials}</span></div>
            <button onclick="removeNurseGlobally('${n.id}')" class="btn-cls-sm" style="color:var(--danger); font-size:16px;" title="Rimuovi personale">🗑</button>
        </div>
    `).join('');
}

function addNewNurse() {
    const name = document.getElementById('newNurseName').value.trim();

    if (!name) { toast('Inserisci il nome del personale.', 'warning'); return; }

    // Automatic Initials Derivation
    const parts = name.split(' ').filter(p => p.length > 0);
    let init = '';
    if (parts.length >= 2) {
        init = (parts[0][0] + parts[parts.length-1][0]).toUpperCase();
    } else {
        init = name.substring(0, 2).toUpperCase();
    }
    
    // Default Quota (handled dynamically by system anyway)
    const nQuota = 5;

    const newId = 'n_' + Date.now();
    const newNurseObj = { id: newId, name: name, initials: init, nightQuota: nQuota };

    NURSES.push(newNurseObj);

    // Adiciona na monthlyOrder a partir DO MES ATUAL e TODOS OS FUTUROS gerados
    let targetVal = currentMonth.getFullYear() * 12 + currentMonth.getMonth();
    for (let mo in monthlyOrder) {
        let [mStr, yStr] = mo.split('_');
        let val = parseInt(yStr) * 12 + parseInt(mStr);
        if (val >= targetVal) {
            monthlyOrder[mo].push(newId);
        }
    }
    
    // Certifique-se de que a array do mes atual existe:
    let currK = `${currentMonth.getMonth()}_${currentMonth.getFullYear()}`;
    if (!monthlyOrder[currK]) {
        getMonthlyNurses(); // Forza a inicialização
        monthlyOrder[currK].push(newId);
    }

    saveData();
    renderActiveNursesList();
    renderCalendar();
    toast(`${name} aggiunto con successo!`, 'success');
    
    document.getElementById('newNurseName').value = '';
}

function removeNurseGlobally(nurseId) {
    const nurse = NURSES.find(n => n.id === nurseId);
    if (!confirm(`Desideri rimuovere ${nurse?.name} dai turni a partire DAL MESE ATTUALE? I mesi vecchi in cui ha lavorato rimarranno intatti, ma non farà più parte del team nelle prossime settimane.`)) return;

    let targetVal = currentMonth.getFullYear() * 12 + currentMonth.getMonth();
    
    // Remove do current month e de TODOS os meses do futuro que já foram visitados e guardados
    for (let mo in monthlyOrder) {
        let [mStr, yStr] = mo.split('_');
        let val = parseInt(yStr) * 12 + parseInt(mStr);
        if (val >= targetVal) {
            monthlyOrder[mo] = monthlyOrder[mo].filter(id => id !== nurseId);
        }
    }

    saveData();
    renderActiveNursesList();
    renderCalendar();
    toast(`${nurse?.name} non fa più parte dei turni a partire da questo mese.`, 'info');
}
