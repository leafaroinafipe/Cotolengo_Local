// ============================================================
//  NurseShift Pro — app.js v3.0
//  Regras de negócio: algoritmo de escala em 4 fases
//  Todas as solicitações são aprovadas automaticamente
// ============================================================

// ── CONFIG ────────────────────────────────────────────────────
const APP_CACHE_VERSION = 'v2.3'; // Incrementar para forçar resync do cloud
const COORD_PASS  = 'coord2026';
const NURSE_PASS  = 'enfermeira123';
const GOOGLE_API_URL = 'https://script.google.com/macros/s/AKfycbw7Hzr4C0V7cIM0pnU7ehbT3rpiwg-BTBpb7hnkgzIICYIbf8tBHXdjw82bFzTVVh2XxA/exec';
const API_KEY = 'cotolengo_2026_secure_key';

// ── UTILS: ID único ──────────────────────────────────────────
function generateId() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
    return Date.now() + '-' + Math.random().toString(36).substring(2, 9);
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
    'P':  { name: 'Pomeriggio', h: 8, color: '#8b5cf6', text: '#fff', period:'afternoon' },
    'PF': { name: 'Pomeriggio Festivo', h: 7.5, color: '#a78bfa', text: '#fff', period:'afternoon' },
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
    { id:'n4', name:'Alves Festa Melissa', initials:'AM', nightQuota:5 },
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
let reportMonthDate = new Date(); // Período do relatório (independente do calendário)
reportMonthDate.setDate(1);
let reportViewMode = 'monthly'; // 'monthly' ou 'annual'


// ── UTILS ─────────────────────────────────────────────────────
function key(nurseId, day) { return `${String(nurseId).trim()}_${currentMonth.getMonth()}_${currentMonth.getFullYear()}_${day}`; }
function daysInMonth(m) { return new Date(m.getFullYear(), m.getMonth()+1, 0).getDate(); }
function isWeekend(m, day) { const d = new Date(m.getFullYear(), m.getMonth(), day); return d.getDay()===0||d.getDay()===6; }

// Feriados fixos italianos (mês 0-indexed, dia)
function getItalianHolidays(year) {
    const fixed = [
        [0, 1],   // Capodanno
        [0, 6],   // Epifania
        [3, 25],  // Festa della Liberazione
        [4, 1],   // Festa del Lavoro
        [5, 2],   // Festa della Repubblica
        [7, 15],  // Ferragosto
        [10, 1],  // Tutti i Santi
        [11, 8],  // Immacolata Concezione
        [11, 25], // Natale
        [11, 26], // Santo Stefano
    ];
    // Pasquetta (Easter Monday) — cálculo algorítmico
    const a = year % 19, b = Math.floor(year / 100), c = year % 100;
    const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4), k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const month = Math.floor((h + l - 7 * m + 114) / 31) - 1;
    const day = ((h + l - 7 * m + 114) % 31) + 1;
    // Pasqua (Easter Sunday) e Pasquetta (Monday)
    const easter = new Date(year, month, day);
    const easterMon = new Date(year, month, day + 1);
    fixed.push([easter.getMonth(), easter.getDate()]);
    fixed.push([easterMon.getMonth(), easterMon.getDate()]);
    return fixed;
}

function isHoliday(monthDate, day) {
    const y = monthDate.getFullYear();
    const m = monthDate.getMonth();
    const holidays = getItalianHolidays(y);
    return holidays.some(([hm, hd]) => hm === m && hd === day);
}

function isFestivo(monthDate, day) {
    return isWeekend(monthDate, day) || isHoliday(monthDate, day);
}
function getShift(nurseId, day) { return schedule[key(nurseId,day)] || 'OFF'; }
function getShiftForMonth(nurseId, day, m, y) { return schedule[`${String(nurseId).trim()}_${m}_${y}_${day}`] || 'OFF'; }

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
        // Verifica se é necessário limpar cache (nova versão do app)
        const savedVersion = localStorage.getItem('escala_app_version');
        if (savedVersion !== APP_CACHE_VERSION) {
            console.warn(`[CACHE] Versione cambiata: ${savedVersion} → ${APP_CACHE_VERSION}. Pulisci schedule per forzare resync dal cloud.`);
            // Limpa APENAS o schedule para forçar resync — mantém nurses, orders, requests
            localStorage.removeItem('escala_schedule');
            localStorage.setItem('escala_app_version', APP_CACHE_VERSION);
        }

        const nr = localStorage.getItem('escala_nurses');
        const s = localStorage.getItem('escala_schedule');
        const o = localStorage.getItem('escala_occurrences');
        const m = localStorage.getItem('escala_monthlyOrder');
        const r = localStorage.getItem('escala_requests');
        if (nr) {
            NURSES = JSON.parse(nr);
            // Normaliza IDs para string (proteção contra tipos inconsistentes)
            NURSES.forEach(n => { n.id = String(n.id).trim(); });
        }
        if (s) schedule = JSON.parse(s);
        if (o) occurrences = JSON.parse(o);
        if (m) monthlyOrder = JSON.parse(m);
        if (r) requests = JSON.parse(r);
        console.log('[LOAD] Dati locali caricati:', NURSES.length, 'infermiere,', Object.keys(schedule).length, 'turni in schedule');
        console.log('[LOAD] NURSES IDs:', NURSES.map(n => `"${n.id}"`).join(', '));
    } catch(e) { console.error('Erro ao carregar dados locais', e); }
}

// ── GOOGLE SHEETS API (CLOUD DB) ──────────────────────────────
async function fetchGoogleDB(action, sheetName, dataObject = null) {
    if (!GOOGLE_API_URL) {
        console.info('Aviso: Operando apenas no Banco Local (Localstorage). URL da Nuvem não fornecida no app.js.');
        return null;
    }
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 12000); // 12s timeout
    try {
        if (action === 'read') {
            const url = `${GOOGLE_API_URL}?action=${action}&sheetName=${sheetName}&apiKey=${API_KEY}`;
            const response = await fetch(url, { signal: controller.signal });
            clearTimeout(timeoutId);
            return await response.json();
        } else if (action === 'write') {
            const url = `${GOOGLE_API_URL}?action=${action}&sheetName=${sheetName}&apiKey=${API_KEY}`;
            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(dataObject || {}),
                signal: controller.signal
            });
            clearTimeout(timeoutId);
            return await response.json();
        }
    } catch (error) {
        clearTimeout(timeoutId);
        if (error.name === 'AbortError') {
            console.warn("Timeout na comunicação com a Base Google — modo offline ativado.");
            toast('Tempo limite excedido. Modo offline ativo.', 'warning');
        } else {
            console.error("Erro na comunicação com a Base Google:", error);
            toast('Conexão instável com Banco Nuvem.', 'warning');
        }
        return null;
    }
}

// ── PUBLICAÇÃO NA NUVEM (para o App Mobile) ──────────────────
async function publishToCloud() {
    const btn = document.getElementById('publishBtn');
    const originalText = btn.innerHTML;
    btn.innerHTML = '<span class="spinner" style="width:14px;height:14px;margin:0;border-width:2px;display:inline-block;"></span> Publicando...';
    btn.style.pointerEvents = 'none';

    try {
        const m = currentMonth.getMonth();
        const y = currentMonth.getFullYear();
        const days = daysInMonth(currentMonth);
        const displayNurses = getMonthlyNurses();

        // AbortController para timeout de segurança em todas as chamadas de publicação
        const pubController = new AbortController();
        const pubTimeout = setTimeout(() => pubController.abort(), 30000); // 30s timeout global

        // 0. Setup headers para Escala (usa nomes padrão do app)
        const setupEscala = `${GOOGLE_API_URL}?action=setupHeaders&sheetName=Escala&apiKey=${API_KEY}`;
        await fetch(setupEscala, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ headers: ['nurseId','month','year','d1','d2','d3','d4','d5','d6','d7','d8','d9','d10','d11','d12','d13','d14','d15','d16','d17','d18','d19','d20','d21','d22','d23','d24','d25','d26','d27','d28','d29','d30','d31'] }),
            signal: pubController.signal
        });

        // 1. Publicar Funcionários — usando os nomes de coluna do Sheet existente
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
            body: JSON.stringify({ clearAll: true, rows: nursesRows }),
            signal: pubController.signal
        });

        // 2. Publicar Escala do mês atual (limpa o mês e regrava)
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
            body: JSON.stringify({ clearFilter: [{ column: 'month', value: String(m + 1) }, { column: 'year', value: String(y) }], rows: escalaRows }),
            signal: pubController.signal
        });

        // 3. Publicar Solicitações (limpa todas e regrava)
        const reqUrl = `${GOOGLE_API_URL}?action=bulkWrite&sheetName=Solicitacoes&apiKey=${API_KEY}`;
        const reqRows = requests.map(r => ({
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
        }));
        await fetch(reqUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ clearAll: true, rows: reqRows }),
            signal: pubController.signal
        });

        clearTimeout(pubTimeout);
        toast('☁️ Turni pubblicati in Cloud! I dipendenti possono già vederli nell\'App.', 'success', 5000);
    } catch (error) {
        if (error.name === 'AbortError') {
            console.warn('[PUBLISH] Timeout na publicação — operação cancelada após 30s.');
            toast('⏱ Tempo limite excedido na publicação. Tenta novamente.', 'warning', 5000);
        } else {
            console.error('Erro ao publicar na nuvem:', error);
            toast('Errore nella pubblicazione nel cloud. Verifica la connessione.', 'error');
        }
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
function showTab(tab, btn) {
    document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
    document.getElementById(tab+'Tab').classList.add('active');
    btn.classList.add('active');
    if (tab==='reports') {
        // Sincroniza o mês do report com o calendário ao abrir a tab
        reportMonthDate = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), 1);
        renderReports();
    }
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
}

async function syncScheduleFromCloud() {
    try {
        const pTxt = document.getElementById('cloudStatusText');
        if(pTxt) pTxt.textContent = 'Sincronizzando...';

        const m = currentMonth.getMonth(); // 0 a 11 base JS
        const y = currentMonth.getFullYear();

        const dbRes = await fetchGoogleDB('read', 'Escala');
        if (dbRes && dbRes.status === 'success' && dbRes.data) {
            let loaded = 0;
            const targetMonth = String(m + 1); // 1 a 12 base Sheets Visual

            console.log(`[SYNC] Escala: ${dbRes.data.length} righe totali dal cloud`);

            // Abordagem idêntica ao Mobile: sem filtros restritivos, carrega tudo
            dbRes.data.forEach((row, idx) => {
                // Normaliza todos os campos para string (Sheets pode retornar números)
                const rowMonth = String(row.month ?? '').trim();
                const rowYear  = String(row.year ?? '').trim();
                const rowNurse = String(row.nurseId ?? '').trim();

                if (!rowNurse || rowNurse === 'undefined' || rowNurse === '') return;

                if (rowMonth === targetMonth && rowYear === String(y)) {
                    console.log(`[SYNC] Riga ${idx+2}: nurseId="${rowNurse}" month=${rowMonth} year=${rowYear}`);

                    for (let d = 1; d <= 31; d++) {
                        const val = row['d' + d];
                        const shiftCode = String(val ?? '').trim();
                        // Carrega qualquer valor não-vazio (idêntico ao Mobile)
                        if (shiftCode && shiftCode !== '' && shiftCode !== 'undefined') {
                            schedule[`${rowNurse}_${m}_${y}_${d}`] = shiftCode;
                            loaded++;
                        }
                    }
                }
            });

            console.log(`[SYNC] ${loaded} turni caricati dal cloud per ${targetMonth}/${y}`);

            // SEMPRE re-renderiza e salva após sync, mesmo que loaded=0
            // (garante que dados do cloud prevaleçam sobre dados locais stale)
            saveData();
            renderCalendar();

            if (loaded > 0) {
                toast(`☁️ Turni del mese scaricati dal cloud!`, 'info', 2000);
            }
            if(pTxt) pTxt.textContent = 'App Sincronizzato';
        } else {
            console.warn('[SYNC] Nessuna risposta valida dal cloud:', dbRes);
        }
    } catch (e) {
        console.error("[SYNC] Errore download turni:", e);
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
        const fest = isFestivo(currentMonth, d);
        const dow = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), d).getDay();
        headerHtml += `<th class="${fest?'wkend':''}">
            <div class="day-num">${d}</div>
            <div class="day-lbl">${dayNames[dow]}</div>
        </th>`;
    }
    hdRow.innerHTML = headerHtml;

    // Body
    const tbody = document.getElementById('calendarBody');
    tbody.innerHTML = '';
    
    const displayNurses = getMonthlyNurses();

    // Debug: mostra dados de schedule para cada enfermeira
    const m = currentMonth.getMonth(), y = currentMonth.getFullYear();
    displayNurses.forEach(n => {
        const shifts = [];
        for (let d = 1; d <= days; d++) {
            const k = `${String(n.id).trim()}_${m}_${y}_${d}`;
            if (schedule[k] && schedule[k] !== 'OFF') shifts.push(`d${d}=${schedule[k]}`);
        }
        console.log(`[CAL] ${n.id} (${n.name}): ${shifts.length > 0 ? shifts.join(', ') : 'NESSUN TURNO'}`);
    });

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

    // Usa getMonthlyNurses() para mostrar apenas enfermeiras ativas no mês
    const summaryNurses = getMonthlyNurses();
    summaryNurses.forEach(n => {
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

// Mostra/oculta painel de upload baseado no tipo de ocorrência selecionado
function toggleAttachmentUI() {
    const reason = document.getElementById('occReason')?.value;
    const panel = document.getElementById('attachmentPanel');
    if (!panel) return;
    // Certificato/Licenza (AT) exige anexo obrigatório
    if (reason === 'AT') {
        panel.classList.remove('hidden');
    } else {
        panel.classList.add('hidden');
        resetAttachment();
    }
}

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

// ── SCHEDULE ALGORITHM — SMART ROSTER (SOLVER HÍBRIDO: HEURÍSTICA + SCORING MULTI-EPOCH) ────
// Arquitetura: Monte Carlo multi-epoch com greedy scoring ponderado.
// Gera N simulações independentes com randomização controlada e seleciona a de maior fitness.
// Inclui lookahead para o próximo mês para evitar restrições de borda.
async function generateSchedule(hourLimits = {}, startDay = 1) {
    document.getElementById('loadingOverlay').classList.remove('hidden');

    // Pequeno atraso para a interface conseguir renderizar a tela de carregamento (Spinner)
    await new Promise(r => setTimeout(r, 60));

    const m = currentMonth.getMonth();
    const y = currentMonth.getFullYear();
    const prefix = `_${m}_${y}_`;

    const days = daysInMonth(currentMonth);
    // Lookahead estendido: simula até o final do próximo mês para resolver pontas cegas
    // e permitir que o solver considere o mês seguinte sem aplicar a escala gerada nele
    const nextMonthDays = new Date(y, m + 2, 0).getDate();
    const simDays = days + nextMonthDays; // Mês atual + próximo mês completo
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
        // G e M2 são distribuídos de forma equilibrada ao longo do mês
        // Rastreamento por enfermeira para equilíbrio individual
        let globalGCount = 0, globalM2Count = 0;
        const nurseGM2Count = {};
        NURSES.forEach(n => { nurseGM2Count[n.id] = { g: 0, m2: 0 }; });

        // Fase: Diurna e Plantões Mistos
        for (let d = startDay; d <= simDays; d++) {
            const dow = new Date(y, m, d).getDay();
            const monthRef = new Date(y, m, 1);
            const festivo = isFestivo(monthRef, d);

            // Espelhamento Domingo: copia turno de sábado para domingo
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
            if (festivo) {
                targets = ['MF', 'PF'];  // Festivo (weekend/holiday): 2 day shifts + N = 3
            } else {
                // Weekday: always M1 + P + exactly ONE of G or M2 = 3 day shifts
                // Distribuição inteligente: alterna G/M2 com randomização ponderada
                let thirdShift;
                if (globalGCount <= globalM2Count) {
                    thirdShift = 'G';
                    globalGCount++;
                } else if (globalM2Count < globalGCount) {
                    thirdShift = 'M2';
                    globalM2Count++;
                } else {
                    // Desempate aleatório para exploração
                    thirdShift = Math.random() < 0.5 ? 'G' : 'M2';
                    if (thirdShift === 'G') globalGCount++; else globalM2Count++;
                }
                targets = ['M1', 'P', thirdShift];  // 3 day shifts + N = 4 total
            }

            // Shuffle target order para evitar viés de alocação
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
                // Tolerância reduzida: no máximo +12h acima do limite individual
                if (free.length === 0) {
                    free = NURSES.filter(n => {
                        if (getSh(tSched, n.id, d)) return false;
                        let p1 = d>1 ? getSh(tSched, n.id, d-1) : null;
                        if (p1 === 'N') return false;
                        if (!checkTransitions(tSched, n.id, d, t)) return false;
                        const individualLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                        if (nurseHoursTemp(tSched, n.id) + SHIFTS[t].h > (individualLimit + 12)) return false;
                        return true;
                    });
                }

                if (free.length > 0) {
                    free.forEach(n => {
                        let p1 = d>1 ? getSh(tSched, n.id, d-1) : null;
                        let seqPenalty = (p1 === t) ? 300 : 0; // Penaliza turnos repetidos consecutivos
                        // Bonus/penalidade para equilíbrio G/M2 por enfermeira
                        let gm2Bias = 0;
                        if (t === 'G') gm2Bias = (nurseGM2Count[n.id]?.g || 0) * 12;
                        if (t === 'M2') gm2Bias = (nurseGM2Count[n.id]?.m2 || 0) * 12;
                        // Penalidade de festivo: favorecer quem trabalhou menos festivos
                        let wkBias = 0;
                        if (festivo) {
                            let wkCount = 0;
                            for (let wd = 1; wd < d; wd++) {
                                if (isFestivo(monthRef, wd) && getSh(tSched, n.id, wd) && !['OFF','FE','AT'].includes(getSh(tSched, n.id, wd))) wkCount++;
                            }
                            wkBias = wkCount * 15;
                        }
                        n.tmpScore = (shiftCountTemp(n.id, t) * 10) + nurseHoursTemp(tSched, n.id) + seqPenalty + gm2Bias + wkBias + (Math.random() * 5);
                    });
                    free.sort((a,b) => a.tmpScore - b.tmpScore);
                    const chosen = free[0];
                    setSh(tSched, chosen.id, d, t);
                    // Atualizar rastreamento G/M2
                    if (t === 'G' && nurseGM2Count[chosen.id]) nurseGM2Count[chosen.id].g++;
                    if (t === 'M2' && nurseGM2Count[chosen.id]) nurseGM2Count[chosen.id].m2++;
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

        // ── FUNÇÃO DE FITNESS COMPOSTA (MULTI-CRITÉRIO) ──
        // Avalia qualidade da escala gerada em múltiplas dimensões
        let maxH = 0, minH = 999;
        let repScore = 0;
        let weekendPenalty = 0;
        let postNightViolations = 0;
        let gm2BalancePenalty = 0;
        const monthRefFit = new Date(y, m, 1);

        NURSES.forEach(n => {
            let h = nurseHoursTemp(tSched, n.id);
            if (h > maxH) maxH = h;
            if (h < minH) minH = h;

            // 1. Repetições sequenciais de turnos iguais
            let cons = 0;
            for (let dt=2; dt<=days; dt++) {
                let curr = getSh(tSched, n.id, dt);
                let prev = getSh(tSched, n.id, dt-1);
                if (curr && curr !== 'OFF' && curr !== 'FE' && curr !== 'AT') {
                    if (curr === prev) cons++;
                }
            }
            repScore += cons;

            // 2. Equilíbrio de festivos trabalhados por enfermeira (weekends + feriados italianos)
            let wkendWorked = 0;
            for (let dd = 1; dd <= days; dd++) {
                if (isFestivo(monthRefFit, dd)) {
                    const s = getSh(tSched, n.id, dd);
                    if (s && !['OFF','FE','AT'].includes(s)) wkendWorked++;
                }
            }
            // Penaliza se uma enfermeira trabalha muitos ou nenhum festivo
            if (wkendWorked > 4) weekendPenalty += (wkendWorked - 4) * 3000;
            if (wkendWorked === 0) weekendPenalty += 2000;

            // 3. Violações de descanso pós-noturno (deve ter OFF depois de noite)
            for (let dd = 1; dd < days; dd++) {
                if (getSh(tSched, n.id, dd) === 'N') {
                    const next = getSh(tSched, n.id, dd+1);
                    if (next && !['OFF','FE','AT','N'].includes(next)) {
                        postNightViolations++;
                    }
                }
            }

            // 4. Equilíbrio G vs M2 por enfermeira
            let gCnt = 0, m2Cnt = 0;
            for (let dd = 1; dd <= days; dd++) {
                const s = getSh(tSched, n.id, dd);
                if (s === 'G') gCnt++;
                if (s === 'M2') m2Cnt++;
            }
            gm2BalancePenalty += Math.abs(gCnt - m2Cnt) * 200;
        });

        // 5. Penalizar dias úteis do mês real que ficaram com < 3 turnos diurnos preenchidos
        let incompleteDays = 0;
        for (let dd = 1; dd <= days; dd++) {
            if (isFestivo(monthRefFit, dd)) continue; // Pula festivos (weekends + feriados)
            let filledCount = 0;
            NURSES.forEach(n => {
                const s = getSh(tSched, n.id, dd);
                if (s && !['OFF','FE','AT','N'].includes(s)) filledCount++;
            });
            if (filledCount < 3) incompleteDays++;
        }

        // Fitness composta ponderada
        let fitness = 100000
            - (emptyShifts * 25000)          // Buracos críticos
            - (incompleteDays * 35000)       // Dias incompletos
            - (repScore * 5000)              // Turnos repetidos consecutivos
            - (weekendPenalty)               // Desequilíbrio de fins de semana
            - (postNightViolations * 8000)   // Violação descanso pós-noturno
            - (gm2BalancePenalty)             // Desequilíbrio G/M2
            - ((maxH - minH) * 10);          // Variância de horas (peso dobrado)

        return { scheduleModel: tSched, fitness: fitness, emptyShifts: emptyShifts };
    }

    // DISPARO E FILTRAGEM DE MODELOS (SOLVER HÍBRIDO: MONTE CARLO MULTI-EPOCH)
    // Gera múltiplas simulações independentes e seleciona a de maior fitness.
    // A randomização em cada epoch garante exploração do espaço de soluções.
    let bestSim = null;
    let validIterCount = 0;
    const MAX_EPOCHS = 1200;  // Aumentado para melhor exploração

    for (let i = 0; i < MAX_EPOCHS; i++) {
        let sim = simulateOneScale();

        if (!bestSim || sim.fitness > bestSim.fitness) {
            bestSim = sim;
            validIterCount = 0; // Reset ao encontrar solução melhor
        }

        // Critério de parada adiantado: se encontrou solução excelente e estabilizou
        if (bestSim.emptyShifts === 0 && bestSim.fitness > 99500) {
            validIterCount++;
            if (validIterCount > 10) break;  // Mais iterações de confirmação
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
    const totalDays = daysInMonth(currentMonth);
    let h = 0;
    for (let d=1; d<=totalDays; d++) h += (SHIFTS[getShift(nurseId,d)]?.h||0);
    return h;
}

async function clearSchedule() {
    if (!confirm('Sei sicuro di voler cancellare tutti i turni costruiti questo mese?\nQuesto rimuoverà i turni dal sistema locale E dal database cloud.')) return;
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

    // Sincronizar exclusão com o cloud: remove linhas do mês+ano na aba Escala
    try {
        const cloudMonth = String(m + 1); // 1-12 base Sheets
        const escalaUrl = `${GOOGLE_API_URL}?action=bulkWrite&sheetName=Escala&apiKey=${API_KEY}`;
        await fetch(escalaUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                clearFilter: [
                    { column: 'month', value: cloudMonth },
                    { column: 'year', value: String(y) }
                ],
                rows: []
            })
        });
        toast('Turni del mese cancellati localmente e dal cloud.', 'success');
        console.log(`[CLEAR] Escala mese ${cloudMonth}/${y} rimossa dal cloud.`);
    } catch (e) {
        console.error('[CLEAR] Errore nella cancellazione cloud:', e);
        toast('Turni cancellati localmente, ma errore nella sincronizzazione cloud. Ri-pubblica per allineare.', 'warning');
    }
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
        filename:    `Escala_Cotolengo_${monthName}_${year}.pdf`,
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
    
    // Strict Business Rule check: Only N, MF, PF, OFF, FE, AT on festivi (weekends + holidays)
    const d = selectedCell.day;
    if (isFestivo(currentMonth, d) && ['M1','M2','P','G'].includes(code)) {
        toast('Attenzione: Nei giorni festivi sono permessi solo turni Festivi (MF, PF) o Notte (N)!', 'error');
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
function onSwapDateChange()  {
    updateSwapNurseFilter();
    updateTargetShift();
}
function updateSwapNurseFilter() {
    const dateStr = document.getElementById('swapDateInput').value;
    const sel = document.getElementById('swapNurseSelect');
    const prevValue = sel.value;
    if (!dateStr || !selectedCell) return;
    const d = new Date(dateStr+'T00:00:00');
    if (d.getMonth()!==currentMonth.getMonth()||d.getFullYear()!==currentMonth.getFullYear()) return;
    const day = d.getDate();
    // Filtra apenas enfermeiras que têm turno naquele dia (excluindo OFF e a própria)
    const filtered = NURSES.filter(n => {
        if (n.id === selectedCell.nurseId) return false;
        const shift = getShift(n.id, day);
        return shift !== 'OFF';
    });
    if (filtered.length === 0) {
        sel.innerHTML = '<option value="">— Nessuna disponibile in questa data —</option>';
    } else {
        sel.innerHTML = '<option value="">— Seleziona —</option>' +
            filtered.map(n => {
                const shift = getShift(n.id, day);
                const shiftName = SHIFTS[shift] ? SHIFTS[shift].name : shift;
                return `<option value="${n.id}">${n.name} (${shift} — ${shiftName})</option>`;
            }).join('');
    }
    // Restaura seleção anterior se ainda disponível
    if (prevValue && filtered.some(n => n.id === prevValue)) {
        sel.value = prevValue;
    }
}
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
            desc: req.desc || req.reason || '',
            swapNurseId: req.toNurseId || req.swapNurseId || '',
            swapNurseName: req.toNurseName || req.swapNurseName || '',
            createdAt: req.createdAt || '', approvedAt: '', approvedBy: ''
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

// ── SYNC COMPLETO DE REQUISIÇÕES NA NUVEM ────────────────────
// Chamado após exclusões locais para garantir que o cloud não "ressuscite"
// requisições deletadas na próxima sincronização.
async function syncAllRequestsToCloud() {
    if (!GOOGLE_API_URL) return;
    const syncCtrl = new AbortController();
    const syncTO = setTimeout(() => syncCtrl.abort(), 15000);
    try {
        const reqRows = requests.map(r => ({
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
        }));
        const reqUrl = `${GOOGLE_API_URL}?action=bulkWrite&sheetName=Solicitacoes&apiKey=${API_KEY}`;
        await fetch(reqUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ clearAll: true, rows: reqRows }),
            signal: syncCtrl.signal
        });
        clearTimeout(syncTO);
        console.log('[SYNC] Lista completa de requests sincronizada na nuvem após exclusão.');
    } catch (e) {
        clearTimeout(syncTO);
        if (e.name === 'AbortError') {
            console.warn('[SYNC] Timeout na sincronização de exclusão.');
        } else {
            console.warn('[SYNC] Erro ao sincronizar lista de requests pós-exclusão:', e);
        }
    }
}

function deleteRequest(id) {
    requests = requests.filter(r=>String(r.id)!==String(id));
    renderRequests();
    updateBadge();
    saveData();
    toast('Richiesta eliminata', 'info');
    // Sincroniza a exclusão na nuvem para evitar que a requisição "ressuscite"
    // na próxima sincronização do cloud (Bug 2 — deleção não persistia no cloud)
    syncAllRequestsToCloud();
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

// ── REPORTS NAV ──────────────────────────────────────────────
function changeReportMonth(dir) {
    reportMonthDate = new Date(reportMonthDate.getFullYear(), reportMonthDate.getMonth() + dir, 1);
    // Carrega dados do cloud para o mês selecionado, se necessário
    loadReportMonthData().then(() => renderReports());
}

async function loadReportMonthData() {
    const m = reportMonthDate.getMonth();
    const y = reportMonthDate.getFullYear();
    // Verifica se já temos dados para este mês
    const hasData = NURSES.some(n => {
        const shift = schedule[`${String(n.id).trim()}_${m}_${y}_1`];
        return shift && shift !== 'OFF';
    });
    if (hasData) return; // Já temos dados

    // Tenta baixar do cloud
    try {
        const dbRes = await fetchGoogleDB('read', 'Escala');
        if (dbRes && dbRes.status === 'success' && dbRes.data) {
            const sheetMonthValue = String(m + 1);
            const maxDays = new Date(y, m + 1, 0).getDate();
            dbRes.data.forEach(row => {
                const rowMonth = String(row.month || '').trim();
                const rowYear  = String(row.year || '').trim();
                const rowNurse = String(row.nurseId || '').trim();
                if (rowMonth === sheetMonthValue && rowYear === String(y)) {
                    if (!NURSES.find(n => String(n.id).trim() === rowNurse)) return;
                    for (let d = 1; d <= maxDays; d++) {
                        let shiftCode = String(row['d' + d] || '').trim();
                        if (shiftCode && shiftCode !== '' && shiftCode !== 'undefined') {
                            schedule[`${rowNurse}_${m}_${y}_${d}`] = shiftCode;
                        }
                    }
                }
            });
            saveData();
        }
    } catch (e) {
        console.warn('[REPORT] Errore download turni per mese report:', e);
    }
}

// ── REPORTS RENDER — DASHBOARD COMPLETO (COORDINATRICE + INFERMIERE) ──────
function toggleReportView(mode) {
    reportViewMode = mode;
    renderReports();
}

function renderReports() {
    if (!currentUser) return;
    if (reportViewMode === 'annual') { renderAnnualReport(); return; }
    renderMonthlyReport();
}

function renderMonthlyReport() {
    const refMonth = reportMonthDate;
    const daysInMo = new Date(refMonth.getFullYear(), refMonth.getMonth() + 1, 0).getDate();
    const m = refMonth.getMonth();
    const y = refMonth.getFullYear();
    const months = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
    const monthLabel = `${months[m]} ${y}`;

    // Seleção de enfermeiras: coordenadora vê todas, enfermeira vê só a dela
    const monthKey = `${m}_${y}`;
    const orderForMonth = monthlyOrder[monthKey] || NURSES.map(n => n.id);
    const displayNurses = isCoordinator
        ? orderForMonth.map(id => NURSES.find(n => n.id === id)).filter(Boolean)
        : [NURSES.find(n => n.id === currentUser.id)].filter(Boolean);

    if (displayNurses.length === 0) {
        document.getElementById('reportsTab').innerHTML = `<div class="empty-state" style="padding:80px"><div class="empty-icon">📊</div><p>Nessun personale attivo questo mese.</p></div>`;
        return;
    }

    // ── COLETA DE DADOS AGREGADOS ──
    let teamTotalH = 0, teamWorkDays = 0, teamRestDays = 0, teamNightShifts = 0;
    let teamAbsences = 0, teamVacations = 0;
    const nurseData = [];

    const monthRef = new Date(y, m, 1);
    // Ore Dovute: dias úteis (non-festivi) × 7.5h (standard contrattuale italiano)
    let giorniFeriali = 0;
    for (let d = 1; d <= daysInMo; d++) { if (!isFestivo(monthRef, d)) giorniFeriali++; }
    const oreDovute = giorniFeriali * 7.5;

    // Copertura giornaliera (quanti turni attivi per giorno)
    const dailyCoverage = [];
    for (let d = 1; d <= daysInMo; d++) {
        const fest = isFestivo(monthRef, d);
        let count = 0;
        const dayShifts = {};
        displayNurses.forEach(n => {
            const code = getShiftForMonth(n.id, d, m, y);
            if (code && !['OFF','FE','AT'].includes(code)) {
                count++;
                dayShifts[code] = (dayShifts[code] || 0) + 1;
            }
        });
        const expected = fest ? 3 : 4;
        dailyCoverage.push({ day: d, count, expected, fest, shifts: dayShifts });
    }

    displayNurses.forEach(n => {
        let totalH = 0, workDays = 0, restDays = 0, nightCount = 0;
        let feCount = 0, atCount = 0, offCount = 0;
        const shiftCounts = {};
        let weekendsWorked = 0;
        const weekendSet = new Set();
        // Breakdown ore: diurno/notturno × feriale/festivo
        let oreDiurneFeriali = 0, oreDiurneFestive = 0;
        let oreNotturneFeriali = 0, oreNotturneFestive = 0;
        let oreFerie = 0, oreMalattia = 0;

        for (let d = 1; d <= daysInMo; d++) {
            const code = getShiftForMonth(n.id, d, m, y);
            const sh = SHIFTS[code];
            if (!sh) continue;

            totalH += sh.h;
            shiftCounts[code] = (shiftCounts[code] || 0) + 1;

            if (['OFF'].includes(code)) { offCount++; restDays++; }
            else if (code === 'FE') { feCount++; restDays++; oreFerie += 7.5; }
            else if (code === 'AT') { atCount++; restDays++; oreMalattia += 7.5; }
            else { workDays++; }
            if (code === 'N') nightCount++;

            // Breakdown ore diurne/notturne × feriali/festive
            const fest = isFestivo(monthRef, d);
            if (code === 'N') {
                if (fest) oreNotturneFestive += sh.h;
                else oreNotturneFeriali += sh.h;
            } else if (!['OFF','FE','AT'].includes(code)) {
                if (fest) oreDiurneFestive += sh.h;
                else oreDiurneFeriali += sh.h;
            }

            // Festivos trabalhados (weekends + feriados italianos)
            if (fest && !['OFF','FE','AT'].includes(code)) {
                weekendSet.add(d);
            }
        }
        weekendsWorked = weekendSet.size;

        teamTotalH += totalH;
        teamWorkDays += workDays;
        teamRestDays += restDays;
        teamNightShifts += nightCount;
        teamAbsences += atCount;
        teamVacations += feCount;

        nurseData.push({
            nurse: n, totalH, workDays, restDays, nightCount,
            feCount, atCount, offCount, shiftCounts, weekendsWorked,
            oreDiurneFeriali, oreDiurneFestive, oreNotturneFeriali, oreNotturneFestive,
            oreFerie, oreMalattia, oreDovute, differenza: totalH - oreDovute
        });
    });

    // Contadores de requests do mês
    const monthRequests = requests.filter(r => {
        const rDate = sanitizeDate(r.startDate || r.date || r.createdAt?.split('T')[0] || '');
        if (!rDate) return false;
        const rd = new Date(rDate + 'T00:00:00');
        return rd.getMonth() === m && rd.getFullYear() === y;
    });
    const pendingReqs = monthRequests.filter(r => r.status === 'pending').length;
    const approvedReqs = monthRequests.filter(r => r.status === 'approved').length;
    const rejectedReqs = monthRequests.filter(r => r.status === 'rejected').length;

    // Taxa de absenteísmo: (dias AT + FE) / (total dias possíveis) * 100
    const totalPossibleDays = displayNurses.length * daysInMo;
    const absentDays = nurseData.reduce((s, nd) => s + nd.atCount + nd.feCount, 0);
    const absenteeismRate = totalPossibleDays > 0 ? ((absentDays / totalPossibleDays) * 100).toFixed(1) : '0.0';

    // Horas médias
    const avgHours = displayNurses.length > 0 ? (teamTotalH / displayNurses.length).toFixed(1) : '0.0';
    const maxH = Math.max(...nurseData.map(nd => nd.totalH));
    const minH = Math.min(...nurseData.map(nd => nd.totalH));
    const hourSpread = (maxH - minH).toFixed(1);

    // ── RENDER HTML ──
    const tab = document.getElementById('reportsTab');
    tab.innerHTML = `
    <div class="reports-wrap" style="max-width:1200px;">
        <div class="rpt-header">
            <h2>📊 Dashboard Operativo</h2>
            <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap;">
                <div class="rpt-view-toggle">
                    <button class="rpt-toggle-btn ${reportViewMode==='monthly'?'active':''}" onclick="toggleReportView('monthly')">Mensile</button>
                    <button class="rpt-toggle-btn ${reportViewMode==='annual'?'active':''}" onclick="toggleReportView('annual')">Annuale</button>
                </div>
                <div class="rpt-month-nav">
                    <button class="rpt-nav-btn" onclick="changeReportMonth(-1)" title="Mese precedente">◀</button>
                    <span class="rpt-month">${monthLabel}</span>
                    <button class="rpt-nav-btn" onclick="changeReportMonth(1)" title="Mese successivo">▶</button>
                </div>
            </div>
        </div>

        <!-- KPI Cards -->
        <div class="stats-grid" style="grid-template-columns: repeat(auto-fit, minmax(160px,1fr));">
            <div class="stat-card stat-primary">
                <div class="stat-lbl">Ore Totali Team</div>
                <div class="stat-val" style="font-size:28px;">${teamTotalH.toFixed(1)}h</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">Media: ${avgHours}h/persona</div>
            </div>
            <div class="stat-card stat-success">
                <div class="stat-lbl">Giorni Lavorati</div>
                <div class="stat-val" style="font-size:28px;">${teamWorkDays}</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">Riposo: ${teamRestDays}</div>
            </div>
            <div class="stat-card stat-night">
                <div class="stat-lbl">Turni Notturni</div>
                <div class="stat-val" style="font-size:28px;">${teamNightShifts}</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">${daysInMo} notti necessarie</div>
            </div>
            <div class="stat-card stat-warning">
                <div class="stat-lbl">Tasso Assenteismo</div>
                <div class="stat-val" style="font-size:28px;">${absenteeismRate}%</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">${absentDays} giorni assenti</div>
            </div>
        </div>

        <!-- Indicadores de RH -->
        <div class="report-section">
            <h3>📋 Indicatori Risorse Umane</h3>
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px,1fr)); gap:16px; margin-bottom:16px;">
                <div style="background:rgba(245,158,11,0.08); border:1px solid rgba(245,158,11,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--warning);">${pendingReqs}</div>
                    <div style="font-size:12px; color:var(--text-3); margin-top:4px;">Richieste In Attesa</div>
                </div>
                <div style="background:rgba(16,185,129,0.08); border:1px solid rgba(16,185,129,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--success);">${approvedReqs}</div>
                    <div style="font-size:12px; color:var(--text-3); margin-top:4px;">Richieste Approvate</div>
                </div>
                <div style="background:rgba(239,68,68,0.08); border:1px solid rgba(239,68,68,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--danger);">${rejectedReqs}</div>
                    <div style="font-size:12px; color:var(--text-3); margin-top:4px;">Richieste Rifiutate</div>
                </div>
                <div style="background:rgba(139,92,246,0.08); border:1px solid rgba(139,92,246,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--primary-light);">${hourSpread}h</div>
                    <div style="font-size:12px; color:var(--text-3); margin-top:4px;">Dispersione Ore (Max-Min)</div>
                </div>
            </div>
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px,1fr)); gap:16px;">
                <div style="background:rgba(16,185,129,0.05); border:1px solid rgba(16,185,129,0.15); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--success);">${teamVacations}</div>
                    <div style="font-size:12px; color:var(--text-3); margin-top:4px;">Giorni Ferie (FE)</div>
                </div>
                <div style="background:rgba(239,68,68,0.05); border:1px solid rgba(239,68,68,0.15); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--danger);">${teamAbsences}</div>
                    <div style="font-size:12px; color:var(--text-3); margin-top:4px;">Certificati/Licenze (AT)</div>
                </div>
            </div>
        </div>

        <!-- Ranking de Horas por Enfermeira -->
        <div class="report-section">
            <h3>⏱ Ranking Ore per Infermiera</h3>
            <table class="rpt-table">
                <thead><tr>
                    <th style="text-align:left">Infermiera</th>
                    <th>Ore Tot.</th>
                    <th>Lavorati</th>
                    <th>Riposo</th>
                    <th>Notti</th>
                    <th>Ferie</th>
                    <th>Certif.</th>
                    <th>Festivi</th>
                    <th>Carico</th>
                </tr></thead>
                <tbody>
                ${nurseData.sort((a,b) => b.totalH - a.totalH).map((nd, idx) => {
                    const pct = maxH > 0 ? ((nd.totalH / maxH) * 100).toFixed(0) : 0;
                    const barColor = nd.totalH >= maxH ? 'var(--danger)' : nd.totalH <= minH ? 'var(--success)' : 'var(--primary)';
                    return `<tr>
                        <td style="text-align:left"><strong>${idx+1}.</strong> ${nd.nurse.name}</td>
                        <td><strong>${nd.totalH.toFixed(1)}h</strong></td>
                        <td>${nd.workDays}</td>
                        <td>${nd.restDays}</td>
                        <td>${nd.nightCount}</td>
                        <td>${nd.feCount}</td>
                        <td>${nd.atCount}</td>
                        <td>${nd.weekendsWorked}</td>
                        <td style="min-width:120px;">
                            <div style="background:rgba(255,255,255,0.06); border-radius:99px; height:8px; overflow:hidden;">
                                <div style="width:${pct}%; height:100%; background:${barColor}; border-radius:99px; transition:width 0.6s;"></div>
                            </div>
                        </td>
                    </tr>`;
                }).join('')}
                </tbody>
            </table>
        </div>

        <!-- Distribuição de Turnos por Enfermeira -->
        <div class="report-section">
            <h3>🔄 Distribuzione Turni per Infermiera</h3>
            <div style="overflow-x:auto;">
            <table class="rpt-table">
                <thead><tr>
                    <th style="text-align:left">Infermiera</th>
                    ${['M1','M2','MF','G','P','PF','N','OFF','FE','AT'].map(c =>
                        `<th><span style="display:inline-block; width:10px; height:10px; border-radius:3px; background:${SHIFTS[c].color}; margin-right:4px; vertical-align:middle;"></span>${c}</th>`
                    ).join('')}
                </tr></thead>
                <tbody>
                ${nurseData.map(nd => `<tr>
                    <td style="text-align:left">${nd.nurse.name}</td>
                    ${['M1','M2','MF','G','P','PF','N','OFF','FE','AT'].map(c => {
                        const val = nd.shiftCounts[c] || 0;
                        return `<td style="font-weight:${val > 0 ? '700' : '400'}; opacity:${val > 0 ? '1' : '0.3'};">${val}</td>`;
                    }).join('')}
                </tr>`).join('')}
                </tbody>
            </table>
            </div>
        </div>

        <!-- Bilancio Ore: Dovute vs Effettive -->
        <div class="report-section">
            <h3>⚖️ Bilancio Ore: Dovute vs Effettive</h3>
            <p style="font-size:12px; color:var(--text-3); margin-bottom:12px;">Ore dovute calcolate su ${giorniFeriali} giorni feriali × 7.5h = ${oreDovute.toFixed(1)}h standard</p>
            <div style="overflow-x:auto;">
            <table class="rpt-table">
                <thead><tr>
                    <th style="text-align:left">Infermiera</th>
                    <th>Ore Dovute</th>
                    <th>Ore Effettive</th>
                    <th>Differenza</th>
                    <th>Ore Ferie</th>
                    <th>Ore Malattia</th>
                </tr></thead>
                <tbody>
                ${nurseData.map(nd => {
                    const diff = nd.differenza;
                    const diffColor = diff > 0 ? 'var(--success)' : diff < 0 ? 'var(--danger)' : 'var(--text-2)';
                    const diffSign = diff > 0 ? '+' : '';
                    return `<tr>
                        <td style="text-align:left">${nd.nurse.name}</td>
                        <td>${nd.oreDovute.toFixed(1)}h</td>
                        <td><strong>${nd.totalH.toFixed(1)}h</strong></td>
                        <td style="color:${diffColor}; font-weight:700;">${diffSign}${diff.toFixed(1)}h</td>
                        <td>${nd.oreFerie.toFixed(1)}h</td>
                        <td>${nd.oreMalattia.toFixed(1)}h</td>
                    </tr>`;
                }).join('')}
                </tbody>
            </table>
            </div>
        </div>

        <!-- Breakdown Ore: Diurno/Notturno × Feriale/Festivo -->
        <div class="report-section">
            <h3>📊 Ripartizione Ore Diurne/Notturne</h3>
            <div style="overflow-x:auto;">
            <table class="rpt-table">
                <thead><tr>
                    <th style="text-align:left">Infermiera</th>
                    <th>Diurne Feriali</th>
                    <th>Diurne Festive</th>
                    <th>Notturne Feriali</th>
                    <th>Notturne Festive</th>
                    <th>Totale</th>
                </tr></thead>
                <tbody>
                ${nurseData.map(nd => {
                    const tot = nd.oreDiurneFeriali + nd.oreDiurneFestive + nd.oreNotturneFeriali + nd.oreNotturneFestive;
                    return `<tr>
                        <td style="text-align:left">${nd.nurse.name}</td>
                        <td>${nd.oreDiurneFeriali.toFixed(1)}h</td>
                        <td>${nd.oreDiurneFestive.toFixed(1)}h</td>
                        <td>${nd.oreNotturneFeriali.toFixed(1)}h</td>
                        <td>${nd.oreNotturneFestive.toFixed(1)}h</td>
                        <td><strong>${tot.toFixed(1)}h</strong></td>
                    </tr>`;
                }).join('')}
                </tbody>
            </table>
            </div>
        </div>

        <!-- Copertura Giornaliera -->
        <div class="report-section">
            <h3>📅 Copertura Giornaliera</h3>
            <p style="font-size:12px; color:var(--text-3); margin-bottom:12px;">Turni attivi per giorno. Feriali: 4 attesi | Festivi: 3 attesi</p>
            <div style="display:flex; flex-wrap:wrap; gap:4px;">
            ${dailyCoverage.map(dc => {
                const ok = dc.count >= dc.expected;
                const bg = ok ? 'rgba(16,185,129,0.15)' : 'rgba(239,68,68,0.2)';
                const border = ok ? 'rgba(16,185,129,0.4)' : 'rgba(239,68,68,0.5)';
                const textColor = ok ? 'var(--success)' : 'var(--danger)';
                const dayNames = ['D','L','M','M','G','V','S'];
                const dow = new Date(y, m, dc.day).getDay();
                return `<div style="
                    display:flex; flex-direction:column; align-items:center; justify-content:center;
                    width:36px; height:48px; border-radius:8px;
                    background:${bg}; border:1px solid ${border};
                    font-size:10px; font-weight:600;
                " title="Giorno ${dc.day} (${dc.fest?'Festivo':'Feriale'}): ${dc.count}/${dc.expected} turni${dc.count < dc.expected ? ' ⚠️ SCOPERTO' : ''}&#10;${Object.entries(dc.shifts).map(([k,v]) => k+':'+v).join(', ')}">
                    <span style="font-size:9px; opacity:0.6;">${dayNames[dow]}</span>
                    <span style="font-size:14px; font-weight:800; color:${textColor};">${dc.count}</span>
                    <span style="font-size:8px; opacity:0.5;">${dc.day}</span>
                </div>`;
            }).join('')}
            </div>
        </div>

        <!-- Quota Noturna -->
        <div class="report-section">
            <h3>🌙 Quota Turni Notturni</h3>
            <div style="display:grid; gap:12px;">
            ${nurseData.map(nd => {
                const quota = nd.nurse.nightQuota || 5;
                const pct = Math.min((nd.nightCount / quota) * 100, 100).toFixed(0);
                const color = nd.nightCount > quota ? 'var(--danger)' : nd.nightCount === quota ? 'var(--warning)' : 'var(--primary)';
                return `<div style="display:flex; align-items:center; gap:12px;">
                    <span style="min-width:140px; font-size:13px; font-weight:600; color:var(--text-2);">${nd.nurse.name}</span>
                    <div style="flex:1; background:rgba(255,255,255,0.06); border-radius:99px; height:10px; overflow:hidden;">
                        <div style="width:${pct}%; height:100%; background:${color}; border-radius:99px; transition:width 0.6s;"></div>
                    </div>
                    <span style="min-width:60px; text-align:right; font-size:13px; font-weight:700; color:${color};">${nd.nightCount}/${quota}</span>
                </div>`;
            }).join('')}
            </div>
        </div>
    </div>`;
}

// ── ANNUAL REPORT ────────────────────────────────────────────
async function loadAnnualData(year) {
    // Tenta carregar dados de todos os 12 meses do ano a partir do cloud
    try {
        const dbRes = await fetchGoogleDB('read', 'Escala');
        if (dbRes && dbRes.status === 'success' && dbRes.data) {
            let loaded = 0;
            dbRes.data.forEach(row => {
                const rowYear = String(row.year ?? '').trim();
                const rowNurse = String(row.nurseId ?? '').trim();
                const rowMonth = parseInt(String(row.month ?? '0').trim());
                if (!rowNurse || rowYear !== String(year) || rowMonth < 1 || rowMonth > 12) return;
                const m = rowMonth - 1; // JS month (0-11)
                for (let d = 1; d <= 31; d++) {
                    const val = row['d' + d];
                    const shiftCode = String(val ?? '').trim();
                    if (shiftCode && shiftCode !== '' && shiftCode !== 'undefined') {
                        schedule[`${rowNurse}_${m}_${year}_${d}`] = shiftCode;
                        loaded++;
                    }
                }
            });
            if (loaded > 0) saveData();
            console.log(`[ANNUAL] ${loaded} turni caricati per anno ${year}`);
        }
    } catch (e) {
        console.warn('[ANNUAL] Errore caricamento dati annuali:', e);
    }
}

function renderAnnualReport() {
    if (!currentUser) return;

    const y = reportMonthDate.getFullYear();
    const months = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
    const displayNurses = isCoordinator ? NURSES : [NURSES.find(n => n.id === currentUser.id)].filter(Boolean);

    if (displayNurses.length === 0) {
        document.getElementById('reportsTab').innerHTML = `<div class="empty-state" style="padding:80px"><div class="empty-icon">📊</div><p>Nessun personale attivo.</p></div>`;
        return;
    }

    // Coleta dados por enfermeira por mês
    const nurseAnnual = [];
    let grandTotalH = 0, grandWorkDays = 0, grandRestDays = 0, grandNights = 0;

    displayNurses.forEach(n => {
        let annualH = 0, annualWork = 0, annualRest = 0, annualNights = 0;
        let annualFE = 0, annualAT = 0, annualOFF = 0, annualWeekends = 0;
        const monthlyH = [];
        const annualShiftCounts = {};

        for (let mo = 0; mo < 12; mo++) {
            const daysInMo = new Date(y, mo + 1, 0).getDate();
            let moH = 0, moWork = 0, moRest = 0, moNights = 0;
            const weekendSet = new Set();

            for (let d = 1; d <= daysInMo; d++) {
                const code = getShiftForMonth(n.id, d, mo, y);
                const sh = SHIFTS[code];
                if (!sh) continue;
                moH += sh.h;
                annualShiftCounts[code] = (annualShiftCounts[code] || 0) + 1;

                if (code === 'OFF') { moRest++; annualOFF++; }
                else if (code === 'FE') { moRest++; annualFE++; }
                else if (code === 'AT') { moRest++; annualAT++; }
                else { moWork++; }
                if (code === 'N') { moNights++; }

                if (isFestivo(new Date(y, mo, 1), d) && !['OFF','FE','AT'].includes(code)) {
                    weekendSet.add(d);
                }
            }
            annualWeekends += weekendSet.size;
            annualH += moH; annualWork += moWork; annualRest += moRest; annualNights += moNights;
            monthlyH.push(moH);
        }

        grandTotalH += annualH; grandWorkDays += annualWork; grandRestDays += annualRest; grandNights += annualNights;
        nurseAnnual.push({
            nurse: n, annualH, annualWork, annualRest, annualNights,
            annualFE, annualAT, annualOFF, annualWeekends, monthlyH, annualShiftCounts
        });
    });

    const avgH = displayNurses.length > 0 ? (grandTotalH / displayNurses.length).toFixed(1) : '0.0';
    const maxH = Math.max(...nurseAnnual.map(nd => nd.annualH));
    const minH = Math.min(...nurseAnnual.map(nd => nd.annualH));

    const tab = document.getElementById('reportsTab');
    tab.innerHTML = `
    <div class="reports-wrap" style="max-width:1200px;">
        <div class="rpt-header">
            <h2>📊 Dashboard Operativo</h2>
            <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap;">
                <div class="rpt-view-toggle">
                    <button class="rpt-toggle-btn ${reportViewMode==='monthly'?'active':''}" onclick="toggleReportView('monthly')">Mensile</button>
                    <button class="rpt-toggle-btn ${reportViewMode==='annual'?'active':''}" onclick="toggleReportView('annual')">Annuale</button>
                </div>
                <div class="rpt-month-nav">
                    <button class="rpt-nav-btn" onclick="changeReportYear(-1)" title="Anno precedente">◀</button>
                    <span class="rpt-month">Anno ${y}</span>
                    <button class="rpt-nav-btn" onclick="changeReportYear(1)" title="Anno successivo">▶</button>
                </div>
            </div>
        </div>

        <!-- KPI Annuali -->
        <div class="stats-grid" style="grid-template-columns: repeat(auto-fit, minmax(160px,1fr));">
            <div class="stat-card stat-primary">
                <div class="stat-lbl">Ore Totali Anno</div>
                <div class="stat-val" style="font-size:28px;">${grandTotalH.toFixed(1)}h</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">Media: ${avgH}h/persona</div>
            </div>
            <div class="stat-card stat-success">
                <div class="stat-lbl">Giorni Lavorati</div>
                <div class="stat-val" style="font-size:28px;">${grandWorkDays}</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">Riposo: ${grandRestDays}</div>
            </div>
            <div class="stat-card stat-night">
                <div class="stat-lbl">Turni Notturni</div>
                <div class="stat-val" style="font-size:28px;">${grandNights}</div>
            </div>
            <div class="stat-card stat-warning">
                <div class="stat-lbl">Dispersione Ore</div>
                <div class="stat-val" style="font-size:28px;">${(maxH - minH).toFixed(1)}h</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">Max-Min annuale</div>
            </div>
        </div>

        <!-- Ranking Anual -->
        <div class="report-section">
            <h3>⏱ Ranking Ore Annuale per Infermiera</h3>
            <table class="rpt-table">
                <thead><tr>
                    <th style="text-align:left">Infermiera</th>
                    <th>Ore Anno</th>
                    <th>Lavorati</th>
                    <th>Riposo</th>
                    <th>Notti</th>
                    <th>Ferie</th>
                    <th>Certif.</th>
                    <th>Festivi</th>
                    <th>Carico</th>
                </tr></thead>
                <tbody>
                ${nurseAnnual.sort((a,b) => b.annualH - a.annualH).map((nd, idx) => {
                    const pct = maxH > 0 ? ((nd.annualH / maxH) * 100).toFixed(0) : 0;
                    const barColor = nd.annualH >= maxH ? 'var(--danger)' : nd.annualH <= minH ? 'var(--success)' : 'var(--primary)';
                    return `<tr>
                        <td style="text-align:left"><strong>${idx+1}.</strong> ${nd.nurse.name}</td>
                        <td><strong>${nd.annualH.toFixed(1)}h</strong></td>
                        <td>${nd.annualWork}</td>
                        <td>${nd.annualRest}</td>
                        <td>${nd.annualNights}</td>
                        <td>${nd.annualFE}</td>
                        <td>${nd.annualAT}</td>
                        <td>${nd.annualWeekends}</td>
                        <td style="min-width:120px;">
                            <div style="background:rgba(255,255,255,0.06); border-radius:99px; height:8px; overflow:hidden;">
                                <div style="width:${pct}%; height:100%; background:${barColor}; border-radius:99px; transition:width 0.6s;"></div>
                            </div>
                        </td>
                    </tr>`;
                }).join('')}
                </tbody>
            </table>
        </div>

        <!-- Ore per Mese (Trend) -->
        <div class="report-section">
            <h3>📈 Ore per Mese — Trend Annuale</h3>
            <div style="overflow-x:auto;">
            <table class="rpt-table">
                <thead><tr>
                    <th style="text-align:left">Infermiera</th>
                    ${months.map((mo, i) => `<th style="font-size:11px;">${mo.substring(0,3)}</th>`).join('')}
                    <th><strong>TOT</strong></th>
                </tr></thead>
                <tbody>
                ${nurseAnnual.map(nd => `<tr>
                    <td style="text-align:left">${nd.nurse.name}</td>
                    ${nd.monthlyH.map(h => {
                        const opacity = h > 0 ? '1' : '0.25';
                        return `<td style="font-weight:${h > 0 ? '700' : '400'}; opacity:${opacity}; font-size:12px;">${h > 0 ? h.toFixed(0) : '—'}</td>`;
                    }).join('')}
                    <td><strong>${nd.annualH.toFixed(0)}h</strong></td>
                </tr>`).join('')}
                </tbody>
            </table>
            </div>
        </div>

        <!-- Distribuição de Turnos Anual -->
        <div class="report-section">
            <h3>🔄 Distribuzione Turni Annuale</h3>
            <div style="overflow-x:auto;">
            <table class="rpt-table">
                <thead><tr>
                    <th style="text-align:left">Infermiera</th>
                    ${['M1','M2','MF','G','P','PF','N','OFF','FE','AT'].map(c =>
                        `<th><span style="display:inline-block; width:10px; height:10px; border-radius:3px; background:${SHIFTS[c].color}; margin-right:4px; vertical-align:middle;"></span>${c}</th>`
                    ).join('')}
                </tr></thead>
                <tbody>
                ${nurseAnnual.map(nd => `<tr>
                    <td style="text-align:left">${nd.nurse.name}</td>
                    ${['M1','M2','MF','G','P','PF','N','OFF','FE','AT'].map(c => {
                        const val = nd.annualShiftCounts[c] || 0;
                        return `<td style="font-weight:${val > 0 ? '700' : '400'}; opacity:${val > 0 ? '1' : '0.3'};">${val}</td>`;
                    }).join('')}
                </tr>`).join('')}
                </tbody>
            </table>
            </div>
        </div>

        <!-- Quota Noturna Anual -->
        <div class="report-section">
            <h3>🌙 Quota Notturna Annuale</h3>
            <div style="display:grid; gap:12px;">
            ${nurseAnnual.map(nd => {
                const annualQuota = (nd.nurse.nightQuota || 5) * 12;
                const pct = Math.min((nd.annualNights / annualQuota) * 100, 100).toFixed(0);
                const color = nd.annualNights > annualQuota ? 'var(--danger)' : nd.annualNights >= annualQuota * 0.9 ? 'var(--warning)' : 'var(--primary)';
                return `<div style="display:flex; align-items:center; gap:12px;">
                    <span style="min-width:140px; font-size:13px; font-weight:600; color:var(--text-2);">${nd.nurse.name}</span>
                    <div style="flex:1; background:rgba(255,255,255,0.06); border-radius:99px; height:10px; overflow:hidden;">
                        <div style="width:${pct}%; height:100%; background:${color}; border-radius:99px; transition:width 0.6s;"></div>
                    </div>
                    <span style="min-width:60px; text-align:right; font-size:13px; font-weight:700; color:${color};">${nd.annualNights}/${annualQuota}</span>
                </div>`;
            }).join('')}
            </div>
        </div>
    </div>`;

    // Carrega dados do cloud para meses que faltam
    loadAnnualData(y).then(() => {
        // Verifica se novos dados foram carregados e re-renderiza se necessário
        let hasNewData = false;
        for (let mo = 0; mo < 12; mo++) {
            if (NURSES.some(n => {
                const shift = schedule[`${String(n.id).trim()}_${mo}_${y}_1`];
                return shift && shift !== 'OFF';
            })) { hasNewData = true; break; }
        }
        // Re-renderiza somente se o report anual está ativo e dados mudaram
        if (hasNewData && reportViewMode === 'annual') {
            // Evita loop: só re-renderiza uma vez marcando flag
            if (!renderAnnualReport._loaded) {
                renderAnnualReport._loaded = true;
                renderAnnualReport();
                setTimeout(() => { renderAnnualReport._loaded = false; }, 5000);
            }
        }
    });
}

function changeReportYear(dir) {
    reportMonthDate = new Date(reportMonthDate.getFullYear() + dir, reportMonthDate.getMonth(), 1);
    renderAnnualReport._loaded = false;
    renderReports();
}

// ── INIT ──────────────────────────────────────────────────────
// ── INIT E GERENCIAMENTO DE FUNCIONÁRIOS ──────────────────────

// Salva dados antes de fechar a página (previne perda por crash/reload)
window.addEventListener('beforeunload', () => {
    clearTimeout(_saveTimer);
    _doSave();
});

// Tratamento global de erros não capturados para evitar tela branca
window.addEventListener('error', (e) => {
    console.error('[NurseShift] Erro global:', e.message, e.filename, e.lineno);
});

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
