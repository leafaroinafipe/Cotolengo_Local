// ============================================================
//  NurseShift Pro — app.js v3.0
//  Regras de negócio: algoritmo de escala em 4 fases
//  Todas as solicitações são aprovadas automaticamente
// ============================================================

// ── CONFIG ────────────────────────────────────────────────────
// Bump para v3.0: schema canônico Solicitacoes, swap cross-date no approve, bulkUpsert em Funcionarios
const APP_CACHE_VERSION = 'v3.0';
// Removidas constantes COORD_PASS / NURSE_PASS — eram dead code com senhas em texto puro.
// O Local não exige login; a autenticação real fica no Mobile (com SHA-256).
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
let editingRequestId = null; // ID della richiesta attualmente in modifica (null = nessuna)
let currentMonth  = new Date();
currentMonth.setDate(1); // Força dia 1 para evitar pular mês em dias 31
let selectedCell  = null;
let occurrences = []; // stores { id, nurseId, type, start, end, desc }
let monthlyOrder  = {}; // key: `${month}_${year}` → array of nurseIds
let reportMonthDate = new Date(); // Período do relatório (independente do calendário)
reportMonthDate.setDate(1);
let reportViewMode = 'monthly'; // 'monthly' ou 'annual'
let reportNurseFilter = 'all'; // 'all' ou nurseId para análise individual


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

        // ── MIGRATION: Restaura Melissa Alves (n4) a partir de abril/2026 ──
        // Gated por flag no localStorage para rodar apenas uma vez
        try {
            const MIGRATION_FLAG = 'escala_migration_n4_restore_v1';
            if (!localStorage.getItem(MIGRATION_FLAG)) {
                // Garante que n4 existe em NURSES
                let hasN4 = NURSES.find(n => String(n.id).trim() === 'n4');
                if (!hasN4) {
                    NURSES.push({ id:'n4', name:'Alves Festa Melissa', initials:'AM', nightQuota:5 });
                    localStorage.setItem('escala_nurses', JSON.stringify(NURSES));
                    console.log('[MIGRATION n4] Infermiera Melissa Alves (n4) riaggiunta a NURSES');
                }
                // Adiciona n4 em todos os meses >= abril/2026 (month=3, year=2026)
                const threshold = 2026 * 12 + 3; // April/2026
                let changed = false;
                Object.keys(monthlyOrder || {}).forEach(key => {
                    const parts = key.split('_');
                    if (parts.length !== 2) return;
                    const mm = parseInt(parts[0], 10);
                    const yy = parseInt(parts[1], 10);
                    if (isNaN(mm) || isNaN(yy)) return;
                    const ordVal = yy * 12 + mm;
                    if (ordVal >= threshold) {
                        const arr = Array.isArray(monthlyOrder[key]) ? monthlyOrder[key] : [];
                        if (!arr.includes('n4')) {
                            arr.push('n4');
                            monthlyOrder[key] = arr;
                            changed = true;
                            console.log(`[MIGRATION n4] n4 aggiunto a monthlyOrder[${key}]`);
                        }
                    }
                });
                if (changed) {
                    localStorage.setItem('escala_monthlyOrder', JSON.stringify(monthlyOrder));
                }
                localStorage.setItem(MIGRATION_FLAG, '1');
                console.log('[MIGRATION n4] Completata ripristino Melissa Alves da aprile/2026 in poi');
            }
        } catch (migErr) {
            console.error('[MIGRATION n4] Errore nella migrazione:', migErr);
        }
    } catch(e) { console.error('Erro ao carregar dados locais', e); }
}

// ── UTILITY: Reattiva un'infermiera a partire da un mese specifico ──
// Uso: reactivateNurseFromMonth('n4', 3, 2026)  // month é 0-indexed (3 = aprile)
function reactivateNurseFromMonth(nurseId, fromMonth, fromYear) {
    const nid = String(nurseId).trim();
    if (!nid) return 0;
    const defaults = {
        n1:{ id:'n1', name:'Balla Sabina',        initials:'BS', nightQuota:5 },
        n2:{ id:'n2', name:'Batista Bianca',      initials:'BB', nightQuota:5 },
        n3:{ id:'n3', name:'De Carvalho Eduarda', initials:'CE', nightQuota:5 },
        n4:{ id:'n4', name:'Alves Festa Melissa', initials:'AM', nightQuota:5 },
        n5:{ id:'n5', name:'Delizzeti Sirlene',   initials:'DS', nightQuota:5 },
        n6:{ id:'n6', name:'Moslih Miriam',       initials:'MM', nightQuota:5 },
        n7:{ id:'n7', name:'Kocevska Kristina',   initials:'KK', nightQuota:5 }
    };
    if (!NURSES.find(n => String(n.id).trim() === nid) && defaults[nid]) {
        NURSES.push(defaults[nid]);
        localStorage.setItem('escala_nurses', JSON.stringify(NURSES));
    }
    const threshold = fromYear * 12 + fromMonth;
    let touched = 0;
    Object.keys(monthlyOrder || {}).forEach(key => {
        const parts = key.split('_');
        if (parts.length !== 2) return;
        const mm = parseInt(parts[0], 10);
        const yy = parseInt(parts[1], 10);
        if (isNaN(mm) || isNaN(yy)) return;
        if ((yy * 12 + mm) >= threshold) {
            const arr = Array.isArray(monthlyOrder[key]) ? monthlyOrder[key] : [];
            if (!arr.includes(nid)) {
                arr.push(nid);
                monthlyOrder[key] = arr;
                touched++;
            }
        }
    });
    if (touched > 0) {
        localStorage.setItem('escala_monthlyOrder', JSON.stringify(monthlyOrder));
        console.log(`[reactivateNurseFromMonth] ${nid} aggiunto a ${touched} mesi`);
    }
    return touched;
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

        // 1. Publicar Funcionários — usa bulkUpsert (não destrói linhas que o Local desconhece;
        //    isso preserva nurses cadastradas em outros clientes/sessões e evita race condition
        //    com criações simultâneas via Mobile).
        const nursesRows = NURSES.map(n => ({
            id:           n.id,
            nome:         n.name,
            attivo:       (n.attivo === false || String(n.attivo) === '0') ? 0 : 1,
            dataInicio:   n.dataInicio || '',
            dataFim:      n.dataFim || '',
            cargaSemanal: n.nightQuota || 5,
            notas:        n.notas || ''
        }));
        const funcUrl = `${GOOGLE_API_URL}?action=bulkUpsert&sheetName=Funcionarios&apiKey=${API_KEY}`;
        await fetch(funcUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ keyColumn: 'id', rows: nursesRows, deleteMissing: false }),
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

        // 3. Publicar Solicitações — schema canônico alinhado com o Mobile.
        //    Usamos APENAS os campos do setupHeaders Mobile: nurseId, nurseName,
        //    nurseIdcambio, nursecambio (sem o legado swapNurseId/toNurseId).
        //    Os aliases continuam aceitos no LEITOR (syncRequestsFromCloud), garantindo
        //    retrocompatibilidade com requests antigas, mas o ESCRITOR só publica o canônico.
        const reqUrl = `${GOOGLE_API_URL}?action=bulkUpsert&sheetName=Solicitacoes&apiKey=${API_KEY}`;
        const reqRows = requests.map(r => ({
            id: String(r.id),
            type: r.type,
            status: r.status,
            nurseId: String(r.nurseId || r.fromNurseId || '').trim(),
            nurseName: r.nurseName || r.fromNurseName || '',
            nurseIdcambio: String(r.nurseIdcambio || r.swapNurseId || r.toNurseId || '').trim(),
            nursecambio: r.nursecambio || r.swapNurseName || r.toNurseName || '',
            startDate: r.startDate || r.date || '',
            endDate: r.endDate || r.startDate || r.date || '',
            // Cross-date swap (vazias para tipos não-swap)
            dataRichiedente: r.dataRichiedente || '',
            dataCambio: r.dataCambio || '',
            turnoRichiedente: r.turnoRichiedente || '',
            turnoCambio: r.turnoCambio || '',
            desc: r.desc || r.reason || '',
            autoApplied: (r.autoApplied === true) ? 'TRUE' : (r.autoApplied || ''),
            createdAt: r.createdAt || '',
            approvedAt: r.approvedAt || '',
            approvedBy: r.approvedBy || ''
        }));
        await fetch(reqUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            // deleteMissing=true: a publicação completa do Local é a fonte da verdade
            // para Solicitacoes naquele instante. Mantém comportamento anterior do clearAll.
            body: JSON.stringify({ keyColumn: 'id', rows: reqRows, deleteMissing: true }),
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
                const pChip = document.getElementById('cloudStatusIndicator');
                const pTxt  = document.getElementById('cloudStatusText');
                const setChipState = (state, text) => {
                    if (pChip) {
                        pChip.classList.remove('is-connecting', 'is-online', 'is-offline');
                        pChip.classList.add(state);
                    }
                    if (pTxt) pTxt.textContent = text;
                };
                
                if (dbTest && dbTest.status === 'success') {
                    if (dbTest.data && dbTest.data.length > 0) {
                        try {
                            let mergedNurses = false;
                            dbTest.data.forEach(cn => {
                                // Schema novo (id/nome/attivo/cargaSemanal) com fallback ao antigo (ID_Funcionario/Nome/Carga_Horaria_Mensal)
                                const id = String(cn.id || cn.ID_Funcionario || '');
                                const name = String(cn.nome || cn.Nome || cn.name || '');
                                const quota = parseInt(cn.cargaSemanal || cn.Carga_Horaria_Mensal || cn.nightQuota) || 5;
                                // Filtros: ignora vazios, Coordinatrice (n0), e inativos (attivo === 0/false)
                                if (!id || id === 'undefined' || !name || name === 'undefined') return;
                                if (id === 'n0') return;
                                const ativoRaw = cn.attivo;
                                if (ativoRaw !== undefined && ativoRaw !== null && ativoRaw !== '') {
                                    const v = String(ativoRaw).trim().toLowerCase();
                                    if (v === '0' || v === 'false' || v === 'no') return;
                                }

                                const localNurse = NURSES.find(n => String(n.id) === id);
                                if (!localNurse) {
                                    NURSES.push({
                                        id: id,
                                        name: name,
                                        initials: name.split(' ').map(w => w[0] || '').join('').substring(0,2).toUpperCase(),
                                        nightQuota: quota,
                                        dataInicio: cn.dataInicio || '',
                                        dataFim:    cn.dataFim || '',
                                        notas:      cn.notas || ''
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

                    setChipState('is-online', 'App Sincronizzato');

                    toast('🟢 Sistema Online connesso al Database cloud!', 'success', 3500);

                    // Dispara a sincronização de turnos pendentes e requisições
                    await syncScheduleFromCloud();
                    await syncRequestsFromCloud();
                    applyApprovedRequests(); // Auto-aplica requisições aprovadas na escala
                    renderRequests();
                    renderOccurrences();
                } else {
                    setChipState('is-offline', 'Offline (Errore)');
                    console.warn("[CLOUD DB] Falha no teste inicial:", dbTest);
                    toast('🔴 Modalità offline attiva. Connessione persa.', 'warning', 5000);
                }
            }, 600);

        }, 500); 
    }, 2200); 
}

function logout() {
    // Limpa estado local antes de recarregar. Sem isso, "logout" era cosmético —
    // o reload reabria com os mesmos dados em cache.
    if (!confirm("Vuoi davvero uscire?\n\nVerranno svuotati i dati in cache su questo dispositivo. Le informazioni nel cloud rimangono intatte.")) return;
    try {
        // Salva pendências antes de limpar (segurança extra)
        clearTimeout(_saveTimer);
        _doSave();
    } catch (e) { /* nothing to do */ }
    try {
        // Apaga apenas as chaves do app, preserva flags de migração e biometria.
        const APP_KEYS = [
            'escala_nurses',
            'escala_schedule',
            'escala_occurrences',
            'escala_monthlyOrder',
            'escala_requests',
            'escala_app_version'
        ];
        APP_KEYS.forEach(k => localStorage.removeItem(k));
        console.log('[LOGOUT] localStorage do app limpo:', APP_KEYS.length, 'chaves');
    } catch (e) {
        console.warn('[LOGOUT] Falha ao limpar localStorage:', e);
    }
    location.reload();
}

// ── LEGEND ────────────────────────────────────────────────────
// OFF tem cor de célula `rgba(255,255,255,0.03)` (proposital sutileza no calendário),
// mas na legenda essa cor fica invisível. Resolvemos com cor de display dedicada.
const LEGEND_DISPLAY_COLOR = {
    OFF: 'var(--text-4)' // cinza médio — mantém código semântico, fica legível
};
function buildLegend() {
    const codes = ['M1','M2','MF','G','P','PF','N','OFF','FE','AT'];
    document.getElementById('shiftLegend').innerHTML = codes.map(c => {
        const s = SHIFTS[c];
        const dotColor = LEGEND_DISPLAY_COLOR[c] || s.color;
        return `<div class="legend-item"><div class="legend-dot" style="background:${dotColor};border:1px solid rgba(255,255,255,0.08)"></div>${s.name}</div>`;
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
                    // Repara nomes/IDs ausentes em requests antigas (ex.: o Mobile gravou só o ID
                    // e a request foi cacheada localmente com nome vazio ANTES da correção)
                    const crSwapId = String(cr.nurseIdcambio || cr.swapNurseId || cr.toNurseId || '').trim();
                    const crSwapName = cr.nursecambio || cr.swapNurseName || cr.toNurseName || '';
                    const crNurseId = String(cr.nurseId || '').trim();
                    const crNurseName = cr.nurseName || '';

                    // Garante ID do cambio preenchido
                    if (crSwapId && !local.nurseIdcambio) { local.nurseIdcambio = crSwapId; isModified = true; }
                    if (crSwapId && !local.swapNurseId)   { local.swapNurseId   = crSwapId; isModified = true; }
                    if (crSwapId && !local.toNurseId)     { local.toNurseId     = crSwapId; isModified = true; }

                    // Garante nome do cambio preenchido (cloud ou fallback via NURSES)
                    if (!local.nursecambio || !local.swapNurseName || !local.toNurseName) {
                        let resolvedSwapName = crSwapName;
                        if (!resolvedSwapName && crSwapId) {
                            const nFound = NURSES.find(n => String(n.id).trim() === crSwapId);
                            if (nFound) resolvedSwapName = nFound.name;
                        }
                        if (resolvedSwapName) {
                            if (!local.nursecambio)   local.nursecambio   = resolvedSwapName;
                            if (!local.swapNurseName) local.swapNurseName = resolvedSwapName;
                            if (!local.toNurseName)   local.toNurseName   = resolvedSwapName;
                            isModified = true;
                        }
                    }

                    // Repara também o solicitante se veio sem nome
                    if (crNurseId && !local.nurseId) { local.nurseId = crNurseId; isModified = true; }
                    if (!local.nurseName && !local.fromNurseName) {
                        let resolvedReqName = crNurseName;
                        if (!resolvedReqName && crNurseId) {
                            const nFound = NURSES.find(n => String(n.id).trim() === crNurseId);
                            if (nFound) resolvedReqName = nFound.name;
                        }
                        if (resolvedReqName) {
                            local.nurseName = resolvedReqName;
                            local.fromNurseName = resolvedReqName;
                            isModified = true;
                        }
                    }

                    updatedRequests.push(local);
                } else {
                    // Sanitiza datas vindas do cloud: Google Sheets pode retornar
                    // "2026-05-23T03:00:00Z" — o split('T')[0] garante "2026-05-23"
                    const rawDate = sanitizeDate(cr.startDate || cr.date || '');
                    // Mapeia os campos 'nursecambio' e 'nurseIdcambio' do Google Sheets para os campos internos do app
                    const swapNurseId = String(cr.nurseIdcambio || cr.swapNurseId || cr.toNurseId || '').trim();
                    // Se o Mobile não gravou o nome (só o ID), resolvemos pelo NURSES local
                    let swapNurseName = cr.nursecambio || cr.swapNurseName || cr.toNurseName || '';
                    if (!swapNurseName && swapNurseId) {
                        const nFound = NURSES.find(n => String(n.id).trim() === swapNurseId);
                        if (nFound) swapNurseName = nFound.name;
                    }
                    // Também resolve o nome do solicitante caso só o ID tenha vindo
                    let reqNurseName = cr.nurseName || '';
                    const reqNurseId = String(cr.nurseId || '').trim();
                    if (!reqNurseName && reqNurseId) {
                        const nFound = NURSES.find(n => String(n.id).trim() === reqNurseId);
                        if (nFound) reqNurseName = nFound.name;
                    }
                    // Nuove colonne swap cross-date dal cloud
                    const dataRichiedenteIso = sanitizeDate(cr.dataRichiedente || '');
                    const dataCambioIso      = sanitizeDate(cr.dataCambio || '');
                    const turnoRichiedente   = String(cr.turnoRichiedente || '').trim();
                    const turnoCambio        = String(cr.turnoCambio || '').trim();
                    updatedRequests.push({
                        id: crId,
                        type: cr.type || 'OFF',
                        status: cr.status || 'pending',
                        nurseId: reqNurseId,
                        fromNurseId: reqNurseId,
                        nurseName: reqNurseName,
                        fromNurseName: reqNurseName,
                        toNurseName: swapNurseName,
                        swapNurseName: swapNurseName,
                        nursecambio: swapNurseName,
                        nurseIdcambio: swapNurseId,
                        swapNurseId: swapNurseId,
                        toNurseId: swapNurseId,
                        startDate: rawDate,
                        date: rawDate,
                        endDate: sanitizeDate(cr.endDate || ''),
                        desc: cr.desc || '',
                        reason: cr.desc || '',
                        // Campi swap cross-date
                        dataRichiedente: dataRichiedenteIso,
                        dataCambio: dataCambioIso,
                        turnoRichiedente: turnoRichiedente,
                        turnoCambio: turnoCambio,
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
                applyApprovedRequests(); // Auto-aplica novas aprovações na escala
                toast('☁️ Richieste sincronizzate e allineate col cloud', 'info');
            }
        }
    } catch (e) {
        console.warn('Sync cloud requests:', e);
    }
}

// ── AUTO-APLICAÇÃO DE REQUISIÇÕES APROVADAS NA ESCALA ──────
// Cruza requisições aprovadas com a escala existente.
// Se uma requisição aprovada (FE, AT, OFF, vacation, justified, swap)
// cai em um dia que já tem turno na escala, substitui automaticamente.
// Isso garante que aprovações feitas DEPOIS da geração da escala
// sejam refletidas sem intervenção manual.
function applyApprovedRequests() {
    const m = currentMonth.getMonth();
    const y = currentMonth.getFullYear();
    const days = daysInMonth(currentMonth);
    let applied = 0;

    requests.forEach(req => {
        if (req.status !== 'approved') return;

        // Requisições de ausência: FE, AT, OFF, vacation, justified, OFF_INJ
        if (['FE', 'OFF', 'AT', 'OFF_INJ', 'vacation', 'justified'].includes(req.type)) {
            let code = 'OFF';
            if (req.type === 'FE' || req.type === 'vacation') code = 'FE';
            else if (req.type === 'AT') code = 'AT';

            const startStr = sanitizeDate(req.startDate || req.date || '');
            const endStr = sanitizeDate(req.endDate || startStr);
            if (!startStr) return;

            const sDate = new Date(startStr + 'T00:00:00');
            const eDate = new Date((endStr || startStr) + 'T00:00:00');

            for (let d = 1; d <= days; d++) {
                const checkDate = new Date(y, m, d);
                checkDate.setHours(0,0,0,0); sDate.setHours(0,0,0,0); eDate.setHours(0,0,0,0);
                if (checkDate >= sDate && checkDate <= eDate) {
                    const currentShift = getShift(req.nurseId, d);
                    // Só aplica se o dia ainda NÃO tem o código correto
                    if (currentShift !== code) {
                        assign(req.nurseId, d, code);
                        applied++;
                    }
                }
            }
        }
        // Requisições de troca (swap) — supporta cross-date (4-cell symmetric swap)
        else if (req.type === 'swap') {
            // Já aplicado? Aceita tanto flag em memória (true) quanto vinda do cloud ('TRUE' string).
            // O Mobile pode ter aplicado antes; nesse caso, pulamos para evitar duplicação.
            if (req.autoApplied === true) return;
            if (typeof req.autoApplied === 'string' && req.autoApplied.trim().toUpperCase() === 'TRUE') {
                req.autoApplied = true; // normaliza para boolean local
                return;
            }

            // Resolve IDs aceitando nomes do Mobile (nurseIdcambio) e do Local (swapNurseId/toNurseId)
            const nurseId     = String(req.nurseId     || req.fromNurseId || '').trim();
            const swapNurseId = String(req.nurseIdcambio || req.swapNurseId || req.toNurseId || '').trim();
            if (!nurseId || !swapNurseId) return;

            // Datas: dataRichiedente = data ceduta dal richiedente; dataCambio = data ceduta dalla controparte.
            // Retrocompatibilidade: swaps antigos só têm startDate/date (uma única data).
            const fromDateStr = sanitizeDate(req.dataRichiedente || req.startDate || req.date || '');
            const toDateStr   = sanitizeDate(req.dataCambio      || req.startDate || req.date || '');
            if (!fromDateStr || !toDateStr) return;

            const fromDate = new Date(fromDateStr + 'T00:00:00');
            const toDate   = new Date(toDateStr   + 'T00:00:00');
            // Ambas as datas devem estar no mês atual para aplicar atomicamente
            if (fromDate.getMonth() !== m || fromDate.getFullYear() !== y) return;
            if (toDate.getMonth()   !== m || toDate.getFullYear()   !== y) return;
            const fromDay = fromDate.getDate();
            const toDay   = toDate.getDate();
            if (fromDay < 1 || fromDay > days || toDay < 1 || toDay > days) return;

            const snapFrom = String(req.turnoRichiedente || '').trim();
            const snapTo   = String(req.turnoCambio      || '').trim();

            // ── SAME-DAY SWAP (legado ou cross-date que calhou no mesmo dia) ──
            if (fromDay === toDay) {
                const shiftA = getShift(nurseId,     fromDay);
                const shiftB = getShift(swapNurseId, fromDay);
                if (shiftA !== shiftB) {
                    assign(nurseId,     fromDay, shiftB);
                    assign(swapNurseId, fromDay, shiftA);
                    req.autoApplied = true;
                    applied++;
                }
                return;
            }

            // ── CROSS-DATE SWAP: 4-cell symmetric swap ──
            const reqOnFrom = getShift(nurseId,     fromDay);
            const cpOnFrom  = getShift(swapNurseId, fromDay);
            const reqOnTo   = getShift(nurseId,     toDay);
            const cpOnTo    = getShift(swapNurseId, toDay);

            // Detecta se já foi aplicado (estado post-swap matching snapshots)
            if (snapFrom && snapTo && cpOnFrom === snapFrom && reqOnTo === snapTo) {
                req.autoApplied = true;
                return;
            }

            // Se há snapshots, valida estado pré-swap antes de aplicar
            if (snapFrom && snapTo) {
                if (reqOnFrom !== snapFrom || cpOnTo !== snapTo) {
                    console.warn(`[AUTO-APPLY] Swap ${req.id}: stato incoerente con snapshot, saltato.`, {
                        atteso: { reqOnFrom: snapFrom, cpOnTo: snapTo },
                        trovato: { reqOnFrom, cpOnTo }
                    });
                    return;
                }
            }

            // Aplica 4-cell symmetric swap
            assign(nurseId,     fromDay, cpOnFrom);
            assign(swapNurseId, fromDay, reqOnFrom);
            assign(nurseId,     toDay,   cpOnTo);
            assign(swapNurseId, toDay,   reqOnTo);
            req.autoApplied = true;
            applied++;
        }
    });

    if (applied > 0) {
        saveData();
        renderCalendar();
        console.log(`[AUTO-APPLY] ${applied} alterações aplicadas de requisições aprovadas.`);
        toast(`🔄 ${applied} turno(i) aggiornato(i) da richieste approvate`, 'info', 3000);

        // Auto-publica alterações no cloud para que o Mobile veja
        autoPublishMonth(m, y, days);

        // Persiste autoApplied=TRUE no cloud para os swaps recém aplicados
        // (proteção contra reaplicação em outro dispositivo)
        requests.forEach(req => {
            if (req.type === 'swap' && req.autoApplied === true && !req._cloudFlagged) {
                req._cloudFlagged = true; // marca local para não republicar toda vez
                fetch(`${GOOGLE_API_URL}?action=update&sheetName=Solicitacoes&apiKey=${API_KEY}`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({ _keyColumn: 'id', _keyValue: String(req.id), autoApplied: 'TRUE' })
                }).catch(e => console.warn('[AUTO-APPLY] Errore flag autoApplied:', e));
            }
        });
    }
}

// Publica apenas a escala do mês alvo no cloud (sem tocar funcionários/requisições)
// IMPORTANTE: usa getShiftForMonth (não getShift) para que funcione mesmo quando
// o mês alvo é DIFERENTE do currentMonth (ex.: aprovação de troca em outro mês)
async function autoPublishMonth(m, y, days) {
    try {
        const displayNurses = getMonthlyNurses();
        const escalaRows = displayNurses.map(nurse => {
            const row = { nurseId: nurse.id, month: String(m + 1), year: String(y) };
            for (let d = 1; d <= 31; d++) {
                row['d' + d] = d <= days ? getShiftForMonth(nurse.id, d, m, y) : '';
            }
            return row;
        });
        const escalaUrl = `${GOOGLE_API_URL}?action=bulkWrite&sheetName=Escala&apiKey=${API_KEY}`;
        await fetch(escalaUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                clearFilter: [{ column: 'month', value: String(m + 1) }, { column: 'year', value: String(y) }],
                rows: escalaRows
            })
        });
        console.log(`[AUTO-PUBLISH] Escala ${m+1}/${y} pubblicata automaticamente nel cloud.`);
    } catch (e) {
        console.warn('[AUTO-PUBLISH] Errore nella pubblicazione automatica:', e);
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
        const pChip = document.getElementById('cloudStatusIndicator');
        const pTxt  = document.getElementById('cloudStatusText');
        if (pChip) {
            pChip.classList.remove('is-online', 'is-offline');
            pChip.classList.add('is-connecting');
        }
        if (pTxt) pTxt.textContent = 'Sincronizzando...';

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
            applyApprovedRequests(); // Auto-aplica requisições aprovadas pós-sync

            if (loaded > 0) {
                toast(`☁️ Turni del mese scaricati dal cloud!`, 'info', 2000);
            }
            if (pChip) {
                pChip.classList.remove('is-connecting', 'is-offline');
                pChip.classList.add('is-online');
            }
            if (pTxt) pTxt.textContent = 'App Sincronizzato';
        } else {
            console.warn('[SYNC] Nessuna risposta valida dal cloud:', dbRes);
            if (pChip) {
                pChip.classList.remove('is-connecting', 'is-online');
                pChip.classList.add('is-offline');
            }
            if (pTxt) pTxt.textContent = 'Offline';
        }
    } catch (e) {
        console.error("[SYNC] Errore download turni:", e);
        const pChip = document.getElementById('cloudStatusIndicator');
        const pTxt  = document.getElementById('cloudStatusText');
        if (pChip) {
            pChip.classList.remove('is-connecting', 'is-online');
            pChip.classList.add('is-offline');
        }
        if (pTxt) pTxt.textContent = 'Offline';
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
// ── Wrapper: tenta executar a geração em Web Worker (não trava a UI). ──
// Se Worker não estiver disponível ou falhar, cai no path inline original.
async function generateSchedule(hourLimits = {}, startDay = 1) {
    const overlay = document.getElementById('loadingOverlay');
    overlay.classList.remove('hidden');
    await new Promise(r => setTimeout(r, 60)); // permite render do spinner

    const m = currentMonth.getMonth();
    const y = currentMonth.getFullYear();
    const days = daysInMonth(currentMonth);

    // Tenta o caminho rápido via Worker
    if (typeof Worker !== 'undefined') {
        try {
            const result = await runGeneratorInWorker({ hourLimits, startDay, m, y });
            applyWorkerResult(result, m, y, days, startDay);
            overlay.classList.add('hidden');
            renderCalendar();
            if (result.emptyShifts > 0) {
                toast(`⚠️ Turni completati, PERÒ ${result.emptyShifts} turni vuoti irrimediabili questo mese!`, 'error', 8000);
            } else {
                toast('Turni Smart generati! Il problema matriciale ha avuto il 100% di completamento. ✨', 'success', 5000);
            }
            saveData();
            return;
        } catch (err) {
            console.warn('[GEN] Worker indisponibile, fallback inline:', err && err.message);
            // continua para o inline abaixo
        }
    }

    // ── FALLBACK INLINE (preserva comportamento legacy se o Worker falhar) ──
    return generateScheduleInline(hourLimits, startDay);
}

// Coloca o melhor model retornado pelo worker no `schedule` global,
// preservando dias < startDay (passado congelado) e descartando lookahead (d > days).
function applyWorkerResult(result, m, y, days, startDay) {
    const prefixRegex = new RegExp(`_${m}_${y}_\\d+$`);
    for (let k in schedule) {
        if (!prefixRegex.test(k)) continue;
        const parts = k.split('_');
        const dInt = parseInt(parts[parts.length - 1], 10);
        if (dInt >= startDay) delete schedule[k];
    }
    for (const key in result.scheduleModel) {
        const parts = key.split('_');
        const dInt = parseInt(parts[parts.length - 1], 10);
        if (dInt >= startDay && dInt <= days) {
            schedule[key] = result.scheduleModel[key];
        }
    }
}

// Promisifica a comunicação com o Web Worker. Resolve com {scheduleModel, fitness, emptyShifts}
// ou rejeita com Error em caso de falha. Captura mensagens de progresso para atualizar
// o texto do overlay e dar feedback visual à coordenadora durante os 1200 epochs.
function runGeneratorInWorker({ hourLimits, startDay, m, y }) {
    return new Promise((resolve, reject) => {
        let worker;
        try {
            worker = new Worker('escala-worker.js?v=3.0');
        } catch (e) {
            reject(new Error('Worker constructor falhou: ' + e.message));
            return;
        }

        const loadingTxt = document.querySelector('.loading-txt');
        const loadingSub = document.querySelector('.loading-sub');
        if (loadingTxt) loadingTxt.textContent = 'Generazione turni intelligenti...';
        if (loadingSub) loadingSub.textContent = 'Esplorazione di soluzioni candidate...';

        worker.onmessage = (ev) => {
            const data = ev.data || {};
            if (data.progress) {
                if (loadingSub) {
                    loadingSub.textContent = `Epoch ${data.epoch}/1200 — fitness ${Math.round(data.bestFitness).toLocaleString('it-IT')}`;
                }
                return;
            }
            if (!data.ok) {
                worker.terminate();
                reject(new Error(data.error || 'Worker reportou erro'));
                return;
            }
            // Resultado final
            worker.terminate();
            resolve(data);
        };

        worker.onerror = (err) => {
            worker.terminate();
            reject(new Error('Worker onerror: ' + (err.message || 'unknown')));
        };

        // Filtra requests apenas relevantes (status approved) para reduzir payload
        const approvedRequests = requests.filter(r => r.status === 'approved');
        worker.postMessage({
            NURSES,
            schedule,                // somente leitura no worker
            occurrences,
            requests: approvedRequests,
            hourLimits,
            m, y,
            startDay,
            MAX_EPOCHS: 1200
        });
    });
}

// ── PATH INLINE (fallback legacy, mantido idêntico ao código original) ──
async function generateScheduleInline(hourLimits = {}, startDay = 1) {
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

    // ── CONTEXTO DOS ÚLTIMOS 4 DIAS DO MÊS ANTERIOR ──
    // Permite que o gerador respeite regras de progressão (ex.: N → OFF → OFF)
    // quando os primeiros dias do mês atual dependem do final do mês anterior.
    // Mapeamento: day=0 → último dia do mês anterior, day=-1 → penúltimo, day=-2, day=-3.
    const prevM = m === 0 ? 11 : m - 1;
    const prevY = m === 0 ? y - 1 : y;
    const prevMonthDaysCount = new Date(prevY, prevM + 1, 0).getDate();
    const prevContext = {}; // { nurseId: { '0': code, '-1': code, '-2': code, '-3': code } }
    try {
        NURSES.forEach(n => {
            prevContext[n.id] = {};
            for (let off = 0; off <= 3; off++) {
                const pDay = prevMonthDaysCount - off;
                if (pDay < 1) continue;
                const code = getShiftForMonth(n.id, pDay, prevM, prevY);
                // Só registra se houver dado real (diferente do default 'OFF') para não impor OFF fictício
                const key = `${String(n.id).trim()}_${prevM}_${prevY}_${pDay}`;
                if (schedule.hasOwnProperty(key) && code) {
                    prevContext[n.id][-off] = code;
                }
            }
        });
        console.log('[GEN] Contexto últimos 4 dias mês anterior carregado:', Object.keys(prevContext).length, 'enfermeiros');
    } catch (prevErr) {
        console.warn('[GEN] Falha ao carregar contexto do mês anterior:', prevErr);
    }

    // Métodos Contextuais Temporários (Trabalham apenas em um objeto genérico, rápido para simulação RAM)
    function getSh(simObj, nId, d) {
        // Para dias <= 0, consulta o contexto dos últimos 4 dias do mês anterior
        if (d <= 0) {
            const ctx = prevContext[nId];
            return ctx ? ctx[d] : undefined;
        }
        return simObj[`${nId}_${m}_${y}_${d}`];
    }
    function setSh(simObj, nId, d, code) {
        // Nunca escreve no passado (mês anterior é somente leitura)
        if (d <= 0) return;
        simObj[`${nId}_${m}_${y}_${d}`] = code;
    }

    function nurseHoursTemp(simObj, nId) {
        let h = 0;
        // Importante: Considera-se impacto contratual de Horas APENAS nos dias pertencentes a este mês (<= days)
        for (let d=1; d<=days; d++) h += (SHIFTS[getSh(simObj,nId,d)]?.h||0);
        return h;
    }

    function canWorkConsecTemp(simObj, nId, day) {
        let count = 1;
        let d = day - 1;
        // Limite inferior: -3 para consultar os últimos 4 dias do mês anterior (offsets 0,-1,-2,-3)
        while (d >= -3 && getSh(simObj, nId, d) && !['OFF','FE','AT'].includes(getSh(simObj, nId, d))) { count++; d--; }
        d = day + 1;
        // Look ahead vai até simDays no contínuo dos plantões
        while (d <= simDays && getSh(simObj, nId, d) && !['OFF','FE','AT'].includes(getSh(simObj, nId, d))) { count++; d++; }
        if (count > MAX_CONSEC) return false;

        // Limite restritivo de fadiga - Bloco grande único
        if (count >= 4) {
            let startS = day;
            while(startS > -3 && getSh(simObj, nId, startS-1) && !['OFF','FE','AT'].includes(getSh(simObj, nId, startS-1))) startS--;
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
        // Consulta dia anterior (inclui dias -3..0 vindos do mês anterior)
        const prev = getSh(simObj, nId, day - 1) || null;
        const next = day < simDays ? getSh(simObj, nId, day + 1) : null;
        if (morningShifts.includes(code) && afternoonShifts.includes(prev)) return false;
        if (afternoonShifts.includes(code) && morningShifts.includes(next)) return false;
        if (afternoonShifts.includes(code) && (afternoonShifts.includes(prev) || afternoonShifts.includes(next))) return false;
        // Regra N → OFF: após um turno de noite, o dia seguinte deve ser repouso
        if (prev === 'N' && !['OFF','FE','AT'].includes(code)) return false;
        return true;
    }

    function canAssignTemp(simObj, nId, day, code) {
        if (!code || ['OFF', 'FE', 'AT'].includes(code)) return true;
        if (!checkTransitions(simObj, nId, day, code)) return false;

        // Impede ilhas de 1 dia isolado para descanso — agora considera também os últimos 4 dias do mês anterior
        const prev = getSh(simObj, nId, day - 1);
        if (prev && ['OFF', 'FE', 'AT'].includes(prev)) {
            const prev2 = getSh(simObj, nId, day - 2);
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
            // Permite d === 1 também: se o último dia do mês anterior foi N, o domingo do novo mês espelha.
            if (dow === 0) {
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
            
            // Scoring inteligente para noturno: considera nightCount, horas acumuladas e limite individual
            eligible.forEach(n => {
                const nLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                const nHours = nurseHoursTemp(tSched, n.id);
                const loadRatio = nHours / nLimit; // % do limite já consumido
                const quota = n.nightQuota || 5;
                const nightRatio = nightCount[n.id] / quota; // % da cota noturna já usada
                // Score composto: menor = melhor candidata para noite
                n._nightScore = (nightRatio * 100) + (loadRatio * 80) + (Math.random() * 5);
            });
            eligible.sort((a,b) => a._nightScore - b._nightScore);

            if (eligible.length > 0) {
                let pair = eligible.find(n => getSh(tSched, n.id, d-1)==='N');
                let chosen = pair;

                if (!chosen) {
                    // Seleciona entre as top candidatas com score similar
                    const topScore = eligible[0]._nightScore;
                    const threshold = topScore + 10;
                    const topNurses = eligible.filter(n => n._nightScore <= threshold);
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
                        const nLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                        const nHours = nurseHoursTemp(tSched, n.id);
                        const p1 = d>1 ? getSh(tSched, n.id, d-1) : null;

                        // 1. Carga horária relativa ao limite individual (0 a 1+)
                        //    Enfermeiras com limite menor (ex: Sirlene) são penalizadas mais cedo
                        const loadRatio = nHours / nLimit;
                        const loadPenalty = loadRatio * 200;

                        // 2. Penaliza turnos consecutivos do mesmo tipo
                        const seqPenalty = (p1 === t) ? 300 : 0;

                        // 3. Equilíbrio do tipo de turno específico: quem já fez mais desse tipo pesa mais
                        const typeCount = shiftCountTemp(n.id, t);
                        const typePenalty = typeCount * 25;

                        // 4. Equilíbrio G/M2 por enfermeira (evita acumular só G ou só M2)
                        let gm2Bias = 0;
                        if (t === 'G') gm2Bias = (nurseGM2Count[n.id]?.g || 0) * 18;
                        if (t === 'M2') gm2Bias = (nurseGM2Count[n.id]?.m2 || 0) * 18;

                        // 5. Equilíbrio P vs M1: quem já tem muitos P ou M1 é desfavorecido
                        let pmBias = 0;
                        if (t === 'P') pmBias = shiftCountTemp(n.id, 'P') * 15;
                        if (t === 'M1') pmBias = shiftCountTemp(n.id, 'M1') * 15;

                        // 6. Penalidade de festivo: favorecer quem trabalhou menos festivos
                        let wkBias = 0;
                        if (festivo) {
                            let wkCount = 0;
                            for (let wd = 1; wd < d; wd++) {
                                if (isFestivo(monthRef, wd) && getSh(tSched, n.id, wd) && !['OFF','FE','AT'].includes(getSh(tSched, n.id, wd))) wkCount++;
                            }
                            wkBias = wkCount * 20;
                        }

                        // 7. Variação diária: quem trabalhou ontem é levemente desfavorecido
                        const workedYesterday = (p1 && !['OFF','FE','AT'].includes(p1)) ? 8 : 0;

                        // Score composto: menor = melhor candidata
                        n.tmpScore = loadPenalty + typePenalty + seqPenalty + gm2Bias + pmBias + wkBias + workedYesterday + (Math.random() * 4);
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
        // Avalia qualidade da escala considerando distribuição proporcional ao limite individual
        let repScore = 0;
        let weekendPenalty = 0;
        let postNightViolations = 0;
        let shiftTypePenalty = 0;
        let overloadPenalty = 0;
        let loadVariancePenalty = 0;
        const monthRefFit = new Date(y, m, 1);

        // Coleta métricas por enfermeira
        const nurseMetrics = [];
        NURSES.forEach(n => {
            const h = nurseHoursTemp(tSched, n.id);
            const nLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
            const loadPct = h / nLimit; // % do limite utilizado (normalizado)

            // 1. Repetições sequenciais de turnos iguais
            let cons = 0;
            for (let dt=2; dt<=days; dt++) {
                const curr = getSh(tSched, n.id, dt);
                const prev = getSh(tSched, n.id, dt-1);
                if (curr && curr !== 'OFF' && curr !== 'FE' && curr !== 'AT') {
                    if (curr === prev) cons++;
                }
            }
            repScore += cons;

            // 2. Equilíbrio de festivos trabalhados
            let festiviWorked = 0;
            for (let dd = 1; dd <= days; dd++) {
                if (isFestivo(monthRefFit, dd)) {
                    const s = getSh(tSched, n.id, dd);
                    if (s && !['OFF','FE','AT'].includes(s)) festiviWorked++;
                }
            }
            if (festiviWorked > 4) weekendPenalty += (festiviWorked - 4) * 3000;
            if (festiviWorked === 0) weekendPenalty += 2000;

            // 3. Violações de descanso pós-noturno
            for (let dd = 1; dd < days; dd++) {
                if (getSh(tSched, n.id, dd) === 'N') {
                    const next = getSh(tSched, n.id, dd+1);
                    if (next && !['OFF','FE','AT','N'].includes(next)) postNightViolations++;
                }
            }

            // 4. Contagem por tipo de turno
            const counts = {};
            ['M1','M2','MF','G','P','PF','N'].forEach(c => { counts[c] = 0; });
            for (let dd = 1; dd <= days; dd++) {
                const s = getSh(tSched, n.id, dd);
                if (counts[s] !== undefined) counts[s]++;
            }

            // 5. Equilíbrio G vs M2
            shiftTypePenalty += Math.abs(counts['G'] - counts['M2']) * 200;
            // 6. Equilíbrio P vs M1 (devem ser similares em dias úteis)
            shiftTypePenalty += Math.abs(counts['P'] - counts['M1']) * 150;
            // 7. Equilíbrio MF vs PF (devem ser similares em festivos)
            shiftTypePenalty += Math.abs(counts['MF'] - counts['PF']) * 150;

            // 8. Penalidade por sobrecarga: se ultrapassou o limite individual
            if (h > nLimit) overloadPenalty += (h - nLimit) * 500;

            nurseMetrics.push({ id: n.id, h, nLimit, loadPct, counts });
        });

        // 9. Variância de carga relativa (loadPct): penaliza desequilíbrio proporcional
        //    Ex: Se Sirlene tem limite 100h e está com 95h (95%), e Sabina tem 135h com 130h (96%),
        //    elas estão equilibradas PROPORCIONALMENTE mesmo com horas absolutas diferentes
        const avgLoadPct = nurseMetrics.reduce((s, nm) => s + nm.loadPct, 0) / nurseMetrics.length;
        nurseMetrics.forEach(nm => {
            loadVariancePenalty += Math.abs(nm.loadPct - avgLoadPct) * 8000;
        });

        // 10. Equilíbrio de noturno entre enfermeiras (desvio da média)
        let nightBalancePenalty = 0;
        const nightCounts = nurseMetrics.map(nm => nm.counts['N']);
        const avgNight = nightCounts.reduce((s,v) => s+v, 0) / nightCounts.length;
        nightCounts.forEach(nc => { nightBalancePenalty += Math.abs(nc - avgNight) * 400; });

        // 11. Equilíbrio de P entre enfermeiras
        let pBalancePenalty = 0;
        const pCounts = nurseMetrics.map(nm => nm.counts['P']);
        const avgP = pCounts.reduce((s,v) => s+v, 0) / pCounts.length;
        pCounts.forEach(pc => { pBalancePenalty += Math.abs(pc - avgP) * 300; });

        // 12. Penalizar dias úteis com cobertura < 3 turnos diurnos
        let incompleteDays = 0;
        for (let dd = 1; dd <= days; dd++) {
            if (isFestivo(monthRefFit, dd)) continue;
            let filledCount = 0;
            NURSES.forEach(n => {
                const s = getSh(tSched, n.id, dd);
                if (s && !['OFF','FE','AT','N'].includes(s)) filledCount++;
            });
            if (filledCount < 3) incompleteDays++;
        }

        // Fitness composta ponderada
        let fitness = 200000
            - (emptyShifts * 25000)          // Buracos críticos
            - (incompleteDays * 35000)       // Dias incompletos
            - (repScore * 5000)              // Turnos repetidos consecutivos
            - (weekendPenalty)               // Desequilíbrio festivos
            - (postNightViolations * 8000)   // Violação descanso pós-noturno
            - (shiftTypePenalty)             // Desequilíbrio tipos de turno (G/M2, P/M1, MF/PF)
            - (overloadPenalty)              // Sobrecarga além do limite individual
            - (loadVariancePenalty)           // Variância proporcional de carga
            - (nightBalancePenalty)           // Desequilíbrio noturno entre enfermeiras
            - (pBalancePenalty);              // Desequilíbrio P entre enfermeiras

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
        if (bestSim.emptyShifts === 0 && bestSim.fitness > 195000) {
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
    // Mostra la data del richiedente (derivata dalla cella cliccata)
    const fromDateObj = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), selectedCell.day);
    const fromDateIso = `${fromDateObj.getFullYear()}-${String(fromDateObj.getMonth()+1).padStart(2,'0')}-${String(fromDateObj.getDate()).padStart(2,'0')}`;
    const fromDateLabel = fromDateObj.toLocaleDateString('it-IT', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
    document.getElementById('swapFromDateDisplay').value = fromDateLabel;
    document.getElementById('swapFromDateDisplay').dataset.iso = fromDateIso;
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
    const dateStr = document.getElementById('swapDateInput').value; // data della controparte
    if (!nurseId||!dateStr) { toast('Compila tutti i campi','warning'); return; }
    const dTo = new Date(dateStr+'T00:00:00');
    if (dTo.getMonth()!==currentMonth.getMonth()) { toast('Data fuori dal mese attuale','error'); return; }
    const toDay    = dTo.getDate();
    const toShift  = getShift(nurseId, toDay);
    const fromDay  = selectedCell.day;
    const fromShift= getShift(selectedCell.nurseId, fromDay);
    const toNurse  = NURSES.find(n=>n.id===nurseId);

    // Impedisce swap tra celle identiche (stessa data + stesso turno finirebbe in no-op)
    if (String(fromDay) === String(toDay) && fromShift === toShift) {
        toast('I turni sono uguali, nulla da scambiare','warning'); return;
    }
    // Validazione: la controparte deve avere effettivamente un turno assegnato
    if (!toShift || toShift === '') {
        toast('La controparte non ha un turno in questa data','error'); return;
    }

    // Costruisce le date ISO (YYYY-MM-DD) di entrambe le parti
    const fromDateObj = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), fromDay);
    const fromDateIso = `${fromDateObj.getFullYear()}-${String(fromDateObj.getMonth()+1).padStart(2,'0')}-${String(fromDateObj.getDate()).padStart(2,'0')}`;
    const toDateIso   = dateStr;

    const st = 'pending'; // Força o registro como pendente, até para Coordenadores.

    // Altera o calendário imediatamente apenas se for Coordenador (autoritário) — ora con swap incrociato 4 celle
    if (st === 'approved') {
        const reqCellOnFrom = getShift(selectedCell.nurseId, fromDay);   // = fromShift
        const cpCellOnFrom  = getShift(nurseId, fromDay);
        const reqCellOnTo   = getShift(selectedCell.nurseId, toDay);
        const cpCellOnTo    = getShift(nurseId, toDay);                  // = toShift
        assign(selectedCell.nurseId, fromDay, cpCellOnFrom);
        assign(nurseId,              fromDay, reqCellOnFrom);
        assign(selectedCell.nurseId, toDay,   cpCellOnTo);
        assign(nurseId,              toDay,   reqCellOnTo);
        renderCalendar();
    }

    requests.push({
        id: generateId(), type:'swap', status: st,
        fromNurseId: selectedCell.nurseId, fromNurseName: currentUser.name,
        fromDay, fromShift,
        toNurseId: nurseId, toNurseName: toNurse.name,
        nursecambio: toNurse.name, swapNurseName: toNurse.name,
        swapNurseId: nurseId,
        toDay, toShift,
        // Nuove colonne standard per il cloud (sheet Solicitacoes)
        dataRichiedente: fromDateIso,
        dataCambio: toDateIso,
        turnoRichiedente: fromShift,
        turnoCambio: toShift,
        // Retrocompat: `date` / `startDate` ora puntano alla data della controparte (comportamento pre-esistente)
        startDate: toDateIso, date: toDateIso,
        createdAt: new Date().toISOString(),
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

// ── EDIT REQUEST (solo richieste in attesa) ───────────────────
function openEditRequestModal(id) {
    const req = requests.find(r => String(r.id) === String(id));
    if (!req) { toast('Richiesta non trovata', 'error'); return; }
    if (req.status !== 'pending') {
        toast('Puoi modificare solo richieste ancora in attesa', 'warning');
        return;
    }
    // Verifica permessi: coordinatore o proprietario della richiesta
    const isOwner = (req.nurseId === currentUser.id) || (req.fromNurseId === currentUser.id);
    if (!isCoordinator && !isOwner) {
        toast('Non hai i permessi per modificare questa richiesta', 'warning');
        return;
    }

    editingRequestId = req.id;

    // Tipo (readonly)
    const typeLabels = { swap:'🔄 Cambio Turno', vacation:'🏖️ Ferie', justified:'📋 Riposo',
                         unexcused:'⚠️ Assenza Ingiustificata',
                         FE:'🏖️ Ferie Programmate', OFF:'📋 Riposo',
                         OFF_INJ:'⚠️ Assenza Ingiustificata', AT:'🏥 Certificato/Licenza' };
    document.getElementById('editReqTypeDisplay').value = typeLabels[req.type] || req.type;

    // Personale (readonly)
    const nurseName = req.nurseName || req.fromNurseName || '';
    document.getElementById('editReqNurseDisplay').value = nurseName;

    // Reset visibilità gruppi dinamici
    const swapGroup = document.getElementById('editReqSwapNurseGroup');
    const endGroup = document.getElementById('editReqEndDateGroup');
    const reasonGroup = document.getElementById('editReqReasonGroup');
    const startLabel = document.getElementById('editReqStartDateLabel');

    swapGroup.style.display = 'none';
    endGroup.style.display = 'none';
    reasonGroup.style.display = 'none';
    startLabel.textContent = 'Data';

    // Configurazione per tipo
    if (req.type === 'swap') {
        // Per swap: data e controparte
        const swapSel = document.getElementById('editReqSwapNurseSelect');
        // Tutte le infermiere eccetto chi richiede
        const requesterId = req.fromNurseId || req.nurseId;
        swapSel.innerHTML = NURSES.filter(n => n.id !== requesterId)
            .map(n => `<option value="${n.id}">${n.name}</option>`).join('');
        const currentSwapId = req.toNurseId || req.swapNurseId || req.nurseIdcambio || '';
        if (currentSwapId) swapSel.value = currentSwapId;
        swapGroup.style.display = 'block';

        const swapDate = req.date || req.startDate || '';
        document.getElementById('editReqStartDate').value = swapDate ? String(swapDate).split('T')[0] : '';
    } else if (req.type === 'vacation' || req.type === 'FE') {
        // Ferie: intervallo date
        startLabel.textContent = 'Data Inizio';
        endGroup.style.display = 'block';
        document.getElementById('editReqStartDate').value = req.startDate ? String(req.startDate).split('T')[0] : '';
        document.getElementById('editReqEndDate').value = req.endDate ? String(req.endDate).split('T')[0] : '';
    } else if (req.type === 'AT') {
        // Certificato: intervallo date + motivo
        startLabel.textContent = 'Data Inizio';
        endGroup.style.display = 'block';
        reasonGroup.style.display = 'block';
        document.getElementById('editReqStartDate').value = req.startDate ? String(req.startDate).split('T')[0] : '';
        document.getElementById('editReqEndDate').value = req.endDate ? String(req.endDate).split('T')[0] : '';
        document.getElementById('editReqReason').value = req.reason || req.desc || '';
    } else {
        // justified / unexcused / OFF / OFF_INJ: data singola + motivo
        reasonGroup.style.display = 'block';
        const dateVal = req.date || req.startDate || '';
        document.getElementById('editReqStartDate').value = dateVal ? String(dateVal).split('T')[0] : '';
        document.getElementById('editReqReason').value = req.reason || req.desc || '';
    }

    document.getElementById('editRequestModal').classList.remove('hidden');
}

function closeEditRequestModal(e) {
    if (e && e.target && !e.target.classList.contains('modal-bd')) return;
    document.getElementById('editRequestModal').classList.add('hidden');
    editingRequestId = null;
}

function submitEditRequest() {
    if (!editingRequestId) { toast('Nessuna richiesta in modifica', 'error'); return; }
    const req = requests.find(r => String(r.id) === String(editingRequestId));
    if (!req) { toast('Richiesta non trovata', 'error'); editingRequestId = null; return; }
    if (req.status !== 'pending') {
        toast('Puoi modificare solo richieste ancora in attesa', 'warning');
        editingRequestId = null;
        return;
    }

    const startDate = document.getElementById('editReqStartDate').value;
    const endDate = document.getElementById('editReqEndDate').value;
    const reason = document.getElementById('editReqReason').value.trim();
    const swapNurseId = document.getElementById('editReqSwapNurseSelect').value;

    if (!startDate) { toast('Compila la data', 'warning'); return; }

    if (req.type === 'swap') {
        if (!swapNurseId) { toast('Seleziona la persona per il cambio', 'warning'); return; }
        const swapNurse = NURSES.find(n => n.id === swapNurseId);
        if (!swapNurse) { toast('Infermiera di cambio non trovata', 'error'); return; }
        const d = new Date(startDate + 'T00:00:00');
        req.date = startDate;
        req.startDate = startDate;
        req.toDay = d.getDate();
        req.toNurseId = swapNurseId;
        req.toNurseName = swapNurse.name;
        req.swapNurseId = swapNurseId;
        req.swapNurseName = swapNurse.name;
        req.nursecambio = swapNurse.name;
        req.nurseIdcambio = swapNurseId;
        req.toShift = getShift(swapNurseId, d.getDate());
    } else if (req.type === 'vacation' || req.type === 'FE') {
        if (!endDate) { toast('Compila la data di fine', 'warning'); return; }
        if (new Date(endDate) < new Date(startDate)) {
            toast('La data di fine deve essere successiva a quella di inizio', 'error');
            return;
        }
        req.startDate = startDate;
        req.endDate = endDate;
    } else if (req.type === 'AT') {
        if (!endDate) { toast('Compila la data di fine', 'warning'); return; }
        if (new Date(endDate) < new Date(startDate)) {
            toast('La data di fine deve essere successiva a quella di inizio', 'error');
            return;
        }
        req.startDate = startDate;
        req.endDate = endDate;
        req.reason = reason;
        req.desc = reason;
    } else {
        // justified / unexcused / OFF / OFF_INJ
        const d = new Date(startDate + 'T00:00:00');
        req.date = startDate;
        req.startDate = startDate;
        req.day = d.getDate();
        if (reason) req.reason = reason;
        if (reason) req.desc = reason;
    }

    // Marca come modificata (timestamp opzionale)
    req.modifiedAt = new Date().toISOString();

    saveData();
    renderRequests();
    updateBadge();
    document.getElementById('editRequestModal').classList.add('hidden');
    editingRequestId = null;

    // Sincronizzazione cloud
    if (typeof syncAllRequestsToCloud === 'function') {
        syncAllRequestsToCloud();
    }

    toast('✏️ Richiesta modificata', 'success');
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
        // Solo le richieste ancora in attesa possono essere eliminate o modificate.
        // Una volta approvate o rifiutate restano nello storico come log immutabile.
        const canDelete = req.status === 'pending' && (isCoordinator || (req.nurseId===currentUser.id||req.fromNurseId===currentUser.id));
        const canEdit = canDelete; // stessi permessi dell'eliminazione
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
                    ${canEdit ? `<button class="req-action-btn" style="background:rgba(59,130,246,0.1); color:#60a5fa; border:1px solid rgba(59,130,246,0.25);" onclick="openEditRequestModal('${req.id}')">✏️ Modifica</button>` : ''}
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
            nursecambio: req.nursecambio || req.toNurseName || req.swapNurseName || '',
            nurseIdcambio: req.nurseIdcambio || req.swapNurseId || req.toNurseId || '',
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
    // Resolve nome do solicitante (se vier só o ID, busca em NURSES)
    let solicitanteName = req.nurseName || req.fromNurseName || '';
    if (!solicitanteName) {
        const sId = String(req.nurseId || req.fromNurseId || '').trim();
        if (sId) {
            const nFound = NURSES.find(n => String(n.id).trim() === sId);
            if (nFound) solicitanteName = nFound.name;
        }
    }
    solicitanteName = solicitanteName || '—';

    let h = `<div class="req-detail-row"><span class="req-detail-icon">👤</span><strong>${solicitanteName}</strong></div>`;
    if (req.type==='swap') {
        // Nome da enfermeira de câmbio: prioriza nursecambio (coluna do Sheets), depois fallbacks internos,
        // e por fim tenta resolver pelo nurseIdcambio/swapNurseId via NURSES (quando o Mobile gravou só o ID).
        let cambioName = req.nursecambio || req.toNurseName || req.swapNurseName || '';
        if (!cambioName) {
            const cambioId = String(req.nurseIdcambio || req.swapNurseId || req.toNurseId || '').trim();
            if (cambioId) {
                const nFound = NURSES.find(n => String(n.id).trim() === cambioId);
                if (nFound) cambioName = nFound.name;
            }
        }
        cambioName = cambioName || '—';
        // Turni: prioriza snapshots turnoRichiedente/turnoCambio, fallback para fromShift/toShift (legado)
        // OFF/empty = "Riposo" (cobertura unidirectional)
        const fromShiftCode = (req.turnoRichiedente || req.fromShift || '').toString().toUpperCase() || 'OFF';
        const toShiftCode   = (req.turnoCambio      || req.toShift   || '').toString().toUpperCase() || 'OFF';
        const fromShiftName = fromShiftCode === 'OFF' ? '🛌 Riposo' : (SHIFTS[fromShiftCode]?.name || fromShiftCode);
        const toShiftName   = toShiftCode   === 'OFF' ? '🛌 Riposo' : (SHIFTS[toShiftCode]?.name   || toShiftCode);
        // Datas: prioriza dataRichiedente/dataCambio (cross-date); fallback para startDate/date
        const fromDateRaw = sanitizeDate(req.dataRichiedente || req.startDate || req.date || '');
        const toDateRaw   = sanitizeDate(req.dataCambio      || req.startDate || req.date || '');
        const fromDateDisplay = fromDateRaw ? fromDateRaw.split('-').reverse().join('/') : '';
        const toDateDisplay   = toDateRaw   ? toDateRaw.split('-').reverse().join('/')   : '';
        const isCrossDate = fromDateRaw && toDateRaw && fromDateRaw !== toDateRaw;

        // Card de troca: visual em 2 colunas mostrando claramente quem cede o quê
        let swapDetail = '';
        if (isCrossDate) {
            swapDetail = `
              <div class="req-swap-summary">
                <div class="req-swap-grid">
                  <div class="req-swap-side">
                    <div class="swap-side-label">${solicitanteName}</div>
                    <div class="swap-side-info">📅 ${fromDateDisplay}</div>
                    <div class="swap-side-info">⏰ ${fromShiftName || '—'}</div>
                  </div>
                  <div class="req-swap-arrow" title="Scambio">⇄</div>
                  <div class="req-swap-side">
                    <div class="swap-side-label">${cambioName}</div>
                    <div class="swap-side-info">📅 ${toDateDisplay}</div>
                    <div class="swap-side-info">⏰ ${toShiftName || '—'}</div>
                  </div>
                </div>
                <div class="req-swap-footnote">Dopo lo scambio: ${solicitanteName} farà ${toDateDisplay} (${toShiftName||'—'}); ${cambioName} farà ${fromDateDisplay} (${fromShiftName||'—'}).</div>
              </div>`;
        } else {
            // Single-date (legado): mesmo dia, troca turnos
            const dateLabel = fromDateDisplay || toDateDisplay || '';
            swapDetail = `
              <div class="req-swap-summary">
                ${dateLabel ? `<div class="req-swap-header"><span>📅</span><strong>${dateLabel}</strong></div>` : ''}
                <div class="req-swap-grid">
                  <div class="req-swap-side">
                    <div class="swap-side-label">${solicitanteName}</div>
                    <div class="swap-side-info">⏰ ${fromShiftName || '—'}</div>
                  </div>
                  <div class="req-swap-arrow" title="Scambio">⇄</div>
                  <div class="req-swap-side">
                    <div class="swap-side-label">${cambioName}</div>
                    <div class="swap-side-info">⏰ ${toShiftName || '—'}</div>
                  </div>
                </div>
              </div>`;
        }
        h += swapDetail;
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
            nursecambio: r.nursecambio || r.toNurseName || r.swapNurseName || '',
            nurseIdcambio: r.nurseIdcambio || r.swapNurseId || r.toNurseId || '',
            swapNurseId: r.swapNurseId || r.toNurseId || '',
            startDate: r.startDate || r.date || '',
            endDate: r.endDate || r.startDate || r.date || '',
            desc: r.desc || r.reason || '',
            // Nuove colonne per swap cross-date (vuote per tipi diversi da swap)
            dataRichiedente: r.dataRichiedente || '',
            dataCambio: r.dataCambio || '',
            turnoRichiedente: r.turnoRichiedente || '',
            turnoCambio: r.turnoCambio || '',
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
    // Difesa in profondità: non si può eliminare una richiesta già approvata o rifiutata.
    // Lo storico delle decisioni deve restare tracciabile.
    const target = requests.find(r => String(r.id) === String(id));
    if (!target) { toast('Richiesta non trovata', 'error'); return; }
    if (target.status !== 'pending') {
        toast('Non puoi eliminare una richiesta già approvata o rifiutata', 'warning');
        return;
    }
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
        // Resolve IDs aceitando ambos os formatos (Mobile e Local legado)
        const nurseAId = String(req.nurseId || req.fromNurseId || '').trim();
        const nurseBId = String(req.nurseIdcambio || req.swapNurseId || req.toNurseId || '').trim();

        // Datas: dataRichiedente = dia cedido pelo solicitante; dataCambio = dia cedido pela contraparte.
        // Retrocompatibilidade: swaps antigos só têm startDate (uma única data) — same-day swap.
        const fromDateStr = sanitizeDate(req.dataRichiedente || req.startDate || req.date || '');
        const toDateStr   = sanitizeDate(req.dataCambio      || req.startDate || req.date || '');

        if (nurseAId && nurseBId && fromDateStr && toDateStr) {
            const fromDate = new Date(fromDateStr + 'T00:00:00');
            const toDate   = new Date(toDateStr   + 'T00:00:00');
            const swapMonth = fromDate.getMonth();
            const swapYear  = fromDate.getFullYear();
            const fromDay   = fromDate.getDate();
            const toDay     = toDate.getDate();

            // Sanity: precisa estar no mesmo mês/ano (limitação documentada do solver)
            if (toDate.getMonth() !== swapMonth || toDate.getFullYear() !== swapYear) {
                console.warn('[APPROVE SWAP] cross-month não suportado:', { fromDateStr, toDateStr });
                toast('⚠️ Cambio fra mesi diversi non supportato.', 'warning');
            } else {
                const k = (nId, d) => `${nId}_${swapMonth}_${swapYear}_${d}`;
                // Snapshots vindos do request (preferencial). Sem snapshot, lê do schedule atual.
                const snapFrom = String(req.turnoRichiedente || '').trim();
                const snapTo   = String(req.turnoCambio      || '').trim();
                const reqOnFrom = schedule[k(nurseAId, fromDay)] || 'OFF';
                const cpOnFrom  = schedule[k(nurseBId, fromDay)] || 'OFF';
                const reqOnTo   = schedule[k(nurseAId, toDay)]   || 'OFF';
                const cpOnTo    = schedule[k(nurseBId, toDay)]   || 'OFF';

                if (fromDay === toDay) {
                    // ── SAME-DAY SWAP (2 células) ──
                    const shiftA = (req.fromShift && req.fromNurseId === nurseAId) ? req.fromShift : reqOnFrom;
                    const shiftB = (req.toShift   && req.toNurseId   === nurseBId) ? req.toShift   : cpOnFrom;
                    if (shiftA !== shiftB) {
                        schedule[k(nurseAId, fromDay)] = shiftB;
                        schedule[k(nurseBId, fromDay)] = shiftA;
                        console.log(`[APPROVE SWAP same-day] ${fromDay}/${swapMonth+1}/${swapYear}: ${nurseAId}(${shiftA}) <-> ${nurseBId}(${shiftB})`);
                    }
                } else {
                    // ── CROSS-DATE 4-CELL SYMMETRIC SWAP ──
                    // Idempotência: se o estado já bate com o pós-swap esperado, pula
                    const alreadyApplied =
                        snapFrom && snapTo &&
                        cpOnFrom === snapFrom && reqOnTo === snapTo;
                    if (alreadyApplied) {
                        console.log('[APPROVE SWAP cross] Já aplicado (snapshot pós-swap detectado), nada a fazer');
                    } else {
                        // Sanity check (apenas log) quando snapshots existem
                        if (snapFrom && snapTo && (reqOnFrom !== snapFrom || cpOnTo !== snapTo)) {
                            console.warn('[APPROVE SWAP cross] estado divergente do snapshot — aplicando com valores correntes', {
                                snapshot: { reqOnFrom: snapFrom, cpOnTo: snapTo },
                                atual:    { reqOnFrom, cpOnTo }
                            });
                        }
                        schedule[k(nurseAId, fromDay)] = cpOnFrom;
                        schedule[k(nurseBId, fromDay)] = reqOnFrom;
                        schedule[k(nurseAId, toDay)]   = cpOnTo;
                        schedule[k(nurseBId, toDay)]   = reqOnTo;
                        req.autoApplied = true;
                        console.log(`[APPROVE SWAP cross] ${fromDay}↔${toDay}/${swapMonth+1}/${swapYear} aplicado`);
                    }
                }

                // Guarda o mês alvo para republicar no cloud após o salvamento
                req._swapAppliedMonth = swapMonth;
                req._swapAppliedYear = swapYear;
            }
        } else {
            console.warn('[APPROVE SWAP] Dados insuficientes — nurseAId:', nurseAId, 'nurseBId:', nurseBId, 'fromDate:', fromDateStr, 'toDate:', toDateStr);
            toast('⚠️ Richiesta di cambio incompleta (manca ID o data).', 'warning');
        }
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

    // Auto-sync: publica o status da request na nuvem
    syncRequestToCloud(req);

    // Se for uma troca aprovada, publica a escala do mês da troca no cloud
    // para que o Mobile também veja a alteração (refresh automático)
    if (req.type === 'swap' && req._swapAppliedMonth != null && req._swapAppliedYear != null) {
        const m = req._swapAppliedMonth;
        const y = req._swapAppliedYear;
        const days = new Date(y, m + 1, 0).getDate();
        autoPublishMonth(m, y, days);
        delete req._swapAppliedMonth;
        delete req._swapAppliedYear;
    }
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

function setReportNurseFilter(nurseId) {
    reportNurseFilter = nurseId;
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
    const dayNamesShort = ['D','L','M','M','G','V','S'];
    const dayNamesFull = ['Domenica','Lunedì','Martedì','Mercoledì','Giovedì','Venerdì','Sabato'];
    const monthLabel = `${months[m]} ${y}`;

    // All nurses for data collection (always compute full team)
    const monthKey = `${m}_${y}`;
    const orderForMonth = monthlyOrder[monthKey] || NURSES.map(n => n.id);
    const allNurses = isCoordinator
        ? orderForMonth.map(id => NURSES.find(n => n.id === id)).filter(Boolean)
        : [NURSES.find(n => n.id === currentUser.id)].filter(Boolean);

    if (allNurses.length === 0) {
        document.getElementById('reportsTab').innerHTML = `<div class="empty-state" style="padding:80px"><div class="empty-icon">📊</div><p>Nessun personale attivo questo mese.</p></div>`;
        return;
    }

    // ── COLETA DE DADOS AGREGADOS (always full team) ──
    let teamTotalH = 0, teamWorkDays = 0, teamRestDays = 0, teamNightShifts = 0;
    let teamAbsences = 0, teamVacations = 0;
    const nurseData = [];

    const monthRef = new Date(y, m, 1);
    let giorniFeriali = 0;
    for (let d = 1; d <= daysInMo; d++) { if (!isFestivo(monthRef, d)) giorniFeriali++; }
    const oreDovute = giorniFeriali * 7.5;

    // Copertura giornaliera
    const dailyCoverage = [];
    for (let d = 1; d <= daysInMo; d++) {
        const fest = isFestivo(monthRef, d);
        let count = 0;
        const dayShifts = {};
        allNurses.forEach(n => {
            const code = getShiftForMonth(n.id, d, m, y);
            if (code && !['OFF','FE','AT'].includes(code)) {
                count++;
                dayShifts[code] = (dayShifts[code] || 0) + 1;
            }
        });
        const expected = fest ? 3 : 4;
        dailyCoverage.push({ day: d, count, expected, fest, shifts: dayShifts });
    }

    allNurses.forEach(n => {
        let totalH = 0, workDays = 0, restDays = 0, nightCount = 0;
        let feCount = 0, atCount = 0, offCount = 0;
        const shiftCounts = {};
        let weekendsWorked = 0;
        const weekendSet = new Set();
        let oreDiurneFeriali = 0, oreDiurneFestive = 0;
        let oreNotturneFeriali = 0, oreNotturneFestive = 0;
        let oreFerie = 0, oreMalattia = 0;
        // Track consecutive work streaks and post-night violations
        let maxConsecutiveWork = 0, currentStreak = 0;
        let postNightViolations = 0, prevWasNight = false;

        for (let d = 1; d <= daysInMo; d++) {
            const code = getShiftForMonth(n.id, d, m, y);
            const sh = SHIFTS[code];
            if (!sh) continue;

            totalH += sh.h;
            shiftCounts[code] = (shiftCounts[code] || 0) + 1;

            if (['OFF'].includes(code)) { offCount++; restDays++; currentStreak = 0; }
            else if (code === 'FE') { feCount++; restDays++; oreFerie += 7.5; currentStreak = 0; }
            else if (code === 'AT') { atCount++; restDays++; oreMalattia += 7.5; currentStreak = 0; }
            else { workDays++; currentStreak++; if (currentStreak > maxConsecutiveWork) maxConsecutiveWork = currentStreak; }
            if (code === 'N') nightCount++;

            // Post-night violation
            if (prevWasNight && !['OFF','FE','AT'].includes(code) && code !== 'N') postNightViolations++;
            prevWasNight = (code === 'N');

            const fest = isFestivo(monthRef, d);
            if (code === 'N') {
                if (fest) oreNotturneFestive += sh.h; else oreNotturneFeriali += sh.h;
            } else if (!['OFF','FE','AT'].includes(code)) {
                if (fest) oreDiurneFestive += sh.h; else oreDiurneFeriali += sh.h;
            }
            if (fest && !['OFF','FE','AT'].includes(code)) weekendSet.add(d);
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
            oreFerie, oreMalattia, oreDovute, differenza: totalH - oreDovute,
            maxConsecutiveWork, postNightViolations
        });
    });

    // Requests do mês
    const monthRequests = requests.filter(r => {
        const rDate = sanitizeDate(r.startDate || r.date || r.createdAt?.split('T')[0] || '');
        if (!rDate) return false;
        const rd = new Date(rDate + 'T00:00:00');
        return rd.getMonth() === m && rd.getFullYear() === y;
    });
    const pendingReqs = monthRequests.filter(r => r.status === 'pending').length;
    const approvedReqs = monthRequests.filter(r => r.status === 'approved').length;
    const rejectedReqs = monthRequests.filter(r => r.status === 'rejected').length;

    const totalPossibleDays = allNurses.length * daysInMo;
    const absentDays = nurseData.reduce((s, nd) => s + nd.atCount + nd.feCount, 0);
    const absenteeismRate = totalPossibleDays > 0 ? ((absentDays / totalPossibleDays) * 100).toFixed(1) : '0.0';
    const avgHours = allNurses.length > 0 ? (teamTotalH / allNurses.length).toFixed(1) : '0.0';
    const maxH = Math.max(...nurseData.map(nd => nd.totalH));
    const minH = Math.min(...nurseData.map(nd => nd.totalH));
    const hourSpread = (maxH - minH).toFixed(1);

    // Coverage issues
    const coverageIssues = dailyCoverage.filter(dc => dc.count < dc.expected).length;
    const coverageRate = daysInMo > 0 ? (((daysInMo - coverageIssues) / daysInMo) * 100).toFixed(0) : '100';

    // ── RENDER HTML ──
    const tab = document.getElementById('reportsTab');

    // ── HEADER ──
    let html = `<div class="reports-wrap" style="max-width:1200px;">
        <div class="rpt-header">
            <h2>📊 Rapporto Operativo</h2>
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
        </div>`;

    // Individual view placeholder removed — reserved for mobile
    if (false) {
        const nd = selectedNd;
        const nightQuota = nd.nurse.nightQuota || 5;
        const nightPct = Math.min((nd.nightCount / nightQuota) * 100, 100).toFixed(0);
        const nightColor = nd.nightCount > nightQuota ? 'var(--danger)' : nd.nightCount === nightQuota ? 'var(--warning)' : 'var(--primary)';
        const diff = nd.differenza;
        const diffColor = diff > 0 ? 'var(--success)' : diff < 0 ? 'var(--danger)' : 'var(--text-2)';
        const diffSign = diff > 0 ? '+' : '';

        // Rank position
        const sortedByH = [...nurseData].sort((a,b) => b.totalH - a.totalH);
        const rank = sortedByH.findIndex(x => x.nurse.id === nd.nurse.id) + 1;

        // Compare to team average
        const avgH = parseFloat(avgHours);
        const vsAvg = nd.totalH - avgH;
        const vsAvgSign = vsAvg > 0 ? '+' : '';
        const vsAvgColor = Math.abs(vsAvg) < 5 ? 'var(--success)' : vsAvg > 0 ? 'var(--warning)' : 'var(--danger)';

        html += `
        <!-- Individual Header Card -->
        <div style="background:linear-gradient(135deg, rgba(139,92,246,0.12), rgba(59,130,246,0.08)); border:1px solid rgba(139,92,246,0.25); border-radius:16px; padding:24px; margin-bottom:20px;">
            <div style="display:flex; align-items:center; gap:16px; margin-bottom:16px;">
                <div style="width:56px; height:56px; border-radius:14px; background:var(--primary); display:flex; align-items:center; justify-content:center; font-size:20px; font-weight:900; color:white;">${nd.nurse.initials}</div>
                <div>
                    <h3 style="margin:0; font-size:20px; font-weight:800; color:var(--text);">${nd.nurse.name}</h3>
                    <p style="margin:2px 0 0; font-size:13px; color:var(--text-3);">Scheda Individuale — ${monthLabel}</p>
                </div>
                <div style="margin-left:auto; text-align:right;">
                    <div style="font-size:32px; font-weight:900; color:var(--primary-light);">#${rank}</div>
                    <div style="font-size:11px; color:var(--text-3);">su ${allNurses.length} infermiere</div>
                </div>
            </div>

            <!-- Executive summary text -->
            <div style="background:rgba(0,0,0,0.15); border-radius:10px; padding:14px 16px; font-size:13px; line-height:1.6; color:var(--text-2);">
                Nel mese di <strong>${monthLabel}</strong>, <strong>${nd.nurse.name}</strong> ha lavorato <strong>${nd.workDays} giorni</strong> per un totale di <strong>${nd.totalH.toFixed(1)} ore</strong>
                (${vsAvgSign}${vsAvg.toFixed(1)}h rispetto alla media team di ${avgH}h).
                ${nd.nightCount > 0 ? `Ha svolto <strong>${nd.nightCount} turni notturni</strong> su ${nightQuota} previsti.` : 'Nessun turno notturno svolto.'}
                ${nd.weekendsWorked > 0 ? `Ha lavorato <strong>${nd.weekendsWorked} giorni festivi</strong>.` : ''}
                ${nd.feCount > 0 ? `Ferie godute: <strong>${nd.feCount} giorni</strong>.` : ''}
                ${nd.atCount > 0 ? `Certificati/licenze: <strong>${nd.atCount} giorni</strong>.` : ''}
                ${nd.postNightViolations > 0 ? `<span style="color:var(--danger);">⚠️ ${nd.postNightViolations} violazioni post-notte rilevate.</span>` : ''}
            </div>
        </div>

        <!-- Individual KPI Cards -->
        <div class="stats-grid" style="grid-template-columns: repeat(auto-fit, minmax(140px,1fr));">
            <div class="stat-card stat-primary">
                <div class="stat-lbl">Ore Lavorate</div>
                <div class="stat-val" style="font-size:28px;">${nd.totalH.toFixed(1)}h</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;color:${vsAvgColor}">${vsAvgSign}${vsAvg.toFixed(1)}h vs media</div>
            </div>
            <div class="stat-card stat-success">
                <div class="stat-lbl">Giorni Lavorati</div>
                <div class="stat-val" style="font-size:28px;">${nd.workDays}</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">Riposo: ${nd.restDays}</div>
            </div>
            <div class="stat-card stat-night">
                <div class="stat-lbl">Turni Notturni</div>
                <div class="stat-val" style="font-size:28px;">${nd.nightCount}/${nightQuota}</div>
            </div>
            <div class="stat-card stat-warning">
                <div class="stat-lbl">Festivi Lavorati</div>
                <div class="stat-val" style="font-size:28px;">${nd.weekendsWorked}</div>
            </div>
            <div class="stat-card" style="background:rgba(16,185,129,0.08); border-color:rgba(16,185,129,0.2);">
                <div class="stat-lbl">Bilancio Ore</div>
                <div class="stat-val" style="font-size:28px; color:${diffColor};">${diffSign}${diff.toFixed(1)}h</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">Dovute: ${oreDovute.toFixed(1)}h</div>
            </div>
            <div class="stat-card" style="background:rgba(239,68,68,0.08); border-color:rgba(239,68,68,0.2);">
                <div class="stat-lbl">Consecutivi Max</div>
                <div class="stat-val" style="font-size:28px; color:${nd.maxConsecutiveWork > 6 ? 'var(--danger)' : 'var(--text)'};">${nd.maxConsecutiveWork}</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">giorni di fila</div>
            </div>
        </div>

        <!-- Distribuzione Turni Individuale (visual chips) -->
        <div class="report-section">
            <h3>🔄 Distribuzione Turni</h3>
            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(90px,1fr)); gap:8px;">
            ${['M1','M2','MF','G','P','PF','N','OFF','FE','AT'].map(c => {
                const cnt = nd.shiftCounts[c] || 0;
                const s = SHIFTS[c];
                const opacity = cnt > 0 ? '1' : '0.3';
                const hrs = (cnt * s.h).toFixed(1);
                return `<div style="background:${s.color}; opacity:${opacity}; border-radius:12px; padding:12px 8px; text-align:center;">
                    <div style="font-size:12px; font-weight:800; color:${s.text};">${c}</div>
                    <div style="font-size:24px; font-weight:900; color:${s.text}; line-height:1.2;">${cnt}</div>
                    <div style="font-size:10px; color:${s.text}; opacity:0.8;">${hrs}h</div>
                </div>`;
            }).join('')}
            </div>
        </div>

        <!-- Ripartizione ore individuale -->
        <div class="report-section">
            <h3>📊 Ripartizione Ore</h3>
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px;">
                <div style="background:rgba(14,165,233,0.08); border:1px solid rgba(14,165,233,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:10px; color:var(--text-3); text-transform:uppercase; font-weight:700;">Diurne Feriali</div>
                    <div style="font-size:22px; font-weight:900; color:#38bdf8; margin:4px 0;">${nd.oreDiurneFeriali.toFixed(1)}h</div>
                </div>
                <div style="background:rgba(245,158,11,0.08); border:1px solid rgba(245,158,11,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:10px; color:var(--text-3); text-transform:uppercase; font-weight:700;">Diurne Festive</div>
                    <div style="font-size:22px; font-weight:900; color:#fbbf24; margin:4px 0;">${nd.oreDiurneFestive.toFixed(1)}h</div>
                </div>
                <div style="background:rgba(30,27,75,0.3); border:1px solid rgba(139,92,246,0.3); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:10px; color:var(--text-3); text-transform:uppercase; font-weight:700;">Notturne Feriali</div>
                    <div style="font-size:22px; font-weight:900; color:#a78bfa; margin:4px 0;">${nd.oreNotturneFeriali.toFixed(1)}h</div>
                </div>
                <div style="background:rgba(239,68,68,0.08); border:1px solid rgba(239,68,68,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:10px; color:var(--text-3); text-transform:uppercase; font-weight:700;">Notturne Festive</div>
                    <div style="font-size:22px; font-weight:900; color:#f87171; margin:4px 0;">${nd.oreNotturneFestive.toFixed(1)}h</div>
                </div>
            </div>
        </div>

        <!-- Quota notturni bar -->
        <div class="report-section">
            <h3>🌙 Quota Turni Notturni</h3>
            <div style="background:rgba(255,255,255,0.03); border:1px solid var(--border); border-radius:12px; padding:16px;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <span style="font-size:13px; font-weight:700; color:var(--text-2);">Progresso</span>
                    <span style="font-size:15px; font-weight:900; color:${nightColor};">${nd.nightCount} / ${nightQuota}</span>
                </div>
                <div style="height:12px; background:rgba(255,255,255,0.06); border-radius:99px; overflow:hidden;">
                    <div style="width:${nightPct}%; height:100%; background:${nightColor}; border-radius:99px; transition:width 0.6s;"></div>
                </div>
                <div style="font-size:11px; color:var(--text-3); margin-top:6px;">${nightPct}% della quota mensile completata</div>
            </div>
        </div>

        <!-- Mappa giornaliera individuale -->
        <div class="report-section">
            <h3>📅 Mappa Turni del Mese</h3>
            <div style="display:grid; grid-template-columns:repeat(7, 1fr); gap:4px;">
            ${(() => {
                const firstDow = new Date(y, m, 1).getDay();
                const emptyCells = (firstDow + 6) % 7;
                let cells = '';
                // Weekday headers
                ['L','M','M','G','V','S','D'].forEach(dn => {
                    cells += `<div style="text-align:center; font-size:10px; font-weight:700; color:var(--text-3); padding:4px 0;">${dn}</div>`;
                });
                // Empty cells
                for (let i = 0; i < emptyCells; i++) cells += '<div></div>';
                // Days
                for (let d = 1; d <= daysInMo; d++) {
                    const code = getShiftForMonth(nd.nurse.id, d, m, y);
                    const s = SHIFTS[code] || SHIFTS['OFF'];
                    const fest = isFestivo(monthRef, d);
                    cells += `<div style="
                        background:${s.color}; border-radius:6px; padding:4px 2px; text-align:center;
                        border:${fest ? '2px solid rgba(245,158,11,0.5)' : '1px solid rgba(255,255,255,0.05)'};
                    " title="Giorno ${d}: ${code} (${s.h}h)${fest ? ' — Festivo' : ''}">
                        <div style="font-size:9px; color:${s.text}; opacity:0.7;">${d}</div>
                        <div style="font-size:11px; font-weight:800; color:${s.text};">${code}</div>
                    </div>`;
                }
                return cells;
            })()}
            </div>
        </div>`;

    } else {
    // Executive summary narrative
    const coverStatus = coverageIssues === 0 ? '✅ completa' : `⚠️ ${coverageIssues} giorni scoperti`;
    const balanceStatus = parseFloat(hourSpread) < 10 ? '✅ equilibrata' : '⚠️ da riequilibrare';
    const topNurse = nurseData.reduce((a,b) => a.totalH > b.totalH ? a : b);
    const bottomNurse = nurseData.reduce((a,b) => a.totalH < b.totalH ? a : b);

    html += `
        <!-- Executive Summary -->
        <div style="background:linear-gradient(135deg, rgba(139,92,246,0.08), rgba(59,130,246,0.05)); border:1px solid rgba(139,92,246,0.15); border-radius:14px; padding:20px; margin-bottom:20px;">
            <h3 style="margin:0 0 10px; font-size:15px; font-weight:800; color:var(--text);">📝 Sintesi Esecutiva</h3>
            <p style="font-size:13px; line-height:1.7; color:var(--text-2); margin:0;">
                Nel mese di <strong>${monthLabel}</strong>, il team di <strong>${allNurses.length} infermiere</strong> ha totalizzato
                <strong>${teamTotalH.toFixed(1)} ore</strong> di servizio (media ${avgHours}h/persona) con <strong>${teamNightShifts} turni notturni</strong>.
                La copertura giornaliera è ${coverStatus}. La distribuzione del carico è ${balanceStatus}
                (dispersione ${hourSpread}h). L'infermiera con più ore è <strong>${topNurse.nurse.name}</strong> (${topNurse.totalH.toFixed(1)}h)
                e quella con meno è <strong>${bottomNurse.nurse.name}</strong> (${bottomNurse.totalH.toFixed(1)}h).
                ${teamAbsences > 0 ? `Assenze per certificato: <strong>${teamAbsences} giorni</strong>.` : ''}
                ${teamVacations > 0 ? `Ferie godute: <strong>${teamVacations} giorni</strong>.` : ''}
                Tasso di assenteismo: <strong>${absenteeismRate}%</strong>.
                ${pendingReqs > 0 ? `<span style="color:var(--warning);">⏳ ${pendingReqs} richieste in attesa di approvazione.</span>` : ''}
            </p>
        </div>

        <!-- KPI Cards -->
        <div class="stats-grid" style="grid-template-columns: repeat(auto-fit, minmax(150px,1fr));">
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
                <div class="stat-lbl">Assenteismo</div>
                <div class="stat-val" style="font-size:28px;">${absenteeismRate}%</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">${absentDays} giorni</div>
            </div>
            <div class="stat-card" style="background:rgba(14,165,233,0.08); border-color:rgba(14,165,233,0.2);">
                <div class="stat-lbl">Copertura</div>
                <div class="stat-val" style="font-size:28px; color:${coverageIssues === 0 ? 'var(--success)' : 'var(--danger)'};">${coverageRate}%</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">${coverageIssues > 0 ? coverageIssues + ' scoperte' : 'Completa'}</div>
            </div>
            <div class="stat-card" style="background:rgba(139,92,246,0.08); border-color:rgba(139,92,246,0.2);">
                <div class="stat-lbl">Equilibrio</div>
                <div class="stat-val" style="font-size:28px; color:${parseFloat(hourSpread) < 10 ? 'var(--success)' : 'var(--warning)'};">${hourSpread}h</div>
                <div style="font-size:11px; opacity:0.7; margin-top:4px;">dispersione max-min</div>
            </div>
        </div>

        <!-- Indicadores de RH -->
        <div class="report-section">
            <h3>📋 Indicatori Risorse Umane</h3>
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(180px,1fr)); gap:12px;">
                <div style="background:rgba(245,158,11,0.08); border:1px solid rgba(245,158,11,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--warning);">${pendingReqs}</div>
                    <div style="font-size:11px; color:var(--text-3); margin-top:4px;">Richieste In Attesa</div>
                </div>
                <div style="background:rgba(16,185,129,0.08); border:1px solid rgba(16,185,129,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--success);">${approvedReqs}</div>
                    <div style="font-size:11px; color:var(--text-3); margin-top:4px;">Richieste Approvate</div>
                </div>
                <div style="background:rgba(239,68,68,0.08); border:1px solid rgba(239,68,68,0.2); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--danger);">${rejectedReqs}</div>
                    <div style="font-size:11px; color:var(--text-3); margin-top:4px;">Richieste Rifiutate</div>
                </div>
                <div style="background:rgba(16,185,129,0.05); border:1px solid rgba(16,185,129,0.15); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--success);">${teamVacations}</div>
                    <div style="font-size:11px; color:var(--text-3); margin-top:4px;">Giorni Ferie (FE)</div>
                </div>
                <div style="background:rgba(239,68,68,0.05); border:1px solid rgba(239,68,68,0.15); border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:24px; font-weight:900; color:var(--danger);">${teamAbsences}</div>
                    <div style="font-size:11px; color:var(--text-3); margin-top:4px;">Certificati (AT)</div>
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
                ${[...nurseData].sort((a,b) => b.totalH - a.totalH).map((nd, idx) => {
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
                const dow = new Date(y, m, dc.day).getDay();
                return `<div style="
                    display:flex; flex-direction:column; align-items:center; justify-content:center;
                    width:36px; height:48px; border-radius:8px;
                    background:${bg}; border:1px solid ${border};
                    font-size:10px; font-weight:600;
                " title="Giorno ${dc.day} (${dc.fest?'Festivo':'Feriale'}): ${dc.count}/${dc.expected} turni${dc.count < dc.expected ? ' ⚠️ SCOPERTO' : ''}&#10;${Object.entries(dc.shifts).map(([k,v]) => k+':'+v).join(', ')}">
                    <span style="font-size:9px; opacity:0.6;">${dayNamesShort[dow]}</span>
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
        </div>`;
    } // end team overview

    html += `</div>`; // close reports-wrap
    tab.innerHTML = html;
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

