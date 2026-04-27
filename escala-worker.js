// ============================================================
//  NurseShift Pro — escala-worker.js v3.0
//  Solver Monte Carlo isolado em Web Worker.
//  Roda fora da main thread → UI nunca trava enquanto gera.
//  Recebe { NURSES, schedule, occurrences, requests, hourLimits, m, y, startDay }
//  Retorna { scheduleModel, fitness, emptyShifts, epochs }
// ============================================================

// SHIFTS local — duplicado do app.js para o worker ser autocontido.
// Quando alterar shifts no app.js, atualizar aqui também.
const SHIFTS = {
    'M1': { name: 'Mattina 1',           h: 7.0, color: '#f59e0b', text: '#1a1a00', period: 'morning'   },
    'M2': { name: 'Mattina 2',           h: 4.5, color: '#fcd34d', text: '#1a1a00', period: 'morning'   },
    'MF': { name: 'Mattina Festivo',     h: 7.5, color: '#f97316', text: '#fff',    period: 'morning'   },
    'G':  { name: 'Giornata Intera',     h: 8,   color: '#0ea5e9', text: '#fff',    period: 'morning'   },
    'P':  { name: 'Pomeriggio',          h: 8,   color: '#8b5cf6', text: '#fff',    period: 'afternoon' },
    'PF': { name: 'Pomeriggio Festivo',  h: 7.5, color: '#a78bfa', text: '#fff',    period: 'afternoon' },
    'N':  { name: 'Notte',               h: 9,   color: '#1e1b4b', text: '#fff',    period: 'night'     },
    'OFF':{ name: 'Riposo',              h: 0,   color: 'rgba(255,255,255,0.03)', text: 'rgba(255,255,255,0.2)', period: 'off' },
    'FE': { name: 'Ferie',               h: 0,   color: '#10b981', text: '#fff',    period: 'off'       },
    'AT': { name: 'Certificato/Licenza', h: 0,   color: '#ef4444', text: '#fff',    period: 'off'       },
};

function sanitizeDate(d) {
    if (!d || typeof d !== 'string') return '';
    return d.split('T')[0];
}

function isWeekend(monthDate, day) {
    const d = new Date(monthDate.getFullYear(), monthDate.getMonth(), day);
    return d.getDay() === 0 || d.getDay() === 6;
}

function getItalianHolidays(year) {
    const fixed = [
        [0, 1], [0, 6], [3, 25], [4, 1], [5, 2], [7, 15],
        [10, 1], [11, 8], [11, 25], [11, 26]
    ];
    const a = year % 19, b = Math.floor(year / 100), c = year % 100;
    const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4), k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const mz = Math.floor((a + 11 * h + 22 * l) / 451);
    const month = Math.floor((h + l - 7 * mz + 114) / 31) - 1;
    const day = ((h + l - 7 * mz + 114) % 31) + 1;
    const easter = new Date(year, month, day);
    const easterMon = new Date(year, month, day + 1);
    fixed.push([easter.getMonth(), easter.getDate()]);
    fixed.push([easterMon.getMonth(), easterMon.getDate()]);
    return fixed;
}

function isHoliday(monthDate, day) {
    const holidays = getItalianHolidays(monthDate.getFullYear());
    return holidays.some(([hm, hd]) => hm === monthDate.getMonth() && hd === day);
}

function isFestivo(monthDate, day) {
    return isWeekend(monthDate, day) || isHoliday(monthDate, day);
}

// ── ENTRY POINT ─────────────────────────────────────────────
self.addEventListener('message', (event) => {
    try {
        const payload = event.data || {};
        const result = runSolver(payload);
        self.postMessage({ ok: true, ...result });
    } catch (err) {
        console.error('[escala-worker] erro:', err);
        self.postMessage({ ok: false, error: err && err.message ? err.message : String(err) });
    }
});

function runSolver(payload) {
    const NURSES        = payload.NURSES || [];
    const occurrences   = payload.occurrences || [];
    const requests      = payload.requests || [];
    const hourLimits    = payload.hourLimits || {};
    const m             = payload.m;
    const y             = payload.y;
    const startDay      = payload.startDay || 1;
    const incomingSched = payload.schedule || {};
    const MAX_EPOCHS    = payload.MAX_EPOCHS || 1200;

    const monthRef = new Date(y, m, 1);
    const days = new Date(y, m + 1, 0).getDate();
    const nextMonthDays = new Date(y, m + 2, 0).getDate();
    const simDays = days + nextMonthDays;
    const hasVacations = occurrences.length > 0;
    const MAX_CONSEC = hasVacations ? 4 : 3;

    // Contexto dos últimos 4 dias do mês anterior
    const prevM = m === 0 ? 11 : m - 1;
    const prevY = m === 0 ? y - 1 : y;
    const prevMonthDaysCount = new Date(prevY, prevM + 1, 0).getDate();
    const prevContext = {};
    NURSES.forEach(n => {
        prevContext[n.id] = {};
        for (let off = 0; off <= 3; off++) {
            const pDay = prevMonthDaysCount - off;
            if (pDay < 1) continue;
            const k = `${String(n.id).trim()}_${prevM}_${prevY}_${pDay}`;
            const code = incomingSched[k];
            if (code) {
                prevContext[n.id][-off] = code;
            }
        }
    });

    function getSh(simObj, nId, d) {
        if (d <= 0) {
            const ctx = prevContext[nId];
            return ctx ? ctx[d] : undefined;
        }
        return simObj[`${nId}_${m}_${y}_${d}`];
    }
    function setSh(simObj, nId, d, code) {
        if (d <= 0) return;
        simObj[`${nId}_${m}_${y}_${d}`] = code;
    }
    function getShiftFromIncoming(nId, d) {
        return incomingSched[`${String(nId).trim()}_${m}_${y}_${d}`] || 'OFF';
    }

    function nurseHoursTemp(simObj, nId) {
        let h = 0;
        for (let d = 1; d <= days; d++) h += (SHIFTS[getSh(simObj, nId, d)]?.h || 0);
        return h;
    }

    function canWorkConsecTemp(simObj, nId, day) {
        let count = 1;
        let d = day - 1;
        while (d >= -3 && getSh(simObj, nId, d) && !['OFF', 'FE', 'AT'].includes(getSh(simObj, nId, d))) { count++; d--; }
        d = day + 1;
        while (d <= simDays && getSh(simObj, nId, d) && !['OFF', 'FE', 'AT'].includes(getSh(simObj, nId, d))) { count++; d++; }
        if (count > MAX_CONSEC) return false;
        if (count >= 4) {
            let startS = day;
            while (startS > -3 && getSh(simObj, nId, startS - 1) && !['OFF', 'FE', 'AT'].includes(getSh(simObj, nId, startS - 1))) startS--;
            let endS = day;
            while (endS < simDays && getSh(simObj, nId, endS + 1) && !['OFF', 'FE', 'AT'].includes(getSh(simObj, nId, endS + 1))) endS++;
            let otherBlocks = 0;
            let tempStreak = 0;
            for (let i = 1; i <= simDays; i++) {
                if (i >= startS && i <= endS) { tempStreak = 0; continue; }
                const s = getSh(simObj, nId, i);
                if (s && !['OFF', 'FE', 'AT'].includes(s)) tempStreak++;
                else { if (tempStreak >= 4) otherBlocks++; tempStreak = 0; }
            }
            if (tempStreak >= 4) otherBlocks++;
            if (otherBlocks > 0) return false;
        }
        return true;
    }

    function checkTransitions(simObj, nId, day, code) {
        const morningShifts = ['M1', 'M2', 'MF', 'G'];
        const afternoonShifts = ['P', 'PF'];
        const prev = getSh(simObj, nId, day - 1) || null;
        const next = day < simDays ? getSh(simObj, nId, day + 1) : null;
        if (morningShifts.includes(code) && afternoonShifts.includes(prev)) return false;
        if (afternoonShifts.includes(code) && morningShifts.includes(next)) return false;
        if (afternoonShifts.includes(code) && (afternoonShifts.includes(prev) || afternoonShifts.includes(next))) return false;
        if (prev === 'N' && !['OFF', 'FE', 'AT'].includes(code)) return false;
        return true;
    }

    function canAssignTemp(simObj, nId, day, code) {
        if (!code || ['OFF', 'FE', 'AT'].includes(code)) return true;
        if (!checkTransitions(simObj, nId, day, code)) return false;
        const prev = getSh(simObj, nId, day - 1);
        if (prev && ['OFF', 'FE', 'AT'].includes(prev)) {
            const prev2 = getSh(simObj, nId, day - 2);
            if (prev2 && !['OFF', 'FE', 'AT'].includes(prev2)) return false;
        }
        return true;
    }

    // Pre-assigned shifts: occurrences + approved requests
    const preAssignedShifts = [];
    occurrences.forEach(occ => {
        const sDate = new Date(occ.start + 'T00:00:00');
        const eDate = new Date(occ.end + 'T00:00:00');
        for (let d = startDay; d <= simDays; d++) {
            const checkDate = new Date(y, m, d);
            checkDate.setHours(0, 0, 0, 0); sDate.setHours(0, 0, 0, 0); eDate.setHours(0, 0, 0, 0);
            if (checkDate >= sDate && checkDate <= eDate) {
                preAssignedShifts.push({ nurseId: occ.nurseId, day: d, code: occ.type });
            }
        }
    });
    requests.forEach(req => {
        if (req.status !== 'approved') return;
        if (['FE', 'OFF', 'AT', 'OFF_INJ', 'vacation', 'justified', 'unexcused'].includes(req.type)) {
            let startStr = sanitizeDate(req.startDate || req.date || '');
            let endStr   = sanitizeDate(req.endDate   || startStr);
            if (!startStr && req.day) {
                const tempD = new Date(y, m, req.day);
                startStr = tempD.toISOString().split('T')[0];
                endStr = startStr;
            }
            if (startStr && endStr) {
                const sDate = new Date(startStr + 'T00:00:00');
                const eDate = new Date(endStr + 'T00:00:00');
                let code = ['AT', 'FE'].includes(req.type) ? req.type : (req.type === 'vacation' ? 'FE' : 'OFF');
                for (let d = startDay; d <= simDays; d++) {
                    const checkDate = new Date(y, m, d);
                    checkDate.setHours(0, 0, 0, 0); sDate.setHours(0, 0, 0, 0); eDate.setHours(0, 0, 0, 0);
                    if (checkDate >= sDate && checkDate <= eDate) {
                        preAssignedShifts.push({ nurseId: req.nurseId, day: d, code: code });
                    }
                }
            }
        } else if (req.type === 'swap') {
            if (req.fromDay >= startDay && req.fromDay <= simDays) {
                preAssignedShifts.push({ nurseId: req.fromNurseId, day: req.fromDay, code: req.toShift });
            }
            if (req.toDay >= startDay && req.toDay <= simDays) {
                preAssignedShifts.push({ nurseId: req.toNurseId, day: req.toDay, code: req.fromShift });
            }
        }
    });

    function simulateOneScale() {
        let tSched = {};

        if (startDay > 1) {
            NURSES.forEach(n => {
                for (let d = 1; d < startDay; d++) {
                    const existingCode = getShiftFromIncoming(n.id, d);
                    if (existingCode && existingCode !== 'FO') {
                        setSh(tSched, n.id, d, existingCode);
                    } else {
                        setSh(tSched, n.id, d, 'OFF');
                    }
                }
            });
        }

        preAssignedShifts.forEach(item => {
            setSh(tSched, item.nurseId, item.day, item.code);
        });

        const nightCount = {}; NURSES.forEach(n => nightCount[n.id] = 0);
        const shiftCountTemp = (nId, type) => {
            let c = 0; for (let d = 1; d <= days; d++) if (getSh(tSched, nId, d) === type) c++; return c;
        };

        let emptyShifts = 0;

        // Fase: Noite
        for (let d = startDay; d <= simDays; d++) {
            const dow = new Date(y, m, d).getDay();
            if (dow === 0) {
                NURSES.forEach(n => {
                    if (getSh(tSched, n.id, d - 1) === 'N') {
                        if (!['FE', 'AT'].includes(getSh(tSched, n.id, d))) {
                            setSh(tSched, n.id, d, 'N');
                            nightCount[n.id]++;
                            if (d + 1 <= simDays) setSh(tSched, n.id, d + 1, 'OFF');
                            if (d + 2 <= simDays) setSh(tSched, n.id, d + 2, 'OFF');
                        }
                    }
                });
                continue;
            }
            if (NURSES.some(n => getSh(tSched, n.id, d) === 'N')) continue;

            let eligible = NURSES.filter(n => {
                if (n.nightQuota === 0) return false;
                if (getSh(tSched, n.id, d)) return false;
                const p1 = getSh(tSched, n.id, d - 1);
                const p2 = getSh(tSched, n.id, d - 2);
                if (dow === 6 && d + 1 <= simDays) {
                    const sunShift = getSh(tSched, n.id, d + 1);
                    if (sunShift && ['FE', 'AT'].includes(sunShift)) return false;
                    if (!canWorkConsecTemp(tSched, n.id, d + 1)) return false;
                }
                if (p1 === 'N' && p2 === 'N') return false;
                if (p1 === 'N') return dow !== 6;
                if (p2 === 'N') return false;
                const individualLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                if (nurseHoursTemp(tSched, n.id) >= individualLimit) return false;
                if (!canWorkConsecTemp(tSched, n.id, d)) return false;
                if (!canAssignTemp(tSched, n.id, d, 'N')) return false;
                return true;
            });

            eligible.forEach(n => {
                const nLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                const nHours = nurseHoursTemp(tSched, n.id);
                const loadRatio = nHours / nLimit;
                const quota = n.nightQuota || 5;
                const nightRatio = nightCount[n.id] / quota;
                n._nightScore = (nightRatio * 100) + (loadRatio * 80) + (Math.random() * 5);
            });
            eligible.sort((a, b) => a._nightScore - b._nightScore);

            if (eligible.length > 0) {
                let pair = eligible.find(n => getSh(tSched, n.id, d - 1) === 'N');
                let chosen = pair;
                if (!chosen) {
                    const topScore = eligible[0]._nightScore;
                    const threshold = topScore + 10;
                    const topNurses = eligible.filter(n => n._nightScore <= threshold);
                    chosen = topNurses[Math.floor(Math.random() * topNurses.length)];
                }
                setSh(tSched, chosen.id, d, 'N');
                nightCount[chosen.id]++;
                if (!pair) {
                    if (d + 1 <= simDays && !getSh(tSched, chosen.id, d + 1)) setSh(tSched, chosen.id, d + 1, 'OFF');
                } else {
                    if (d + 1 <= simDays) setSh(tSched, chosen.id, d + 1, 'OFF');
                    if (d + 2 <= simDays) setSh(tSched, chosen.id, d + 2, 'OFF');
                }
            } else {
                emptyShifts += 1;
            }
        }

        // Fase: Diurna
        let globalGCount = 0, globalM2Count = 0;
        const nurseGM2Count = {};
        NURSES.forEach(n => { nurseGM2Count[n.id] = { g: 0, m2: 0 }; });

        for (let d = startDay; d <= simDays; d++) {
            const dow = new Date(y, m, d).getDay();
            const festivo = isFestivo(monthRef, d);

            if (dow === 0 && d > 1) {
                NURSES.forEach(n => {
                    const satShift = getSh(tSched, n.id, d - 1);
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
                targets = ['MF', 'PF'];
            } else {
                let thirdShift;
                if (globalGCount <= globalM2Count) {
                    thirdShift = 'G'; globalGCount++;
                } else if (globalM2Count < globalGCount) {
                    thirdShift = 'M2'; globalM2Count++;
                } else {
                    thirdShift = Math.random() < 0.5 ? 'G' : 'M2';
                    if (thirdShift === 'G') globalGCount++; else globalM2Count++;
                }
                targets = ['M1', 'P', thirdShift];
            }

            targets.sort(() => Math.random() - 0.5);

            for (let t of targets) {
                let free = NURSES.filter(n => {
                    if (getSh(tSched, n.id, d)) return false;
                    let p1 = d > 1 ? getSh(tSched, n.id, d - 1) : null;
                    if (p1 === 'N') return false;
                    if (!canWorkConsecTemp(tSched, n.id, d)) return false;
                    const individualLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
                    if (nurseHoursTemp(tSched, n.id) + SHIFTS[t].h > (individualLimit + 1)) return false;
                    if (!canAssignTemp(tSched, n.id, d, t)) return false;
                    if (dow === 6 && d + 1 <= simDays) {
                        const sunShift = getSh(tSched, n.id, d + 1);
                        if (sunShift && ['FE', 'AT'].includes(sunShift)) return false;
                    }
                    return true;
                });

                if (free.length === 0) {
                    free = NURSES.filter(n => {
                        if (getSh(tSched, n.id, d)) return false;
                        let p1 = d > 1 ? getSh(tSched, n.id, d - 1) : null;
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
                        const p1 = d > 1 ? getSh(tSched, n.id, d - 1) : null;
                        const loadRatio = nHours / nLimit;
                        const loadPenalty = loadRatio * 200;
                        const seqPenalty = (p1 === t) ? 300 : 0;
                        const typeCount = shiftCountTemp(n.id, t);
                        const typePenalty = typeCount * 25;
                        let gm2Bias = 0;
                        if (t === 'G') gm2Bias = (nurseGM2Count[n.id]?.g || 0) * 18;
                        if (t === 'M2') gm2Bias = (nurseGM2Count[n.id]?.m2 || 0) * 18;
                        let pmBias = 0;
                        if (t === 'P') pmBias = shiftCountTemp(n.id, 'P') * 15;
                        if (t === 'M1') pmBias = shiftCountTemp(n.id, 'M1') * 15;
                        let wkBias = 0;
                        if (festivo) {
                            let wkCount = 0;
                            for (let wd = 1; wd < d; wd++) {
                                if (isFestivo(monthRef, wd) && getSh(tSched, n.id, wd) && !['OFF', 'FE', 'AT'].includes(getSh(tSched, n.id, wd))) wkCount++;
                            }
                            wkBias = wkCount * 20;
                        }
                        const workedYesterday = (p1 && !['OFF', 'FE', 'AT'].includes(p1)) ? 8 : 0;
                        n.tmpScore = loadPenalty + typePenalty + seqPenalty + gm2Bias + pmBias + wkBias + workedYesterday + (Math.random() * 4);
                    });
                    free.sort((a, b) => a.tmpScore - b.tmpScore);
                    const chosen = free[0];
                    setSh(tSched, chosen.id, d, t);
                    if (t === 'G' && nurseGM2Count[chosen.id]) nurseGM2Count[chosen.id].g++;
                    if (t === 'M2' && nurseGM2Count[chosen.id]) nurseGM2Count[chosen.id].m2++;
                } else {
                    emptyShifts += 1;
                }
            }
        }

        NURSES.forEach(n => {
            for (let d = startDay; d <= simDays; d++) {
                if (!getSh(tSched, n.id, d)) setSh(tSched, n.id, d, 'OFF');
            }
        });

        // Fitness
        let repScore = 0, weekendPenalty = 0, postNightViolations = 0;
        let shiftTypePenalty = 0, overloadPenalty = 0, loadVariancePenalty = 0;
        const nurseMetrics = [];
        NURSES.forEach(n => {
            const h = nurseHoursTemp(tSched, n.id);
            const nLimit = hourLimits[n.id] ? hourLimits[n.id] : 130;
            const loadPct = h / nLimit;
            let cons = 0;
            for (let dt = 2; dt <= days; dt++) {
                const curr = getSh(tSched, n.id, dt);
                const prev = getSh(tSched, n.id, dt - 1);
                if (curr && curr !== 'OFF' && curr !== 'FE' && curr !== 'AT') {
                    if (curr === prev) cons++;
                }
            }
            repScore += cons;
            let festiviWorked = 0;
            for (let dd = 1; dd <= days; dd++) {
                if (isFestivo(monthRef, dd)) {
                    const s = getSh(tSched, n.id, dd);
                    if (s && !['OFF', 'FE', 'AT'].includes(s)) festiviWorked++;
                }
            }
            if (festiviWorked > 4) weekendPenalty += (festiviWorked - 4) * 3000;
            if (festiviWorked === 0) weekendPenalty += 2000;
            for (let dd = 1; dd < days; dd++) {
                if (getSh(tSched, n.id, dd) === 'N') {
                    const next = getSh(tSched, n.id, dd + 1);
                    if (next && !['OFF', 'FE', 'AT', 'N'].includes(next)) postNightViolations++;
                }
            }
            const counts = {};
            ['M1', 'M2', 'MF', 'G', 'P', 'PF', 'N'].forEach(c => { counts[c] = 0; });
            for (let dd = 1; dd <= days; dd++) {
                const s = getSh(tSched, n.id, dd);
                if (counts[s] !== undefined) counts[s]++;
            }
            shiftTypePenalty += Math.abs(counts['G'] - counts['M2']) * 200;
            shiftTypePenalty += Math.abs(counts['P'] - counts['M1']) * 150;
            shiftTypePenalty += Math.abs(counts['MF'] - counts['PF']) * 150;
            if (h > nLimit) overloadPenalty += (h - nLimit) * 500;
            nurseMetrics.push({ id: n.id, h, nLimit, loadPct, counts });
        });
        const avgLoadPct = nurseMetrics.reduce((s, nm) => s + nm.loadPct, 0) / nurseMetrics.length;
        nurseMetrics.forEach(nm => {
            loadVariancePenalty += Math.abs(nm.loadPct - avgLoadPct) * 8000;
        });
        let nightBalancePenalty = 0;
        const nightCounts = nurseMetrics.map(nm => nm.counts['N']);
        const avgNight = nightCounts.reduce((s, v) => s + v, 0) / nightCounts.length;
        nightCounts.forEach(nc => { nightBalancePenalty += Math.abs(nc - avgNight) * 400; });
        let pBalancePenalty = 0;
        const pCounts = nurseMetrics.map(nm => nm.counts['P']);
        const avgP = pCounts.reduce((s, v) => s + v, 0) / pCounts.length;
        pCounts.forEach(pc => { pBalancePenalty += Math.abs(pc - avgP) * 300; });
        let incompleteDays = 0;
        for (let dd = 1; dd <= days; dd++) {
            if (isFestivo(monthRef, dd)) continue;
            let filledCount = 0;
            NURSES.forEach(n => {
                const s = getSh(tSched, n.id, dd);
                if (s && !['OFF', 'FE', 'AT', 'N'].includes(s)) filledCount++;
            });
            if (filledCount < 3) incompleteDays++;
        }

        const fitness = 200000
            - (emptyShifts * 25000)
            - (incompleteDays * 35000)
            - (repScore * 5000)
            - (weekendPenalty)
            - (postNightViolations * 8000)
            - (shiftTypePenalty)
            - (overloadPenalty)
            - (loadVariancePenalty)
            - (nightBalancePenalty)
            - (pBalancePenalty);

        return { scheduleModel: tSched, fitness, emptyShifts };
    }

    let bestSim = null;
    let validIterCount = 0;
    let epochs = 0;
    for (let i = 0; i < MAX_EPOCHS; i++) {
        const sim = simulateOneScale();
        epochs++;
        if (!bestSim || sim.fitness > bestSim.fitness) {
            bestSim = sim;
            validIterCount = 0;
        }
        if (bestSim.emptyShifts === 0 && bestSim.fitness > 195000) {
            validIterCount++;
            if (validIterCount > 10) break;
        }
        // Reporta progresso a cada 100 epochs (não bloqueia o solver — apenas posta)
        if ((i + 1) % 100 === 0) {
            self.postMessage({ ok: true, progress: true, epoch: i + 1, bestFitness: bestSim.fitness, emptyShifts: bestSim.emptyShifts });
        }
    }

    return {
        scheduleModel: bestSim.scheduleModel,
        fitness: bestSim.fitness,
        emptyShifts: bestSim.emptyShifts,
        epochs
    };
}
