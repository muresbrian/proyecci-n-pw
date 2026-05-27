/* ========================================
   WUZI Analytics — Main Application
   ======================================== */

(() => {
    'use strict';

    // ── CSV file paths ──
    const CSV_BASE = 'https://raw.githubusercontent.com/muresbrian/proyecci-n-pw/main/webapp/Reportes_Individuales_CSV/';
    const CSV_FILES = {
        tendencias:     'Resumen_Tendencias_Actual.csv',
        proyeccion:     'Proyeccion_Cierre_Mes.csv',
        abonosMensuales:'Abonos_Mensuales.csv',
        efectoNomina:   'Efecto_Nomina.csv',
        patronesDia:    'Patrones_Dia_Semana.csv',
        variacionDiaria:'Variacion_Diaria.csv',
        abonosDiarios:  'Abonos_Diarios.csv',
        patronesQuincena:'Patrones_Quincena.csv',
        abonosSemanales:'Abonos_Semanales.csv',
        patronesSemana: 'Patrones_Semana_Mes.csv',
        variacionMensual:'Variacion_Mensual.csv',
        variacionSemanal:'Variacion_Semanal.csv',
        semaforo:       'Semaforo_Salud.csv',
        ranking:        'Ranking.csv',
        trxSem:         'TRX_SEM_clean.csv',
        contexto:       'Contexto_Holders.csv',
    };

    // ── Global data store ──
    const DATA = {};
    let allHolders = [];
    let currentDetailHolder = null;
    let showOnlyTop300 = false;
    const abonosTotalesMap = new Map();
    let top300Holders = [];
    const normalizedToCanonical = new Map();
    let weekCols = [];

    // Chart instances (for cleanup)
    const charts = {};

    // ── Utility helpers ──
    const $ = sel => document.querySelector(sel);
    const $$ = sel => document.querySelectorAll(sel);
    const fmt = (n) => {
        if (n == null || isNaN(n)) return '—';
        return new Intl.NumberFormat('es-MX', { maximumFractionDigits: 2 }).format(n);
    };
    const fmtCurrency = (n) => {
        if (n == null || isNaN(n)) return '—';
        return '$' + new Intl.NumberFormat('es-MX', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(n);
    };
    const parseNum = (v) => {
        if (v == null || v === '') return 0;
        const s = String(v).replace(/[$,\s]/g, '').replace(/−/g, '-');
        const n = parseFloat(s);
        return isNaN(n) ? 0 : n;
    };
    const normalizeHolder = (name) => {
        if (!name) return '';
        return name.trim()
            .toUpperCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '');
    };

    function extractWeekCols() {
        if (weekCols.length > 0) return;
        if (!DATA.trxSem || DATA.trxSem.length === 0) return;

        const row0 = DATA.trxSem[0];
        const weekHeaders = Object.keys(row0).filter(k => k.startsWith('Semana '));

        const parsedWeeks = weekHeaders.map(h => {
            const m = h.match(/Semana (\d+) (\d{4})/);
            return {
                original: h,
                week: m ? parseInt(m[1]) : 0,
                year: m ? parseInt(m[2]) : 0
            };
        });

        parsedWeeks.sort((a, b) => {
            if (a.year !== b.year) return a.year - b.year;
            return a.week - b.week;
        });

        weekCols = parsedWeeks.map(pw => pw.original);
    }

    function formatWeekHeader(weekStr) {
        const m = weekStr.match(/Semana (\d+) (\d{4})/);
        if (m) {
            const week = m[1];
            const year = m[2].slice(-2);
            return `S${week} '${year}`;
        }
        return weekStr;
    }

    // ── CSV Loading ──
    function cleanRow(row) {
        const clean = {};
        Object.keys(row).forEach(k => {
            let cleanKey = k.trim().replace(/^\ufeff/, '');
            const norm = cleanKey.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
            if (norm === 'holder') {
                cleanKey = 'Holder';
            } else if (norm === 'director') {
                cleanKey = 'Director';
            } else if (norm === 'vendedor') {
                cleanKey = 'Vendedor';
            } else if (norm === 'ranking') {
                cleanKey = 'Ranking';
            } else if (norm === 'semaforo') {
                cleanKey = 'Semáforo';
            } else if (norm === 'contexto') {
                cleanKey = 'Contexto';
            }
            clean[cleanKey] = row[k];
        });
        return clean;
    }

    async function loadCSV(filename) {
        try {
            const response = await fetch(CSV_BASE + filename);
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            const text = await response.text();
            return new Promise((resolve) => {
                Papa.parse(text, {
                    header: true,
                    skipEmptyLines: true,
                    dynamicTyping: false,
                    complete: (results) => {
                        const cleaned = results.data.map(row => cleanRow(row));
                        resolve(cleaned);
                    },
                    error: (err) => {
                        console.error('CSV parse error:', filename, err);
                        resolve([]);
                    }
                });
            });
        } catch (err) {
            console.error('CSV load error:', filename, err);
            return [];
        }
    }

    async function loadAllData() {
        const keys = Object.keys(CSV_FILES);
        const total = keys.length;
        let loaded = 0;

        const statusEl = $('#loading-status');
        const progressEl = $('#loading-progress');

        for (const key of keys) {
            statusEl.textContent = `Cargando ${CSV_FILES[key]}...`;
            DATA[key] = await loadCSV(CSV_FILES[key]);
            loaded++;
            progressEl.style.width = `${(loaded / total) * 100}%`;
        }

        // Normalize and build canonical name map
        normalizedToCanonical.clear();
        function addHolder(rawName) {
            if (!rawName) return;
            const name = rawName.trim();
            const norm = normalizeHolder(name);
            if (!norm || norm === 'HOLDER') return;
            if (!normalizedToCanonical.has(norm)) {
                normalizedToCanonical.set(norm, name);
            } else {
                const existing = normalizedToCanonical.get(norm);
                const isExistingAllCaps = existing === existing.toUpperCase();
                const isNewAllCaps = name === name.toUpperCase();
                if (isExistingAllCaps && !isNewAllCaps) {
                    normalizedToCanonical.set(norm, name);
                }
            }
        }

        // Preferred order for canonical names: tendencias, proyeccion, semaforo
        const preferredKeys = ['tendencias', 'proyeccion', 'semaforo'];
        preferredKeys.forEach(k => {
            (DATA[k] || []).forEach(r => { if (r.Holder) addHolder(r.Holder); });
        });
        keys.forEach(k => {
            if (!preferredKeys.includes(k)) {
                (DATA[k] || []).forEach(r => { if (r.Holder) addHolder(r.Holder); });
            }
        });

        allHolders = Array.from(normalizedToCanonical.values()).sort((a, b) => a.localeCompare(b, 'es'));

        statusEl.textContent = 'Preparando dashboard...';
        progressEl.style.width = '100%';
    }

    // ── Lookup helpers ──
    function findRow(dataset, holder) {
        if (!DATA[dataset]) return null;
        const norm = normalizeHolder(holder);
        return DATA[dataset].find(r => r.Holder && normalizeHolder(r.Holder) === norm) || null;
    }

    // ── Top 300 helper functions ──
    function calculateAbonosTotales() {
        abonosTotalesMap.clear();
        // First populate from abonosMensuales
        (DATA.abonosMensuales || []).forEach(row => {
            if (!row.Holder) return;
            let sum = 0;
            Object.keys(row).forEach(k => {
                if (k !== 'Holder') {
                    sum += parseNum(row[k]);
                }
            });
            abonosTotalesMap.set(normalizeHolder(row.Holder), sum);
        });

        // Overwrite/update with full historical total from trxSem if available
        if (DATA.trxSem && DATA.trxSem.length > 1) {
            for (let i = 1; i < DATA.trxSem.length; i++) {
                const row = DATA.trxSem[i];
                if (!row.Holder) continue;
                let sum = 0;
                Object.keys(row).forEach(k => {
                    if (k !== 'Director' && k !== 'Holder') {
                        sum += parseNum(row[k]);
                    }
                });
                abonosTotalesMap.set(normalizeHolder(row.Holder), sum);
            }
        }
    }

    function computeTop300() {
        top300Holders = [...allHolders]
            .map(h => {
                const rrow = getRow('ranking', h);
                const rank = rrow && rrow.Ranking ? parseInt(rrow.Ranking) : null;
                return { holder: h, rank };
            })
            .filter(x => x.rank !== null && x.rank >= 1 && x.rank <= 300)
            .sort((a, b) => a.rank - b.rank)
            .map(x => x.holder);
    }

    function getActiveHolders() {
        return showOnlyTop300 ? top300Holders : allHolders;
    }

    // ── Index maps for quick lookups ──
    const INDEX = {};
    function buildIndexes() {
        for (const key of Object.keys(DATA)) {
            INDEX[key] = {};
            (DATA[key] || []).forEach(row => {
                if (row.Holder) {
                    const norm = normalizeHolder(row.Holder);
                    INDEX[key][norm] = row;
                }
            });
        }
    }
    function getRow(dataset, holder) {
        const norm = normalizeHolder(holder);
        return (INDEX[dataset] && INDEX[dataset][norm]) || null;
    }

    function getHolderContext(holder) {
        const crow = getRow('contexto', holder);
        return crow ? (crow.Contexto || '') : '';
    }

    // ══════════════════════════════
    //  TAB NAVIGATION
    // ══════════════════════════════
    function initTabs() {
        $$('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const tab = btn.dataset.tab;
                switchTab(tab);
            });
        });
    }

    function switchTab(tab) {
        $$('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
        $$('.tab-content').forEach(s => s.classList.toggle('active', s.id === 'section-' + tab));
    }

    // ══════════════════════════════
    //  ANIMATED COUNTER
    // ══════════════════════════════
    function animateCounter(el, target, duration = 1200) {
        const start = 0;
        const startTime = performance.now();
        function step(now) {
            const elapsed = now - startTime;
            const progress = Math.min(elapsed / duration, 1);
            const eased = 1 - Math.pow(1 - progress, 3);
            const current = Math.round(start + (target - start) * eased);
            el.textContent = new Intl.NumberFormat('es-MX').format(current);
            if (progress < 1) requestAnimationFrame(step);
        }
        requestAnimationFrame(step);
    }

    // ══════════════════════════════
    //  SECTION 1: OVERVIEW
    // ══════════════════════════════
    function renderOverview() {
        const activeSet = new Set(getActiveHolders());
        const activeNormalizedSet = new Set(Array.from(activeSet).map(h => normalizeHolder(h)));

        // --- KPIs ---
        const tend = DATA.tendencias || [];
        const proy = DATA.proyeccion || [];

        const totalComercios = activeSet.size;
        let alza = 0, baja = 0, estable = 0;
        tend.forEach(r => {
            if (!r.Holder || !activeNormalizedSet.has(normalizeHolder(r.Holder))) return;
            const t = (r.Tendencia_Mes || '').toLowerCase();
            if (t.includes('alza')) alza++;
            else if (t.includes('baja')) baja++;
            else if (t.includes('estable')) estable++;
        });

        let meta = 0, riesgo = 0, noLlegara = 0, nuevo = 0;
        proy.forEach(r => {
            if (!r.Holder || !activeNormalizedSet.has(normalizeHolder(r.Holder))) return;
            const d = (r.Diagnostico || '').toLowerCase();
            if (d.includes('no llegar')) noLlegara++;
            else if (d.includes('llegar')) meta++;
            if (d.includes('riesgo')) riesgo++;
            if (d.includes('nuevo')) nuevo++;
        });

        let constante = 0, intermitente = 0, abandono = 0, sinActividad = 0;
        activeSet.forEach(holder => {
            const srow = getRow('semaforo', holder);
            if (!srow) {
                sinActividad++;
                return;
            }
            const s = (srow['Semáforo'] || '').toLowerCase();
            if (s.includes('constante')) constante++;
            else if (s.includes('intermitente')) intermitente++;
            else if (s.includes('riesgo') || s.includes('abandono')) abandono++;
            else if (s.includes('sin actividad')) sinActividad++;
        });

        // --- TPV & SPEI KPIs ---
        let tpvCount = 0;
        let wuziOnly = 0;
        let bpOnly = 0;
        let tpvBoth = 0;

        let speiCount = 0;
        let speiOnly = 0;
        let speiBoth = 0;

        proy.forEach(r => {
            if (!r.Holder || !activeNormalizedSet.has(normalizeHolder(r.Holder))) return;
            const wuziVal = parseNum(r.Wuzi);
            const bpVal = parseNum(r.BP);
            const speiVal = parseNum(r.SPEI);

            const hasWuzi = wuziVal > 0;
            const hasBp = bpVal > 0;
            const hasTpv = hasWuzi || hasBp;
            const hasSpei = speiVal > 0;

            if (hasTpv) {
                tpvCount++;
                if (hasWuzi && hasBp) {
                    tpvBoth++;
                } else if (hasWuzi) {
                    wuziOnly++;
                } else if (hasBp) {
                    bpOnly++;
                }
            }

            if (hasSpei) {
                speiCount++;
                if (hasTpv) {
                    speiBoth++;
                } else {
                    speiOnly++;
                }
            }
        });

        animateCounter($('#kpi-total'), totalComercios);
        animateCounter($('#kpi-alza'), alza);
        animateCounter($('#kpi-baja'), baja);
        animateCounter($('#kpi-salud-constante'), constante);
        animateCounter($('#kpi-salud-intermitente'), intermitente);
        animateCounter($('#kpi-salud-riesgo'), abandono);
        
        const kpiSinActividad = $('#kpi-salud-sinactividad');
        if (kpiSinActividad) animateCounter(kpiSinActividad, sinActividad);

        const kpiTpv = $('#kpi-tpv');
        if (kpiTpv) animateCounter(kpiTpv, tpvCount);

        const kpiSpei = $('#kpi-spei');
        if (kpiSpei) animateCounter(kpiSpei, speiCount);

        const kpiTpvBreakdown = $('#kpi-tpv-breakdown');
        if (kpiTpvBreakdown) kpiTpvBreakdown.textContent = `Wuzi: ${wuziOnly} | BP: ${bpOnly} | Ambas: ${tpvBoth}`;

        const kpiSpeiBreakdown = $('#kpi-spei-breakdown');
        if (kpiSpeiBreakdown) kpiSpeiBreakdown.textContent = `Solo SPEI: ${speiOnly} | Ambas: ${speiBoth}`;

        const sinActividadCard = $('.kpi-card.kpi-salud-sinactividad');
        if (sinActividadCard) {
            sinActividadCard.style.display = showOnlyTop300 ? 'none' : '';
        }

        $('#total-comercios-badge').textContent = `${totalComercios} comercios`;

        // --- Chart: Semáforo Doughnut ---
        renderSemaforoChart(constante, intermitente, abandono, sinActividad);

        // --- Chart: Top 15 Bar ---
        renderTop15Chart(proy, activeSet);

        // --- Chart: Radar Day of Week ---
        renderRadarChart(activeSet);
    }

    function renderSemaforoChart(constante, intermitente, abandono, sinActividad) {
        const ctx = $('#chart-semaforo').getContext('2d');
        if (charts.semaforo) charts.semaforo.destroy();
        charts.semaforo = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Constante ✅', 'Intermitente ⚠️', 'En Riesgo 🚨', 'Sin Actividad 💤'],
                datasets: [{
                    data: [constante, intermitente, abandono, sinActividad],
                    backgroundColor: ['#00e676', '#ffab00', '#ff4466', '#5a6478'],
                    borderColor: 'transparent',
                    borderWidth: 0,
                    hoverOffset: 14,
                    spacing: 3,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                cutout: '65%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { color: '#8a94a8', font: { family: "'Inter'", size: 12 }, padding: 18, usePointStyle: true, pointStyleWidth: 10 }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(18,22,42,0.95)',
                        titleColor: '#e8ecf4', bodyColor: '#8a94a8',
                        borderColor: 'rgba(255,255,255,0.08)', borderWidth: 1,
                        cornerRadius: 8, padding: 12,
                        titleFont: { family: "'Inter'", weight: '600' },
                        bodyFont: { family: "'Inter'" },
                    }
                }
            }
        });
    }
    function renderTop15Chart(proy, activeSet) {
        const activeNormalizedSet = new Set(Array.from(activeSet).map(h => normalizeHolder(h)));
        const sorted = [...proy]
            .map(r => {
                const rawName = (r.Holder || '').trim();
                const norm = normalizeHolder(rawName);
                const canonicalName = normalizedToCanonical.get(norm) || rawName;
                return { holder: canonicalName, norm, val: parseNum(r.Abonos_Actuales_Mayo) };
            })
            .filter(r => r.holder && r.val > 0 && activeNormalizedSet.has(r.norm))
            .sort((a, b) => b.val - a.val)
            .slice(0, 15);

        const ctx = $('#chart-top15').getContext('2d');
        if (charts.top15) charts.top15.destroy();

        const gradient = ctx.createLinearGradient(0, 0, ctx.canvas.width, 0);
        gradient.addColorStop(0, 'rgba(0,212,255,0.7)');
        gradient.addColorStop(1, 'rgba(123,47,247,0.7)');

        charts.top15 = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: sorted.map(r => r.holder.length > 30 ? r.holder.slice(0, 28) + '…' : r.holder),
                datasets: [{
                    data: sorted.map(r => r.val),
                    backgroundColor: gradient,
                    borderRadius: 6,
                    borderSkipped: false,
                    maxBarThickness: 22,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                indexAxis: 'y',
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: 'rgba(18,22,42,0.95)',
                        titleColor: '#e8ecf4', bodyColor: '#8a94a8',
                        borderColor: 'rgba(255,255,255,0.08)', borderWidth: 1,
                        cornerRadius: 8, padding: 12,
                        callbacks: {
                            label: (ctx) => fmtCurrency(ctx.parsed.x),
                        },
                        titleFont: { family: "'Inter'", weight: '600' },
                        bodyFont: { family: "'Inter'" },
                    }
                },
                scales: {
                    y: {
                        ticks: { color: '#8a94a8', font: { family: "'Inter'", size: 11 } },
                        grid: { display: false },
                        border: { display: false },
                    },
                    x: {
                        ticks: {
                            color: '#5a6478',
                            font: { family: "'Inter'", size: 10 },
                            callback: v => fmtCurrency(v),
                        },
                        grid: { color: 'rgba(255,255,255,0.04)' },
                        border: { display: false },
                    }
                },
                onClick: (evt, elements) => {
                    if (elements.length > 0) {
                        const idx = elements[0].index;
                        const holder = sorted[idx].holder;
                        openDetail(holder);
                    }
                }
            }
        });
    }

    function renderRadarChart(activeSet) {
        const activeNormalizedSet = new Set(Array.from(activeSet).map(h => normalizeHolder(h)));
        const dias = DATA.patronesDia || [];
        // Exclude Monday (Lunes) as weekend terminal cuts skew the scale
        const dayKeys = ['2-Martes', '3-Miércoles', '4-Jueves', '5-Viernes', '6-Sábado', '7-Domingo'];
        const dayLabels = ['Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
        const sums = dayKeys.map(() => 0);

        dias.forEach(row => {
            if (!row.Holder || !activeNormalizedSet.has(normalizeHolder(row.Holder))) return;
            dayKeys.forEach((k, i) => {
                sums[i] += parseNum(row[k]);
            });
        });

        const ctx = $('#chart-radar').getContext('2d');
        if (charts.radar) charts.radar.destroy();
        charts.radar = new Chart(ctx, {
            type: 'radar',
            data: {
                labels: dayLabels,
                datasets: [{
                    label: 'Abonos Totales',
                    data: sums,
                    backgroundColor: 'rgba(0,212,255,0.12)',
                    borderColor: '#00d4ff',
                    borderWidth: 2,
                    pointBackgroundColor: '#00d4ff',
                    pointBorderColor: '#0a0e1a',
                    pointBorderWidth: 2,
                    pointRadius: 5,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: 'rgba(18,22,42,0.95)',
                        titleColor: '#e8ecf4', bodyColor: '#8a94a8',
                        borderColor: 'rgba(255,255,255,0.08)', borderWidth: 1,
                        cornerRadius: 8, padding: 12,
                        callbacks: { label: ctx => fmtCurrency(ctx.parsed.r) },
                        titleFont: { family: "'Inter'", weight: '600' },
                        bodyFont: { family: "'Inter'" },
                    }
                },
                scales: {
                    r: {
                        angleLines: { color: 'rgba(255,255,255,0.06)' },
                        grid: { color: 'rgba(255,255,255,0.06)' },
                        pointLabels: { color: '#8a94a8', font: { family: "'Inter'", size: 12 } },
                        ticks: {
                            color: '#5a6478',
                            backdropColor: 'transparent',
                            font: { family: "'Inter'", size: 10 },
                            callback: v => fmtCurrency(v),
                        }
                    }
                }
            }
        });
    }

    // ══════════════════════════════
    //  SECTION 2: TABLE
    // ══════════════════════════════
    const TABLE_PAGE_SIZE = 50;
    let tableData = [];
    let tableFiltered = [];
    let tablePage = 0;
    let tableSort = { key: 'holder', dir: 'asc' };

    function buildTableHeader() {
        extractWeekCols();
        const thead = $('#comercios-table thead');
        if (!thead) return;

        const sortActiveKey = tableSort.key;
        const sortActiveDir = tableSort.dir;

        const getIcon = (k) => {
            if (sortActiveKey === k) {
                return sortActiveDir === 'asc' ? '↑' : '↓';
            }
            return '⇅';
        };

        const getClass = (k) => {
            return sortActiveKey === k ? 'sortable sort-active' : 'sortable';
        };

        let html = '<tr>';
        html += `<th data-sort="holder" class="${getClass('holder')}">Comercio <span class="sort-icon">${getIcon('holder')}</span></th>`;
        html += `<th data-sort="ranking" class="${getClass('ranking')} sort-num">Ranking <span class="sort-icon">${getIcon('ranking')}</span></th>`;
        html += `<th data-sort="director" class="${getClass('director')}">Director <span class="sort-icon">${getIcon('director')}</span></th>`;
        html += `<th data-sort="vendedor" class="${getClass('vendedor')}">Vendedor <span class="sort-icon">${getIcon('vendedor')}</span></th>`;
        html += `<th data-sort="abonos" class="${getClass('abonos')} sort-num">Abonos Mayo <span class="sort-icon">${getIcon('abonos')}</span></th>`;
        html += `<th data-sort="promedio" class="${getClass('promedio')} sort-num">Promedio Histórico <span class="sort-icon">${getIcon('promedio')}</span></th>`;
        html += `<th data-sort="proyeccion" class="${getClass('proyeccion')} sort-num">Proyección Cierre <span class="sort-icon">${getIcon('proyeccion')}</span></th>`;
        html += `<th data-sort="diagnostico" class="${getClass('diagnostico')}">Diagnóstico <span class="sort-icon">${getIcon('diagnostico')}</span></th>`;
        html += `<th data-sort="salud" class="${getClass('salud')}">Salud <span class="sort-icon">${getIcon('salud')}</span></th>`;

        weekCols.forEach(col => {
            const label = formatWeekHeader(col);
            const monthName = DATA.trxSem && DATA.trxSem[0] ? DATA.trxSem[0][col] : '';
            const tooltip = `${col}${monthName ? ' (' + monthName + ')' : ''}`;
            const extraClass = getClass(col) + ' sort-num';
            html += `<th data-sort="${col}" class="${extraClass}" title="${escHtml(tooltip)}">${escHtml(label)} <span class="sort-icon">${getIcon(col)}</span></th>`;
        });
        html += '</tr>';
        thead.innerHTML = html;

        initTableSortEvents();
    }

    function initTableSortEvents() {
        $$('#comercios-table thead th.sortable').forEach(th => {
            th.addEventListener('click', () => {
                const key = th.dataset.sort;
                if (tableSort.key === key) {
                    tableSort.dir = tableSort.dir === 'asc' ? 'desc' : 'asc';
                } else {
                    tableSort.key = key;
                    tableSort.dir = key === 'ranking' ? 'asc' : (th.classList.contains('sort-num') ? 'desc' : 'asc');
                }
                $$('#comercios-table thead th').forEach(t => {
                    t.classList.remove('sort-active');
                    const icon = t.querySelector('.sort-icon');
                    if (icon) icon.textContent = '⇅';
                });
                th.classList.add('sort-active');
                const activeIcon = th.querySelector('.sort-icon');
                if (activeIcon) {
                    activeIcon.textContent = tableSort.dir === 'asc' ? '↑' : '↓';
                }
                applyTableFilters();
            });
        });
    }

    function buildTableData() {
        extractWeekCols();
        tableData = getActiveHolders().map(h => {
            const prow = getRow('proyeccion', h);
            const trow = getRow('tendencias', h);
            const vmrow = getRow('variacionMensual', h);
            const srow = getRow('semaforo', h);
            const rrow = getRow('ranking', h);
            const trxSemRow = getRow('trxSem', h);
            // find last variacion mensual column
            let varMes = 0;
            if (vmrow) {
                const cols = Object.keys(vmrow).filter(k => k !== 'Holder');
                if (cols.length > 0) varMes = parseNum(vmrow[cols[cols.length - 1]]);
            }
            const rPos = rrow && rrow.Ranking ? parseInt(rrow.Ranking) : null;
            const rowObj = {
                holder: h,
                ranking: rPos,
                director: rrow ? (rrow.Director || '').trim() : '',
                vendedor: rrow ? (rrow.Vendedor || '').trim() : '',
                abonos: prow ? parseNum(prow.Abonos_Actuales_Mayo) : 0,
                promedio: prow ? parseNum(prow.Promedio_Mensual_Historico) : 0,
                proyeccion: prow ? parseNum(prow.Proyeccion_Cierre_Mayo) : 0,
                variacion: varMes,
                tendencia: trow ? (trow.Tendencia_Mes || '').trim() : '',
                diagnostico: prow ? (prow.Diagnostico || '').trim() : '',
                salud: srow ? (srow['Semáforo'] || '').trim() : 'Sin Actividad',
            };

            // Populate all week columns
            weekCols.forEach(col => {
                rowObj[col] = trxSemRow ? parseNum(trxSemRow[col]) : 0;
            });

            return rowObj;
        });
    }

    function applyTableFilters() {
        const search = ($('#table-search').value || '').toLowerCase().trim();
        const filtTend = $('#filter-tendencia').value;
        const filtDiag = $('#filter-diagnostico').value;
        const filtSalud = $('#filter-salud').value;
        const filtDir = $('#filter-director') ? $('#filter-director').value : '';

        tableFiltered = tableData.filter(r => {
            if (search && !r.holder.toLowerCase().includes(search)) return false;
            if (filtTend && !r.tendencia.toLowerCase().includes(filtTend.toLowerCase())) return false;
            if (filtDiag && !r.diagnostico.toLowerCase().includes(filtDiag.toLowerCase())) return false;
            if (filtSalud && !r.salud.toLowerCase().includes(filtSalud.toLowerCase())) return false;
            if (filtDir && r.director.toLowerCase() !== filtDir.toLowerCase()) return false;
            return true;
        });

        // Apply sort
        const { key, dir } = tableSort;
        tableFiltered.sort((a, b) => {
            let va = a[key], vb = b[key];
            if (key === 'ranking') {
                const valA = (va === null || va === undefined || isNaN(va)) ? Infinity : va;
                const valB = (vb === null || vb === undefined || isNaN(vb)) ? Infinity : vb;
                if (valA === Infinity && valB === Infinity) return 0;
                if (valA === Infinity) return 1;
                if (valB === Infinity) return -1;
                return dir === 'asc' ? valA - valB : valB - valA;
            }
            if (typeof va === 'string') {
                const cmp = va.localeCompare(vb, 'es');
                return dir === 'asc' ? cmp : -cmp;
            }
            return dir === 'asc' ? va - vb : vb - va;
        });

        tablePage = 0;
        renderTable();
    }

    function renderTable() {
        const tbody = $('#table-body');
        const start = tablePage * TABLE_PAGE_SIZE;
        const page = tableFiltered.slice(start, start + TABLE_PAGE_SIZE);

        tbody.innerHTML = page.map(r => {
            const rankingText = r.ranking != null ? `<span style="color: var(--accent-cyan); font-weight: 600;">#${r.ranking}</span>` : '—';
            
            let weekCells = '';
            weekCols.forEach((col, idx) => {
                const val = r[col];
                let trendClass = '';
                if (idx > 0) {
                    const prevCol = weekCols[idx - 1];
                    const prevVal = r[prevCol];
                    if (val > prevVal) {
                        trendClass = 'class="trend-up"';
                    } else if (val < prevVal) {
                        trendClass = 'class="trend-down"';
                    }
                }
                weekCells += `<td ${trendClass}>${fmtCurrency(val)}</td>`;
            });

            return `<tr data-holder="${escHtml(r.holder)}">
                <td title="${escHtml(r.holder)}">${escHtml(r.holder)}</td>
                <td>${rankingText}</td>
                <td>${escHtml(r.director || '—')}</td>
                <td>${escHtml(r.vendedor || '—')}</td>
                <td>${fmtCurrency(r.abonos)}</td>
                <td>${fmtCurrency(r.promedio)}</td>
                <td>${fmtCurrency(r.proyeccion)}</td>
                <td>${diagnosticoBadge(r.diagnostico, r.promedio, r.proyeccion)}</td>
                <td>${semaforoBadge(r.salud)}</td>
                ${weekCells}
            </tr>`;
        }).join('');

        // Row click
        tbody.querySelectorAll('tr').forEach(tr => {
            tr.addEventListener('click', () => openDetail(tr.dataset.holder));
        });

        // Info
        $('#table-count').textContent = `${tableFiltered.length} comercios`;

        renderPagination();
    }

    function renderPagination() {
        const totalPages = Math.ceil(tableFiltered.length / TABLE_PAGE_SIZE);
        const pag = $('#pagination');
        if (totalPages <= 1) { pag.innerHTML = ''; return; }

        let html = '';
        html += `<button class="${tablePage === 0 ? 'disabled' : ''}" data-page="prev">‹</button>`;

        const maxButtons = 7;
        let startP = Math.max(0, tablePage - 3);
        let endP = Math.min(totalPages - 1, startP + maxButtons - 1);
        if (endP - startP < maxButtons - 1) startP = Math.max(0, endP - maxButtons + 1);

        if (startP > 0) {
            html += `<button data-page="0">1</button>`;
            if (startP > 1) html += `<span class="page-info">…</span>`;
        }
        for (let i = startP; i <= endP; i++) {
            html += `<button class="${i === tablePage ? 'active' : ''}" data-page="${i}">${i + 1}</button>`;
        }
        if (endP < totalPages - 1) {
            if (endP < totalPages - 2) html += `<span class="page-info">…</span>`;
            html += `<button data-page="${totalPages - 1}">${totalPages}</button>`;
        }

        html += `<button class="${tablePage === totalPages - 1 ? 'disabled' : ''}" data-page="next">›</button>`;

        pag.innerHTML = html;
        pag.querySelectorAll('button').forEach(btn => {
            btn.addEventListener('click', () => {
                const p = btn.dataset.page;
                if (p === 'prev') tablePage = Math.max(0, tablePage - 1);
                else if (p === 'next') tablePage = Math.min(totalPages - 1, tablePage + 1);
                else tablePage = parseInt(p);
                renderTable();
            });
        });
    }

    function initTableEvents() {
        $('#table-search').addEventListener('input', debounce(applyTableFilters, 250));
        $('#filter-tendencia').addEventListener('change', applyTableFilters);
        $('#filter-diagnostico').addEventListener('change', applyTableFilters);
        $('#filter-salud').addEventListener('change', applyTableFilters);
        const filterDir = $('#filter-director');
        if (filterDir) {
            filterDir.addEventListener('change', applyTableFilters);
        }
    }

    function getPaymentMethodsHTML(wuziVal, bpVal, speiVal) {
        const hasWuzi = wuziVal > 0;
        const hasBp = bpVal > 0;
        const hasSpei = speiVal > 0;

        let badges = [];

        if (hasWuzi && hasBp) {
            badges.push(`<span class="badge" style="background: rgba(0, 212, 255, 0.08); color: #00d4ff; border: 1px solid rgba(0, 212, 255, 0.2); font-weight: 600; padding: 4px 10px;">💳 TPV: Ambas (Wuzi + BP)</span>`);
        } else if (hasWuzi) {
            badges.push(`<span class="badge" style="background: rgba(0, 212, 255, 0.08); color: #00d4ff; border: 1px solid rgba(0, 212, 255, 0.15); font-weight: 600; padding: 4px 10px;">💳 TPV: Wuzi</span>`);
        } else if (hasBp) {
            badges.push(`<span class="badge" style="background: rgba(123, 47, 247, 0.08); color: #b18eff; border: 1px solid rgba(123, 47, 247, 0.15); font-weight: 600; padding: 4px 10px;">💳 TPV: BP</span>`);
        }

        if (hasSpei) {
            if (hasWuzi || hasBp) {
                badges.push(`<span class="badge" style="background: rgba(255, 68, 102, 0.08); color: #ff4466; border: 1px solid rgba(255, 68, 102, 0.15); font-weight: 600; padding: 4px 10px;">🏦 SPEI (Multicanal)</span>`);
            } else {
                badges.push(`<span class="badge" style="background: rgba(255, 68, 102, 0.08); color: #ff4466; border: 1px solid rgba(255, 68, 102, 0.15); font-weight: 600; padding: 4px 10px;">🏦 Solo SPEI</span>`);
            }
        }

        if (badges.length === 0) {
            return `<span class="badge" style="background: rgba(255, 255, 255, 0.04); color: var(--text-muted); border: 1px solid rgba(255, 255, 255, 0.08); font-weight: 600; padding: 4px 10px;">Sin transacciones</span>`;
        }

        return `<div style="display: flex; gap: 8px; flex-wrap: wrap;">${badges.join('')}</div>`;
    }

    // Badge builders
    function tendenciaBadge(t) {
        const tl = (t || '').toLowerCase();
        if (tl.includes('alza')) return `<span class="badge badge-alza">📈 Alza</span>`;
        if (tl.includes('baja')) return `<span class="badge badge-baja">📉 Baja</span>`;
        if (tl.includes('estable')) return `<span class="badge badge-estable">➖ Estable</span>`;
        return `<span class="badge">${escHtml(t) || '—'}</span>`;
    }
    function diagnosticoBadge(d, promedio, proyeccion) {
        if (!d) return '—';
        const dl = d.toLowerCase();
        
        let pctText = '';
        if (promedio && promedio > 0 && proyeccion !== undefined) {
            const pct = ((proyeccion - promedio) / promedio) * 100;
            const sign = pct >= 0 ? '+' : '';
            pctText = `${sign}${Math.round(pct)}% `;
        }
        
        let cleanText = d.replace(/^[✅⚠️🚨]\s*/, '');
        const textWithPct = `${pctText}${cleanText}`;
        
        let badgeClass = '';
        if (dl.includes('no llegar') || dl.includes('inactivo') || dl.includes('abandono')) {
            badgeClass = 'badge-no';
        } else if (dl.includes('llegar')) {
            badgeClass = 'badge-meta';
        } else if (dl.includes('riesgo') || dl.includes('peligro')) {
            badgeClass = 'badge-riesgo';
        } else if (dl.includes('nuevo')) {
            badgeClass = 'badge-nuevo';
        }
        
        return `<span class="badge ${badgeClass}">${escHtml(textWithPct)}</span>`;
    }
    function semaforoBadge(s) {
        const sl = (s || '').toLowerCase();
        if (sl.includes('constante')) return `<span class="badge badge-salud-constante">Constante ✅</span>`;
        if (sl.includes('intermitente')) return `<span class="badge badge-salud-intermitente">Intermitente ⚠️</span>`;
        if (sl.includes('riesgo') || sl.includes('abandono')) return `<span class="badge badge-salud-riesgo">Riesgo 🚨</span>`;
        if (sl.includes('nuevo')) return `<span class="badge badge-nuevo">Nuevo 🌟</span>`;
        return `<span class="badge badge-salud-sinactividad">Sin Actividad</span>`;
    }

    function populateDirectorFilter() {
        const directors = new Set();
        (DATA.ranking || []).forEach(r => {
            if (r.Director) {
                const dir = r.Director.trim();
                if (dir) directors.add(dir);
            }
        });

        const sortedDirectors = Array.from(directors).sort((a, b) => a.localeCompare(b, 'es'));
        const select = $('#filter-director');
        if (!select) return;

        select.innerHTML = '<option value="">Todos los Directores</option>' +
            sortedDirectors.map(d => `<option value="${escHtml(d)}">${escHtml(d)}</option>`).join('');
    }

    // ══════════════════════════════
    //  SECTION 3: DETAIL
    // ══════════════════════════════
    function openDetail(holder) {
        currentDetailHolder = holder;
        const detailTab = $('#detail-tab');
        detailTab.style.display = '';
        $('#detail-tab-label').textContent = holder.length > 22 ? holder.slice(0, 20) + '…' : holder;
        switchTab('detail');
        renderDetail(holder);
    }

    function renderDetail(holder) {
        const prow = getRow('proyeccion', holder);
        const trow = getRow('tendencias', holder);
        const enrow = getRow('efectoNomina', holder);
        const pdrow = getRow('patronesDia', holder);
        const pqrow = getRow('patronesQuincena', holder);
        const srow = getRow('semaforo', holder);
        const rrow = getRow('ranking', holder);
        const trxSemRow = getRow('trxSem', holder);

        // Title
        $('#detail-title').textContent = holder;

        const wuziVal = prow ? parseNum(prow.Wuzi) : 0;
        const bpVal = prow ? parseNum(prow.BP) : 0;
        const speiVal = prow ? parseNum(prow.SPEI) : 0;
        const paymentMethodsEl = $('#detail-payment-methods');
        if (paymentMethodsEl) {
            paymentMethodsEl.innerHTML = getPaymentMethodsHTML(wuziVal, bpVal, speiVal);
        }

        // Subtitle (Ranking, Director, Vendedor, Context)
        const subEl = $('#detail-subtitle');
        const contextVal = getHolderContext(holder);
        
        const posText = (rrow && rrow.Ranking) ? `Rank #${rrow.Ranking}` : 'Sin Ranking';
        const dirName = (rrow && rrow.Director) || '—';
        const vendName = (rrow && rrow.Vendedor) || '—';
        subEl.innerHTML = `
            <span class="rank-badge">🏆 ${posText}</span>
            <span>👤 Director: <strong>${escHtml(dirName)}</strong></span>
            <span>💼 Vendedor: <strong>${escHtml(vendName)}</strong></span>
            <span class="context-badge">📝 Contexto: <strong>${escHtml(contextVal || 'Sin contexto')}</strong></span>
        `;

        // KPIs
        $('#det-abonos').textContent = prow ? fmtCurrency(parseNum(prow.Abonos_Actuales_Mayo)) : '—';
        $('#det-promedio').textContent = prow ? fmtCurrency(parseNum(prow.Promedio_Mensual_Historico)) : '—';
        $('#det-proyeccion').textContent = prow ? fmtCurrency(parseNum(prow.Proyeccion_Cierre_Mayo)) : '—';

        const diagEl = $('#det-diagnostico');
        if (prow && prow.Diagnostico) {
            const promedio = parseNum(prow.Promedio_Mensual_Historico);
            const proyeccion = parseNum(prow.Proyeccion_Cierre_Mayo);
            diagEl.innerHTML = diagnosticoBadge(prow.Diagnostico, promedio, proyeccion);
        } else {
            diagEl.textContent = '—';
        }

        // Trend badges
        if (trow) {
            const tendMes = $('#det-tend-mes');
            tendMes.innerHTML = tendenciaBadge(trow.Tendencia_Mes || '');
            $('#det-var-mes').textContent = trow.Variacion_Ultimo_Mes != null ? `${parseNum(trow.Variacion_Ultimo_Mes) > 0 ? '+' : ''}${fmt(parseNum(trow.Variacion_Ultimo_Mes))}%` : '—';
            const tendSem = $('#det-tend-sem');
            tendSem.innerHTML = tendenciaBadge(trow.Tendencia_Semana || '');
            $('#det-var-sem').textContent = trow.Variacion_Ultima_Semana != null ? `${parseNum(trow.Variacion_Ultima_Semana) > 0 ? '+' : ''}${fmt(parseNum(trow.Variacion_Ultima_Semana))}%` : '—';
        } else {
            $('#det-tend-mes').textContent = '—';
            $('#det-var-mes').textContent = '—';
            $('#det-tend-sem').textContent = '—';
            $('#det-var-sem').textContent = '—';
        }

        // Payroll effect
        if (enrow) {
            $('#det-nomina').textContent = fmtCurrency(parseNum(enrow.Dias_Nomina));
            $('#det-normal').textContent = fmtCurrency(parseNum(enrow.Dias_Normales));
            const diff = parseNum(enrow.Diferencia_Absoluta);
            const diffEl = $('#det-diff-nomina');
            diffEl.textContent = `${diff >= 0 ? '+' : ''}${fmtCurrency(diff)}`;
            diffEl.style.color = diff >= 0 ? '#00e676' : '#ff4466';
        } else {
            $('#det-nomina').textContent = '—';
            $('#det-normal').textContent = '—';
            $('#det-diff-nomina').textContent = '—';
        }

        // Health stats
        if (srow) {
            $('#det-salud-status').innerHTML = semaforoBadge(srow['Semáforo']);
            $('#det-salud-semanas').textContent = `${srow.Semanas_De_Vida || 0} / ${srow.Semanas_En_Ceros || 0}`;
            $('#det-salud-inactividad').textContent = srow.Porcentaje_Inactividad || '0.0%';
        } else {
            $('#det-salud-status').textContent = '—';
            $('#det-salud-semanas').textContent = '—';
            $('#det-salud-inactividad').textContent = '—';
        }

        // Charts
        renderDetailMensual(trxSemRow);
        renderDetailDiaSemana(pdrow);
        renderDetailQuincena(pqrow);
        renderDetailSemanal(trxSemRow);
    }

    function renderDetailMensual(row) {
        const ctx = $('#detail-chart-mensual').getContext('2d');
        if (charts.detMensual) charts.detMensual.destroy();

        if (!row) {
            const fallbackRow = getRow('abonosMensuales', currentDetailHolder);
            if (fallbackRow) {
                renderDetailMensualOld(fallbackRow, ctx);
                return;
            }
            charts.detMensual = emptyChart(ctx, 'Sin datos mensuales');
            return;
        }

        const monthsRow = DATA.trxSem ? DATA.trxSem[0] : null;
        if (!monthsRow) {
            charts.detMensual = emptyChart(ctx, 'Sin datos mensuales');
            return;
        }

        const monthlySums = {};
        const monthMap = {
            'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
            'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
        };
        const monthNamesAbbr = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

        Object.keys(row).forEach(col => {
            if (col === 'Director' || col === 'Holder') return;
            const monthName = (monthsRow[col] || '').toLowerCase().trim();
            const monthNum = monthMap[monthName];
            if (!monthNum) return;

            const yearMatch = col.match(/\d{4}$/);
            if (!yearMatch) return;
            const year = parseInt(yearMatch[0]);

            const key = `${year}-${String(monthNum).padStart(2, '0')}`;
            const val = parseNum(row[col]);
            monthlySums[key] = (monthlySums[key] || 0) + val;
        });

        const sortedKeys = Object.keys(monthlySums).sort();
        if (sortedKeys.length === 0) {
            charts.detMensual = emptyChart(ctx, 'Sin datos mensuales');
            return;
        }

        const labels = sortedKeys.map(k => {
            const [y, m] = k.split('-');
            return monthNamesAbbr[parseInt(m)] + ' ' + y;
        });
        const values = sortedKeys.map(k => monthlySums[k]);

        const gradient = ctx.createLinearGradient(0, 0, 0, ctx.canvas.height || 280);
        gradient.addColorStop(0, 'rgba(0,212,255,0.3)');
        gradient.addColorStop(1, 'rgba(0,212,255,0.01)');

        charts.detMensual = new Chart(ctx, {
            type: 'line',
            data: {
                labels,
                datasets: [{
                    label: 'Abonos Mensuales',
                    data: values,
                    borderColor: '#00d4ff',
                    backgroundColor: gradient,
                    fill: true,
                    tension: 0.4,
                    pointBackgroundColor: '#00d4ff',
                    pointBorderColor: '#0a0e1a',
                    pointBorderWidth: 2,
                    pointRadius: 5,
                    pointHoverRadius: 7,
                    borderWidth: 2.5,
                }]
            },
            options: chartLineOpts('Abonos'),
        });
    }

    function renderDetailMensualOld(row, ctx) {
        const cols = Object.keys(row).filter(k => k !== 'Holder');
        const labels = cols.map(c => {
            const parts = c.split('-');
            const monthNames = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
            return parts.length === 2 ? monthNames[parseInt(parts[1])] + ' ' + parts[0] : c;
        });
        const values = cols.map(c => parseNum(row[c]));

        const gradient = ctx.createLinearGradient(0, 0, 0, ctx.canvas.height || 280);
        gradient.addColorStop(0, 'rgba(0,212,255,0.3)');
        gradient.addColorStop(1, 'rgba(0,212,255,0.01)');

        charts.detMensual = new Chart(ctx, {
            type: 'line',
            data: {
                labels,
                datasets: [{
                    label: 'Abonos',
                    data: values,
                    borderColor: '#00d4ff',
                    backgroundColor: gradient,
                    fill: true,
                    tension: 0.4,
                    pointBackgroundColor: '#00d4ff',
                    pointBorderColor: '#0a0e1a',
                    pointBorderWidth: 2,
                    pointRadius: 6,
                    pointHoverRadius: 8,
                    borderWidth: 2.5,
                }]
            },
            options: chartLineOpts('Abonos'),
        });
    }

    function renderDetailDiaSemana(row) {
        const ctx = $('#detail-chart-diasemana').getContext('2d');
        if (charts.detDia) charts.detDia.destroy();

        if (!row) {
            charts.detDia = emptyChart(ctx, 'Sin datos de día');
            return;
        }

        // Exclude Monday (Lunes) as weekend terminal cuts skew the scale
        const dayKeys = ['2-Martes', '3-Miércoles', '4-Jueves', '5-Viernes', '6-Sábado', '7-Domingo'];
        const dayLabels = ['Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];
        const values = dayKeys.map(k => parseNum(row[k]));

        const colors = values.map((v, i) => {
            const max = Math.max(...values);
            return v === max ? 'rgba(0,212,255,0.8)' : 'rgba(123,47,247,0.5)';
        });

        charts.detDia = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: dayLabels,
                datasets: [{
                    label: 'Abonos',
                    data: values,
                    backgroundColor: colors,
                    borderRadius: 8,
                    borderSkipped: false,
                    maxBarThickness: 44,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: tooltipOpts({ label: ctx => fmtCurrency(ctx.parsed.y) }),
                },
                scales: scaleOpts(),
            }
        });
    }

    function renderDetailQuincena(row) {
        const ctx = $('#detail-chart-quincena').getContext('2d');
        if (charts.detQuincena) charts.detQuincena.destroy();

        if (!row) {
            charts.detQuincena = emptyChart(ctx, 'Sin datos quincenales');
            return;
        }

        const keys = Object.keys(row).filter(k => k !== 'Holder' && k !== 'Quincena_Fuerte');
        const labels = keys.map(k => k.includes('1ra') ? '1ra Quincena' : '2da Quincena');
        const values = keys.map(k => parseNum(row[k]));

        charts.detQuincena = new Chart(ctx, {
            type: 'pie',
            data: {
                labels,
                datasets: [{
                    data: values,
                    backgroundColor: ['rgba(0,212,255,0.7)', 'rgba(123,47,247,0.7)'],
                    borderColor: 'transparent',
                    hoverOffset: 12,
                    spacing: 3,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { color: '#8a94a8', font: { family: "'Inter'", size: 12 }, padding: 18, usePointStyle: true }
                    },
                    tooltip: tooltipOpts({ label: ctx => fmtCurrency(ctx.parsed) }),
                },
            }
        });
    }

    function renderDetailSemanal(row) {
        const ctx = $('#detail-chart-semanal').getContext('2d');
        if (charts.detSemanal) charts.detSemanal.destroy();

        if (!row) {
            const fallbackRow = getRow('abonosSemanales', currentDetailHolder);
            if (fallbackRow) {
                renderDetailSemanalOld(fallbackRow, ctx);
                return;
            }
            charts.detSemanal = emptyChart(ctx, 'Sin datos semanales');
            return;
        }

        const val2025 = Array(53).fill(null);
        const val2026 = Array(53).fill(null);

        Object.keys(row).forEach(col => {
            if (col === 'Director' || col === 'Holder') return;
            const match = col.match(/Semana (\d+) (\d{4})/);
            if (!match) return;
            const weekNum = parseInt(match[1]);
            const year = parseInt(match[2]);

            const val = parseNum(row[col]);
            if (year === 2025 && weekNum >= 1 && weekNum <= 53) {
                val2025[weekNum - 1] = val;
            } else if (year === 2026 && weekNum >= 1 && weekNum <= 53) {
                val2026[weekNum - 1] = val;
            }
        });

        const labels = Array.from({ length: 53 }, (_, i) => String(i + 1));

        charts.detSemanal = new Chart(ctx, {
            type: 'line',
            data: {
                labels,
                datasets: [
                    {
                        label: '2025',
                        data: val2025,
                        borderColor: '#7b2ff7',
                        backgroundColor: 'transparent',
                        fill: false,
                        tension: 0.35,
                        pointBackgroundColor: '#7b2ff7',
                        pointBorderColor: '#0a0e1a',
                        pointBorderWidth: 1.5,
                        pointRadius: 0,
                        pointHitRadius: 10,
                        pointHoverRadius: 6,
                        borderWidth: 2.5,
                    },
                    {
                        label: '2026',
                        data: val2026,
                        borderColor: '#00d4ff',
                        backgroundColor: 'transparent',
                        fill: false,
                        tension: 0.35,
                        pointBackgroundColor: '#00d4ff',
                        pointBorderColor: '#0a0e1a',
                        pointBorderWidth: 1.5,
                        pointRadius: 0,
                        pointHitRadius: 10,
                        pointHoverRadius: 6,
                        borderWidth: 2.5,
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: {
                            color: '#8a94a8',
                            font: { family: "'Inter'", size: 11 },
                            boxWidth: 12,
                            padding: 10,
                        }
                    },
                    tooltip: tooltipOpts({
                        title: (tooltipItems) => `Semana ${tooltipItems[0].label}`,
                        label: ctx => `${ctx.dataset.label}: ${fmtCurrency(ctx.parsed.y)}`
                    }),
                },
                scales: {
                    x: {
                        ticks: {
                            color: '#8a94a8',
                            font: { family: "'Inter'", size: 10 },
                            autoSkip: true,
                            maxTicksLimit: 12,
                            maxRotation: 0,
                            minRotation: 0
                        },
                        grid: { display: false },
                        border: { display: false },
                    },
                    y: {
                        ticks: {
                            color: '#5a6478',
                            font: { family: "'Inter'", size: 11 },
                            callback: v => fmtCurrency(v)
                        },
                        grid: { color: 'rgba(255,255,255,0.04)' },
                        border: { display: false },
                    }
                },
                interaction: { intersect: false, mode: 'index' },
            }
        });
    }

    function renderDetailSemanalOld(row, ctx) {
        const cols = Object.keys(row).filter(k => k !== 'Holder');
        const labels = cols.map((c, i) => {
            const parts = c.split(' a ');
            if (parts.length === 2) {
                const d = parts[0].split('-');
                return `${d[1]}/${d[2]}`;
            }
            return `S${i + 1}`;
        });
        const values = cols.map(c => parseNum(row[c]));

        const gradient = ctx.createLinearGradient(0, 0, 0, ctx.canvas.height || 280);
        gradient.addColorStop(0, 'rgba(123,47,247,0.25)');
        gradient.addColorStop(1, 'rgba(123,47,247,0.01)');

        charts.detSemanal = new Chart(ctx, {
            type: 'line',
            data: {
                labels,
                datasets: [{
                    label: 'Abonos Semanales (2026)',
                    data: values,
                    borderColor: '#7b2ff7',
                    backgroundColor: gradient,
                    fill: true,
                    tension: 0.35,
                    pointBackgroundColor: '#7b2ff7',
                    pointBorderColor: '#0a0e1a',
                    pointBorderWidth: 2,
                    pointRadius: 0,
                    pointHitRadius: 10,
                    pointHoverRadius: 6,
                    borderWidth: 2.5,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: tooltipOpts({ label: ctx => fmtCurrency(ctx.parsed.y) }),
                },
                scales: {
                    x: {
                        ticks: {
                            color: '#8a94a8',
                            font: { family: "'Inter'", size: 10 },
                            autoSkip: true,
                            maxTicksLimit: 8,
                            maxRotation: 0,
                            minRotation: 0
                        },
                        grid: { display: false },
                        border: { display: false },
                    },
                    y: {
                        ticks: {
                            color: '#5a6478',
                            font: { family: "'Inter'", size: 11 },
                            callback: v => fmtCurrency(v)
                        },
                        grid: { color: 'rgba(255,255,255,0.04)' },
                        border: { display: false },
                    }
                },
                interaction: { intersect: false, mode: 'index' },
            },
        });
    }

    // ── Chart option helpers ──
    function tooltipOpts(callbacks = {}) {
        return {
            backgroundColor: 'rgba(18,22,42,0.95)',
            titleColor: '#e8ecf4', bodyColor: '#8a94a8',
            borderColor: 'rgba(255,255,255,0.08)', borderWidth: 1,
            cornerRadius: 8, padding: 12,
            titleFont: { family: "'Inter'", weight: '600' },
            bodyFont: { family: "'Inter'" },
            callbacks,
        };
    }
    function scaleOpts() {
        return {
            x: {
                ticks: { color: '#8a94a8', font: { family: "'Inter'", size: 11 } },
                grid: { display: false }, border: { display: false },
            },
            y: {
                ticks: { color: '#5a6478', font: { family: "'Inter'", size: 11 }, callback: v => fmtCurrency(v) },
                grid: { color: 'rgba(255,255,255,0.04)' }, border: { display: false },
            }
        };
    }
    function chartLineOpts(lbl) {
        return {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: tooltipOpts({ label: ctx => fmtCurrency(ctx.parsed.y) }),
            },
            scales: scaleOpts(),
            interaction: { intersect: false, mode: 'index' },
        };
    }
    function emptyChart(ctx, msg) {
        return new Chart(ctx, {
            type: 'bar',
            data: { labels: [msg], datasets: [{ data: [0], backgroundColor: 'rgba(255,255,255,0.05)' }] },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false }, tooltip: { enabled: false } },
                scales: { x: { display: false }, y: { display: false } },
            }
        });
    }

    // ══════════════════════════════
    //  GLOBAL SEARCH
    // ══════════════════════════════
    function initGlobalSearch() {
        const input = $('#global-search');
        const dropdown = $('#search-results');

        input.addEventListener('input', debounce(() => {
            const q = (input.value || '').toLowerCase().trim();
            if (q.length < 2) { dropdown.classList.remove('open'); return; }
            const matches = getActiveHolders().filter(h => h.toLowerCase().includes(q)).slice(0, 20);
            if (matches.length === 0) {
                dropdown.innerHTML = '<div class="search-item">Sin resultados</div>';
            } else {
                dropdown.innerHTML = matches.map(h => `<div class="search-item" data-holder="${escHtml(h)}">${highlightMatch(h, q)}</div>`).join('');
                dropdown.querySelectorAll('.search-item[data-holder]').forEach(item => {
                    item.addEventListener('click', () => {
                        openDetail(item.dataset.holder);
                        dropdown.classList.remove('open');
                        input.value = '';
                    });
                });
            }
            dropdown.classList.add('open');
        }, 200));

        // Close on click outside
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.header-search')) dropdown.classList.remove('open');
        });
    }

    function highlightMatch(text, query) {
        const idx = text.toLowerCase().indexOf(query);
        if (idx === -1) return escHtml(text);
        return escHtml(text.slice(0, idx)) +
            '<strong style="color:#00d4ff">' + escHtml(text.slice(idx, idx + query.length)) + '</strong>' +
            escHtml(text.slice(idx + query.length));
    }

    // ── Misc ──
    function escHtml(s) {
        const div = document.createElement('div');
        div.textContent = s || '';
        return div.innerHTML;
    }
    function debounce(fn, ms) {
        let t;
        return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
    }

    // ══════════════════════════════
    //  BACK BUTTON
    // ══════════════════════════════
    function initBackButton() {
        $('#btn-back').addEventListener('click', () => {
            switchTab('table');
        });
    }

    // ══════════════════════════════
    //  TOP 300 TOGGLE
    // ══════════════════════════════
    function initTop300Toggle() {
        const toggle = $('#top300-toggle');
        if (!toggle) return;

        toggle.addEventListener('change', () => {
            showOnlyTop300 = toggle.checked;

            // Re-render everything that depends on active holders
            renderOverview();
            buildTableData();
            applyTableFilters();

            // If a client is selected but not in active list, close details
            if (currentDetailHolder && showOnlyTop300) {
                const activeSet = new Set(getActiveHolders());
                if (!activeSet.has(currentDetailHolder)) {
                    $('#detail-tab').style.display = 'none';
                    switchTab('overview');
                    currentDetailHolder = null;
                }
            }
        });
    }



    // ══════════════════════════════
    //  INIT
    // ══════════════════════════════
    async function init() {
        try {
            await loadAllData();
            buildIndexes();
            calculateAbonosTotales();
            computeTop300();
            populateDirectorFilter();

            renderOverview();

            buildTableHeader();
            buildTableData();
            initTableEvents();
            applyTableFilters();

            initTabs();
            initGlobalSearch();
            initBackButton();
            initTop300Toggle();
            initExportButtons();

            // Hide loader
            setTimeout(() => {
                $('#loading-overlay').classList.add('hidden');
            }, 400);
        } catch (err) {
            console.error('Init error:', err);
            $('#loading-status').textContent = 'Error al cargar los datos. Revise la consola.';
        }
    }

    function showToast(msg) {
        let toast = $('#toast-notification');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'toast-notification';
            document.body.appendChild(toast);
        }
        toast.textContent = msg;
        toast.classList.add('show');
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }

    function downloadTextFile(text, filename) {
        const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);
    }

    function getExportText() {
        if (tableFiltered.length === 0) return null;

        let txt = `*REPORTE DE COMERCIOS — PROYECCIÓN*\n`;
        txt += `Fecha: ${new Date().toLocaleDateString('es-MX')}\n`;
        
        const searchVal = $('#table-search').value.trim();
        const tendVal = $('#filter-tendencia').value;
        const diagVal = $('#filter-diagnostico').value;
        const saludVal = $('#filter-salud').value;
        const dirVal = $('#filter-director') ? $('#filter-director').value : '';

        const activeFilters = [];
        if (searchVal) activeFilters.push(`Búsqueda: "${searchVal}"`);
        if (tendVal) activeFilters.push(`Tendencia: ${tendVal}`);
        if (diagVal) activeFilters.push(`Diagnóstico: ${diagVal}`);
        if (saludVal) activeFilters.push(`Salud: ${saludVal}`);
        if (dirVal) activeFilters.push(`Director: ${dirVal}`);
        if (showOnlyTop300) activeFilters.push(`Top 300`);

        txt += `Filtros: ${activeFilters.length > 0 ? activeFilters.join(', ') : 'Ninguno'}\n`;
        txt += `Total: ${tableFiltered.length} comercios\n`;
        txt += `---------------------------------------------\n\n`;

        tableFiltered.forEach((r, idx) => {
            const rankText = r.ranking != null ? `#${r.ranking}` : '—';
            
            // Calculate dynamic percentage
            let pctText = '';
            if (r.promedio && r.promedio > 0 && r.proyeccion !== undefined) {
                const pct = ((r.proyeccion - r.promedio) / r.promedio) * 100;
                const sign = pct >= 0 ? '+' : '';
                pctText = `${sign}${Math.round(pct)}% `;
            }
            const cleanDiag = r.diagnostico ? r.diagnostico.replace(/^[✅⚠️🚨]\s*/, '') : '—';
            const fullDiag = `${pctText}${cleanDiag}`;

            txt += `*${idx + 1}. ${r.holder}* (Rank: ${rankText})\n`;
            txt += `• Dir: ${r.director || '—'} | Vend: ${r.vendedor || '—'}\n`;
            txt += `• Abonos Mayo: ${fmtCurrency(r.abonos)} | Prom. Hist: ${fmtCurrency(r.promedio)} | Proy. Cierre: ${fmtCurrency(r.proyeccion)}\n`;
            txt += `• Salud: ${r.salud || '—'} | Diagnóstico: *${fullDiag}*\n`;
            txt += `---------------------------------------------\n`;
        });
        return txt;
    }

    function initExportButtons() {
        const copyBtn = $('#btn-export-copy');
        const dlBtn = $('#btn-export-download');

        if (copyBtn) {
            copyBtn.addEventListener('click', () => {
                const text = getExportText();
                if (!text) {
                    showToast('No hay comercios en la lista para copiar.');
                    return;
                }
                navigator.clipboard.writeText(text).then(() => {
                    showToast('📋 ¡Reporte copiado para WhatsApp!');
                }).catch(err => {
                    console.error('Error al copiar:', err);
                    showToast('Error al copiar. Descargando como archivo...');
                    downloadTextFile(text, 'reporte_proyeccion.txt');
                });
            });
        }

        if (dlBtn) {
            dlBtn.addEventListener('click', () => {
                const text = getExportText();
                if (!text) {
                    showToast('No hay comercios en la lista para descargar.');
                    return;
                }
                downloadTextFile(text, 'reporte_proyeccion.txt');
                showToast('📥 ¡Archivo de texto descargado!');
            });
        }
    }

    // Start
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
