// ========================
// multiCAD-MCP Dashboard
// ========================

let cadStatus = { connected: false, cad_type: 'None', drawings: [], current_drawing: 'None' };
let layers = [];
let blocks = [];
let entitiesByType = {}; // Stores entities per type
let entityPagination = {}; // Stores pagination state per type

const entitiesPerPage = 500;

const PANEL_TITLES = {
    overview: 'Resumen',
    entities: 'Entidades',
    layers: 'Capas',
    blocks: 'Bloques',
    logs: 'Consola'
};

let logPollInterval = null;
let logSeq = 0;

document.addEventListener('DOMContentLoaded', () => {
    setupNav();
    setupEventListeners();
    refreshData();
    // Pre-load log buffer so Consola tab shows history immediately on first visit
    fetchLogs();
});

// ========================
// Navigation
// ========================

function setupNav() {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', () => {
            switchPanel(item.getAttribute('data-panel'));
        });
    });
}

function switchPanel(panelId) {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.toggle('active', item.getAttribute('data-panel') === panelId);
    });
    document.querySelectorAll('.panel').forEach(panel => {
        panel.classList.toggle('active', panel.id === `panel-${panelId}`);
    });
    const titleEl = document.getElementById('panel-title');
    if (titleEl) titleEl.textContent = PANEL_TITLES[panelId] || panelId;

    if (panelId === 'logs') {
        startLogPolling();
    } else {
        stopLogPolling();
    }
}

// ========================
// Event Listeners
// ========================

function setupEventListeners() {
    document.getElementById('refresh-all')?.addEventListener('click', handleManualRefresh);
    document.getElementById('export-all')?.addEventListener('click', handleExport);
    document.getElementById('drawing-selector')?.addEventListener('change', handleDrawingSwitch);

    document.getElementById('filter-entities')?.addEventListener('input', renderEntities);
    document.getElementById('filter-entities-layer')?.addEventListener('change', renderEntities);
    document.getElementById('filter-layers')?.addEventListener('input', renderLayers);
    document.getElementById('filter-blocks')?.addEventListener('input', renderBlocks);

    document.getElementById('expand-all-entities')?.addEventListener('click', () => toggleAllAccordions('entities-container', true));
    document.getElementById('collapse-all-entities')?.addEventListener('click', () => toggleAllAccordions('entities-container', false));
    document.getElementById('expand-all-blocks')?.addEventListener('click', () => toggleAllAccordions('blocks-container', true));
    document.getElementById('collapse-all-blocks')?.addEventListener('click', () => toggleAllAccordions('blocks-container', false));
}

// ========================
// Data
// ========================

async function handleManualRefresh() {
    const btn = document.getElementById('refresh-all');
    if (!btn || btn.classList.contains('loading')) return;
    btn.classList.add('loading');

    // Deshabilitar selector de dibujos mientras carga
    const selDropdown = document.getElementById('drawing-selector');
    if (selDropdown) selDropdown.disabled = true;

    try {
        const res = await fetch('/api/cad/refresh', { method: 'POST' });
        const result = await res.json();
        if (result.success) {
            await refreshData();
            showToast('Datos actualizados', 'success');
        } else {
            showToast('Error de refresco', 'error');
        }
    } catch {
        showToast('Error de conexión', 'error');
    } finally {
        setTimeout(() => btn.classList.remove('loading'), 400);
    }
}

async function handleExport() {
    const btn = document.getElementById('export-all');
    if (!btn || btn.classList.contains('loading')) return;
    btn.classList.add('loading');

    try {
        const res = await fetch('/api/cad/export', { method: 'POST' });
        const result = await res.json();
        if (result.success) {
            showToast('Exportación completada con éxito', 'success');
        } else {
            showToast('Error al exportar: ' + (result.detail || 'Desconocido'), 'error');
        }
    } catch {
        showToast('Error de conexión al exportar', 'error');
    } finally {
        setTimeout(() => btn.classList.remove('loading'), 400);
    }
}

async function handleDrawingSwitch(e) {
    const drawingName = e.target.value;
    if (!drawingName) return;

    const selDropdown = document.getElementById('drawing-selector');
    if (selDropdown) selDropdown.disabled = true;

    try {
        const res = await fetch('/api/cad/switch_drawing', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ drawing_name: drawingName })
        });
        const result = await res.json();

        if (result.success) {
            showToast(`Cambio a ${getFilename(drawingName)}`, 'success');
            setTimeout(handleManualRefresh, 1000); // Give CAD a moment to switch and cache to mark dirty before refreshing
        } else {
            showToast('Error: No se pudo cambiar el dibujo', 'error');
            // Revertir a current_drawing anterior
            populateDrawingSelector(cadStatus.connected);
        }
    } catch {
        showToast('Error de red al cambiar dibujo', 'error');
        populateDrawingSelector(cadStatus.connected);
    }
}

async function refreshData() {
    try {
        const res = await fetch('/api/cad/status');
        const data = await res.json();
        if (data.success) {
            cadStatus = data.status;
            if (cadStatus.connected) {
                await fetchDetails();
            } else {
                layers = []; blocks = []; entities = [];
            }
            updateUI();
        }
    } catch (err) {
        console.error('Refresh error:', err);
    }
}

async function fetchDetails() {
    try {
        const [lR, bR] = await Promise.all([
            fetch('/api/cad/layers'),
            fetch('/api/cad/blocks')
        ]);
        const [lD, bD] = await Promise.all([lR.json(), bR.json()]);
        layers = lD.success ? lD.layers : [];
        blocks = bD.success ? bD.blocks : [];

        // Reset per-type data on refresh
        entitiesByType = {};
        entityPagination = {};
    } catch (err) {
        console.error('Details error:', err);
    }
}

async function fetchEntitiesPage(type, page) {
    try {
        const res = await fetch(`/api/cad/entities?type=${encodeURIComponent(type)}&page=${page}&limit=${entitiesPerPage}`);
        const data = await res.json();

        if (data.success) {
            entitiesByType[type] = data.entities;
            entityPagination[type] = {
                page: data.pagination.page,
                total: data.pagination.total,
                totalPages: data.pagination.total_pages
            };

            // Re-render the specific section content
            const contentEl = document.querySelector(`.accordion-section[data-type="${type}"] .accordion-body`);
            if (contentEl) {
                contentEl.innerHTML = '';
                contentEl.appendChild(makeTypePaginationHeader(type));
                contentEl.appendChild(makeEntityTable(data.entities));
            }
        }
    } catch (err) {
        console.error('Error fetching entities page:', err);
    }
}

function makeTypePaginationHeader(type) {
    const pag = entityPagination[type] || { page: 1, totalPages: 1 };
    const div = document.createElement('div');
    div.className = 'type-pagination';
    div.innerHTML = `
        <button class="btn-sm" ${pag.page <= 1 ? 'disabled' : ''} onclick="fetchEntitiesPage('${type}', ${pag.page - 1})">◀</button>
        <span class="page-info">Pág. ${pag.page} de ${pag.totalPages}</span>
        <button class="btn-sm" ${pag.page >= pag.totalPages ? 'disabled' : ''} onclick="fetchEntitiesPage('${type}', ${pag.page + 1})">▶</button>
    `;
    return div;
}

// ========================
// UI Update
// ========================

function updateUI() {
    const c = cadStatus.connected;

    // Sidebar status
    const dot = document.getElementById('status-dot');
    const statusText = document.getElementById('status-text');
    if (dot) dot.className = `status-dot${c ? ' online' : ''}`;
    if (statusText) statusText.textContent = c ? 'Conectado' : 'Desconectado';

    // Overview drawing label (formerly in header)
    const headerDrawing = document.getElementById('header-drawing');
    if (headerDrawing) {
        headerDrawing.textContent = c
            ? (cadStatus.current_drawing || 'Sin dibujo')
            : 'Sin conexión';
        headerDrawing.title = cadStatus.current_drawing || '';
    }

    // Topbar drawing selector
    populateDrawingSelector(c);

    // Stats bar
    setText('cad-type-display', cadStatus.cad_type || '—');
    setText('entities-count', cadStatus.total_entities || 0); // Always use total_entities for overview
    setText('layers-count', layers.length);
    setText('blocks-count', blocks.length);

    // Panel count pills
    setText('entities-panel-count', `${cadStatus.total_entities || 0} entidades totales`);
    setText('layers-panel-count', `${layers.length} capas`);
    setText('blocks-panel-count', `${blocks.length} bloques`);

    populateLayerFilter();
    renderOverview();
    renderEntities();
    renderLayers();
    renderBlocks();
}

function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function getFilename(path) {
    if (!path) return null;
    return path.split(/[/\\]/).pop() || path;
}

function populateDrawingSelector(isConnected) {
    const sel = document.getElementById('drawing-selector');
    if (!sel) return;

    sel.innerHTML = '';
    sel.disabled = !isConnected;

    if (!isConnected) {
        sel.innerHTML = '<option value="">Sin conexión</option>';
        return;
    }

    if (!cadStatus.drawings || cadStatus.drawings.length === 0) {
        sel.innerHTML = '<option value="">Sin dibujos</option>';
        sel.disabled = true;
        return;
    }

    cadStatus.drawings.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = getFilename(d);
        if (d === cadStatus.current_drawing) opt.selected = true;
        sel.appendChild(opt);
    });
}

function populateLayerFilter() {
    const sel = document.getElementById('filter-entities-layer');
    if (!sel) return;
    const cur = sel.value;
    sel.innerHTML = '<option value="">Todas las capas</option>';
    // This filter now needs to work across all fetched entities, or be removed if not feasible.
    // For now, it will only filter what's currently loaded in entitiesByType.
    // A more robust solution would involve fetching entities by layer from the backend.
    const allLoadedEntities = Object.values(entitiesByType).flat();
    [...new Set(allLoadedEntities.map(e => e.Layer).filter(Boolean))].sort().forEach(n => {
        const o = document.createElement('option');
        o.value = n; o.textContent = n;
        sel.appendChild(o);
    });
    sel.value = cur;
}

// ========================
// Overview Renderers
// ========================

function renderOverview() {
    renderOverviewEntities();
    renderOverviewLayers();
    renderOverviewBlocks();
}

/**
 * Centralized entity classification for consistent naming.
 * @param {string} typeStr - Internal object type (e.g., 'AcDbPolyline').
 * @returns {string} Friendly name.
 */
function classifyEntityType(typeStr) {
    if (!typeStr) return 'Desconocido';
    const s = String(typeStr).toUpperCase();
    if (s.includes('POLYLINE') || s.includes('LWPOLYLINE')) return 'Polilínea';
    if (s.includes('CIRCLE')) return 'Círculo';
    if (s.includes('ARC')) return 'Arco';
    if (s.includes('LINE')) return 'Línea';
    if (s.includes('TEXT') || s.includes('MTEXT')) return 'Texto';
    if (s.includes('INSERT') || s.includes('BLOCKREFERENCE') || s.includes('BLOCK')) return 'Bloque';
    if (s.includes('DIMENSION')) return 'Cota';
    if (s.includes('SPLINE')) return 'Spline';
    if (s.includes('POINT')) return 'Punto';
    if (s.includes('HATCH')) return 'Sombreado';
    return typeStr.replace('AcDb', '');
}

function renderOverviewEntities() {
    const container = document.getElementById('overview-entities');
    if (!container) return;

    // Use actual counts from backend
    const counts = cadStatus.entity_counts || {};
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);

    if (sorted.length === 0) {
        container.innerHTML = '<div class="empty-state">Sin entidades</div>';
    } else {
        container.innerHTML = sorted.map(([type, count]) => `
            <div class="ov-row">
                <span class="ov-name" title="${esc(classifyEntityType(type))}">${esc(classifyEntityType(type))}</span>
                <span class="ov-count">${fmtNum(count)}</span>
            </div>`).join('');
    }
}

const ICON_EYE = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="width:12px;height:12px;"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;
const ICON_EYE_OFF = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="width:12px;height:12px;"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path><line x1="1" y1="1" x2="23" y2="23"></line></svg>`;

function renderOverviewLayers() {
    const container = document.getElementById('overview-layers');
    if (!container) return;
    if (layers.length === 0) {
        container.innerHTML = '<div class="empty-state">Sin capas</div>';
        return;
    }
    container.innerHTML = layers.map(l => {
        const ch = getColorHex(l.Color);
        const visIcon = l.IsVisible ? ICON_EYE : ICON_EYE_OFF;
        const visTitle = l.IsVisible ? 'Visible' : 'Oculta';

        return `<div class="ov-row">
            <span class="ov-dot" style="background:${ch}"></span>
            <span class="ov-name" title="${esc(l.Name)}">${esc(l.Name)}</span>
            <span class="vis-icon" title="${visTitle}">${visIcon}</span>
            <span class="ov-count" title="Entidades">${l.ObjectCount || 0} <small>ent.</small></span>
        </div>`;
    }).join('');
}

function renderOverviewBlocks() {
    const container = document.getElementById('overview-blocks');
    if (!container) return;
    if (blocks.length === 0) {
        container.innerHTML = '<div class="empty-state">Sin bloques</div>';
        return;
    }
    container.innerHTML = blocks.map(b =>
        `<div class="ov-row">
            <span class="ov-name" title="${esc(b.Name)}">${esc(b.Name)}</span>
            <span class="ov-count" title="Instancias">${b.Count || 0} <small>ins.</small></span>
        </div>`
    ).join('');
}

// ========================
// Panel Renderers
// ========================

function renderEntities() {
    const container = document.getElementById('entities-container');
    if (!container) return;

    const filterText = (document.getElementById('filter-entities')?.value || '').toLowerCase();
    const filterLayer = (document.getElementById('filter-entities-layer')?.value || '');

    // We use actual counts from backend as the categories
    const counts = cadStatus.entity_counts || {};

    // Filter the categories themselves? Or the entities inside?
    // Since we lazy load, we can only filter the visible categories/accordions here.
    // If the category name doesn't match filterText, we skip it.

    container.innerHTML = '';

    // Group the counts by friendly name to avoid duplicate sections
    // (e.g., "Line" and "Línea" should be one accordion)
    const friendlyGroups = {};
    Object.entries(counts).forEach(([type, count]) => {
        const friendly = classifyEntityType(type);
        if (!friendlyGroups[friendly]) {
            friendlyGroups[friendly] = { count: 0, rawTypes: [] };
        }
        friendlyGroups[friendly].count += count;
        friendlyGroups[friendly].rawTypes.push(type);
    });

    const sortedFriendly = Object.keys(friendlyGroups).sort();

    sortedFriendly.forEach(friendlyType => {
        if (filterText && !friendlyType.toLowerCase().includes(filterText)) return;

        const data = friendlyGroups[friendlyType];
        // Use the first rawType as the primary identifier for the API
        const primaryRawType = data.rawTypes[0];

        // Use a wrapper to trigger fetch on first open
        const section = createAccordionSection(friendlyType, `${data.count}`, null, (isOpen) => {
            if (isOpen && !entitiesByType[primaryRawType]) {
                const content = section.querySelector('.accordion-body');
                content.innerHTML = '<div class="empty-state">Cargando datos...</div>';
                fetchEntitiesPage(primaryRawType, 1);
            } else if (isOpen) {
                // Refresh table if already loaded (sync pagination)
                const content = section.querySelector('.accordion-body');
                content.innerHTML = '';
                content.appendChild(makeTypePaginationHeader(primaryRawType));
                content.appendChild(makeEntityTable(entitiesByType[primaryRawType]));
            }
        });

        section.setAttribute('data-type', primaryRawType);
        container.appendChild(section);
    });
}

function makeEntityTable(items) {
    if (!items || !Array.isArray(items) || items.length === 0) {
        const div = document.createElement('div');
        div.className = 'empty-state';
        div.textContent = 'No hay entidades de este tipo en la página actual.';
        return div;
    }

    const headers = ['Handle', 'Capa', 'Color', 'Longitud / Radio', 'Área', 'Perímetro'];
    const rows = items.map(e => {
        const friendlyType = classifyEntityType(e.ObjectType);
        const isCircleOrArc = friendlyType === 'Círculo' || friendlyType === 'Arco';

        // Resolve "ByLayer" color
        let effectiveColor = e.Color;
        let colorText = String(e.Color || '');
        if (String(effectiveColor).toLowerCase() === 'bylayer' || effectiveColor === 256) {
            const entLayerNorm = String(e.Layer || '').trim().toLowerCase();
            const layerObj = layers.find(l => String(l.Name || '').trim().toLowerCase() === entLayerNorm);

            effectiveColor = (layerObj && layerObj.Color && String(layerObj.Color).toLowerCase() !== 'bylayer')
                ? layerObj.Color : 'white';
            colorText = `ByLayer (${effectiveColor})`;
        }
        const ch = getColorHex(effectiveColor);

        const mainVal = isCircleOrArc ? e.Radius : e.Length;
        const perimVal = isCircleOrArc ? e.Circumference : null;

        return [
            `<td class="mono">${esc(e.Handle || '')}</td>`,
            `<td>${esc(e.Layer || '')}</td>`,
            `<td><span class="color-dot" style="background:${ch}"></span>${esc(colorText)}</td>`,
            `<td class="num">${fmtNum(mainVal)}</td>`,
            `<td class="num">${fmtNum(e.Area)}</td>`,
            `<td class="num">${perimVal !== null ? fmtNum(perimVal) : '—'}</td>`
        ];
    });

    return createTable(headers, rows, 'ent-table entities');
}

function renderLayers() {
    const container = document.getElementById('layers-container');
    if (!container) return;
    const ft = (document.getElementById('filter-layers')?.value || '').toLowerCase();
    const filtered = layers.filter(l => (l.Name || '').toLowerCase().includes(ft));

    if (filtered.length === 0) {
        container.innerHTML = '<div class="empty-state">No hay datos de capas disponibles</div>';
        return;
    }

    container.innerHTML = '';
    const headers = ['Capa', 'Color', 'Estado', 'Entidades'];
    const rows = filtered.map(l => {
        const ch = getColorHex(l.Color);
        return [
            `<td><strong>${esc(l.Name)}</strong></td>`,
            `<td><span class="color-dot" style="background:${ch}"></span>${esc(String(l.Color || ''))}</td>`,
            `<td><span class="vis-pill ${l.IsVisible ? 'on' : 'off'}">${l.IsVisible ? 'Visible' : 'Oculta'}</span></td>`,
            `<td class="num">${l.ObjectCount || 0}</td>`
        ];
    });

    container.appendChild(createTable(headers, rows, 'ent-table layers'));
}

function renderBlocks() {
    const container = document.getElementById('blocks-container');
    if (!container) return;
    const ft = (document.getElementById('filter-blocks')?.value || '').toLowerCase();
    const filtered = blocks.filter(b => (b.Name || '').toLowerCase().includes(ft));

    if (filtered.length === 0) {
        container.innerHTML = '<div class="empty-state">No hay bloques disponibles</div>';
        return;
    }

    container.innerHTML = '';
    filtered.forEach(b => {
        const body = document.createElement('div');
        body.className = 'block-detail';
        body.innerHTML = `<strong>Instancias:</strong> ${b.Count || 0} &middot; <strong>Entidades:</strong> ${b.ObjectCount || 0}`;
        container.appendChild(createAccordionSection(b.Name, `${b.Count || 0}×`, body));
    });
}

// ========================
// Components
// ========================

function createAccordionSection(title, badge, contentEl, onToggle = null) {
    const section = document.createElement('div');
    section.className = 'accordion-section';
    section.innerHTML = `
        <div class="accordion-header">
            <span class="accordion-icon">▶</span>
            <span class="accordion-title">${esc(title)}</span>
            <span class="accordion-badge">${badge}</span>
        </div>
        <div class="accordion-body"></div>
    `;

    const header = section.querySelector('.accordion-header');
    const body = section.querySelector('.accordion-body');

    if (contentEl) {
        body.appendChild(contentEl);
    }

    header.addEventListener('click', () => {
        const isOpen = section.classList.toggle('open');
        body.classList.toggle('open', isOpen);
        header.classList.toggle('open', isOpen);

        if (onToggle) onToggle(isOpen);
    });

    return section;
}

function toggleAllAccordions(id, expand) {
    const c = document.getElementById(id);
    if (!c) return;
    c.querySelectorAll('.accordion-section').forEach(s => {
        const header = s.querySelector('.accordion-header');
        const body = s.querySelector('.accordion-body');

        const isCurrentOpen = s.classList.contains('open');
        if (isCurrentOpen === expand) return; // No change needed

        s.classList.toggle('open', expand);
        if (body) body.classList.toggle('open', expand);
        if (header) header.classList.toggle('open', expand);

        // If expanding and data not loaded, trigger fetch (via attribute/manual)
        if (expand) {
            const type = s.getAttribute('data-type');
            if (type && !entitiesByType[type]) {
                if (body) body.innerHTML = '<div class="empty-state">Cargando datos...</div>';
                fetchEntitiesPage(type, 1);
            }
        }
    });
}

// ========================
// Log Polling
// ========================

function startLogPolling() {
    if (logPollInterval) return;
    fetchLogs();
    logPollInterval = setInterval(fetchLogs, 3000);
}

function stopLogPolling() {
    clearInterval(logPollInterval);
    logPollInterval = null;
}

async function fetchLogs() {
    try {
        const res = await fetch(`/api/logs?since=${logSeq}`);
        const data = await res.json();
        if (!data.success || data.entries.length === 0) return;

        const container = document.getElementById('console-output');
        if (!container) return;

        data.entries.forEach(entry => {
            logSeq = Math.max(logSeq, entry.seq);
            const el = document.createElement('div');
            el.className = `log-entry ${logLevelClass(entry.level)}`;
            el.textContent = `[${entry.time}] ${entry.level.padEnd(5)}  ${entry.msg}`;
            container.appendChild(el);
        });
        container.scrollTop = container.scrollHeight;
    } catch (err) {
        console.error('Log fetch error:', err);
    }
}

function logLevelClass(level) {
    if (level === 'ERROR' || level === 'CRITICAL') return 'error';
    if (level === 'WARNING') return 'warning';
    return 'info';
}

// ========================
// Utils
// ========================

function addLog(msg, type = 'info') {
    const c = document.getElementById('console-output');
    if (!c) return;
    const e = document.createElement('div');
    e.className = `log-entry ${type}`;
    e.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
    c.appendChild(e);
    c.scrollTop = c.scrollHeight;
}

function showToast(msg, type = 'success') {
    const t = document.getElementById('message-box');
    if (!t) return;
    t.textContent = msg;
    t.className = 'toast show';
    t.style.backgroundColor = type === 'success' ? '#10b981' : '#ef4444';
    setTimeout(() => { t.className = 'toast'; }, 2500);
}

function esc(text) {
    const d = document.createElement('div');
    d.textContent = text;
    return d.innerHTML;
}

function fmtNum(val) {
    if (val === null || val === undefined || isNaN(val)) return '—';
    return Number(val).toLocaleString('es-ES', {
        minimumFractionDigits: 3,
        maximumFractionDigits: 3
    });
}

function getColorHex(color) {
    if (!color) return '#475569';
    const c = String(color).toLowerCase();

    // Use ACI_PALETTE from aci_colors.js
    if (typeof ACI_PALETTE !== 'undefined' && ACI_PALETTE[c]) {
        return ACI_PALETTE[c];
    }

    return '#475569'; // Fallback gray
}

/**
 * Creates a standard styled table.
 * @param {string[]} headers - Array of header labels.
 * @param {string[][]} rows - Array of row data (HTML strings). Each row is an array of cells.
 * @param {string} [tableClass='ent-table'] - CSS class for the table.
 * @returns {HTMLTableElement} 
 */
function createTable(headers, rows, tableClass = 'ent-table') {
    const t = document.createElement('table');
    t.className = tableClass;

    const thead = headers.map(h => {
        // Detect if it should be a numeric column based on name or presence of class
        const isNum = h.toLowerCase().includes('área') ||
            h.toLowerCase().includes('longitud') ||
            h.toLowerCase().includes('radio') ||
            h.toLowerCase().includes('perímetro') ||
            h.toLowerCase().includes('entidades') ||
            h.toLowerCase().includes('instancias');
        return `<th class="${isNum ? 'num' : ''}">${h}</th>`;
    }).join('');

    t.innerHTML = `<thead><tr>${thead}</tr></thead><tbody></tbody>`;
    const tb = t.querySelector('tbody');

    rows.forEach(rowData => {
        const r = document.createElement('tr');
        r.innerHTML = rowData.join('');
        tb.appendChild(r);
    });

    return t;
}

