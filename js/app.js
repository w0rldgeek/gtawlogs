/**
 * ============================================================
 * Актуализация номеров абонентов
 * База данных = сам .xlsx файл (без хостинга и сервера)
 * ============================================================
 * Приложение читает данные ПРЯМО из выбранного Excel-файла и
 * пишет изменения обратно в него же (File System Access API).
 * Файл портативен: откройте его на другом компьютере — увидите
 * все записи. Данные разложены по месяцам, каждый своего цвета.
 *
 * Браузеры без File System Access (Firefox/Safari): открытие
 * файла через выбор + сохранение кнопкой «Сохранить файл».
 * ============================================================
 */

(function () {
    'use strict';

    const MONTH_NAMES = [
        'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
        'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь',
    ];
    const MONTH_COLORS = [
        '60A5FA', '22D3EE', '34D399', 'A3E635', 'FACC15', 'FB923C',
        'F87171', 'F472B6', 'E879F9', 'C084FC', '818CF8', '38BDF8',
    ];
    const CACHE_KEY = 'subscribers_cache_v2';

    const supportsFS = typeof window.showOpenFilePicker === 'function';

    // Состояние: данные в памяти = рабочая копия базы (файла)
    let rows = [];
    let fileHandle = null;
    let fileName = null;

    // ============================================================
    // ЧТЕНИЕ / ЗАПИСЬ EXCEL (ExcelJS)
    // ============================================================
    async function readRowsFromBlob(blob) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await blob.arrayBuffer());
        const out = [];
        wb.eachSheet(ws => {
            ws.eachRow((row, rowNumber) => {
                if (rowNumber < 3) return; // 1 — заголовок месяца, 2 — шапка
                const address = String(cellText(row.getCell(2).value)).trim();
                const login = String(cellText(row.getCell(3).value)).trim();
                const phone = String(cellText(row.getCell(4).value)).trim();
                if (!address && !login && !phone) return;
                out.push({
                    id: uid(),
                    date: cellDate(row.getCell(1).value),
                    address, login, phone,
                });
            });
        });
        return out;
    }

    async function buildBlob() {
        const wb = new ExcelJS.Workbook();

        const groups = new Map();
        for (const s of rows) {
            const d = parseISO(s.date) || new Date();
            const key = `${d.getUTCFullYear()}-${String(d.getUTCMonth()).padStart(2, '0')}`;
            if (!groups.has(key)) groups.set(key, { y: d.getUTCFullYear(), m: d.getUTCMonth(), items: [] });
            groups.get(key).items.push(s);
        }
        if (groups.size === 0) {
            const now = new Date();
            groups.set('e', { y: now.getUTCFullYear(), m: now.getUTCMonth(), items: [] });
        }

        for (const key of Array.from(groups.keys()).sort()) {
            const g = groups.get(key);
            const argb = 'FF' + MONTH_COLORS[g.m];
            const ws = wb.addWorksheet(`${MONTH_NAMES[g.m]} ${g.y}`, {
                properties: { tabColor: { argb } },
                views: [{ state: 'frozen', ySplit: 2 }],
            });
            ws.columns = [{ width: 14 }, { width: 38 }, { width: 20 }, { width: 20 }];

            ws.mergeCells('A1:D1');
            const title = ws.getCell('A1');
            title.value = `${MONTH_NAMES[g.m]} ${g.y}`;
            title.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
            title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
            title.alignment = { horizontal: 'center', vertical: 'middle' };
            ws.getRow(1).height = 24;

            const header = ws.getRow(2);
            ['Дата', 'Адрес', 'Логин', 'Номер телефона'].forEach((h, i) => {
                const c = header.getCell(i + 1);
                c.value = h;
                c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
                c.alignment = { horizontal: 'center', vertical: 'middle' };
            });
            header.height = 20;

            g.items.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
            for (const s of g.items) {
                const r = ws.addRow([parseISO(s.date), s.address, s.login, s.phone]);
                r.getCell(1).numFmt = 'dd.mm.yyyy';
            }

            const last = ws.rowCount;
            for (let rr = 2; rr <= last; rr++) {
                for (let cc = 1; cc <= 4; cc++) {
                    ws.getCell(rr, cc).border = {
                        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        right: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                    };
                }
            }
        }

        const buffer = await wb.xlsx.writeBuffer();
        return new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
    }

    // Сохранить рабочую копию обратно в файл-базу
    async function persist(opts) {
        opts = opts || {};
        if (supportsFS && fileHandle) {
            try {
                const blob = await buildBlob();
                const w = await fileHandle.createWritable();
                await w.write(blob);
                await w.close();
                if (opts.notify) toast('Сохранено в файл', 'success');
                return true;
            } catch (e) {
                console.error('write error', e);
                toast('Не удалось записать в файл базы', 'error');
                return false;
            }
        }
        // Браузер без File System Access — держим кэш, чтобы не терять данные
        if (!supportsFS) {
            localStorage.setItem(CACHE_KEY, JSON.stringify(rows));
            return true;
        }
        toast('База не открыта — откройте или создайте файл', 'warning');
        return false;
    }

    // ============================================================
    // ОТКРЫТЬ / СОЗДАТЬ БАЗУ
    // ============================================================
    async function openDatabase() {
        if (!supportsFS) { document.getElementById('file-input').click(); return; }
        try {
            const [handle] = await window.showOpenFilePicker({
                types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
            });
            fileHandle = handle;
            fileName = handle.name;
            await HandleDB.set(handle);
            const file = await handle.getFile();
            rows = await readRowsFromBlob(file);
            setStatus();
            render();
            toast(`База открыта: ${rows.length} записей`, 'success');
        } catch (e) {
            if (e && e.name === 'AbortError') return;
            console.error(e);
            toast('Не удалось открыть базу', 'error');
        }
    }

    async function createDatabase() {
        if (!supportsFS) {
            rows = [];
            localStorage.setItem(CACHE_KEY, '[]');
            render();
            toast('Новая база. Заполняйте и нажмите «Сохранить файл».', 'info');
            return;
        }
        try {
            fileHandle = await window.showSaveFilePicker({
                suggestedName: `Абоненты_${new Date().getFullYear()}.xlsx`,
                types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
            });
            fileName = fileHandle.name;
            rows = [];
            await HandleDB.set(fileHandle);
            await persist();
            setStatus();
            render();
            toast('Новая база создана', 'success');
        } catch (e) {
            if (e && e.name === 'AbortError') return;
            console.error(e);
            toast('Не удалось создать базу', 'error');
        }
    }

    // Открытие файла в браузерах без File System Access (read-only загрузка)
    async function openViaInput(file) {
        try {
            rows = await readRowsFromBlob(file);
            fileName = file.name;
            localStorage.setItem(CACHE_KEY, JSON.stringify(rows));
            setStatus();
            render();
            toast(`Загружено: ${rows.length} записей. Изменения сохраняйте кнопкой «Сохранить файл».`, 'success');
        } catch (e) {
            console.error(e);
            toast('Не удалось прочитать файл', 'error');
        }
    }

    function downloadFile() {
        buildBlob().then(blob => {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName || `Абоненты_${todayISO()}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);
        });
    }

    async function tryRestore() {
        if (!supportsFS) {
            try { rows = JSON.parse(localStorage.getItem(CACHE_KEY) || '[]'); } catch { rows = []; }
            setStatus();
            return;
        }
        const stored = await HandleDB.get();
        if (!stored) { setStatus(); return; }
        try {
            const perm = await stored.queryPermission({ mode: 'readwrite' });
            fileName = stored.name;
            if (perm === 'granted') {
                fileHandle = stored;
                rows = await readRowsFromBlob(await stored.getFile());
                render();
            }
        } catch (e) { /* потребуется повторно открыть */ }
        setStatus();
    }

    function setStatus() {
        const el = document.getElementById('db-status');
        if (fileHandle) { el.textContent = `База: ${fileName}`; el.classList.add('linked'); }
        else if (fileName) { el.textContent = `${fileName} (откройте заново)`; el.classList.remove('linked'); }
        else { el.textContent = 'База не открыта'; el.classList.remove('linked'); }
    }

    // ============================================================
    // ДОБАВЛЕНИЕ / УДАЛЕНИЕ
    // ============================================================
    function tryAdd(opts) {
        opts = opts || {};
        const address = document.getElementById('f-address').value.trim();
        const login = document.getElementById('f-login').value.trim();
        const phone = document.getElementById('f-phone').value.trim();
        let date = document.getElementById('f-date').value;

        if (!address || !login || !phone) {
            if (opts.notify) toast('Заполните адрес, логин и номер телефона', 'warning');
            return;
        }
        if (supportsFS && !fileHandle) {
            toast('Сначала откройте или создайте базу', 'warning');
            return;
        }
        if (!date) date = todayISO();

        rows.push({ id: uid(), address, login, phone, date });

        document.getElementById('f-address').value = '';
        document.getElementById('f-login').value = '';
        document.getElementById('f-phone').value = '';
        if (opts.focus) document.getElementById('f-address').focus();

        render();
        persist({ notify: true });
    }

    function removeRow(id) {
        rows = rows.filter(s => s.id !== id);
        render();
        persist();
    }

    // ============================================================
    // RENDER
    // ============================================================
    function render() {
        const container = document.getElementById('table-container');
        const filter = (document.getElementById('f-filter').value || '').trim().toLowerCase();

        let view = rows;
        if (filter) {
            view = rows.filter(s =>
                (s.address || '').toLowerCase().includes(filter) ||
                (s.login || '').toLowerCase().includes(filter) ||
                (s.phone || '').toLowerCase().includes(filter)
            );
        }

        if (view.length === 0) {
            const noDb = supportsFS && !fileHandle;
            container.innerHTML = `<div class="empty-state">${
                noDb ? 'Откройте существующую базу или создайте новую — данные хранятся прямо в .xlsx файле.'
                     : (rows.length === 0 ? 'Пусто. Заполните форму и нажмите Enter.' : 'Ничего не найдено по фильтру.')
            }</div>`;
            return;
        }

        const groups = new Map();
        for (const s of view) {
            const d = parseISO(s.date) || new Date();
            const key = `${d.getUTCFullYear()}-${String(d.getUTCMonth()).padStart(2, '0')}`;
            if (!groups.has(key)) groups.set(key, { year: d.getUTCFullYear(), month: d.getUTCMonth(), items: [] });
            groups.get(key).items.push(s);
        }
        const sortedKeys = Array.from(groups.keys()).sort().reverse();

        container.innerHTML = sortedKeys.map(key => {
            const g = groups.get(key);
            g.items.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
            const rowsHtml = g.items.map(s => `
                <tr>
                    <td>${s.date ? formatDate(s.date) : '—'}</td>
                    <td>${escapeHtml(s.address)}</td>
                    <td>${escapeHtml(s.login)}</td>
                    <td><strong>${escapeHtml(s.phone)}</strong></td>
                    <td class="sub-actions">
                        <button class="btn btn-ghost btn-sm sub-delete" data-id="${s.id}" title="Удалить">🗑</button>
                    </td>
                </tr>
            `).join('');
            return `
                <div class="month-card month-${g.month}">
                    <div class="month-card-header">
                        <h3>${MONTH_NAMES[g.month]} ${g.year}</h3>
                        <span class="badge month-badge">${g.items.length}</span>
                    </div>
                    <div class="table-container">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th style="width:120px;">Дата</th>
                                    <th>Адрес</th>
                                    <th>Логин</th>
                                    <th style="width:160px;">Номер телефона</th>
                                    <th style="width:60px;"></th>
                                </tr>
                            </thead>
                            <tbody>${rowsHtml}</tbody>
                        </table>
                    </div>
                </div>
            `;
        }).join('');

        container.querySelectorAll('.sub-delete').forEach(btn => {
            btn.addEventListener('click', () => {
                if (confirm('Удалить запись?')) removeRow(btn.dataset.id);
            });
        });
    }

    // ============================================================
    // МАЛАЯ БАЗА ДЛЯ ССЫЛКИ НА ФАЙЛ (IndexedDB)
    // ============================================================
    const HandleDB = {
        _db() {
            return new Promise((resolve, reject) => {
                const req = indexedDB.open('abonenty', 1);
                req.onupgradeneeded = () => req.result.createObjectStore('handles');
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
            });
        },
        async set(h) { try { const db = await this._db(); db.transaction('handles', 'readwrite').objectStore('handles').put(h, 'db'); } catch {} },
        async get() {
            try {
                const db = await this._db();
                return await new Promise(res => {
                    const r = db.transaction('handles', 'readonly').objectStore('handles').get('db');
                    r.onsuccess = () => res(r.result || null);
                    r.onerror = () => res(null);
                });
            } catch { return null; }
        },
    };

    // ============================================================
    // UTILS
    // ============================================================
    function uid() { return Date.now() + '-' + Math.random().toString(36).slice(2, 8); }
    function todayISO() { return new Date().toISOString().slice(0, 10); }
    function parseISO(iso) {
        if (!iso) return null;
        const [y, m, d] = String(iso).split('-').map(Number);
        if (!y || !m || !d) return null;
        return new Date(Date.UTC(y, m - 1, d));
    }
    function formatDate(iso) {
        const [y, m, d] = String(iso).split('-');
        return (y && m && d) ? `${d}.${m}.${y}` : iso;
    }
    function cellText(v) {
        if (v == null) return '';
        if (v instanceof Date) return v;
        if (typeof v === 'object') {
            if (v.richText) return v.richText.map(t => t.text).join('');
            if (v.text != null) return v.text;
            if (v.result != null) return v.result;
            return '';
        }
        return v;
    }
    function cellDate(v) {
        if (v instanceof Date) {
            return `${v.getUTCFullYear()}-${String(v.getUTCMonth() + 1).padStart(2, '0')}-${String(v.getUTCDate()).padStart(2, '0')}`;
        }
        if (v == null || v === '') return '';
        const d = new Date(v);
        return isNaN(d) ? '' : `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
    }
    function escapeHtml(text) {
        if (text == null) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    function toast(message, type = 'info', duration = 3500) {
        const container = document.getElementById('toast-container');
        const el = document.createElement('div');
        el.className = `toast ${type}`;
        el.textContent = message;
        container.appendChild(el);
        setTimeout(() => {
            el.style.opacity = '0';
            el.style.transform = 'translateX(40px)';
            el.style.transition = 'all 0.3s ease';
            setTimeout(() => el.remove(), 300);
        }, duration);
    }

    // ============================================================
    // INIT
    // ============================================================
    function init() {
        document.getElementById('f-date').value = todayISO();

        document.getElementById('btn-open').addEventListener('click', openDatabase);
        document.getElementById('btn-create').addEventListener('click', createDatabase);
        document.getElementById('btn-download').addEventListener('click', downloadFile);
        document.getElementById('f-filter').addEventListener('input', render);

        const fi = document.getElementById('file-input');
        fi.addEventListener('change', () => { if (fi.files[0]) openViaInput(fi.files[0]); fi.value = ''; });

        ['f-address', 'f-login', 'f-phone', 'f-date'].forEach(id => {
            document.getElementById(id).addEventListener('keydown', (e) => {
                if (e.key === 'Enter') { e.preventDefault(); tryAdd({ focus: true, notify: true }); }
            });
        });

        // Кнопка «Сохранить файл» нужна только без File System Access
        if (supportsFS) document.getElementById('btn-download').hidden = true;
        else {
            document.getElementById('btn-open').textContent = '📂 Открыть файл';
            document.getElementById('btn-create').textContent = '✨ Новая база';
        }

        initPWA();
        tryRestore();
        render();
    }

    // ============================================================
    // PWA — установка приложения + офлайн (service worker)
    // ============================================================
    function initPWA() {
        // Регистрируем service worker (только по http/https)
        if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
            window.addEventListener('load', () => {
                navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW:', e));
            });
        }

        const btn = document.getElementById('btn-install');
        let deferred = null;

        window.addEventListener('beforeinstallprompt', (e) => {
            e.preventDefault();
            deferred = e;
            btn.hidden = false;
        });

        btn.addEventListener('click', async () => {
            if (!deferred) return;
            deferred.prompt();
            await deferred.userChoice;
            deferred = null;
            btn.hidden = true;
        });

        window.addEventListener('appinstalled', () => {
            btn.hidden = true;
            toast('Приложение установлено', 'success');
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
