/**
 * ============================================================
 * Абоненты — настольное приложение со своей базой данных
 * ============================================================
 * Своя база = отдельный ТЕКСТОВЫЙ файл (.txt), независимый от Excel.
 * Файл человекочитаемый: записи разбиты по месяцам, аккуратными
 * строками — можно открыть в «Блокноте» и просто читать.
 *
 * Приложение читает базу прямо из файла и пишет изменения обратно
 * (File System Access API). Excel — отдельная кнопка экспорта
 * (цветная таблица по месяцам), базой не является.
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
    const CACHE_KEY = 'abonenty_cache_v3';
    const supportsFS = typeof window.showOpenFilePicker === 'function';

    let rows = [];
    let fileHandle = null;
    let fileName = null;

    // ============================================================
    // ТЕКСТОВАЯ БАЗА: сериализация / разбор
    // ============================================================
    function groupByMonth(list) {
        const groups = new Map();
        for (const s of list) {
            const d = parseISO(s.date) || new Date();
            const key = `${d.getUTCFullYear()}-${String(d.getUTCMonth()).padStart(2, '0')}`;
            if (!groups.has(key)) groups.set(key, { y: d.getUTCFullYear(), m: d.getUTCMonth(), items: [] });
            groups.get(key).items.push(s);
        }
        return groups;
    }

    function serializeTxt() {
        const lines = [];
        lines.push('АБОНЕНТЫ — база данных');
        lines.push('Обновлено: ' + nowStr());
        lines.push('Записей: ' + rows.length);
        lines.push('Формат строки:  ДД.ММ.ГГГГ | Адрес | Логин | Номер');
        lines.push('');

        const groups = groupByMonth(rows);
        const keys = Array.from(groups.keys()).sort().reverse(); // новые месяцы сверху
        for (const key of keys) {
            const g = groups.get(key);
            lines.push(`================  ${MONTH_NAMES[g.m]} ${g.y}  ================`);
            g.items.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
            for (const s of g.items) {
                lines.push(`${formatDate(s.date)} | ${s.address} | ${s.login} | ${s.phone}`);
            }
            lines.push('');
        }
        return lines.join('\r\n');
    }

    function parseTxt(text) {
        const out = [];
        for (const raw of String(text).split(/\r?\n/)) {
            const m = raw.match(/^\s*(\d{2})\.(\d{2})\.(\d{4})\s*\|(.*)$/);
            if (!m) continue;
            const date = `${m[3]}-${m[2]}-${m[1]}`;
            const rest = m[4].split('|').map(x => x.trim());
            const address = rest[0] || '', login = rest[1] || '', phone = rest[2] || '';
            if (!address && !login && !phone) continue;
            out.push({ id: uid(), date, address, login, phone });
        }
        return out;
    }

    // ============================================================
    // ЗАПИСЬ / ЧТЕНИЕ БАЗЫ (файл .txt)
    // ============================================================
    async function persist(opts) {
        opts = opts || {};
        if (supportsFS && fileHandle) {
            try {
                const w = await fileHandle.createWritable();
                await w.write(serializeTxt());
                await w.close();
                if (opts.notify) toast('Сохранено в базу', 'success');
                return true;
            } catch (e) {
                console.error('write error', e);
                toast('Не удалось записать в файл базы', 'error');
                return false;
            }
        }
        if (!supportsFS) {
            localStorage.setItem(CACHE_KEY, JSON.stringify(rows));
            return true;
        }
        toast('База не открыта — откройте или создайте файл', 'warning');
        return false;
    }

    async function openDatabase() {
        if (!supportsFS) { document.getElementById('file-input').click(); return; }
        try {
            const [handle] = await window.showOpenFilePicker({
                types: [{ description: 'База абонентов (текст)', accept: { 'text/plain': ['.txt'] } }],
            });
            fileHandle = handle;
            fileName = handle.name;
            await HandleDB.set(handle);
            rows = parseTxt(await (await handle.getFile()).text());
            setStatus(); render();
            toast(`База открыта: ${rows.length} записей`, 'success');
        } catch (e) {
            if (e && e.name === 'AbortError') return;
            console.error(e); toast('Не удалось открыть базу', 'error');
        }
    }

    async function createDatabase() {
        if (!supportsFS) {
            rows = []; localStorage.setItem(CACHE_KEY, '[]'); render();
            toast('Новая база. Заполняйте и нажмите «Сохранить базу».', 'info');
            return;
        }
        try {
            fileHandle = await window.showSaveFilePicker({
                suggestedName: 'База_абонентов.txt',
                types: [{ description: 'База абонентов (текст)', accept: { 'text/plain': ['.txt'] } }],
            });
            fileName = fileHandle.name;
            rows = [];
            await HandleDB.set(fileHandle);
            await persist();
            setStatus(); render();
            toast('Новая база создана', 'success');
        } catch (e) {
            if (e && e.name === 'AbortError') return;
            console.error(e); toast('Не удалось создать базу', 'error');
        }
    }

    async function openViaInput(file) {
        try {
            rows = parseTxt(await file.text());
            fileName = file.name;
            localStorage.setItem(CACHE_KEY, JSON.stringify(rows));
            setStatus(); render();
            toast(`Загружено: ${rows.length} записей. Сохраняйте кнопкой «Сохранить базу».`, 'success');
        } catch (e) { console.error(e); toast('Не удалось прочитать файл', 'error'); }
    }

    function downloadTxt() {
        const blob = new Blob(['﻿' + serializeTxt()], { type: 'text/plain;charset=utf-8' });
        triggerDownload(blob, fileName || 'База_абонентов.txt');
    }

    async function tryRestore() {
        if (!supportsFS) {
            try { rows = JSON.parse(localStorage.getItem(CACHE_KEY) || '[]'); } catch { rows = []; }
            setStatus(); return;
        }
        const stored = await HandleDB.get();
        if (!stored) { setStatus(); return; }
        try {
            fileName = stored.name;
            if (await stored.queryPermission({ mode: 'readwrite' }) === 'granted') {
                fileHandle = stored;
                rows = parseTxt(await (await stored.getFile()).text());
                render();
            }
        } catch (e) { /* потребуется открыть заново */ }
        setStatus();
    }

    function setStatus() {
        const el = document.getElementById('db-status');
        if (fileHandle) { el.textContent = `База: ${fileName}`; el.classList.add('linked'); }
        else if (fileName) { el.textContent = `${fileName} (откройте заново)`; el.classList.remove('linked'); }
        else { el.textContent = 'База не открыта'; el.classList.remove('linked'); }
    }

    // ============================================================
    // ЭКСПОРТ В EXCEL (цветная таблица по месяцам) — не база, по запросу
    // ============================================================
    async function exportExcel() {
        if (rows.length === 0) { toast('База пуста — нечего экспортировать', 'warning'); return; }
        const wb = new ExcelJS.Workbook();
        const groups = groupByMonth(rows);
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
            for (let rr = 2; rr <= last; rr++)
                for (let cc = 1; cc <= 4; cc++)
                    ws.getCell(rr, cc).border = {
                        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        right: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                    };
        }
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        triggerDownload(blob, (fileName ? fileName.replace(/\.txt$/i, '') : 'Абоненты') + '.xlsx');
        toast('Excel выгружен', 'success');
    }

    // ============================================================
    // ДОБАВЛЕНИЕ / УДАЛЕНИЕ
    // ============================================================
    function tryAdd(opts) {
        opts = opts || {};
        const address = clean(document.getElementById('f-address').value);
        const login = clean(document.getElementById('f-login').value);
        const phone = clean(document.getElementById('f-phone').value);
        let date = document.getElementById('f-date').value;

        if (!address || !login || !phone) {
            if (opts.notify) toast('Заполните адрес, логин и номер телефона', 'warning');
            return;
        }
        if (supportsFS && !fileHandle) { toast('Сначала откройте или создайте базу', 'warning'); return; }
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
                (s.phone || '').toLowerCase().includes(filter));
        }

        if (view.length === 0) {
            const noDb = supportsFS && !fileHandle;
            container.innerHTML = `<div class="empty-state">${
                noDb ? 'Откройте базу или создайте новую — данные хранятся в отдельном текстовом файле.'
                     : (rows.length === 0 ? 'Пусто. Заполните форму и нажмите Enter.' : 'Ничего не найдено по фильтру.')
            }</div>`;
            return;
        }

        const groups = groupByMonth(view);
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
                </tr>`).join('');
            return `
                <div class="month-card month-${g.m}">
                    <div class="month-card-header">
                        <h3>${MONTH_NAMES[g.m]} ${g.y}</h3>
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
                </div>`;
        }).join('');

        container.querySelectorAll('.sub-delete').forEach(btn => {
            btn.addEventListener('click', () => { if (confirm('Удалить запись?')) removeRow(btn.dataset.id); });
        });
    }

    // ============================================================
    // IndexedDB — ссылка на файл базы
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
                    r.onsuccess = () => res(r.result || null); r.onerror = () => res(null);
                });
            } catch { return null; }
        },
    };

    // ============================================================
    // PWA
    // ============================================================
    function initPWA() {
        if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
            window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW:', e)));
        }
        const btn = document.getElementById('btn-install');
        let deferred = null;
        window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferred = e; btn.hidden = false; });
        btn.addEventListener('click', async () => { if (!deferred) return; deferred.prompt(); await deferred.userChoice; deferred = null; btn.hidden = true; });
        window.addEventListener('appinstalled', () => { btn.hidden = true; toast('Приложение установлено', 'success'); });
    }

    // ============================================================
    // UTILS
    // ============================================================
    function uid() { return Date.now() + '-' + Math.random().toString(36).slice(2, 8); }
    function clean(v) { return String(v || '').trim().replace(/\|/g, '/'); } // защищаем формат базы
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
    function nowStr() {
        const d = new Date();
        const p = n => String(n).padStart(2, '0');
        return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
    }
    function triggerDownload(blob, name) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = name; a.click();
        URL.revokeObjectURL(url);
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
        document.getElementById('btn-excel').addEventListener('click', exportExcel);
        document.getElementById('btn-download').addEventListener('click', downloadTxt);
        document.getElementById('f-filter').addEventListener('input', render);

        const fi = document.getElementById('file-input');
        fi.addEventListener('change', () => { if (fi.files[0]) openViaInput(fi.files[0]); fi.value = ''; });

        ['f-address', 'f-login', 'f-phone', 'f-date'].forEach(id => {
            document.getElementById(id).addEventListener('keydown', (e) => {
                if (e.key === 'Enter') { e.preventDefault(); tryAdd({ focus: true, notify: true }); }
            });
        });

        if (supportsFS) document.getElementById('btn-download').hidden = true;
        else {
            document.getElementById('btn-open').textContent = '📂 Открыть файл';
            document.getElementById('btn-create').textContent = '✨ Новая база';
        }

        initPWA();
        tryRestore();
        render();
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
