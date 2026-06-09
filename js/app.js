/**
 * ============================================================
 * Актуализация номеров абонентов — запись прямо в Excel
 * ============================================================
 * Что заполняется в форме — автоматически уходит в .xlsx файл,
 * разбитый по месяцам (каждый месяц своего цвета).
 *
 * В Chrome/Edge: один раз «Привязать Excel» — и дальше каждая
 * запись пишется прямо в выбранный файл (File System Access API).
 * В других браузерах: кнопка «Скачать» отдаёт готовый .xlsx.
 *
 * localStorage хранит данные между перезагрузками и служит
 * источником для пересборки файла.
 * ============================================================
 */

(function () {
    'use strict';

    const STORAGE_KEY = 'subscribers_v1';

    const MONTH_NAMES = [
        'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
        'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь',
    ];

    // Цвета месяцев (RGB) — те же, что на странице
    const MONTH_COLORS = [
        '60A5FA', '22D3EE', '34D399', 'A3E635', 'FACC15', 'FB923C',
        'F87171', 'F472B6', 'E879F9', 'C084FC', '818CF8', '38BDF8',
    ];

    const supportsFS = typeof window.showSaveFilePicker === 'function';
    let fileHandle = null;   // привязанный .xlsx
    let fileName = null;

    // ============================================================
    // STORE — данные (localStorage)
    // ============================================================
    const Store = {
        all() {
            try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); }
            catch { return []; }
        },
        save(list) { localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); },
        add(item) {
            const list = this.all();
            item.id = Date.now() + '-' + Math.random().toString(36).slice(2, 8);
            list.push(item);
            this.save(list);
            return item;
        },
        remove(id) { this.save(this.all().filter(s => s.id !== id)); },
    };

    // ============================================================
    // IndexedDB — хранение ссылки на файл между перезагрузками
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
        async set(handle) {
            try {
                const db = await this._db();
                const tx = db.transaction('handles', 'readwrite');
                tx.objectStore('handles').put(handle, 'excel');
            } catch (e) { /* not critical */ }
        },
        async get() {
            try {
                const db = await this._db();
                return await new Promise((resolve) => {
                    const r = db.transaction('handles', 'readonly')
                        .objectStore('handles').get('excel');
                    r.onsuccess = () => resolve(r.result || null);
                    r.onerror = () => resolve(null);
                });
            } catch (e) { return null; }
        },
    };

    // ============================================================
    // EXCEL — сборка книги (ExcelJS), цвета по месяцам
    // ============================================================
    async function buildWorkbook() {
        const wb = new ExcelJS.Workbook();
        const rows = Store.all();

        // Группируем по (год, месяц)
        const groups = new Map();
        for (const s of rows) {
            const d = s.date ? new Date(s.date) : new Date();
            const key = `${d.getFullYear()}-${String(d.getMonth()).padStart(2, '0')}`;
            if (!groups.has(key)) groups.set(key, { y: d.getFullYear(), m: d.getMonth(), items: [] });
            groups.get(key).items.push(s);
        }
        // Если данных нет — создаём лист текущего месяца, чтобы файл был валидным
        if (groups.size === 0) {
            const now = new Date();
            groups.set('empty', { y: now.getFullYear(), m: now.getMonth(), items: [] });
        }

        const keys = Array.from(groups.keys()).sort(); // хронологически
        for (const key of keys) {
            const g = groups.get(key);
            const argb = 'FF' + MONTH_COLORS[g.m];
            const ws = wb.addWorksheet(`${MONTH_NAMES[g.m]} ${g.y}`, {
                properties: { tabColor: { argb } },
                views: [{ state: 'frozen', ySplit: 2 }],
            });

            ws.columns = [
                { width: 14 }, { width: 38 }, { width: 20 }, { width: 20 },
            ];

            // Заголовок месяца
            ws.mergeCells('A1:D1');
            const title = ws.getCell('A1');
            title.value = `${MONTH_NAMES[g.m]} ${g.y}`;
            title.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
            title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
            title.alignment = { horizontal: 'center', vertical: 'middle' };
            ws.getRow(1).height = 24;

            // Шапка
            const header = ws.getRow(2);
            ['Дата', 'Адрес', 'Логин', 'Номер телефона'].forEach((h, i) => {
                const c = header.getCell(i + 1);
                c.value = h;
                c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
                c.alignment = { horizontal: 'center', vertical: 'middle' };
            });
            header.height = 20;

            // Данные (новые сверху)
            g.items.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
            for (const s of g.items) {
                const r = ws.addRow([
                    s.date ? new Date(s.date) : null,
                    s.address, s.login, s.phone,
                ]);
                r.getCell(1).numFmt = 'dd.mm.yyyy';
            }

            // Рамки на всю заполненную область
            const lastRow = ws.rowCount;
            for (let rr = 2; rr <= lastRow; rr++) {
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

    // Записать книгу: в привязанный файл (тихо) либо скачиванием
    async function saveExcel(opts) {
        opts = opts || {};
        const blob = await buildWorkbook();

        if (fileHandle) {
            try {
                const w = await fileHandle.createWritable();
                await w.write(blob);
                await w.close();
                if (opts.notify) toast('Сохранено в Excel', 'success');
                return true;
            } catch (e) {
                console.error('Excel write error:', e);
                toast('Не удалось записать в файл — скачиваю копию', 'warning');
            }
        }

        if (opts.download || !supportsFS) {
            downloadBlob(blob);
            return true;
        }

        if (opts.notify) toast('Файл Excel не привязан — нажмите «Привязать Excel»', 'warning');
        return false;
    }

    function downloadBlob(blob) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName || `Абоненты_номера_${todayISO()}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
    }

    // ============================================================
    // Привязка Excel-файла (File System Access API)
    // ============================================================
    async function linkExcel() {
        if (!supportsFS) {
            toast('Этот браузер не умеет писать в файл. Используйте «Скачать» (или Chrome/Edge).', 'warning');
            return;
        }
        try {
            // Если есть сохранённая ссылка — пробуем переиспользовать
            const stored = await HandleDB.get();
            if (stored) {
                const perm = await stored.requestPermission({ mode: 'readwrite' });
                if (perm === 'granted') {
                    fileHandle = stored;
                    fileName = stored.name;
                    setExcelStatus();
                    await saveExcel({ notify: true });
                    return;
                }
            }
            // Иначе — выбираем/создаём файл
            fileHandle = await window.showSaveFilePicker({
                suggestedName: `Абоненты_номера_${new Date().getFullYear()}.xlsx`,
                types: [{
                    description: 'Excel',
                    accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
                }],
            });
            fileName = fileHandle.name;
            await HandleDB.set(fileHandle);
            setExcelStatus();
            await saveExcel({ notify: true });
        } catch (e) {
            if (e && e.name === 'AbortError') return; // пользователь отменил
            console.error('Link error:', e);
            toast('Не удалось привязать файл', 'error');
        }
    }

    async function tryRestoreHandle() {
        if (!supportsFS) return;
        const stored = await HandleDB.get();
        if (!stored) return;
        try {
            const perm = await stored.queryPermission({ mode: 'readwrite' });
            if (perm === 'granted') {
                fileHandle = stored;
                fileName = stored.name;
            } else {
                fileName = stored.name; // покажем имя, но потребуется повторно привязать
            }
        } catch (e) { /* ignore */ }
        setExcelStatus();
    }

    function setExcelStatus() {
        const el = document.getElementById('excel-status');
        if (fileHandle) {
            el.textContent = `Excel: ${fileName}`;
            el.classList.add('linked');
        } else if (fileName) {
            el.textContent = `Excel: ${fileName} (привяжите заново)`;
            el.classList.remove('linked');
        } else {
            el.textContent = supportsFS ? 'Excel не привязан' : 'Режим скачивания';
            el.classList.remove('linked');
        }
    }

    // ============================================================
    // Добавление записи → сразу в Excel
    // ============================================================
    function tryAutoAdd(opts) {
        opts = opts || {};
        const address = document.getElementById('f-address').value.trim();
        const login = document.getElementById('f-login').value.trim();
        const phone = document.getElementById('f-phone').value.trim();
        let date = document.getElementById('f-date').value;

        if (!address || !login || !phone) {
            if (opts.notify) toast('Заполните адрес, логин и номер телефона', 'warning');
            return;
        }
        if (!date) date = todayISO();

        Store.add({ address, login, phone, date });

        document.getElementById('f-address').value = '';
        document.getElementById('f-login').value = '';
        document.getElementById('f-phone').value = '';
        if (opts.focus) document.getElementById('f-address').focus();

        render();
        saveExcel({ notify: true }); // пишем прямо в файл
    }

    // ============================================================
    // RENDER — предпросмотр того, что в Excel
    // ============================================================
    function render() {
        const container = document.getElementById('table-container');
        const filter = (document.getElementById('f-filter').value || '').trim().toLowerCase();

        const allRows = Store.all();
        let rows = allRows;
        if (filter) {
            rows = rows.filter(s =>
                (s.address || '').toLowerCase().includes(filter) ||
                (s.login || '').toLowerCase().includes(filter) ||
                (s.phone || '').toLowerCase().includes(filter)
            );
        }

        if (rows.length === 0) {
            container.innerHTML = `<div class="empty-state">${
                allRows.length === 0
                    ? 'Пока пусто. Заполните форму — строка уйдёт в Excel и появится здесь.'
                    : 'Ничего не найдено по фильтру.'
            }</div>`;
            return;
        }

        const groups = new Map();
        for (const s of rows) {
            const d = s.date ? new Date(s.date) : new Date();
            const key = `${d.getFullYear()}-${String(d.getMonth()).padStart(2, '0')}`;
            if (!groups.has(key)) groups.set(key, { year: d.getFullYear(), month: d.getMonth(), items: [] });
            groups.get(key).items.push(s);
        }
        const sortedKeys = Array.from(groups.keys()).sort().reverse();

        container.innerHTML = sortedKeys.map(key => {
            const g = groups.get(key);
            g.items.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
            const monthLabel = `${MONTH_NAMES[g.month]} ${g.year}`;
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
                        <h3>${monthLabel}</h3>
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
                if (confirm('Удалить запись абонента?')) {
                    Store.remove(btn.dataset.id);
                    toast('Удалено', 'success');
                    render();
                    saveExcel();
                }
            });
        });
    }

    // ============================================================
    // UTILS
    // ============================================================
    function todayISO() { return new Date().toISOString().slice(0, 10); }
    function formatDate(iso) {
        const d = new Date(iso);
        return isNaN(d) ? iso : d.toLocaleDateString('ru-RU');
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

        document.getElementById('btn-link').addEventListener('click', linkExcel);
        document.getElementById('btn-download').addEventListener('click', () => saveExcel({ download: true }));
        document.getElementById('f-filter').addEventListener('input', render);

        // Быстрый ввод: Enter из любого поля добавляет строку и
        // возвращает курсор в «Адрес» для следующей записи.
        ['f-address', 'f-login', 'f-phone', 'f-date'].forEach(id => {
            document.getElementById(id).addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    tryAutoAdd({ focus: true, notify: true });
                }
            });
        });

        setExcelStatus();
        tryRestoreHandle();
        render();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
