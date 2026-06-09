/**
 * ============================================================
 * Актуализация номеров абонентов — standalone
 * ============================================================
 * Готовый статичный сайт. Данные пока хранятся локально в браузере
 * (localStorage). Бэкенд подключается позже — достаточно заменить
 * методы объекта Store на сетевые запросы (см. комментарии ниже).
 * ============================================================
 */

(function () {
    'use strict';

    const STORAGE_KEY = 'subscribers_v1';

    const MONTH_NAMES = [
        'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
        'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь',
    ];

    // ============================================================
    // STORE — слой данных
    // Сейчас localStorage. Чтобы подключить бэкенд позже —
    // заменить тела методов на fetch(...) к вашему API.
    // ============================================================
    const Store = {
        all() {
            try {
                return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
            } catch {
                return [];
            }
        },
        save(list) {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
        },
        add(item) {
            const list = this.all();
            item.id = Date.now() + '-' + Math.random().toString(36).slice(2, 8);
            list.push(item);
            this.save(list);
            return item;
        },
        remove(id) {
            this.save(this.all().filter(s => s.id !== id));
        },
    };

    // ============================================================
    // RENDER
    // ============================================================
    function render() {
        const container = document.getElementById('table-container');
        const totalBadge = document.getElementById('total-count');
        const filter = (document.getElementById('f-filter').value || '').trim().toLowerCase();

        const allRows = Store.all();
        totalBadge.textContent = `${allRows.length} записей`;

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
                    ? 'Нет абонентов. Добавьте первую запись через форму выше.'
                    : 'Ничего не найдено по фильтру.'
            }</div>`;
            return;
        }

        // Группируем по месяцу (ключ YYYY-MM по дате)
        const groups = new Map();
        for (const s of rows) {
            const d = s.date ? new Date(s.date) : new Date();
            const key = `${d.getFullYear()}-${String(d.getMonth()).padStart(2, '0')}`;
            if (!groups.has(key)) {
                groups.set(key, { year: d.getFullYear(), month: d.getMonth(), items: [] });
            }
            groups.get(key).items.push(s);
        }

        // Новые месяцы — сверху; внутри месяца — новые даты сверху
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
                    toast('Запись удалена', 'success');
                    render();
                }
            });
        });
    }

    // ============================================================
    // ACTIONS
    // ============================================================
    function addSubscriber() {
        const address = document.getElementById('f-address').value.trim();
        const login = document.getElementById('f-login').value.trim();
        const phone = document.getElementById('f-phone').value.trim();
        let date = document.getElementById('f-date').value;

        if (!address || !login || !phone) {
            toast('Заполните адрес, логин и номер телефона', 'warning');
            return;
        }
        if (!date) date = todayISO();

        Store.add({ address, login, phone, date });
        toast('Абонент добавлен', 'success');

        // Очищаем поля, дату оставляем для серии записей
        document.getElementById('f-address').value = '';
        document.getElementById('f-login').value = '';
        document.getElementById('f-phone').value = '';
        document.getElementById('f-address').focus();

        render();
    }

    function exportCsv() {
        const rows = Store.all();
        if (rows.length === 0) {
            toast('Нет данных для экспорта', 'warning');
            return;
        }
        const header = ['Дата', 'Адрес', 'Логин', 'Номер телефона'];
        const lines = [header.join(';')].concat(
            rows
                .sort((a, b) => (b.date || '').localeCompare(a.date || ''))
                .map(s => [s.date || '', s.address, s.login, s.phone].map(csvCell).join(';'))
        );
        const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `abonenty_${todayISO()}.csv`;
        a.click();
        URL.revokeObjectURL(url);
    }

    // ============================================================
    // UTILS
    // ============================================================
    function todayISO() { return new Date().toISOString().slice(0, 10); }

    function formatDate(iso) {
        const d = new Date(iso);
        if (isNaN(d)) return iso;
        return d.toLocaleDateString('ru-RU');
    }

    function csvCell(v) {
        const s = String(v ?? '');
        return /[";\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
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

        document.getElementById('btn-add').addEventListener('click', addSubscriber);
        document.getElementById('btn-export').addEventListener('click', exportCsv);
        document.getElementById('f-filter').addEventListener('input', render);
        ['f-address', 'f-login', 'f-phone'].forEach(id => {
            document.getElementById(id).addEventListener('keydown', (e) => {
                if (e.key === 'Enter') addSubscriber();
            });
        });

        render();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
