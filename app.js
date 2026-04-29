// ============================================================
// 智能记账本 - Smart Ledger App
// ============================================================

const app = (() => {
    // ========== Default Data ==========
    const DEFAULT_CATEGORIES = {
        expense: [
            { id: 'food', name: '餐饮', icon: '🍜' },
            { id: 'transport', name: '交通', icon: '🚗' },
            { id: 'shopping', name: '购物', icon: '🛒' },
            { id: 'entertainment', name: '娱乐', icon: '🎮' },
            { id: 'housing', name: '住房', icon: '🏠' },
            { id: 'medical', name: '医疗', icon: '💊' },
            { id: 'education', name: '教育', icon: '📚' },
            { id: 'communication', name: '通讯', icon: '📱' },
            { id: 'beauty', name: '美容', icon: '💄' },
            { id: 'social', name: '社交', icon: '🤝' },
            { id: 'pet', name: '宠物', icon: '🐱' },
            { id: 'travel', name: '旅行', icon: '✈️' },
            { id: 'digital', name: '数码', icon: '💻' },
            { id: 'daily', name: '日用', icon: '🧴' },
            { id: 'other_expense', name: '其他', icon: '📦' },
        ],
        income: [
            { id: 'salary', name: '工资', icon: '💰' },
            { id: 'bonus', name: '奖金', icon: '🎁' },
            { id: 'invest', name: '投资', icon: '📈' },
            { id: 'sidejob', name: '兼职', icon: '💼' },
            { id: 'refund', name: '退款', icon: '↩️' },
            { id: 'redpacket', name: '红包', icon: '🧧' },
            { id: 'other_income', name: '其他', icon: '💎' },
        ]
    };

    // Keyword → category auto-matching rules
    const DEFAULT_KEYWORD_RULES = [
        { keywords: ['美团外卖', '饿了么', '肯德基', '麦当劳', '星巴克', '瑞幸', '喜茶', '海底捞', '餐厅', '火锅', '奶茶', '早餐', '午餐', '晚餐', '零食', '水果', '菜市场', '超市'], category: 'food', type: 'expense' },
        { keywords: ['滴滴', '出租车', '地铁', '公交', '高铁', '火车', '机票', '停车', '加油', '打车', '共享单车', '哈啰'], category: 'transport', type: 'expense' },
        { keywords: ['淘宝', '京东', '拼多多', '天猫', '购物', '买', '商场'], category: 'shopping', type: 'expense' },
        { keywords: ['电影', '游戏', 'KTV', '酒吧', '演出', '演唱会', '门票', '会员', 'VIP', '视频会员', '音乐会员', 'Netflix', '爱奇艺', '腾讯视频', '优酷', 'B站'], category: 'entertainment', type: 'expense' },
        { keywords: ['房租', '水电', '物业', '燃气', '暖气', '宽带', '网费'], category: 'housing', type: 'expense' },
        { keywords: ['医院', '药店', '门诊', '体检', '挂号', '看病', '医药'], category: 'medical', type: 'expense' },
        { keywords: ['课程', '培训', '考试', '教材', '书', '学费'], category: 'education', type: 'expense' },
        { keywords: ['话费', '流量', '手机', '充值'], category: 'communication', type: 'expense' },
        { keywords: ['理发', '美甲', '化妆品', '护肤', '美容'], category: 'beauty', type: 'expense' },
        { keywords: ['请客', '聚餐', '份子钱', '红包', '礼物', '人情'], category: 'social', type: 'expense' },
        { keywords: ['猫粮', '狗粮', '宠物', '兽医'], category: 'pet', type: 'expense' },
        { keywords: ['酒店', '民宿', '景点', '旅游', '签证'], category: 'travel', type: 'expense' },
        { keywords: ['电脑', '手机壳', '耳机', '数据线', '充电器', '键盘', '鼠标'], category: 'digital', type: 'expense' },
        { keywords: ['纸巾', '洗衣液', '牙膏', '日用品', '收纳'], category: 'daily', type: 'expense' },
        { keywords: ['工资', '薪资', '薪水', '月薪'], category: 'salary', type: 'income' },
        { keywords: ['奖金', '年终奖', '绩效'], category: 'bonus', type: 'income' },
        { keywords: ['利息', '理财', '基金', '股票', '分红', '收益'], category: 'invest', type: 'income' },
        { keywords: ['兼职', '外快', '稿费', '私活'], category: 'sidejob', type: 'income' },
        { keywords: ['退款', '退货', '赔偿'], category: 'refund', type: 'income' },
        { keywords: ['红包', '转账'], category: 'redpacket', type: 'income' },
    ];

    // ========== State ==========
    let state = {
        records: [],
        categories: { ...DEFAULT_CATEGORIES },
        keywordRules: [...DEFAULT_KEYWORD_RULES],
        currentPage: 'dashboard',
        recordFormType: 'expense',
        selectedCategory: null,
        // Summary state
        summaryPeriod: 'month',
        summaryDate: new Date(),
        // Records pagination
        recordsPage: 1,
        recordsPerPage: 20,
        // Batch mode
        batchMode: false,
        selectedIds: new Set(),
        // Import temp
        importData: null,
        importHeaders: null,
        importMapping: {},
        importWorkbook: null,
    };

    // Chart instances
    let charts = {};

    // Debounce utility
    function debounce(fn, delay) {
        let timer;
        return (...args) => {
            clearTimeout(timer);
            timer = setTimeout(() => fn(...args), delay);
        };
    }

    // Simple filter cache to avoid re-filtering large datasets on pagination-only changes
    let _filterCache = { key: '', result: [] };

    // ========== Storage ==========
    function saveState() {
        // Invalidate filter cache on any data change
        _filterCache = { key: '', result: [] };
        const data = {
            records: state.records,
            categories: state.categories,
            keywordRules: state.keywordRules,
        };
        localStorage.setItem('smart-ledger-data', JSON.stringify(data));
    }

    function loadState() {
        try {
            const raw = localStorage.getItem('smart-ledger-data');
            if (raw) {
                const data = JSON.parse(raw);
                state.records = data.records || [];
                if (data.categories) {
                    state.categories = data.categories;
                }
                if (data.keywordRules) {
                    state.keywordRules = data.keywordRules;
                }
            }
        } catch (e) {
            console.error('Failed to load data:', e);
        }
    }

    // ========== Utilities ==========
    function generateId() {
        return Date.now().toString(36) + Math.random().toString(36).substr(2, 6);
    }

    function formatMoney(amount) {
        return '¥' + Math.abs(amount).toFixed(2);
    }

    // Prevent XSS when inserting user-provided text into innerHTML
    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    function formatDate(dateStr) {
        const d = new Date(dateStr);
        // Use UTC methods to avoid timezone offset issues
        // (e.g. '2024-01-15' parsed as UTC midnight → local time shifts to previous day in UTC+ zones)
        return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
    }

    function toast(message, type = 'success') {
        const container = document.getElementById('toast-container');
        const el = document.createElement('div');
        el.className = `toast ${type}`;
        el.textContent = message;
        container.appendChild(el);
        setTimeout(() => el.remove(), 3000);
    }

    function getCategoryInfo(categoryId, type) {
        const list = state.categories[type] || [];
        return list.find(c => c.id === categoryId) || { icon: '❓', name: '未分类' };
    }

    // Auto-classify a description
    function autoClassify(desc, type = null) {
        if (!desc) return null;
        const lower = desc.toLowerCase();
        for (const rule of state.keywordRules) {
            if (type && rule.type !== type) continue;
            for (const kw of rule.keywords) {
                if (lower.includes(kw.toLowerCase())) {
                    return { category: rule.category, type: rule.type };
                }
            }
        }
        return null;
    }

    // ========== AA / Group Payment Logic ==========
    // The KEY feature: when you pay for everyone and collect via group payment,
    // only your personal share counts as actual expense (not the whole amount).
    function calculateActualExpense(record) {
        if (record.type !== 'expense') return 0;
        if (record.aa && record.aa.enabled) {
            return record.aa.myShare;
        }
        return record.amount;
    }

    // Get totals for a set of records
    function calcTotals(records) {
        let income = 0, expense = 0, actualExpense = 0;
        for (const r of records) {
            if (r.type === 'income') {
                income += r.amount;
            } else {
                expense += r.amount;
                actualExpense += calculateActualExpense(r);
            }
        }
        return {
            income: Math.round(income * 100) / 100,
            expense: Math.round(expense * 100) / 100,
            actualExpense: Math.round(actualExpense * 100) / 100,
            balance: Math.round((income - actualExpense) * 100) / 100,
        };
    }

    // ========== Navigation ==========
    function navigateTo(page) {
        state.currentPage = page;
        document.querySelectorAll('.nav-item').forEach(el => {
            el.classList.toggle('active', el.dataset.page === page);
        });
        document.querySelectorAll('.page').forEach(el => {
            el.classList.toggle('active', el.id === `page-${page}`);
        });
        renderPage(page);
    }

    function renderPage(page) {
        switch (page) {
            case 'dashboard': renderDashboard(); break;
            case 'add': renderAddForm(); break;
            case 'records': renderRecords(); break;
            case 'import': break; // static
            case 'summary': renderSummary(); break;
            case 'categories': renderCategories(); break;
        }
    }

    // ========== Dashboard ==========
    function renderDashboard() {
        const now = new Date();
        document.getElementById('current-date').textContent =
            `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;

        // Month records
        const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const monthRecords = state.records.filter(r => r.date.startsWith(monthKey));
        const totals = calcTotals(monthRecords);

        document.getElementById('month-income').textContent = formatMoney(totals.income);
        document.getElementById('month-expense').textContent = formatMoney(totals.expense);
        document.getElementById('month-actual-expense').textContent = formatMoney(totals.actualExpense);
        document.getElementById('month-balance').textContent = formatMoney(totals.balance);

        // Recent records
        const recent = [...state.records].sort((a, b) => b.date.localeCompare(a.date) || b.createdAt - a.createdAt).slice(0, 8);
        const listEl = document.getElementById('recent-list');
        if (recent.length === 0) {
            listEl.innerHTML = '<div class="empty-state"><div class="empty-icon">📝</div><p>还没有记录，去记一笔吧</p></div>';
        } else {
            listEl.innerHTML = recent.map(r => renderRecordItem(r)).join('');
        }

        // Charts
        renderDashboardCharts(monthRecords, now);
    }

    function renderDashboardCharts(monthRecords, now) {
        // Pie chart - expense by category
        const expenseRecords = monthRecords.filter(r => r.type === 'expense');
        const catTotals = {};
        for (const r of expenseRecords) {
            const actual = calculateActualExpense(r);
            catTotals[r.category] = (catTotals[r.category] || 0) + actual;
        }

        const catEntries = Object.entries(catTotals).sort((a, b) => b[1] - a[1]);
        const pieLabels = catEntries.map(([catId]) => getCategoryInfo(catId, 'expense').name);
        const pieData = catEntries.map(([, val]) => val);
        const pieColors = [
            '#5B7FFF', '#66D4A0', '#FF8A80', '#9B8FFF', '#FFB74D',
            '#CE93D8', '#80DEEA', '#F48FB1', '#81D4FA', '#FFF176',
            '#BCAAA4', '#B0BEC5', '#90A4AE', '#78909C', '#4DB6AC'
        ];

        if (charts.dashPie) charts.dashPie.destroy();
        const pieCtx = document.getElementById('dashboard-pie-chart');
        if (catEntries.length > 0) {
            charts.dashPie = new Chart(pieCtx, {
                type: 'doughnut',
                data: {
                    labels: pieLabels,
                    datasets: [{ data: pieData, backgroundColor: pieColors.slice(0, pieData.length), borderWidth: 0 }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {
                        legend: { position: 'right', labels: { boxWidth: 12, font: { size: 12 } } }
                    },
                    cutout: '60%',
                }
            });
        } else {
            charts.dashPie = new Chart(pieCtx, {
                type: 'doughnut',
                data: { labels: ['暂无数据'], datasets: [{ data: [1], backgroundColor: ['rgba(0,0,0,0.05)'], borderWidth: 0 }] },
                options: { responsive: true, plugins: { legend: { display: false } }, cutout: '60%' }
            });
        }

        // Trend chart - last 7 days
        if (charts.dashTrend) charts.dashTrend.destroy();
        const days = [];
        for (let i = 6; i >= 0; i--) {
            const d = new Date(now);
            d.setDate(d.getDate() - i);
            days.push(formatDate(d.toISOString()));
        }
        const dayIncome = days.map(d => {
            return state.records.filter(r => r.date === d && r.type === 'income').reduce((s, r) => s + r.amount, 0);
        });
        const dayExpense = days.map(d => {
            return state.records.filter(r => r.date === d && r.type === 'expense').reduce((s, r) => s + calculateActualExpense(r), 0);
        });

        const trendCtx = document.getElementById('dashboard-trend-chart');
        charts.dashTrend = new Chart(trendCtx, {
            type: 'bar',
            data: {
                labels: days.map(d => d.slice(5)),
                datasets: [
                    { label: '收入', data: dayIncome, backgroundColor: 'rgba(52,199,89,0.75)', borderRadius: 6 },
                    { label: '实际支出', data: dayExpense, backgroundColor: 'rgba(255,59,48,0.65)', borderRadius: 6 },
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { legend: { labels: { boxWidth: 12, font: { size: 12 } } } },
                scales: {
                    x: { grid: { display: false }, ticks: { font: { size: 11 }, color: 'rgba(0,0,0,0.36)' } },
                    y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.04)' }, ticks: { font: { size: 11 }, color: 'rgba(0,0,0,0.36)' } },
                },
            }
        });
    }

    // ========== Record Item Render ==========
    function renderRecordItem(r, showActions = false) {
        const catInfo = getCategoryInfo(r.category, r.type);
        const isIncome = r.type === 'income';
        const actualAmt = r.type === 'expense' ? calculateActualExpense(r) : r.amount;
        const aaBadge = (r.aa && r.aa.enabled)
            ? `<span class="record-aa-badge">AA | 实付${formatMoney(r.aa.myShare)}</span>`
            : '';

        const actionsHtml = showActions && !state.batchMode ? `
            <div class="record-actions">
                <button onclick="app.editRecord('${r.id}')">编辑</button>
                <button class="delete" onclick="app.deleteRecord('${r.id}')">删除</button>
            </div>` : '';

        const checkboxHtml = state.batchMode ? `
            <label class="batch-checkbox" onclick="event.stopPropagation()">
                <input type="checkbox" ${state.selectedIds.has(r.id) ? 'checked' : ''}
                    onchange="app.toggleSelect('${r.id}', this.checked)">
                <span class="batch-checkmark"></span>
            </label>` : '';

        return `
            <div class="record-item ${state.batchMode ? 'batch-mode' : ''} ${state.selectedIds.has(r.id) ? 'batch-selected' : ''}">
                ${checkboxHtml}
                <div class="record-icon" style="background: ${isIncome ? 'var(--color-income-bg)' : 'var(--color-expense-bg)'}">${catInfo.icon}</div>
                <div class="record-info">
                    <div class="record-desc">${escapeHtml(r.description || catInfo.name)}</div>
                    <div class="record-meta">
                        <span>${escapeHtml(r.date)}</span>
                        <span>${escapeHtml(catInfo.name)}</span>
                        ${aaBadge}
                    </div>
                </div>
                <div class="record-amount ${r.type}">${isIncome ? '+' : '-'}${formatMoney(r.amount)}</div>
                ${actionsHtml}
            </div>`;
    }

    // ========== Add Record Form ==========
    function renderAddForm() {
        const type = state.recordFormType;
        renderCategoryGrid(type);
        document.getElementById('form-date').value = formatDate(new Date().toISOString());
        document.getElementById('aa-section').style.display = type === 'expense' ? 'block' : 'none';
        updateAAPreview();
    }

    function renderCategoryGrid(type) {
        const grid = document.getElementById('category-grid');
        const cats = state.categories[type] || [];
        grid.innerHTML = cats.map(c => `
            <div class="category-item ${state.selectedCategory === c.id ? 'selected' : ''}" data-id="${c.id}">
                <span class="cat-icon">${c.icon}</span>
                <span>${c.name}</span>
            </div>
        `).join('');

        grid.querySelectorAll('.category-item').forEach(el => {
            el.addEventListener('click', () => {
                state.selectedCategory = el.dataset.id;
                grid.querySelectorAll('.category-item').forEach(e => e.classList.remove('selected'));
                el.classList.add('selected');
            });
        });
    }

    // Find income records related to a group payment (AA collect)
    // Returns { recommended: close matches, all: all income records in window }
    function findMatchingIncomes(date, totalCollect, people) {
        if (!date || !totalCollect || totalCollect <= 0) return { recommended: [], all: [] };
        // Calculate 15-day window centered on the record date
        const recordDate = new Date(date + 'T00:00:00');
        const minDate = new Date(recordDate);
        minDate.setDate(minDate.getDate() - 15);
        const maxDate = new Date(recordDate);
        maxDate.setDate(maxDate.getDate() + 15);
        const minStr = formatDate(minDate.toISOString());
        const maxStr = formatDate(maxDate.toISOString());

        // Per-person share: what each other person should pay back
        const othersCount = (people && people > 1) ? (people - 1) : 1;
        const perPerson = Math.round(totalCollect / othersCount * 100) / 100;
        const threshold = Math.max(perPerson * 0.1, 0.5);

        // All income records in the 15-day window
        const allIncomes = state.records.filter(r =>
            r.type === 'income' &&
            r.date >= minStr &&
            r.date <= maxStr
        ).sort((a, b) => b.date.localeCompare(a.date) || b.createdAt - a.createdAt);

        // Recommended: amount close to per-person share
        const recommended = allIncomes.filter(r =>
            Math.abs(r.amount - perPerson) <= threshold
        );

        return { recommended, all: allIncomes };
    }

    function updateAAPreview() {
        const amount = parseFloat(document.getElementById('form-amount').value) || 0;
        const people = parseInt(document.getElementById('aa-people').value) || 2;
        const mode = document.querySelector('input[name="aa-mode"]:checked')?.value || 'equal';
        const preview = document.getElementById('aa-preview');
        const date = document.getElementById('form-date').value;

        let myShare;
        if (mode === 'equal') {
            myShare = amount > 0 ? Math.round(amount / people * 100) / 100 : 0;
        } else {
            myShare = parseFloat(document.getElementById('aa-my-share').value) || 0;
        }

        const othersTotal = Math.round((amount - myShare) * 100) / 100;

        // Build preview HTML
        let html = '';
        if (amount <= 0) {
            html = '<div style="color:var(--text-tertiary);">请先输入金额，系统将自动计算你的实际支出</div>';
            if (people > 1) {
                html += `<div style="margin-top:6px;font-size:12px;color:var(--text-tertiary);">当前 ${people} 人均摊模式</div>`;
            }
        } else {
            html = `
                <div>总金额: <strong>${formatMoney(amount)}</strong> | ${people}人参与</div>
                <div style="margin-top:8px;">你的实际支出: <span class="aa-highlight">${formatMoney(myShare)}</span></div>
                <div style="margin-top:4px;font-size:12px;color:var(--text-tertiary);">群收款可回收: ${formatMoney(othersTotal)}</div>
            `;
        }

        // Check for matching income records
        if (amount > 0 && othersTotal > 0 && date) {
            const { recommended, all } = findMatchingIncomes(date, othersTotal, people);
            if (recommended.length > 0) {
                html += `<div class="aa-match-warning">💡 发现 ${recommended.length} 条金额接近的收入记录，保存后可选择删除</div>`;
            } else if (all.length > 0) {
                html += `<div class="aa-match-hint">📋 近15天有 ${all.length} 条收入记录，保存后可手动选择删除</div>`;
            }
        }

        preview.innerHTML = html;
    }

    function handleFormSubmit(e) {
        e.preventDefault();
        const type = state.recordFormType;
        const amount = parseFloat(document.getElementById('form-amount').value);
        const category = state.selectedCategory;
        const desc = document.getElementById('form-desc').value.trim();
        const date = document.getElementById('form-date').value;

        if (!amount || amount <= 0) return toast('请输入有效金额', 'error');
        if (!category) return toast('请选择分类', 'error');
        if (!date) return toast('请选择日期', 'error');

        const record = {
            id: generateId(),
            type,
            amount,
            category,
            description: desc,
            date,
            createdAt: Date.now(),
        };

        // AA info
        if (type === 'expense' && document.getElementById('aa-enabled').checked) {
            const people = parseInt(document.getElementById('aa-people').value) || 2;
            const mode = document.querySelector('input[name="aa-mode"]:checked')?.value || 'equal';
            let myShare;
            if (mode === 'equal') {
                myShare = Math.round(amount / people * 100) / 100;
            } else {
                myShare = parseFloat(document.getElementById('aa-my-share').value) || 0;
            }
            record.aa = {
                enabled: true,
                people,
                mode,
                myShare,
                totalCollect: Math.round((amount - myShare) * 100) / 100,
            };
        }

        state.records.push(record);
        saveState();
        toast('记录成功！');

        // Reset form
        document.getElementById('form-amount').value = '';
        document.getElementById('form-desc').value = '';
        state.selectedCategory = null;
        document.getElementById('aa-enabled').checked = false;
        document.getElementById('aa-details').style.display = 'none';
        renderCategoryGrid(type);

        // After saving an AA record, always show income matching modal
        if (record.aa && record.aa.enabled && record.aa.totalCollect > 0) {
            const matchResult = findMatchingIncomes(record.date, record.aa.totalCollect, record.aa.people);
            showMatchedIncomesModal(record, matchResult);
        }
    }

    // Show modal to review and delete matching income records after AA save
    function showMatchedIncomesModal(aaRecord, matchResult) {
        const totalCollect = aaRecord.aa.totalCollect;
        const { recommended, all } = matchResult;

        // Build recommended section
        let recommendedHtml = '';
        if (recommended.length > 0) {
            recommendedHtml = `
                <div class="aa-match-section">
                    <div class="aa-match-section-title">💡 推荐匹配（金额接近 ¥${(totalCollect / ((aaRecord.aa.people || 2) - 1)).toFixed(0)}）</div>
                    <div class="aa-match-list">${recommended.map(r => renderMatchItem(r)).join('')}</div>
                </div>
            `;
        }

        // Build all incomes section
        const otherIncomes = all.filter(r => !recommended.find(rec => rec.id === r.id));
        let allHtml = '';
        if (all.length > 0) {
            allHtml = `
                <div class="aa-match-section">
                    <div class="aa-match-section-title">📋 近期全部收入记录（前后15天）</div>
                    <div class="aa-match-list">${all.map(r => renderMatchItem(r)).join('')}</div>
                </div>
            `;
        }

        const noMatchHint = (recommended.length === 0 && all.length === 0)
            ? '<div class="aa-match-empty">近15天没有收入记录</div>'
            : '';

        showModal('清理群收款收入', `
            <div class="aa-match-explain">
                你刚记录了一笔 AA 垫付支出，群收款可回收 <strong>${formatMoney(totalCollect)}</strong>。
                选择需要删除的收入记录以避免虚增收支，没有则直接关闭。
            </div>
            ${recommendedHtml}
            ${allHtml}
            ${noMatchHint}
            <div class="modal-actions">
                <button class="btn-secondary" onclick="app.closeModal()">关闭</button>
                <button class="btn-primary" onclick="app.deleteMatchedIncomes()">删除选中的收入</button>
            </div>
        `);
    }

    function renderMatchItem(r) {
        const catInfo = getCategoryInfo(r.category, 'income');
        return `
            <div class="aa-match-item" data-id="${r.id}">
                <label class="aa-match-check">
                    <input type="checkbox" onchange="app.toggleMatchSelect('${r.id}', this.checked)">
                    <span class="batch-checkmark"></span>
                </label>
                <div class="aa-match-info">
                    <div class="aa-match-desc">${escapeHtml(r.description || catInfo.name)}</div>
                    <div class="aa-match-meta">${escapeHtml(r.date)} · ${catInfo.icon} ${escapeHtml(catInfo.name)}</div>
                </div>
                <div class="aa-match-amount">+${formatMoney(r.amount)}</div>
            </div>
        `;
    }

    function toggleMatchSelect(id, checked) {
        // Visual toggle handled by checkbox state
    }

    function deleteMatchedIncomes() {
        const checkboxes = document.querySelectorAll('.aa-match-item input[type="checkbox"]');
        let deleted = 0;
        checkboxes.forEach(cb => {
            if (cb.checked) {
                const item = cb.closest('.aa-match-item');
                const id = item.dataset.id;
                state.records = state.records.filter(r => r.id !== id);
                deleted++;
            }
        });
        saveState();
        closeModal();
        if (deleted > 0) {
            toast(`已删除 ${deleted} 条匹配的收入记录`);
        }
    }

    // ========== Records Page ==========
    function renderRecords() {
        const typeFilter = document.getElementById('filter-type').value;
        const catFilter = document.getElementById('filter-category').value;
        const monthFilter = document.getElementById('filter-month').value;
        const searchFilter = document.getElementById('filter-search').value.toLowerCase();

        // Update category filter options
        updateCategoryFilterOptions();

        // Build cache key from filter state + records count to detect data changes
        const cacheKey = `${typeFilter}|${catFilter}|${monthFilter}|${searchFilter}|${state.records.length}`;
        let filtered;
        if (_filterCache.key === cacheKey) {
            filtered = _filterCache.result;
        } else {
            filtered = [...state.records];
            if (typeFilter !== 'all') filtered = filtered.filter(r => r.type === typeFilter);
            if (catFilter !== 'all') filtered = filtered.filter(r => r.category === catFilter);
            if (monthFilter) filtered = filtered.filter(r => r.date.startsWith(monthFilter));
            if (searchFilter) filtered = filtered.filter(r => (r.description || '').toLowerCase().includes(searchFilter));
            filtered.sort((a, b) => b.date.localeCompare(a.date) || b.createdAt - a.createdAt);
            _filterCache = { key: cacheKey, result: filtered };
        }

        // Pagination
        const total = filtered.length;
        const totalPages = Math.max(1, Math.ceil(total / state.recordsPerPage));
        state.recordsPage = Math.min(state.recordsPage, totalPages);
        const start = (state.recordsPage - 1) * state.recordsPerPage;
        const pageRecords = filtered.slice(start, start + state.recordsPerPage);

        const listEl = document.getElementById('records-list');
        if (pageRecords.length === 0) {
            listEl.innerHTML = '<div class="empty-state"><div class="empty-icon">🔍</div><p>没有找到匹配的记录</p></div>';
        } else {
            listEl.innerHTML = pageRecords.map(r => renderRecordItem(r, true)).join('');
        }

        // Pagination buttons
        const pagEl = document.getElementById('records-pagination');
        if (totalPages <= 1) {
            pagEl.innerHTML = '';
        } else {
            let btns = '';
            for (let i = 1; i <= totalPages; i++) {
                btns += `<button class="${i === state.recordsPage ? 'active' : ''}" onclick="app.goToPage(${i})">${i}</button>`;
            }
            pagEl.innerHTML = btns;
        }
    }

    function updateCategoryFilterOptions() {
        const select = document.getElementById('filter-category');
        const current = select.value;
        const allCats = [...state.categories.expense, ...state.categories.income];
        select.innerHTML = '<option value="all">全部分类</option>' +
            allCats.map(c => `<option value="${c.id}">${c.icon} ${c.name}</option>`).join('');
        select.value = current || 'all';
    }

    function goToPage(page) {
        state.recordsPage = page;
        renderRecords();
    }

    function editRecord(id) {
        const record = state.records.find(r => r.id === id);
        if (!record) return;

        const allCats = state.categories[record.type] || [];
        const catOptions = allCats.map(c =>
            `<option value="${c.id}" ${c.id === record.category ? 'selected' : ''}>${c.icon} ${c.name}</option>`
        ).join('');

        const aaHtml = record.type === 'expense' ? `
            <div class="form-group">
                <label>
                    <input type="checkbox" id="edit-aa-enabled" ${record.aa?.enabled ? 'checked' : ''}>
                    AA制/群收款
                </label>
            </div>
            <div id="edit-aa-fields" style="${record.aa?.enabled ? '' : 'display:none'}">
                <div class="form-group">
                    <label>参与人数（含自己）</label>
                    <input type="number" id="edit-aa-people" min="2" value="${record.aa?.people || 2}">
                </div>
                <div class="form-group">
                    <label>分摊方式</label>
                    <div class="radio-group">
                        <label><input type="radio" name="edit-aa-mode" value="equal" ${(!record.aa || record.aa.mode === 'equal') ? 'checked' : ''}> 平均分摊</label>
                        <label><input type="radio" name="edit-aa-mode" value="custom" ${record.aa?.mode === 'custom' ? 'checked' : ''}> 自定义我的份额</label>
                    </div>
                </div>
                <div class="form-group" id="edit-aa-custom-amount" style="${record.aa?.mode === 'custom' ? '' : 'display:none'}">
                    <label>我实际应付</label>
                    <input type="number" id="edit-aa-myshare" step="0.01" value="${record.aa?.myShare || 0}">
                </div>
                <div class="aa-preview" id="edit-aa-preview"></div>
            </div>
        ` : '';

        showModal('编辑记录', `
            <div class="form-group">
                <label>金额</label>
                <input type="number" id="edit-amount" step="0.01" value="${record.amount}">
            </div>
            <div class="form-group">
                <label>分类</label>
                <select id="edit-category">${catOptions}</select>
            </div>
            <div class="form-group">
                <label>描述</label>
                <input type="text" id="edit-desc" value="${record.description || ''}">
            </div>
            <div class="form-group">
                <label>日期</label>
                <input type="date" id="edit-date" value="${record.date}">
            </div>
            ${aaHtml}
            <div class="modal-actions">
                <button class="btn-secondary" onclick="app.closeModal()">取消</button>
                <button class="btn-primary" onclick="app.saveEditRecord('${id}')">保存</button>
            </div>
        `);

        // AA toggle
        const aaCheck = document.getElementById('edit-aa-enabled');
        if (aaCheck) {
            aaCheck.addEventListener('change', () => {
                document.getElementById('edit-aa-fields').style.display = aaCheck.checked ? '' : 'none';
                updateEditAAPreview();
            });

            // Bind real-time calculation for edit AA fields
            const editAmountEl = document.getElementById('edit-amount');
            const editPeopleEl = document.getElementById('edit-aa-people');
            const editMyShareEl = document.getElementById('edit-aa-myshare');
            const editDateEl = document.getElementById('edit-date');

            if (editAmountEl) editAmountEl.addEventListener('input', updateEditAAPreview);
            if (editPeopleEl) editPeopleEl.addEventListener('input', updateEditAAPreview);
            if (editMyShareEl) editMyShareEl.addEventListener('input', updateEditAAPreview);
            if (editDateEl) editDateEl.addEventListener('input', updateEditAAPreview);

            // Radio mode toggle
            document.querySelectorAll('input[name="edit-aa-mode"]').forEach(el => {
                el.addEventListener('change', () => {
                    const customDiv = document.getElementById('edit-aa-custom-amount');
                    if (customDiv) {
                        customDiv.style.display = el.value === 'custom' ? '' : 'none';
                    }
                    updateEditAAPreview();
                });
            });

            // Initial preview
            updateEditAAPreview();
        }
    }

    // Real-time AA preview for edit modal
    function updateEditAAPreview() {
        const preview = document.getElementById('edit-aa-preview');
        if (!preview) return;

        const amount = parseFloat(document.getElementById('edit-amount')?.value) || 0;
        const people = parseInt(document.getElementById('edit-aa-people')?.value) || 2;
        const mode = document.querySelector('input[name="edit-aa-mode"]:checked')?.value || 'equal';
        const date = document.getElementById('edit-date')?.value || '';

        let myShare;
        if (mode === 'equal') {
            myShare = amount > 0 ? Math.round(amount / people * 100) / 100 : 0;
        } else {
            myShare = parseFloat(document.getElementById('edit-aa-myshare')?.value) || 0;
        }
        const othersTotal = Math.round((amount - myShare) * 100) / 100;

        let html = '';
        if (amount <= 0) {
            html = '<div style="color:var(--text-tertiary);">请输入金额</div>';
        } else {
            html = `
                <div>总金额: <strong>${formatMoney(amount)}</strong> | ${people}人参与</div>
                <div style="margin-top:8px;">你的实际支出: <span class="aa-highlight">${formatMoney(myShare)}</span></div>
                <div style="margin-top:4px;font-size:12px;color:var(--text-tertiary);">群收款可回收: ${formatMoney(othersTotal)}</div>
            `;
        }

        // Check for matching income records
        if (amount > 0 && othersTotal > 0 && date) {
            const { recommended, all } = findMatchingIncomes(date, othersTotal, people);
            if (recommended.length > 0) {
                html += `<div class="aa-match-warning">💡 发现 ${recommended.length} 条金额接近的收入记录，保存后可选择删除</div>`;
            } else if (all.length > 0) {
                html += `<div class="aa-match-hint">📋 近15天有 ${all.length} 条收入记录，保存后可手动选择删除</div>`;
            }
        }

        preview.innerHTML = html;

        // Auto-fill myShare in equal mode
        if (mode === 'equal' && amount > 0) {
            const myShareInput = document.getElementById('edit-aa-myshare');
            if (myShareInput) myShareInput.value = myShare;
        }
    }

    function saveEditRecord(id) {
        const record = state.records.find(r => r.id === id);
        if (!record) return;

        record.amount = parseFloat(document.getElementById('edit-amount').value) || 0;
        record.category = document.getElementById('edit-category').value;
        record.description = document.getElementById('edit-desc').value.trim();
        record.date = document.getElementById('edit-date').value;

        const aaCheck = document.getElementById('edit-aa-enabled');
        if (aaCheck) {
            if (aaCheck.checked) {
                const people = parseInt(document.getElementById('edit-aa-people').value) || 2;
                const mode = document.querySelector('input[name="edit-aa-mode"]:checked')?.value || 'equal';
                let myShare;
                if (mode === 'equal') {
                    myShare = Math.round(record.amount / people * 100) / 100;
                } else {
                    myShare = parseFloat(document.getElementById('edit-aa-myshare').value) || 0;
                }
                record.aa = {
                    enabled: true,
                    people,
                    mode,
                    myShare,
                    totalCollect: Math.round((record.amount - myShare) * 100) / 100,
                };
            } else {
                record.aa = null;
            }
        }

        saveState();
        closeModal();
        renderRecords();
        toast('修改成功');

        // After editing a record with AA enabled, always show income matching modal
        if (record.aa && record.aa.enabled && record.aa.totalCollect > 0) {
            const matchResult = findMatchingIncomes(record.date, record.aa.totalCollect, record.aa.people);
            showMatchedIncomesModal(record, matchResult);
        }
    }

    function deleteRecord(id) {
        if (!confirm('确定要删除这条记录吗？')) return;
        state.records = state.records.filter(r => r.id !== id);
        saveState();
        renderRecords();
        toast('已删除');
    }

    // ========== Batch Operations ==========
    function toggleBatchMode() {
        state.batchMode = !state.batchMode;
        state.selectedIds.clear();
        document.getElementById('batch-toolbar').style.display = state.batchMode ? 'flex' : 'none';
        document.getElementById('btn-batch-toggle').textContent = state.batchMode ? '退出批量' : '批量操作';
        document.getElementById('btn-batch-toggle').classList.toggle('active', state.batchMode);
        document.getElementById('batch-select-all').checked = false;
        updateBatchCount();
        renderRecords();
    }

    function toggleSelect(id, checked) {
        if (checked) {
            state.selectedIds.add(id);
        } else {
            state.selectedIds.delete(id);
        }
        updateBatchCount();
        // Update visual selection state without full re-render
        const items = document.querySelectorAll('.record-item');
        items.forEach(el => {
            const cb = el.querySelector('.batch-checkbox input[type="checkbox"]');
            if (cb) {
                const rid = cb.getAttribute('onchange').match(/'([^']+)'/)?.[1];
                if (rid) el.classList.toggle('batch-selected', state.selectedIds.has(rid));
            }
        });
        // Sync select-all checkbox
        syncSelectAllCheckbox();
    }

    function toggleSelectAll(checked) {
        // Get current page's record IDs from the rendered checkboxes
        const checkboxes = document.querySelectorAll('.batch-checkbox input[type="checkbox"]');
        checkboxes.forEach(cb => {
            const rid = cb.getAttribute('onchange').match(/'([^']+)'/)?.[1];
            if (rid) {
                if (checked) {
                    state.selectedIds.add(rid);
                } else {
                    state.selectedIds.delete(rid);
                }
                cb.checked = checked;
            }
        });
        // Update visual state
        document.querySelectorAll('.record-item').forEach(el => {
            el.classList.toggle('batch-selected', checked);
        });
        updateBatchCount();
    }

    function syncSelectAllCheckbox() {
        const checkboxes = document.querySelectorAll('.batch-checkbox input[type="checkbox"]');
        if (checkboxes.length === 0) return;
        const allChecked = Array.from(checkboxes).every(cb => cb.checked);
        document.getElementById('batch-select-all').checked = allChecked;
    }

    function updateBatchCount() {
        const count = state.selectedIds.size;
        document.getElementById('batch-count').textContent = `已选 ${count} 条`;
    }

    function showBatchCategory() {
        if (state.selectedIds.size === 0) {
            return toast('请先选择要修改的记录', 'error');
        }

        // Figure out which types are selected to show appropriate category lists
        const selectedRecords = state.records.filter(r => state.selectedIds.has(r.id));
        const hasExpense = selectedRecords.some(r => r.type === 'expense');
        const hasIncome = selectedRecords.some(r => r.type === 'income');

        let catListHtml = '';
        if (hasExpense) {
            catListHtml += `<div class="batch-cat-section"><div class="batch-cat-section-label">支出分类</div><div class="batch-cat-grid">`;
            catListHtml += state.categories.expense.map(c =>
                `<div class="batch-cat-item" data-id="${c.id}" data-type="expense" onclick="app.selectBatchCat(this)">
                    <span class="cat-icon">${c.icon}</span><span>${escapeHtml(c.name)}</span>
                </div>`
            ).join('');
            catListHtml += `</div></div>`;
        }
        if (hasIncome) {
            catListHtml += `<div class="batch-cat-section"><div class="batch-cat-section-label">收入分类</div><div class="batch-cat-grid">`;
            catListHtml += state.categories.income.map(c =>
                `<div class="batch-cat-item" data-id="${c.id}" data-type="income" onclick="app.selectBatchCat(this)">
                    <span class="cat-icon">${c.icon}</span><span>${escapeHtml(c.name)}</span>
                </div>`
            ).join('');
            catListHtml += `</div></div>`;
        }

        const mixedWarning = (hasExpense && hasIncome)
            ? `<div class="batch-mixed-warning">你选中的记录包含收支两种类型，请分别选择对应的分类。收入记录只会应用收入分类，支出记录只会应用支出分类。</div>`
            : '';

        showModal(`批量修改分类（${state.selectedIds.size} 条）`, `
            ${mixedWarning}
            ${catListHtml}
            <div class="modal-actions">
                <button class="btn-secondary" onclick="app.closeModal()">取消</button>
                <button class="btn-primary" id="batch-cat-confirm" onclick="app.confirmBatchCategory()" disabled>确认修改</button>
            </div>
        `);
    }

    let _batchSelectedCat = { expense: null, income: null };

    function selectBatchCat(el) {
        const catId = el.dataset.id;
        const catType = el.dataset.type;
        // Deselect siblings in the same section
        el.parentElement.querySelectorAll('.batch-cat-item').forEach(item => item.classList.remove('selected'));
        el.classList.add('selected');
        _batchSelectedCat[catType] = catId;
        // Enable confirm button if at least one type is selected
        document.getElementById('batch-cat-confirm').disabled = false;
    }

    function confirmBatchCategory() {
        const expCat = _batchSelectedCat.expense;
        const incCat = _batchSelectedCat.income;
        if (!expCat && !incCat) {
            return toast('请选择一个分类', 'error');
        }

        let modified = 0;
        for (const r of state.records) {
            if (!state.selectedIds.has(r.id)) continue;
            if (r.type === 'expense' && expCat) {
                r.category = expCat;
                modified++;
            } else if (r.type === 'income' && incCat) {
                r.category = incCat;
                modified++;
            }
        }

        saveState();
        closeModal();
        _batchSelectedCat = { expense: null, income: null };
        state.selectedIds.clear();
        updateBatchCount();
        document.getElementById('batch-select-all').checked = false;
        renderRecords();
        toast(`已修改 ${modified} 条记录的分类`);
    }

    function batchDelete() {
        const count = state.selectedIds.size;
        if (count === 0) return toast('请先选择要删除的记录', 'error');
        if (!confirm(`确定要删除选中的 ${count} 条记录吗？`)) return;

        state.records = state.records.filter(r => !state.selectedIds.has(r.id));
        saveState();
        state.selectedIds.clear();
        updateBatchCount();
        document.getElementById('batch-select-all').checked = false;
        renderRecords();
        toast(`已删除 ${count} 条记录`);
    }

    // ========== Excel Import ==========
    function initImport() {
        const dropZone = document.getElementById('drop-zone');
        const fileInput = document.getElementById('file-input');

        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('drag-over');
        });
        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('drag-over');
        });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            const file = e.dataTransfer.files[0];
            if (file) processFile(file);
        });
        fileInput.addEventListener('change', (e) => {
            if (e.target.files[0]) processFile(e.target.files[0]);
        });
    }

    function processFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];

                // Try to find header row (skip metadata rows in WeChat/Alipay bills)
                const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                // Find the header row - look for rows with date/amount-like headers
                let headerRowIdx = 0;
                for (let i = 0; i < Math.min(rawRows.length, 20); i++) {
                    const row = rawRows[i];
                    if (!row) continue;
                    const rowStr = row.join(',').toLowerCase();
                    if (rowStr.includes('日期') || rowStr.includes('时间') ||
                        rowStr.includes('金额') || rowStr.includes('交易') ||
                        rowStr.includes('date') || rowStr.includes('amount')) {
                        headerRowIdx = i;
                        break;
                    }
                }

                const headers = rawRows[headerRowIdx] || [];
                const rows = rawRows.slice(headerRowIdx + 1).filter(r => r && r.length > 0 && r.some(cell => cell !== null && cell !== undefined && cell !== ''));

                state.importHeaders = headers.map(h => String(h || '').trim());
                state.importData = rows;
                state.importWorkbook = workbook;

                // Detect bill platform
                state.importPlatform = detectBillPlatform(state.importHeaders);

                // Auto-map columns
                autoMapColumns(state.importHeaders);

                // Show mapping UI
                document.getElementById('drop-zone').style.display = 'none';
                document.getElementById('import-mapping').style.display = 'block';

                renderImportMapping();
                renderImportPreview();
            } catch (err) {
                console.error(err);
                toast('文件解析失败，请检查文件格式', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // Detect if the bill comes from WeChat, Alipay, or unknown
    function detectBillPlatform(headers) {
        const joined = headers.join(',');
        if (joined.includes('交易对方') && joined.includes('商品') && joined.includes('收/支')) {
            return 'wechat';
        }
        // Future: Alipay detection
        // if (joined.includes('交易对方') && joined.includes('商品说明') && joined.includes('收/付款方式'))
        //     return 'alipay';
        return 'generic';
    }

    function autoMapColumns(headers) {
        const mapping = { date: -1, amount: -1, description: -1, type: -1, category: -1 };
        const lowerHeaders = headers.map(h => h.toLowerCase());
        const platform = state.importPlatform;

        if (platform === 'wechat') {
            // WeChat: precise mapping based on known column structure
            // 交易时间 | 交易类型 | 交易对方 | 商品 | 收/支 | 金额(元) | 支付方式 | 当前状态 | ...
            for (let i = 0; i < lowerHeaders.length; i++) {
                const h = lowerHeaders[i];
                if (mapping.date === -1 && h.includes('交易时间')) mapping.date = i;
                if (mapping.amount === -1 && h.includes('金额')) mapping.amount = i;
                if (mapping.type === -1 && h === '收/支') mapping.type = i;
            }
            // For WeChat, description is synthesized from 交易对方+商品 in confirmImport,
            // but we still map it to 交易对方 for the preview
            for (let i = 0; i < lowerHeaders.length; i++) {
                if (mapping.description === -1 && lowerHeaders[i].includes('交易对方')) mapping.description = i;
            }
            // Store extra column indices for WeChat-specific merging
            state.importWechatCols = {
                counterparty: headers.findIndex(h => h === '交易对方'),
                product: headers.findIndex(h => h === '商品'),
                txType: headers.findIndex(h => h === '交易类型'),
                status: headers.findIndex(h => h === '当前状态'),
                inOut: headers.findIndex(h => h === '收/支'),
            };
        } else {
            // Generic / Alipay fallback
            for (let i = 0; i < lowerHeaders.length; i++) {
                const h = lowerHeaders[i];
                if (mapping.date === -1 && (h.includes('日期') || h.includes('时间') || h.includes('date') || h.includes('交易时间'))) {
                    mapping.date = i;
                }
                if (mapping.amount === -1 && (h.includes('金额') || h.includes('amount') || h.includes('支出') || h.includes('价格'))) {
                    mapping.amount = i;
                }
                if (mapping.description === -1 && (h.includes('描述') || h.includes('商品') || h.includes('备注') || h.includes('说明') || h.includes('对方') || h.includes('交易对方') || h.includes('商户') || h.includes('description'))) {
                    mapping.description = i;
                }
                if (mapping.type === -1 && (h.includes('类型') || h.includes('收/支') || h.includes('收支') || h.includes('type'))) {
                    mapping.type = i;
                }
                if (mapping.category === -1 && (h.includes('分类') || h.includes('category') || h.includes('交易类型'))) {
                    mapping.category = i;
                }
            }
        }

        state.importMapping = mapping;
    }

    function renderImportMapping() {
        const grid = document.getElementById('mapping-grid');
        const fields = [
            { key: 'date', label: '日期 *', required: true },
            { key: 'amount', label: '金额 *', required: true },
            { key: 'description', label: '描述/商户名', required: false },
            { key: 'type', label: '收入/支出类型', required: false },
            { key: 'category', label: '分类', required: false },
        ];

        grid.innerHTML = fields.map(f => {
            const options = ['<option value="-1">不映射</option>']
                .concat(state.importHeaders.map((h, i) => `<option value="${i}" ${state.importMapping[f.key] === i ? 'selected' : ''}>${h}</option>`))
                .join('');
            return `
                <div class="mapping-label">${f.label}</div>
                <div class="mapping-arrow">→</div>
                <select data-field="${f.key}" onchange="app.updateMapping(this)">${options}</select>
            `;
        }).join('');
    }

    function updateMapping(selectEl) {
        state.importMapping[selectEl.dataset.field] = parseInt(selectEl.value);
        renderImportPreview();
    }

    // Build a unified description string for one row, handling WeChat merging
    function buildRowDescription(row, m) {
        if (state.importPlatform === 'wechat' && state.importWechatCols) {
            const wc = state.importWechatCols;
            const counterparty = wc.counterparty >= 0 ? String(row[wc.counterparty] || '').trim() : '';
            const product = wc.product >= 0 ? String(row[wc.product] || '').trim() : '';
            // Filter out placeholder '/' values
            const parts = [counterparty, product].filter(p => p && p !== '/');
            return parts.join(' - ') || '未知交易';
        }
        return m.description >= 0 ? String(row[m.description] || '').trim() : '';
    }

    // Determine if a WeChat row should be skipped (non-income/expense transactions)
    function shouldSkipWechatRow(row) {
        if (state.importPlatform !== 'wechat' || !state.importWechatCols) return false;
        const wc = state.importWechatCols;
        // Skip rows where 收/支 is '/' (transfers, top-ups, withdrawals, etc.)
        if (wc.inOut >= 0) {
            const inOutVal = String(row[wc.inOut] || '').trim();
            if (inOutVal === '/' || inOutVal === '') return true;
        }
        // Skip fully refunded rows
        if (wc.status >= 0) {
            const statusVal = String(row[wc.status] || '').trim();
            if (statusVal.includes('已全额退款')) return true;
        }
        return false;
    }

    // Parse the 收/支 type from a row
    function parseRowType(row, m) {
        if (m.type >= 0) {
            const typeVal = String(row[m.type] || '').trim();
            if (typeVal.includes('收入') || typeVal.includes('income') || typeVal === '收') return 'income';
            if (typeVal.includes('支出') || typeVal.includes('expense') || typeVal === '支') return 'expense';
        }
        return 'expense'; // default
    }

    function renderImportPreview() {
        const preview = document.getElementById('import-preview-table');
        const m = state.importMapping;
        const isWechat = state.importPlatform === 'wechat';

        // For preview, pick first 5 displayable rows (skip non-收支 for WeChat)
        const previewRows = [];
        for (const row of state.importData) {
            if (previewRows.length >= 5) break;
            if (isWechat && shouldSkipWechatRow(row)) continue;
            previewRows.push(row);
        }

        let platformBadge = '';
        if (isWechat) {
            platformBadge = '<div class="import-platform-badge">已识别为 <strong>微信支付</strong> 账单，将自动合并「交易对方」+「商品」为描述，并跳过非收支记录</div>';
        }

        let html = platformBadge + '<table><thead><tr><th>日期</th><th>金额</th><th>描述</th><th>类型</th><th>自动分类</th></tr></thead><tbody>';
        for (const row of previewRows) {
            const date = m.date >= 0 ? String(row[m.date] || '') : '-';
            const amount = m.amount >= 0 ? row[m.amount] : '-';
            const desc = escapeHtml(buildRowDescription(row, m));
            const type = parseRowType(row, m);
            const typeLabel = type === 'income' ? '收入' : '支出';

            const classified = autoClassify(desc, type);
            const catName = classified ? getCategoryInfo(classified.category, classified.type).name : '未分类';

            html += `<tr><td>${date}</td><td>${amount}</td><td>${desc}</td><td>${typeLabel}</td><td>${catName}</td></tr>`;
        }
        html += '</tbody></table>';
        preview.innerHTML = html;
    }

    function confirmImport() {
        const m = state.importMapping;
        if (m.date < 0 || m.amount < 0) {
            return toast('请至少映射日期和金额列', 'error');
        }

        const isWechat = state.importPlatform === 'wechat';
        let imported = 0;
        let skipped = 0;

        for (const row of state.importData) {
            try {
                // WeChat: skip non-income/expense rows (transfers, top-ups, withdrawals, refunded)
                if (isWechat && shouldSkipWechatRow(row)) { skipped++; continue; }

                // Parse date
                let dateRaw = row[m.date];
                if (!dateRaw) { skipped++; continue; }
                let dateStr = parseDateValue(dateRaw, state.importWorkbook);
                if (!dateStr) { skipped++; continue; }

                // Parse amount
                let amountRaw = row[m.amount];
                let amount = parseFloat(String(amountRaw).replace(/[¥￥,，\s]/g, ''));
                if (isNaN(amount) || amount === 0) { skipped++; continue; }

                // Parse type
                let type = parseRowType(row, m);

                // Negative amount handling
                if (amount < 0) {
                    amount = Math.abs(amount);
                    if (m.type < 0) type = 'expense';
                }

                // Description: use merged version for WeChat
                const desc = buildRowDescription(row, m);

                // Auto classify
                let category = type === 'income' ? 'other_income' : 'other_expense';
                if (m.category >= 0 && row[m.category]) {
                    const catVal = String(row[m.category]).trim();
                    const found = [...state.categories.expense, ...state.categories.income].find(c => c.name === catVal);
                    if (found) category = found.id;
                }
                const classified = autoClassify(desc, type);
                if (classified) {
                    category = classified.category;
                    if (!type || type === 'expense') type = classified.type;
                }
                if (type === 'income' && category.includes('expense')) {
                    category = 'other_income';
                }

                const record = {
                    id: generateId(),
                    type,
                    amount,
                    category,
                    description: desc,
                    date: dateStr,
                    createdAt: Date.now(),
                    source: isWechat ? 'wechat' : 'import',
                };

                state.records.push(record);
                imported++;
            } catch (e) {
                skipped++;
            }
        }

        saveState();
        toast(`成功导入 ${imported} 条记录${skipped > 0 ? `，跳过 ${skipped} 条` : ''}`, 'success');

        // Reset import UI
        cancelImport();
    }

    function parseDateValue(val, wb) {
        if (!val) return null;
        // Excel serial number (supports both 1900 and 1904 date systems)
        if (typeof val === 'number') {
            // 1904 date system (common in Mac-generated Excel): epoch starts 1904-01-01
            // 1900 date system (default): epoch offset is 25569 days from Unix epoch
            const epoch1904 = wb && wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904;
            const offset = epoch1904 ? 24107 : 25569;
            const date = new Date((val - offset) * 86400 * 1000);
            if (!isNaN(date.getTime())) return formatDate(date.toISOString());
        }
        const str = String(val).trim();
        // Try common formats
        // 2024-01-15, 2024/01/15, 2024.01.15
        let match = str.match(/(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
        if (match) {
            return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
        }
        // 01-15-2024 or 01/15/2024
        match = str.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
        if (match) {
            return `${match[3]}-${match[1].padStart(2, '0')}-${match[2].padStart(2, '0')}`;
        }
        // Try Date parse
        const d = new Date(str);
        if (!isNaN(d.getTime())) return formatDate(d.toISOString());
        return null;
    }

    function cancelImport() {
        state.importData = null;
        state.importHeaders = null;
        state.importMapping = {};
        state.importWorkbook = null;
        state.importPlatform = null;
        state.importWechatCols = null;
        document.getElementById('drop-zone').style.display = '';
        document.getElementById('import-mapping').style.display = 'none';
        document.getElementById('file-input').value = '';
    }

    // ========== Summary Page ==========
    function renderSummary() {
        const period = state.summaryPeriod;
        const date = state.summaryDate;

        // Period label
        const label = getPeriodLabel(period, date);
        document.getElementById('summary-period-label').textContent = label;

        // Get records for period
        const { start, end } = getPeriodRange(period, date);
        const periodRecords = state.records.filter(r => r.date >= start && r.date <= end);
        const totals = calcTotals(periodRecords);

        document.getElementById('summary-income').textContent = formatMoney(totals.income);
        document.getElementById('summary-expense').textContent = formatMoney(totals.expense);
        document.getElementById('summary-actual').textContent = formatMoney(totals.actualExpense);
        document.getElementById('summary-balance').textContent = formatMoney(totals.balance);

        renderSummaryCharts(periodRecords, period, start, end);
        renderSummaryCategoryDetail(periodRecords);
    }

    function getPeriodLabel(period, date) {
        const y = date.getFullYear();
        const m = date.getMonth() + 1;
        const d = date.getDate();
        switch (period) {
            case 'day': return `${y}年${m}月${d}日`;
            case 'week': {
                const ws = getWeekStart(date);
                const we = new Date(ws);
                we.setDate(we.getDate() + 6);
                return `${formatDate(ws.toISOString())} ~ ${formatDate(we.toISOString())}`;
            }
            case 'month': return `${y}年${m}月`;
            case 'year': return `${y}年`;
        }
    }

    function getPeriodRange(period, date) {
        const y = date.getFullYear();
        const m = date.getMonth();
        const d = date.getDate();
        switch (period) {
            case 'day': {
                const s = formatDate(date.toISOString());
                return { start: s, end: s };
            }
            case 'week': {
                const ws = getWeekStart(date);
                const we = new Date(ws);
                we.setDate(we.getDate() + 6);
                return { start: formatDate(ws.toISOString()), end: formatDate(we.toISOString()) };
            }
            case 'month': {
                const s = `${y}-${String(m + 1).padStart(2, '0')}-01`;
                const lastDay = new Date(y, m + 1, 0).getDate();
                const e = `${y}-${String(m + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
                return { start: s, end: e };
            }
            case 'year': {
                return { start: `${y}-01-01`, end: `${y}-12-31` };
            }
        }
    }

    // Week starts on Monday (ISO standard)
    function getWeekStart(date) {
        const d = new Date(date);
        const day = d.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
        const diff = d.getDate() - day + (day === 0 ? -6 : 1); // Roll back to Monday
        return new Date(d.setDate(diff));
    }

    function navigateSummary(dir) {
        const d = state.summaryDate;
        switch (state.summaryPeriod) {
            case 'day':
                d.setDate(d.getDate() + dir);
                break;
            case 'week':
                d.setDate(d.getDate() + dir * 7);
                break;
            case 'month':
                d.setMonth(d.getMonth() + dir);
                break;
            case 'year':
                d.setFullYear(d.getFullYear() + dir);
                break;
        }
        renderSummary();
    }

    function renderSummaryCharts(records, period, start, end) {
        const expenseRecords = records.filter(r => r.type === 'expense');
        const catTotals = {};
        for (const r of expenseRecords) {
            const actual = calculateActualExpense(r);
            catTotals[r.category] = (catTotals[r.category] || 0) + actual;
        }

        const catEntries = Object.entries(catTotals).sort((a, b) => b[1] - a[1]);
        const pieLabels = catEntries.map(([catId]) => getCategoryInfo(catId, 'expense').name);
        const pieData = catEntries.map(([, val]) => val);
        const pieColors = [
            '#5B7FFF', '#66D4A0', '#FF8A80', '#9B8FFF', '#FFB74D',
            '#CE93D8', '#80DEEA', '#F48FB1', '#81D4FA', '#FFF176',
            '#BCAAA4', '#B0BEC5'
        ];

        // Pie chart
        if (charts.summaryPie) charts.summaryPie.destroy();
        const pieCtx = document.getElementById('summary-category-chart');
        if (catEntries.length > 0) {
            charts.summaryPie = new Chart(pieCtx, {
                type: 'doughnut',
                data: {
                    labels: pieLabels,
                    datasets: [{ data: pieData, backgroundColor: pieColors.slice(0, pieData.length), borderWidth: 0 }]
                },
                options: { responsive: true, plugins: { legend: { position: 'right', labels: { boxWidth: 12, font: { size: 12 } } } }, cutout: '60%' }
            });
        } else {
            charts.summaryPie = new Chart(pieCtx, {
                type: 'doughnut',
                data: { labels: ['暂无数据'], datasets: [{ data: [1], backgroundColor: ['rgba(0,0,0,0.05)'], borderWidth: 0 }] },
                options: { responsive: true, plugins: { legend: { display: false } }, cutout: '60%' }
            });
        }

        // Trend chart
        if (charts.summaryTrend) charts.summaryTrend.destroy();
        const trendCtx = document.getElementById('summary-trend-chart');

        const timePoints = getTimePoints(period, start, end);
        const incomeData = timePoints.map(tp => {
            const recs = records.filter(r => r.type === 'income' && matchTimePoint(r.date, tp, period));
            return recs.reduce((s, r) => s + r.amount, 0);
        });
        const expenseData = timePoints.map(tp => {
            const recs = records.filter(r => r.type === 'expense' && matchTimePoint(r.date, tp, period));
            return recs.reduce((s, r) => s + calculateActualExpense(r), 0);
        });

        charts.summaryTrend = new Chart(trendCtx, {
            type: 'line',
            data: {
                labels: timePoints,
                datasets: [
                    { label: '收入', data: incomeData, borderColor: '#34c759', backgroundColor: 'rgba(52,199,89,0.06)', fill: true, tension: 0.4, pointRadius: 3, borderWidth: 2 },
                    { label: '实际支出', data: expenseData, borderColor: '#ff3b30', backgroundColor: 'rgba(255,59,48,0.06)', fill: true, tension: 0.4, pointRadius: 3, borderWidth: 2 },
                ]
            },
            options: {
                responsive: true,
                plugins: { legend: { labels: { boxWidth: 12, font: { size: 11 } } } },
                scales: {
                    x: { grid: { display: false }, ticks: { font: { size: 11 }, color: 'rgba(0,0,0,0.36)' } },
                    y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.04)' }, ticks: { font: { size: 11 }, color: 'rgba(0,0,0,0.36)' } },
                },
            }
        });
    }

    function getTimePoints(period, start, end) {
        const points = [];
        switch (period) {
            case 'day': {
                // Hours
                for (let h = 0; h < 24; h += 3) points.push(`${h}:00`);
                break;
            }
            case 'week': {
                const d = new Date(start);
                while (formatDate(d.toISOString()) <= end) {
                    points.push(formatDate(d.toISOString()));
                    d.setDate(d.getDate() + 1);
                }
                break;
            }
            case 'month': {
                const d = new Date(start);
                while (formatDate(d.toISOString()) <= end) {
                    points.push(formatDate(d.toISOString()));
                    d.setDate(d.getDate() + 1);
                }
                break;
            }
            case 'year': {
                for (let m = 1; m <= 12; m++) {
                    points.push(`${m}月`);
                }
                break;
            }
        }
        return points;
    }

    function matchTimePoint(dateStr, tp, period) {
        switch (period) {
            case 'day': return true; // all records match since already filtered by day
            case 'week':
            case 'month':
                return dateStr === tp;
            case 'year': {
                const month = parseInt(dateStr.split('-')[1]);
                return tp === `${month}月`;
            }
        }
    }

    function renderSummaryCategoryDetail(records) {
        const expenseRecords = records.filter(r => r.type === 'expense');
        const catTotals = {};
        for (const r of expenseRecords) {
            const actual = calculateActualExpense(r);
            catTotals[r.category] = (catTotals[r.category] || 0) + actual;
        }

        const catEntries = Object.entries(catTotals).sort((a, b) => b[1] - a[1]);
        const totalExpense = catEntries.reduce((s, [, v]) => s + v, 0) || 1;
        const maxVal = catEntries[0] ? catEntries[0][1] : 1;

        const listEl = document.getElementById('summary-category-list');
        if (catEntries.length === 0) {
            listEl.innerHTML = '<div class="empty-state"><p>暂无支出数据</p></div>';
            return;
        }

        listEl.innerHTML = catEntries.map(([catId, amount]) => {
            const cat = getCategoryInfo(catId, 'expense');
            const percent = Math.round(amount / totalExpense * 100);
            const barWidth = Math.round(amount / maxVal * 100);
            return `
                <div class="category-detail-item">
                    <div class="cat-icon">${cat.icon}</div>
                    <div class="cat-info">
                        <div class="cat-name">${cat.name}</div>
                        <div class="cat-bar"><div class="cat-bar-fill" style="width:${barWidth}%"></div></div>
                    </div>
                    <div class="cat-amount">${formatMoney(amount)}</div>
                    <div class="cat-percent">${percent}%</div>
                </div>
            `;
        }).join('');
    }

    // ========== Categories Management ==========
    function renderCategories() {
        const expGrid = document.getElementById('expense-categories');
        const incGrid = document.getElementById('income-categories');

        expGrid.innerHTML = state.categories.expense.map(c => `
            <div class="category-manage-item">
                <span class="cat-icon">${c.icon}</span>
                <span>${c.name}</span>
                <button class="delete-cat" onclick="app.deleteCategory('expense','${c.id}')">×</button>
            </div>
        `).join('');

        incGrid.innerHTML = state.categories.income.map(c => `
            <div class="category-manage-item">
                <span class="cat-icon">${c.icon}</span>
                <span>${c.name}</span>
                <button class="delete-cat" onclick="app.deleteCategory('income','${c.id}')">×</button>
            </div>
        `).join('');

        renderKeywordRules();
    }

    function renderKeywordRules() {
        const listEl = document.getElementById('keyword-rules');
        listEl.innerHTML = state.keywordRules.map((rule, idx) => {
            const cat = getCategoryInfo(rule.category, rule.type);
            return `
                <div class="keyword-rule">
                    <div class="rule-keywords">
                        ${rule.keywords.map(k => `<span class="keyword-tag">${k}</span>`).join('')}
                    </div>
                    <span class="rule-arrow">→</span>
                    <span class="rule-category">${cat.icon} ${cat.name}</span>
                    <button class="delete-rule" onclick="app.deleteKeywordRule(${idx})">×</button>
                </div>
            `;
        }).join('');
    }

    function showAddCategory() {
        showModal('添加分类', `
            <div class="form-group">
                <label>类型</label>
                <select id="new-cat-type">
                    <option value="expense">支出</option>
                    <option value="income">收入</option>
                </select>
            </div>
            <div class="form-group">
                <label>图标 (emoji)</label>
                <input type="text" id="new-cat-icon" placeholder="🏷️" maxlength="4">
            </div>
            <div class="form-group">
                <label>名称</label>
                <input type="text" id="new-cat-name" placeholder="分类名称">
            </div>
            <div class="modal-actions">
                <button class="btn-secondary" onclick="app.closeModal()">取消</button>
                <button class="btn-primary" onclick="app.addCategory()">添加</button>
            </div>
        `);
    }

    function addCategory() {
        const type = document.getElementById('new-cat-type').value;
        const icon = document.getElementById('new-cat-icon').value || '🏷️';
        const name = document.getElementById('new-cat-name').value.trim();
        if (!name) return toast('请输入分类名称', 'error');

        const id = name.toLowerCase().replace(/\s/g, '_') + '_' + Date.now().toString(36);
        state.categories[type].push({ id, name, icon });
        saveState();
        closeModal();
        renderCategories();
        toast('分类已添加');
    }

    function deleteCategory(type, id) {
        if (!confirm('确定要删除这个分类吗？')) return;
        state.categories[type] = state.categories[type].filter(c => c.id !== id);
        saveState();
        renderCategories();
        toast('分类已删除');
    }

    function showAddKeyword() {
        const allCats = [...state.categories.expense.map(c => ({ ...c, type: 'expense' })),
            ...state.categories.income.map(c => ({ ...c, type: 'income' }))];
        const options = allCats.map(c => `<option value="${c.id}" data-type="${c.type}">${c.icon} ${c.name} (${c.type === 'income' ? '收入' : '支出'})</option>`).join('');

        showModal('添加关键词规则', `
            <div class="form-group">
                <label>关键词（逗号分隔多个）</label>
                <input type="text" id="new-kw-keywords" placeholder="美团外卖, 饿了么, 外卖">
            </div>
            <div class="form-group">
                <label>对应分类</label>
                <select id="new-kw-category">${options}</select>
            </div>
            <div class="modal-actions">
                <button class="btn-secondary" onclick="app.closeModal()">取消</button>
                <button class="btn-primary" onclick="app.addKeywordRule()">添加</button>
            </div>
        `);
    }

    // Merge keyword rules with the same category+type into one rule
    function mergeKeywordRules() {
        const merged = {};
        for (const rule of state.keywordRules) {
            const key = `${rule.type}:${rule.category}`;
            if (merged[key]) {
                // Merge keywords, skip duplicates
                const existing = new Set(merged[key].keywords);
                for (const kw of rule.keywords) {
                    if (!existing.has(kw)) merged[key].keywords.push(kw);
                }
            } else {
                merged[key] = { ...rule, keywords: [...rule.keywords] };
            }
        }
        state.keywordRules = Object.values(merged);
    }

    function addKeywordRule() {
        const kwInput = document.getElementById('new-kw-keywords').value;
        const catSelect = document.getElementById('new-kw-category');
        const catId = catSelect.value;
        const catType = catSelect.options[catSelect.selectedIndex].dataset.type;

        const keywords = kwInput.split(/[,，]/).map(k => k.trim()).filter(k => k);
        if (keywords.length === 0) return toast('请输入关键词', 'error');

        // Merge into existing rule for same category if exists
        const existingRule = state.keywordRules.find(r => r.category === catId && r.type === catType);
        if (existingRule) {
            const existingKws = new Set(existingRule.keywords);
            for (const kw of keywords) {
                if (!existingKws.has(kw)) existingRule.keywords.push(kw);
            }
        } else {
            state.keywordRules.push({ keywords, category: catId, type: catType });
        }

        saveState();
        closeModal();
        renderCategories();
        toast('规则已添加');
    }

    function deleteKeywordRule(idx) {
        state.keywordRules.splice(idx, 1);
        saveState();
        renderCategories();
    }

    // ========== Modal ==========
    let _previousFocusEl = null;

    function showModal(title, bodyHtml) {
        _previousFocusEl = document.activeElement;
        document.getElementById('modal-title').textContent = title;
        document.getElementById('modal-body').innerHTML = bodyHtml;
        const overlay = document.getElementById('modal-overlay');
        overlay.style.display = 'flex';
        // Focus first focusable element inside modal
        requestAnimationFrame(() => {
            const focusable = overlay.querySelector('input, select, button, textarea, [tabindex]');
            if (focusable) focusable.focus();
        });
    }

    function closeModal() {
        document.getElementById('modal-overlay').style.display = 'none';
        // Restore focus to the element that opened the modal
        if (_previousFocusEl && typeof _previousFocusEl.focus === 'function') {
            _previousFocusEl.focus();
        }
        _previousFocusEl = null;
    }

    // ========== Export / Import Data ==========
    function exportData() {
        const data = {
            records: state.records,
            categories: state.categories,
            keywordRules: state.keywordRules,
            exportDate: new Date().toISOString(),
        };
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `smart-ledger-backup-${formatDate(new Date().toISOString())}.json`;
        a.click();
        URL.revokeObjectURL(url);
        toast('数据已导出');
    }

    function importData() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.json';
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (ev) => {
                try {
                    const data = JSON.parse(ev.target.result);
                    if (data.records) state.records = data.records;
                    if (data.categories) state.categories = data.categories;
                    if (data.keywordRules) state.keywordRules = data.keywordRules;
                    saveState();
                    toast(`已恢复 ${state.records.length} 条记录`);
                    navigateTo(state.currentPage);
                } catch (err) {
                    toast('数据文件格式错误', 'error');
                }
            };
            reader.readAsText(file);
        };
        input.click();
    }

    // ========== Init ==========
    function init() {
        loadState();
        mergeKeywordRules();
        saveState();

        // Navigation
        document.querySelectorAll('.nav-item').forEach(el => {
            el.addEventListener('click', () => navigateTo(el.dataset.page));
        });

        // Form tabs
        document.querySelectorAll('.form-tab').forEach(el => {
            el.addEventListener('click', () => {
                state.recordFormType = el.dataset.type;
                state.selectedCategory = null;
                document.querySelectorAll('.form-tab').forEach(t => t.classList.remove('active'));
                el.classList.add('active');
                renderAddForm();
            });
        });

        // Form submit
        document.getElementById('add-form').addEventListener('submit', handleFormSubmit);

        // AA toggle
        document.getElementById('aa-enabled').addEventListener('change', (e) => {
            document.getElementById('aa-details').style.display = e.target.checked ? 'block' : 'none';
            updateAAPreview();
        });

        // AA mode radio
        document.querySelectorAll('input[name="aa-mode"]').forEach(el => {
            el.addEventListener('change', () => {
                document.getElementById('aa-custom-amount').style.display =
                    document.querySelector('input[name="aa-mode"]:checked').value === 'custom' ? '' : 'none';
                updateAAPreview();
            });
        });

        // AA inputs (amount, people, my-share, date all trigger preview update)
        ['form-amount', 'aa-people', 'aa-my-share', 'form-date'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('input', updateAAPreview);
        });

        // Filters — select elements use 'change', search input uses debounced 'input'
        ['filter-type', 'filter-category', 'filter-month'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('change', () => { state.recordsPage = 1; renderRecords(); });
        });
        const searchEl = document.getElementById('filter-search');
        if (searchEl) {
            const debouncedSearch = debounce(() => { state.recordsPage = 1; renderRecords(); }, 300);
            searchEl.addEventListener('input', debouncedSearch);
        }

        // Summary tabs
        document.querySelectorAll('.summary-tab').forEach(el => {
            el.addEventListener('click', () => {
                state.summaryPeriod = el.dataset.period;
                state.summaryDate = new Date();
                document.querySelectorAll('.summary-tab').forEach(t => t.classList.remove('active'));
                el.classList.add('active');
                renderSummary();
            });
        });

        // Summary navigation
        document.getElementById('summary-prev').addEventListener('click', () => navigateSummary(-1));
        document.getElementById('summary-next').addEventListener('click', () => navigateSummary(1));

        // Import
        initImport();

        // Modal close on overlay click
        document.getElementById('modal-overlay').addEventListener('click', (e) => {
            if (e.target === e.currentTarget) closeModal();
        });

        // ESC key closes modal
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && document.getElementById('modal-overlay').style.display !== 'none') {
                closeModal();
            }
        });

        // Render initial page
        renderDashboard();
    }

    // Start
    init();

    // Public API
    return {
        navigateTo,
        editRecord,
        saveEditRecord,
        deleteRecord,
        goToPage,
        exportData,
        importData,
        updateMapping,
        confirmImport,
        cancelImport,
        showAddCategory,
        addCategory,
        deleteCategory,
        showAddKeyword,
        addKeywordRule,
        deleteKeywordRule,
        closeModal,
        // Batch operations
        toggleBatchMode,
        toggleSelect,
        toggleSelectAll,
        showBatchCategory,
        selectBatchCat,
        confirmBatchCategory,
        batchDelete,
        // AA match
        toggleMatchSelect,
        deleteMatchedIncomes,
    };
})();
