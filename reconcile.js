/**
 * reconcile.js — 月底校對比對模組
 * 功能：解析 Excel 貼上資料 → 模糊比對 → 產出差異報告
 * 與 script.js 完全獨立，互不干擾
 */

// ==================== 資料解析 ====================

/**
 * 將 Excel 複製貼上的 Tab 分隔文字解析為物件陣列
 * 自動偵測表頭欄位對應，支援彈性欄位順序
 * @param {string} text - Tab 分隔的文字（含表頭）
 * @returns {Array<Object>} 結構化資料陣列
 */
function parseReconcileData(text) {
    if (!text || !text.trim()) return [];

    const lines = text.trim().split('\n').map(l => l.replace(/\r$/, ''));
    if (lines.length === 0) return [];

    const firstLineCells = lines[0].split('\t').map(h => h.trim());

    // 先嘗試關鍵字表頭匹配
    const headerFieldMap = buildFieldMap(firstLineCells);
    const hasHeaderMatch = Object.keys(headerFieldMap).length > 0;

    // 判斷第一行是否為資料列（日期格式或中文日期）
    const isFirstLineData = looksLikeDate(firstLineCells[0]);

    let startIndex = 1;
    let finalFieldMap = headerFieldMap;

    if (isFirstLineData || !hasHeaderMatch) {
        // 無表頭：用第一列資料的格式特徵自動偵測欄位
        startIndex = 0;
        finalFieldMap = autoDetectFieldMap(firstLineCells);
    }

    const results = [];
    for (let i = startIndex; i < lines.length; i++) {
        const cells = lines[i].split('\t');
        if (cells.length < 2) continue;

        const row = {
            date: getCellValue(cells, finalFieldMap, 'date'),
            name: getCellValue(cells, finalFieldMap, 'name'),
            location: getCellValue(cells, finalFieldMap, 'location'),
            workContent: getCellValue(cells, finalFieldMap, 'workContent'),
            amount: getCellValue(cells, finalFieldMap, 'amount'),
            startTime: getCellValue(cells, finalFieldMap, 'startTime'),
        };

        if (row.date) {
            results.push(row);
        }
    }
    return results;
}

/**
 * 判斷字串是否像日期
 */
function looksLikeDate(str) {
    if (!str) return false;
    // 數字分隔格式：2026/2/1, 2/7
    if (/^\d{1,4}[\/\-]\d{1,2}/.test(str)) return true;
    // 中文格式：2月7日, 2/7日
    if (/^\d{1,2}月\d{1,2}日?/.test(str)) return true;
    return false;
}

/**
 * 無表頭時，分析各欄格式自動建立欄位對照表
 * 策略：
 *   - date: 第一個符合日期格式的欄
 *   - startTime: 符合 4 位數時間格式（0400, 0600）的欄
 *   - amount: 最後一個數值在 100~99999 之間的欄（最終金額通常在最後方）
 *   - name: 第一個 2~4 字且非日期/時間/數字的欄
 *   - workContent: 最後一個符合工作關鍵字的欄
 *   - location: 剩餘中最短的非數字欄
 */
function autoDetectFieldMap(cells) {
    const map = {};
    const workKeywords = ['接體', '出殯', '入殮', '洗身', '禮生', '冷凍', '協助', '告別式'];

    // 1. 找日期欄
    for (let i = 0; i < cells.length; i++) {
        if (looksLikeDate(cells[i])) {
            map.date = i;
            break;
        }
    }

    // 2. 找時間欄（0400, 0600, 0800... 4 位純數字，首位為 0 或 1）
    for (let i = 0; i < cells.length; i++) {
        if (i === map.date) continue;
        if (/^[012]\d{3}$/.test(cells[i].replace(':', ''))) {
            map.startTime = i;
            break;
        }
    }

    // 3. 找金額欄（從右往左，找第一個 100~99999 之間純數字欄）
    for (let i = cells.length - 1; i >= 0; i--) {
        if (i === map.date || i === map.startTime) continue;
        const num = parseInt(cells[i].replace(/[^0-9]/g, ''), 10);
        if (!isNaN(num) && num >= 100 && num <= 99999) {
            map.amount = i;
            break;
        }
    }

    // 4. 找姓名欄（通常在日期欄後第一個 2~4 字的非數字欄）
    const dateIdx = map.date ?? 0;
    for (let i = dateIdx + 1; i < cells.length; i++) {
        if (i === map.startTime || i === map.amount) continue;
        const v = cells[i];
        if (v.length >= 2 && v.length <= 4 && !/^\d/.test(v)) {
            map.name = i;
            break;
        }
    }

    // 5. 找工作內容欄（符合關鍵字），從左往右最後一個匹配
    for (let i = 0; i < cells.length; i++) {
        if (i === map.date || i === map.startTime || i === map.amount || i === map.name) continue;
        if (workKeywords.some(kw => cells[i].includes(kw))) {
            map.workContent = i;
        }
    }

    // 6. 剩餘未分配的欄位中，找地點（選擇最接近工作內容左側且不是已分配欄位的欄）
    if (map.workContent !== undefined) {
        for (let i = map.workContent - 1; i > (map.name ?? -1); i--) {
            if (i === map.date || i === map.startTime || i === map.amount || i === map.name) continue;
            const v = cells[i];
            if (v && !/^\d+$/.test(v)) {
                map.location = i;
                break;
            }
        }
    }

    // 7. Fallback：若仍有未對應的關鍵欄位，套用保守預設值
    if (map.date === undefined) map.date = 0;
    if (map.name === undefined) map.name = 1;
    if (map.location === undefined) map.location = map.workContent !== undefined ? map.workContent - 1 : 2;
    if (map.workContent === undefined) map.workContent = 3;
    if (map.amount === undefined) map.amount = 4;
    if (map.startTime === undefined) map.startTime = 2;

    return map;
}

/**
 * 建立表頭欄位名稱到索引的對照表
 */
function buildFieldMap(headers) {
    const map = {};
    const fieldAliases = {
        date: ['日期'],
        name: ['姓名', '名字', '人員'],
        location: ['地點', '場地', '館別'],
        workContent: ['工作內容', '內容', '項目', '工作項目'],
        amount: ['金額', '金額2', '費用', '報酬'],
        startTime: ['開始時間', '時間', '開始'],
    };

    for (const [field, aliases] of Object.entries(fieldAliases)) {
        for (const alias of aliases) {
            const idx = headers.findIndex(h => h.includes(alias));
            if (idx !== -1) {
                map[field] = idx;
                break;
            }
        }
    }
    return map;
}

/**
 * 從 cells 陣列中依欄位名取值
 * @param {string[]} cells
 * @param {Object} fieldMap
 * @param {string} fieldName
 * @returns {string}
 */
function getCellValue(cells, fieldMap, fieldName) {
    const idx = fieldMap[fieldName];
    if (idx === undefined || idx >= cells.length) return '';
    return (cells[idx] || '').trim();
}

// ==================== 模糊比對 ====================

/**
 * 模糊比對兩個字串是否相似
 * 判斷邏輯：完全相等 → A 包含 B → B 包含 A
 * 例：'龍圓' 與 '龍圓殯儀館' → true
 * @param {string} a
 * @param {string} b
 * @returns {boolean}
 */
function fuzzyMatch(a, b) {
    if (!a || !b) return false;
    const na = a.trim();
    const nb = b.trim();
    if (na === nb) return true;
    if (na.includes(nb) || nb.includes(na)) return true;
    return false;
}

/**
 * 標準化日期字串為統一格式 (YYYY/MM/DD)
 * 支援：2026/2/1, 2026/02/01, 2/1 等格式
 * @param {string} dateStr
 * @returns {string}
 */
function normalizeDate(dateStr) {
    if (!dateStr) return '';
    let s = dateStr.trim();

    // 中文格式：2月7日 → 轉換為 MM/DD
    const chineseMatch = s.match(/^(\d{1,4})月(\d{1,2})日?/);
    if (chineseMatch) {
        const part1 = chineseMatch[1];
        const part2 = chineseMatch[2];
        // 若第一組 > 12，視為年份（罕見但支援）
        if (parseInt(part1, 10) > 31) {
            // 有年份：例如「2026年2月7日」—先不處理這情況，直接跳到下面
        } else {
            const year = new Date().getFullYear();
            s = `${year}/${part1.padStart(2, '0')}/${part2.padStart(2, '0')}`;
            return s;
        }
    }

    const cleaned = s.replace(/-/g, '/');
    const parts = cleaned.split('/');

    if (parts.length === 3) {
        return `${parts[0]}/${parts[1].padStart(2, '0')}/${parts[2].padStart(2, '0')}`;
    }
    if (parts.length === 2) {
        const year = new Date().getFullYear();
        return `${year}/${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}`;
    }
    return cleaned;
}

/**
 * 標準化金額為純數字字串
 * @param {string} amountStr
 * @returns {string}
 */
function normalizeAmount(amountStr) {
    if (!amountStr) return '0';
    return String(parseInt(amountStr.replace(/[^0-9]/g, ''), 10) || 0);
}

/**
 * 標準化時間格式 (去除冒號、補零)
 * @param {string} timeStr
 * @returns {string}
 */
function normalizeTime(timeStr) {
    if (!timeStr) return '';
    const cleaned = timeStr.replace(/[:\s]/g, '');
    if (cleaned.length === 3) return '0' + cleaned;
    if (cleaned.length === 4) return cleaned;
    return cleaned;
}

// ==================== 核心比對邏輯 ====================

/**
 * 比對基準資料與回報資料
 * 比對依據：日期 + 地點（模糊）+ 工作內容
 * @param {Array<Object>} baseData  - 左側基準資料
 * @param {Array<Object>} reportData - 右側回報資料
 * @returns {Array<Object>} 比對結果陣列
 */
function reconcile(baseData, reportData) {
    const results = [];

    // 標準化所有資料
    const base = baseData.map(row => ({
        ...row,
        _date: normalizeDate(row.date),
        _amount: normalizeAmount(row.amount),
        _time: normalizeTime(row.startTime),
        _matched: false,
    }));

    const report = reportData.map(row => ({
        ...row,
        _date: normalizeDate(row.date),
        _amount: normalizeAmount(row.amount),
        _time: normalizeTime(row.startTime),
        _matched: false,
    }));

    // 第一輪：嘗試為每筆 report 找到 base 中的匹配項
    for (const rRow of report) {
        let bestMatch = null;
        let bestScore = 0;

        for (const bRow of base) {
            if (bRow._matched) continue;

            // 日期必須完全相同
            if (rRow._date !== bRow._date) continue;

            // 計算匹配分數
            let score = 0;

            // 地點模糊比對
            if (fuzzyMatch(rRow.location, bRow.location)) score += 2;

            // 工作內容模糊比對
            if (fuzzyMatch(rRow.workContent, bRow.workContent)) score += 2;

            // 姓名比對（額外加分）
            if (rRow.name && bRow.name && fuzzyMatch(rRow.name, bRow.name)) score += 1;

            // 至少地點或工作內容要有一個匹配
            if (score >= 2 && score > bestScore) {
                bestScore = score;
                bestMatch = bRow;
            }
        }

        if (bestMatch) {
            bestMatch._matched = true;
            rRow._matched = true;

            // 檢查金額與時間是否一致
            const amountDiff = rRow._amount !== bestMatch._amount;
            const timeDiff = rRow._time && bestMatch._time && rRow._time !== bestMatch._time;

            if (amountDiff || timeDiff) {
                // [差異]：匹配成功但金額或時間不同
                const diffs = [];
                if (amountDiff) diffs.push(`金額：基準 ${bestMatch._amount} / 回報 ${rRow._amount}`);
                if (timeDiff) diffs.push(`時間：基準 ${bestMatch._time} / 回報 ${rRow._time}`);

                results.push({
                    status: 'diff',
                    date: rRow._date,
                    name: rRow.name || bestMatch.name,
                    location: bestMatch.location,
                    workContent: bestMatch.workContent,
                    baseAmount: bestMatch._amount,
                    reportAmount: rRow._amount,
                    baseTime: bestMatch._time,
                    reportTime: rRow._time,
                    description: diffs.join('；'),
                    amountDiff,
                    timeDiff,
                });
            } else {
                // [正確]：完全一致
                results.push({
                    status: 'match',
                    date: rRow._date,
                    name: rRow.name || bestMatch.name,
                    location: bestMatch.location,
                    workContent: bestMatch.workContent,
                    baseAmount: bestMatch._amount,
                    reportAmount: rRow._amount,
                    baseTime: bestMatch._time,
                    reportTime: rRow._time,
                    description: '一致',
                    amountDiff: false,
                    timeDiff: false,
                });
            }
        } else {
            // [多報]：右側（回報）有，左側（基準）找不到
            rRow._matched = true;
            results.push({
                status: 'over_reported',
                date: rRow._date,
                name: rRow.name,
                location: rRow.location,
                workContent: rRow.workContent,
                baseAmount: '-',
                reportAmount: rRow._amount,
                baseTime: '-',
                reportTime: rRow._time,
                description: '基準無此筆，月底回報疑似多報',
                amountDiff: false,
                timeDiff: false,
            });
        }
    }

    // 第二輪：找出 base 中未被匹配的項目 → [漏單]
    for (const bRow of base) {
        if (!bRow._matched) {
            results.push({
                status: 'under_reported',
                date: bRow._date,
                name: bRow.name,
                location: bRow.location,
                workContent: bRow.workContent,
                baseAmount: bRow._amount,
                reportAmount: '-',
                baseTime: bRow._time,
                reportTime: '-',
                description: '月底回報未見此筆，疑似漏單',
                amountDiff: false,
                timeDiff: false,
            });
        }
    }

    // 依日期排序，異常項目優先（漏單 > 差異 > 多報 > 正確）
    const statusOrder = { under_reported: 0, diff: 1, over_reported: 2, match: 3 };
    results.sort((a, b) => {
        const sDiff = statusOrder[a.status] - statusOrder[b.status];
        if (sDiff !== 0) return sDiff;
        return a.date.localeCompare(b.date);
    });

    return results;
}

// ==================== UI 互動邏輯 ====================

let reconcileResults = [];        // 完整比對結果
let reconcileFilteredResults = []; // 過濾後的結果
let currentReconcileFilter = 'problem'; // 預設過濾模式

/**
 * 切換 Tab 頁面
 * @param {string} tabName - 'convert' 或 'reconcile'
 */
function switchTab(tabName) {
    const tabConvert = document.getElementById('tab-convert');
    const tabReconcile = document.getElementById('tab-reconcile');
    const btnConvert = document.getElementById('tabBtnConvert');
    const btnReconcile = document.getElementById('tabBtnReconcile');

    if (tabName === 'convert') {
        tabConvert.classList.remove('hidden');
        tabReconcile.classList.add('hidden');
        btnConvert.classList.add('tab-btn-active');
        btnReconcile.classList.remove('tab-btn-active');
    } else {
        tabConvert.classList.add('hidden');
        tabReconcile.classList.remove('hidden');
        btnConvert.classList.remove('tab-btn-active');
        btnReconcile.classList.add('tab-btn-active');
    }
}

/**
 * 切換校對結果過濾標籤
 * @param {HTMLElement} btn - 被點擊的按鈕元素
 */
function toggleReconcileFilter(btn) {
    // 移除所有 filter-chip 的 active 狀態
    document.querySelectorAll('.filter-chip').forEach(b => b.classList.remove('filter-chip-active'));
    btn.classList.add('filter-chip-active');

    currentReconcileFilter = btn.dataset.filter;
    applyReconcileFilter();
}

/**
 * 依據目前的過濾模式篩選結果
 */
function applyReconcileFilter() {
    if (currentReconcileFilter === 'all') {
        reconcileFilteredResults = [...reconcileResults];
    } else if (currentReconcileFilter === 'problem') {
        // 「異常項目」= 多報 + 差異 + 少報
        reconcileFilteredResults = reconcileResults.filter(r => r.status !== 'match');
    } else {
        reconcileFilteredResults = reconcileResults.filter(r => r.status === currentReconcileFilter);
    }
    renderReconcileTable();
}

/**
 * 渲染校對結果表格
 */
function renderReconcileTable() {
    const tbody = document.getElementById('reconcileResultBody');

    if (reconcileFilteredResults.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-400 text-sm">
                    目前過濾條件下沒有符合的項目
                </td>
            </tr>`;
        return;
    }

    tbody.innerHTML = reconcileFilteredResults.map(r => {
        // 狀態標籤
        const statusBadge = getStatusBadge(r.status);

        // 金額差異高亮
        const baseAmountClass = r.amountDiff ? 'diff-highlight-base' : '';
        const reportAmountClass = r.amountDiff ? 'diff-highlight-report' : '';

        // 時間差異高亮
        const baseTimeClass = r.timeDiff ? 'diff-highlight-base' : '';
        const reportTimeClass = r.timeDiff ? 'diff-highlight-report' : '';

        // 行底色
        const rowClass = getRowClass(r.status);

        return `
            <tr class="${rowClass} hover:bg-indigo-50/30 transition-colors">
                <td class="p-3 text-xs">${statusBadge}</td>
                <td class="p-3 text-xs text-slate-600">${r.date}</td>
                <td class="p-3 text-xs text-slate-800 font-medium">${r.name}</td>
                <td class="p-3 text-xs text-slate-600">${r.location}</td>
                <td class="p-3 text-xs text-slate-600">${r.workContent}</td>
                <td class="p-3 text-xs text-slate-700 font-medium ${baseAmountClass}">${r.baseAmount}</td>
                <td class="p-3 text-xs text-slate-700 font-medium ${reportAmountClass}">${r.reportAmount}</td>
                <td class="p-3 text-xs text-slate-600 font-mono ${baseTimeClass}">${r.baseTime}</td>
                <td class="p-3 text-xs text-slate-600 font-mono ${reportTimeClass}">${r.reportTime}</td>
                <td class="p-3 text-xs text-slate-500 italic">${r.description}</td>
            </tr>`;
    }).join('');
}

/**
 * 取得狀態標籤 HTML
 * @param {string} status
 * @returns {string}
 */
function getStatusBadge(status) {
    const badges = {
        under_reported: '<span class="status-badge status-badge-extra">漏單</span>',
        diff: '<span class="status-badge status-badge-diff">差異</span>',
        over_reported: '<span class="status-badge status-badge-missing">多報</span>',
        match: '<span class="status-badge status-badge-match">正確</span>',
    };
    return badges[status] || '';
}

/**
 * 狀態轉純文字（供複製到 Excel 用）
 * @param {string} status
 * @returns {string}
 */
function getStatusText(status) {
    const map = {
        under_reported: '漏單',
        diff: '差異',
        over_reported: '多報',
        match: '正確',
    };
    return map[status] || '';
}

/**
 * 複製差異清單（Tab 分隔），可直接貼到 Excel
 */
async function handleReconcileCopy() {
    const headers = ['狀態', '日期', '姓名', '地點', '工作內容', '基準金額', '回報金額', '基準時間', '回報時間', '說明'];
    const rows = reconcileResults.map(r => [
        getStatusText(r.status),
        r.date,
        r.name,
        r.location,
        r.workContent,
        r.baseAmount,
        r.reportAmount,
        r.baseTime,
        r.reportTime,
        r.description,
    ].join('\t'));

    const content = [headers.join('\t'), ...rows].join('\n');

    try {
        await navigator.clipboard.writeText(content);
        const btn = document.getElementById('reconcileCopyBtn');
        const successMsg = document.getElementById('reconcileCopySuccess');
        const oldHTML = btn.innerHTML;
        const oldClass = btn.className;

        btn.innerHTML = '<svg class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg> 已複製！';
        btn.className = 'bg-emerald-500 hover:bg-emerald-600 text-white py-2 px-5 rounded-lg text-sm flex items-center gap-2 transition-all shadow-md font-medium';
        successMsg.classList.remove('hidden');

        setTimeout(() => {
            btn.innerHTML = oldHTML;
            btn.className = oldClass;
            successMsg.classList.add('hidden');
        }, 3000);
    } catch (e) {
        alert('複製失敗，請手動框選表格後複製。');
    }
}

/**
 * 取得行底色 class
 * @param {string} status
 * @returns {string}
 */
function getRowClass(status) {
    const classes = {
        over_reported: 'bg-red-50/40',
        diff: 'bg-amber-50/40',
        under_reported: 'bg-blue-50/40',
        match: '',
    };
    return classes[status] || '';
}

/**
 * 更新統計摘要數字
 */
function updateReconcileStats() {
    const counts = { over_reported: 0, diff: 0, under_reported: 0, match: 0 };
    for (const r of reconcileResults) {
        counts[r.status] = (counts[r.status] || 0) + 1;
    }

    document.getElementById('statMissing').textContent = counts.over_reported;
    document.getElementById('statDiff').textContent = counts.diff;
    document.getElementById('statExtra').textContent = counts.under_reported;
    document.getElementById('statMatch').textContent = counts.match;
}

// ==================== 事件綁定 ====================

window.addEventListener('DOMContentLoaded', () => {
    const baseTextarea = document.getElementById('reconcileBaseText');
    const reportTextarea = document.getElementById('reconcileReportText');
    const reconcileBtn = document.getElementById('reconcileBtn');
    const clearBtn = document.getElementById('reconcileClearBtn');

    // 啟用/停用「開始校對」按鈕
    function checkReconcileReady() {
        reconcileBtn.disabled = !(baseTextarea.value.trim() && reportTextarea.value.trim());
    }

    baseTextarea.addEventListener('input', checkReconcileReady);
    reportTextarea.addEventListener('input', checkReconcileReady);

    // 開始校對
    reconcileBtn.addEventListener('click', () => {
        const baseData = parseReconcileData(baseTextarea.value);
        const reportData = parseReconcileData(reportTextarea.value);

        if (baseData.length === 0) {
            alert('左側「基準資料」解析結果為空。\n請確認為 Tab 分隔格式（從 Excel 直接複製），至少需有日期欄位。');
            return;
        }
        if (reportData.length === 0) {
            alert('右側「月底回報」解析結果為空。\n請確認為 Tab 分隔格式（從 Excel 直接複製），至少需有日期欄位。');
            return;
        }

        reconcileResults = reconcile(baseData, reportData);
        currentReconcileFilter = 'problem';

        // 重置過濾標籤為「異常項目」
        document.querySelectorAll('.filter-chip').forEach(b => b.classList.remove('filter-chip-active'));
        document.querySelector('[data-filter="problem"]').classList.add('filter-chip-active');

        updateReconcileStats();
        applyReconcileFilter();

        // 顯示結果區域
        document.getElementById('reconcileResultSection').classList.remove('hidden');

        // 捲動到結果區
        setTimeout(() => {
            document.getElementById('reconcileResultSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
    });

    // 複製差異清單按鈕
    document.getElementById('reconcileCopyBtn').addEventListener('click', handleReconcileCopy);

    // 清除按鈕
    clearBtn.addEventListener('click', () => {
        baseTextarea.value = '';
        reportTextarea.value = '';
        reconcileResults = [];
        reconcileFilteredResults = [];
        reconcileBtn.disabled = true;
        document.getElementById('reconcileResultSection').classList.add('hidden');
    });
});
