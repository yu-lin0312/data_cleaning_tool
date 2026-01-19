// Data Configuration
const projectVendors = ['松佑', '尚閔', '龍霖', '銘棋', '高盛', '龍圓', '新光人生', '采邑', '鴻德', '祐安', '和榕', '協和', '上德', '高誠', '公庭', '福祥', '出云', '恆善蔡嘉薇', '群豐東昇', '南華靜觀', '南都', '榮祥', '南華'];
const locationTypoFixes = { '人本': '仁本', '橋殯': '橋頭殯儀館', '會館': '龍巖會館' };
const vendorTypoFixes = { '新光': '新光人生' };
const nameMapping = {
    '勳': '王健勳', '健勳': '王健勳', '千': '林千智', '千智': '林千智', '黑': '郭育帆', '小黑': '郭育帆',
    '潘': '潘信安', '小潘': '潘信安', '魚': '謝瑋育', '鮪魚': '謝瑋育', '郭': '郭鴻億', '小郭': '郭鴻億',
    '邱': '邱暐傑', '暐傑': '邱暐傑', '聰': '戴耀聰', '燿聰': '戴耀聰', '耀聰': '戴耀聰', '鈞': '鄒至鈞',
    '至鈞': '鄒至鈞', '科': '何科賢', '賢': '何科賢', '風': '王德風', '玲': '蔡采玲', '小玲': '蔡采玲',
    '九': '黃湘玲', '小九': '黃湘玲', '瑩': '戴嘉瑩', '嘉瑩': '戴嘉瑩', '珼': '袁翊珼', '貝貝': '袁翊珼',
    '軒': '謝立軒', '立軒': '謝立軒', '文': '羅文昕', '阿文': '羅文昕', '承': '葉承恩', '恩': '葉承恩',
    '承恩': '葉承恩', '羊羊': '楊洋', '洋洋': '楊洋', '力宏': '林立宏', '立宏': '林立宏', '皮皮': '尤怡蘋'
};
const preservedNames = ['潘調', '勳調'];

const MAIN_FEES = {
    '入殮': 1000, '出殯': 1200, '入殮出殯': 1700, '入殮扛夫': 1700, '入殮火化': 1000,
    '禮生': 1000, '禮生出殯': 1400, '禮生扶棺': 1500, '午夜功德': 2000, '半日功德': 1000,
    '招待': 1200, '接體': 1200, '晉塔': 1000
};

const OTHER_FEES = {
    '頭七~滿七': 500, '女兒旬': 500, '接體空跑': 500, '退冰': 300, '驗屍/復驗': 500,
    '佈置': 500, '安主/安位': 1000, '返主': 1000, '顧SPA': 500, '教會出殯': 1200
};

// Helper Functions
function isProjectVendor(vendor) { return vendor && projectVendors.some(pv => vendor.includes(pv)); }
function fixLocationTypo(location) { if (!location) return location; let fixed = location; for (const [typo, correct] of Object.entries(locationTypoFixes)) fixed = fixed.replace(typo, correct); return fixed; }
function fixVendorTypo(vendor) { if (!vendor) return vendor; let fixed = vendor; for (const [typo, correct] of Object.entries(vendorTypoFixes)) fixed = fixed.replace(typo, correct); return fixed; }
function convertName(name) { const trimmed = name.trim(); return preservedNames.includes(trimmed) ? trimmed : (nameMapping[trimmed] || trimmed); }
function formatTime(timeStr) { let cleaned = timeStr.replace(/[:\s]/g, ''); if (cleaned.length === 1) return '0' + cleaned + '00'; if (cleaned.length === 2) return cleaned + '00'; if (cleaned.length === 3) return '0' + cleaned; if (cleaned.length === 4) return cleaned; return timeStr; }

function normalizeWorkContent(content) {
    if (!content) return content;
    let normalized = content.trim();
    normalized = normalized.replace(/入出/g, '入殮出殯');
    normalized = normalized.replace(/禮出/g, '禮生出殯');
    normalized = normalized.replace(/禮扶/g, '禮生扶棺');
    normalized = normalized.replace(/禮扛/g, '禮生扛棺');
    normalized = normalized.replace(/入冰/g, '入殮退冰');
    normalized = normalized.replace(/午夜(?!功德)/g, '午夜功德');
    return normalized;
}

function isDateLine(line) { return /^\d{1,2}\/\d{1,2}\s*$/.test(line.trim()); }
function isSeparatorLine(line) { return /^[-—─]+$/.test(line.trim()) || line.trim() === ''; }
function isScheduleLine(line) { return /^\d{1,2}[:\s]?\d{0,2}\s+/.test(line.trim()); }
function isCaseNameLine(line) { return line.trim().startsWith('案名：') || line.trim().startsWith('案名:'); }
function isRitualistLine(line) { return line.trim().startsWith('禮儀師：') || line.trim().startsWith('禮儀師:'); }

function isNamesLine(line) {
    const trimmed = line.trim();
    // Fast Path: Check for Chinese characters first
    if (!/^[\u4e00-\u9fa5\s]+$/.test(trimmed)) return false;

    // Exclude specific non-name lines that might contain Chinese
    if (isCaseNameLine(line) || isRitualistLine(line) || isScheduleLine(line)) return false;

    return true;
}

function parseScheduleLine(line) {
    const trimmed = line.trim();
    const timeMatch = trimmed.match(/^(\d{1,2}[:\s]?\d{0,2})\s+(.+)$/);
    if (!timeMatch) return null;
    const time = formatTime(timeMatch[1]);
    const parts = timeMatch[2].split(/\s+/);
    if (parts.length >= 3) return { time, location: fixLocationTypo(parts[0]), vendor: fixVendorTypo(parts[1]), workContent: normalizeWorkContent(parts.slice(2).join('')) };
    if (parts.length === 2) return { time, location: fixLocationTypo(parts[0]), vendor: fixVendorTypo(parts[1]), workContent: '' };
    return { time, location: fixLocationTypo(timeMatch[2]), vendor: '', workContent: '' };
}

const calculationCache = new Map();

function calculateAmount(workContent, vendor = '', location = '') {
    const cacheKey = `${workContent}|${vendor}|${location}`;
    if (calculationCache.has(cacheKey)) {
        return calculationCache.get(cacheKey);
    }

    const result = (function () {
        if (!workContent) return { amount: 0, needsManualCheck: true };
        const content = workContent.trim();

        // Special Rules
        if (/台南山上鄉/.test(content)) { if (/入殮出殯/.test(content)) return { amount: 1500, needsManualCheck: false }; if (/禮生/.test(content)) return { amount: 1500, needsManualCheck: false }; }
        if (/林園/.test(content) && /禮生/.test(content)) return { amount: 1200, needsManualCheck: false };
        if (/臭臭/.test(content) && /入殮出殯/.test(content)) return { amount: 2500, needsManualCheck: false };
        if (/接臭屍/.test(content)) return { amount: 2000, needsManualCheck: false };
        if ((/柳營/.test(location) || /新營/.test(location)) && /禮生出殯/.test(content)) return { amount: 2000, needsManualCheck: false };
        if ((/台南聖恩/.test(vendor) || /台南龍巖/.test(vendor)) && /禮生出殯/.test(content)) return { amount: 1500, needsManualCheck: false };
        if (/高雄聖恩/.test(vendor) && /禮生出殯/.test(content)) return { amount: 1400, needsManualCheck: false };

        // Combinations
        if (/洗穿/.test(content) && /化妝/.test(content) && /入殮出殯/.test(content)) return { amount: 2400, needsManualCheck: false };
        if (/更衣入驗/.test(content) && /禮生出殯/.test(content)) return { amount: 2400, needsManualCheck: false };
        if (/洗穿/.test(content) && /入殮出殯/.test(content)) return { amount: 2200, needsManualCheck: false };
        if (/洗穿/.test(content) && /入殮扛夫/.test(content)) return { amount: 2200, needsManualCheck: false };
        if (/化妝/.test(content) && /入殮出殯/.test(content) && !/更衣/.test(content)) return { amount: 2200, needsManualCheck: false };
        if (/更衣/.test(content) && /入殮出殯/.test(content) && !/化妝/.test(content)) return { amount: 2200, needsManualCheck: false };
        if ((/加衣/.test(content) || /更衣/.test(content)) && /入殮/.test(content) && !/出殯/.test(content)) return { amount: 1500, needsManualCheck: false };
        if (/出殯/.test(content) && /回洗/.test(content)) return { amount: 1500, needsManualCheck: false };
        if (/入殮出殯/.test(content) && /\+禮生/.test(content)) return { amount: 1900, needsManualCheck: false };
        if (/入殮/.test(content) && /出殯/.test(content) && /禮生/.test(content)) return { amount: 1900, needsManualCheck: false };

        // Main Items
        if (/午夜功德|午夜/.test(content)) return { amount: 2000, needsManualCheck: false };
        if (/換罐樹葬/.test(content)) return { amount: 2000, needsManualCheck: false };
        if (/半日功德|半日燒庫|^半日$/.test(content)) return { amount: 1000, needsManualCheck: false };
        if (/佛教藥懺/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/頭七.*燒庫|頭七\+燒庫/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/頭七|二七|三七|五七|滿七|女兒旬|女兒七/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/接體空跑/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/接體/.test(content)) return { amount: 1200, needsManualCheck: false };
        if (/入殮退冰/.test(content)) return { amount: 1300, needsManualCheck: false };
        if (/退冰/.test(content)) return { amount: 300, needsManualCheck: false };
        if (/驗屍|復驗|相相驗/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/豎靈/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/引魂/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/佈置/.test(content)) return { amount: 500, needsManualCheck: false };
        if (/安主|安位/.test(content)) return { amount: 1000, needsManualCheck: false };
        if (/返主/.test(content)) return { amount: 1000, needsManualCheck: false };
        if (/晉塔|進塔/.test(content)) return { amount: 1000, needsManualCheck: false };
        if (/顧spa/i.test(content)) return { amount: 500, needsManualCheck: false };
        if (/招待/.test(content)) return { amount: 1200, needsManualCheck: false };
        if (/教會出殯/.test(content)) return { amount: 1200, needsManualCheck: false };

        // Standard Items
        if (/扶棺/.test(content) && !/禮生/.test(content)) return { amount: 700, needsManualCheck: false };
        if (/入殮扛夫/.test(content) || (/入殮/.test(content) && /扛夫/.test(content))) return { amount: 1700, needsManualCheck: false };
        if (/入殮火化|入殮送火/.test(content)) return { amount: 1000, needsManualCheck: false };
        if (/入殮出殯/.test(content) || (/入殮/.test(content) && /出殯/.test(content))) return { amount: 1700, needsManualCheck: false };
        if (/禮生扶棺|禮生扛棺/.test(content)) return { amount: 1500, needsManualCheck: false };
        if (/禮生出殯/.test(content)) return { amount: 1400, needsManualCheck: false };
        if (/禮生/.test(content) && !/出殯|扶棺|扛棺/.test(content)) return { amount: 1000, needsManualCheck: false };
        if (/入殮/.test(content) && !/出殯|扛夫|火化|送火/.test(content)) return { amount: 1000, needsManualCheck: false };
        if (/出殯/.test(content) && !/入殮|禮生|回洗/.test(content)) return { amount: 1200, needsManualCheck: false };

        return { amount: 0, needsManualCheck: true };
    })();

    calculationCache.set(cacheKey, result);
    return result;
}

function parseScheduleData(text) {
    const lines = text.split('\n'), results = [];
    let currentDate = '', lastSchedule = null, currentCaseName = '', currentRitualist = '';

    for (let line of lines) {
        if (isSeparatorLine(line)) continue;
        if (isDateLine(line)) { const m = line.trim().match(/(\d{1,2})\/(\d{1,2})/); if (m) currentDate = `2025/${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}`; lastSchedule = null; currentCaseName = ''; currentRitualist = ''; continue; }
        if (isScheduleLine(line)) {
            const p = parseScheduleLine(line);
            if (p) {
                lastSchedule = { date: currentDate, startTime: p.time, location: p.location, vendor: p.vendor.replace(/[哥姐]/g, ''), workContent: p.workContent };
                currentCaseName = '';
                currentRitualist = '';
            }
            continue;
        }
        if (isCaseNameLine(line)) { currentCaseName = line.trim().replace(/^案名[：:]/, '').trim(); continue; }
        if (isRitualistLine(line)) { currentRitualist = line.trim().replace(/^禮儀師[：:]/, '').trim(); continue; }
        if (isNamesLine(line) && lastSchedule) {
            const names = line.trim().split(/\s+/).filter(n => n);
            for (let name of names) {
                const fullName = convertName(name);
                const { amount, needsManualCheck } = calculateAmount(lastSchedule.workContent, lastSchedule.vendor, lastSchedule.location);
                let notes = currentRitualist || currentCaseName || '';
                if (needsManualCheck) notes = (notes ? notes + '；' : '') + '需人工確認金額';
                if (isProjectVendor(lastSchedule.vendor)) notes = (notes ? notes + '；' : '') + '專案';
                results.push({ date: lastSchedule.date, name: fullName, startTime: lastSchedule.startTime, location: lastSchedule.location, vendor: lastSchedule.vendor, workContent: lastSchedule.workContent, amount: amount.toString(), paymentStatus: '未收', notes: notes, amount2: amount.toString() });
            }
            lastSchedule = null;
        }
    }
    return results;
}

// UI Logic
let parsedData = [], filteredData = [];

window.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    renderRules();

    const input = document.getElementById('inputText');
    const parseBtn = document.getElementById('parseBtn');
    const clearBtn = document.getElementById('clearBtn');
    const nameFilter = document.getElementById('nameFilter');
    const copyBtn = document.getElementById('copyBtn');
    const downloadCsvBtn = document.getElementById('downloadCsvBtn');

    input.addEventListener('input', () => {
        parseBtn.disabled = !input.value.trim();
        updateInputCount();
    });

    parseBtn.addEventListener('click', () => {
        const startTime = performance.now();

        // Clear cache for a fresh run
        calculationCache.clear();

        parsedData = parseScheduleData(input.value);
        filterData();

        const endTime = performance.now();
        const duration = (endTime - startTime).toFixed(2);
        console.log(`Parsing completed in ${duration}ms`);

        updateUIState(duration);
    });

    clearBtn.addEventListener('click', () => {
        input.value = '';
        parsedData = [];
        filteredData = [];
        nameFilter.value = '';
        parseBtn.disabled = true;
        updateInputCount();
        updateUIState();
    });

    nameFilter.addEventListener('input', filterData);
    copyBtn.addEventListener('click', handleCopyToClipboard);
    downloadCsvBtn.addEventListener('click', handleDownloadCsv);
});

function renderRules() {
    // Name Mapping
    const nameList = document.getElementById('nameMappingList');
    nameList.innerHTML = Object.entries(nameMapping).map(([k, v]) => `<span>${k} → ${v}</span>`).join('');

    // Main Fees
    const mainList = document.getElementById('mainFeeList');
    mainList.innerHTML = Object.entries(MAIN_FEES).map(([k, v]) =>
        `<div class="flex justify-between items-center"><span>${k}</span><span class="font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">${v}</span></div>`
    ).join('');

    // Other Fees
    const otherList = document.getElementById('otherFeeList');
    otherList.innerHTML = Object.entries(OTHER_FEES).map(([k, v]) =>
        `<div class="flex justify-between items-center"><span>${k}</span><span class="font-semibold text-slate-800 bg-slate-100 px-2 py-0.5 rounded">${v}</span></div>`
    ).join('');

    // Project Vendors
    const vendorList = document.getElementById('projectVendorList');
    vendorList.innerHTML = projectVendors.map(v =>
        `<span class="px-2 py-1 bg-slate-100 rounded text-slate-600 border border-slate-200">${v}</span>`
    ).join('');
}

function updateInputCount() {
    const text = document.getElementById('inputText').value;
    const lines = text.split('\n');
    const count = lines.filter(line => isScheduleLine(line)).length;
    const badge = document.getElementById('inputCountBadge');

    if (count > 0) {
        badge.textContent = `${count} 筆`;
        badge.classList.remove('hidden');
    } else {
        badge.classList.add('hidden');
    }
}

function filterData() {
    const query = document.getElementById('nameFilter').value.trim();
    filteredData = query ? parsedData.filter(r => r.name.includes(query)) : [...parsedData];
    // Sort by date
    filteredData.sort((a, b) => new Date(a.date) - new Date(b.date));
    renderTable();
}

function updateUIState(duration = null) {
    const hasData = parsedData.length > 0;
    document.getElementById('emptyState').classList.toggle('hidden', hasData);
    document.getElementById('resultContainer').classList.toggle('hidden', !hasData);
    document.getElementById('actionButtons').classList.toggle('hidden', !hasData);

    const badge = document.getElementById('countBadge');
    badge.classList.toggle('hidden', !hasData);

    let badgeText = nameFilter.value ? `${filteredData.length} / ${parsedData.length} 筆` : `${parsedData.length} 筆`;
    if (duration) {
        badgeText += ` (耗時 ${duration}ms)`;
    }
    badge.textContent = badgeText;
}

function renderTable() {
    const tbody = document.getElementById('resultBody');
    tbody.innerHTML = filteredData.map(r => `
        <tr class="hover:bg-indigo-50/30 transition-colors group">
            <td class="p-4 text-slate-600 text-sm">${r.date}</td>
            <td class="p-4 text-slate-800 font-medium text-sm">${r.name}</td>
            <td class="p-4 text-slate-600 font-mono text-xs">${r.startTime}</td>
            <td class="p-4 text-slate-600 text-sm">${r.location}</td>
            <td class="p-4 text-slate-600 text-sm">${r.vendor}</td>
            <td class="p-4 text-slate-600 text-sm">${r.workContent}</td>
            <td class="p-4 text-slate-700 font-medium text-sm">${r.amount}</td>
            <td class="p-4 text-slate-500 text-sm">${r.paymentStatus}</td>
            <td class="p-4 text-slate-500 text-xs italic">${r.notes}</td>
            <td class="p-4 text-slate-700 font-medium text-sm">${r.amount2}</td>
        </tr>
    `).join('');
}

function handleDownloadCsv() {
    const headers = ['日期', '姓名', '開始時間', '地點', '廠商/單位', '工作內容', '金額', '收款狀態', '備註', '金額2'];
    const rows = filteredData.map(r => [
        r.date, r.name, r.startTime, r.location, r.vendor, r.workContent, r.amount, r.paymentStatus, r.notes, r.amount2
    ].map(c => `"${String(c || '').replace(/"/g, '""')}"`).join(','));

    const blob = new Blob(['\uFEFF' + [headers.join(','), ...rows].join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `排班資料_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

async function handleCopyToClipboard() {
    const headers = ['日期', '姓名', '開始時間', '地點', '廠商/單位', '工作內容', '金額', '收款狀態', '備註', '金額2'];
    const rows = filteredData.map(r => [
        r.date, r.name, r.startTime, r.location, r.vendor, r.workContent, r.amount, r.paymentStatus, r.notes, r.amount2
    ].join('\t'));

    try {
        await navigator.clipboard.writeText([headers.join('\t'), ...rows].join('\n'));
        const btn = document.getElementById('copyBtn');
        const oldHTML = btn.innerHTML;
        const oldClass = btn.className;

        btn.innerHTML = '<i data-lucide="check" class="w-4 h-4"></i> 已複製';
        btn.className = 'bg-emerald-500 hover:bg-emerald-600 text-white py-2 px-4 rounded-lg text-sm flex items-center gap-2 transition-all shadow-md';
        lucide.createIcons();

        setTimeout(() => {
            btn.innerHTML = oldHTML;
            btn.className = oldClass;
            lucide.createIcons();
        }, 2000);
    } catch (e) { alert('複製失敗'); }
}
