// Data Configuration
const projectVendors = ['松佑', '尚閔', '龍霖', '銘棋', '高盛', '龍圓', '新光人生', '采邑', '鴻德', '祐安', '和榕', '協和', '上德', '高誠', '公庭', '福祥', '出云', '恆善蔡嘉薇', '群豐東昇', '南華靜觀', '南都', '榮祥', '南華'];
const locationTypoFixes = { '人本': '仁本', '橋殯': '橋頭殯儀館', '會館': '龍巖會館' };
const vendorTypoFixes = { '新光': '新光人生', '仁本': '台灣仁本' };
const nameMapping = {
    '勳': '王健勳', '健勳': '王健勳', '千': '林千智', '千智': '林千智', '黑': '郭育帆', '小黑': '郭育帆',
    '潘': '潘信安', '小潘': '潘信安', '魚': '謝瑋育', '鮪魚': '謝瑋育', '郭': '郭鴻億', '小郭': '郭鴻億',
    '邱': '邱暐傑', '暐傑': '邱暐傑', '聰': '戴耀聰', '燿聰': '戴耀聰', '耀聰': '戴耀聰', '鈞': '鄒至鈞',
    '至鈞': '鄒至鈞', '科': '何科賢', '賢': '何科賢', '風': '王德風', '玲': '蔡采玲', '小玲': '蔡采玲',
    '九': '黃湘玲', '小九': '黃湘玲', '瑩': '戴嘉瑩', '嘉瑩': '戴嘉瑩', '珼': '袁翊珼', '貝貝': '袁翊珼', '珼珼': '袁翊珼',
    '軒': '謝立軒', '立軒': '謝立軒', '文': '羅文昕', '阿文': '羅文昕', '承': '葉承恩', '恩': '葉承恩',
    '承恩': '葉承恩', '羊羊': '楊洋', '洋洋': '楊洋', '洋': '楊洋', '楊': '楊洋', '力宏': '林立宏', '立宏': '林立宏', '皮皮': '尤怡蘋',
    '靖': '周靖', '周靖': '周靖', '雨': '陳佳琳', '小雨': '陳佳琳', '陳佳琳': '陳佳琳',
    '宇': '吳庭宇', '小宇': '吳庭宇', '吳庭宇': '吳庭宇', '姸': '姸姸', '姸姸': '姸姸'
};
const preservedNames = ['潘調', '勳調'];

const MAIN_FEES = {
    '入殮': 1000, '出殯': 1200, '入殮出殯': 1700, '入殮扛夫': 1700, '入殮火化': 1000,
    '禮生': 1000, '禮生出殯': 1400, '禮生扶棺': 1500, '午夜功德': 2000, '半日功德': 1000,
    '招待': 1200, '接體': 1200, '晉塔': 1000, '拼廳': 1200
};

const OTHER_FEES = {
    '頭七~滿七': 500, '女兒旬': 500, '接體空跑': 500, '退冰': 300, '驗屍/復驗': 500,
    '佈置': 500, '扶棺': 500, '安主/安位': 1000, '返主': 1000, '顧SPA': 500, '教會出殯': 1200, '移靈': 500
};

// Helper Functions
function isProjectVendor(vendor) { return vendor && projectVendors.some(pv => vendor.includes(pv)); }
function fixLocationTypo(location) { if (!location) return location; let fixed = location; for (const [typo, correct] of Object.entries(locationTypoFixes)) fixed = fixed.replace(typo, correct); return fixed; }
function fixVendorTypo(vendor) { if (!vendor) return vendor; let fixed = vendor; for (const [typo, correct] of Object.entries(vendorTypoFixes)) fixed = fixed.replace(typo, correct); return fixed; }
function convertName(name) { const trimmed = name.trim(); return preservedNames.includes(trimmed) ? trimmed : (nameMapping[trimmed] || trimmed); }
function formatTime(timeStr) { let cleaned = timeStr.replace(/」/g, '0').replace(/[:\s]/g, ''); if (cleaned.length === 1) return '0' + cleaned + '00'; if (cleaned.length === 2) return cleaned + '00'; if (cleaned.length === 3) return '0' + cleaned; if (cleaned.length === 4) return cleaned; return timeStr; }

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

// 判斷是否為「日期開頭行」(支援 YYYY/MM/DD, MM/DD, 或中文日期)
function isDateLine(line) {
    const t = line.trim();
    return /^(?:\d{4}[\/\-])?\d{1,2}[\/\-]\d{1,2}/.test(t) || /^\d{1,2}月\d{1,2}日?/.test(t);
}
// 單純只有日期，沒有工作內容
function isPureDateLine(line) {
    const t = line.trim();
    return /^(?:\d{4}[\/\-])?\d{1,2}[\/\-]\d{1,2}\s*$/.test(t) || /^\d{1,2}月\d{1,2}日?\s*$/.test(t);
}
function isSeparatorLine(line) { return /^[-—─]+$/.test(line.trim()) || line.trim() === ''; }
function isScheduleLine(line) {
    const trimmed = line.trim();
    // 略過星號備註行 (例：*采邑185 接體 14:30)
    if (!trimmed || trimmed.startsWith('*')) return false;
    // 格式一：時間在前 (11:00 龍圓 岡山 接體 或 06:30橋殯 邑威 出殯)
    if (/^\d{1,2}[:\s]?\d{2}(\s+|(?=[\u4e00-\u9fa5]))/.test(trimmed)) return true;
    if (/^\d{1,2}[:\s]?\d{0,2}\s+/.test(trimmed)) return true;
    // 格式二：時間在後 (龍圓 岡山 接體 11:00)
    if (/\s\d{1,2}[:\s]?\d{2}\s*$/.test(trimmed)) return true;
    // 格式三：廠商→時間→地點→工作 (松佑 13:30 靜心10 半日)
    if (/^[\u4e00-\u9fa5]+\s+\d{1,2}[:\s]?\d{2}\s+/.test(trimmed)) return true;
    return false;
}
function isCaseNameLine(line) { return line.trim().startsWith('案名：') || line.trim().startsWith('案名:'); }
function isRitualistLine(line) { return line.trim().startsWith('禮儀師：') || line.trim().startsWith('禮儀師:'); }

function isNamesLine(line) {
    const trimmed = line.trim();
    if (!trimmed) return false;
    // 支援中文字、空白、常見分隔符號（如頓號、逗號、斜線等）
    if (!/^[\u4e00-\u9fa5\s、，,／/\(\)（）]+$/.test(trimmed)) return false;

    // 排除特定非人名行（案名、禮儀師、排班場次行、日期行）
    if (isCaseNameLine(line) || isRitualistLine(line) || isScheduleLine(line) || isDateLine(line)) return false;

    return true;
}

function parseScheduleLine(line) {
    const trimmed = line.trim();

    // 格式一：時間在前 → 11:00 龍圓 岡山 接體 或 06:30橋殯 邑威 出殯
    const frontMatch = trimmed.match(/^(\d{1,2}[:\s]?\d{0,2})(?:\s+|(?=[\u4e00-\u9fa5]))(.+)$/);
    if (frontMatch) {
        const time = formatTime(frontMatch[1]);
        const rest = frontMatch[2];
        // 處理時間後面緊黏地點的情況 (如 06:30橋殯 → 時間=06:30, rest=橋殯 ...)
        const parts = rest.split(/\s+/);
        if (parts.length >= 3) return { time, location: fixLocationTypo(parts[0]), vendor: fixVendorTypo(parts[1]), workContent: normalizeWorkContent(parts.slice(2).join(' ')) };
        if (parts.length === 2) return { time, location: fixLocationTypo(parts[0]), vendor: fixVendorTypo(parts[1]), workContent: '' };
        return { time, location: fixLocationTypo(rest), vendor: '', workContent: '' };
    }

    // 格式二：時間在後 → 龍圓 岡山 接體 11:00
    const backMatch = trimmed.match(/^(.+)\s+(\d{1,2}[:\s]?\d{2})\s*$/);
    if (backMatch) {
        const time = formatTime(backMatch[2]);
        const parts = backMatch[1].split(/\s+/);
        if (parts.length >= 3) return { time, location: fixLocationTypo(parts[0]), vendor: fixVendorTypo(parts[1]), workContent: normalizeWorkContent(parts.slice(2).join(' ')) };
        if (parts.length === 2) return { time, location: fixLocationTypo(parts[0]), vendor: fixVendorTypo(parts[1]), workContent: '' };
        return { time, location: fixLocationTypo(backMatch[1]), vendor: '', workContent: '' };
    }

    // 格式三：廠商→時間→地點→工作 → 松佑 13:30 靜心10 半日
    const vendorFirstMatch = trimmed.match(/^([\u4e00-\u9fa5]+)\s+(\d{1,2}[:\s]?\d{2})\s+(.+)$/);
    if (vendorFirstMatch) {
        const vendor = fixVendorTypo(vendorFirstMatch[1]);
        const time = formatTime(vendorFirstMatch[2]);
        const parts = vendorFirstMatch[3].split(/\s+/);
        const location = fixLocationTypo(parts[0]);
        const workContent = normalizeWorkContent(parts.slice(1).join(' '));
        return { time, location, vendor, workContent };
    }

    return null;
}

const calculationCache = new Map();

/** 純計算邏輯，不含 cache。由 calculateAmount 呼叫 */
function _computeAmount(workContent, vendor, location) {
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
    if (/高雄龍巖/.test(vendor) && /禮生扶棺/.test(content)) return { amount: 1400, needsManualCheck: false };

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
    if (/拼廳/.test(content)) return { amount: 1200, needsManualCheck: false };
    if (/教會出殯/.test(content)) return { amount: 1200, needsManualCheck: false };
    if (/移靈/.test(content)) return { amount: 500, needsManualCheck: false };

    // Standard Items
    if (/扶棺/.test(content) && !/禮生/.test(content)) return { amount: 500, needsManualCheck: false };
    if (/入殮扛夫/.test(content) || (/入殮/.test(content) && /扛夫/.test(content))) return { amount: 1700, needsManualCheck: false };
    if (/入殮火化|入殮送火/.test(content)) return { amount: 1000, needsManualCheck: false };
    if (/入殮出殯/.test(content) || (/入殮/.test(content) && /出殯/.test(content))) return { amount: 1700, needsManualCheck: false };
    if (/禮生扶棺|禮生扛棺/.test(content)) return { amount: 1500, needsManualCheck: false };
    if (/禮生出殯/.test(content)) return { amount: 1400, needsManualCheck: false };
    if (/禮生/.test(content) && !/出殯|扶棺|扛棺/.test(content)) return { amount: 1000, needsManualCheck: false };
    if (/入殮/.test(content) && !/出殯|扛夫|火化|送火/.test(content)) return { amount: 1000, needsManualCheck: false };
    if (/出殯/.test(content) && !/入殮|禮生|回洗/.test(content)) return { amount: 1200, needsManualCheck: false };

    return { amount: 0, needsManualCheck: true };
}

function calculateAmount(workContent, vendor = '', location = '') {
    const cacheKey = `${workContent}|${vendor}|${location}`;
    if (calculationCache.has(cacheKey)) return calculationCache.get(cacheKey);
    const result = _computeAmount(workContent, vendor, location);
    calculationCache.set(cacheKey, result);
    return result;
}

const WORK_KEYWORDS = [
    '入殮', '出殯', '禮生', '功德', '接體', '退冰', '驗屍', '復驗', '相驗', '豎靈', '引魂', '佈置',
    '安主', '安位', '返主', '晉塔', '進塔', '招待', '拼廳', '扶棺', '洗身', '洗穿', '化妝', '更衣',
    '燒庫', '回洗', '加衣', '火化', '扛棺', '扛夫', '半日', '午夜', '藥懺', '頭七', '二七', '三七',
    '五七', '滿七', '女兒旬', '女兒七', 'SPA', 'spa', '移靈', '教會', '送火', '入出', '禮出',
    '禮扶', '禮扛', '入冰', '換罐', '樹葬', '協助', '告別式', '調'
];

function splitNamesAndWork(text, unknownNamesSet = null) {
    if (!text) return { names: [], extraWork: '' };
    // 支援以空白、頓號、逗號、斜線等拆分
    const parts = text.trim().split(/[、，,\s/／]+/).filter(Boolean);
    const names = [];
    const rest = [];

    for (const part of parts) {
        // 1. 檢查是否在既有名單或保留名單中
        const isKnownName = nameMapping[part] || Object.values(nameMapping).includes(part) || preservedNames.includes(part);
        if (isKnownName) {
            names.push(part);
            continue;
        }

        // 2. 檢查是否包含工作關鍵字
        const isWork = WORK_KEYWORDS.some(kw => part.includes(kw));
        if (isWork) {
            rest.push(part);
            continue;
        }

        // 3. 既非已知人名、亦非工作關鍵字，若符合中文名稱特徵（1~4字），視為未知人名候選
        if (/^[\u4e00-\u9fa5]{1,4}$/.test(part)) {
            names.push(part);
            if (unknownNamesSet) unknownNamesSet.add(part);
        } else {
            rest.push(part);
        }
    }
    return { names, extraWork: rest.join(' ') };
}

function parseScheduleData(text) {
    const lines = text.split('\n');
    const results = [];
    let currentDate = '';
    let lastSchedule = null;
    let lastScheduleLineNum = 0;
    let currentCaseName = '';
    let currentRitualist = '';

    const unknownNamesSet = new Set();
    const unrecognizedLines = [];
    const orphanSchedules = [];
    let scheduleCount = 0;

    // 輔助函式：若有未分配人員的場次，建立「（未指定人員）」以防遺漏
    function flushPendingSchedule(reason = '未分配人員') {
        if (!lastSchedule) return;
        orphanSchedules.push({
            lineNum: lastScheduleLineNum,
            date: lastSchedule.date,
            startTime: lastSchedule.startTime,
            location: lastSchedule.location,
            vendor: lastSchedule.vendor,
            workContent: lastSchedule.workContent,
        });

        const { amount, needsManualCheck } = calculateAmount(lastSchedule.workContent, lastSchedule.vendor, lastSchedule.location);
        let notes = currentRitualist || currentCaseName || '';
        notes = (notes ? notes + '；' : '') + reason;
        if (needsManualCheck) notes = (notes ? notes + '；' : '') + '需人工確認金額';
        if (isProjectVendor(lastSchedule.vendor)) notes = (notes ? notes + '；' : '') + '專案';

        results.push({
            date: lastSchedule.date,
            name: '（未指定人員）',
            startTime: lastSchedule.startTime,
            location: lastSchedule.location,
            vendor: lastSchedule.vendor,
            workContent: lastSchedule.workContent,
            amount: amount.toString(),
            paymentStatus: '未收',
            notes: notes,
            amount2: amount.toString(),
        });
        lastSchedule = null;
    }

    for (let i = 0; i < lines.length; i++) {
        const lineNum = i + 1;
        const line = lines[i];
        const trimmed = line.trim();

        if (isSeparatorLine(line)) continue;
        if (trimmed.startsWith('*')) continue;

        if (isDateLine(line)) {
            flushPendingSchedule();

            const fullYearMatch = trimmed.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
            const monthDayMatch = trimmed.match(/^(\d{1,2})[\/\-](\d{1,2})/);
            const chineseMatch = trimmed.match(/^(\d{1,2})月(\d{1,2})日?/);

            if (fullYearMatch) {
                currentDate = `${fullYearMatch[1]}/${fullYearMatch[2].padStart(2, '0')}/${fullYearMatch[3].padStart(2, '0')}`;
            } else if (monthDayMatch) {
                const currentYear = new Date().getFullYear();
                currentDate = `${currentYear}/${monthDayMatch[1].padStart(2, '0')}/${monthDayMatch[2].padStart(2, '0')}`;
            } else if (chineseMatch) {
                const currentYear = new Date().getFullYear();
                currentDate = `${currentYear}/${chineseMatch[1].padStart(2, '0')}/${chineseMatch[2].padStart(2, '0')}`;
            }

            // 判斷日期後面是否緊接工作內容 (例如 "2/1 06:30橋殯 邑威 出殯")
            const afterDate = trimmed.replace(/^(?:\d{4}[\/\-])?\d{1,2}[\/\-]\d{1,2}\s*/, '').replace(/^\d{1,2}月\d{1,2}日?\s*/, '').trim();
            if (afterDate && isScheduleLine(afterDate)) {
                const p = parseScheduleLine(afterDate);
                if (p) {
                    scheduleCount++;
                    const { names, extraWork } = splitNamesAndWork(p.workContent, unknownNamesSet);
                    const finalWorkContent = extraWork || p.workContent;

                    lastSchedule = {
                        date: currentDate,
                        startTime: p.time,
                        location: p.location,
                        vendor: p.vendor.replace(/[哥姐]/g, ''),
                        workContent: normalizeWorkContent(finalWorkContent),
                    };
                    lastScheduleLineNum = lineNum;

                    if (names.length > 0) {
                        for (let name of names) {
                            const fullName = convertName(name);
                            const isUnknown = unknownNamesSet.has(name);
                            const { amount, needsManualCheck } = calculateAmount(lastSchedule.workContent, lastSchedule.vendor, lastSchedule.location);
                            let notes = currentRitualist || currentCaseName || '';
                            if (isUnknown) notes = (notes ? notes + '；' : '') + `未建檔姓名(${name})`;
                            if (needsManualCheck) notes = (notes ? notes + '；' : '') + '需人工確認金額';
                            if (isProjectVendor(lastSchedule.vendor)) notes = (notes ? notes + '；' : '') + '專案';

                            results.push({
                                date: lastSchedule.date,
                                name: fullName,
                                startTime: lastSchedule.startTime,
                                location: lastSchedule.location,
                                vendor: lastSchedule.vendor,
                                workContent: lastSchedule.workContent,
                                amount: amount.toString(),
                                paymentStatus: '未收',
                                notes: notes,
                                amount2: amount.toString(),
                            });
                        }
                        lastSchedule = null;
                    }
                }
            } else {
                currentCaseName = '';
                currentRitualist = '';
            }
            continue;
        }

        if (isScheduleLine(line)) {
            flushPendingSchedule();

            const p = parseScheduleLine(line);
            if (p) {
                scheduleCount++;
                const { names, extraWork } = splitNamesAndWork(p.workContent, unknownNamesSet);
                const finalWorkContent = extraWork || p.workContent;

                lastSchedule = {
                    date: currentDate,
                    startTime: p.time,
                    location: p.location,
                    vendor: p.vendor.replace(/[哥姐]/g, ''),
                    workContent: normalizeWorkContent(finalWorkContent),
                };
                lastScheduleLineNum = lineNum;

                if (names.length > 0) {
                    for (let name of names) {
                        const fullName = convertName(name);
                        const isUnknown = unknownNamesSet.has(name);
                        const { amount, needsManualCheck } = calculateAmount(lastSchedule.workContent, lastSchedule.vendor, lastSchedule.location);
                        let notes = currentRitualist || currentCaseName || '';
                        if (isUnknown) notes = (notes ? notes + '；' : '') + `未建檔姓名(${name})`;
                        if (needsManualCheck) notes = (notes ? notes + '；' : '') + '需人工確認金額';
                        if (isProjectVendor(lastSchedule.vendor)) notes = (notes ? notes + '；' : '') + '專案';

                        results.push({
                            date: lastSchedule.date,
                            name: fullName,
                            startTime: lastSchedule.startTime,
                            location: lastSchedule.location,
                            vendor: lastSchedule.vendor,
                            workContent: lastSchedule.workContent,
                            amount: amount.toString(),
                            paymentStatus: '未收',
                            notes: notes,
                            amount2: amount.toString(),
                        });
                    }
                    lastSchedule = null;
                }
                currentCaseName = '';
                currentRitualist = '';
            }
            continue;
        }

        if (isCaseNameLine(line)) {
            currentCaseName = trimmed.replace(/^案名[：:]/, '').trim();
            continue;
        }

        if (isRitualistLine(line)) {
            currentRitualist = trimmed.replace(/^禮儀師[：:]/, '').trim();
            continue;
        }

        if (isNamesLine(line)) {
            if (lastSchedule) {
                const { names, extraWork } = splitNamesAndWork(trimmed, unknownNamesSet);
                if (extraWork) {
                    lastSchedule.workContent = normalizeWorkContent(lastSchedule.workContent + ' ' + extraWork);
                }

                if (names.length > 0) {
                    for (let name of names) {
                        const fullName = convertName(name);
                        const isUnknown = unknownNamesSet.has(name);
                        const { amount, needsManualCheck } = calculateAmount(lastSchedule.workContent, lastSchedule.vendor, lastSchedule.location);
                        let notes = currentRitualist || currentCaseName || '';
                        if (isUnknown) notes = (notes ? notes + '；' : '') + `未建檔姓名(${name})`;
                        if (needsManualCheck) notes = (notes ? notes + '；' : '') + '需人工確認金額';
                        if (isProjectVendor(lastSchedule.vendor)) notes = (notes ? notes + '；' : '') + '專案';

                        results.push({
                            date: lastSchedule.date,
                            name: fullName,
                            startTime: lastSchedule.startTime,
                            location: lastSchedule.location,
                            vendor: lastSchedule.vendor,
                            workContent: lastSchedule.workContent,
                            amount: amount.toString(),
                            paymentStatus: '未收',
                            notes: notes,
                            amount2: amount.toString(),
                        });
                    }
                    lastSchedule = null;
                } else {
                    unrecognizedLines.push({ lineNum, text: line, reason: '未能識別出有效人員姓名' });
                }
            } else {
                unrecognizedLines.push({ lineNum, text: line, reason: '此行看似人名但無對應場次' });
            }
            continue;
        }

        // 非空白且未被以上任何規則匹配
        if (trimmed) {
            unrecognizedLines.push({ lineNum, text: line, reason: '無法辨識的格式' });
        }
    }

    // 文字結束時檢查是否有殘餘未排人場次
    flushPendingSchedule();

    results.audit = {
        totalLines: lines.length,
        scheduleCount,
        generatedCount: results.length,
        unrecognizedLines,
        orphanSchedules,
        unknownNames: Array.from(unknownNamesSet),
    };

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
        if (parsedData.audit) {
            renderAuditBanner(parsedData.audit);
        }
        filterData();

        const duration = (performance.now() - startTime).toFixed(2);
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
        const banner = document.getElementById('parseAuditBanner');
        if (banner) banner.classList.add('hidden');
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

function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function renderAuditBanner(audit) {
    const banner = document.getElementById('parseAuditBanner');
    const titleContainer = document.getElementById('auditTitleContainer');
    const toggleBtn = document.getElementById('auditToggleBtn');
    const details = document.getElementById('auditDetails');
    const toggleIcon = document.getElementById('auditToggleIcon');
    if (!banner || !titleContainer || !details) return;

    const issueCount = (audit.unrecognizedLines?.length || 0) +
        (audit.orphanSchedules?.length || 0) +
        (audit.unknownNames?.length || 0);

    if (issueCount === 0) {
        banner.className = 'mx-6 mt-4 p-3.5 rounded-xl border bg-emerald-50/80 border-emerald-200 text-emerald-800 transition-all';
        titleContainer.innerHTML = `
            <svg class="w-5 h-5 text-emerald-600 flex-shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
                <polyline points="22 4 12 14.01 9 11.01"></polyline>
            </svg>
            <div class="text-xs">
                <span class="font-semibold text-emerald-900">解析檢查正常：</span>
                <span class="text-slate-600">共識別 ${audit.scheduleCount} 個場次，生成 ${audit.generatedCount} 筆排班明細，無遺漏或未辨識項目。</span>
            </div>
        `;
        toggleBtn.classList.add('hidden');
        details.classList.add('hidden');
        banner.classList.remove('hidden');
    } else {
        banner.className = 'mx-6 mt-4 p-4 rounded-xl border bg-amber-50/90 border-amber-200 text-amber-900 shadow-sm transition-all';
        titleContainer.innerHTML = `
            <svg class="w-5 h-5 text-amber-600 flex-shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"></path>
                <line x1="12" y1="9" x2="12" y2="13"></line>
                <line x1="12" y1="17" x2="12.01" y2="17"></line>
            </svg>
            <div class="text-xs">
                <span class="font-semibold text-amber-950">比對檢查提醒：</span>
                <span class="text-slate-700">生成 ${audit.generatedCount} 筆，發現 <strong class="text-amber-700 font-bold">${issueCount} 個需確認項目</strong>（點擊右側可展開/收合）</span>
            </div>
        `;
        toggleBtn.classList.remove('hidden');

        // 組裝詳細資訊清單
        let detailsHtml = '';

        if (audit.unknownNames && audit.unknownNames.length > 0) {
            detailsHtml += `
                <div class="p-2.5 bg-white/80 rounded-lg border border-amber-200/80">
                    <div class="font-semibold text-amber-900 flex items-center gap-1.5 mb-1">
                        <svg class="w-4 h-4 text-amber-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path>
                            <circle cx="9" cy="7" r="4"></circle>
                        </svg>
                        未建檔姓名（共 ${audit.unknownNames.length} 位）
                    </div>
                    <p class="text-slate-600 mb-1.5 leading-relaxed">已暫先為其產生成員資料（防止漏算），但因不在姓名對照表中，請確認是否為新同仁：</p>
                    <div class="flex flex-wrap gap-1.5">
                        ${audit.unknownNames.map(n => `<span class="px-2 py-0.5 bg-amber-100 text-amber-800 rounded font-medium border border-amber-300/60">${escapeHtml(n)}</span>`).join('')}
                    </div>
                </div>
            `;
        }

        if (audit.orphanSchedules && audit.orphanSchedules.length > 0) {
            detailsHtml += `
                <div class="p-2.5 bg-white/80 rounded-lg border border-red-200">
                    <div class="font-semibold text-red-800 flex items-center gap-1.5 mb-1">
                        <svg class="w-4 h-4 text-red-500" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="12" y1="8" x2="12" y2="12"></line>
                            <line x1="12" y1="16" x2="12.01" y2="16"></line>
                        </svg>
                        未分配人員場次（共 ${audit.orphanSchedules.length} 場）
                    </div>
                    <p class="text-slate-600 mb-1 leading-relaxed">以下場次有工作時間與內容，但下方未找到人員姓名，已暫列為「（未指定人員）」避免漏算：</p>
                    <ul class="list-disc list-inside space-y-0.5 text-slate-700">
                        ${audit.orphanSchedules.map(s => `<li>第 ${s.lineNum} 行：${escapeHtml(s.date)} ${escapeHtml(s.startTime || '')} ${escapeHtml(s.location || '')} ${escapeHtml(s.workContent || '')}</li>`).join('')}
                    </ul>
                </div>
            `;
        }

        if (audit.unrecognizedLines && audit.unrecognizedLines.length > 0) {
            detailsHtml += `
                <div class="p-2.5 bg-white/80 rounded-lg border border-slate-200">
                    <div class="font-semibold text-slate-700 flex items-center gap-1.5 mb-1">
                        <svg class="w-4 h-4 text-slate-500" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="4.93" y1="4.93" x2="19.07" y2="19.07"></line>
                        </svg>
                        未辨識或格式不符的文字行（共 ${audit.unrecognizedLines.length} 行）
                    </div>
                    <p class="text-slate-500 mb-1 leading-relaxed">以下內容不符合日期、場次或人名格式，已被略過，請確認是否有排班被漏掉：</p>
                    <ul class="space-y-1 font-mono text-slate-600">
                        ${audit.unrecognizedLines.map(l => `<li class="bg-slate-100/80 px-2 py-1 rounded">第 ${l.lineNum} 行：${escapeHtml(l.text)} <span class="text-slate-400 font-sans">(${escapeHtml(l.reason)})</span></li>`).join('')}
                    </ul>
                </div>
            `;
        }

        details.innerHTML = detailsHtml;
        details.classList.remove('hidden');
        if (toggleIcon) toggleIcon.style.transform = 'rotate(180deg)';
        banner.classList.remove('hidden');
    }
}

function toggleAuditDetails() {
    const details = document.getElementById('auditDetails');
    const icon = document.getElementById('auditToggleIcon');
    if (!details) return;
    const isHidden = details.classList.contains('hidden');
    details.classList.toggle('hidden', !isHidden);
    if (icon) {
        icon.style.transform = isHidden ? 'rotate(180deg)' : 'rotate(0deg)';
    }
}

function renderTable() {
    const tbody = document.getElementById('resultBody');
    tbody.innerHTML = filteredData.map(r => {
        const isOrphan = r.name === '（未指定人員）';
        const isUnknown = r.notes && r.notes.includes('未建檔姓名');
        const rowClass = isOrphan ? 'bg-red-50/50' : (isUnknown ? 'bg-amber-50/40' : '');
        const nameDisplay = isOrphan
            ? '<span class="text-red-600 font-bold bg-red-100/80 px-2 py-0.5 rounded text-xs">（未指定人員）</span>'
            : (isUnknown
                ? `<span class="text-amber-900 font-medium">${escapeHtml(r.name)}</span> <span class="text-[10px] bg-amber-100 text-amber-700 px-1 py-0.5 rounded font-normal">未建檔</span>`
                : escapeHtml(r.name));

        return `
        <tr class="hover:bg-indigo-50/30 transition-colors group ${rowClass}">
            <td class="p-4 text-slate-600 text-sm">${escapeHtml(r.date)}</td>
            <td class="p-4 text-slate-800 font-medium text-sm">${nameDisplay}</td>
            <td class="p-4 text-slate-600 font-mono text-xs">${escapeHtml(r.startTime)}</td>
            <td class="p-4 text-slate-600 text-sm">${escapeHtml(r.location)}</td>
            <td class="p-4 text-slate-600 text-sm">${escapeHtml(r.vendor)}</td>
            <td class="p-4 text-slate-600 text-sm">${escapeHtml(r.workContent)}</td>
            <td class="p-4 text-slate-700 font-medium text-sm">${escapeHtml(r.amount)}</td>
            <td class="p-4 text-slate-500 text-sm">${escapeHtml(r.paymentStatus)}</td>
            <td class="p-4 text-slate-500 text-xs italic">${escapeHtml(r.notes)}</td>
            <td class="p-4 text-slate-700 font-medium text-sm">${escapeHtml(r.amount2)}</td>
        </tr>`;
    }).join('');
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
    const rows = filteredData.map(r => [
        r.date, r.name, r.startTime, r.location, r.vendor, r.workContent, r.amount, r.paymentStatus, r.notes, r.amount2
    ].join('\t'));

    try {
        await navigator.clipboard.writeText(rows.join('\n'));
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
