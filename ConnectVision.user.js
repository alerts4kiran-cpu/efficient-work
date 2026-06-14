// ==UserScript==
// @name         ConnectVision - Amazon Connect Ultimate Monitoring Suite
// @namespace    http://tampermonkey.net/
// @version      8.0
// @description  ConnectVision: Unified dashboard for Amazon Connect with duration-based highlighting, activity tracking, break schedule compliance, single-tab enforcement, and pause detection
// @author       alerts4kiran-cpu
// @match        https://c2-na-prod.my.connect.aws/real-time-metrics*
// @match        https://c2-na-prod.awsapps.com/connect/real-time-metrics*
// @match        https://prod.aria.ats.a2z.com/shift-execution2*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_listValues
// @grant        GM_deleteValue
// @grant        GM_xmlhttpRequest
// @updateURL    https://raw.githubusercontent.com/alerts4kiran-cpu/efficient-work/main/ConnectVision.user.js
// @downloadURL  https://raw.githubusercontent.com/alerts4kiran-cpu/efficient-work/main/ConnectVision.user.js
// ==/UserScript==

(function () {
    'use strict';

    // ==================== CONFIGURATION ====================
    const CONFIG = {
        DEBOUNCE:         500,
        UPDATE_THROTTLE:  500,
        UNIFIED_INTERVAL: 5000,
        LEASE_KEY:        'connectvision_active_tab',
        LEASE_TTL:        8000,   // ms — lease expires if not renewed
        BROADCAST_CH:     'connectvision_channel',
        TIME_REGEX:       /^\d{1,2}:\d{2}(:\d{2})?$/,
        LETTER_REGEX:     /[a-zA-Z]/,
        TIME_RANGE_REGEX: /^\d{2}:\d{2}-\d{2}:\d{2}$/,
        DURATION_REGEX:   /^\d{2}:\d{2}:\d{2}$/,
        LOGIN_REGEX:      /^[a-z]{3,20}$/i,
        ACTIVITY_SUMMARY_KEYWORDS: [
            '=== ACTIVITY SUMMARY ===', 'Activity', 'Total', 'After contact work',
            'Available', 'On contact', 'Break', 'Lunch', 'Personal', 'Training',
            'Meeting', 'Project', 'Manager 1-1', 'Incoming', 'Missed', 'Outage',
            'Manager Approved', 'System/Power/Internet Outage', 'Skip Meeting',
            'Start Up', 'Team Huddle', 'Summary'
        ],
        QUIP_THREAD_ID: 'uIuCAKVKQGQU',
        QUIP_API_BASE:  'https://platform.quip-amazon.com/1',
        SP_SITE:        'https://amazon.sharepoint.com/sites/ComplianceCentral',
        SP_LIST:        'OutOfSlot_Breaks_Log'
    };

    const ATR_STORAGE_KEY = 'connectvision_atr_agents';
    const CRLF            = '\r\n';

    // ==================== SANITIZER ====================
    const SANITIZE_MAP = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
    const sanitize = (str) => String(str).replace(/[&<>"']/g, c => SANITIZE_MAP[c]);

    // Strip zero-width and invisible Unicode characters that Quip injects into cells
    const cleanText = (str) => String(str || '').replace(/[\u200b\u200c\u200d\ufeff\u00ad\u200e\u200f]/g, '').trim();

    // ==================== TAB IDENTITY ====================
    // Each tab gets a unique ID for this session
    const MY_TAB_ID = `cv_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;

    // ==================== STATE ====================
    const state = {
        settings: {
            yellowEnabled:    GM_getValue('yellowEnabled',    true),
            redEnabled:       GM_getValue('redEnabled',       true),
            blueEnabled:      GM_getValue('blueEnabled',      true),
            orangeEnabled:    GM_getValue('orangeEnabled',    true),
            yellowMinMinutes: GM_getValue('yellowMinMinutes', 2),
            yellowMinSeconds: GM_getValue('yellowMinSeconds', 0),
            yellowMaxMinutes: GM_getValue('yellowMaxMinutes', 5),
            yellowMaxSeconds: GM_getValue('yellowMaxSeconds', 0),
            redMinMinutes:    GM_getValue('redMinMinutes',    5),
            redMinSeconds:    GM_getValue('redMinSeconds',    0),
            breakMinMinutes:  GM_getValue('breakMinMinutes',  20),
            breakMinSeconds:  GM_getValue('breakMinSeconds',  0),
            lunchMinMinutes:  GM_getValue('lunchMinMinutes',  30),
            lunchMinSeconds:  GM_getValue('lunchMinSeconds',  0)
        },
        breakSchedules:           {},
        bufferMinutes:            0,
        bufferSeconds:            0,
        isScheduleMonitoring:     false,
        outOfSlotAgents:          [],
        hiddenOutOfSlotAgents:    new Set(),
        managerComments:          {},
        commentSaveTimers:        {},    // local save only — no uploads
        violationDetectionTimes:  {},
        activeTab:                'activity',
        activityDetailsSortColumn:    'duration',
        activityDetailsSortDirection: 'desc',
        isPaused:                 false,
        isActive:                 false,  // true when this tab holds the lease
        observer:                 null,
        updateTimeout:            null,
        leaseInterval:            null,
        dashboardEl:              null,
        takeoverBannerEl:         null,
        isUpdating:               false,
        lastUpdateTime:           0,
        cachedActivityDetails:    {},
        cachedAgentsByActivity:   {},
        tableCache:               null,
        tableCacheTime:           0,
        lastTableCount:           0,
        checkCount:               0,
        highAvailViolations:      new Map(),
        currentSelectedActivity:  null,
        violationLog:             (() => {
            try {
                const saved = JSON.parse(GM_getValue('cv_violationLog', '[]'));
                const savedDate = GM_getValue('cv_violationDate', '');
                const today = new Date().toISOString().slice(0, 10);
                return savedDate === today ? saved : [];
            } catch(e) { return []; }
        })(),
        uploadedHashes:           (() => {
            try {
                const saved = JSON.parse(GM_getValue('cv_uploadedHashes', '[]'));
                const savedDate = GM_getValue('cv_violationDate', '');
                const today = new Date().toISOString().slice(0, 10);
                return savedDate === today ? new Set(saved) : new Set();
            } catch(e) { return new Set(); }
        })(),
        lastViolationDate:        GM_getValue('cv_violationDate', null),
        lateToBreakAlertMin:      GM_getValue('lateToBreakAlertMin', 2),
        uploaderAlias:            ''
    };

    // Persist violation log to GM storage (survives reload/crash)
    function saveViolationLog() {
        GM_setValue('cv_violationLog', JSON.stringify(state.violationLog));
        GM_setValue('cv_uploadedHashes', JSON.stringify([...state.uploadedHashes]));
        GM_setValue('cv_violationDate', state.lastViolationDate || new Date().toISOString().slice(0, 10));
    }

    // ==================== UTILITY FUNCTIONS ====================

    const parseTimeToMinutes = (timeStr) => {
        if (!timeStr || timeStr === '-') return 0;
        const parts = timeStr.split(':');
        const len   = parts.length;
        return len === 2
            ? parseInt(parts[0]) + parseInt(parts[1]) / 60
            : len === 3
                ? parseInt(parts[0]) * 60 + parseInt(parts[1]) + parseInt(parts[2]) / 60
                : 0;
    };

    const convertToTotalMinutes = (min, sec) => min + sec / 60;

    const isActivitySummaryRow = (agentLogin) => {
        if (!agentLogin) return true;
        const lower = agentLogin.toString().trim().toLowerCase();
        return CONFIG.ACTIVITY_SUMMARY_KEYWORDS.some(kw => lower === kw.toLowerCase());
    };

    const getISTTimestamp = () => {
        const now = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Kolkata' }));
        return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}-${String(now.getMinutes()).padStart(2,'0')}-${String(now.getSeconds()).padStart(2,'0')}`;
    };

    const getTables = () => {
        const now = Date.now();
        if (state.tableCache && now - state.tableCacheTime < CONFIG.UNIFIED_INTERVAL) return state.tableCache;
        state.tableCache     = document.querySelectorAll('table tbody');
        state.tableCacheTime = now;
        return state.tableCache;
    };

    const showScheduleMessage = (message, type) => {
        const el = document.getElementById('cv-uploadMessage');
        if (!el) return;
        el.textContent = message;
        el.style.color = type === 'success' ? '#067d62' : type === 'error' ? '#d13212' : '#666';
        if (type !== 'error') setTimeout(() => { el.textContent = ''; }, 5000);
    };

    const getUniqueSlots = () => {
        const seen   = new Set();
        const unique = [];
        for (const login in state.breakSchedules) {
            const sd = state.breakSchedules[login];
            if (!sd.breaks) continue;
            for (const slot of sd.breaks) {
                const dur     = slot.end - slot.start;
                const display = `${Math.floor(slot.start/60)}:${String(slot.start%60).padStart(2,'0')}-${Math.floor(slot.end/60)}:${String(slot.end%60).padStart(2,'0')}`;
                if (!seen.has(display)) { seen.add(display); unique.push({ start: slot.start, end: slot.end, duration: dur, display }); }
            }
        }
        return unique;
    };

    // ==================== PAUSE DETECTION ====================
    const detectConnectPauseState = () => {
        // When paused, the RESUME (play triangle) SVG button is visible
        // Path signature unique to the play/resume triangle icon
        const svgs = document.querySelectorAll('svg');
        for (const svg of svgs) {
            if (svg.innerHTML.includes('M8 6.82v10.36c0 .79.87 1.27 1.54.84')) {
                return true;  // resume button is showing → page IS paused
            }
        }
        return false;         // pause button is showing → page is running
    };

    // ==================== MANAGER HELPER ====================
    const getManagerForAgent = (login, rowCells) => {
        const hierarchy = rowCells[5]?.textContent?.trim();
        if (hierarchy?.includes('/')) {
            const parts = hierarchy.split('/').map(p => p.trim());
            if (parts.length >= 2 && parts[1]) return parts[1];
        }
        return state.breakSchedules[login.toLowerCase()]?.manager || 'N/A';
    };

    // ==================== XLSX PARSER ====================

    async function parseXLSXFile(file) {
        const uint8Array    = new Uint8Array(await file.arrayBuffer());
        const files         = await extractZipFiles(uint8Array);
        const sharedStrings = parseSharedStrings(files['xl/sharedStrings.xml'] || '');
        return parseSheetData(files['xl/worksheets/sheet1.xml'] || '', sharedStrings);
    }

    async function extractZipFiles(data) {
        const files = {};
        let cdOffset = -1;
        for (let i = data.length - 22; i >= 0; i--) {
            if (data[i]===0x50 && data[i+1]===0x4b && data[i+2]===0x05 && data[i+3]===0x06) {
                cdOffset = readUInt32LE(data, i + 16); break;
            }
        }
        if (cdOffset === -1) throw new Error('Invalid ZIP file');
        let offset = cdOffset;
        while (offset < data.length - 4) {
            if (data[offset]!==0x50||data[offset+1]!==0x4b||data[offset+2]!==0x01||data[offset+3]!==0x02) break;
            const nameLen        = readUInt16LE(data, offset+28);
            const extraLen       = readUInt16LE(data, offset+30);
            const commentLen     = readUInt16LE(data, offset+32);
            const localHdrOffset = readUInt32LE(data, offset+42);
            const fileName       = new TextDecoder().decode(data.slice(offset+46, offset+46+nameLen));
            const compSize       = readUInt32LE(data, localHdrOffset+18);
            const localNameLen   = readUInt16LE(data, localHdrOffset+26);
            const localExtraLen  = readUInt16LE(data, localHdrOffset+28);
            const fileDataOffset = localHdrOffset + 30 + localNameLen + localExtraLen;
            const fileData       = data.slice(fileDataOffset, fileDataOffset + compSize);
            const compression    = readUInt16LE(data, localHdrOffset+8);
            if (compression === 0) {
                files[fileName] = new TextDecoder().decode(fileData);
            } else if (compression === 8) {
                try { files[fileName] = new TextDecoder().decode(await decompressDeflate(fileData)); }
                catch (e) { console.warn('Cannot decompress:', fileName); }
            }
            offset += 46 + nameLen + extraLen + commentLen;
        }
        return files;
    }

    async function decompressDeflate(data) {
        if (typeof DecompressionStream === 'undefined') return data;
        const ds = new DecompressionStream('deflate-raw');
        const w  = ds.writable.getWriter();
        w.write(data); w.close();
        const r = ds.readable.getReader();
        const chunks = [];
        while (true) { const {done, value} = await r.read(); if (done) break; chunks.push(value); }
        const total  = chunks.reduce((a, c) => a + c.length, 0);
        const result = new Uint8Array(total);
        let pos = 0;
        for (const c of chunks) { result.set(c, pos); pos += c.length; }
        return result;
    }

    const readUInt16LE = (d, o) => d[o] | (d[o+1] << 8);
    const readUInt32LE = (d, o) => d[o] | (d[o+1]<<8) | (d[o+2]<<16) | (d[o+3]<<24);

    function parseSharedStrings(xml) {
        const strings = []; const re = /<t[^>]*>([^<]*)<\/t>/g; let m;
        while ((m = re.exec(xml)) !== null) strings.push(m[1]);
        return strings;
    }

    function parseSheetData(xml, sharedStrings) {
        const cellData = {};
        const apply = (re, fn) => { let m; while ((m = re.exec(xml)) !== null) fn(m); };
        apply(/<c r="([A-Z]+)(\d+)"[^>]*t="s"[^>]*><v>(\d+)<\/v><\/c>/g,
            m => { const r=parseInt(m[2]); if(!cellData[r])cellData[r]={}; cellData[r][m[1]]=sharedStrings[parseInt(m[3])]||''; });
        apply(/<c r="([A-Z]+)(\d+)"[^>]*t="inlineStr"[^>]*><is><t>([^<]*)<\/t><\/is><\/c>/g,
            m => { const r=parseInt(m[2]); if(!cellData[r])cellData[r]={}; cellData[r][m[1]]=m[3]; });
        apply(/<c r="([A-Z]+)(\d+)"[^>]*><v>([^<]*)<\/v><\/c>/g,
            m => { const r=parseInt(m[2]); if(!cellData[r])cellData[r]={}; if(!cellData[r][m[1]])cellData[r][m[1]]=m[3]; });
        const maxRow = Object.keys(cellData).reduce((mx, k) => Math.max(mx, Number(k)), 0);
        const rows = [];
        for (let i = 1; i <= maxRow; i++) {
            const rowData = [];
            if (cellData[i]) {
                for (const col of Object.keys(cellData[i]).sort()) {
                    const idx = columnToIndex(col);
                    while (rowData.length < idx) rowData.push('');
                    rowData[idx] = cellData[i][col];
                }
            }
            rows.push(rowData);
        }
        return rows;
    }

    function columnToIndex(col) {
        let idx = 0;
        for (let i = 0; i < col.length; i++) idx = idx * 26 + (col.charCodeAt(i) - 64);
        return idx - 1;
    }

    // ==================== SCHEDULE PARSING ====================

    // ── Schedule grid helpers ────────────────────────────────────────────────

    function extractManagerAlias(raw) {
        if (!raw) return 'N/A';
        let s = String(raw).trim();
        const slashIdx = s.lastIndexOf(' / ');
        if (slashIdx !== -1) s = s.slice(slashIdx + 3).trim();
        s = s.replace(/_[A-Za-z0-9]+$/i, '').trim();
        return s || 'N/A';
    }

    function parseDayHeaderWeekday(headerStr) {
        const m = String(headerStr || '').trim().match(/^([A-Za-z]+)/);
        return m ? m[1].toLowerCase() : '';
    }

    function isScheduleSkipRow(row, col) {
        if (!row) return true;
        const cell = row[col];
        if (cell === null || cell === undefined || String(cell).trim() === '') return true;
        const lower = String(cell).trim().toLowerCase();
        // Skip header/label rows (login, day abbreviations, manager, etc.)
        const skipWords = ['login','agent','employee','manager','sun','mon','tue','tues','wed','thu','fri','sat',
                           'sunday','monday','tuesday','wednesday','thursday','friday','saturday','name','alias'];
        return skipWords.includes(lower);
    }

    function slotToDisplay(slot) {
        const sh = Math.floor(slot.start / 60), sm = slot.start % 60;
        const eh = Math.floor(slot.end   / 60), em = slot.end   % 60;
        return `${sh}:${String(sm).padStart(2,'0')}-${eh}:${String(em).padStart(2,'0')}`;
    }

    // ── Grid-aware schedule parser (multi-day side-by-side layout) ────────────

    function parseScheduleFromGrid(data) {
        state.breakSchedules = {};
        let successCount = 0;

        console.log('[ConnectVision] parseScheduleFromGrid called:', data.length, 'rows');
        if (data[0]) console.log('[ConnectVision] Row 0 (first 10 cols):', JSON.stringify(data[0].slice(0, 10)));
        if (data[1]) console.log('[ConnectVision] Row 1 (first 10 cols):', JSON.stringify(data[1].slice(0, 10)));
        if (data[2]) console.log('[ConnectVision] Row 2 (first 10 cols):', JSON.stringify(data[2].slice(0, 10)));

        // Build today's identifiers for matching
        const now = new Date();
        const dayNum = now.getDate();
        const suffix = ([,'st','nd','rd'][(dayNum%100-20)%10] || [,'st','nd','rd'][dayNum%100] || 'th');
        const todayDate = (dayNum + suffix).toLowerCase();  // e.g. "14th"
        const todayWeekday = now.toLocaleDateString('en-US', { weekday: 'long' }).toLowerCase();

        console.log('[ConnectVision] Matching:', todayWeekday, todayDate);

        // Scan first 3 rows for the header row (merged day+date cells like "Sunday(14th June)")
        let headerRowIdx = -1;
        let targetCol = -1;

        for (let r = 0; r < Math.min(4, data.length); r++) {
            const row = data[r] || [];
            // Pass 1: find a cell that starts with today's weekday AND contains today's date
            for (let c = 0; c < row.length; c++) {
                const cell = String(row[c] || '').trim().toLowerCase();
                if (cell.startsWith(todayWeekday) && cell.includes(todayDate)) {
                    headerRowIdx = r;
                    targetCol = c;
                    break;
                }
            }
            if (targetCol !== -1) break;
        }

        // Pass 2: find by today's date only (in case weekday format differs)
        if (targetCol === -1) {
            for (let r = 0; r < Math.min(4, data.length); r++) {
                const row = data[r] || [];
                for (let c = 0; c < row.length; c++) {
                    const cell = String(row[c] || '').trim().toLowerCase();
                    if (cell.includes(todayDate) && cell.length > 3) {
                        headerRowIdx = r;
                        targetCol = c;
                        break;
                    }
                }
                if (targetCol !== -1) break;
            }
        }

        // Pass 3: find by today's weekday only (if no date in headers)
        if (targetCol === -1) {
            for (let r = 0; r < Math.min(4, data.length); r++) {
                const row = data[r] || [];
                for (let c = 0; c < row.length; c++) {
                    const cell = String(row[c] || '').trim().toLowerCase();
                    if (cell.startsWith(todayWeekday)) {
                        headerRowIdx = r;
                        targetCol = c;
                        break;
                    }
                }
                if (targetCol !== -1) break;
            }
        }

        if (targetCol === -1) {
            console.log('[ConnectVision] Could not find today\'s section in header rows — falling back to simple parser');
            return parseScheduleSimple(data);
        }

        const cOff = targetCol;
        const dayLabel = String(data[headerRowIdx][cOff] || '').trim();
        console.log('[ConnectVision] Found today at row', headerRowIdx, 'col', cOff, ':', dayLabel);

        // Data starts after header row + sub-header row (login/manager/break labels)
        let dataStart = -1;
        for (let r = headerRowIdx + 2; r < Math.min(headerRowIdx + 8, data.length); r++) {
            if (isScheduleSkipRow(data[r], cOff)) continue;
            const cell = String(data[r][cOff] || '').trim().toLowerCase();
            if (/^[a-z]{3,20}$/.test(cell)) { dataStart = r; break; }
        }

        if (dataStart === -1) {
            showScheduleMessage('\u274c No data rows found for ' + dayLabel, 'error');
            return;
        }

        console.log('[ConnectVision] Data starts at row', dataStart, '- first login:', String(data[dataStart][cOff] || '').trim());

        for (let r = dataStart; r < data.length; r++) {
            const row = data[r];
            if (isScheduleSkipRow(row, cOff)) continue;
            const login = cleanText(row[cOff]).toLowerCase();
            if (!login || !/^[a-z]{3,20}$/.test(login)) continue;
            try {
                const manager  = extractManagerAlias(cleanText(row[cOff + 1]));
                const b1       = cleanText(row[cOff + 3]);
                const b2       = cleanText(row[cOff + 4]);
                const b3       = cleanText(row[cOff + 5]);
                const allSlots = [
                    ...(b1 ? parseTimeSlot(b1) : []),
                    ...(b2 ? parseTimeSlot(b2) : []),
                    ...(b3 ? parseTimeSlot(b3) : [])
                ].filter(s => s != null);  // Keep all parsed slots (ATR-aligned)
                if (allSlots.length === 0) continue;
                state.breakSchedules[login] = {
                    manager, breaks: allSlots,
                    break10: allSlots[0] ? slotToDisplay(allSlots[0]) : 'N/A',
                    break20: allSlots[1] ? slotToDisplay(allSlots[1]) : 'N/A',
                    break30: allSlots[2] ? slotToDisplay(allSlots[2]) : 'N/A'
                };
                successCount++;
            } catch (err) { console.error(`Schedule parse error for ${login}:`, err); }
        }

        const statusEl = document.getElementById('cv-scheduleStatus');
        if (successCount > 0) {
            if (statusEl) statusEl.textContent = `\u2705 ${successCount} agents loaded (${dayLabel})`;
            showScheduleMessage(`\u2705 Loaded ${successCount} schedules (${dayLabel})`, 'success');
        } else {
            showScheduleMessage('\u274c No valid schedules found for ' + dayLabel, 'error');
        }
    }

    // ── Simple single-section fallback (original logic) ───────────────────────

    function parseScheduleSimple(data) {
        state.breakSchedules = {};
        let successCount     = 0;
        let headerRowIndex   = -1;
        for (let i = 0; i < Math.min(5, data.length); i++) {
            if (!data[i]) continue;
            // Search all columns in the row for 'login'
            const hasLogin = data[i].some(cell => cell && String(cell).toLowerCase().includes('login'));
            if (hasLogin) { headerRowIndex = i; break; }
        }
        if (headerRowIndex === -1) { showScheduleMessage('\u274c Header row not found.', 'error'); return; }
        const headers    = data[headerRowIndex];
        const loginCol   = headers.findIndex(h => h && String(h).toLowerCase().includes('login'));
        const managerCol = headers.findIndex(h => h && String(h).toLowerCase().includes('manager'));
        const break10Col = headers.findIndex(h => h && String(h).toLowerCase().includes('break') && String(h).includes('10'));
        const break20Col = headers.findIndex(h => h && String(h).toLowerCase().includes('break') && String(h).includes('20'));
        const break30Col = headers.findIndex(h => h && String(h).toLowerCase().includes('break') && String(h).includes('30'));
        if (loginCol===-1||break10Col===-1||break20Col===-1||break30Col===-1) {
            showScheduleMessage('\u274c Required columns not found.', 'error'); return;
        }
        for (let i = headerRowIndex + 1; i < data.length; i++) {
            const row = data[i];
            if (!row || !row[loginCol]) continue;
            const login = cleanText(row[loginCol]).toLowerCase();
            if (!login) continue;
            try {
                const break10 = row[break10Col] ? String(row[break10Col]).trim() : '';
                const break20 = row[break20Col] ? String(row[break20Col]).trim() : '';
                const break30 = row[break30Col] ? String(row[break30Col]).trim() : '';
                const breaks  = [
                    ...(break10 ? parseTimeSlot(break10) : []),
                    ...(break20 ? parseTimeSlot(break20) : []),
                    ...(break30 ? parseTimeSlot(break30) : [])
                ];
                if (breaks.length > 0) {
                    state.breakSchedules[login] = {
                        manager: managerCol !== -1 ? extractManagerAlias(String(row[managerCol] || 'N/A').trim()) : 'N/A',
                        breaks, break10: break10||'N/A', break20: break20||'N/A', break30: break30||'N/A'
                    };
                    successCount++;
                }
            } catch (err) { console.error(`Schedule parse error for ${login}:`, err); }
        }
        const statusEl = document.getElementById('cv-scheduleStatus');
        if (successCount > 0) {
            if (statusEl) statusEl.textContent = `\u2705 ${successCount} agent schedules loaded`;
            showScheduleMessage(`\u2705 Loaded ${successCount} schedules`, 'success');
        } else {
            showScheduleMessage('\u274c No valid schedules found', 'error');
        }
    }

    function parseTimeSlot(timeStr) {
        if (!timeStr) return [];
        return String(timeStr).trim().split(/[,;]/).reduce((slots, range) => {
            const m = range.trim().match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
            if (m) slots.push({ start: parseInt(m[1])*60+parseInt(m[2]), end: parseInt(m[3])*60+parseInt(m[4]) });
            return slots;
        }, []);
    }

    function checkIfOutOfSchedule(agentLogin) {
        const sd = state.breakSchedules[agentLogin.toLowerCase()];
        if (!sd?.breaks?.length) return false;
        const now     = new Date();
        const current = now.getHours()*60 + now.getMinutes() + now.getSeconds()/60;
        const buffer  = state.bufferMinutes + state.bufferSeconds/60;
        for (const slot of sd.breaks) {
            if (current >= slot.start - buffer && current <= slot.end + buffer) return false;
        }
        return true;
    }

    // Calculate Out-of-Slot time: how long past the nearest ended slot
    function getOutOfSlotTime(agentLoginLower) {
        const sd = state.breakSchedules[agentLoginLower];
        if (!sd?.breaks?.length) return 'N/A';
        const now = new Date();
        const currentMin = now.getHours()*60 + now.getMinutes() + now.getSeconds()/60;
        // Find the nearest slot that has already ended (current > slot.end)
        let nearestSlot = null, minDist = Infinity;
        for (const slot of sd.breaks) {
            if (currentMin > slot.end) {
                const dist = currentMin - slot.end;
                if (dist < minDist) { minDist = dist; nearestSlot = slot; }
            }
        }
        if (!nearestSlot) return 'N/A';
        const oosMinutes = Math.floor(minDist);
        const oosSeconds = Math.round((minDist - oosMinutes) * 60);
        return `${String(oosMinutes).padStart(2,'0')}:${String(oosSeconds).padStart(2,'0')}`;
    }

    // Returns the agent's scheduled slot matching current time (or nearest)
    function getSlotForViolation(agentLoginLower) {
        const sd = state.breakSchedules[agentLoginLower];
        if (!sd?.breaks?.length) return null;
        const buf     = state.bufferMinutes + state.bufferSeconds / 60;
        const current = new Date().getHours()*60 + new Date().getMinutes() + new Date().getSeconds()/60;
        for (const slot of sd.breaks) {
            if (current >= slot.start - buf && current <= slot.end + buf) return slot;
        }
        let best = null, bestDist = Infinity;
        for (const slot of sd.breaks) {
            const dist = Math.abs(current - (slot.start + slot.end) / 2);
            if (dist < bestDist) { bestDist = dist; best = slot; }
        }
        return best;
    }

    function buildViolationHash(login, slot) {
        const d    = new Date();
        const date = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
        const hhmm = `${String(Math.floor(slot.start/60)).padStart(2,'0')}${String(slot.start%60).padStart(2,'0')}`;
        const bucket = Math.floor((d.getHours()*60 + d.getMinutes()) / 30); // 30-min window
        return `${login.toLowerCase()}|${date}|${hhmm}|${bucket}`;
    }

    function getUploaderAlias() {
        try {
            for (const svg of document.querySelectorAll('svg')) {
                if (svg.innerHTML.includes('m8 11 4-6H4l4 6Z')) {
                    const parent = svg.closest('div');
                    if (parent) {
                        const txt = Array.from(parent.childNodes)
                            .find(n => n.nodeType === 3 && n.textContent.trim());
                        if (txt) return txt.textContent.trim();
                    }
                }
            }
        } catch (e) {}
        return '';
    }

    function getCurrentSlotDisplay() {
        const current = new Date().getHours()*60 + new Date().getMinutes() + new Date().getSeconds()/60;
        const found   = getUniqueSlots().find(s => current >= s.start && current <= s.end);
        return found ? `${found.display} (${found.duration} Mins)` : 'No active slot';
    }

    function showOOSMessage(msg, type) {
        const el = document.getElementById('cv-oos-uploadMsg');
        if (!el) return;
        el.textContent = msg;
        el.style.color = type === 'success' ? '#067d62' : type === 'error' ? '#d13212' : '#555';
        if (type !== 'error') setTimeout(() => { el.textContent = ''; }, 6000);
    }

    // ==================== ARIA ATR SCANNER ====================
    const scanAndStoreARIAStatus = () => {
        if (!window.location.href.includes('aria.ats.a2z.com')) return;
        const atrStatusMap = {};
        try {
            let activeTab = false;
            const tab = document.querySelector('#mat-tab-label-0-1');
            if (tab?.getAttribute('aria-selected') === 'true') {
                activeTab = true;
            } else {
                for (const t of document.querySelectorAll('[role="tab"]')) {
                    if (t.getAttribute('aria-selected')==='true' && t.textContent.trim().toLowerCase().includes('employee')) { activeTab = true; break; }
                }
            }
            if (!activeTab) return;
            const table = document.querySelector('table tbody');
            if (!table) return;
            for (const row of table.querySelectorAll('tr')) {
                const cells = row.querySelectorAll('td');
                if (cells.length < 7) continue;
                const empText = cells[1]?.textContent?.trim();
                if (!empText) continue;
                const lm = empText.match(/@(\w+)/);
                if (!lm?.[1]) continue;
                const login      = lm[1].toLowerCase();
                const statusText = cells[6]?.textContent?.trim() || '';
                atrStatusMap[login] = statusText.includes('Online') && statusText.includes('ATR')
                    ? 'Online' : statusText.includes('Offline') ? 'Offline' : 'Unknown';
            }
            GM_setValue(ATR_STORAGE_KEY, JSON.stringify({ timestamp: Date.now(), agents: atrStatusMap }));
        } catch (e) { console.error('[ConnectVision ARIA]', e); }
    };

    const getATRStatusFromStorage = () => {
        try {
            const raw = GM_getValue(ATR_STORAGE_KEY, null);
            if (!raw) return {};
            const data = JSON.parse(raw);
            if (Date.now() - data.timestamp > 120000) return {};
            return data.agents || {};
        } catch (e) { return {}; }
    };
    // ==================== SINGLE-TAB GUARD RAIL ====================
    // Uses localStorage as a cross-tab shared lease.
    // BroadcastChannel sends instant "takeover" and "shutdown" messages.

    let _bc = null;  // BroadcastChannel instance

    const getLease = () => {
        try {
            const raw = localStorage.getItem(CONFIG.LEASE_KEY);
            return raw ? JSON.parse(raw) : null;
        } catch (e) { return null; }
    };

    const writeLease = () => {
        try {
            localStorage.setItem(CONFIG.LEASE_KEY, JSON.stringify({
                tabId:     MY_TAB_ID,
                timestamp: Date.now()
            }));
        } catch (e) { /* ignore */ }
    };

    const clearLease = () => {
        try {
            const lease = getLease();
            // Only clear if WE own it
            if (lease?.tabId === MY_TAB_ID) localStorage.removeItem(CONFIG.LEASE_KEY);
        } catch (e) { /* ignore */ }
    };

    const isLeaseValid = (lease) => {
        if (!lease) return false;
        return (Date.now() - lease.timestamp) < CONFIG.LEASE_TTL;
    };

    const isMine = (lease) => lease?.tabId === MY_TAB_ID;

    // ── Takeover Banner (shown on non-active tabs) ──────────────────

    const showTakeoverBanner = () => {
        if (state.takeoverBannerEl) return;

        const banner = document.createElement('div');
        banner.id    = 'cv-takeoverBanner';
        banner.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            z-index: 99999;
            background: #232f3e;
            color: #fff;
            border: 3px solid #ff9900;
            border-radius: 12px;
            padding: 28px 36px;
            text-align: center;
            box-shadow: 0 8px 32px rgba(0,0,0,0.5);
            font-family: Arial, sans-serif;
            min-width: 360px;
        `;
        banner.innerHTML = `
            <button id="cv-bannerMinBtn" style="
                position:absolute; top:8px; right:10px;
                background:#ff9900; border:none; color:#232f3e;
                border-radius:4px; padding:1px 10px; cursor:pointer;
                font-size:16px; font-weight:bold; line-height:1.4;
            " title="Minimise">−</button>
            <div style="font-size:32px;margin-bottom:10px">🖥️</div>
            <div style="font-size:17px;font-weight:bold;margin-bottom:8px;color:#ff9900">ConnectVision</div>
            <div style="font-size:14px;color:#ccc;margin-bottom:20px;line-height:1.6">
                ConnectVision is already running<br>on another tab in this browser.
            </div>
            <button id="cv-takeoverBtn" style="
                padding: 10px 28px;
                background: #ff9900;
                color: #232f3e;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                font-size: 15px;
                font-weight: bold;
                box-shadow: 0 2px 8px rgba(0,0,0,0.3);
            ">⚡ Take Over</button>
            <div style="margin-top:12px;font-size:11px;color:#888">Clicking Take Over will close ConnectVision on the other tab.</div>
        `;
        document.body.appendChild(banner);
        state.takeoverBannerEl = banner;

        // Mini pill shown next to the ConnectVision button when banner is minimised
        const mini = document.createElement('button');
        mini.id           = 'cv-takeoverMini';
        mini.textContent  = '⚡ Take Over';
        mini.style.cssText = `
            display: none;
            position: fixed;
            top: 10px;
            z-index: 10001;
            padding: 8px 18px;
            background: #ff9900;
            color: #232f3e;
            border: 2px solid #232f3e;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
            box-shadow: 0 2px 8px rgba(0,0,0,0.3);
        `;
        document.body.appendChild(mini);

        const positionMini = () => {
            const openBtn = document.getElementById('cv-openBtn');
            if (openBtn) {
                const rect = openBtn.getBoundingClientRect();
                mini.style.top        = rect.top + 'px';
                mini.style.left       = (rect.right + 10) + 'px';
                mini.style.transform  = 'none';
            } else {
                mini.style.top   = '10px';
                mini.style.left  = 'auto';
                mini.style.right = '10px';
            }
        };

        document.getElementById('cv-bannerMinBtn').addEventListener('click', () => {
            banner.style.display = 'none';
            positionMini();
            mini.style.display = 'block';
        });

        mini.addEventListener('click', () => {
            mini.style.display   = 'none';
            banner.style.display = 'block';
        });

        document.getElementById('cv-takeoverBtn').addEventListener('click', () => {
            // Broadcast takeover command to other tabs
            _bc.postMessage({ type: 'TAKEOVER', fromTabId: MY_TAB_ID });
            // Give the other tab 200ms to acknowledge, then claim the lease
            setTimeout(() => {
                writeLease();
                banner.remove();
                mini.remove();
                state.takeoverBannerEl = null;
                activateThisTab();
            }, 200);
        });
    };

    const hideTakeoverBanner = () => {
        if (state.takeoverBannerEl) {
            state.takeoverBannerEl.remove();
            state.takeoverBannerEl = null;
        }
        const mini = document.getElementById('cv-takeoverMini');
        if (mini) mini.remove();
    };

    // ── Deactivate this tab (called when another tab takes over) ────

    const deactivateThisTab = () => {
        console.log('[ConnectVision] 🛑 Deactivated — another tab took over');
        state.isActive = false;

        // Stop observer and intervals
        if (state.observer)      { state.observer.disconnect(); state.observer = null; }
        if (state.leaseInterval) { clearInterval(state.leaseInterval); state.leaseInterval = null; }
        clearTimeout(state.updateTimeout);

        // Hide dashboard
        if (state.dashboardEl)   state.dashboardEl.style.display = 'none';
        const openBtn = document.getElementById('cv-openBtn');
        if (openBtn) openBtn.style.display = 'none';

        // Show the takeover banner so this tab can reclaim if needed
        showTakeoverBanner();
    };

    // ── Activate this tab ────────────────────────────────────────────

    const activateThisTab = () => {
        state.isActive = true;
        writeLease();

        // Start lease renewal heartbeat
        if (state.leaseInterval) clearInterval(state.leaseInterval);
        state.leaseInterval = setInterval(() => {
            writeLease();
        }, Math.floor(CONFIG.LEASE_TTL / 2));

        // Show UI
        const openBtn = document.getElementById('cv-openBtn');
        if (openBtn) openBtn.style.display = 'block';

        if (!state.dashboardEl) {
            createDashboard();
        } else {
            state.dashboardEl.style.display = 'flex';
        }

        // Boot monitoring
        unifiedUpdate();
        attachObserver();
        setInterval(unifiedUpdate, CONFIG.UNIFIED_INTERVAL);

        // Auto-load schedule from Quip (if token exists) then auto-start monitoring
        const _autoToken = GM_getValue('quipToken', '');
        if (_autoToken) {
            console.log('[ConnectVision] Auto-loading schedule from Quip...');
            setTimeout(() => {
                loadScheduleFromQuip(() => {
                    if (Object.keys(state.breakSchedules).length > 0 && !state.isScheduleMonitoring) {
                        state.isScheduleMonitoring = true;
                        const _btn = document.getElementById('cv-toggleMonitor');
                        if (_btn) { _btn.textContent = '⏹ Stop Schedule Monitoring'; _btn.style.background = '#d13212'; }
                        showScheduleMessage('✅ Auto-monitoring started', 'success');
                        console.log('[ConnectVision] Auto-monitoring started');
                    }
                });
            }, 3000);
        }

        // Re-attach observer guard
        setInterval(() => {
            const count = document.querySelectorAll('table tbody').length;
            if (count !== state.lastTableCount) {
                state.tableCache = null;
                attachObserver();
            }
        }, 30000);

        console.log('[ConnectVision] ✅ Active on this tab:', MY_TAB_ID);
    };

    // ── Initialise guard rail ────────────────────────────────────────

    const initTabGuard = () => {
        // Open BroadcastChannel for instant cross-tab messaging
        try {
            _bc = new BroadcastChannel(CONFIG.BROADCAST_CH);
        } catch (e) {
            // BroadcastChannel not available — just run without guard
            console.warn('[ConnectVision] BroadcastChannel unavailable, running without tab guard');
            activateThisTab();
            return;
        }

        _bc.onmessage = (event) => {
            const msg = event.data;
            if (!msg) return;

            if (msg.type === 'TAKEOVER' && msg.fromTabId !== MY_TAB_ID) {
                // Another tab is claiming ownership
                if (state.isActive) {
                    deactivateThisTab();
                }
            }

            if (msg.type === 'PING') {
                // Another tab is checking if someone is active
                if (state.isActive) {
                    _bc.postMessage({ type: 'PONG', fromTabId: MY_TAB_ID });
                }
            }
        };

        const lease = getLease();

        if (!isLeaseValid(lease)) {
            // No valid lease — claim it
            writeLease();
            activateThisTab();
        } else if (isMine(lease)) {
            // We somehow already have the lease (e.g. page reload)
            activateThisTab();
        } else {
            // Another tab holds a valid lease
            console.log('[ConnectVision] 🔒 Another tab is active. Showing takeover banner.');
            // Create the open button (hidden) and banner
            createOpenButton();
            showTakeoverBanner();
        }
    };

    // ==================== MONITORING ENGINE ====================

    const unifiedUpdate = () => {
        if (!state.isActive) return;
        state.checkCount++;

        state.isPaused = detectConnectPauseState();

        const atrStatusMap   = getATRStatusFromStorage();
        state.outOfSlotAgents = [];

        // Reset violation log + uploaded hashes at day boundary
        const _today = new Date().toISOString().slice(0, 10);
        if (state.lastViolationDate !== _today) {
            state.lastViolationDate = _today;
            state.violationLog      = [];
            state.uploadedHashes    = new Set();
            saveViolationLog();
        }

        const tables = getTables();
        const len    = tables.length;

        const yellowMin = convertToTotalMinutes(state.settings.yellowMinMinutes, state.settings.yellowMinSeconds);
        const yellowMax = convertToTotalMinutes(state.settings.yellowMaxMinutes, state.settings.yellowMaxSeconds);
        const redMin    = convertToTotalMinutes(state.settings.redMinMinutes,    state.settings.redMinSeconds);
        const breakMin  = convertToTotalMinutes(state.settings.breakMinMinutes,  state.settings.breakMinSeconds);
        const lunchMin  = convertToTotalMinutes(state.settings.lunchMinMinutes,  state.settings.lunchMinSeconds);

        const activityDetails  = {};
        const agentsByActivity = {};
        const seenOutOfSlot    = new Set();
        const processedAgents  = new Set();
        let   highlightedCount = 0;

        const managerLoginsSet = new Set(
            Object.values(state.breakSchedules).map(s => s.manager?.toLowerCase()).filter(Boolean)
        );

        for (let t = 0; t < len; t++) {
            const rows   = tables[t].rows;
            const rowLen = rows.length;
            for (let i = 0; i < rowLen; i++) {
                const row   = rows[i];
                const cells = row.cells;

                if (cells.length <= 2 || row.querySelectorAll('th').length > 5) continue;

                row.style.cssText = '';
                row.removeAttribute('data-highlighted');

                const agentLogin      = cells[0]?.textContent?.trim();
                const agentLoginLower = agentLogin?.toLowerCase();
                const activity        = cells[2]?.textContent?.trim();
                const activityLower   = activity?.toLowerCase();
                const durationText    = cells[4]?.textContent?.trim();

                if (!agentLogin || agentLogin === 'Agent Login') continue;
                if (isActivitySummaryRow(agentLogin))           continue;
                if (/^\d+$/.test(agentLogin))                   continue;
                if (!CONFIG.LOGIN_REGEX.test(agentLogin))       continue;
                if (processedAgents.has(agentLoginLower))       continue;
                if (CONFIG.TIME_RANGE_REGEX.test(durationText)) continue;
                if (!activity || activity === 'N/A' || activity === '-' || activity === '') continue;

                let hasTimeRange = false;
                for (let j = 0; j < cells.length; j++) {
                    if (CONFIG.TIME_RANGE_REGEX.test(cells[j]?.textContent?.trim())) { hasTimeRange = true; break; }
                }
                if (hasTimeRange) continue;

                processedAgents.add(agentLoginLower);

                // PRIORITY 1: Purple (out-of-slot) — only when NOT paused
                if (!state.isPaused && state.isScheduleMonitoring &&
                    (activityLower === 'break' || activityLower === 'lunch')) {
                    const hasSchedule = !!state.breakSchedules[agentLoginLower];
                    const isOutOfSlot = checkIfOutOfSchedule(agentLoginLower);
                    if (state.checkCount <= 3 && (activityLower === 'lunch' || !hasSchedule)) {
                        console.log(`[OOS Debug] ${agentLoginLower}: hasSchedule=${hasSchedule}, isOutOfSlot=${isOutOfSlot}, activity=${activity}`);
                    }
                    if (isOutOfSlot) {
                        row.style.cssText = 'background-color:#e6b3ff;font-weight:bold;transition:background 0.3s';
                        row.setAttribute('data-highlighted', 'purple');
                        highlightedCount++;
                        if (!seenOutOfSlot.has(agentLoginLower)) {
                            seenOutOfSlot.add(agentLoginLower);
                            const sd = state.breakSchedules[agentLoginLower];
                            if (!state.violationDetectionTimes[agentLoginLower]) {
                                state.violationDetectionTimes[agentLoginLower] = new Date();
                            }
                            state.outOfSlotAgents.push({
                                login: agentLogin, manager: sd?.manager || 'N/A', activity,
                                duration: durationText || 'N/A', oosTime: getOutOfSlotTime(agentLoginLower),
                                break10: sd?.break10 || 'N/A', break20: sd?.break20 || 'N/A', break30: sd?.break30 || 'N/A'
                            });
                            // Persist in violation log (deduplicated per slot per day)
                            const _pvSlot = getSlotForViolation(agentLoginLower);
                            if (_pvSlot) {
                                const _pvHash = buildViolationHash(agentLoginLower, _pvSlot);
                                if (!state.violationLog.some(v => v.hash === _pvHash)) {
                                    state.violationLog.push({
                                        hash: _pvHash, login: agentLoginLower,
                                        manager: sd?.manager||'N/A', activity,
                                        duration: durationText||'N/A', oosTime: getOutOfSlotTime(agentLoginLower),
                                        slotStart: _pvSlot.start, slotEnd: _pvSlot.end,
                                        break10: sd?.break10||'N/A', break20: sd?.break20||'N/A', break30: sd?.break30||'N/A',
                                        detectedAt: new Date().toISOString(), uploaded: false
                                    });
                                    saveViolationLog();
                                }
                            }
                        }
                        const dm = parseTimeToMinutes(durationText);
                        if (!activityDetails[activity]) {
                            activityDetails[activity] = { count: 0, maxDuration: durationText, maxDurationMinutes: 0, maxAgent: agentLogin };
                        }
                        activityDetails[activity].count++;
                        if (dm > activityDetails[activity].maxDurationMinutes) {
                            activityDetails[activity].maxDuration = durationText;
                            activityDetails[activity].maxDurationMinutes = dm;
                            activityDetails[activity].maxAgent = agentLogin;
                        }
                        continue;
                    }
                }

                // Activity tracking
                if (!managerLoginsSet.has(activityLower)) {
                    const cleanDur   = durationText || '-';
                    const durMinutes = parseTimeToMinutes(cleanDur);
                    if (!activityDetails[activity]) {
                        activityDetails[activity] = { count: 0, maxDuration: cleanDur, maxDurationMinutes: durMinutes, maxAgent: agentLogin };
                    }
                    activityDetails[activity].count++;
                    if (durMinutes > activityDetails[activity].maxDurationMinutes) {
                        activityDetails[activity].maxDuration        = cleanDur;
                        activityDetails[activity].maxDurationMinutes = durMinutes;
                        activityDetails[activity].maxAgent           = agentLogin;
                    }
                    if (!agentsByActivity[activity]) agentsByActivity[activity] = [];
                    agentsByActivity[activity].push({
                        login: agentLogin, manager: getManagerForAgent(agentLogin, cells),
                        activity, duration: cleanDur, durationMinutes: durMinutes,
                        atrStatus: atrStatusMap[agentLoginLower] || 'Unknown'
                    });
                }

                // Duration-based highlights
                const duration = parseTimeToMinutes(durationText);
                if (activityLower === 'available') {
                    if (state.settings.redEnabled && duration >= redMin) {
                        row.style.cssText = 'background-color:#ffcccc;font-weight:bold';
                        row.setAttribute('data-highlighted', 'red'); highlightedCount++;
                    } else if (state.settings.yellowEnabled && duration >= yellowMin && duration < yellowMax) {
                        row.style.cssText = 'background-color:#ffff99;font-weight:bold';
                        row.setAttribute('data-highlighted', 'yellow'); highlightedCount++;
                    }
                } else if (activityLower === 'break' && state.settings.blueEnabled && duration > breakMin) {
                    row.style.cssText = 'background-color:#cce5ff;font-weight:bold';
                    row.setAttribute('data-highlighted', 'blue'); highlightedCount++;
                    if (state.isScheduleMonitoring && state.breakSchedules[agentLoginLower]) {
                        const _bvSlot = getSlotForViolation(agentLoginLower);
                        if (_bvSlot) {
                            const _bvHash = buildViolationHash(agentLoginLower, _bvSlot);
                            if (!state.violationLog.some(v => v.hash === _bvHash)) {
                                const _bvSd = state.breakSchedules[agentLoginLower];
                                state.violationLog.push({
                                    hash: _bvHash, login: agentLoginLower,
                                    manager: _bvSd?.manager||'N/A', activity,
                                    duration: durationText||'N/A', oosTime: getOutOfSlotTime(agentLoginLower),
                                    slotStart: _bvSlot.start, slotEnd: _bvSlot.end,
                                    break10: _bvSd?.break10||'N/A', break20: _bvSd?.break20||'N/A', break30: _bvSd?.break30||'N/A',
                                    detectedAt: new Date().toISOString(), uploaded: false
                                });
                                    saveViolationLog();
                            }
                        }
                    }
                } else if (activityLower === 'lunch' && state.settings.orangeEnabled && duration > lunchMin) {
                    row.style.cssText = 'background-color:#ffe5cc;font-weight:bold';
                    row.setAttribute('data-highlighted', 'orange'); highlightedCount++;
                    if (state.isScheduleMonitoring && state.breakSchedules[agentLoginLower]) {
                        const _ovSlot = getSlotForViolation(agentLoginLower);
                        if (_ovSlot) {
                            const _ovHash = buildViolationHash(agentLoginLower, _ovSlot);
                            if (!state.violationLog.some(v => v.hash === _ovHash)) {
                                const _ovSd = state.breakSchedules[agentLoginLower];
                                state.violationLog.push({
                                    hash: _ovHash, login: agentLoginLower,
                                    manager: _ovSd?.manager||'N/A', activity,
                                    duration: durationText||'N/A', oosTime: getOutOfSlotTime(agentLoginLower),
                                    slotStart: _ovSlot.start, slotEnd: _ovSlot.end,
                                    break10: _ovSd?.break10||'N/A', break20: _ovSd?.break20||'N/A', break30: _ovSd?.break30||'N/A',
                                    detectedAt: new Date().toISOString(), uploaded: false
                                });
                                    saveViolationLog();
                            }
                        }
                    }
                }

                // Green: agent is in scheduled break slot but not yet on Break/Lunch
                if (!state.isPaused && state.isScheduleMonitoring && state.lateToBreakAlertMin > 0 &&
                    activityLower !== 'break' && activityLower !== 'lunch') {
                    const _ltbSd = state.breakSchedules[agentLoginLower];
                    if (_ltbSd?.breaks?.length) {
                        const _ltbCur = new Date().getHours()*60 + new Date().getMinutes() + new Date().getSeconds()/60;
                        for (const _ltbSlot of _ltbSd.breaks) {
                            if (_ltbCur >= _ltbSlot.start + state.lateToBreakAlertMin && _ltbCur <= _ltbSlot.end) {
                                if (!row.getAttribute('data-highlighted')) {
                                    row.style.cssText = 'background-color:#ccffcc;font-weight:bold';
                                    row.setAttribute('data-highlighted', 'green');
                                    highlightedCount++;
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }

        // Clean stale violation detection times
        const currentViolations = new Set(state.outOfSlotAgents.map(a => a.login.toLowerCase()));
        for (const login of Object.keys(state.violationDetectionTimes)) {
            if (!currentViolations.has(login)) delete state.violationDetectionTimes[login];
        }

        // Sort by duration desc
        for (const act in agentsByActivity) {
            agentsByActivity[act].sort((a, b) => b.durationMinutes - a.durationMinutes);
        }

        // ── Virtual activity: NPT-(Breaks+Lunch+Project) ──────────────────────
        // NPT = everything except Available / On Contact / After Contact Work /
        //       Monitoring / Monitoring Ended.
        // Subtract the scheduled/unavoidable auxes (Break, Lunch, Project)
        // to produce an aggregate of the remaining unplanned NPT time.
        const NPT_LABEL    = 'NPT-(Breaks+Lunch+Project)';
        const NPT_EXCLUDED = new Set([
            'available', 'on contact', 'after contact work',
            'monitoring', 'monitoring ended',
            'break', 'lunch', 'project'
        ]);
        const nptAgents = [];
        let nptMaxDurMin = 0, nptMaxDur = '-', nptMaxAgent = '-';
        for (const [act, agents] of Object.entries(agentsByActivity)) {
            if (!NPT_EXCLUDED.has(act.toLowerCase())) {
                for (const agent of agents) {
                    nptAgents.push(agent);
                    if (agent.durationMinutes > nptMaxDurMin) {
                        nptMaxDurMin = agent.durationMinutes;
                        nptMaxDur    = agent.duration;
                        nptMaxAgent  = agent.login;
                    }
                }
            }
        }
        if (nptAgents.length > 0) {
            nptAgents.sort((a, b) => b.durationMinutes - a.durationMinutes);
            activityDetails[NPT_LABEL]  = { count: nptAgents.length, maxDuration: nptMaxDur, maxDurationMinutes: nptMaxDurMin, maxAgent: nptMaxAgent };
            agentsByActivity[NPT_LABEL] = nptAgents;
        }

        state.cachedActivityDetails  = activityDetails;
        state.cachedAgentsByActivity = agentsByActivity;

        // ✅ Only refresh the active tab — no full re-renders on inactive tabs
        refreshActiveTab();

        const total = Object.values(activityDetails).reduce((s, d) => s + d.count, 0);
        const dbg   = document.getElementById('cv-debugInfo');
        if (dbg) dbg.textContent = `#${state.checkCount} | ${total} agents | ${highlightedCount} highlighted | ${seenOutOfSlot.size} out-of-slot | ${state.isPaused ? '⏸ PAUSED' : '▶ LIVE'}`;
    };

    const debouncedUpdate = () => {
        clearTimeout(state.updateTimeout);
        state.updateTimeout = setTimeout(unifiedUpdate, CONFIG.DEBOUNCE);
    };

    const attachObserver = () => {
        if (state.observer) state.observer.disconnect();
        state.observer = new MutationObserver(debouncedUpdate);
        const tables = getTables();
        tables.forEach(t => state.observer.observe(t, { childList: true, subtree: false }));
        state.lastTableCount = tables.length;
    };
    // ==================== DASHBOARD SHELL ====================

    const TABS = [
        { id: 'activity',  label: '📊 Activity'  },
        { id: 'outofslot', label: '🚨 Out-of-Slot'},
        { id: 'schedule',  label: '🟣 Schedule'   },
        { id: 'settings',  label: '⚙️ Settings'   }
    ];

    const createDashboard = () => {
        const el = document.createElement('div');
        el.id    = 'cv-dashboard';
        el.style.cssText = `
            position:fixed; top:60px; left:10px;
            width:800px; height:560px;
            background:#fff; border:2px solid #232f3e;
            border-radius:10px; z-index:9999;
            box-shadow:0 6px 24px rgba(0,0,0,0.22);
            font-family:Arial,sans-serif;
            display:flex; flex-direction:column;
            overflow:hidden; cursor:default;
        `;
        el.innerHTML = `
            <div id="cv-titlebar" style="background:#232f3e;color:#fff;padding:8px 14px;display:flex;justify-content:space-between;align-items:center;cursor:move;flex-shrink:0;border-radius:8px 8px 0 0;user-select:none">
                <span style="font-weight:bold;font-size:15px">🖥️ ConnectVision <span style="font-size:11px;opacity:.6;margin-left:6px">v8.0</span></span>
                <div style="display:flex;gap:8px;align-items:center">
                    <span id="cv-pauseIndicator" style="display:none;background:#d13212;color:#fff;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:bold">⏸ PAUSED</span>
                    <span id="cv-tabOwner"        style="background:#067d62;color:#fff;padding:2px 10px;border-radius:12px;font-size:11px">● Active</span>
                    <button id="cv-minimizeBtn"  style="background:#ff9900;border:none;color:#fff;border-radius:4px;padding:1px 10px;cursor:pointer;font-size:16px;line-height:1.4">−</button>
                    <button id="cv-maximizeBtn" style="background:#067d62;border:none;color:#fff;border-radius:4px;padding:1px 10px;cursor:pointer;font-size:14px;line-height:1.4" title="Maximize">□</button>
                    <button id="cv-closeBtn"    style="background:#d13212;border:none;color:#fff;border-radius:4px;padding:1px 10px;cursor:pointer;font-size:16px;line-height:1.4">×</button>
                </div>
            </div>
            <div id="cv-tabbar" style="display:flex;background:#f5f5f5;border-bottom:2px solid #232f3e;flex-shrink:0">
                ${TABS.map(t => `<button class="cv-tab" data-tab="${t.id}" style="flex:1;padding:9px 4px;border:none;background:transparent;cursor:pointer;font-size:13px;font-weight:bold;color:#555;border-bottom:3px solid transparent">${t.label}</button>`).join('')}
            </div>
            <div id="cv-tabContent" style="flex:1;overflow:hidden;display:flex;flex-direction:column;min-height:0"></div>
            <div style="background:#f8f8f8;border-top:1px solid #ddd;padding:3px 12px;font-size:11px;color:#888;flex-shrink:0;border-radius:0 0 8px 8px">
                <span id="cv-debugInfo">Initializing...</span>
            </div>
        `;
        document.body.appendChild(el);
        state.dashboardEl = el;

        makeDraggable(el);
        makeResizable(el);
        wireTabBar();

        document.getElementById('cv-minimizeBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            const content = document.getElementById('cv-tabContent');
            const tabbar  = document.getElementById('cv-tabbar');
            const btn     = document.getElementById('cv-minimizeBtn');
            const isMin   = content.style.display === 'none';
            content.style.display = isMin ? 'flex'  : 'none';
            tabbar.style.display  = isMin ? 'flex'  : 'none';
            btn.textContent       = isMin ? '−'     : '+';
            el.style.height       = isMin ? '560px' : 'auto';
        });

        document.getElementById('cv-closeBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            el.style.display = 'none';
        });

        // Maximize / Restore
        let _cvMaximized = false;
        let _cvPrevBounds = null;
        document.getElementById('cv-maximizeBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            const btn = document.getElementById('cv-maximizeBtn');
            if (!_cvMaximized) {
                _cvPrevBounds = { top: el.style.top, left: el.style.left, width: el.style.width, height: el.style.height };
                el.style.top = '0'; el.style.left = '0';
                el.style.width = '100vw'; el.style.height = '100vh';
                el.style.borderRadius = '0';
                btn.textContent = '⧉'; btn.title = 'Restore';
                _cvMaximized = true;
            } else {
                el.style.top = _cvPrevBounds.top; el.style.left = _cvPrevBounds.left;
                el.style.width = _cvPrevBounds.width || '800px'; el.style.height = _cvPrevBounds.height || '560px';
                el.style.borderRadius = '10px';
                btn.textContent = '□'; btn.title = 'Maximize';
                _cvMaximized = false;
            }
        });

        // Blink animation for pause badge
        if (!document.getElementById('cv-styles')) {
            const s = document.createElement('style');
            s.id    = 'cv-styles';
            s.textContent = `@keyframes cv-blink{50%{opacity:0}} #cv-pauseIndicator{animation:cv-blink 1s step-start infinite}`;
            document.head.appendChild(s);
        }

        switchTab(state.activeTab);
    };

    const wireTabBar = () => {
        document.querySelectorAll('.cv-tab').forEach(btn => {
            btn.addEventListener('click', () => switchTab(btn.getAttribute('data-tab')));
        });
    };

    const switchTab = (tabId) => {
        state.activeTab = tabId;
        document.querySelectorAll('.cv-tab').forEach(btn => {
            const active = btn.getAttribute('data-tab') === tabId;
            btn.style.background   = active ? '#232f3e'           : 'transparent';
            btn.style.color        = active ? '#ff9900'           : '#555';
            btn.style.borderBottom = active ? '3px solid #ff9900' : '3px solid transparent';
        });
        const content = document.getElementById('cv-tabContent');
        if (!content) return;
        content.innerHTML = '';
        switch (tabId) {
            case 'activity':  content.appendChild(buildActivityTab());  wireActivityTab();  break;
            case 'outofslot': content.appendChild(buildOutOfSlotTab()); wireOutOfSlotTab(); break;
            case 'schedule':  content.appendChild(buildScheduleTab());  wireScheduleTab();  break;
            case 'settings':  content.appendChild(buildSettingsTab());  wireSettingsTab();  break;
        }
        // ✅ Populate immediately after building — no waiting for next 5s cycle
        refreshActiveTab();
    };

    const refreshActiveTab = () => {
        const pi = document.getElementById('cv-pauseIndicator');
        if (pi) pi.style.display = state.isPaused ? 'inline-block' : 'none';
        switch (state.activeTab) {
            case 'activity':  renderActivityContent();  break;
            case 'outofslot': renderOutOfSlotContent(); break;
            case 'schedule':  renderScheduleStatus();   break;
        }
    };

    // ── DRAGGABLE & RESIZABLE ──────────────────────────────────────

    const makeDraggable = (el) => {
        let p1=0,p2=0,p3=0,p4=0;
        el.onmousedown = (e) => {
            if (!e.target.closest('#cv-titlebar')) return;
            if (e.target.closest('button')) return;
            e.preventDefault();
            p3=e.clientX; p4=e.clientY;
            document.onmouseup   = () => { document.onmouseup=null; document.onmousemove=null; };
            document.onmousemove = (e) => {
                e.preventDefault();
                p1=p3-e.clientX; p2=p4-e.clientY; p3=e.clientX; p4=e.clientY;
                el.style.top  = Math.max(0, el.offsetTop  - p2) + 'px';
                el.style.left = Math.max(0, el.offsetLeft - p1) + 'px';
                el.style.right='auto'; el.style.bottom='auto';
            };
        };
    };

    const makeResizable = (el) => {
        const r = document.createElement('div');
        r.style.cssText = 'position:absolute;right:0;bottom:0;width:18px;height:18px;cursor:nwse-resize;z-index:10;background:linear-gradient(135deg,transparent 50%,#ff9900 50%);border-radius:0 0 8px 0';
        el.appendChild(r);
        let sx,sy,sw,sh;
        r.onmousedown = (e) => {
            e.preventDefault(); e.stopPropagation();
            sx=e.clientX; sy=e.clientY;
            sw=parseInt(getComputedStyle(el).width);
            sh=parseInt(getComputedStyle(el).height);
            document.onmousemove = (e) => {
                el.style.width  = Math.max(600, Math.min(window.innerWidth-20,  sw+e.clientX-sx)) + 'px';
                el.style.height = Math.max(400, Math.min(window.innerHeight-20, sh+e.clientY-sy)) + 'px';
                el.style.maxHeight = 'none';
            };
            document.onmouseup = () => { document.onmousemove=null; document.onmouseup=null; };
        };
    };

    // ════════════════════════════════════════════
    // TAB: ACTIVITY
    // ════════════════════════════════════════════

    const buildActivityTab = () => {
        const w = document.createElement('div');
        w.style.cssText = 'display:flex;flex-direction:column;height:100%;padding:10px;gap:8px;overflow:hidden';
        w.innerHTML = `
            <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;flex-shrink:0">
                <div>
                    <label style="font-weight:bold;font-size:13px;margin-right:6px">Activity:</label>
                    <select id="cv-actDropdown" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;min-width:130px">
                        <option value="">-- Select --</option>
                    </select>
                </div>
                <div>
                    <label style="font-weight:bold;font-size:13px;margin-right:6px">Manager:</label>
                    <select id="cv-mgrDropdown" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px;min-width:130px">
                        <option value="">All Managers</option>
                    </select>
                </div>
                <div>
                    <label style="font-weight:bold;font-size:13px;margin-right:6px">Min (min):</label>
                    <input type="number" id="cv-threshold" min="0" max="999" value="0" style="width:60px;padding:4px 6px;border:1px solid #ccc;border-radius:4px;font-size:13px;text-align:center">
                </div>
                <div>
                    <label style="font-weight:bold;font-size:13px;margin-right:6px">Show:</label>
                    <select id="cv-limitDropdown" style="padding:4px 8px;border:1px solid #ccc;border-radius:4px;font-size:13px">
                        <option value="5">Top 5</option>
                        <option value="10" selected>Top 10</option>
                        <option value="20">Top 20</option>
                        <option value="all">All</option>
                    </select>
                </div>
                <button id="cv-actCSVBtn" style="padding:4px 12px;background:#232f3e;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:bold;margin-left:auto">📥 CSV</button>
            </div>
            <!-- Activity chip strip -->
            <div id="cv-actSummaryStrip" style="display:flex;gap:5px;flex-wrap:wrap;flex-shrink:0;min-height:24px"></div>
            <!-- Agent table -->
            <div style="flex:1;overflow-y:auto;overflow-x:auto;border:1px solid #eee;border-radius:6px;min-height:0">
                <table id="cv-actTable" style="width:100%;border-collapse:collapse;font-size:13px">
                    <thead>
                        <tr style="background:#232f3e;color:#fff;position:sticky;top:0;z-index:2">
                            <th style="padding:7px 8px;text-align:center;width:36px">#</th>
                            <th id="cv-sortLogin"    style="padding:7px 8px;text-align:left;cursor:pointer;user-select:none">Login <span id="cv-iconLogin">⇅</span></th>
                            <th id="cv-sortManager"  style="padding:7px 8px;text-align:left;cursor:pointer;user-select:none">Manager <span id="cv-iconManager">⇅</span></th>
                            <th id="cv-sortActivity" style="padding:7px 8px;text-align:left;cursor:pointer;user-select:none">Aux <span id="cv-iconActivity">⇅</span></th>
                            <th id="cv-sortDuration" style="padding:7px 8px;text-align:center;cursor:pointer;user-select:none">Duration <span id="cv-iconDuration">▼</span></th>
                            <th id="cv-sortATR"      style="padding:7px 8px;text-align:center;cursor:pointer;user-select:none">ATR <span id="cv-iconATR">⇅</span></th>
                        </tr>
                    </thead>
                    <tbody id="cv-actTableBody">
                        <tr><td colspan="6" style="padding:20px;text-align:center;color:#999">Select an activity above</td></tr>
                    </tbody>
                </table>
            </div>
            <div style="text-align:right;font-size:12px;color:#888;flex-shrink:0">
                Showing <span id="cv-actShowing">0</span> of <span id="cv-actTotal">0</span> agents
            </div>
        `;
        return w;
    };

    const wireActivityTab = () => {
        const SORT_MAP = {
            'cv-sortLogin':'login','cv-sortManager':'manager',
            'cv-sortActivity':'activity','cv-sortDuration':'duration','cv-sortATR':'atrStatus'
        };
        for (const [id, col] of Object.entries(SORT_MAP)) {
            document.getElementById(id)?.addEventListener('click', () => {
                state.activityDetailsSortColumn    = col;
                state.activityDetailsSortDirection = (state.activityDetailsSortColumn === col && state.activityDetailsSortDirection === 'desc') ? 'asc' : (col === 'duration' ? 'desc' : 'asc');
                updateSortIcons();
                renderActivityTable();
            });
        }
        document.getElementById('cv-actDropdown')?.addEventListener('change',  () => renderActivityTable());
        document.getElementById('cv-mgrDropdown')?.addEventListener('change',  () => renderActivityTable());
        document.getElementById('cv-limitDropdown')?.addEventListener('change',() => renderActivityTable());
        document.getElementById('cv-threshold')?.addEventListener('input',     () => renderActivityTable());
        document.getElementById('cv-actCSVBtn')?.addEventListener('click',     downloadActivityCSV);
    };

    const updateSortIcons = () => {
        const ICON_MAP = {
            'cv-sortLogin':'cv-iconLogin','cv-sortManager':'cv-iconManager',
            'cv-sortActivity':'cv-iconActivity','cv-sortDuration':'cv-iconDuration','cv-sortATR':'cv-iconATR'
        };
        const COL_MAP = {
            'cv-sortLogin':'login','cv-sortManager':'manager',
            'cv-sortActivity':'activity','cv-sortDuration':'duration','cv-sortATR':'atrStatus'
        };
        for (const [btnId, iconId] of Object.entries(ICON_MAP)) {
            const icon = document.getElementById(iconId);
            if (!icon) continue;
            icon.textContent = COL_MAP[btnId] === state.activityDetailsSortColumn
                ? (state.activityDetailsSortDirection === 'asc' ? '▲' : '▼') : '⇅';
        }
    };

    // ✅ FIX: renderActivityContent populates dropdown AND table immediately
    const renderActivityContent = () => {
        const dropdown = document.getElementById('cv-actDropdown');
        if (!dropdown) return;

        const activities  = Object.keys(state.cachedAgentsByActivity || {}).filter(a => a?.trim()).sort();
        const currentSel  = dropdown.value;
        const currentOpts = Array.from(dropdown.options).slice(2).map(o => o.value);

        // Only rebuild dropdown options if the list itself changed
        if (JSON.stringify(currentOpts) !== JSON.stringify(activities)) {
            dropdown.innerHTML = '<option value="">-- Select --</option><option value="All">All</option>';
            activities.forEach(a => {
                const opt = document.createElement('option');
                opt.value = opt.textContent = a;
                dropdown.appendChild(opt);
            });
        }
        // Always restore selection — never let a background update wipe the user's choice
        if (currentSel && (currentSel === 'All' || activities.includes(currentSel))) {
            dropdown.value = currentSel;
        } else if (!dropdown.value && activities.includes('Available')) {
            dropdown.value = 'Available';
        }

        // Rebuild chip strip
        const strip = document.getElementById('cv-actSummaryStrip');
        if (strip) {
            const details = state.cachedActivityDetails;
            strip.innerHTML = Object.keys(details).sort().map(act => {
                const d   = details[act];
                const col = act.toLowerCase() === 'available' ? '#067d62'
                          : act.toLowerCase() === 'break'     ? '#2196F3'
                          : act.toLowerCase() === 'lunch'     ? '#ff9800' : '#555';
                return `<span style="background:#f0f0f0;border:1px solid #ddd;border-radius:12px;padding:2px 10px;font-size:12px;cursor:pointer;color:${col};font-weight:bold" data-act="${sanitize(act)}">${sanitize(act)}: <b>${d.count}</b></span>`;
            }).join('');
            strip.querySelectorAll('[data-act]').forEach(chip => {
                chip.addEventListener('click', () => {
                    dropdown.value = chip.getAttribute('data-act');
                    renderActivityTable();
                });
            });
        }

        // ✅ Always render the table immediately — no conditional skipping
        renderActivityTable();
    };

    const renderActivityTable = () => {
        const dropdown = document.getElementById('cv-actDropdown');
        const mgrDrop  = document.getElementById('cv-mgrDropdown');
        const limDrop  = document.getElementById('cv-limitDropdown');
        const tbody    = document.getElementById('cv-actTableBody');
        const showEl   = document.getElementById('cv-actShowing');
        const totEl    = document.getElementById('cv-actTotal');
        if (!dropdown || !tbody) return;

        const selAct = dropdown.value;
        const selMgr = mgrDrop?.value || '';
        const limit  = limDrop?.value || '10';

        if (!selAct) {
            tbody.innerHTML = '<tr><td colspan="6" style="padding:20px;text-align:center;color:#999">Select an activity above</td></tr>';
            if (showEl) showEl.textContent = '0';
            if (totEl)  totEl.textContent  = '0';
            return;
        }

        let agents = selAct === 'All'
            ? Object.values(state.cachedAgentsByActivity || {}).flat()
            : (state.cachedAgentsByActivity?.[selAct] || []);

        if (agents.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" style="padding:20px;text-align:center;color:#999">No agents found</td></tr>';
            if (showEl) showEl.textContent = '0';
            if (totEl)  totEl.textContent  = '0';
            return;
        }

        // Update manager filter only if changed
        if (mgrDrop) {
            const curMgr  = mgrDrop.value;
            const mgrs    = [...new Set(agents.map(a => a.manager).filter(m => m?.trim() && m !== 'N/A'))].sort();
            const existing = Array.from(mgrDrop.options).slice(1).map(o => o.value).sort();
            if (JSON.stringify(existing) !== JSON.stringify(mgrs)) {
                mgrDrop.innerHTML = '<option value="">All Managers</option>';
                mgrs.forEach(m => { const o = document.createElement('option'); o.value = o.textContent = m; mgrDrop.appendChild(o); });
                if (mgrs.includes(curMgr)) mgrDrop.value = curMgr;
            }
        }

        const threshold = parseFloat(document.getElementById('cv-threshold')?.value) || 0;
        let filtered    = agents.filter(a => a.durationMinutes >= threshold);
        if (selMgr)     filtered = filtered.filter(a => a.manager === selMgr);

        const col = state.activityDetailsSortColumn || 'duration';
        const dir = state.activityDetailsSortDirection || 'desc';
        filtered.sort((a, b) => {
            let ca, cb;
            switch (col) {
                case 'duration':  ca=a.durationMinutes||0;            cb=b.durationMinutes||0;            break;
                case 'login':     ca=(a.login||'').toLowerCase();     cb=(b.login||'').toLowerCase();     break;
                case 'manager':   ca=(a.manager||'').toLowerCase();   cb=(b.manager||'').toLowerCase();   break;
                case 'activity':  ca=(a.activity||'').toLowerCase();  cb=(b.activity||'').toLowerCase();  break;
                case 'atrStatus': ca=(a.atrStatus||'').toLowerCase(); cb=(b.atrStatus||'').toLowerCase(); break;
                default:          ca=a.durationMinutes||0;            cb=b.durationMinutes||0;
            }
            if (col === 'duration') return dir === 'asc' ? ca-cb : cb-ca;
            return dir === 'asc' ? (ca<cb?-1:ca>cb?1:0) : (ca>cb?-1:ca<cb?1:0);
        });

        const displayed = limit === 'all' ? filtered : filtered.slice(0, parseInt(limit));

        if (displayed.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" style="padding:20px;text-align:center;color:#999">No agents match filters</td></tr>';
            if (showEl) showEl.textContent = '0';
            if (totEl)  totEl.textContent  = String(filtered.length);
            return;
        }

        const frag = document.createDocumentFragment();
        displayed.forEach((agent, idx) => {
            const tr     = document.createElement('tr');
            tr.style.background = idx % 2 === 0 ? '#fff' : '#f9f9f9';
            const atrCol = agent.atrStatus === 'Offline' ? '#d13212' : agent.atrStatus === 'Online' ? '#067d62' : '#888';
            const atrWt  = agent.atrStatus === 'Offline' ? 'bold' : 'normal';
            tr.innerHTML = `
                <td style="padding:6px 8px;text-align:center;color:#aaa">${idx+1}</td>
                <td class="cv-copyLogin" data-login="${sanitize(agent.login)}" style="padding:6px 8px;cursor:pointer;color:#0066c0;font-weight:bold" title="Click to copy">${sanitize(agent.login)}</td>
                <td style="padding:6px 8px">${sanitize(agent.manager||'-')}</td>
                <td style="padding:6px 8px">${sanitize(agent.activity||'-')}</td>
                <td style="padding:6px 8px;text-align:center;font-weight:bold">${sanitize(agent.duration||'-')}</td>
                <td style="padding:6px 8px;text-align:center;color:${atrCol};font-weight:${atrWt}">${sanitize(agent.atrStatus||'-')}</td>`;
            frag.appendChild(tr);
        });
        tbody.replaceChildren(frag);

        tbody.querySelectorAll('.cv-copyLogin').forEach(cell => {
            cell.addEventListener('click', () => {
                navigator.clipboard.writeText(cell.getAttribute('data-login')).then(() => {
                    cell.style.background = '#c8f7c5';
                    setTimeout(() => { cell.style.background = ''; }, 600);
                });
            });
        });

        if (showEl) showEl.textContent = String(displayed.length);
        if (totEl)  totEl.textContent  = String(filtered.length);
    };

    const downloadActivityCSV = () => {
        const details = state.cachedActivityDetails;
        const acts    = Object.keys(details).sort();
        let csv = `Activity,HC,Highest Duration,Agent${CRLF}`;
        let tot = 0;
        for (const act of acts) { const d=details[act]; tot+=d.count; csv+=`${act},${d.count},${d.maxDuration},${d.maxAgent}${CRLF}`; }
        csv += `Total,${tot},,${CRLF}`;
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `Activity_${getISTTimestamp()}.csv` });
        link.style.display = 'none';
        document.body.appendChild(link); link.click(); document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
    };


    // ════════════════════════════════════════════
    // SHAREPOINT INTEGRATION
    // ════════════════════════════════════════════

    function getRequestDigest(callback) {
        console.log('[ConnectVision SP] Requesting digest from:', `${CONFIG.SP_SITE}/_api/contextinfo`);
        GM_xmlhttpRequest({
            method:  'POST',
            url:     `${CONFIG.SP_SITE}/_api/contextinfo`,
            headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' },
            onload:  (res) => {
                console.log('[ConnectVision SP] Digest response:', res.status, res.responseText?.slice(0, 200));
                try {
                    const d = JSON.parse(res.responseText);
                    callback(null, d.d.GetContextWebInformation.FormDigestValue);
                } catch (e) { console.error('[ConnectVision SP] Digest parse error:', e); callback(e); }
            },
            onerror: (e) => { console.error('[ConnectVision SP] Digest network error:', e); callback(new Error('contextinfo request failed')); }
        });
    }

    function fetchExistingHashesFromSP(dateStr, callback) {
        const filter = encodeURIComponent(`Date eq '${dateStr}'`);
        GM_xmlhttpRequest({
            method:  'GET',
            url:     `${CONFIG.SP_SITE}/_api/web/lists/getbytitle('${CONFIG.SP_LIST}')/items?$select=ViolationHash&$filter=${filter}&$top=500`,
            headers: { 'Accept': 'application/json;odata=verbose' },
            onload:  (res) => {
                try {
                    const d = JSON.parse(res.responseText);
                    const hashes = new Set((d.d?.results || []).map(i => i.ViolationHash).filter(Boolean));
                    callback(null, hashes);
                } catch (e) { callback(e); }
            },
            onerror: () => callback(new Error('SP fetch failed'))
        });
    }

    // ── SharePoint Duplicate Cleanup (runs every 30 min) ──────────────────
    function cleanupDuplicatesFromSP() {
        const today = new Date().toISOString().slice(0, 10);
        const filter = encodeURIComponent(`Date eq '${today}'`);
        console.log('[ConnectVision SP Cleanup] Starting duplicate scan...');

        GM_xmlhttpRequest({
            method: 'GET',
            url: `${CONFIG.SP_SITE}/_api/web/lists/getbytitle('${CONFIG.SP_LIST}')/items?$select=Id,ViolationHash&$filter=${filter}&$top=1000`,
            headers: { 'Accept': 'application/json;odata=verbose' },
            onload: (res) => {
                try {
                    const data = JSON.parse(res.responseText);
                    const items = data.d?.results || [];
                    if (items.length === 0) { console.log('[ConnectVision SP Cleanup] No entries today'); return; }

                    // Group by ViolationHash — keep lowest ID (first uploaded), delete rest
                    const groups = {};
                    items.forEach(item => {
                        const hash = item.ViolationHash;
                        if (!hash) return;
                        if (!groups[hash]) groups[hash] = [];
                        groups[hash].push(item.Id);
                    });

                    const toDelete = [];
                    for (const hash in groups) {
                        if (groups[hash].length > 1) {
                            groups[hash].sort((a, b) => a - b); // lowest ID first
                            toDelete.push(...groups[hash].slice(1)); // delete all except first
                        }
                    }

                    if (toDelete.length === 0) {
                        console.log('[ConnectVision SP Cleanup] No duplicates found (${items.length} entries checked)');
                        return;
                    }

                    console.log(`[ConnectVision SP Cleanup] Found ${toDelete.length} duplicate(s) to remove`);

                    // Get digest then delete
                    getRequestDigest((err, digest) => {
                        if (err) { console.error('[ConnectVision SP Cleanup] Auth error:', err); return; }
                        let idx = 0;
                        const deleteNext = () => {
                            if (idx >= toDelete.length) {
                                console.log(`[ConnectVision SP Cleanup] Removed ${toDelete.length} duplicate(s)`);
                                return;
                            }
                            const id = toDelete[idx++];
                            GM_xmlhttpRequest({
                                method: 'POST',
                                url: `${CONFIG.SP_SITE}/_api/web/lists/getbytitle('${CONFIG.SP_LIST}')/items(${id})`,
                                headers: {
                                    'Accept': 'application/json;odata=verbose',
                                    'X-RequestDigest': digest,
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'DELETE'
                                },
                                onload: (r) => {
                                    if (r.status >= 200 && r.status < 300) {
                                        console.log(`[ConnectVision SP Cleanup] Deleted item ID ${id}`);
                                    } else {
                                        console.error(`[ConnectVision SP Cleanup] Failed to delete ID ${id}: HTTP ${r.status}`);
                                    }
                                    setTimeout(deleteNext, 500); // throttle to avoid rate limiting
                                },
                                onerror: () => {
                                    console.error(`[ConnectVision SP Cleanup] Network error deleting ID ${id}`);
                                    setTimeout(deleteNext, 500);
                                }
                            });
                        };
                        deleteNext();
                    });
                } catch (e) { console.error('[ConnectVision SP Cleanup] Parse error:', e); }
            },
            onerror: () => console.error('[ConnectVision SP Cleanup] Network error fetching list')
        });
    }

    // Run cleanup every 30 minutes
    setInterval(cleanupDuplicatesFromSP, 30 * 60 * 1000);
    // Also run once 2 minutes after boot (give time for initial uploads)
    setTimeout(cleanupDuplicatesFromSP, 2 * 60 * 1000);

    function postViolationToSP(item, digest, callback) {
        const payload = JSON.stringify({
            __metadata:      { type: 'SP.Data.OutOfSlot_x005f_Breaks_x005f_LogListItem' },
            Login:           item.login,
            Manager:         item.manager,
            Activity:        item.activity,
            Duration:        item.duration,
            Break10Mins:     item.break10,
            Break20Mins:     item.break20,
            Break30Mins:     item.break30,
            BufferTime:      `${String(state.bufferMinutes).padStart(2,'0')}:${String(state.bufferSeconds).padStart(2,'0')}`,
            CurrentTime:     new Date().toTimeString().slice(0, 8),
            CurrentSlot:     getCurrentSlotDisplay(),
            ViolationHash:   item.hash,
            OutOfSlotTime:   item.oosTime || 'N/A',
            ManagerComments: state.managerComments[item.login] || '',
            Date:            new Date().toISOString().slice(0, 10),
            UploadedBy:      state.uploaderAlias || getUploaderAlias()
        });
        GM_xmlhttpRequest({
            method:  'POST',
            url:     `${CONFIG.SP_SITE}/_api/web/lists/getbytitle('${CONFIG.SP_LIST}')/items`,
            headers: {
                'Accept':          'application/json;odata=verbose',
                'Content-Type':    'application/json;odata=verbose',
                'X-RequestDigest': digest
            },
            data:    payload,
            onload:  (res) => {
                if (res.status >= 200 && res.status < 300) { console.log('[ConnectVision SP] POST success for:', item.login); callback(null); }
                else { console.error('[ConnectVision SP] POST failed:', res.status, res.responseText?.slice(0, 300)); callback(new Error(`HTTP ${res.status}`)); }
            },
            onerror: (e) => { console.error('[ConnectVision SP] POST network error:', e); callback(new Error('POST failed')); }
        });
    }

    function uploadToSharePoint() {
        const btn = document.getElementById('cv-oos-uploadBtn');
        if (btn) { btn.disabled = true; btn.textContent = '\u23f3 Uploading\u2026'; }

        state.uploaderAlias = getUploaderAlias();
        const today   = new Date().toISOString().slice(0, 10);
        const pending = state.violationLog.filter(v => !v.uploaded);

        if (pending.length === 0) {
            showOOSMessage('\u2139\ufe0f No new violations to upload', 'info');
            if (btn) { btn.disabled = false; btn.textContent = '\u2b06\ufe0f Upload to SP'; }
            return;
        }

        showOOSMessage('\u23f3 Checking SharePoint for existing entries\u2026', 'info');

        fetchExistingHashesFromSP(today, (err, existingHashes) => {
            if (err) {
                showOOSMessage('\u274c SP fetch error: ' + err.message, 'error');
                if (btn) { btn.disabled = false; btn.textContent = '\u2b06\ufe0f Upload to SP'; }
                return;
            }

            pending.forEach(v => { if (existingHashes.has(v.hash)) v.uploaded = true; });
            const toUpload = pending.filter(v => !v.uploaded);

            if (toUpload.length === 0) {
                showOOSMessage('\u2705 All violations already in SharePoint', 'success');
                renderViolationLog();
                if (btn) { btn.disabled = false; btn.textContent = '\u2b06\ufe0f Upload to SP'; }
                return;
            }

            getRequestDigest((err, digest) => {
                if (err) {
                    showOOSMessage('\u274c Auth error: ' + err.message, 'error');
                    if (btn) { btn.disabled = false; btn.textContent = '\u2b06\ufe0f Upload to SP'; }
                    return;
                }

                showOOSMessage(`\u23f3 Uploading ${toUpload.length} violation(s)\u2026`, 'info');
                let idx = 0;
                const next = () => {
                    if (idx >= toUpload.length) {
                        showOOSMessage(`\u2705 Uploaded ${toUpload.length} violation(s)`, 'success');
                        renderViolationLog();
                        if (btn) { btn.disabled = false; btn.textContent = '\u2b06\ufe0f Upload to SP'; }
                        return;
                    }
                    const item = toUpload[idx++];
                    postViolationToSP(item, digest, (err) => {
                        if (err) {
                            showOOSMessage(`\u274c Error for ${item.login}: ${err.message}`, 'error');
                            if (btn) { btn.disabled = false; btn.textContent = '\u2b06\ufe0f Upload to SP'; }
                            return;
                        }
                        item.uploaded = true; saveViolationLog();
                        next();
                    });
                };
                next();
            });
        });
    }

    // ════════════════════════════════════════════
    // QUIP SCHEDULE LOADER
    // ════════════════════════════════════════════

    function loadScheduleFromQuip(onSuccess) {
        const token = GM_getValue('quipToken', '');
        if (!token) {
            showScheduleMessage('\u274c No Quip token set \u2014 add it in \u2699\ufe0f Settings', 'error');
            return;
        }
        const btn = document.getElementById('cv-quipLoadBtn');
        if (btn) { btn.disabled = true; btn.textContent = '\u23f3 Loading\u2026'; }
        showScheduleMessage('\u23f3 Fetching schedule from Quip\u2026', 'info');

        GM_xmlhttpRequest({
            method:  'GET',
            url:     `${CONFIG.QUIP_API_BASE}/threads/${CONFIG.QUIP_THREAD_ID}`,
            headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json' },
            onload:  (res) => {
                if (btn) { btn.disabled = false; btn.textContent = '\ud83d\udd17 Load from Quip'; }
                try {
                    const data = JSON.parse(res.responseText);
                    const html = data?.html || '';
                    if (!html) { showScheduleMessage('\u274c Empty Quip response', 'error'); return; }
                    parseScheduleFromQuipHTML(html);
                    if (typeof onSuccess === "function") onSuccess();
                } catch (e) { showScheduleMessage('\u274c Quip parse error: ' + e.message, 'error'); }
            },
            onerror: () => {
                if (btn) { btn.disabled = false; btn.textContent = '\ud83d\udd17 Load from Quip'; }
                showScheduleMessage('\u274c Network error fetching Quip', 'error');
            }
        });
    }

    function parseScheduleFromQuipHTML(html) {
        try {
            const parser = new DOMParser();
            const doc    = parser.parseFromString(html, 'text/html');
            const allTables = doc.querySelectorAll('table');

            // Filter out hidden sheets — Quip marks hidden sheets with class="hidden",
            // style="display:none", or aria-hidden on the table or its ancestors
            const isHidden = (el) => {
                let node = el;
                while (node && node !== doc.body) {
                    if (node.style && node.style.display === 'none') return true;
                    if (node.getAttribute && node.getAttribute('aria-hidden') === 'true') return true;
                    if (node.classList && (node.classList.contains('hidden') || node.classList.contains('collapsed'))) return true;
                    node = node.parentElement;
                }
                return false;
            };

            const tables = Array.from(allTables).filter(t => !isHidden(t));

            console.log('[ConnectVision] Quip HTML: Total tables:', allTables.length, '| Visible:', tables.length);

            // Build today's identifiers
            const now = new Date();
            const dayNum = now.getDate();
            const suffix = ([,'st','nd','rd'][(dayNum%100-20)%10] || [,'st','nd','rd'][dayNum%100] || 'th');
            const todayDate = (dayNum + suffix).toLowerCase();  // e.g. "14th"
            const todayWeekday = now.toLocaleDateString('en-US', { weekday: 'long' }).toLowerCase();

            console.log('[ConnectVision] Looking for today:', todayWeekday, todayDate);

            // Target the FIRST visible table — this is the active/current schedule sheet
            // Then apply date/weekday matching within it to find today's section
            let scheduleTable = tables.length > 0 ? tables[0] : null;

            // If first visible table doesn't have today's date, check other visible tables
            if (scheduleTable && tables.length > 1) {
                const firstRows = scheduleTable.querySelectorAll('tr');
                let firstHasToday = false;
                for (let r = 0; r < Math.min(3, firstRows.length); r++) {
                    const cells = firstRows[r].querySelectorAll('td, th');
                    for (const cell of cells) {
                        const text = (cell.textContent || '').trim().toLowerCase();
                        if ((text.startsWith(todayWeekday) && text.includes(todayDate)) || text.includes(todayDate)) {
                            firstHasToday = true;
                            break;
                        }
                    }
                    if (firstHasToday) break;
                }

                // If first visible table doesn't have today's data, scan others as safety fallback
                if (!firstHasToday) {
                    console.log('[ConnectVision] First visible table does not have today\'s date, checking others...');
                    for (const tbl of tables.slice(1)) {
                        const rows = tbl.querySelectorAll('tr');
                        for (let r = 0; r < Math.min(3, rows.length); r++) {
                            const cells = rows[r].querySelectorAll('td, th');
                            for (const cell of cells) {
                                const text = (cell.textContent || '').trim().toLowerCase();
                                if (text.startsWith(todayWeekday) && text.includes(todayDate)) {
                                    scheduleTable = tbl;
                                    break;
                                }
                            }
                            if (scheduleTable !== tables[0]) break;
                        }
                        if (scheduleTable !== tables[0]) break;
                    }
                }
            }

            if (!scheduleTable) {
                showScheduleMessage('\u274c No visible schedule table found in Quip', 'error');
                return;
            }

            const rows = scheduleTable.querySelectorAll('tr');
            const grid = Array.from(rows).map(row =>
                Array.from(row.querySelectorAll('td, th')).map(c => { const t = (c.textContent || '').replace(/[\u200b\u200c\u200d\ufeff\u00ad]/g, '').trim(); return t || null; })
            );
            console.log('[ConnectVision] Quip schedule table:', grid.length, 'rows,', (grid[0]||[]).length, 'cols');
            if (grid[0]) console.log('[ConnectVision] Grid Row 0:', JSON.stringify(grid[0].slice(0, 14)));
            if (grid[1]) console.log('[ConnectVision] Grid Row 1:', JSON.stringify(grid[1].slice(0, 14)));
            if (grid.length === 0) { showScheduleMessage('\u274c No rows in schedule table', 'error'); return; }
            parseScheduleFromGrid(grid);
        } catch (e) { showScheduleMessage('\u274c Quip HTML parse error: ' + e.message, 'error'); }
    }

    // ════════════════════════════════════════════
    // TAB: OUT-OF-SLOT
    // ════════════════════════════════════════════

    const buildOutOfSlotTab = () => {
        const w = document.createElement('div');
        w.style.cssText = 'display:flex;flex-direction:column;height:100%;padding:10px;gap:8px;overflow:hidden';
        w.innerHTML = `
            <div style="display:flex;gap:8px;align-items:center;background:#fff3cd;border-left:4px solid #ff9900;padding:7px 12px;border-radius:4px;flex-shrink:0;flex-wrap:wrap">
                <span><b>Buffer:</b> <span id="cv-oos-buffer">00:00</span></span>
                <span><b>Time:</b>   <span id="cv-oos-time">--:--:--</span></span>
                <span><b>Slot:</b>   <span id="cv-oos-slot" style="font-weight:bold;color:#067d62">--</span></span>
                <button id="cv-unhideAllBtn"  style="display:none;padding:3px 10px;background:#067d62;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:bold;margin-left:auto">👁️ Unhide All</button>
                <button id="cv-oos-csvBtn"    style="padding:3px 10px;background:#232f3e;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:bold">📥 CSV</button>
                <button id="cv-oos-uploadBtn" style="padding:3px 10px;background:#0073bb;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:bold">⬆️ Upload to SP</button>
            </div>
            <div id="cv-oos-uploadMsg" style="font-size:12px;min-height:16px;flex-shrink:0;padding:0 4px"></div>
            <!-- Pause warning -->
            <div id="cv-pauseWarning" style="display:none;background:#d13212;color:#fff;padding:12px 16px;border-radius:6px;font-size:14px;font-weight:bold;text-align:center;flex-shrink:0;line-height:1.7">
                ⚠️ The Connect page is paused. Out-of-Slot break data is not accurate.<br>
                <span style="font-size:13px;font-weight:normal;opacity:.9">Please resume the Connect page to see accurate out-of-slot breaks.</span>
            </div>
            <div id="cv-oos-tableWrap" style="flex:1;overflow-y:auto;overflow-x:auto;border:1px solid #eee;border-radius:6px;min-height:0">
                <div style="text-align:center;color:#999;padding:30px">No out-of-slot breaks detected</div>
            </div>
            <div style="display:flex;justify-content:space-between;align-items:center;font-size:12px;color:#888;flex-shrink:0">
                <span>Total: <b><span id="cv-oos-count">0</span></b> agents</span>
            </div>
            <div id="cv-violationLogWrap" style="flex-shrink:0;max-height:190px;overflow-y:auto;border:1px solid #cce5ff;border-radius:6px">
                <div style="padding:6px 12px;font-weight:bold;font-size:13px;background:#0073bb;color:#fff;position:sticky;top:0;display:flex;align-items:center;gap:8px">
                    📋 Violation Log
                    <span id="cv-vlog-count" style="font-size:12px;font-weight:normal;background:rgba(255,255,255,.25);padding:1px 8px;border-radius:10px">0</span>
                    <button id="cv-vlog-toggle" style="margin-left:auto;background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.4);color:#fff;border-radius:3px;padding:0 7px;cursor:pointer;font-size:14px;line-height:1.4" title="Minimize/Expand Violation Log">−</button>
                </div>
                <div id="cv-violationLogBody" style="background:#f0f8ff">
                    <div style="padding:8px 12px;color:#999;font-size:12px">No violations logged yet</div>
                </div>
            </div>
        `;
        return w;
    };

    const wireOutOfSlotTab = () => {
        // Violation log minimize/expand toggle
        document.getElementById('cv-vlog-toggle')?.addEventListener('click', () => {
            const body = document.getElementById('cv-violationLogBody');
            const btn  = document.getElementById('cv-vlog-toggle');
            if (!body || !btn) return;
            const isHidden = body.style.display === 'none';
            body.style.display = isHidden ? '' : 'none';
            btn.textContent    = isHidden ? '−' : '+';
        });
        document.getElementById('cv-unhideAllBtn')?.addEventListener('click', () => {
            state.hiddenOutOfSlotAgents.clear();
            renderOutOfSlotContent();
        });
        document.getElementById('cv-oos-csvBtn')?.addEventListener('click',    downloadOutOfSlotCSV);
        document.getElementById('cv-oos-uploadBtn')?.addEventListener('click', uploadToSharePoint);
    };

    const renderOutOfSlotContent = () => {
        const wrap    = document.getElementById('cv-oos-tableWrap');
        const warning = document.getElementById('cv-pauseWarning');
        const countEl = document.getElementById('cv-oos-count');
        const bufEl   = document.getElementById('cv-oos-buffer');
        const timeEl  = document.getElementById('cv-oos-time');
        const slotEl  = document.getElementById('cv-oos-slot');
        const unhide  = document.getElementById('cv-unhideAllBtn');
        if (!wrap) return;

        // Pause state
        if (warning) warning.style.display = state.isPaused ? 'block' : 'none';
        if (wrap)    wrap.style.display     = state.isPaused ? 'none'  : 'block';
        if (countEl) countEl.textContent    = state.isPaused ? '-'     : String(state.outOfSlotAgents.filter(a => !state.hiddenOutOfSlotAgents.has(a.login.toLowerCase())).length);

        if (bufEl) bufEl.textContent = `${String(state.bufferMinutes).padStart(2,'0')}:${String(state.bufferSeconds).padStart(2,'0')}`;

        const now = new Date();
        if (timeEl) timeEl.textContent = `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')}`;

        const current = now.getHours()*60 + now.getMinutes() + now.getSeconds()/60;
        const found   = getUniqueSlots().find(s => current >= s.start && current <= s.end);
        if (slotEl) {
            slotEl.textContent  = found ? `Break Slot: ${found.display} (${found.duration} Mins)` : 'No scheduled slots right now';
            slotEl.style.color  = found ? '#067d62' : '#888';
        }

        if (unhide) {
            unhide.style.display = state.hiddenOutOfSlotAgents.size > 0 ? 'inline-block' : 'none';
            if (state.hiddenOutOfSlotAgents.size > 0) unhide.textContent = `👁️ Unhide All (${state.hiddenOutOfSlotAgents.size})`;
        }

        if (state.isPaused) return;

        // Skip refresh if user is typing a comment
        if (document.activeElement?.classList.contains('cv-comment-input')) return;

        const visible = state.outOfSlotAgents.filter(a => !state.hiddenOutOfSlotAgents.has(a.login.toLowerCase()));

        if (visible.length === 0) {
            wrap.innerHTML = `<div style="text-align:center;color:#666;padding:30px">${state.outOfSlotAgents.length > 0 ? 'All agents hidden' : 'No out-of-slot breaks detected'}</div>`;
            return;
        }

        let html = `<table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
                <tr style="background:#d13212;color:#fff;position:sticky;top:0;z-index:2">
                    <th style="padding:6px 8px;width:36px">Hide</th>
                    <th style="padding:6px 8px;text-align:left">Login</th>
                    <th style="padding:6px 8px;text-align:left">Manager</th>
                    <th style="padding:6px 8px;text-align:left">Activity</th>
                    <th style="padding:6px 8px;text-align:center">Duration</th>
                    <th style="padding:6px 8px;text-align:center;background:#8b0000">OOS Time</th>
                    <th style="padding:6px 8px;text-align:center">Break 10</th>
                    <th style="padding:6px 8px;text-align:center">Break 20</th>
                    <th style="padding:6px 8px;text-align:center">Break 30</th>
                    <th style="padding:6px 8px;text-align:left;min-width:120px">Comments</th>
                </tr>
            </thead><tbody>`;

        visible.forEach((agent, idx) => {
            const bg      = idx % 2 === 0 ? '#fff' : '#fdf0f0';
            const loginUp = sanitize(agent.login.toUpperCase());
            const loginLw = agent.login.toLowerCase();
            const comment = sanitize(state.managerComments[loginLw] || '-');
            html += `<tr style="background:${bg}">
                <td style="padding:5px 8px;text-align:center">
                    <button class="cv-hideBtn" data-login="${loginLw}" style="background:#ff9900;color:#fff;border:none;border-radius:3px;padding:1px 7px;cursor:pointer;font-size:12px">👁️</button>
                </td>
                <td class="cv-oosLogin" data-login="${loginUp}" style="padding:5px 8px;font-weight:bold;color:#d13212;cursor:pointer">${loginUp}</td>
                <td style="padding:5px 8px">${sanitize(agent.manager)}</td>
                <td style="padding:5px 8px">${sanitize(agent.activity)}</td>
                <td style="padding:5px 8px;text-align:center;font-weight:bold">${sanitize(agent.duration)}</td>
                <td style="padding:5px 8px;text-align:center;font-weight:bold;color:#8b0000">${sanitize(agent.oosTime || 'N/A')}</td>
                <td style="padding:5px 8px;text-align:center">${sanitize(agent.break10)}</td>
                <td style="padding:5px 8px;text-align:center">${sanitize(agent.break20)}</td>
                <td style="padding:5px 8px;text-align:center">${sanitize(agent.break30)}</td>
                <td style="padding:5px 8px">
                    <input type="text" class="cv-comment-input" data-login="${loginLw}" value="${comment}"
                        style="width:100%;padding:3px 5px;border:1px solid #ccc;border-radius:3px;font-size:12px" placeholder="Comment...">
                </td>
            </tr>`;
        });

        html += '</tbody></table>';
        wrap.innerHTML = html;

        wrap.querySelectorAll('.cv-hideBtn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                state.hiddenOutOfSlotAgents.add(btn.getAttribute('data-login'));
                renderOutOfSlotContent();
            });
        });
        wrap.querySelectorAll('.cv-oosLogin').forEach(cell => {
            cell.addEventListener('click', () => {
                navigator.clipboard.writeText(cell.getAttribute('data-login')).then(() => {
                    cell.style.background = '#ffd'; setTimeout(() => { cell.style.background=''; }, 500);
                });
            });
        });
        // Comment inputs — local save only, no uploads
        wrap.querySelectorAll('.cv-comment-input').forEach(input => {
            input.addEventListener('input', () => {
                const login = input.getAttribute('data-login');
                state.managerComments[login] = input.value.trim() || '-';
                clearTimeout(state.commentSaveTimers[login]);
                state.commentSaveTimers[login] = setTimeout(() => {
                    console.log(`[ConnectVision] 💬 Comment stored locally for ${login}: ${state.managerComments[login]}`);
                }, 800);
            });
        });

        renderViolationLog();
    };

    const renderViolationLog = () => {
        const body    = document.getElementById('cv-violationLogBody');
        const countEl = document.getElementById('cv-vlog-count');
        if (!body) return;
        const log = state.violationLog;
        if (countEl) countEl.textContent = String(log.length);
        if (log.length === 0) {
            body.innerHTML = '<div style="padding:8px 12px;color:#999;font-size:12px">No violations logged yet</div>';
            return;
        }
        let html = '';
        for (const v of [...log].reverse()) {
            const slot    = (v.slotStart !== undefined)
                ? `${Math.floor(v.slotStart/60)}:${String(v.slotStart%60).padStart(2,'0')}-${Math.floor(v.slotEnd/60)}:${String(v.slotEnd%60).padStart(2,'0')}`
                : 'N/A';
            const status  = v.uploaded ? '\u2705' : '\u23f3';
            const timeStr = v.detectedAt ? new Date(v.detectedAt).toLocaleTimeString() : '';
            html += `<div style="padding:4px 12px;border-bottom:1px solid #dde;font-size:12px;display:flex;gap:8px;align-items:center">
                <span>${status}</span>
                <span style="font-weight:bold;min-width:80px">${sanitize(v.login.toUpperCase())}</span>
                <span style="color:#555;min-width:55px">${sanitize(v.activity)}</span>
                <span style="color:#0073bb;font-size:11px;min-width:95px">${slot}</span>
                <span style="color:#888;font-size:11px">${timeStr}</span>
            </div>`;
        }
        body.innerHTML = html;
    };

    const downloadOutOfSlotCSV = () => {
        if (state.outOfSlotAgents.length === 0) { alert('❌ No out-of-slot breaks to download'); return; }
        const now     = new Date();
        const cTime   = `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}:${String(now.getSeconds()).padStart(2,'0')}`;
        const buf     = `${String(state.bufferMinutes).padStart(2,'0')}:${String(state.bufferSeconds).padStart(2,'0')}`;
        const current = now.getHours()*60 + now.getMinutes() + now.getSeconds()/60;
        const found   = getUniqueSlots().find(s => current >= s.start && current <= s.end);
        const slot    = found ? `Break Slot: ${found.display} (${found.duration} Mins)` : 'No scheduled slots';
        let csv = `Buffer:${buf},Time:${cTime},${slot}${CRLF}`;
        csv    += `Login,Manager,Activity,Duration,Break 10,Break 20,Break 30,Comments${CRLF}`;
        state.outOfSlotAgents
            .filter(a => !state.hiddenOutOfSlotAgents.has(a.login.toLowerCase()))
            .forEach(a => { csv += [a.login.toLowerCase(),a.manager,a.activity,a.duration,a.break10,a.break20,a.break30,state.managerComments[a.login.toLowerCase()]||'-'].join(',')+CRLF; });
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `OOS_${getISTTimestamp()}.csv` });
        link.style.display = 'none';
        document.body.appendChild(link); link.click(); document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
        alert('✅ CSV downloaded');
    };

    // ════════════════════════════════════════════
    // TAB: SCHEDULE
    // ════════════════════════════════════════════

    const buildScheduleTab = () => {
        const w = document.createElement('div');
        w.style.cssText = 'display:flex;flex-direction:column;height:100%;padding:16px;gap:14px;overflow-y:auto';
        w.innerHTML = `
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
                <div style="background:#f9f9f9;border:1px solid #ddd;border-radius:8px;padding:14px">
                    <div style="font-weight:bold;font-size:14px;margin-bottom:10px;color:#232f3e">📂 Upload Schedule (.xlsx)</div>
                    <input type="file" id="cv-scheduleFile" accept=".xlsx" style="width:100%;font-size:13px;cursor:pointer;margin-bottom:6px">
                    <button id="cv-quipLoadBtn" style="width:100%;padding:6px;background:#232f3e;color:#ff9900;border:1px solid #ff9900;border-radius:4px;cursor:pointer;font-size:13px;font-weight:bold">🔗 Load from Quip</button>
                    <button id="cv-downloadScheduleBtn" style="width:100%;padding:6px;margin-top:6px;background:#0073bb;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:bold">📥 Download Loaded Schedule</button>
                    <div id="cv-scheduleStatus"  style="margin-top:8px;font-size:13px;color:#555">No schedule loaded</div>
                    <div id="cv-uploadMessage"   style="margin-top:4px;font-size:12px;min-height:16px"></div>
                </div>
                <div style="background:#f9f9f9;border:1px solid #ddd;border-radius:8px;padding:14px">
                    <div style="font-weight:bold;font-size:14px;margin-bottom:10px;color:#232f3e">⏱️ Buffer Time (MM : SS)</div>
                    <div style="display:flex;gap:8px;align-items:center">
                        <input type="number" id="cv-bufMin" min="0" value="${state.bufferMinutes}" style="width:58px;padding:5px;font-size:14px;border:1px solid #ccc;border-radius:4px">
                        <span style="font-weight:bold;font-size:16px">:</span>
                        <input type="number" id="cv-bufSec" min="0" max="59" value="${state.bufferSeconds}" style="width:58px;padding:5px;font-size:14px;border:1px solid #ccc;border-radius:4px">
                        <button id="cv-applyBuffer" style="padding:5px 14px;background:#ff9900;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:14px;font-weight:bold">Apply</button>
                    </div>
                    <div style="margin-top:8px;font-size:13px;color:#555">Current: <b><span id="cv-currentBuffer">${String(state.bufferMinutes).padStart(2,'0')}:${String(state.bufferSeconds).padStart(2,'0')}</span></b></div>
                </div>
            </div>
            <div style="background:#f9f9f9;border:1px solid #ddd;border-radius:8px;padding:14px">
                <button id="cv-toggleMonitor" style="width:100%;padding:10px;background:${state.isScheduleMonitoring?'#d13212':'#232f3e'};color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:bold">
                    ${state.isScheduleMonitoring ? '⏹ Stop Schedule Monitoring' : '▶ Start Schedule Monitoring'}
                </button>
                <div style="margin-top:10px;font-size:13px;color:#666;line-height:1.7">
                    🟣 <b>Purple</b> = Break/Lunch outside scheduled slot (Highest Priority)<br>
                    🔴 <b>Red</b> = Available ≥ red threshold &nbsp;|&nbsp; 🟡 <b>Yellow</b> = Available in yellow range<br>
                    🔵 <b>Blue</b> = Break > threshold &nbsp;|&nbsp; 🟠 <b>Orange</b> = Lunch > threshold
                </div>
            </div>
        `;
        return w;
    };

    const wireScheduleTab = () => {
        document.getElementById('cv-scheduleFile')?.addEventListener('change', handleFileUpload);
        document.getElementById('cv-quipLoadBtn')?.addEventListener('click',   loadScheduleFromQuip);
        document.getElementById('cv-downloadScheduleBtn')?.addEventListener('click', downloadLoadedSchedule);
        document.getElementById('cv-applyBuffer')?.addEventListener('click',  applyBufferTime);
        document.getElementById('cv-toggleMonitor')?.addEventListener('click', toggleScheduleMonitoring);
    };

    const renderScheduleStatus = () => {
        const btn = document.getElementById('cv-toggleMonitor');
        if (!btn) return;
        btn.textContent      = state.isScheduleMonitoring ? '⏹ Stop Schedule Monitoring' : '▶ Start Schedule Monitoring';
        btn.style.background = state.isScheduleMonitoring ? '#d13212' : '#232f3e';
    };

    async function handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        if (!file.name.endsWith('.xlsx')) { showScheduleMessage('❌ Please upload an .xlsx file', 'error'); return; }
        if (typeof DecompressionStream === 'undefined') { showScheduleMessage('❌ DecompressionStream not supported. Use Chrome/Edge.', 'error'); return; }
        try {
            showScheduleMessage('⏳ Processing...', 'info');
            const data = await parseXLSXFile(file);
            if (!data?.length) { showScheduleMessage('❌ Could not read file', 'error'); return; }
            parseScheduleFromGrid(data);
        } catch (err) { showScheduleMessage('❌ Error: ' + err.message, 'error'); }
    }


    function downloadLoadedSchedule() {
        const schedules = state.breakSchedules;
        const logins = Object.keys(schedules).sort();
        if (logins.length === 0) {
            showScheduleMessage('\u274c No schedule loaded to download', 'error');
            return;
        }
        // Build CSV content
        const CRLF = '\r\n';
        let csv = 'Login,Manager,Break 10 Mins,Break 20 Mins,Break 30 Mins' + CRLF;
        for (const login of logins) {
            const s = schedules[login];
            csv += [
                login,
                s.manager || 'N/A',
                s.break10 || 'N/A',
                s.break20 || 'N/A',
                s.break30 || 'N/A'
            ].join(',') + CRLF;
        }
        csv += CRLF + `Total: ${logins.length} agents` + CRLF;

        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = Object.assign(document.createElement('a'), {
            href: URL.createObjectURL(blob),
            download: `LoadedSchedule_${new Date().toISOString().slice(0,10)}.csv`
        });
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
        showScheduleMessage(`\u2705 Downloaded ${logins.length} agent schedules`, 'success');
    }

    function applyBufferTime() {
        state.bufferMinutes = Math.max(0, parseInt(document.getElementById('cv-bufMin')?.value)||0);
        state.bufferSeconds = Math.max(0, Math.min(59, parseInt(document.getElementById('cv-bufSec')?.value)||0));
        const el = document.getElementById('cv-currentBuffer');
        if (el) el.textContent = `${String(state.bufferMinutes).padStart(2,'0')}:${String(state.bufferSeconds).padStart(2,'0')}`;
        showScheduleMessage('✅ Buffer updated', 'success');
    }

    function toggleScheduleMonitoring() {
        if (!state.isScheduleMonitoring && Object.keys(state.breakSchedules).length === 0) {
            showScheduleMessage('❌ Upload a schedule file first', 'error'); return;
        }
        state.isScheduleMonitoring = !state.isScheduleMonitoring;
        renderScheduleStatus();
        showScheduleMessage(state.isScheduleMonitoring ? '✅ Monitoring started' : '⏸ Monitoring stopped',
            state.isScheduleMonitoring ? 'success' : 'info');
    }
    // ════════════════════════════════════════════
    // TAB: SETTINGS
    // ════════════════════════════════════════════

    const buildSettingsTab = () => {
        const w = document.createElement('div');
        w.style.cssText = 'display:flex;flex-direction:column;height:100%;padding:14px;gap:10px;overflow-y:auto';
        const s = state.settings;

        const block = (color, bg, label, chkId, minMId, minSId, maxMId, maxSId, showMax) => `
            <div style="background:${bg};border-radius:6px;border-left:4px solid ${color};padding:11px 14px;display:flex;flex-wrap:wrap;align-items:center;gap:10px">
                <label style="display:flex;align-items:center;gap:6px;font-weight:bold;font-size:13px;min-width:160px;cursor:pointer">
                    <input type="checkbox" id="${chkId}" ${s[chkId.replace('cv-','')] ? 'checked' : ''} style="width:15px;height:15px;cursor:pointer">
                    ${label}
                </label>
                <span style="font-size:13px;color:#666">≥</span>
                <input type="number" id="${minMId}" value="${s[minMId.replace('cv-','')]}" min="0" step="1" style="width:52px;padding:3px;border:1px solid #ccc;border-radius:4px;font-size:13px">
                <span style="font-size:12px;color:#666">min</span>
                <input type="number" id="${minSId}" value="${s[minSId.replace('cv-','')]}" min="0" max="59" style="width:52px;padding:3px;border:1px solid #ccc;border-radius:4px;font-size:13px">
                <span style="font-size:12px;color:#666">sec</span>
                ${showMax ? `
                    <span style="font-size:13px;color:#666">&lt;</span>
                    <input type="number" id="${maxMId}" value="${s[maxMId.replace('cv-','')]}" min="0" step="1" style="width:52px;padding:3px;border:1px solid #ccc;border-radius:4px;font-size:13px">
                    <span style="font-size:12px;color:#666">min</span>
                    <input type="number" id="${maxSId}" value="${s[maxSId.replace('cv-','')]}" min="0" max="59" style="width:52px;padding:3px;border:1px solid #ccc;border-radius:4px;font-size:13px">
                    <span style="font-size:12px;color:#666">sec</span>
                ` : ''}
            </div>`;

        w.innerHTML = `
            <div style="font-weight:bold;font-size:14px;color:#232f3e;margin-bottom:2px">⚙️ Duration-Based Highlight Settings</div>
            ${block('#ffeb3b','#fffef0','🟡 Yellow (Available)','cv-yellowEnabled','cv-yellowMinMinutes','cv-yellowMinSeconds','cv-yellowMaxMinutes','cv-yellowMaxSeconds',true)}
            ${block('#f44336','#fff0f0','🔴 Red (Available)',   'cv-redEnabled',   'cv-redMinMinutes',   'cv-redMinSeconds',   '','',false)}
            ${block('#2196F3','#f0f8ff','🔵 Blue (Break)',      'cv-blueEnabled',  'cv-breakMinMinutes', 'cv-breakMinSeconds', '','',false)}
            ${block('#ff9800','#fff5f0','🟠 Orange (Lunch)',    'cv-orangeEnabled','cv-lunchMinMinutes', 'cv-lunchMinSeconds', '','',false)}
            <div style="background:#f0f0ff;border-left:4px solid #9c27b0;border-radius:6px;padding:9px 14px;font-size:13px;color:#555">
                🟣 <b>Purple (Out-of-Slot)</b> — Configured in the Schedule tab. Overrides all other colours.
            </div>
            <div style="background:#f0fff0;border-left:4px solid #067d62;border-radius:6px;padding:11px 14px">
                <div style="font-weight:bold;font-size:13px;margin-bottom:8px;color:#232f3e">🔐 Quip API Token</div>
                <input type="password" id="cv-quipToken" value="${GM_getValue('quipToken','')}" placeholder="Paste Quip API token…"
                    style="width:100%;padding:5px;border:1px solid #ccc;border-radius:4px;font-size:13px;box-sizing:border-box">
                <div style="font-size:11px;color:#888;margin-top:4px">Stored locally. Used to load break schedule from Quip.</div>
            </div>
            <div style="background:#f0fff0;border-left:4px solid #067d62;border-radius:6px;padding:11px 14px;display:flex;flex-wrap:wrap;align-items:center;gap:10px">
                <label style="font-weight:bold;font-size:13px;min-width:200px">🟢 Late-to-Break Alert</label>
                <span style="font-size:13px;color:#666">≥</span>
                <input type="number" id="cv-lateToBreakAlertMin" value="${state.lateToBreakAlertMin}" min="0" step="1"
                    style="width:52px;padding:3px;border:1px solid #ccc;border-radius:4px;font-size:13px">
                <span style="font-size:12px;color:#666">min into slot (0 = disabled)</span>
            </div>
            <div style="display:flex;gap:10px;justify-content:flex-end;padding-top:4px">
                <button id="cv-resetSettings" style="padding:7px 18px;background:#ccc;border:none;border-radius:5px;cursor:pointer;font-size:13px">Reset</button>
                <button id="cv-saveSettings"  style="padding:7px 22px;background:#4CAF50;color:#fff;border:none;border-radius:5px;cursor:pointer;font-size:13px;font-weight:bold">💾 Save</button>
            </div>
        `;
        return w;
    };

    const wireSettingsTab = () => {
        document.getElementById('cv-saveSettings')?.addEventListener('click',  saveSettings);
        document.getElementById('cv-resetSettings')?.addEventListener('click', resetSettings);
    };

    const saveSettings = () => {
        const getInt  = id => parseInt(document.getElementById(id)?.value) || 0;
        const getBool = id => document.getElementById(id)?.checked ?? false;
        const newS = {
            yellowEnabled:    getBool('cv-yellowEnabled'),
            redEnabled:       getBool('cv-redEnabled'),
            blueEnabled:      getBool('cv-blueEnabled'),
            orangeEnabled:    getBool('cv-orangeEnabled'),
            yellowMinMinutes: getInt('cv-yellowMinMinutes'),
            yellowMinSeconds: getInt('cv-yellowMinSeconds'),
            yellowMaxMinutes: getInt('cv-yellowMaxMinutes'),
            yellowMaxSeconds: getInt('cv-yellowMaxSeconds'),
            redMinMinutes:    getInt('cv-redMinMinutes'),
            redMinSeconds:    getInt('cv-redMinSeconds'),
            breakMinMinutes:  getInt('cv-breakMinMinutes'),
            breakMinSeconds:  getInt('cv-breakMinSeconds'),
            lunchMinMinutes:  getInt('cv-lunchMinMinutes'),
            lunchMinSeconds:  getInt('cv-lunchMinSeconds')
        };
        const yMin = convertToTotalMinutes(newS.yellowMinMinutes, newS.yellowMinSeconds);
        const yMax = convertToTotalMinutes(newS.yellowMaxMinutes, newS.yellowMaxSeconds);
        const rMin = convertToTotalMinutes(newS.redMinMinutes,    newS.redMinSeconds);
        if (newS.yellowEnabled && yMin >= yMax) { alert('⚠️ Yellow min must be < yellow max'); return; }
        if (newS.yellowEnabled && newS.redEnabled && rMin < yMax) { alert('⚠️ Red min should be ≥ yellow max'); return; }
        state.settings = newS;
        for (const [k, v] of Object.entries(newS)) GM_setValue(k, v);
        // Persist Quip token and late-to-break threshold
        const _qt  = document.getElementById('cv-quipToken')?.value?.trim() || '';
        GM_setValue('quipToken', _qt);
        const _ltb = parseInt(document.getElementById('cv-lateToBreakAlertMin')?.value) || 0;
        state.lateToBreakAlertMin = _ltb;
        GM_setValue('lateToBreakAlertMin', _ltb);
        unifiedUpdate();
        const btn = document.getElementById('cv-saveSettings');
        if (btn) { btn.textContent='✅ Saved!'; btn.style.background='#067d62'; setTimeout(()=>{ btn.textContent='💾 Save'; btn.style.background='#4CAF50'; },2000); }
    };

    const resetSettings = () => {
        if (!confirm('Reset all highlight settings to defaults?')) return;
        const def = { yellowEnabled:true,redEnabled:true,blueEnabled:true,orangeEnabled:true, yellowMinMinutes:2,yellowMinSeconds:0,yellowMaxMinutes:5,yellowMaxSeconds:0, redMinMinutes:5,redMinSeconds:0,breakMinMinutes:20,breakMinSeconds:0,lunchMinMinutes:30,lunchMinSeconds:0 };
        state.settings = { ...def };
        for (const [k,v] of Object.entries(def)) GM_setValue(k,v);
        state.lateToBreakAlertMin = 2;
        GM_setValue('lateToBreakAlertMin', 2);
        switchTab('settings');
        unifiedUpdate();
    };

    // ==================== OPEN BUTTON ====================

    const createOpenButton = () => {
        const btn = document.createElement('button');
        btn.id    = 'cv-openBtn';
        btn.innerHTML = '🖥️ ConnectVision';
        btn.style.cssText = `
            position:fixed; top:10px; left:50%; transform:translateX(-50%);
            z-index:10000; padding:8px 22px;
            background:#232f3e; color:#ff9900;
            border:2px solid #ff9900; border-radius:6px;
            cursor:pointer; font-size:14px; font-weight:bold;
            box-shadow:0 2px 8px rgba(0,0,0,0.3);
        `;
        btn.onmouseover = function() { this.style.background='#ff9900'; this.style.color='#232f3e'; };
        btn.onmouseout  = function() { this.style.background='#232f3e'; this.style.color='#ff9900'; };
        btn.onclick = () => {
            if (!state.isActive) return; // Safety guard
            if (state.dashboardEl) {
                state.dashboardEl.style.display = 'flex';
                // Restore if minimized
                const content = document.getElementById('cv-tabContent');
                const tabbar  = document.getElementById('cv-tabbar');
                const minBtn  = document.getElementById('cv-minimizeBtn');
                if (content) content.style.display = 'flex';
                if (tabbar)  tabbar.style.display  = 'flex';
                if (minBtn)  minBtn.textContent    = '−';
                state.dashboardEl.style.height = '560px';
            }
        };
        document.body.appendChild(btn);
    };

    // ==================== CLEANUP ====================

    window.addEventListener('beforeunload', () => {
        if (state.observer)      state.observer.disconnect();
        if (state.leaseInterval) clearInterval(state.leaseInterval);
        clearTimeout(state.updateTimeout);
        clearLease();
        if (_bc) _bc.close();
    });

    // ==================== ENTRY POINT ====================

    const currentUrl = window.location.href;

    if (currentUrl.includes('aria.ats.a2z.com')) {
        // ARIA page — ATR scanner only, no tab guard needed
        console.log('[ConnectVision ARIA] ATR scanner mode active');
        scanAndStoreARIAStatus();
        setInterval(scanAndStoreARIAStatus, 5000);

    } else {
        // Connect page — full dashboard with tab guard
        console.log('[ConnectVision] Starting with tab guard... Tab ID:', MY_TAB_ID);

        const boot = () => {
            createOpenButton();
            initTabGuard();
        };

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => setTimeout(boot, 2000));
        } else {
            setTimeout(boot, 2000);
        }
    }

})();
