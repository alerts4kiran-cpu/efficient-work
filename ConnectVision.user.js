// ==UserScript==
// @name         ConnectVision - Amazon Connect Ultimate Monitoring Suite
// @namespace    http://tampermonkey.net/
// @version      6.0
// @description  ConnectVision: Ultimate monitoring suite for Amazon Connect with duration-based highlighting, activity tracking, break schedule compliance, resizable panels, and CSV exports
// @author       alerts4kiran-cpu
// @match        https://c2-na-prod.my.connect.aws/real-time-metrics*
// @match        https://c2-na-prod.awsapps.com/connect/real-time-metrics*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_listValues
// @grant        GM_deleteValue
// @updateURL    https://raw.githubusercontent.com/alerts4kiran-cpu/efficient-work/main/ConnectVision.user.js
// @downloadURL  https://raw.githubusercontent.com/alerts4kiran-cpu/efficient-work/main/ConnectVision.user.js
// ==/UserScript==

(function() {
    'use strict';

    // ==================== CONFIGURATION ====================
    const CONFIG = {
        DEBOUNCE: 500,
        UPDATE_THROTTLE: 500,
        UNIFIED_INTERVAL: 5000, // 5 seconds for all updates
        TIME_REGEX: /^\d{1,2}:\d{2}(:\d{2})?$/,
        LETTER_REGEX: /[a-zA-Z]/,
        ACTIVITY_SUMMARY_KEYWORDS: [
            '=== ACTIVITY SUMMARY ===', 'Activity', 'Total', 'After contact work',
            'Available', 'On contact', 'Break', 'Lunch', 'Personal', 'Training',
            'Meeting', 'Project', 'Manager 1-1', 'Incoming', 'Missed', 'Outage',
            'Manager Approved', 'System/Power/Internet Outage', 'Skip Meeting',
            'Start Up', 'Team Huddle'
        ]
    };

    // ==================== STATE MANAGEMENT ====================
    const state = {
        // Duration-based highlighting settings
        settings: {
            yellowEnabled: GM_getValue('yellowEnabled', true),
            redEnabled: GM_getValue('redEnabled', true),
            blueEnabled: GM_getValue('blueEnabled', true),
            orangeEnabled: GM_getValue('orangeEnabled', true),
            yellowMinMinutes: GM_getValue('yellowMinMinutes', 2),
            yellowMinSeconds: GM_getValue('yellowMinSeconds', 0),
            yellowMaxMinutes: GM_getValue('yellowMaxMinutes', 5),
            yellowMaxSeconds: GM_getValue('yellowMaxSeconds', 0),
            redMinMinutes: GM_getValue('redMinMinutes', 5),
            redMinSeconds: GM_getValue('redMinSeconds', 0),
            breakMinMinutes: GM_getValue('breakMinMinutes', 20),
            breakMinSeconds: GM_getValue('breakMinSeconds', 0),
            lunchMinMinutes: GM_getValue('lunchMinMinutes', 30),
            lunchMinSeconds: GM_getValue('lunchMinSeconds', 0)
        },
        // Break schedule monitoring
        breakSchedules: {},
        bufferMinutes: 0,
        bufferSeconds: 0,
        isScheduleMonitoring: false,
        outOfSlotAgents: [],
        // UI state
        isPanelMinimized: false,
        isOutOfSlotBoxMinimized: false,
        // Performance optimization
        observer: null,
        updateTimeout: null,
        summaryBox: null,
        outOfSlotBox: null,
        isUpdating: false,
        lastUpdateTime: 0,
        cachedActivityCounts: {},
        cachedActivityDetails: {},
        tableCache: null,
        tableCacheTime: 0,
        checkCount: 0
    };

    // ==================== UTILITY FUNCTIONS ====================
    const parseTimeToMinutes = (timeStr) => {
        if (!timeStr || timeStr === '-') return 0;
        const parts = timeStr.split(':');
        const len = parts.length;
        return len === 2 ? parseInt(parts[0]) + parseInt(parts[1]) / 60 :
               len === 3 ? parseInt(parts[0]) * 60 + parseInt(parts[1]) + parseInt(parts[2]) / 60 : 0;
    };

    const containsLetters = (str) => CONFIG.LETTER_REGEX.test(str);

    const isActivitySummaryRow = (agentLogin) => {
        if (!agentLogin) return true;
        const loginStr = agentLogin.toString().trim().toLowerCase();
        return CONFIG.ACTIVITY_SUMMARY_KEYWORDS.some(keyword => loginStr === keyword.toLowerCase());
    };

    const getISTTimestamp = () => {
        const now = new Date();
        return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}-${String(now.getSeconds()).padStart(2, '0')}`;
    };

    const getTables = () => {
        const now = Date.now();
        if (state.tableCache && now - state.tableCacheTime < 1000) {
            return state.tableCache;
        }
        state.tableCache = document.querySelectorAll('table tbody');
        state.tableCacheTime = now;
        return state.tableCache;
    };

    const convertToTotalMinutes = (minutes, seconds) => {
        return minutes + (seconds / 60);
    };

    const showScheduleMessage = (message, type) => {
        const messageDiv = document.getElementById('uploadMessage');
        if (!messageDiv) return;
        messageDiv.textContent = message;
        if (type === 'success') {
            messageDiv.style.color = '#067d62';
        } else if (type === 'error') {
            messageDiv.style.color = '#d13212';
        } else {
            messageDiv.style.color = '#666';
        }
        if (type !== 'error') {
            setTimeout(() => {
                messageDiv.textContent = '';
            }, 5000);
        }
    };

    // ==================== XLSX PARSER ====================
    async function parseXLSXFile(file) {
        const arrayBuffer = await file.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        const files = await extractZipFiles(uint8Array);
        const sharedStrings = parseSharedStrings(files['xl/sharedStrings.xml'] || '');
        const sheetXML = files['xl/worksheets/sheet1.xml'] || '';
        const rows = parseSheetData(sheetXML, sharedStrings);
        return rows;
    }

    async function extractZipFiles(data) {
        const files = {};
        let cdOffset = -1;
        for (let i = data.length - 22; i >= 0; i--) {
            if (data[i] === 0x50 && data[i+1] === 0x4b &&
                data[i+2] === 0x05 && data[i+3] === 0x06) {
                cdOffset = readUInt32LE(data, i + 16);
                break;
            }
        }
        if (cdOffset === -1) throw new Error('Invalid ZIP file');

        let offset = cdOffset;
        while (offset < data.length - 4) {
            if (data[offset] !== 0x50 || data[offset+1] !== 0x4b ||
                data[offset+2] !== 0x01 || data[offset+3] !== 0x02) break;

            const nameLength = readUInt16LE(data, offset + 28);
            const extraLength = readUInt16LE(data, offset + 30);
            const commentLength = readUInt16LE(data, offset + 32);
            const localHeaderOffset = readUInt32LE(data, offset + 42);
            const nameBytes = data.slice(offset + 46, offset + 46 + nameLength);
            const fileName = new TextDecoder().decode(nameBytes);
            const compressedSize = readUInt32LE(data, localHeaderOffset + 18);
            const localNameLength = readUInt16LE(data, localHeaderOffset + 26);
            const localExtraLength = readUInt16LE(data, localHeaderOffset + 28);
            const fileDataOffset = localHeaderOffset + 30 + localNameLength + localExtraLength;
            const fileData = data.slice(fileDataOffset, fileDataOffset + compressedSize);
            const compressionMethod = readUInt16LE(data, localHeaderOffset + 8);

            if (compressionMethod === 0) {
                files[fileName] = new TextDecoder().decode(fileData);
            } else if (compressionMethod === 8) {
                try {
                    const decompressed = await decompressDeflate(fileData);
                    files[fileName] = new TextDecoder().decode(decompressed);
                } catch (e) {
                    console.warn('Could not decompress:', fileName);
                }
            }
            offset += 46 + nameLength + extraLength + commentLength;
        }
        return files;
    }

    async function decompressDeflate(data) {
        if (typeof DecompressionStream !== 'undefined') {
            const ds = new DecompressionStream('deflate-raw');
            const writer = ds.writable.getWriter();
            writer.write(data);
            writer.close();
            const reader = ds.readable.getReader();
            const chunks = [];
            while (true) {
                const {done, value} = await reader.read();
                if (done) break;
                chunks.push(value);
            }
            const totalLength = chunks.reduce((acc, chunk) => acc + chunk.length, 0);
            const result = new Uint8Array(totalLength);
            let offset = 0;
            for (const chunk of chunks) {
                result.set(chunk, offset);
                offset += chunk.length;
            }
            return result;
        }
        return data;
    }

    function readUInt16LE(data, offset) {
        return data[offset] | (data[offset + 1] << 8);
    }

    function readUInt32LE(data, offset) {
        return data[offset] | (data[offset + 1] << 8) |
               (data[offset + 2] << 16) | (data[offset + 3] << 24);
    }

    function parseSharedStrings(xml) {
        const strings = [];
        const regex = /<t[^>]*>([^<]*)<\/t>/g;
        let match;
        while ((match = regex.exec(xml)) !== null) {
            strings.push(match[1]);
        }
        return strings;
    }

    function parseSheetData(xml, sharedStrings) {
        const rows = [];
        const cellData = {};
        const cellRegex1 = /<c r="([A-Z]+)(\d+)"[^>]*t="s"[^>]*><v>(\d+)<\/v><\/c>/g;
        let match;
        while ((match = cellRegex1.exec(xml)) !== null) {
            const col = match[1];
            const row = parseInt(match[2]);
            const stringIndex = parseInt(match[3]);
            if (!cellData[row]) cellData[row] = {};
            cellData[row][col] = sharedStrings[stringIndex] || '';
        }
        const cellRegex2 = /<c r="([A-Z]+)(\d+)"[^>]*t="inlineStr"[^>]*><is><t>([^<]*)<\/t><\/is><\/c>/g;
        while ((match = cellRegex2.exec(xml)) !== null) {
            const col = match[1];
            const row = parseInt(match[2]);
            const value = match[3];
            if (!cellData[row]) cellData[row] = {};
            cellData[row][col] = value;
        }
        const cellRegex3 = /<c r="([A-Z]+)(\d+)"[^>]*><v>([^<]*)<\/v><\/c>/g;
        while ((match = cellRegex3.exec(xml)) !== null) {
            const col = match[1];
            const row = parseInt(match[2]);
            const value = match[3];
            if (!cellData[row]) cellData[row] = {};
            if (!cellData[row][col]) {
                cellData[row][col] = value;
            }
        }
        const maxRow = Math.max(...Object.keys(cellData).map(Number));
        for (let i = 1; i <= maxRow; i++) {
            const rowData = [];
            if (cellData[i]) {
                const cols = Object.keys(cellData[i]).sort();
                for (const col of cols) {
                    const colIndex = columnToIndex(col);
                    while (rowData.length < colIndex) {
                        rowData.push('');
                    }
                    rowData[colIndex] = cellData[i][col];
                }
            }
            rows.push(rowData);
        }
        return rows;
    }

    function columnToIndex(col) {
        let index = 0;
        for (let i = 0; i < col.length; i++) {
            index = index * 26 + (col.charCodeAt(i) - 64);
        }
        return index - 1;
    }

    function parseSchedule(data) {
        state.breakSchedules = {};
        let successCount = 0;

        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(5, data.length); i++) {
            const row = data[i];
            if (row && String(row[0]).toLowerCase().includes('login')) {
                headerRowIndex = i;
                break;
            }
        }

        if (headerRowIndex === -1) {
            showScheduleMessage('❌ Error: Header row not found.', 'error');
            return;
        }

        const headers = data[headerRowIndex];
        const loginCol = headers.findIndex(h => h && String(h).toLowerCase().includes('login'));
        const managerCol = headers.findIndex(h => h && String(h).toLowerCase().includes('manager'));
        const break10Col = headers.findIndex(h => h && String(h).toLowerCase().includes('break') && String(h).includes('10'));
        const break20Col = headers.findIndex(h => h && String(h).toLowerCase().includes('break') && String(h).includes('20'));
        const break30Col = headers.findIndex(h => h && String(h).toLowerCase().includes('break') && String(h).includes('30'));

        if (loginCol === -1 || break10Col === -1 || break20Col === -1 || break30Col === -1) {
            showScheduleMessage('❌ Error: Required columns not found.', 'error');
            return;
        }

        for (let i = headerRowIndex + 1; i < data.length; i++) {
            const row = data[i];
            if (!row || !row[loginCol]) continue;
            const login = String(row[loginCol]).trim().toLowerCase();
            if (!login) continue;

            try {
                const breaks = [];
                const break10 = row[break10Col] ? String(row[break10Col]).trim() : '';
                const break20 = row[break20Col] ? String(row[break20Col]).trim() : '';
                const break30 = row[break30Col] ? String(row[break30Col]).trim() : '';

                if (break10) breaks.push(...parseTimeSlot(break10));
                if (break20) breaks.push(...parseTimeSlot(break20));
                if (break30) breaks.push(...parseTimeSlot(break30));

                if (breaks.length > 0) {
                    state.breakSchedules[login] = {
                        manager: managerCol !== -1 ? String(row[managerCol] || 'N/A').trim() : 'N/A',
                        breaks: breaks,
                        break10: break10 || 'N/A',
                        break20: break20 || 'N/A',
                        break30: break30 || 'N/A'
                    };
                    successCount++;
                }
            } catch (error) {
                console.error(`Error parsing schedule for ${login}:`, error);
            }
        }

        if (successCount > 0) {
            document.getElementById('scheduleStatus').textContent = `✅ ${successCount} agent schedules loaded`;
            showScheduleMessage(`✅ Success! Loaded ${successCount} agent schedules`, 'success');
            console.log('📋 Loaded schedules:', Object.keys(state.breakSchedules).slice(0, 10));
        } else {
            showScheduleMessage('❌ Error: No valid schedules found', 'error');
        }
    }

    function parseTimeSlot(timeStr) {
        if (!timeStr) return [];
        const slots = [];
        const timeString = String(timeStr).trim();
        const ranges = timeString.split(/[,;]/).map(s => s.trim());
        for (const range of ranges) {
            const match = range.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
            if (match) {
                const startHour = parseInt(match[1]);
                const startMin = parseInt(match[2]);
                const endHour = parseInt(match[3]);
                const endMin = parseInt(match[4]);
                slots.push({
                    start: startHour * 60 + startMin,
                    end: endHour * 60 + endMin
                });
            }
        }
        return slots;
    }

    function checkIfOutOfSchedule(agentLogin) {
        const scheduleData = state.breakSchedules[agentLogin.toLowerCase()];
        if (!scheduleData || !scheduleData.breaks || scheduleData.breaks.length === 0) {
            return false;
        }

        const now = new Date();
        const currentMinutes = now.getHours() * 60 + now.getMinutes() + (now.getSeconds() / 60);
        const bufferInMinutes = state.bufferMinutes + (state.bufferSeconds / 60);

        for (const slot of scheduleData.breaks) {
    const slotStart = slot.start - bufferInMinutes;  // ✅ SUBTRACT BUFFER FROM START
    const slotEnd = slot.end + bufferInMinutes;      // ✅ ADD BUFFER TO END

    if (currentMinutes >= slotStart && currentMinutes <= slotEnd) {
        return false;
    }
}

        return true;
    }
// ✅ NEW FUNCTION 1: Check if a duration string is actually a time range (schedule data)
    function isTimeRange(durationStr) {
        // Matches patterns like "22:15-22:25" or "16:00-16:20"
        return /^\d{2}:\d{2}-\d{2}:\d{2}$/.test(durationStr);
    }

    // ✅ NEW FUNCTION 2: Extract only active agents from real-time metrics tables
    function getActiveAgentsFromPage() {
        const activeAgents = new Set();
        const tables = getTables();

        for (let t = 0; t < tables.length; t++) {
            const rows = tables[t].rows;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const cells = row.cells;
                if (cells.length < 3) continue;

                const agentCell = cells[0]?.textContent?.trim();

                // Check if this row has channel information (Voice/Chat)
                const hasChannelInfo = Array.from(cells).some(cell =>
                    cell.textContent?.includes('Voice') ||
                    cell.textContent?.includes('Chat')
                );

                // Check if row has duration in HH:MM:SS format (not time range)
                const hasDuration = Array.from(cells).some(cell => {
                    const text = cell.textContent?.trim();
                    return text && /^\d{2}:\d{2}:\d{2}$/.test(text);
                });

                // Only add if this looks like a real agent row
                if (agentCell && hasChannelInfo && hasDuration) {
                    const loginMatch = agentCell.match(/[a-z]{6,8}/i);
                    if (loginMatch) {
                        activeAgents.add(loginMatch[0].toLowerCase());
                    }
                }
            }
        }

        return activeAgents;
    }

    // END OF PART 1
// PART 2 - UNIFIED MONITORING ENGINE & ACTIVITY SUMMARY
    // This continues from Part 1

    // ==================== UNIFIED HIGHLIGHTING & MONITORING ====================
    const unifiedUpdate = () => {
        state.checkCount++;
        state.outOfSlotAgents = [];
// ✅ NEW: Get active agents FIRST to prevent ghost entries
        const activeAgentsSet = getActiveAgentsFromPage();
        console.log(`[v6.0 Fixed] Active agents detected: ${activeAgentsSet.size}`);
        const tables = getTables();
        const len = tables.length;

        // Calculate threshold values
        const yellowMin = convertToTotalMinutes(state.settings.yellowMinMinutes, state.settings.yellowMinSeconds);
        const yellowMax = convertToTotalMinutes(state.settings.yellowMaxMinutes, state.settings.yellowMaxSeconds);
        const redMin = convertToTotalMinutes(state.settings.redMinMinutes, state.settings.redMinSeconds);
        const breakMin = convertToTotalMinutes(state.settings.breakMinMinutes, state.settings.breakMinSeconds);
        const lunchMin = convertToTotalMinutes(state.settings.lunchMinMinutes, state.settings.lunchMinSeconds);

        // Activity tracking
        const activityDetails = {};
        const seenOutOfSlotAgents = new Set();
        const processedAgents = new Set();
        let processedCount = 0;
        let highlightedCount = 0;

        for (let t = 0; t < len; t++) {
            const rows = tables[t].rows;
            const rowLen = rows.length;

            for (let i = 0; i < rowLen; i++) {
                const row = rows[i];
                const cells = row.cells;

                if (cells.length <= 2 || row.querySelectorAll('th').length > 5) continue;

                // Reset highlighting
                row.style.cssText = '';
                row.removeAttribute('data-highlighted');

                const agentLogin = cells[0]?.textContent?.trim();
                const agentLoginLower = agentLogin?.toLowerCase();
                // Skip if already processed this agent
                const normalizedLogin = agentLogin?.toLowerCase();
               if (processedAgents.has(normalizedLogin)) continue;
                   processedAgents.add(normalizedLogin);
                const activity = cells[2]?.textContent?.trim();
                const activityLower = activity?.toLowerCase();
                const durationText = cells[4]?.textContent?.trim();

                if (!agentLogin || agentLogin === 'Agent Login' || isActivitySummaryRow(agentLogin)) continue;



                // ✅ NEW VALIDATION CHECK #1: Skip if agent is not in active agents set
                if (!activeAgentsSet.has(agentLoginLower)) {
                    continue;
                }

                // ✅ NEW VALIDATION CHECK #2: Skip rows with time ranges (schedule data)
                if (isTimeRange(durationText)) {
                    continue;
                }

                // ✅ NEW VALIDATION CHECK #3: Skip if any cell contains time range format
                let hasTimeRangeInRow = false;
                for (let j = 0; j < cells.length; j++) {
                    if (isTimeRange(cells[j]?.textContent?.trim())) {
                        hasTimeRangeInRow = true;
                        break;
                    }
                }
                if (hasTimeRangeInRow) continue;
                 processedCount++;

               // PRIORITY 1: Check for out-of-slot breaks (PURPLE - HIGHEST PRIORITY)
if (state.isScheduleMonitoring &&
    (activityLower === 'break' || activityLower === 'lunch')) {

    const isOutOfSchedule = checkIfOutOfSchedule(agentLoginLower);

    if (isOutOfSchedule) {
        row.style.cssText = 'background-color:#e6b3ff;font-weight:bold;transition:background 0.3s';
        row.setAttribute('data-highlighted', 'purple');
        highlightedCount++;

        // Add to out-of-slot list (deduplicated)
        if (!seenOutOfSlotAgents.has(agentLoginLower)) {
            seenOutOfSlotAgents.add(agentLoginLower);
            const scheduleData = state.breakSchedules[agentLoginLower];
            const actualDuration = cells[4]?.textContent?.trim() || durationText || 'N/A';

state.outOfSlotAgents.push({
    login: agentLogin,
    manager: scheduleData?.manager || 'N/A',
    activity: activity,
    duration: actualDuration,
    break10: scheduleData?.break10 || 'N/A',
    break20: scheduleData?.break20 || 'N/A',
    break30: scheduleData?.break30 || 'N/A'
});
        }

        // ✅ ACTIVITY TRACKING MOVED HERE - BEFORE continue statement
if (activity && containsLetters(activity)) {
    const cleanDuration = cells[4]?.textContent?.trim() || '-';

    const durationMinutes = parseTimeToMinutes(cleanDuration);

            if (!activityDetails[activity]) {
                activityDetails[activity] = {
                    count: 0,
                    maxDuration: cleanDuration,
                    maxDurationMinutes: durationMinutes,
                    maxAgent: agentLogin
                };
            }

            activityDetails[activity].count++;

            if (durationMinutes > activityDetails[activity].maxDurationMinutes) {
                activityDetails[activity].maxDuration = cleanDuration;
                activityDetails[activity].maxDurationMinutes = durationMinutes;
                activityDetails[activity].maxAgent = agentLogin;
            }
        }

        continue; // Skip other highlighting checks
    }
}

// Activity tracking (for all OTHER agents - not out-of-slot)
if (activity && containsLetters(activity)) {
    const cleanDuration = cells[4]?.textContent?.trim() || '-';
    const durationMinutes = parseTimeToMinutes(cleanDuration);

    if (!activityDetails[activity]) {
        activityDetails[activity] = {
            count: 0,
            maxDuration: cleanDuration,
            maxDurationMinutes: durationMinutes,
            maxAgent: agentLogin
        };
    }

    activityDetails[activity].count++;

    if (durationMinutes > activityDetails[activity].maxDurationMinutes) {
        activityDetails[activity].maxDuration = cleanDuration;
        activityDetails[activity].maxDurationMinutes = durationMinutes;
        activityDetails[activity].maxAgent = agentLogin;
    }
}

                // PRIORITY 2-5: Duration-based highlighting
                const duration = parseTimeToMinutes(durationText);

                if (activityLower === 'available') {
                    if (state.settings.redEnabled && duration >= redMin) {
                        row.style.cssText = 'background-color:#ffcccc;font-weight:bold';
                        row.setAttribute('data-highlighted', 'red');
                        highlightedCount++;
                    } else if (state.settings.yellowEnabled && duration >= yellowMin && duration < yellowMax) {
                        row.style.cssText = 'background-color:#ffff99;font-weight:bold';
                        row.setAttribute('data-highlighted', 'yellow');
                        highlightedCount++;
                    }
                } else if (activityLower === 'break' && state.settings.blueEnabled && duration > breakMin) {
                    row.style.cssText = 'background-color:#cce5ff;font-weight:bold';
                    row.setAttribute('data-highlighted', 'blue');
                    highlightedCount++;
                } else if (activityLower === 'lunch' && state.settings.orangeEnabled && duration > lunchMin) {
                    row.style.cssText = 'background-color:#ffe5cc;font-weight:bold';
                    row.setAttribute('data-highlighted', 'orange');
                    highlightedCount++;
                }
            }
        }

        // Update activity summary cache
        state.cachedActivityCounts = {};
        for (const activity in activityDetails) {
            state.cachedActivityCounts[activity] = activityDetails[activity].count;
        }
        state.cachedActivityDetails = activityDetails;

        // Update UI components
        updateActivitySummary();
        updateOutOfSlotBox();

        // Debug info
        const debugInfo = document.getElementById('debugInfo');
        if (debugInfo) {
            debugInfo.textContent = `Check #${state.checkCount}: ${processedCount} agents, ${highlightedCount} highlighted, ${seenOutOfSlotAgents.size} out-of-slot`;
        }

        console.log(`✅ Unified Update #${state.checkCount}: ${processedCount} agents, ${highlightedCount} highlighted, ${seenOutOfSlotAgents.size} out-of-slot`);
    };

    const debouncedUpdate = () => {
        clearTimeout(state.updateTimeout);
        state.updateTimeout = setTimeout(unifiedUpdate, CONFIG.DEBOUNCE);
    };

    // ==================== ACTIVITY SUMMARY ====================
    const updateActivitySummary = () => {
        if (!state.summaryBox || state.summaryBox.style.display === 'none' || state.isUpdating) return;

        const now = Date.now();
        if (now - state.lastUpdateTime < CONFIG.UPDATE_THROTTLE) return;

        state.isUpdating = true;
        state.lastUpdateTime = now;

        const activityDetails = state.cachedActivityDetails;
        const tableBody = document.getElementById('activity-summary-tbody');
        if (!tableBody) {
            state.isUpdating = false;
            return;
        }

        const sortedActivities = Object.keys(activityDetails).sort();
        let totalCount = 0;
        const rows = [];

        for (let i = 0; i < sortedActivities.length; i++) {
            const activity = sortedActivities[i];
            const details = activityDetails[activity];
            totalCount += details.count;
            rows.push(`<tr><td style="padding:8px;border-bottom:1px solid #ddd">${activity}</td><td style="padding:8px;border-bottom:1px solid #ddd;text-align:center;font-weight:bold">${details.count}</td><td style="padding:8px;border-bottom:1px solid #ddd;text-align:center">${details.maxDuration}</td><td style="padding:8px;border-bottom:1px solid #ddd;text-align:center">${details.maxAgent}</td></tr>`);
        }
        rows.push(`<tr><td style="padding:8px;font-weight:bold;background:#f0f0f0">Total</td><td style="padding:8px;font-weight:bold;background:#f0f0f0;text-align:center">${totalCount}</td><td style="padding:8px;font-weight:bold;background:#f0f0f0;text-align:center">-</td><td style="padding:8px;font-weight:bold;background:#f0f0f0;text-align:center">-</td></tr>`);

        tableBody.innerHTML = rows.join('');
        state.isUpdating = false;
    };

    const createActivitySummaryBox = () => {
        state.summaryBox = document.createElement('div');
        state.summaryBox.id = 'activity-summary-box';
        state.summaryBox.style.cssText = 'position:fixed;top:10px;left:10px;z-index:9999;background:#fff;padding:15px;border-radius:8px;box-shadow:0 4px 15px rgba(0,0,0,0.3);width:600px;height:500px;cursor:move;font-family:Arial,sans-serif;overflow:hidden';
        state.summaryBox.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;border-bottom:2px solid #FF9900;padding-bottom:8px"><h3 style="margin:0;color:#232F3E;font-size:16px">📊 Activity Summary</h3><div><button id="download-summary-btn" style="background:#FF9900;border:none;font-size:14px;cursor:pointer;color:#fff;padding:4px 8px;border-radius:4px;margin-right:8px;font-weight:bold" title="Download CSV">⬇</button><button id="close-summary-btn" style="background:none;border:none;font-size:20px;cursor:pointer;color:#666;padding:0;width:24px;height:24px">×</button></div></div><div style="max-height:calc(100% - 100px);overflow-y:auto"><table style="width:100%;border-collapse:collapse;font-size:14px"><thead><tr style="background:#f8f8f8"><th style="padding:8px;border-bottom:2px solid #ddd;text-align:left;font-weight:bold">Activity</th><th style="padding:8px;border-bottom:2px solid #ddd;text-align:center;font-weight:bold">HC</th><th style="padding:8px;border-bottom:2px solid #ddd;text-align:center;font-weight:bold">Highest Duration</th><th style="padding:8px;border-bottom:2px solid #ddd;text-align:center;font-weight:bold">Agent</th></tr></thead><tbody id="activity-summary-tbody"><tr><td colspan="4" style="padding:20px;text-align:center;color:#999">Loading...</td></tr></tbody></table></div><div style="margin-top:10px;padding-top:8px;border-top:1px solid #ddd;font-size:11px;color:#666;text-align:center">Updates every 5 seconds • Drag to move</div>';
        document.body.appendChild(state.summaryBox);

        makeDraggable(state.summaryBox);
        makeResizable(state.summaryBox);
        document.getElementById('download-summary-btn').onclick = (e) => {
            e.stopPropagation();
            downloadActivityData();
        };
        document.getElementById('close-summary-btn').onclick = (e) => {
            e.stopPropagation();
            state.summaryBox.style.display = 'none';
        };

        updateActivitySummary();
    };

    const makeDraggable = (element) => {
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        element.onmousedown = (e) => {
    if (e.target.id === 'close-summary-btn' || e.target.id === 'download-summary-btn' ||
        e.target.id === 'minimizeOutOfSlotBtn' || e.target.id === 'downloadCSVBtn' ||
        e.target.id === 'minimizeBtn' || e.target.closest('input') || e.target.closest('button') ||
        e.target.closest('td') || e.target.closest('table')) return;
    e.preventDefault();
            pos3 = e.clientX;
            pos4 = e.clientY;
            document.onmouseup = closeDrag;
            document.onmousemove = drag;
        };

        const drag = (e) => {
            e.preventDefault();
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;
            element.style.top = (element.offsetTop - pos2) + "px";
            element.style.left = (element.offsetLeft - pos1) + "px";
            element.style.right = 'auto';
            element.style.bottom = 'auto';
        };

        const closeDrag = () => {
            document.onmouseup = null;
            document.onmousemove = null;
        };
    };
const closeDrag = () => {
        document.onmouseup = null;
        document.onmousemove = null;
    };
  // ← This is the END of makeDraggable function

// ← ADD THE NEW FUNCTION HERE (after this blank line)
const makeResizable = (element) => {
    const resizer = document.createElement('div');
    resizer.style.cssText = 'position:absolute;right:0;bottom:0;width:15px;height:15px;background:linear-gradient(135deg,transparent 50%,#666 50%);cursor:nwse-resize;z-index:10';
    element.appendChild(resizer);
    element.style.position = 'fixed';

    let startX, startY, startWidth, startHeight;

    resizer.onmousedown = (e) => {
        e.preventDefault();
        e.stopPropagation();
        startX = e.clientX;
        startY = e.clientY;
        startWidth = parseInt(getComputedStyle(element).width, 10);
        startHeight = parseInt(getComputedStyle(element).height, 10);

        document.onmousemove = resize;
        document.onmouseup = stopResize;
    };

    const resize = (e) => {
        const width = startWidth + (e.clientX - startX);
        const height = startHeight + (e.clientY - startY);
        element.style.width = Math.max(400, width) + 'px';
        element.style.height = Math.max(300, height) + 'px';
    };

    const stopResize = () => {
        document.onmousemove = null;
        document.onmouseup = null;
    };
};
    // ==================== DATA EXTRACTION & CSV EXPORT ====================
    const extractTableData = (highlightedOnly) => {
        const data = [];
        const tables = getTables();
        const headers = ['Agent Login','Channels','Activity','Next activity','Duration','Agent Hierarchy','Routing Profile','Capacity','Active','Availability','Contact State','Duration_Contact','Queue','Avg ACW','Agent non-response','Handled in','Handled out','AHT','Occupancy'];
        if (highlightedOnly) headers.push('Highlight Status');

        for (let t = 0; t < tables.length; t++) {
            const rows = tables[t].rows;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                if (row.querySelectorAll('th').length > 5) continue;
                if (highlightedOnly && !row.hasAttribute('data-highlighted')) continue;

                const cells = row.cells;
                if (cells.length === 0) continue;

                const agentLogin = cells[0]?.textContent?.trim();

                // Skip if no agent login, header row, or activity summary row
if (!agentLogin || agentLogin === 'Agent Login' || isActivitySummaryRow(agentLogin)) continue;

// ✅ NEW: Skip rows from Out-of-Slot Breaks table (uppercase logins)
if (agentLogin === agentLogin.toUpperCase() && agentLogin.length > 0) continue;

// ✅ NEW: Skip rows where Channels column doesn't contain "Voice"
const channels = cells?.textContent?.trim();
if (channels && channels !== 'Voice') continue;

                const rowData = {};
                for (let j = 0; j < 19 && j < cells.length; j++) {
                    rowData[headers[j]] = cells[j]?.textContent?.trim() || '';
                }
                if (highlightedOnly) {
                    rowData['Highlight Status'] = row.getAttribute('data-highlighted');
                }
                data.push(rowData);
            }
        }
        return data;
    };

    const convertToCSV = (data) => {
    if (!data.length) return '';
    const headers = Object.keys(data[0]);
    let csv = headers.map(h => `"${h}"`).join(',') + '\r';
    for (let i = 0; i < data.length; i++) {
        csv += headers.map(h => `"${(data[i][h] || '').toString().replace(/"/g, '""')}"`).join(',') + '\r';
    }
    return csv;
};

    const getActivitySummaryCSV = () => {
    const activityDetails = state.cachedActivityDetails;
    const sortedActivities = Object.keys(activityDetails).sort();
    let csv = '\r\r\r"=== ACTIVITY SUMMARY ==="\r"Activity","HC","Highest Duration","Agent"\r';
    let totalCount = 0;
    for (let i = 0; i < sortedActivities.length; i++) {
        const activity = sortedActivities[i];
        const details = activityDetails[activity];
        totalCount += details.count;
        csv += `"${activity}",${details.count},"${details.maxDuration}","${details.maxAgent}"\r
`;
    }
    csv += `"Total",${totalCount},"",""\r
`;
    return csv;
};

    const downloadData = (highlightedOnly) => {
        const data = extractTableData(highlightedOnly);
        if (!data.length) {
            alert(highlightedOnly ? 'No highlighted agents found.' : 'No data found.');
            return;
        }

        const timestamp = getISTTimestamp();
        const agentCSV = convertToCSV(data);
        const activityCSV = getActivitySummaryCSV();
        const combinedCSV = agentCSV + activityCSV;

        const blob = new Blob([combinedCSV], {type: 'text/csv;charset=utf-8;'});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `connect_${highlightedOnly ? 'highlighted' : 'realtime'}_${timestamp}.csv`;
        link.click();
        URL.revokeObjectURL(link.href);
        alert(`Downloaded ${data.length} agent records with Activity Summary!`);
    };

    const downloadActivityData = () => {
    const activityDetails = state.cachedActivityDetails;
    const sortedActivities = Object.keys(activityDetails).sort();
    let csvContent = 'Activity,HC,Highest Duration,Agent\r';
    let totalCount = 0;

    for (let i = 0; i < sortedActivities.length; i++) {
        const activity = sortedActivities[i];
        const details = activityDetails[activity];
        totalCount += details.count;
        csvContent += `${activity},${details.count},${details.maxDuration},${details.maxAgent}\r
`;
    }
    csvContent += `Total,${totalCount},,\r
`;

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        const timestamp = getISTTimestamp();
        link.setAttribute('href', url);
        link.setAttribute('download', `Activity_Summary_${timestamp}.csv`);
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    };

    // END OF PART 2
  // PART 3 - OUT-OF-SLOT BREAKS UI & SCHEDULE MANAGEMENT
    // This continues from Part 2

    // ==================== OUT-OF-SLOT BREAKS DISPLAY ====================
    const updateOutOfSlotBox = () => {
        const listDiv = document.getElementById('outOfSlotList');
        const countSpan = document.getElementById('outOfSlotCount');
        const bufferTimeSpan = document.getElementById('displayBufferTime');

        if (!listDiv || !countSpan || !bufferTimeSpan) return;

        bufferTimeSpan.textContent = `${String(state.bufferMinutes).padStart(2, '0')}:${String(state.bufferSeconds).padStart(2, '0')}`;
// ✅ ADD THIS NEW CODE:
const currentTimeSpan = document.getElementById('displayCurrentTime');
if (currentTimeSpan) {
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    currentTimeSpan.textContent = `${hours}:${minutes}:${seconds}`;
}
        if (state.outOfSlotAgents.length === 0) {
            listDiv.innerHTML = `
                <div style="text-align: center; color: #666; padding: 20px;">
                    No out-of-slot breaks detected
                </div>
            `;
            countSpan.textContent = '0';
            return;
        }

        let html = `
            <table style="width: 100%; border-collapse: collapse; font-size: 11px; user-select: text;">
                <thead>
                    <tr style="background: #f8f9fa; border-bottom: 2px solid #d13212;">
                        <th style="padding: 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;">Login</th>
                        <th style="padding: 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;">Manager</th>
                        <th style="padding: 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;">Activity</th>
                        <th style="padding: 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;">Duration</th>
                        <th style="padding: 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;">Break (10 Mins)</th>
                        <th style="padding: 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;">Break (20 Mins)</th>
                        <th style="padding: 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;">Break (30 Mins)</th>
                    </tr>
                </thead>
                <tbody>
        `;

        state.outOfSlotAgents.forEach((agent, index) => {
            const rowBg = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
            html += `
                <tr style="background: ${rowBg};">
                    <td class="copy-login" data-login="${agent.login.toUpperCase()}" style="padding: 6px; border: 1px solid #ddd; font-weight: bold; color: #d13212; cursor: pointer;" title="Click to copy">${agent.login.toUpperCase()}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${agent.manager}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${agent.activity.charAt(0).toUpperCase() + agent.activity.slice(1)}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${agent.duration}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${agent.break10}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${agent.break20}</td>
                    <td style="padding: 6px; border: 1px solid #ddd;">${agent.break30}</td>
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
        `;

        listDiv.innerHTML = html;
        countSpan.textContent = state.outOfSlotAgents.length;
    };

    const createOutOfSlotBox = () => {
        state.outOfSlotBox = document.createElement('div');
        state.outOfSlotBox.id = 'outOfSlotBox';
        state.outOfSlotBox.style.cssText = `
    position: fixed;
    bottom: 10px;
    right: 10px;
    width: 700px;
    max-height: 600px;
    background: white;
    border: 2px solid #d13212;
    border-radius: 8px;
    padding: 0;
    z-index: 9998;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    font-family: Arial, sans-serif;
    cursor: move;
    user-select: text;
    overflow: hidden;
`;

        state.outOfSlotBox.innerHTML = `
            <div id="outOfSlotHeader" style="background: #d13212; color: white; padding: 10px 15px; border-radius: 6px 6px 0 0; cursor: move; display: flex; justify-content: space-between; align-items: center;">
                <h3 style="margin: 0; font-size: 16px;">🚨 Out-of-Slot Breaks</h3>
                <div style="display: flex; gap: 10px;">
                    <button id="minimizeOutOfSlotBtn" style="background: #ff9900; color: white; border: none; border-radius: 3px; padding: 2px 8px; cursor: pointer; font-size: 14px; font-weight: bold;">−</button>
                </div>
            </div>
            <div id="outOfSlotContent" style="padding: 15px; height: calc(100% - 50px); overflow-y: auto;">
                <div style="margin-bottom: 10px; padding: 8px; background: #fff3cd; border-left: 4px solid #ff9900; font-size: 12px; display: flex; justify-content: space-between; align-items: center;">
                     <div>
                          <div><strong>Buffer Time:</strong> <span id="displayBufferTime">00:00</span></div>
                          <div style="margin-top: 5px;"><strong>Current Time:</strong> <span id="displayCurrentTime">--:--:--</span></div>
                    </div>
                    <button id="downloadCSVBtn" style="padding: 5px 10px; background: #232f3e; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 11px; font-weight: bold;">📥 Download CSV</button>
                </div>
                <div id="outOfSlotList" style="font-size: 12px; overflow-x: auto; user-select: text;">
                    <div style="text-align: center; color: #666; padding: 20px;">
                        No out-of-slot breaks detected
                    </div>
                </div>
                <div id="outOfSlotSummary" style="margin-top: 15px; padding: 10px; background: #f8f9fa; border-top: 2px solid #d13212; font-weight: bold; font-size: 13px; text-align: center;">
                    Total: <span id="outOfSlotCount">0</span> agents
                </div>
            </div>
        `;

        document.body.appendChild(state.outOfSlotBox);
        makeDraggable(state.outOfSlotBox);
        makeResizable(state.outOfSlotBox);
        document.getElementById('minimizeOutOfSlotBtn').addEventListener('click', toggleOutOfSlotBox);
        document.getElementById('downloadCSVBtn').addEventListener('click', downloadOutOfSlotCSV);
        // ✅ NEW: Add click-to-copy functionality for agent logins
document.getElementById('outOfSlotList').addEventListener('click', (e) => {
    const cell = e.target.closest('.copy-login');
    if (cell) {
        e.stopPropagation();
        const login = cell.getAttribute('data-login');
        navigator.clipboard.writeText(login).then(() => {
            alert(`✅ Copied: ${login}`);
        }).catch(err => {
            console.error('Failed to copy:', err);
            alert('❌ Failed to copy to clipboard');
        });
    }
});
    };

    const toggleOutOfSlotBox = () => {
    const content = document.getElementById('outOfSlotContent');
    const minimizeBtn = document.getElementById('outOfSlotMinimizeBtn');
    const box = document.getElementById('outOfSlotBox');

    state.isOutOfSlotBoxMinimized = !state.isOutOfSlotBoxMinimized;

    if (state.isOutOfSlotBoxMinimized) {
        content.style.display = 'none';
        minimizeBtn.textContent = '+';
    } else {
        content.style.display = 'block';
        minimizeBtn.textContent = '−';
    }
};

    const downloadOutOfSlotCSV = () => {
    if (state.outOfSlotAgents.length === 0) {
        alert('❌ No out-of-slot breaks to download');
        return;
    }

    const headers = ['Login', 'Manager', 'Activity', 'Duration', 'Break (10 Mins)', 'Break (20 Mins)', 'Break (30 Mins)'];
    let csvContent = headers.join(',') + '\r';

    state.outOfSlotAgents.forEach(agent => {
        const row = [
            agent.login.toLowerCase(),
            agent.manager,
            agent.activity.charAt(0).toUpperCase() + agent.activity.slice(1),
            agent.duration,
            agent.break10,
            agent.break20,
            agent.break30
        ];
        csvContent += row.join(',') + '\r';
    });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        const timestamp = getISTTimestamp();

        link.setAttribute('href', url);
        link.setAttribute('download', `Out-of-Slot-Breaks_${timestamp}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

        alert('✅ CSV downloaded successfully');
    };

    // ==================== SCHEDULE CONTROL PANEL ====================
    const createScheduleControlPanel = () => {
        const panel = document.createElement('div');
        panel.id = 'scheduleControlPanel';
        panel.style.cssText = `
            position: fixed;
            top: 10px;
            right: 10px;
            width: 350px;
            background: white;
            border: 2px solid #232f3e;
            border-radius: 8px;
            padding: 0;
            z-index: 9997;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            font-family: Arial, sans-serif;
            cursor: move;
        `;

        panel.innerHTML = `
            <div id="panelHeader" style="background: #232f3e; color: white; padding: 10px 15px; border-radius: 6px 6px 0 0; cursor: move; display: flex; justify-content: space-between; align-items: center;">
                <h3 style="margin: 0; font-size: 16px;">🟣 Break Schedule Monitor</h3>
                <div style="display: flex; gap: 10px;">
                    <button id="minimizeBtn" style="background: #ff9900; color: white; border: none; border-radius: 3px; padding: 2px 8px; cursor: pointer; font-size: 14px; font-weight: bold;">−</button>
                </div>
            </div>
            <div id="panelContent" style="padding: 15px;">
                <div style="margin-bottom: 15px;">
                    <div style="font-size: 12px; color: #666; margin-bottom: 10px;">
                        <div id="scheduleStatus">No schedule loaded</div>
                    </div>
                </div>

                <div style="margin-bottom: 15px;">
                    <label style="display: block; margin-bottom: 5px; font-weight: bold; font-size: 13px;">Upload Schedule (.xlsx):</label>
                    <input type="file" id="scheduleFileInput" accept=".xlsx" style="width: 100%; padding: 5px; font-size: 12px; cursor: pointer;">
                    <div id="uploadMessage" style="margin-top: 5px; font-size: 12px; min-height: 18px;"></div>
                </div>

                <div style="margin-bottom: 15px;">
                    <label style="display: block; margin-bottom: 5px; font-weight: bold; font-size: 13px;">Buffer Time (MM:SS):</label>
                    <div style="display: flex; gap: 10px; align-items: center;">
                        <input type="number" id="bufferMinutes" min="0" value="0"
                               style="width: 60px; padding: 5px; font-size: 12px; cursor: text;" placeholder="MM">
                        <span style="font-weight: bold;">:</span>
                        <input type="number" id="bufferSeconds" min="0" max="59" value="0"
                               style="width: 60px; padding: 5px; font-size: 12px; cursor: text;" placeholder="SS">
                        <button id="applyBuffer" style="padding: 5px 10px; background: #ff9900; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">Apply</button>
                    </div>
                    <div style="margin-top: 5px; font-size: 11px; color: #666;">Current: <span id="currentBuffer">00:00</span></div>
                </div>

                <div style="margin-bottom: 10px;">
                    <button id="toggleScheduleMonitoring" style="width: 100%; padding: 8px; background: #232f3e; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: bold;">
                        Start Schedule Monitoring
                    </button>
                </div>

                <div style="font-size: 11px; color: #666; border-top: 1px solid #ddd; padding-top: 10px;">
                    <div style="margin-bottom: 3px;">🟣 Purple = Out of schedule break (Highest Priority)</div>
                    <div>Updates every 5 seconds</div>
                    <div id="debugInfo" style="margin-top: 5px; color: #999; font-size: 10px;"></div>
                </div>
            </div>
        `;

        document.body.appendChild(panel);
        makeDraggable(panel);

        document.getElementById('scheduleFileInput').addEventListener('change', handleFileUpload);
        document.getElementById('applyBuffer').addEventListener('click', applyBufferTime);
        document.getElementById('toggleScheduleMonitoring').addEventListener('click', toggleScheduleMonitoring);
        document.getElementById('minimizeBtn').addEventListener('click', toggleMinimize);
    };

    const toggleMinimize = () => {
        const content = document.getElementById('panelContent');
        const minimizeBtn = document.getElementById('minimizeBtn');
        state.isPanelMinimized = !state.isPanelMinimized;
        if (state.isPanelMinimized) {
            content.style.display = 'none';
            minimizeBtn.textContent = '+';
        } else {
            content.style.display = 'block';
            minimizeBtn.textContent = '−';
        }
    };

    async function handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        if (!file.name.endsWith('.xlsx')) {
            showScheduleMessage('❌ Error: Please upload an Excel file (.xlsx)', 'error');
            return;
        }
        try {
            showScheduleMessage('⏳ Processing file...', 'info');
            if (typeof DecompressionStream === 'undefined') {
                showScheduleMessage('❌ Error: Your browser does not support DecompressionStream. Please use Chrome/Edge.', 'error');
                return;
            }
            const jsonData = await parseXLSXFile(file);
            if (!jsonData || jsonData.length === 0) {
                showScheduleMessage('❌ Error: Could not read file.', 'error');
                return;
            }
            parseSchedule(jsonData);
        } catch (error) {
            console.error('File upload error:', error);
            showScheduleMessage('❌ Error: ' + error.message, 'error');
        }
    }

    function applyBufferTime() {
        const minutes = parseInt(document.getElementById('bufferMinutes').value) || 0;
        const seconds = parseInt(document.getElementById('bufferSeconds').value) || 0;
        state.bufferMinutes = Math.max(0, minutes);
        state.bufferSeconds = Math.max(0, Math.min(59, seconds));
        document.getElementById('currentBuffer').textContent =
            `${String(state.bufferMinutes).padStart(2, '0')}:${String(state.bufferSeconds).padStart(2, '0')}`;
        showScheduleMessage(`✅ Buffer time updated`, 'success');

        if (state.outOfSlotAgents.length > 0) {
            updateOutOfSlotBox();
        }
    }

    function toggleScheduleMonitoring() {
        const button = document.getElementById('toggleScheduleMonitoring');
        if (Object.keys(state.breakSchedules).length === 0) {
            showScheduleMessage('❌ Please upload a schedule file first', 'error');
            return;
        }
        state.isScheduleMonitoring = !state.isScheduleMonitoring;
        if (state.isScheduleMonitoring) {
            button.textContent = 'Stop Schedule Monitoring';
            button.style.background = '#d13212';
            showScheduleMessage('✅ Schedule monitoring started', 'success');
            console.log('🟣 SCHEDULE MONITORING STARTED');
        } else {
            button.textContent = 'Start Schedule Monitoring';
            button.style.background = '#232f3e';
            showScheduleMessage('⏸️ Schedule monitoring stopped', 'info');
            console.log('⏸️ SCHEDULE MONITORING STOPPED');
        }
    }

    // END OF PART 3
// PART 4 - SETTINGS DIALOG, UI BUTTONS & INITIALIZATION
    // This continues from Part 3

    // ==================== SETTINGS DIALOG ====================
    const showSettingsDialog = () => {
        document.getElementById('highlightSettingsDialog')?.remove();
        document.getElementById('highlightSettingsBackdrop')?.remove();

        const backdrop = document.createElement('div');
        backdrop.id = 'highlightSettingsBackdrop';
        backdrop.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:10000';
        backdrop.onclick = closeDialog;

        const dialog = document.createElement('div');
        dialog.id = 'highlightSettingsDialog';
        dialog.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:10001;background:#fff;padding:25px;border-radius:10px;box-shadow:0 4px 20px rgba(0,0,0,0.3);min-width:550px;max-height:80vh;overflow-y:auto';
        dialog.innerHTML = `
            <h2 style="margin-top:0;color:#333">Duration-Based Highlight Settings</h2>

            <!-- Yellow Highlighting -->
            <div style="margin-bottom:20px;padding:15px;background:#fffef0;border-radius:5px;border-left:4px solid #ffeb3b">
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
                    <label style="font-weight:bold;font-size:16px">🟡 Yellow (Available)</label>
                    <label style="display:flex;align-items:center;cursor:pointer">
                        <input type="checkbox" id="yellowEnabled" ${state.settings.yellowEnabled ? 'checked' : ''} style="width:20px;height:20px;cursor:pointer">
                        <span style="margin-left:8px;font-size:14px">Enable</span>
                    </label>
                </div>
                <div style="display:flex;align-items:center;gap:10px;margin-top:10px">
                    <span>≥</span>
                    <input type="number" id="yellowMinMinutes" value="${state.settings.yellowMinMinutes}" min="0" step="1" style="width:60px;padding:5px">
                    <span>min</span>
                    <input type="number" id="yellowMinSeconds" value="${state.settings.yellowMinSeconds}" min="0" max="59" step="1" style="width:60px;padding:5px">
                    <span>sec</span>
                    <span style="margin:0 5px">&lt;</span>
                    <input type="number" id="yellowMaxMinutes" value="${state.settings.yellowMaxMinutes}" min="0" step="1" style="width:60px;padding:5px">
                    <span>min</span>
                    <input type="number" id="yellowMaxSeconds" value="${state.settings.yellowMaxSeconds}" min="0" max="59" step="1" style="width:60px;padding:5px">
                    <span>sec</span>
                </div>
            </div>

            <!-- Red Highlighting -->
            <div style="margin-bottom:20px;padding:15px;background:#fff0f0;border-radius:5px;border-left:4px solid #f44336">
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
                    <label style="font-weight:bold;font-size:16px">🔴 Red (Available)</label>
                    <label style="display:flex;align-items:center;cursor:pointer">
                        <input type="checkbox" id="redEnabled" ${state.settings.redEnabled ? 'checked' : ''} style="width:20px;height:20px;cursor:pointer">
                        <span style="margin-left:8px;font-size:14px">Enable</span>
                    </label>
                </div>
                <div style="display:flex;align-items:center;gap:10px;margin-top:10px">
                    <span>≥</span>
                    <input type="number" id="redMinMinutes" value="${state.settings.redMinMinutes}" min="0" step="1" style="width:60px;padding:5px">
                    <span>min</span>
                    <input type="number" id="redMinSeconds" value="${state.settings.redMinSeconds}" min="0" max="59" step="1" style="width:60px;padding:5px">
                    <span>sec</span>
                </div>
            </div>

            <!-- Blue Highlighting -->
            <div style="margin-bottom:20px;padding:15px;background:#f0f8ff;border-radius:5px;border-left:4px solid #2196F3">
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
                    <label style="font-weight:bold;font-size:16px">🔵 Blue (Break)</label>
                    <label style="display:flex;align-items:center;cursor:pointer">
                        <input type="checkbox" id="blueEnabled" ${state.settings.blueEnabled ? 'checked' : ''} style="width:20px;height:20px;cursor:pointer">
                        <span style="margin-left:8px;font-size:14px">Enable</span>
                    </label>
                </div>
                <div style="display:flex;align-items:center;gap:10px;margin-top:10px">
                    <span>&gt;</span>
                    <input type="number" id="breakMinMinutes" value="${state.settings.breakMinMinutes}" min="0" step="1" style="width:60px;padding:5px">
                    <span>min</span>
                    <input type="number" id="breakMinSeconds" value="${state.settings.breakMinSeconds}" min="0" max="59" step="1" style="width:60px;padding:5px">
                    <span>sec</span>
                </div>
            </div>

            <!-- Orange Highlighting -->
            <div style="margin-bottom:20px;padding:15px;background:#fff5f0;border-radius:5px;border-left:4px solid #ff9800">
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
                    <label style="font-weight:bold;font-size:16px">🟠 Orange (Lunch)</label>
                    <label style="display:flex;align-items:center;cursor:pointer">
                        <input type="checkbox" id="orangeEnabled" ${state.settings.orangeEnabled ? 'checked' : ''} style="width:20px;height:20px;cursor:pointer">
                        <span style="margin-left:8px;font-size:14px">Enable</span>
                    </label>
                </div>
                <div style="display:flex;align-items:center;gap:10px;margin-top:10px">
                    <span>&gt;</span>
                    <input type="number" id="lunchMinMinutes" value="${state.settings.lunchMinMinutes}" min="0" step="1" style="width:60px;padding:5px">
                    <span>min</span>
                    <input type="number" id="lunchMinSeconds" value="${state.settings.lunchMinSeconds}" min="0" max="59" step="1" style="width:60px;padding:5px">
                    <span>sec</span>
                </div>
            </div>

            <div style="padding:10px;background:#f0f0ff;border-radius:5px;border-left:4px solid #9c27b0;margin-bottom:20px">
                <div style="font-weight:bold;font-size:14px;margin-bottom:5px">🟣 Purple (Out-of-Slot Break)</div>
                <div style="font-size:12px;color:#666">Configured in Break Schedule Monitor panel. Purple highlighting has the highest priority and will override all other colors.</div>
            </div>

            <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:20px">
                <button id="cancelBtn" style="padding:8px 20px;background:#ccc;border:none;border-radius:5px;cursor:pointer">Cancel</button>
                <button id="saveBtn" style="padding:8px 20px;background:#4CAF50;color:#fff;border:none;border-radius:5px;cursor:pointer;font-weight:bold">Save</button>
            </div>
        `;

        document.body.appendChild(backdrop);
        document.body.appendChild(dialog);

        document.getElementById('saveBtn').onclick = saveSettings;
        document.getElementById('cancelBtn').onclick = closeDialog;
    };

    const saveSettings = () => {
        const yellowEnabled = document.getElementById('yellowEnabled').checked;
        const redEnabled = document.getElementById('redEnabled').checked;
        const blueEnabled = document.getElementById('blueEnabled').checked;
        const orangeEnabled = document.getElementById('orangeEnabled').checked;

        const yellowMinMinutes = parseInt(document.getElementById('yellowMinMinutes').value);
        const yellowMinSeconds = parseInt(document.getElementById('yellowMinSeconds').value);
        const yellowMaxMinutes = parseInt(document.getElementById('yellowMaxMinutes').value);
        const yellowMaxSeconds = parseInt(document.getElementById('yellowMaxSeconds').value);
        const redMinMinutes = parseInt(document.getElementById('redMinMinutes').value);
        const redMinSeconds = parseInt(document.getElementById('redMinSeconds').value);
        const breakMinMinutes = parseInt(document.getElementById('breakMinMinutes').value);
        const breakMinSeconds = parseInt(document.getElementById('breakMinSeconds').value);
        const lunchMinMinutes = parseInt(document.getElementById('lunchMinMinutes').value);
        const lunchMinSeconds = parseInt(document.getElementById('lunchMinSeconds').value);

        // Validation
        const yellowMinTotal = convertToTotalMinutes(yellowMinMinutes, yellowMinSeconds);
        const yellowMaxTotal = convertToTotalMinutes(yellowMaxMinutes, yellowMaxSeconds);
        const redMinTotal = convertToTotalMinutes(redMinMinutes, redMinSeconds);

        if (yellowEnabled && yellowMinTotal >= yellowMaxTotal) {
            alert('Yellow minimum must be less than yellow maximum!');
            return;
        }

        if (yellowEnabled && redEnabled && redMinTotal < yellowMaxTotal) {
            alert('Red minimum should be greater than or equal to yellow maximum!');
            return;
        }

        // Save all settings
        state.settings = {
            yellowEnabled,
            redEnabled,
            blueEnabled,
            orangeEnabled,
            yellowMinMinutes,
            yellowMinSeconds,
            yellowMaxMinutes,
            yellowMaxSeconds,
            redMinMinutes,
            redMinSeconds,
            breakMinMinutes,
            breakMinSeconds,
            lunchMinMinutes,
            lunchMinSeconds
        };

        // Save to GM storage
        GM_setValue('yellowEnabled', yellowEnabled);
        GM_setValue('redEnabled', redEnabled);
        GM_setValue('blueEnabled', blueEnabled);
        GM_setValue('orangeEnabled', orangeEnabled);
        GM_setValue('yellowMinMinutes', yellowMinMinutes);
        GM_setValue('yellowMinSeconds', yellowMinSeconds);
        GM_setValue('yellowMaxMinutes', yellowMaxMinutes);
        GM_setValue('yellowMaxSeconds', yellowMaxSeconds);
        GM_setValue('redMinMinutes', redMinMinutes);
        GM_setValue('redMinSeconds', redMinSeconds);
        GM_setValue('breakMinMinutes', breakMinMinutes);
        GM_setValue('breakMinSeconds', breakMinSeconds);
        GM_setValue('lunchMinMinutes', lunchMinMinutes);
        GM_setValue('lunchMinSeconds', lunchMinSeconds);

        unifiedUpdate();
        closeDialog();

        const notif = document.createElement('div');
        notif.textContent = 'Settings saved!';
        notif.style.cssText = 'position:fixed;top:70px;right:10px;z-index:10002;background:#4CAF50;color:#fff;padding:15px 20px;border-radius:5px;box-shadow:0 2px 5px rgba(0,0,0,0.3);font-weight:bold';
        document.body.appendChild(notif);
        setTimeout(() => notif.remove(), 3000);
    };

    const closeDialog = () => {
        document.getElementById('highlightSettingsDialog')?.remove();
        document.getElementById('highlightSettingsBackdrop')?.remove();
    };

    // ==================== UI BUTTONS ====================
    const createButton = (text, right, bg, hoverBg, onClick) => {
        const btn = document.createElement('button');
        btn.innerHTML = text;
        btn.style.cssText = `position:fixed;top:10px;right:${right}px;z-index:10000;padding:10px 15px;background:${bg};color:#fff;border:none;border-radius:5px;cursor:pointer;font-size:14px;font-weight:bold;box-shadow:0 2px 5px rgba(0,0,0,0.3)`;
        btn.onmouseover = function() { this.style.background = hoverBg; };
        btn.onmouseout = function() { this.style.background = bg; };
        btn.onclick = onClick;
        document.body.appendChild(btn);
    };

    const createToggleActivityButton = () => {
        const button = document.createElement('button');
        button.innerHTML = '📊 Activity';
        button.id = 'toggle-summary-btn';
        button.style.cssText = 'position:fixed;top:10px;right:540px;z-index:10000;padding:10px 15px;background:#2196F3;color:#fff;border:none;border-radius:5px;cursor:pointer;font-size:14px;font-weight:bold;box-shadow:0 2px 5px rgba(0,0,0,0.3)';
        button.onmouseover = function() { this.style.background = '#1976D2'; };
        button.onmouseout = function() { this.style.background = '#2196F3'; };
        button.onclick = () => {
            if (state.summaryBox) {
                state.summaryBox.style.display = state.summaryBox.style.display === 'none' ? 'block' : 'none';
            }
        };
        document.body.appendChild(button);
    };

    // ==================== INITIALIZATION ====================
    const init = () => {
        console.log('🚀 Amazon Connect Ultimate Monitoring Suite v6.0 initialized');
        console.log('📋 Features: Duration Highlighting + Activity Tracking + Break Schedule Compliance');
        console.log('🔄 Update Interval: 5 seconds (unified)');
        console.log('🎨 Highlight Priority: Purple > Red > Yellow > Blue > Orange');

        // Create UI components
        createButton('⚙️ Highlight Settings', 380, '#4CAF50', '#45a049', showSettingsDialog);
        createButton('⭐ Download Highlighted', 180, '#9C27B0', '#7B1FA2', () => downloadData(true));
        createButton('📥 Download All', 10, '#FF9900', '#EC7211', () => downloadData(false));
        createToggleActivityButton();
        createActivitySummaryBox();
        createOutOfSlotBox();
        createScheduleControlPanel();

        // Initial update
        unifiedUpdate();

        // Set up MutationObserver for dynamic updates
        const tables = getTables();
        if (tables.length > 0) {
            state.observer = new MutationObserver(debouncedUpdate);

            for (let i = 0; i < tables.length; i++) {
                state.observer.observe(tables[i], {childList: true, subtree: false});
            }
        }

        // Set up unified interval (5 seconds)
        setInterval(unifiedUpdate, CONFIG.UNIFIED_INTERVAL);

        console.log('✅ Initialization complete - monitoring active');
    };

    // ==================== CLEANUP ====================
    window.addEventListener('beforeunload', () => {
        if (state.observer) state.observer.disconnect();
        clearTimeout(state.updateTimeout);
        console.log('🛑 Amazon Connect Ultimate Monitoring Suite stopped');
    });

    // ==================== START ====================
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => setTimeout(init, 2000));
    } else {
        setTimeout(init, 2000);
    }

})();
