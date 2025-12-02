// ==UserScript==
// @name         Connect - Combined Highlight, Download & Activity Counter (Optimized)
// @namespace    https://github.com/alerts4kiran-cpu/efficient-work
// @version      5.1
// @description  Maximum optimized script with toggle switches and seconds precision for highlighting
// @author       alerts4kiran-cpu
// @match        https://c2-na-prod.my.connect.aws/real-time-metrics*
// @match        https://c2-na-prod.awsapps.com/connect/real-time-metrics*
// @grant        GM_setValue
// @grant        GM_getValue
// @updateURL    https://raw.githubusercontent.com/alerts4kiran-cpu/efficient-work/main/Connect%20Monitoring%20-%20Highlight,%20Download%20&%20Activity%20Counter%20(Optimized)-5.0.user.js
// @downloadURL  https://raw.githubusercontent.com/alerts4kiran-cpu/efficient-work/main/Connect%20Monitoring%20-%20Highlight,%20Download%20&%20Activity%20Counter%20(Optimized)-5.0.user.js
// @supportURL   https://github.com/alerts4kiran-cpu/efficient-work/issues
// ==/UserScript==
(function() {
    'use strict';

    // ==================== CONFIGURATION ====================
    const CONFIG = {
        DEBOUNCE: 500,
        UPDATE_THROTTLE: 500,
        HIGHLIGHT_INTERVAL: 5000,
        ACTIVITY_UPDATE_INTERVAL: 1000,
        TIME_REGEX: /^\d{1,2}:\d{2}(:\d{2})?$/,
        LETTER_REGEX: /[a-zA-Z]/,
        ACTIVITY_SUMMARY_KEYWORDS: [
            '=== ACTIVITY SUMMARY ===',
            'Activity',
            'Total',
            'After contact work',
            'Available',
            'On contact',
            'Break',
            'Lunch',
            'Personal',
            'Training',
            'Meeting',
            'Project',
            'Manager 1-1',
            'Incoming',
            'Missed',
            'Outage',
            'Manager Approved',
            'System/Power/Internet Outage',
            'Skip Meeting',
            'Start Up',
            'Team Huddle'
        ]
    };

    // ==================== STATE MANAGEMENT ====================
    const state = {
        settings: {
            // Toggle switches for each highlight category
            yellowEnabled: GM_getValue('yellowEnabled', true),
            redEnabled: GM_getValue('redEnabled', true),
            blueEnabled: GM_getValue('blueEnabled', true),
            orangeEnabled: GM_getValue('orangeEnabled', true),
            // Duration settings with seconds precision
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
        observer: null,
        activityObserver: null,
        highlightTimeout: null,
        updateInterval: null,
        summaryBox: null,
        isUpdating: false,
        lastUpdateTime: 0,
        cachedActivityCounts: {},
        cachedActivityDetails: {},
        tableCache: null,
        tableCacheTime: 0
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

    // Convert minutes and seconds to total minutes
    const convertToTotalMinutes = (minutes, seconds) => {
        return minutes + (seconds / 60);
    };

    // ==================== HIGHLIGHTING ====================
    const highlightAgents = () => {
        const tables = getTables();
        const len = tables.length;

        // Calculate threshold values in minutes
        const yellowMin = convertToTotalMinutes(state.settings.yellowMinMinutes, state.settings.yellowMinSeconds);
        const yellowMax = convertToTotalMinutes(state.settings.yellowMaxMinutes, state.settings.yellowMaxSeconds);
        const redMin = convertToTotalMinutes(state.settings.redMinMinutes, state.settings.redMinSeconds);
        const breakMin = convertToTotalMinutes(state.settings.breakMinMinutes, state.settings.breakMinSeconds);
        const lunchMin = convertToTotalMinutes(state.settings.lunchMinMinutes, state.settings.lunchMinSeconds);

        for (let t = 0; t < len; t++) {
            const rows = tables[t].rows;
            const rowLen = rows.length;

            for (let i = 0; i < rowLen; i++) {
                const row = rows[i];
                const cells = row.cells;

                if (cells.length <= 2 || row.querySelectorAll('th').length > 5) continue;

                row.style.cssText = '';
                row.removeAttribute('data-highlighted');

                let status = '', duration = 0;
                const cellLen = cells.length;

                for (let j = 0; j < cellLen; j++) {
                    const text = cells[j].textContent.trim();

                    if (text === 'Available') {
                        status = 'Available';
                        for (let k = j + 1; k < Math.min(j + 4, cellLen); k++) {
                            const nextText = cells[k].textContent.trim();
                            if (CONFIG.TIME_REGEX.test(nextText)) {
                                duration = parseTimeToMinutes(nextText);
                                break;
                            }
                        }
                        break;
                    } else if (text.toLowerCase().includes('break')) {
                        status = 'Break';
                        for (let k = j + 1; k < Math.min(j + 4, cellLen); k++) {
                            const nextText = cells[k].textContent.trim();
                            if (CONFIG.TIME_REGEX.test(nextText)) {
                                duration = parseTimeToMinutes(nextText);
                                break;
                            }
                        }
                        break;
                    } else if (text.toLowerCase().includes('lunch')) {
                        status = 'Lunch';
                        for (let k = j + 1; k < Math.min(j + 4, cellLen); k++) {
                            const nextText = cells[k].textContent.trim();
                            if (CONFIG.TIME_REGEX.test(nextText)) {
                                duration = parseTimeToMinutes(nextText);
                                break;
                            }
                        }
                        break;
                    }
                }

                if (status === 'Available') {
                    if (state.settings.redEnabled && duration >= redMin) {
                        row.style.cssText = 'background-color:#ffcccc;font-weight:bold';
                        row.setAttribute('data-highlighted', 'red');
                    } else if (state.settings.yellowEnabled && duration >= yellowMin && duration < yellowMax) {
                        row.style.cssText = 'background-color:#ffff99;font-weight:bold';
                        row.setAttribute('data-highlighted', 'yellow');
                    }
                } else if (status === 'Break' && state.settings.blueEnabled && duration > breakMin) {
                    row.style.cssText = 'background-color:#cce5ff;font-weight:bold';
                    row.setAttribute('data-highlighted', 'blue');
                } else if (status === 'Lunch' && state.settings.orangeEnabled && duration > lunchMin) {
                    row.style.cssText = 'background-color:#ffe5cc;font-weight:bold';
                    row.setAttribute('data-highlighted', 'orange');
                }
            }
        }
    };

    const debouncedHighlight = () => {
        clearTimeout(state.highlightTimeout);
        state.highlightTimeout = setTimeout(highlightAgents, CONFIG.DEBOUNCE);
    };

    // ==================== DATA EXTRACTION (EXPORT FILTERING ONLY) ====================
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

                if (!agentLogin || agentLogin === 'Agent Login' || isActivitySummaryRow(agentLogin)) continue;

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
            csv += `"${activity}",${details.count},"${details.maxDuration}","${details.maxAgent}"\r`;
        }
        csv += `"Total",${totalCount},"",""\r`;
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

    // ==================== ACTIVITY COUNTER (NO EXPORT FILTERING) ====================
    const getActivityCounts = () => {
        const activityDetails = {};
        const tables = getTables();

        for (let t = 0; t < tables.length; t++) {
            const rows = tables[t].rows;
            for (let i = 0; i < rows.length; i++) {
                const cells = rows[i].cells;
                if (cells.length <= 2 || rows[i].querySelectorAll('th').length > 5) continue;

                const agentLogin = cells[0]?.textContent?.trim();
                const activity = cells[2]?.textContent?.trim();
                const durationText = cells[4]?.textContent?.trim();

                if (agentLogin && activity && agentLogin !== 'Agent Login' && containsLetters(activity)) {
                    const durationMinutes = parseTimeToMinutes(durationText);

                    if (!activityDetails[activity]) {
                        activityDetails[activity] = {
                            count: 0,
                            maxDuration: durationText || '-',
                            maxDurationMinutes: durationMinutes,
                            maxAgent: agentLogin
                        };
                    }

                    activityDetails[activity].count++;

                    if (durationMinutes > activityDetails[activity].maxDurationMinutes) {
                        activityDetails[activity].maxDuration = durationText || '-';
                        activityDetails[activity].maxDurationMinutes = durationMinutes;
                        activityDetails[activity].maxAgent = agentLogin;
                    }
                }
            }
        }

        state.cachedActivityCounts = {};
        for (const activity in activityDetails) {
            state.cachedActivityCounts[activity] = activityDetails[activity].count;
        }
        state.cachedActivityDetails = activityDetails;

        return activityDetails;
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
            csvContent += `${activity},${details.count},${details.maxDuration},${details.maxAgent}\r`;
        }
        csvContent += `Total,${totalCount},,\r`;

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
const updateActivitySummary = () => {
        if (!state.summaryBox || state.summaryBox.style.display === 'none' || state.isUpdating) return;

        const now = Date.now();
        if (now - state.lastUpdateTime < CONFIG.UPDATE_THROTTLE) return;

        state.isUpdating = true;
        state.lastUpdateTime = now;

        const activityDetails = getActivityCounts();
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
        state.summaryBox.style.cssText = 'position:fixed;top:10px;left:10px;z-index:9999;background:#fff;padding:15px;border-radius:8px;box-shadow:0 4px 15px rgba(0,0,0,0.3);min-width:400px;max-width:600px;cursor:move;font-family:Arial,sans-serif';
        state.summaryBox.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;border-bottom:2px solid #FF9900;padding-bottom:8px"><h3 style="margin:0;color:#232F3E;font-size:16px">📊 Activity Summary</h3><div><button id="download-summary-btn" style="background:#FF9900;border:none;font-size:14px;cursor:pointer;color:#fff;padding:4px 8px;border-radius:4px;margin-right:8px;font-weight:bold" title="Download CSV">⬇</button><button id="close-summary-btn" style="background:none;border:none;font-size:20px;cursor:pointer;color:#666;padding:0;width:24px;height:24px">×</button></div></div><div style="max-height:400px;overflow-y:auto"><table style="width:100%;border-collapse:collapse;font-size:14px"><thead><tr style="background:#f8f8f8"><th style="padding:8px;border-bottom:2px solid #ddd;text-align:left;font-weight:bold">Activity</th><th style="padding:8px;border-bottom:2px solid #ddd;text-align:center;font-weight:bold">HC</th><th style="padding:8px;border-bottom:2px solid #ddd;text-align:center;font-weight:bold">Highest Duration</th><th style="padding:8px;border-bottom:2px solid #ddd;text-align:center;font-weight:bold">Agent</th></tr></thead><tbody id="activity-summary-tbody"><tr><td colspan="4" style="padding:20px;text-align:center;color:#999">Loading...</td></tr></tbody></table></div><div style="margin-top:10px;padding-top:8px;border-top:1px solid #ddd;font-size:11px;color:#666;text-align:center">Updates every second • Drag to move</div>';
        document.body.appendChild(state.summaryBox);

        makeDraggable(state.summaryBox);

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
            if (e.target.id === 'close-summary-btn' || e.target.id === 'download-summary-btn') return;
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
        };

        const closeDrag = () => {
            document.onmouseup = null;
            document.onmousemove = null;
        };
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
            <h2 style="margin-top:0;color:#333">Agent Highlight Settings</h2>

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

        highlightAgents();
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

    // ==================== INITIALIZATION ====================
    const init = () => {
        createButton('⚙️ Highlight Settings', 10, '#4CAF50', '#45a049', showSettingsDialog);
        createButton('⭐ Download Highlighted', 180, '#9C27B0', '#7B1FA2', () => downloadData(true));
        createButton('📥 Download All', 360, '#FF9900', '#EC7211', () => downloadData(false));
        createToggleActivityButton();
        createActivitySummaryBox();

        highlightAgents();

        const tables = getTables();
        if (tables.length > 0) {
            state.observer = new MutationObserver(debouncedHighlight);
            state.activityObserver = new MutationObserver(() => {
                if (state.summaryBox && state.summaryBox.style.display !== 'none') {
                    updateActivitySummary();
                }
            });

            for (let i = 0; i < tables.length; i++) {
                state.observer.observe(tables[i], {childList: true, subtree: false});
                state.activityObserver.observe(tables[i], {childList: true, subtree: false});
            }
        }

        setInterval(highlightAgents, CONFIG.HIGHLIGHT_INTERVAL);
        state.updateInterval = setInterval(updateActivitySummary, CONFIG.ACTIVITY_UPDATE_INTERVAL);
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => setTimeout(init, 2000));
    } else {
        setTimeout(init, 2000);
    }

    window.addEventListener('beforeunload', () => {
        if (state.observer) state.observer.disconnect();
        if (state.activityObserver) state.activityObserver.disconnect();
        if (state.updateInterval) clearInterval(state.updateInterval);
        clearTimeout(state.highlightTimeout);
    });
})();
