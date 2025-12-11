// ==UserScript==
// @name         WIMS Maximum Age Alert with Dual Visual Boxes - Split Critical Filter (OPTIMIZED)
// @namespace    tampermonkey.net/
// @version      11.1
// @description  Display two independent Category-Severity-Age alert boxes with IMMEDIATE detection and optimized performance
// @author       You
// @match        optimus-internal.amazon.com/wims/v2/tasks*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    let lastAnnouncedAge1 = 0;
    let lastAnnouncedDetails1 = '';
    let lastAnnouncedAge2 = 0;
    let lastAnnouncedDetails2 = '';

    let selectedSeverities1 = ['CRITICAL_CARRIER_DRIVER', 'CRITICAL_OTHERS', 'HIGH', 'MEDIUM', 'LOW'];
    let selectedSeverities2 = ['CRITICAL_CARRIER_DRIVER', 'CRITICAL_OTHERS', 'HIGH', 'MEDIUM', 'LOW'];

    let voiceAlertsEnabled = true;

    let categoryColumnIndex = -1;
    let severityColumnIndex = -1;
    let ageColumnIndex = -1;

    // Performance optimization: Cache DOM elements
    let alertBox1, alertBox2;
    let updateTimer = null;

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    function init() {
        console.log('🚀 Initializing WIMS Dual Alert System (OPTIMIZED)...');
        setTimeout(() => {
            findColumnIndices();
            createAlertBoxes();
            createFilters();
            createVoiceToggle();
            createDownloadButton();
            console.log('✅ All elements created successfully');
        }, 2000); // Reduced from 3000ms
    }

    function findColumnIndices() {
        let headers = document.querySelectorAll('thead th');
        if (headers.length === 0) {
            headers = document.querySelectorAll('th');
        }

        console.log(`Found ${headers.length} header elements`);

        headers.forEach((header, index) => {
            const headerText = header.textContent.trim();
            if (headerText === 'Category') categoryColumnIndex = index;
            else if (headerText === 'Severity') severityColumnIndex = index;
            else if (headerText === 'Age') ageColumnIndex = index;
        });

        console.log(`Column indices - Category: ${categoryColumnIndex}, Severity: ${severityColumnIndex}, Age: ${ageColumnIndex}`);
    }

    function createAlertBoxes() {
        // Alert Box 1 (Right side)
        alertBox1 = document.createElement('div');
        alertBox1.id = 'wim-age-alert-1';
        alertBox1.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            width: 300px;
            min-height: 150px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 15px;
            font-size: 24px;
            font-weight: bold;
            color: #000;
            border: 3px solid #333;
            border-radius: 10px;
            z-index: 99999;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            background-color: #FFFFFF;
            cursor: move;
            user-select: none;
            text-align: center;
            line-height: 1.4;
        `;
        alertBox1.innerHTML = '<div style="font-size: 18px;">Alert Box 1<br>No WIMs</div>';
        document.body.appendChild(alertBox1);
        makeDraggable(alertBox1);
        console.log('✅ Alert Box 1 created');

        // Alert Box 2 (Left side)
        alertBox2 = document.createElement('div');
        alertBox2.id = 'wim-age-alert-2';
        alertBox2.style.cssText = `
            position: fixed;
            top: 20px;
            left: 20px;
            width: 300px;
            min-height: 150px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 15px;
            font-size: 24px;
            font-weight: bold;
            color: #000;
            border: 3px solid #666;
            border-radius: 10px;
            z-index: 99999;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            background-color: #FFFFFF;
            cursor: move;
            user-select: none;
            text-align: center;
            line-height: 1.4;
        `;
        alertBox2.innerHTML = '<div style="font-size: 18px;">Alert Box 2<br>No WIMs</div>';
        document.body.appendChild(alertBox2);
        makeDraggable(alertBox2);
        console.log('✅ Alert Box 2 created');

        // Initial update
        updateAlert(alertBox1, selectedSeverities1, 1);
        updateAlert(alertBox2, selectedSeverities2, 2);

        // OPTIMIZED: Reduced interval from 15000ms to 5000ms for faster detection
        setInterval(() => {
            updateAlert(alertBox1, selectedSeverities1, 1);
            updateAlert(alertBox2, selectedSeverities2, 2);
        }, 5000);

        // OPTIMIZED: More efficient MutationObserver with debouncing
        const observer = new MutationObserver(() => {
            // Debounce: Clear existing timer and set new one
            if (updateTimer) clearTimeout(updateTimer);

            // IMMEDIATE update (reduced delay)
            updateTimer = setTimeout(() => {
                updateAlert(alertBox1, selectedSeverities1, 1);
                updateAlert(alertBox2, selectedSeverities2, 2);
            }, 100); // Reduced from 1000ms to 100ms for near-instant detection
        });

        // OPTIMIZED: More targeted observation
        const tableContainer = document.querySelector('tbody') || document.querySelector('table') || document.querySelector('main') || document.body;
        observer.observe(tableContainer, {
            childList: true,
            subtree: true,
            characterData: true // Also watch for text changes
        });
    }

    function createFilters() {
        // Filter 1 (Right side - Blue)
        const filter1 = document.createElement('div');
        filter1.id = 'severity-filter-1';
        filter1.style.cssText = `
            position: fixed;
            top: 200px;
            right: 20px;
            width: 300px;
            min-width: 250px;
            padding: 15px;
            background-color: #e3f2fd;
            border: 3px solid #1976d2;
            border-radius: 10px;
            z-index: 99999;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            cursor: move;
            user-select: none;
            font-family: Arial, sans-serif;
            overflow: auto;
        `;
filter1.innerHTML = createFilterHTML('Alert Box 1 Filter', 'filter1');
        document.body.appendChild(filter1);
        makeDraggableAndResizable(filter1);
        attachFilterListeners('filter1', updateSeverityFilter1);
        console.log('✅ Filter 1 created');

        // Filter 2 (Left side - Orange)
        const filter2 = document.createElement('div');
        filter2.id = 'severity-filter-2';
        filter2.style.cssText = `
            position: fixed;
            top: 200px;
            left: 20px;
            width: 300px;
            min-width: 250px;
            padding: 15px;
            background-color: #fff3e0;
            border: 3px solid #f57c00;
            border-radius: 10px;
            z-index: 99999;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            cursor: move;
            user-select: none;
            font-family: Arial, sans-serif;
            overflow: auto;
        `;
        filter2.innerHTML = createFilterHTML('Alert Box 2 Filter', 'filter2');
        document.body.appendChild(filter2);
        makeDraggableAndResizable(filter2);
        attachFilterListeners('filter2', updateSeverityFilter2);
        console.log('✅ Filter 2 created');
    }

    function createVoiceToggle() {
        const toggleBox = document.createElement('div');
        toggleBox.id = 'voice-toggle-box';
        toggleBox.style.cssText = `
            position: fixed;
            top: 200px;
            right: 340px;
            width: 150px;
            min-width: 120px;
            padding: 15px;
            background-color: #f0f0f0;
            border: 3px solid #666;
            border-radius: 10px;
            z-index: 99999;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            cursor: move;
            user-select: none;
            font-family: Arial, sans-serif;
            text-align: center;
            overflow: auto;
        `;

        toggleBox.innerHTML = `
            <div style="font-size: 14px; font-weight: bold; margin-bottom: 10px; color: #333;">
                🔊 Voice Alerts
            </div>
            <label style="display: flex; align-items: center; justify-content: center; gap: 10px; cursor: pointer;">
                <span style="font-size: 13px; font-weight: 600;">OFF</span>
                <div style="position: relative; width: 50px; height: 26px; background-color: #4CAF50; border-radius: 13px; transition: background-color 0.3s;" id="toggle-switch">
                    <div style="position: absolute; top: 3px; right: 3px; width: 20px; height: 20px; background-color: white; border-radius: 50%; transition: right 0.3s;" id="toggle-slider"></div>
                </div>
                <span style="font-size: 13px; font-weight: 600;">ON</span>
            </label>
            <div style="font-size: 11px; margin-top: 8px; color: #666;" id="toggle-status">
                Voice alerts are <strong>ON</strong>
            </div>
        `;

        document.body.appendChild(toggleBox);
        makeDraggableAndResizable(toggleBox);

        const toggleSwitch = document.getElementById('toggle-switch');
        const toggleSlider = document.getElementById('toggle-slider');
        const toggleStatus = document.getElementById('toggle-status');

        toggleSwitch.addEventListener('click', () => {
            voiceAlertsEnabled = !voiceAlertsEnabled;

            if (voiceAlertsEnabled) {
                toggleSwitch.style.backgroundColor = '#4CAF50';
                toggleSlider.style.right = '3px';
                toggleSlider.style.left = 'auto';
                toggleStatus.innerHTML = 'Voice alerts are <strong>ON</strong>';
                console.log('🔊 Voice alerts ENABLED');
            } else {
                toggleSwitch.style.backgroundColor = '#ccc';
                toggleSlider.style.right = 'auto';
                toggleSlider.style.left = '3px';
                toggleStatus.innerHTML = 'Voice alerts are <strong>OFF</strong>';
                console.log('🔇 Voice alerts DISABLED');
                window.speechSynthesis.cancel();
            }
        });

        console.log('✅ Voice toggle created');
    }

    function createFilterHTML(title, filterId) {
        return `
            <div class="filter-content" style="font-size: 14px; font-weight: bold; margin-bottom: 10px; text-align: center; color: #333;">
                📊 ${title}
            </div>
            <div class="filter-content" style="display: flex; flex-direction: column; gap: 8px;">
                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; padding: 5px;">
                    <input type="checkbox" class="${filterId}-checkbox" value="CRITICAL_CARRIER_DRIVER" checked style="width: 18px; height: 18px; flex-shrink: 0;">
                    <span style="font-size: 13px; font-weight: 600; color: #dc3545; word-wrap: break-word;">Critical - Carrier/Driver Callback (≥3min)</span>
                </label>
                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; padding: 5px;">
                    <input type="checkbox" class="${filterId}-checkbox" value="CRITICAL_OTHERS" checked style="width: 18px; height: 18px; flex-shrink: 0;">
                    <span style="font-size: 13px; font-weight: 600; color: #dc3545; word-wrap: break-word;">Critical - Others (≥8min)</span>
                </label>
                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; padding: 5px;">
                    <input type="checkbox" class="${filterId}-checkbox" value="HIGH" checked style="width: 18px; height: 18px; flex-shrink: 0;">
                    <span style="font-size: 14px; font-weight: 600; color: #fd7e14;">High (≥20min)</span>
                </label>
                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; padding: 5px;">
                    <input type="checkbox" class="${filterId}-checkbox" value="MEDIUM" checked style="width: 18px; height: 18px; flex-shrink: 0;">
                    <span style="font-size: 14px; font-weight: 600; color: #ffc107;">Medium (≥20min)</span>
                </label>
                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; padding: 5px;">
                    <input type="checkbox" class="${filterId}-checkbox" value="LOW" checked style="width: 18px; height: 18px; flex-shrink: 0;">
                    <span style="font-size: 14px; font-weight: 600; color: #28a745;">Low (≥20min)</span>
                </label>
            </div>
        `;
    }

    function attachFilterListeners(filterId, callback) {
        const checkboxes = document.querySelectorAll(`.${filterId}-checkbox`);
        checkboxes.forEach(cb => cb.addEventListener('change', callback));
    }

    function updateSeverityFilter1() {
        const checkboxes = document.querySelectorAll('.filter1-checkbox');
        selectedSeverities1 = Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value);
        console.log('Updated Alert Box 1 filter:', selectedSeverities1);
        if (alertBox1) updateAlert(alertBox1, selectedSeverities1, 1);
    }

    function updateSeverityFilter2() {
        const checkboxes = document.querySelectorAll('.filter2-checkbox');
        selectedSeverities2 = Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value);
        console.log('Updated Alert Box 2 filter:', selectedSeverities2);
        if (alertBox2) updateAlert(alertBox2, selectedSeverities2, 2);
    }

    function createDownloadButton() {
        const downloadBtn = document.createElement('div');
        downloadBtn.id = 'wim-download-btn';
        downloadBtn.style.cssText = `
            position: fixed;
            top: 20px;
            right: 340px;
            width: 150px;
            height: 150px;
            min-width: 100px;
            min-height: 100px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            font-size: 14px;
            font-weight: bold;
            color: #fff;
            background-color: #0073bb;
            border: 3px solid #005a8d;
            border-radius: 10px;
            z-index: 99999;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            cursor: move;
            user-select: none;
            overflow: auto;
        `;
        downloadBtn.innerHTML = '<div style="font-size: 32px;">📥</div><div style="margin-top: 8px;">Double-Click</div><div style="font-size: 12px; margin-top: 4px;">to Download</div>';
        downloadBtn.addEventListener('dblclick', (e) => {
            e.preventDefault();
            e.stopPropagation();
            downloadWIMData();
        });
        document.body.appendChild(downloadBtn);
        makeDraggableAndResizable(downloadBtn);
        console.log('✅ Download button created');
    }

    function downloadWIMData() {
        try {
            const headers = [];
            const headerElements = document.querySelectorAll('thead th, th');
            let idColumnIndex = -1;

            headerElements.forEach((header, index) => {
                const text = header.textContent.trim();
                if (text && !headers.includes(text)) {
                    headers.push(text);
                    if (text === 'ID') idColumnIndex = index;
                }
            });

            if (headers.length === 0) {
                alert('No table headers found.');
                return;
            }

            const rows = [];
            const tableRows = document.querySelectorAll('tbody tr, table tr');

            tableRows.forEach(row => {
                const cells = row.querySelectorAll('td');
                if (cells.length > 0) {
                    const rowData = [];
                    cells.forEach((cell, cellIndex) => {
                        let text = cell.textContent.trim();
                        if (cellIndex === idColumnIndex) {
                            const link = cell.querySelector('a');
                            if (link && link.href) text = link.href;
                        }
                        text = text.replace(/"/g, '""');
                        rowData.push(`"${text}"`);
                    });
                    rows.push(rowData);
                }
            });

            if (rows.length === 0) {
                alert('No data rows found.');
                return;
            }

            const now = new Date();
            const downloadTimestamp = now.toLocaleString('en-US', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: false
            });

            const NEWLINE = String.fromCharCode(13, 10);
            const timestampRow = `"Downloaded on: ${downloadTimestamp}"`;
            const csvHeaders = headers.map(h => `"${h.replace(/"/g, '""')}"`).join(',');
            const csvRows = rows.map(row => row.join(',')).join(NEWLINE);
            const csvContent = [timestampRow, '', csvHeaders, csvRows].join(NEWLINE);

            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            const fileTimestamp = now.toISOString().replace(/[:.]/g, '-').slice(0, -5).replace('T', '_');
            const filename = `WIMS_Data_${fileTimestamp}.csv`;

            link.setAttribute('href', url);
            link.setAttribute('download', filename);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            console.log(`✓ Downloaded ${rows.length} rows`);
        } catch (error) {
            console.error('Error downloading:', error);
        }
    }
 function makeDraggable(element) {
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        element.onmousedown = dragMouseDown;

        function dragMouseDown(e) {
            if (e.target.tagName === 'INPUT' ||
                e.target.tagName === 'LABEL' ||
                e.target.closest('#toggle-switch') ||
                e.target.closest('label')) {
                return;
            }

            e.preventDefault();
            pos3 = e.clientX;
            pos4 = e.clientY;
            document.onmouseup = closeDragElement;
            document.onmousemove = elementDrag;
        }

        function elementDrag(e) {
            e.preventDefault();
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;
            element.style.top = (element.offsetTop - pos2) + "px";
            element.style.left = (element.offsetLeft - pos1) + "px";
            element.style.right = "auto";
        }

        function closeDragElement() {
            document.onmouseup = null;
            document.onmousemove = null;
        }
    }

    function makeDraggableAndResizable(element) {
        const resizeHandle = document.createElement('div');
        resizeHandle.style.cssText = `
            position: absolute;
            bottom: 0;
            right: 0;
            width: 20px;
            height: 20px;
            cursor: nwse-resize;
            background: linear-gradient(135deg, transparent 50%, rgba(0,0,0,0.3) 50%);
            border-bottom-right-radius: 7px;
            z-index: 10;
        `;
        element.style.position = 'fixed';
        element.appendChild(resizeHandle);

        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        let isDragging = false;

        let isResizing = false;
        let startWidth = 0;
        let startHeight = 0;
        let startX = 0;
        let startY = 0;

        element.onmousedown = function(e) {
            if (e.target.tagName === 'INPUT' ||
                e.target.tagName === 'LABEL' ||
                e.target.closest('#toggle-switch') ||
                e.target.closest('label') ||
                e.target === resizeHandle) {
                return;
            }

            e.preventDefault();
            isDragging = true;
            pos3 = e.clientX;
            pos4 = e.clientY;
            document.onmouseup = closeDragElement;
            document.onmousemove = elementDrag;
        };

        function elementDrag(e) {
            if (!isDragging) return;
            e.preventDefault();
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;
            element.style.top = (element.offsetTop - pos2) + "px";
            element.style.left = (element.offsetLeft - pos1) + "px";
            element.style.right = "auto";
        }

        function closeDragElement() {
            isDragging = false;
            document.onmouseup = null;
            document.onmousemove = null;
        }

        resizeHandle.onmousedown = function(e) {
            e.preventDefault();
            e.stopPropagation();
            isResizing = true;
            startWidth = element.offsetWidth;
            startHeight = element.offsetHeight;
            startX = e.clientX;
            startY = e.clientY;

            document.onmousemove = elementResize;
            document.onmouseup = closeResizeElement;
        };

        function elementResize(e) {
            if (!isResizing) return;
            e.preventDefault();

            const newWidth = startWidth + (e.clientX - startX);
            const newHeight = startHeight + (e.clientY - startY);

            if (newWidth > 100) {
                element.style.width = newWidth + 'px';
            }
            if (newHeight > 100) {
                element.style.height = newHeight + 'px';
                element.style.minHeight = newHeight + 'px';
            }
        }

        function closeResizeElement() {
            isResizing = false;
            document.onmousemove = null;
            document.onmouseup = null;
        }
    }

    function isCarrierOrDriverCallback(category) {
        const categoryLower = category.toLowerCase();
        return categoryLower.includes('carrier callback') || categoryLower.includes('driver callback');
    }

    function mapSeverityToFilterCategory(severity, category) {
        if (severity === 'CRITICAL') {
            if (isCarrierOrDriverCallback(category)) {
                return 'CRITICAL_CARRIER_DRIVER';
            } else {
                return 'CRITICAL_OTHERS';
            }
        }
        return severity;
    }

    function speakAlert(ageMinutes, category, severity, boxNumber) {
        if (!voiceAlertsEnabled) {
            return;
        }

        const filterCategory = mapSeverityToFilterCategory(severity, category);
        const selectedSeverities = boxNumber === 1 ? selectedSeverities1 : selectedSeverities2;

        if (!selectedSeverities.includes(filterCategory)) return;

        let ageThreshold = 3;

        if (filterCategory === 'CRITICAL_CARRIER_DRIVER') {
            ageThreshold = 3;
        } else if (filterCategory === 'CRITICAL_OTHERS') {
            ageThreshold = 8;
        } else if (severity === 'HIGH' || severity === 'MEDIUM' || severity === 'LOW') {
            ageThreshold = 20;
        }

        if (ageMinutes >= ageThreshold) {
            const currentDetails = `${category}-${severity}-${ageMinutes}`;
            const lastDetails = boxNumber === 1 ? lastAnnouncedDetails1 : lastAnnouncedDetails2;

            if (currentDetails !== lastDetails) {
                const utterance = new SpeechSynthesisUtterance();
                utterance.text = `Alert Box ${boxNumber}: ${category} ${severity} WIM at ${ageMinutes} min`;
                utterance.rate = 1.0;
                utterance.pitch = 1.0;
                utterance.volume = 1.0;
                window.speechSynthesis.speak(utterance);

                if (boxNumber === 1) {
                    lastAnnouncedDetails1 = currentDetails;
                    lastAnnouncedAge1 = ageMinutes;
                } else {
                    lastAnnouncedDetails2 = currentDetails;
                    lastAnnouncedAge2 = ageMinutes;
                }
                console.log(`🔊 Voice alert Box ${boxNumber}: ${utterance.text} (threshold: ${ageThreshold}min)`);
            }
        }
    }

    function getBackgroundColor(ageMinutes) {
        if (ageMinutes < 2) return '#FFFFFF';
        else if (ageMinutes < 3) return '#FFFF00';
        else if (ageMinutes < 4) return '#FFA500';
        else if (ageMinutes < 5) return '#FF0000';
        else return '#000000';
    }

    function getTextColor(ageMinutes) {
        return ageMinutes >= 5 ? '#FFFFFF' : '#000000';
    }

    function parseAgeToMinutes(ageString) {
        if (!ageString) return 0;
        let totalMinutes = 0;
        const dayMatch = ageString.match(/(\d+)d/);
        if (dayMatch) totalMinutes += parseInt(dayMatch) * 24 * 60;
        const hourMatch = ageString.match(/(\d+)h/);
        if (hourMatch) totalMinutes += parseInt(hourMatch) * 60;
        const minMatch = ageString.match(/(\d+)m/);
        if (minMatch) totalMinutes += parseInt(minMatch);
        return totalMinutes;
    }

    function findMaxAgeWithDetails(selectedSeverities) {
        try {
            let maxAge = 0;
            let maxCategory = 'Work Item';
            let maxSeverity = 'Alert';

            if (categoryColumnIndex === -1 || severityColumnIndex === -1 || ageColumnIndex === -1) {
                findColumnIndices();
            }

            const tableRows = document.querySelectorAll('tbody tr, table tr');

            tableRows.forEach((row) => {
                const cells = row.querySelectorAll('td');
                if (cells.length === 0) return;

                if (categoryColumnIndex >= 0 && severityColumnIndex >= 0 && ageColumnIndex >= 0) {
                    if (cells.length > Math.max(categoryColumnIndex, severityColumnIndex, ageColumnIndex)) {
                        const categoryText = cells[categoryColumnIndex]?.textContent.trim() || '';
                        const severityText = cells[severityColumnIndex]?.textContent.trim() || '';
                        const ageText = cells[ageColumnIndex]?.textContent.trim() || '';

                        const filterCategory = mapSeverityToFilterCategory(severityText, categoryText);

                        if (!selectedSeverities.includes(filterCategory)) return;

                        if (ageText.match(/\d+d\s+\d+h\s+\d+m/)) {
                            const ageInMinutes = parseAgeToMinutes(ageText);
                            if (ageInMinutes > maxAge) {
                                maxAge = ageInMinutes;
                                maxCategory = categoryText || 'Work Item';
                                maxSeverity = severityText || 'Alert';
                            }
                        }
                    }
                }
            });

            return { age: maxAge, category: maxCategory, severity: maxSeverity };
        } catch (error) {
            console.error('Error finding max age:', error);
            return { age: 0, category: 'Work Item', severity: 'Alert' };
        }
    }
function updateAlert(alertBox, selectedSeverities, boxNumber) {
        try {
            const details = findMaxAgeWithDetails(selectedSeverities);
            const maxAge = details.age;
            const backgroundColor = getBackgroundColor(maxAge);
            const textColor = getTextColor(maxAge);

            alertBox.style.backgroundColor = backgroundColor;
            alertBox.style.color = textColor;

            if (maxAge === 0) {
                alertBox.innerHTML = `<div style="font-size: 18px;">Alert Box ${boxNumber}<br>No WIMs</div>`;
            } else {
                const categoryDisplay = details.category || 'Work Item';
                const severityDisplay = details.severity || 'Alert';
                alertBox.innerHTML = `
                    <div style="font-size: 16px; margin-bottom: 5px;">${categoryDisplay}</div>
                    <div style="font-size: 20px; font-weight: bold; margin-bottom: 5px;">${severityDisplay}</div>
                    <div style="font-size: 32px; font-weight: bold;">${maxAge} min</div>
                `;
            }

            speakAlert(maxAge, details.category, details.severity, boxNumber);
        } catch (error) {
            console.error('Error updating alert:', error);
        }
    }

})();