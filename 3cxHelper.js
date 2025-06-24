// ==UserScript==
// @name         3CX Helper - Tools
// @namespace    http://tampermonkey.net/
// @version      3.4
// @description  Bulk Copy/ Import Holiday , User Password, Export.
// @match        *://*.3cx.be/*
// @match        *://*.3cx.com/*
// @match        *://*.3cx.eu/*
// @match        *://*.my3cx.be/*
// @match        *://*.my3cx.eu/*
// @updateURL   https://raw.githubusercontent.com/th0mas137/TM-3cxHelper/refs/heads/main/3cxHelper.js
// @downloadURL https://raw.githubusercontent.com/th0mas137/TM-3cxHelper/refs/heads/main/3cxHelper.js
// @grant        GM_notification
// @grant        GM_setClipboard
// @grant        GM_addStyle
// ==/UserScript==

(function () {
    'use strict';

    // ==================== CAPTURE TOKEN! ====================
    (function() {
        const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
        XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
            if (header.toLowerCase() === 'authorization' && value.startsWith('Bearer ')) {
                window.__activeBearerToken = value.substring(7);
            }
            return originalSetRequestHeader.apply(this, arguments);
        };
    })();

    // ==================== CAPTURE USER DATA ====================
    let capturedUserData = null;
    
    (function() {
        const originalXHROpen = XMLHttpRequest.prototype.open;
        const originalXHRSend = XMLHttpRequest.prototype.send;
        
        XMLHttpRequest.prototype.open = function(method, url) {
            const urlString = typeof url === 'string' ? url : '';
            this._isUserDataRequest = method === 'GET' && urlString.includes('/xapi/v1/Users(');
            return originalXHROpen.apply(this, arguments);
        };
        
        XMLHttpRequest.prototype.send = function(body) {
            if (this._isUserDataRequest) {
                const originalOnReadyStateChange = this.onreadystatechange;
                this.onreadystatechange = function() {
                    if (this.readyState === 4 && this.status === 200) {
                        try {
                            const userData = JSON.parse(this.responseText);
                            if (userData.AuthID || userData.AuthPassword) {
                                capturedUserData = {
                                    AuthID: userData.AuthID || 'N/A',
                                    AuthPassword: userData.AuthPassword || 'N/A'
                                };
                                setTimeout(() => {
                                    injectUserDataFields();
                                }, 1000);
                            }
                        } catch (e) {}
                    }
                    if (originalOnReadyStateChange) {
                        originalOnReadyStateChange.apply(this, arguments);
                    }
                };
            }
            return originalXHRSend.apply(this, arguments);
        };
    })();

    /**
     * Returns the active bearer token captured from outgoing API calls.
     * @returns {string|null} The bearer token (without the "Bearer " prefix), or null if not yet captured.
     */
    function getActiveBearerToken() {
        return window.__activeBearerToken || null;
    }

    // -------------------- BELGIUM HOLIDAYS (EDIT AS NEEDED) --------------------
    const belgiumHolidays = [
        { id: 39, name: "Nieuwjaar", month: 1, day: 1 },
        { id: 40, name: "Feest van de arbeid", month: 5, day: 1 },
        { id: 41, name: "Kerstmis", month: 12, day: 25 }
    ];

    // ==================== URL-BASED INJECTION HELPERS ====================
    function maybeInjectPasswordField() {
        if (window.location.hash.startsWith('#/office/users/edit/')) {
            injectPasswordField();
        }
    }
    function maybeInjectUserDataFields() {
        if (window.location.hash.startsWith('#/office/users/edit/')) {
            injectUserDataFields();
        }
    }
    function maybeInjectExportButtonUI() {
        if (window.location.hash.startsWith('#/office/call-handling')) {
            injectExportButtonUI();
        } else {
            removeExportButtonUI();
        }
    }
    function maybeInjectImportHolidaysButtonUI() {
        if (window.location.hash.startsWith('#/office/office-hours/')) {
            injectImportHolidaysButtonUI();
            injectCopyHolidaysButtonUI();
        } else {
            removeImportHolidaysButtonUI();
            removeCopyHolidaysButtonUI();
        }
    }

    // ==================== REMOVE BUTTONS WHEN NEEDED ====================
    function removeExportButtonUI() {
        const exportBtn = document.getElementById('export-queues-btn');
        if (exportBtn) {
            const wrapper = exportBtn.closest('app-button-add');
            if (wrapper) wrapper.remove();
        }
    }
    function removeImportHolidaysButtonUI() {
        const importBtn = document.getElementById('import-holidays-btn');
        if (importBtn) {
            const wrapper = importBtn.closest('app-button-add');
            if (wrapper) wrapper.remove();
        }
    }
    function removeCopyHolidaysButtonUI() {
        const copyBtn = document.getElementById('copy-holidays-btn');
        if (copyBtn) {
            const wrapper = copyBtn.closest('app-button-add');
            if (wrapper) wrapper.remove();
        }
    }

    // ==================== PASSWORD FIELD FUNCTIONALITY ====================
    function injectPasswordField() {
        if (document.getElementById('custom-password-field')) return;
        const formLabel = document.querySelector('label[for^="did-select"]');
        if (formLabel) {
            const passwordDiv = document.createElement('div');
            passwordDiv.innerHTML = `
                <div class="d-flex align-items-center form-label">
                    <label for="custom-password-field">Access Password</label>
                </div>
                <input type="password" id="custom-password-field" class="form-control" style="width:100%;" placeholder="Enter password (optional)">
            `;
            formLabel.parentNode.parentNode.appendChild(passwordDiv);
        }
    }
    
    // ==================== USER DATA FIELDS FUNCTIONALITY ====================
    function injectUserDataFields() {
        if (!capturedUserData) return;
        if (document.getElementById('user-auth-id-field')) return;
        
        const accessPasswordField = document.getElementById('custom-password-field');
        let injectionPoint = accessPasswordField ? accessPasswordField.parentNode : 
            document.querySelector('label[for^="did-select"]')?.parentNode?.parentNode;
        
        if (!injectionPoint) {
            setTimeout(() => { injectUserDataFields(); }, 2000);
            return;
        }
        
        const userDataDiv = document.createElement('div');
        userDataDiv.style.marginTop = '15px';
        userDataDiv.innerHTML = `
            <div style="padding: 15px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;">
                <h6 style="margin-bottom: 15px; color: #495057; font-weight: bold;">Phone Auth</h6>
                
                <div style="display: flex; gap: 15px; flex-wrap: wrap;">
                    <div style="flex: 1; min-width: 250px;">
                        <div class="d-flex align-items-center form-label">
                            <label for="user-auth-id-field" style="font-weight: 500;">Auth ID</label>
                        </div>
                        <input type="text" id="user-auth-id-field" class="form-control" style="width:100%; background: #fff;" value="${capturedUserData.AuthID}" readonly>
                    </div>
                    
                    <div style="flex: 1; min-width: 250px;">
                        <div class="d-flex align-items-center form-label">
                            <label for="user-auth-password-field" style="font-weight: 500;">Auth Password</label>
                        </div>
                        <div style="position: relative;">
                            <input type="password" id="user-auth-password-field" class="form-control" style="width:100%; background: #fff; padding-right: 40px;" value="${capturedUserData.AuthPassword}" readonly>
                            <button type="button" id="toggle-auth-password" style="position: absolute; right: 8px; top: 50%; transform: translateY(-50%); border: none; background: none; cursor: pointer; padding: 4px 6px; border-radius: 3px; color: #6c757d; font-size: 14px; font-family: monospace; font-weight: bold;" title="Toggle password visibility">👁</button>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        injectionPoint.insertAdjacentElement('afterend', userDataDiv);
        
        // Add event listener for password toggle (avoiding CSP violation)
        const toggleButton = document.getElementById('toggle-auth-password');
        const passwordField = document.getElementById('user-auth-password-field');
        
        if (toggleButton && passwordField) {
            toggleButton.addEventListener('click', function() {
                if (passwordField.type === 'password') {
                    passwordField.type = 'text';
                    toggleButton.textContent = '🔒';
                    toggleButton.title = 'Hide password';
                } else {
                    passwordField.type = 'password';
                    toggleButton.textContent = '👁';
                    toggleButton.title = 'Show password';
                }
            });
        }
    }
    
    const passwordObserver = new MutationObserver(() => {
        maybeInjectPasswordField();
        maybeInjectUserDataFields();
    });
    passwordObserver.observe(document.body, { childList: true, subtree: true });

    // ==================== MODIFYING REQUEST FOR PASSWORD ====================
    const originalOpen = XMLHttpRequest.prototype.open;
    const originalSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (method, url) {
        this._isTargetRequest = (method === 'POST' || method === 'PATCH') && url.includes('/xapi/v1/Users(');
        return originalOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function (body) {
        if (this._isTargetRequest && body) {
            try {
                const jsonBody = JSON.parse(body);
                const passwordField = document.getElementById('custom-password-field');
                if (passwordField?.value?.trim()) {
                    jsonBody.AccessPassword = passwordField.value.trim();
                }
                arguments[0] = JSON.stringify(jsonBody);
            } catch (e) {
                console.error('Error modifying request:', e);
            }
        }
        return originalSend.apply(this, arguments);
    };

    // ==================== STYLE INJECTIONS ====================
    GM_addStyle(`
        /* Export dropdown styling */
        .download-dropdown-menu {
            position: absolute;
            top: 100%;
            right: 0;
            background: white;
            border: 1px solid #ccc;
            border-radius: 4px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            z-index: 9999;
            min-width: 150px;
            display: none;
        }
        .download-dropdown-menu div {
            padding: 8px 12px;
            cursor: pointer;
        }
        .download-dropdown-menu div:hover {
            background: #f0f0f0;
        }

        /* Unified holidays dropdown styling for Bulk Import and Copy Holidays */
        .holidays-dropdown-menu {
            position: absolute;
            top: 100%;
            right: 0;
            background: white;
            border: 1px solid #ccc;
            border-radius: 4px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            z-index: 2147483647;
            min-width: 200px;
            padding: 10px;
            display: none;
        }

        .department-list {
            max-height: 300px;
            overflow-y: auto;
            margin-bottom: 10px;
        }

        .department-list div {
            margin-bottom: 5px;
        }

        .holidays-dropdown-menu button {
            margin-top: 10px;
            padding: 5px 10px;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }

        .holidays-dropdown-menu button:hover {
            background: #0069d9;
        }
    `);

    // ==================== QUEUE EXPORT FUNCTIONALITY ====================
    let lastExportData = null;
    function injectExportButtonUI() {
        if (document.getElementById('export-queues-btn')) return;
        const container = document.querySelector('.d-flex.flex-wrap.gap-2.flex-grow-1.mw-0');
        if (container) {
            const btnWrapper = document.createElement('app-button-add');
            btnWrapper.setAttribute('data-qa', 'exportQueue');
            btnWrapper.setAttribute('text', 'Export Queue');
            const innerDiv = document.createElement('div');
            innerDiv.setAttribute('container', 'body');
            innerDiv.className = 'd-inline-block';
            innerDiv.style.position = 'relative';
            const button = document.createElement('button');
            button.type = 'button';
            button.setAttribute('data-qa', 'button');
            button.className = 'btn btn-primary';
            button.id = 'export-queues-btn';
            button.textContent = 'Export Queue';
            button.onclick = handleQueueExport;
            innerDiv.appendChild(button);
            btnWrapper.appendChild(innerDiv);
            container.appendChild(btnWrapper);
        }
    }
    const exportObserver = new MutationObserver(() => maybeInjectExportButtonUI());
    exportObserver.observe(document.body, { childList: true, subtree: true });

    async function handleQueueExport() {
        if (lastExportData) return;
        let token;
        try {
            token = await getBearerToken();
        } catch (error) {
            console.error("Token error:", error);
            showNotification(`ERROR: ${error.message}`, false);
            return;
        }
        try {
            const queues = await fetchQueues(token);
            if (!queues?.value?.length) throw new Error('No queues found in the system');
            lastExportData = await processQueues(queues.value, token);
            const exportBtn = document.getElementById('export-queues-btn');
            if (exportBtn) {
                exportBtn.textContent = 'Download Options';
                exportBtn.onclick = toggleDownloadDropdown;
                setupDownloadDropdown(exportBtn);
                showNotification('Export complete! Choose your download option.', true);
            }
        } catch (error) {
            console.error('Export Error:', error);
            showNotification(`ERROR: ${error.message}`, false);
        }
    }

    function setupDownloadDropdown(anchorButton) {
        let dropdown = anchorButton.parentElement.querySelector('.download-dropdown-menu');
        if (!dropdown) {
            dropdown = document.createElement('div');
            dropdown.className = 'download-dropdown-menu';
            anchorButton.parentElement.appendChild(dropdown);
        } else {
            dropdown.innerHTML = '';
        }
        const downloadCsv = document.createElement('div');
        downloadCsv.textContent = 'Download CSV';
        downloadCsv.onclick = () => {
            dropdown.style.display = 'none';
            downloadAsCsv(lastExportData);
        };
        dropdown.appendChild(downloadCsv);

        const downloadJson = document.createElement('div');
        downloadJson.textContent = 'Download JSON';
        downloadJson.onclick = () => {
            dropdown.style.display = 'none';
            const jsonContent = JSON.stringify(lastExportData, null, 2);
            downloadAsFile(jsonContent, 'json');
        };
        dropdown.appendChild(downloadJson);

        const copyJson = document.createElement('div');
        copyJson.textContent = 'Copy JSON';
        copyJson.onclick = async () => {
            dropdown.style.display = 'none';
            const jsonContent = JSON.stringify(lastExportData, null, 2);
            try {
                await GM_setClipboard(jsonContent);
                showNotification('JSON copied to clipboard!', true);
            } catch (e) {
                showNotification('Failed to copy JSON to clipboard.', false);
            }
        };
        dropdown.appendChild(copyJson);
    }

    function toggleDownloadDropdown() {
        const exportBtn = document.getElementById('export-queues-btn');
        if (!exportBtn) return;
        const dropdown = exportBtn.parentElement.querySelector('.download-dropdown-menu');
        if (!dropdown) return;
        dropdown.style.display = (dropdown.style.display === 'block') ? 'none' : 'block';
    }

    // Hide the export dropdown if clicked outside
    document.addEventListener('click', function(e) {
        const exportBtn = document.getElementById('export-queues-btn');
        if (exportBtn) {
            const container = exportBtn.parentElement;
            if (!container.contains(e.target)) {
                const dropdown = container.querySelector('.download-dropdown-menu');
                if (dropdown) dropdown.style.display = 'none';
            }
        }
    });

    // ==================== BULK IMPORT HOLIDAYS FUNCTIONALITY ====================
    function injectImportHolidaysButtonUI() {
        if (document.getElementById('import-holidays-btn')) return;

        // Find the native +Add holiday button
        const existingAddButton = document.querySelector('app-button-add[data-qa="add-holiday"]');
        if (!existingAddButton) return;

        // Create our new button
        const btnWrapper = document.createElement('app-button-add');
        btnWrapper.setAttribute('data-qa', 'importHolidays');
        btnWrapper.setAttribute('text', 'Bulk Import Holidays');
        btnWrapper.classList.add('mb-3'); // same class 3cx uses for spacing
        btnWrapper.style.marginLeft = '10px'; // extra spacing from the +Add button

        const innerDiv = document.createElement('div');
        innerDiv.setAttribute('container', 'body');
        innerDiv.className = 'd-inline-block';
        innerDiv.style.position = 'relative';

        const button = document.createElement('button');
        button.type = 'button';
        button.setAttribute('data-qa', 'button');
        button.className = 'btn btn-primary';
        button.id = 'import-holidays-btn';
        button.textContent = 'Bulk Import Holidays';
        button.onclick = function() {
            showImportHolidaysDropdown(button);
        };

        innerDiv.appendChild(button);
        btnWrapper.appendChild(innerDiv);

        // Insert our new button AFTER the existing +Add button
        existingAddButton.insertAdjacentElement('afterend', btnWrapper);
    }

    const importHolidaysObserver = new MutationObserver(() => maybeInjectImportHolidaysButtonUI());
    importHolidaysObserver.observe(document.body, { childList: true, subtree: true });

    function showImportHolidaysDropdown(anchorButton) {
        let dropdown = anchorButton.parentElement.querySelector('.holidays-dropdown-menu');
        if (!dropdown) {
            dropdown = document.createElement('div');
            dropdown.className = 'holidays-dropdown-menu';
            anchorButton.parentElement.appendChild(dropdown);
        } else {
            dropdown.innerHTML = '';
        }

        // Create a form with checkboxes for each holiday
        const form = document.createElement('form');
        belgiumHolidays.forEach(h => {
            const div = document.createElement('div');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `holiday-${h.id}`;
            checkbox.value = h.id;
            checkbox.dataset.holiday = JSON.stringify(h);

            const label = document.createElement('label');
            label.htmlFor = `holiday-${h.id}`;
            label.textContent = h.name;

            div.appendChild(checkbox);
            div.appendChild(label);
            form.appendChild(div);
        });

        // Confirm button
        const confirmBtn = document.createElement('button');
        confirmBtn.type = 'button';
        confirmBtn.textContent = 'Confirm Import';
        confirmBtn.onclick = async (e) => {
            e.preventDefault();
            const parts = window.location.hash.split('/');
            const deptId = parts[parts.length - 1];
            const selectedHolidays = [];
            const checkboxes = form.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => {
                if (cb.checked) {
                    const holiday = JSON.parse(cb.dataset.holiday);
                    selectedHolidays.push(holiday);
                }
            });
            if (selectedHolidays.length === 0) {
                alert("Please select at least one holiday.");
                return;
            }
            try {
                await addHolidaysToDepartment(deptId, selectedHolidays);
                showNotification('Holidays updated successfully!', true);
                dropdown.style.display = 'none';
            } catch (error) {
                showNotification(`Error updating holidays: ${error.message}`, false);
            }
        };
        form.appendChild(confirmBtn);
        dropdown.appendChild(form);
        dropdown.style.display = 'block';
    }

    async function addHolidaysToDepartment(departmentId, selectedHolidays) {
        const token = await getBearerToken();
        // GET current group data with Authorization header
        const groupResponse = await fetch(`${window.location.origin}/xapi/v1/Groups(${departmentId})`, {
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            }
        });
        if (!groupResponse.ok) {
            throw new Error("Failed to fetch current group data");
        }
        const groupData = await groupResponse.json();
        let currentHolidays = groupData.OfficeHolidays || [];

        // Merge current and selected holidays using a map
        const holidayMap = {};
        currentHolidays.forEach(h => { holidayMap[h.Id] = h; });
        selectedHolidays.forEach(h => {
            holidayMap[h.id] = {
                Day: h.day,
                DayEnd: h.day,
                HolidayPrompt: "",
                IsRecurrent: true,
                Month: h.month,
                MonthEnd: h.month,
                Name: h.name,
                TimeOfEndDate: "PT23H59M",
                TimeOfStartDate: "PT0S",
                Year: 0,
                YearEnd: 0,
                Id: h.id
            };
        });
        const newHolidays = Object.values(holidayMap);

        // PATCH request with Authorization header
        const patchResponse = await fetch(`${window.location.origin}/xapi/v1/Groups(${departmentId})`, {
            method: 'PATCH',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({
                OfficeHolidays: newHolidays
            })
        });
        if (!patchResponse.ok) {
            throw new Error(`Failed to update holidays: ${patchResponse.status}`);
        }
        return await patchResponse.json();
    }

    // ==================== COPY HOLIDAYS FUNCTIONALITY ====================
    function injectCopyHolidaysButtonUI() {
        if (document.getElementById('copy-holidays-btn')) return;

        // Find the Bulk Import Holidays button
        const importHolidaysButton = document.getElementById('import-holidays-btn');
        if (!importHolidaysButton) return;

        const importButtonWrapper = importHolidaysButton.closest('app-button-add');
        if (!importButtonWrapper) return;

        // Create our new button
        const btnWrapper = document.createElement('app-button-add');
        btnWrapper.setAttribute('data-qa', 'copyHolidays');
        btnWrapper.setAttribute('text', 'Copy Holidays');
        btnWrapper.classList.add('mb-3'); // same class 3cx uses for spacing
        btnWrapper.style.marginLeft = '10px'; // extra spacing from the previous button

        const innerDiv = document.createElement('div');
        innerDiv.setAttribute('container', 'body');
        innerDiv.className = 'd-inline-block';
        innerDiv.style.position = 'relative';

        const button = document.createElement('button');
        button.type = 'button';
        button.setAttribute('data-qa', 'button');
        button.className = 'btn btn-primary';
        button.id = 'copy-holidays-btn';
        button.textContent = 'Copy Holidays';
        button.onclick = function() {
            showCopyHolidaysDropdown(button);
        };

        innerDiv.appendChild(button);
        btnWrapper.appendChild(innerDiv);

        // Insert our new button AFTER the Bulk Import Holidays button
        importButtonWrapper.insertAdjacentElement('afterend', btnWrapper);
    }

    async function showCopyHolidaysDropdown(anchorButton) {
        let dropdown = anchorButton.parentElement.querySelector('.holidays-dropdown-menu');
        if (!dropdown) {
            dropdown = document.createElement('div');
            dropdown.className = 'holidays-dropdown-menu';
            anchorButton.parentElement.appendChild(dropdown);
        } else {
            dropdown.innerHTML = '';
        }

        dropdown.innerHTML = '<p>Loading departments...</p>';
        dropdown.style.display = 'block';

        try {
            // Get current department ID from URL
            const hashParts = window.location.hash.split('/').filter(part => part !== '');
            const currentDeptId = hashParts[hashParts.length - 1];

            // Get token
            const token = await getBearerToken();

            // Fetch all departments
            const departments = await fetchDepartments(token);

            // Create dropdown content
            dropdown.innerHTML = '<h4>Select target departments:</h4>';

            const departmentListDiv = document.createElement('div');
            departmentListDiv.className = 'department-list';

            // Add departments to the list (excluding current one)
            departments.forEach(dept => {
                if (dept.Id.toString() !== currentDeptId.toString()) {
                    const div = document.createElement('div');
                    const checkbox = document.createElement('input');
                    checkbox.type = 'checkbox';
                    checkbox.id = `dept-${dept.Id}`;
                    checkbox.value = dept.Id;

                    const label = document.createElement('label');
                    label.htmlFor = `dept-${dept.Id}`;
                    label.textContent = dept.Name || `Department ${dept.Id}`;

                    div.appendChild(checkbox);
                    div.appendChild(label);
                    departmentListDiv.appendChild(div);
                }
            });

            dropdown.appendChild(departmentListDiv);

            // Add confirm button
            const confirmBtn = document.createElement('button');
            confirmBtn.type = 'button';
            confirmBtn.textContent = 'Copy Holidays';
            confirmBtn.onclick = async () => {
                const selectedDepts = [];
                const checkboxes = departmentListDiv.querySelectorAll('input[type="checkbox"]');

                checkboxes.forEach(cb => {
                    if (cb.checked) {
                        selectedDepts.push(cb.value);
                    }
                });

                if (selectedDepts.length === 0) {
                    alert("Please select at least one department.");
                    return;
                }

                try {
                    dropdown.innerHTML = '<p>Copying holidays...</p>';
                    await copyHolidaysToDepartments(currentDeptId, selectedDepts, token);
                    showNotification('Holidays copied successfully!', true);
                    dropdown.style.display = 'none';
                } catch (error) {
                    console.error('Copy Error:', error);
                    showNotification(`Error copying holidays: ${error.message}`, false);
                    dropdown.style.display = 'none';
                }
            };

            dropdown.appendChild(confirmBtn);

        } catch (error) {
            console.error('Error loading departments:', error);
            dropdown.innerHTML = `<p>Error: ${error.message}</p>`;
        }
    }

    async function fetchDepartments(token) {
        try {
            const response = await fetch(`${window.location.origin}/xapi/v1/Groups?$filter=not%20startsWith(Name,%20%27___FAVORITES___%27)&$orderby=Name&$select=Id,Name`, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                }
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            return data.value || [];
        } catch (error) {
            throw new Error(`Failed to fetch departments: ${error.message}`);
        }
    }

    async function copyHolidaysToDepartments(sourceDeptId, targetDeptIds, token) {
        // Get holidays from source department
        const sourceResponse = await fetch(`${window.location.origin}/xapi/v1/Groups(${sourceDeptId})?$expand=OfficeHolidays`, {
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            }
        });

        if (!sourceResponse.ok) {
            throw new Error(`Failed to fetch source department: ${sourceResponse.status}`);
        }

        const sourceData = await sourceResponse.json();
        let holidays = [];
        if (sourceData.OfficeHolidays) {
            if (Array.isArray(sourceData.OfficeHolidays)) {
                holidays = sourceData.OfficeHolidays;
            } else if (sourceData.OfficeHolidays.results) {
                holidays = sourceData.OfficeHolidays.results;
            }
        }

        // Copy holidays to each target department
        const results = [];

        for (const targetDeptId of targetDeptIds) {
            try {
                // PATCH request to update target department
                const patchResponse = await fetch(`${window.location.origin}/xapi/v1/Groups(${targetDeptId})`, {
                    method: 'PATCH',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${token}`
                    },
                    body: JSON.stringify({
                        OfficeHolidays: holidays
                    })
                });

                if (!patchResponse.ok) {
                    throw new Error(`Failed to update department ${targetDeptId}: ${patchResponse.status}`);
                }

                results.push({
                    departmentId: targetDeptId,
                    success: true
                });
            } catch (error) {
                console.error(`Error updating department ${targetDeptId}:`, error);
                results.push({
                    departmentId: targetDeptId,
                    success: false,
                    error: error.message
                });
            }
        }

        // Check if any operations failed
        const failures = results.filter(r => !r.success);
        if (failures.length > 0) {
            throw new Error(`Failed to update ${failures.length} out of ${targetDeptIds.length} departments`);
        }

        return results;
    }

    // Hide the copy holidays dropdown if clicked outside
    document.addEventListener('click', function(e) {
        const copyBtn = document.getElementById('copy-holidays-btn');
        if (copyBtn) {
            const container = copyBtn.parentElement;
            if (!container.contains(e.target)) {
                const dropdown = container.querySelector('.holidays-dropdown-menu');
                if (dropdown) dropdown.style.display = 'none';
            }
        }
    });

    // Hide the import holidays dropdown if clicked outside
    document.addEventListener('click', function(e) {
        const importBtn = document.getElementById('import-holidays-btn');
        if (importBtn) {
            const container = importBtn.parentElement;
            if (!container.contains(e.target)) {
                const dropdown = container.querySelector('.holidays-dropdown-menu');
                if (dropdown) dropdown.style.display = 'none';
            }
        }
    });

    // ==================== CORE QUEUE FUNCTIONS ====================
    async function getBearerToken() {
        const token = getActiveBearerToken();
        if (token) {
            return token;
        }
        const username = prompt("Enter your 3CX admin username:");
        if (!username) throw new Error("Username required");
        const password = prompt("Enter your 3CX admin password:");
        if (!password) throw new Error("Password required");
        const response = await fetch(`${window.location.origin}/webclient/api/Login/GetAccessToken`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                SecurityCode: '',
                Username: username,
                Password: password
            })
        });
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const data = await response.json();
        if (data?.Status !== 'AuthSuccess') throw new Error(data?.Message || 'Authentication failed');
        window.__activeBearerToken = data.Token.access_token;
        return window.__activeBearerToken;
    }

    async function fetchQueues(token) {
        try {
            const response = await fetch(`${window.location.origin}/xapi/v1/Queues?%24orderby=Number&%24select=Id%2CName%2CNumber%2CIsRegistered%2CPollingStrategy%2CForwardNoAnswer%2CGroups&%24expand=Groups(%24select%3DGroupId%2CName%3B%24filter%3Dnot%20startsWith(Name%2C%27___FAVORITES___%27)%3B)`, {
                headers: { Authorization: `Bearer ${token}` }
            });
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            return await response.json();
        } catch (error) {
            throw new Error(`Failed to fetch queues: ${error.message}`);
        }
    }

    async function processQueues(queues, token) {
        const results = [];
        for (const queue of queues) {
            try {
                const response = await fetch(`${window.location.origin}/xapi/v1/Queues(${queue.Id})?%24expand=Agents%2CManagers%2CGroups(%24select%3DGroupId%2CName%2CNumber%3B%24filter%3Dnot%20startsWith(Name%2C%27___FAVORITES___%27)%3B)`, {
                    headers: { Authorization: `Bearer ${token}` }
                });
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                const details = await response.json();
                results.push({
                    Queue: { Name: queue.Name, Number: queue.Number },
                    Agents: (details.Agents || []).map(a => ({ Number: a.Number, Name: a.Name })),
                    Managers: (details.Managers || []).map(m => ({ Number: m.Number, Name: m.Name }))
                });
            } catch (error) {
                console.warn(`Skipping queue ${queue.Name}: ${error.message}`);
            }
        }
        return results;
    }

    // ==================== CSV & FILE DOWNLOAD FUNCTIONS ====================
    function downloadAsCsv(data) {
        try {
            const csvContent = convertToCSV(data);
            downloadAsFile(csvContent, 'csv');
        } catch (error) {
            showNotification(`CSV Conversion Failed: ${error.message}`, false);
        }
    }

    function convertToCSV(data) {
        const csvRows = [];
        csvRows.push('Queue Name,Queue Number,Type,User Name,User Number');
        data.forEach(queueData => {
            const queueName = `"${queueData.Queue.Name.replace(/"/g, '""')}"`;
            const queueNumber = queueData.Queue.Number;
            queueData.Agents.forEach(agent => {
                csvRows.push(
                    `${queueName},${queueNumber},Agent,` +
                    `"${agent.Name.replace(/"/g, '""')}",${agent.Number}`
                );
            });
            queueData.Managers.forEach(manager => {
                csvRows.push(
                    `${queueName},${queueNumber},Manager,` +
                    `"${manager.Name.replace(/"/g, '""')}",${manager.Number}`
                );
            });
        });
        return csvRows.join('\n');
    }

    function downloadAsFile(data, type = 'json') {
        const extension = type === 'csv' ? 'csv' : 'json';
        const mimeType = type === 'csv' ? 'text/csv' : 'application/json';
        const blob = new Blob([data], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `3cx-queues-export-${new Date().toISOString()}.${extension}`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // ==================== UTILITY NOTIFICATION FUNCTION ====================
    function showNotification(message, isSuccess) {
        if (typeof GM_notification === 'function') {
            GM_notification({
                title: '3CX Export',
                text: message,
                timeout: 5000,
                highlight: isSuccess
            });
        } else {
            alert(message);
        }
    }

    // ==================== INITIALIZATION ====================
    maybeInjectPasswordField();
    maybeInjectUserDataFields();
    maybeInjectExportButtonUI();
    maybeInjectImportHolidaysButtonUI();

    window.addEventListener('hashchange', () => {
        maybeInjectPasswordField();
        maybeInjectUserDataFields();
        maybeInjectExportButtonUI();
        maybeInjectImportHolidaysButtonUI();
        // Clear captured data when navigating away from user edit
        if (!window.location.hash.startsWith('#/office/users/edit/')) {
            capturedUserData = null;
        }
    });
})();
