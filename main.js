let workbookData = null;
        let rawData = [];
        let globalData = [];
        let columns = [];
        let lastCalcResultData = [];

        // State untuk file kedua (Merge & Append)
        let workbookData2 = null;
        let rawData2 = [];
        let columns2 = [];

        window.addEventListener('DOMContentLoaded', function() {
            try {
                const savedData = localStorage.getItem('mini_excel_raw_data');
                const savedColumns = localStorage.getItem('mini_excel_columns');
                if (savedData && savedColumns) {
                    rawData = JSON.parse(savedData);
                    columns = JSON.parse(savedColumns);
                    globalData = JSON.parse(JSON.stringify(rawData));
                    setupDropdowns();
                    renderMainTable(globalData);
                    document.getElementById('appContent').style.display = 'block';
                }
            } catch (err) {
                console.error("Gagal memuat auto-save localStorage:", err);
            }
        });

        document.getElementById('fileInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                workbookData = XLSX.read(data, { type: 'array' });

                const sheets = workbookData.SheetNames;
                if (sheets.length > 1) {
                    let options = sheets.map(s => `<option value="${s}">${s}</option>`).join('');
                    document.getElementById('sheetSelect').innerHTML = options;
                    document.getElementById('sheetSelectorPanel').classList.remove('hidden');
                    document.getElementById('appContent').style.display = 'none';
                } else {
                    document.getElementById('sheetSelectorPanel').classList.add('hidden');
                    loadSheetData(sheets[0]);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        document.getElementById('btnLoadSheet').addEventListener('click', function() {
            const selectedSheet = document.getElementById('sheetSelect').value;
            loadSheetData(selectedSheet);
        });

        function loadSheetData(sheetName) {
            const sheet = workbookData.Sheets[sheetName];
            let rawJson = XLSX.utils.sheet_to_json(sheet, {header: 1});
            if (rawJson.length === 0) { alert("Sheet kosong!"); return; }

            let firstRow = rawJson[0] || [];
            let hasValidHeader = firstRow.some(cell => cell !== undefined && cell !== "" && isNaN(cell));

            if (!hasValidHeader) {
                rawData = XLSX.utils.sheet_to_json(sheet, {header: 'A'});
            } else {
                let headerRowIndex = 0;
                for (let i = 0; i < rawJson.length; i++) {
                    let row = rawJson[i];
                    let hasText = row.some(cell => cell !== undefined && cell !== "" && isNaN(cell));
                    if (hasText) { headerRowIndex = i; break; }
                }
                rawData = XLSX.utils.sheet_to_json(sheet, {range: headerRowIndex});
            }

            globalData = JSON.parse(JSON.stringify(rawData));
            if (globalData.length === 0) { alert("Format data tidak valid!"); return; }

            columns = Object.keys(globalData[0]);

            try {
                localStorage.setItem('mini_excel_raw_data', JSON.stringify(rawData));
                localStorage.setItem('mini_excel_columns', JSON.stringify(columns));
            } catch (err) { console.error(err); }

            setupDropdowns();
            renderMainTable(globalData);
            document.getElementById('appContent').style.display = 'block';
        }

        function setupDropdowns() {
            let colOpts = columns.map(c => `<option value="${c}">${c}</option>`).join('');
            let colOptsWithBlank = '<option value="">-- Pilih Kolom --</option>' + colOpts;

            document.getElementById('criteriaField').innerHTML = colOptsWithBlank;
            document.getElementById('sortField').innerHTML = colOptsWithBlank;
            document.getElementById('cfField').innerHTML = colOptsWithBlank;
            document.getElementById('valSelect').innerHTML = colOpts;

            let pivotRowsHtml = '';
            let pivotColsHtml = '';
            columns.forEach(c => {
                pivotRowsHtml += `<div class="flex items-center gap-2"><input type="checkbox" class="pivot-row-cb w-4 h-4 text-purple-600 bg-gray-700 border-gray-500 rounded focus:ring-purple-500" value="${c}"><label class="text-xs text-gray-300 cursor-pointer">${c}</label></div>`;
                pivotColsHtml += `<div class="flex items-center gap-2"><input type="checkbox" class="pivot-col-cb w-4 h-4 text-purple-600 bg-gray-700 border-gray-500 rounded focus:ring-purple-500" value="${c}"><label class="text-xs text-gray-300 cursor-pointer">${c}</label></div>`;
            });
            document.getElementById('pivotRowsContainer').innerHTML = pivotRowsHtml;
            document.getElementById('pivotColsContainer').innerHTML = pivotColsHtml;

            document.getElementById('opCol1').innerHTML = colOpts;
            document.getElementById('opCol2').innerHTML = colOpts;
            document.getElementById('filterColSelect').innerHTML = colOpts;
            document.getElementById('ifColSelect').innerHTML = colOpts;
            document.getElementById('targetValSelect').innerHTML = colOpts;

            updateItemCheckboxes();
            document.getElementById('criteriaRowsContainer').innerHTML = "";
            addCriteriaRow();
        }

        function renderMainTable(dataArray) {
            if (dataArray.length === 0) {
                document.getElementById('mainTableContainer').innerHTML = "<p class='p-4 text-gray-400 text-sm'>Tidak ada data.</p>";
                return;
            }

            const cfField = document.getElementById('cfField').value;
            const cfRule = document.getElementById('cfRule').value;
            const cfValue = parseFloat(document.getElementById('cfValue').value);

            let html = `<table class="w-full text-sm text-left text-gray-300" id="tableMainExport"><thead class="text-xs text-gray-300 uppercase bg-gray-700 sticky-header"><tr><th class="px-4 py-3">No</th>`;
            columns.forEach(c => html += `<th class="px-4 py-3">${c}</th>`);
            html += `</tr></thead><tbody>`;

            dataArray.forEach((row, index) => {
                html += `<tr class="border-b border-gray-700 hover:bg-gray-700/50"><td class="px-4 py-3 text-gray-400">${index + 1}</td>`;
                columns.forEach(c => {
                    let val = row[c] !== undefined ? row[c] : "";
                    let cssClass = "";
                    if (cfField === c && !isNaN(cfValue)) {
                        let numVal = parseFloat(val);
                        if (cfRule === 'gt' && numVal > cfValue) cssClass = "bg-green-900/40 text-green-300";
                        if (cfRule === 'lt' && numVal < cfValue) cssClass = "bg-red-900/40 text-red-300";
                    }
                    html += `<td class="px-4 py-3 ${cssClass}">${val}</td>`;
                });
                html += `</tr>`;
            });
            html += `</tbody></table>`;
            document.getElementById('mainTableContainer').innerHTML = html;
        }

        document.getElementById('btnApplyCriteria').addEventListener('click', function() {
            let workingData = JSON.parse(JSON.stringify(rawData));
            const field = document.getElementById('criteriaField').value;
            const condition = document.getElementById('criteriaCondition').value;
            const keyword = document.getElementById('criteriaValue').value;

            if (!field) { alert("Pilih kolom!"); return; }

            workingData = workingData.filter(row => {
                let val = row[field] !== undefined ? row[field] : "";
                if (condition === 'contains') return String(val).toLowerCase().includes(keyword.toLowerCase());
                if (condition === 'equals') return String(val).toLowerCase() === keyword.toLowerCase();
                if (condition === 'gt') return parseFloat(val) > parseFloat(keyword);
                if (condition === 'lt') return parseFloat(val) < parseFloat(keyword);
                return true;
            });

            globalData = workingData;
            renderMainTable(globalData);
        });

        document.getElementById('btnApplySort').addEventListener('click', function() {
            const sortField = document.getElementById('sortField').value;
            const sortOrder = document.getElementById('sortOrder').value;
            if (!sortField) return;

            globalData.sort((a, b) => {
                let valA = a[sortField] !== undefined ? a[sortField] : "";
                let valB = b[sortField] !== undefined ? b[sortField] : "";
                if (!isNaN(valA) && !isNaN(valB)) {
                    return sortOrder === 'asc' ? valA - valB : valB - valA;
                } else {
                    return sortOrder === 'asc' ? String(valA).localeCompare(String(valB)) : String(valB).localeCompare(String(valA));
                }
            });
            renderMainTable(globalData);
        });

        document.getElementById('btnApplyCF').addEventListener('click', function() { renderMainTable(globalData); });

        document.getElementById('btnRemoveDuplicates').addEventListener('click', function() {
            if (globalData.length === 0) return;
            let seen = new Set();
            globalData = globalData.filter(row => {
                let identifier = JSON.stringify(row);
                if (seen.has(identifier)) return false;
                seen.add(identifier);
                return true;
            });
            renderMainTable(globalData);
            alert(`Duplikat dihapus! Sisa baris: ${globalData.length}`);
        });

        document.getElementById('btnResetData').addEventListener('click', function() {
            globalData = JSON.parse(JSON.stringify(rawData));
            document.getElementById('criteriaValue').value = "";
            document.getElementById('cfValue').value = "";
            renderMainTable(globalData);
        });

        document.getElementById('btnDownloadMain').addEventListener('click', function() {
            let table = document.getElementById('tableMainExport');
            if (!table) { alert("Tidak ada data!"); return; }
            let wb = XLSX.utils.table_to_book(table, {sheet: "DataUtama"});
            XLSX.writeFile(wb, "Hasil_Filter_Utama.xlsx");
        });

        // ===================== MERGE & APPEND DATA =====================

        document.getElementById('fileInput2').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                workbookData2 = XLSX.read(data, { type: 'array' });

                const sheets = workbookData2.SheetNames;
                if (sheets.length > 1) {
                    let options = sheets.map(s => `<option value="${s}">${s}</option>`).join('');
                    document.getElementById('sheetSelect2').innerHTML = options;
                    document.getElementById('sheetSelectorPanel2').classList.remove('hidden');
                    document.getElementById('mergeAppendControls').classList.add('hidden');
                } else {
                    document.getElementById('sheetSelectorPanel2').classList.add('hidden');
                    loadSheetData2(sheets[0]);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        document.getElementById('btnLoadSheet2').addEventListener('click', function() {
            const selectedSheet = document.getElementById('sheetSelect2').value;
            loadSheetData2(selectedSheet);
        });

        function loadSheetData2(sheetName) {
            const sheet = workbookData2.Sheets[sheetName];
            let rawJson = XLSX.utils.sheet_to_json(sheet, {header: 1});
            if (rawJson.length === 0) { alert("Sheet kedua kosong!"); return; }

            let firstRow = rawJson[0] || [];
            let hasValidHeader = firstRow.some(cell => cell !== undefined && cell !== "" && isNaN(cell));

            if (!hasValidHeader) {
                rawData2 = XLSX.utils.sheet_to_json(sheet, {header: 'A'});
            } else {
                let headerRowIndex = 0;
                for (let i = 0; i < rawJson.length; i++) {
                    let row = rawJson[i];
                    let hasText = row.some(cell => cell !== undefined && cell !== "" && isNaN(cell));
                    if (hasText) { headerRowIndex = i; break; }
                }
                rawData2 = XLSX.utils.sheet_to_json(sheet, {range: headerRowIndex});
            }

            if (rawData2.length === 0) { alert("Format data file kedua tidak valid!"); return; }
            columns2 = Object.keys(rawData2[0]);

            document.getElementById('file2Info').textContent = `${rawData2.length} baris, ${columns2.length} kolom (${sheetName})`;
            document.getElementById('mergeAppendControls').classList.remove('hidden');

            populateMergeDropdowns();
        }

        function populateMergeDropdowns() {
            let colOpts = columns.map(c => `<option value="${c}">${c}</option>`).join('');
            let colOpts2 = columns2.map(c => `<option value="${c}">${c}</option>`).join('');
            document.getElementById('mergeKeyMain').innerHTML = colOpts;
            document.getElementById('mergeKeySecond').innerHTML = colOpts2;

            let cbHtml = '';
            columns2.forEach(c => {
                cbHtml += `<div class="flex items-center gap-2"><input type="checkbox" class="merge-col-cb w-4 h-4 text-teal-500 bg-gray-700 border-gray-500 rounded focus:ring-teal-500" value="${c}" checked><label class="text-xs text-gray-300 cursor-pointer">${c}</label></div>`;
            });
            document.getElementById('mergeColsCheckbox').innerHTML = cbHtml;
        }

        document.getElementById('mergeModeSelect').addEventListener('change', function() {
            const mode = this.value;
            document.getElementById('panelAppend').classList.add('hidden');
            document.getElementById('panelMergeNoKey').classList.add('hidden');
            document.getElementById('panelMergeKey').classList.add('hidden');

            if (mode === 'append') document.getElementById('panelAppend').classList.remove('hidden');
            else if (mode === 'merge_nokey') document.getElementById('panelMergeNoKey').classList.remove('hidden');
            else if (mode === 'merge_key') document.getElementById('panelMergeKey').classList.remove('hidden');
        });

        document.getElementById('btnRunMergeAppend').addEventListener('click', function() {
            if (rawData2.length === 0) { alert("Muat file kedua terlebih dahulu!"); return; }
            const mode = document.getElementById('mergeModeSelect').value;
            let resultRowCount = 0;
            let resultColCount = 0;

            if (mode === 'append') {
                let unionCols = [...columns];
                columns2.forEach(c => { if (!unionCols.includes(c)) unionCols.push(c); });

                let baseRows = rawData.map(r => {
                    let o = {};
                    unionCols.forEach(c => { o[c] = r[c] !== undefined ? r[c] : ""; });
                    return o;
                });
                let appendRows = rawData2.map(r => {
                    let o = {};
                    unionCols.forEach(c => { o[c] = r[c] !== undefined ? r[c] : ""; });
                    return o;
                });

                rawData = baseRows.concat(appendRows);
                columns = unionCols;

            } else if (mode === 'merge_nokey') {
                let cols2Renamed = columns2.map(c => columns.includes(c) ? (c + "_2") : c);
                let maxLen = Math.max(rawData.length, rawData2.length);
                let merged = [];

                for (let i = 0; i < maxLen; i++) {
                    let o = {};
                    columns.forEach(c => { o[c] = rawData[i] ? (rawData[i][c] !== undefined ? rawData[i][c] : "") : ""; });
                    columns2.forEach((c, idx) => {
                        let newName = cols2Renamed[idx];
                        o[newName] = rawData2[i] ? (rawData2[i][c] !== undefined ? rawData2[i][c] : "") : "";
                    });
                    merged.push(o);
                }

                rawData = merged;
                columns = [...columns, ...cols2Renamed];

            } else if (mode === 'merge_key') {
                const keyMain = document.getElementById('mergeKeyMain').value;
                const keySecond = document.getElementById('mergeKeySecond').value;
                const innerJoin = document.getElementById('mergeInnerJoin').checked;
                const selectedCols2 = Array.from(document.querySelectorAll('.merge-col-cb:checked')).map(cb => cb.value);

                if (selectedCols2.length === 0) { alert("Pilih minimal satu kolom dari file kedua untuk digabungkan!"); return; }

                let renamedMap = {};
                selectedCols2.forEach(c => { renamedMap[c] = columns.includes(c) ? (c + "_2") : c; });

                let merged = [];
                rawData.forEach(row => {
                    let matchRow = rawData2.find(r2 => String(r2[keySecond]) === String(row[keyMain]));
                    if (!matchRow && innerJoin) return;

                    let o = { ...row };
                    selectedCols2.forEach(c => {
                        o[renamedMap[c]] = matchRow ? (matchRow[c] !== undefined ? matchRow[c] : "") : "";
                    });
                    merged.push(o);
                });

                rawData = merged;
                columns = [...columns, ...selectedCols2.map(c => renamedMap[c])];
            }

            resultRowCount = rawData.length;
            resultColCount = columns.length;

            globalData = JSON.parse(JSON.stringify(rawData));
            try {
                localStorage.setItem('mini_excel_raw_data', JSON.stringify(rawData));
                localStorage.setItem('mini_excel_columns', JSON.stringify(columns));
            } catch (err) { console.error(err); }

            setupDropdowns();
            renderMainTable(globalData);

            // Reset state file kedua
            rawData2 = []; columns2 = []; workbookData2 = null;
            document.getElementById('fileInput2').value = "";
            document.getElementById('file2Info').textContent = "";
            document.getElementById('mergeAppendControls').classList.add('hidden');

            alert(`Berhasil digabung! Tabel utama sekarang punya ${resultRowCount} baris & ${resultColCount} kolom.`);
        });

        // ===================== END MERGE & APPEND DATA =====================

        document.getElementById('calcModeSelect').addEventListener('change', function() {
            const mode = this.value;
            document.getElementById('panelOperator').classList.add('hidden');
            document.getElementById('panelFilterItem').classList.add('hidden');
            document.getElementById('multiCriteriaBox').classList.add('hidden');
            document.getElementById('conditionalIfBox').classList.add('hidden');
            document.getElementById('standardCheckboxArea').classList.add('hidden');

            if (mode === 'operator') {
                document.getElementById('panelOperator').classList.remove('hidden');
            } else {
                document.getElementById('panelFilterItem').classList.remove('hidden');
                if (mode === 'sumifs' || mode === 'countifs' || mode === 'averageifs') {
                    document.getElementById('multiCriteriaBox').classList.remove('hidden');
                } else if (mode === 'sumif' || mode === 'countif' || mode === 'averageif') {
                    document.getElementById('conditionalIfBox').classList.remove('hidden');
                } else {
                    document.getElementById('standardCheckboxArea').classList.remove('hidden');
                }
            }
        });

        function addCriteriaRow() {
            let colOpts = columns.map(c => `<option value="${c}">${c}</option>`).join('');
            let rowDiv = document.createElement('div');
            rowDiv.className = 'grid grid-cols-1 sm:grid-cols-4 gap-2 items-center bg-gray-800 p-3 rounded-lg border border-gray-700';
            rowDiv.innerHTML = `
                <div><select class="multi-col bg-gray-900 border border-gray-600 text-gray-200 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2">${colOpts}</select></div>
                <div><select class="multi-op bg-gray-900 border border-gray-600 text-gray-200 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2"><option value="equals">=</option><option value="contains">Mengandung</option><option value="gt">&gt;</option><option value="lt">&lt;</option></select></div>
                <div><input type="text" class="multi-val bg-gray-900 border border-gray-600 text-gray-200 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2" placeholder="Nilai..."></div>
                <div><button type="button" class="btnRemoveRow w-full text-white bg-red-600 hover:bg-red-700 focus:ring-4 focus:ring-red-800 font-medium rounded-lg text-xs px-3 py-2">Hapus</button></div>
            `;
            rowDiv.querySelector('.btnRemoveRow').addEventListener('click', function() { rowDiv.remove(); });
            document.getElementById('criteriaRowsContainer').appendChild(rowDiv);
        }

        document.getElementById('btnAddCriteriaRow').addEventListener('click', addCriteriaRow);

        function updateItemCheckboxes() {
            const selectedCol = document.getElementById('filterColSelect').value;
            if (!selectedCol || rawData.length === 0) return;
            const uniqueItems = [...new Set(rawData.map(row => String(row[selectedCol])))];
            let html = "";
            uniqueItems.forEach((item) => {
                html += `<div class="flex items-center gap-2"><input type="checkbox" class="item-checkbox w-4 h-4 text-amber-500 bg-gray-700 border-gray-500 rounded focus:ring-amber-500" value="${item}" checked><label class="text-xs text-gray-300 cursor-pointer">${item}</label></div>`;
            });
            document.getElementById('checkboxContainer').innerHTML = html;
        }

        document.getElementById('filterColSelect').addEventListener('change', updateItemCheckboxes);
        document.getElementById('btnSelectAll').addEventListener('click', function() { document.querySelectorAll('.item-checkbox').forEach(cb => cb.checked = true); });
        document.getElementById('btnClearAll').addEventListener('click', function() { document.querySelectorAll('.item-checkbox').forEach(cb => cb.checked = false); });

        // Custom Multi-Pivot Generator Engine (Tanpa pivottable.js)
        document.getElementById('btnGenerate').addEventListener('click', function() {
            let selectedRows = Array.from(document.querySelectorAll('.pivot-row-cb:checked')).map(cb => cb.value);
            let selectedCols = Array.from(document.querySelectorAll('.pivot-col-cb:checked')).map(cb => cb.value);
            const valCol = document.getElementById('valSelect').value;
            const agg = document.getElementById('aggSelect').value;

            if (selectedRows.length === 0) {
                alert("Pilih minimal satu Kolom Baris untuk Pivot!");
                return;
            }
            if (!valCol) {
                alert("Pilih Kolom Nilai (Value)!");
                return;
            }

            // Ekstraksi unik Kolom Samping (Cols) jika ada
            let colValues = ['Total'];
            if (selectedCols.length > 0) {
                let rawColVals = [];
                rawData.forEach(row => {
                    let keyParts = selectedCols.map(c => row[c] !== undefined ? String(row[c]) : 'N/A');
                    rawColVals.push(keyParts.join(' / '));
                });
                colValues = [...new Set(rawColVals)];
            }

            // Mapping data berdasarkan baris gabungan
            let groupedData = {};
            rawData.forEach(row => {
                let rowKeyParts = selectedRows.map(c => row[c] !== undefined ? String(row[c]) : 'N/A');
                let rowKey = rowKeyParts.join(' | ');

                let colKey = 'Total';
                if (selectedCols.length > 0) {
                    let keyParts = selectedCols.map(c => row[c] !== undefined ? String(row[c]) : 'N/A');
                    colKey = keyParts.join(' / ');
                }

                if (!groupedData[rowKey]) {
                    groupedData[rowKey] = {};
                }
                if (!groupedData[rowKey][colKey]) {
                    groupedData[rowKey][colKey] = [];
                }
                let numVal = parseFloat(row[valCol]) || 0;
                groupedData[rowKey][colKey].push(numVal);
            });

            // Helper kalkulasi agregat
            function calculateAgg(arr) {
                if (arr.length === 0) return 0;
                if (agg === 'Sum') return arr.reduce((a, b) => a + b, 0);
                if (agg === 'Count') return arr.length;
                if (agg === 'Average') return arr.reduce((a, b) => a + b, 0) / arr.length;
                if (agg === 'Max') return Math.max(...arr);
                if (agg === 'Min') return Math.min(...arr);
                return arr.reduce((a, b) => a + b, 0);
            }

            // Render Tabel Pivot Custom
            let html = `<table class="w-full text-sm text-left text-gray-300 border-collapse" id="tablePivotExport">`;
            html += `<thead class="text-xs text-purple-300 uppercase bg-gray-800 sticky-header"><tr>`;

            selectedRows.forEach(r => {
                html += `<th class="px-4 py-3 border border-gray-700">${r}</th>`;
            });
            colValues.forEach(c => {
                html += `<th class="px-4 py-3 border border-gray-700 text-right">${c}</th>`;
            });
            if (selectedCols.length > 0) {
                html += `<th class="px-4 py-3 border border-gray-700 text-right font-bold text-amber-400">Grand Total</th>`;
            }
            html += `</tr></thead><tbody>`;

            Object.keys(groupedData).forEach(rowKey => {
                html += `<tr class="border-b border-gray-700 hover:bg-gray-700/50">`;
                let rowSplits = rowKey.split(' | ');
                rowSplits.forEach(part => {
                    html += `<td class="px-4 py-3 border border-gray-700 font-medium text-gray-200">${part}</td>`;
                });

                let rowAllValues = [];
                colValues.forEach(c => {
                    let arr = groupedData[rowKey][c] || [];
                    let res = calculateAgg(arr);
                    rowAllValues = rowAllValues.concat(arr);
                    html += `<td class="px-4 py-3 border border-gray-700 text-right text-blue-400 font-medium">${res.toLocaleString()}</td>`;
                });

                if (selectedCols.length > 0) {
                    let grandRes = calculateAgg(rowAllValues);
                    html += `<td class="px-4 py-3 border border-gray-700 text-right text-amber-400 font-bold">${grandRes.toLocaleString()}</td>`;
                }
                html += `</tr>`;
            });

            html += `</tbody></table>`;
            document.getElementById('output').innerHTML = html;
        });

        document.getElementById('btnDownloadPivot').addEventListener('click', function() {
            let pivotTableElem = document.getElementById('tablePivotExport');
            if (!pivotTableElem) { alert("Tampilkan pivot dulu!"); return; }
            let wb = XLSX.utils.table_to_book(pivotTableElem, {sheet: "PivotTable"});
            XLSX.writeFile(wb, "Hasil_Pivot.xlsx");
        });

        document.getElementById('btnRunCalcTable').addEventListener('click', function() {
            const mode = document.getElementById('calcModeSelect').value;
            if (rawData.length === 0) return;
            let tableHTML = "";
            lastCalcResultData = [];

            if (mode === 'operator') {
                const col1 = document.getElementById('opCol1').value;
                const col2 = document.getElementById('opCol2').value;
                const symbol = document.getElementById('opSymbol').value;
                const resName = document.getElementById('resultColName').value || "Hasil";

                tableHTML = `<table class="w-full text-sm text-left text-gray-300" id="tableCalcExport"><thead class="text-xs text-gray-300 uppercase bg-gray-700 sticky-header"><tr><th class="px-4 py-3">No</th><th class="px-4 py-3">${col1}</th><th class="px-4 py-3 text-center">Op</th><th class="px-4 py-3">${col2}</th><th class="px-4 py-3">${resName}</th></tr></thead><tbody>`;

                rawData.forEach((row, index) => {
                    let v1 = parseFloat(row[col1]) || 0;
                    let v2 = parseFloat(row[col2]) || 0;
                    let hasil = symbol === "*" ? v1 * v2 : symbol === "/" ? (v2 !== 0 ? v1 / v2 : 0) : symbol === "+" ? v1 + v2 : v1 - v2;

                    let exportRow = { No: index + 1 };
                    exportRow[col1] = v1; exportRow['Operator'] = symbol; exportRow[col2] = v2; exportRow[resName] = hasil;
                    lastCalcResultData.push(exportRow);

                    tableHTML += `<tr class="border-b border-gray-700 hover:bg-gray-700/50"><td class="px-4 py-3 text-gray-400">${index + 1}</td><td class="px-4 py-3">${v1}</td><td class="px-4 py-3 text-center text-gray-400">${symbol}</td><td class="px-4 py-3">${v2}</td><td class="px-4 py-3 font-bold text-amber-400">${hasil}</td></tr>`;
                });
                tableHTML += `</tbody></table>`;

            } else if (mode === 'sumifs' || mode === 'countifs' || mode === 'averageifs') {
                const targetVal = document.getElementById('targetValSelect').value;
                let criteriaList = [];
                document.querySelectorAll('#criteriaRowsContainer > div').forEach(r => {
                    criteriaList.push({
                        col: r.querySelector('.multi-col').value,
                        op: r.querySelector('.multi-op').value,
                        val: r.querySelector('.multi-val').value
                    });
                });

                const filteredData = rawData.filter(row => {
                    return criteriaList.every(crit => {
                        let cellVal = row[crit.col] !== undefined ? String(row[crit.col]) : "";
                        let targetText = crit.val;
                        if (crit.op === 'equals') return cellVal.toLowerCase() === targetText.toLowerCase();
                        if (crit.op === 'contains') return cellVal.toLowerCase().includes(targetText.toLowerCase());
                        if (crit.op === 'gt') return parseFloat(cellVal) > parseFloat(targetText);
                        if (crit.op === 'lt') return parseFloat(cellVal) < parseFloat(targetText);
                        return true;
                    });
                });

                tableHTML = `<table class="w-full text-sm text-left text-gray-300" id="tableCalcExport"><thead class="text-xs text-gray-300 uppercase bg-gray-700 sticky-header"><tr><th class="px-4 py-3">No</th><th class="px-4 py-3">Target (${targetVal})</th></tr></thead><tbody>`;
                let numbers = filteredData.map(row => parseFloat(row[targetVal]) || 0);
                let resultVal = mode === 'sumifs' ? numbers.reduce((a, b) => a + b, 0) : mode === 'countifs' ? numbers.length : (numbers.length ? numbers.reduce((a, b) => a + b, 0) / numbers.length : 0);

                filteredData.forEach((row, index) => {
                    let val = parseFloat(row[targetVal]) || 0;
                    lastCalcResultData.push({ No: index + 1, [targetVal]: val });
                    tableHTML += `<tr class="border-b border-gray-700 hover:bg-gray-700/50"><td class="px-4 py-3 text-gray-400">${index + 1}</td><td class="px-4 py-3">${val}</td></tr>`;
                });

                lastCalcResultData.push({ No: "HASIL AKHIR", [targetVal]: resultVal });
                tableHTML += `<tr class="bg-gray-700 font-bold text-amber-400 border-t border-gray-600"><td class="px-4 py-3">HASIL: ${mode.toUpperCase()} (Cocok: ${filteredData.length})</td><td class="px-4 py-3">${resultVal}</td></tr></tbody></table>`;

            } else if (mode === 'sumif' || mode === 'countif' || mode === 'averageif') {
                const filterCol = document.getElementById('ifColSelect').value;
                const targetVal = document.getElementById('targetValSelect').value;
                const ifOp = document.getElementById('ifOperator').value;
                const criteriaVal = document.getElementById('ifCriteriaValue').value;

                const filteredData = rawData.filter(row => {
                    let cellVal = row[filterCol] !== undefined ? String(row[filterCol]) : "";
                    if (ifOp === 'equals') return cellVal.toLowerCase() === criteriaVal.toLowerCase();
                    if (ifOp === 'contains') return cellVal.toLowerCase().includes(criteriaVal.toLowerCase());
                    if (ifOp === 'gt') return parseFloat(cellVal) > parseFloat(criteriaVal);
                    if (ifOp === 'lt') return parseFloat(cellVal) < parseFloat(criteriaVal);
                    return true;
                });

                tableHTML = `<table class="w-full text-sm text-left text-gray-300" id="tableCalcExport"><thead class="text-xs text-gray-300 uppercase bg-gray-700 sticky-header"><tr><th class="px-4 py-3">No</th><th class="px-4 py-3">${filterCol}</th><th class="px-4 py-3">${targetVal}</th></tr></thead><tbody>`;
                let numbers = filteredData.map(row => parseFloat(row[targetVal]) || 0);
                let resultVal = mode === 'sumif' ? numbers.reduce((a, b) => a + b, 0) : mode === 'countif' ? numbers.length : (numbers.length ? numbers.reduce((a, b) => a + b, 0) / numbers.length : 0);

                filteredData.forEach((row, index) => {
                    let val = parseFloat(row[targetVal]) || 0;
                    lastCalcResultData.push({ No: index + 1, [filterCol]: row[filterCol], [targetVal]: val });
                    tableHTML += `<tr class="border-b border-gray-700 hover:bg-gray-700/50"><td class="px-4 py-3 text-gray-400">${index + 1}</td><td class="px-4 py-3">${row[filterCol]}</td><td class="px-4 py-3">${val}</td></tr>`;
                });

                lastCalcResultData.push({ No: "HASIL AKHIR", [filterCol]: mode.toUpperCase(), [targetVal]: resultVal });
                tableHTML += `<tr class="bg-gray-700 font-bold text-amber-400 border-t border-gray-600"><td colspan="2" class="px-4 py-3">HASIL: ${mode.toUpperCase()}</td><td class="px-4 py-3">${resultVal}</td></tr></tbody></table>`;

            } else {
                const filterCol = document.getElementById('filterColSelect').value;
                const targetVal = document.getElementById('targetValSelect').value;
                const checkedItems = Array.from(document.querySelectorAll('.item-checkbox:checked')).map(cb => cb.value);

                if (checkedItems.length === 0) { alert("Pilih minimal satu item!"); return; }

                const filteredData = rawData.filter(row => checkedItems.includes(String(row[filterCol])));
                let numbers = filteredData.map(row => parseFloat(row[targetVal]) || 0);
                let resultVal = mode === 'item_sum' ? numbers.reduce((a, b) => a + b, 0) : mode === 'item_avg' ? (numbers.length ? numbers.reduce((a, b) => a + b, 0) / numbers.length : 0) : mode === 'item_max' ? Math.max(...numbers) : mode === 'item_min' ? Math.min(...numbers) : numbers.length;

                tableHTML = `<table class="w-full text-sm text-left text-gray-300" id="tableCalcExport"><thead class="text-xs text-gray-300 uppercase bg-gray-700 sticky-header"><tr><th class="px-4 py-3">No</th><th class="px-4 py-3">${filterCol}</th><th class="px-4 py-3">${targetVal}</th></tr></thead><tbody>`;
                filteredData.forEach((row, index) => {
                    let val = parseFloat(row[targetVal]) || 0;
                    lastCalcResultData.push({ No: index + 1, [filterCol]: row[filterCol], [targetVal]: val });
                    tableHTML += `<tr class="border-b border-gray-700 hover:bg-gray-700/50"><td class="px-4 py-3 text-gray-400">${index + 1}</td><td class="px-4 py-3">${row[filterCol]}</td><td class="px-4 py-3">${val}</td></tr>`;
                });

                lastCalcResultData.push({ No: "HASIL AKHIR", [filterCol]: mode, [targetVal]: resultVal });
                tableHTML += `<tr class="bg-gray-700 font-bold text-amber-400 border-t border-gray-600"><td colspan="2" class="px-4 py-3">TOTAL (${filteredData.length} baris):</td><td class="px-4 py-3">${resultVal}</td></tr></tbody></table>`;
            }

            document.getElementById('calcOutputTable').innerHTML = tableHTML;
        });

        document.getElementById('btnDownloadCalc').addEventListener('click', function() {
            if (lastCalcResultData.length === 0) { alert("Jalankan kalkulator dulu!"); return; }
            let worksheet = XLSX.utils.json_to_sheet(lastCalcResultData);
            let workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "HasilKalkulasi");
            XLSX.writeFile(workbook, "Hasil_Kalkulator_Excel.xlsx");
        });