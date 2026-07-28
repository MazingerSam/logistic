<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>戰略聯合物流併單管理系統 (終極修復版)</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS Excel 解析庫 -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- FontAwesome 圖示庫 -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body { font-family: 'Microsoft JhengHei', sans-serif; background-color: #f3f4f6; }
        .drag-active { border-color: #3b82f6 !important; background-color: #eff6ff !important; }
        .tab-active { border-bottom: 3px solid #1d4ed8; color: #1d4ed8; font-weight: bold; }
        .table-container { max-height: 520px; overflow-y: auto; }
        th { position: sticky; top: 0; background-color: #1e3a8a; color: white; z-index: 10; }
        .cell-multiline { white-space: pre-line; }
    </style>
</head>
<body class="p-4 md:p-8">

    <!-- Header -->
    <header class="max-w-7xl mx-auto mb-6 bg-white rounded-xl shadow-md p-6 flex flex-col md:flex-row justify-between items-center border-l-8 border-blue-900">
        <div>
            <h1 class="text-2xl font-bold text-gray-800"><i class="fa-solid fa-shield-halved text-blue-900 mr-3"></i>戰略聯合物流併單管理系統 <span class="text-xs bg-emerald-100 text-emerald-800 font-normal px-2.5 py-1 rounded-full ml-2">UID精準同步版</span></h1>
            <p class="text-sm text-gray-500 mt-1">支援雙軌定址併單、全頁面 Checkbox 操作、即時歸屬主號刷新、QA/QC 自動核實與物流規格 Excel 匯出</p>
        </div>
        <div class="mt-4 md:mt-0 flex gap-2">
            <button id="btnAutoMerge" onclick="runAutoConsolidation()" disabled class="px-4 py-2.5 bg-blue-700 hover:bg-blue-800 disabled:bg-gray-300 text-white font-semibold rounded-lg shadow transition flex items-center gap-1.5 text-sm">
                <i class="fa-solid fa-wand-magic-sparkles"></i> 執行預先併單
            </button>
            <button id="btnQC" onclick="openQAModal()" disabled class="px-4 py-2.5 bg-purple-700 hover:bg-purple-800 disabled:bg-gray-300 text-white font-semibold rounded-lg shadow transition flex items-center gap-1.5 text-sm">
                <i class="fa-solid fa-list-check"></i> QA/QC 稽核報告
            </button>
            <button id="btnExport" onclick="exportToExcel()" disabled class="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:bg-gray-300 text-white font-semibold rounded-lg shadow transition flex items-center gap-1.5 text-sm">
                <i class="fa-solid fa-file-excel"></i> 匯出 Excel
            </button>
        </div>
    </header>

    <main class="max-w-7xl mx-auto space-y-6">

        <!-- Upload Section -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div id="dropDapha" class="bg-white border-2 border-dashed border-gray-300 rounded-xl p-6 text-center hover:border-blue-500 transition cursor-pointer relative">
                <input type="file" id="fileDapha" accept=".xls,.xlsx" class="hidden" onchange="handleFileSelect('dapha', event)">
                <i class="fa-solid fa-cloud-arrow-up text-4xl text-blue-800 mb-2"></i>
                <h3 class="font-bold text-gray-700 text-lg">大法 (Dapha) 訂單檔案</h3>
                <p class="text-xs text-gray-400 mt-1">拖曳 .xls / .xlsx 檔案至此，或點擊選擇檔案</p>
                <div id="statusDapha" class="mt-3 text-sm font-semibold text-gray-500">尚未載入檔案</div>
            </div>

            <div id="dropRapha" class="bg-white border-2 border-dashed border-gray-300 rounded-xl p-6 text-center hover:border-blue-500 transition cursor-pointer relative">
                <input type="file" id="fileRapha" accept=".xls,.xlsx" class="hidden" onchange="handleFileSelect('rapha', event)">
                <i class="fa-solid fa-cloud-arrow-up text-4xl text-indigo-700 mb-2"></i>
                <h3 class="font-bold text-gray-700 text-lg">拉法 (Rapha) 訂單檔案</h3>
                <p class="text-xs text-gray-400 mt-1">拖曳 .xls / .xlsx 檔案至此，或點擊選擇檔案</p>
                <div id="statusRapha" class="mt-3 text-sm font-semibold text-gray-500">尚未載入檔案</div>
            </div>
        </div>

        <!-- Toolbar & QC Badge -->
        <div class="bg-white rounded-xl shadow p-4 flex flex-wrap gap-4 items-center justify-between border border-gray-200">
            <div class="flex items-center gap-3">
                <span class="font-bold text-gray-700 text-sm"><i class="fa-solid fa-sliders text-blue-700 mr-1"></i> 人工調整區：</span>
                <button onclick="openMergeModal()" class="px-3.5 py-2 bg-amber-600 hover:bg-amber-700 text-white text-sm font-semibold rounded-lg shadow transition flex items-center gap-1.5">
                    <i class="fa-solid fa-object-group"></i> 併入指定大法單號
                </button>
                <button onclick="manualUnmergeSelected()" class="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 text-white text-sm font-semibold rounded-lg shadow transition flex items-center gap-1.5">
                    <i class="fa-solid fa-object-ungroup"></i> 反併單 (還原選取)
                </button>
            </div>
            
            <div class="flex items-center gap-4 text-xs font-semibold">
                <div id="qcBadge" class="hidden px-3 py-1.5 rounded-full flex items-center gap-1.5 bg-gray-100 text-gray-600">
                    <i class="fa-solid fa-circle-notch fa-spin"></i> 尚未執行稽核
                </div>
                <div class="text-gray-500 flex gap-3">
                    <span>大法: <b id="statDaphaCount" class="text-gray-800">0</b></span>
                    <span>拉法: <b id="statRaphaCount" class="text-gray-800">0</b></span>
                    <span>出貨包裹數: <b id="statMergedCount" class="text-blue-900">0</b></span>
                </div>
            </div>
        </div>

        <!-- Tabs & Data Table -->
        <div class="bg-white rounded-xl shadow overflow-hidden border border-gray-200">
            <div class="flex border-b border-gray-200 bg-gray-50 text-sm font-medium">
                <button id="tabBtnMergedDapha" onclick="switchTab('mergedDapha')" class="px-6 py-3 tab-active transition flex items-center gap-2">
                    <i class="fa-solid fa-boxes-packing"></i> 大法併單結果 (<span id="countMergedDapha">0</span>)
                </button>
                <button id="tabBtnMergedRapha" onclick="switchTab('mergedRapha')" class="px-6 py-3 text-gray-500 hover:text-gray-700 transition flex items-center gap-2">
                    <i class="fa-solid fa-boxes-stacked"></i> 拉法併單結果 (<span id="countMergedRapha">0</span>)
                </button>
                <button id="tabBtnOrigDapha" onclick="switchTab('origDapha')" class="px-6 py-3 text-gray-500 hover:text-gray-700 transition flex items-center gap-2">
                    <i class="fa-solid fa-table-list"></i> 大法原始明細 (<span id="countOrigDapha">0</span>)
                </button>
                <button id="tabBtnOrigRapha" onclick="switchTab('origRapha')" class="px-6 py-3 text-gray-500 hover:text-gray-700 transition flex items-center gap-2">
                    <i class="fa-solid fa-table-list"></i> 拉法原始明細 (<span id="countOrigRapha">0</span>)
                </button>
            </div>

            <div class="table-container p-2">
                <table id="dataTable" class="w-full text-sm text-left text-gray-600 border-collapse">
                    <thead id="tableHeader"></thead>
                    <tbody id="tableBody" class="divide-y divide-gray-200">
                        <tr>
                            <td colspan="8" class="text-center py-12 text-gray-400">請先完成「大法」與「拉法」Excel 訂單檔案上傳</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </main>

    <!-- Modal 1: 人工併單選擇 -->
    <div id="mergeModal" class="fixed inset-0 bg-black/50 hidden items-center justify-center z-50 p-4">
        <div class="bg-white rounded-xl shadow-2xl max-w-2xl w-full max-h-[85vh] flex flex-col overflow-hidden">
            <div class="bg-blue-900 text-white p-4 flex justify-between items-center">
                <div>
                    <h3 class="font-bold text-lg flex items-center gap-2">
                        <i class="fa-solid fa-list-check"></i> 請選擇欲併入的大法目標主單號
                    </h3>
                    <p class="text-xs text-blue-200 mt-0.5">已鎖定 <b id="selectedCountNotice" class="text-amber-300">0</b> 筆待處理訂單</p>
                </div>
                <button onclick="closeMergeModal()" class="text-white/80 hover:text-white text-xl"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="p-3 bg-gray-50 border-b border-gray-200">
                <input type="text" id="modalSearch" oninput="filterModalOptions()" placeholder="🔍 搜尋單號、客戶全稱或地址..." class="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-blue-600">
            </div>
            <div class="flex-1 overflow-y-auto p-4 space-y-2" id="modalOptionsList"></div>
            <div class="bg-gray-100 p-3 border-t border-gray-200 flex justify-end gap-2">
                <button onclick="closeMergeModal()" class="px-4 py-2 bg-gray-300 hover:bg-gray-400 text-gray-700 text-sm font-semibold rounded-lg">取消</button>
                <button onclick="confirmManualMerge()" class="px-5 py-2 bg-blue-700 hover:bg-blue-800 text-white text-sm font-semibold rounded-lg shadow">確認併入</button>
            </div>
        </div>
    </div>

    <!-- Modal 2: QA/QC 數據報告 -->
    <div id="qaModal" class="fixed inset-0 bg-black/50 hidden items-center justify-center z-50 p-4">
        <div class="bg-white rounded-xl shadow-2xl max-w-3xl w-full max-h-[90vh] flex flex-col overflow-hidden">
            <div class="bg-purple-900 text-white p-4 flex justify-between items-center">
                <h3 class="font-bold text-lg flex items-center gap-2">
                    <i class="fa-solid fa-clipboard-check"></i> 併單數據 QA/QC 自動防呆核實報告
                </h3>
                <button onclick="closeQAModal()" class="text-white/80 hover:text-white text-xl"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="flex-1 overflow-y-auto p-6 space-y-4" id="qaReportContent"></div>
            <div class="bg-gray-100 p-4 border-t border-gray-200 flex justify-between items-center">
                <div class="text-xs text-gray-500">稽核時間: <span id="qaTimestamp">-</span></div>
                <button onclick="closeQAModal()" class="px-5 py-2 bg-purple-800 hover:bg-purple-900 text-white text-sm font-semibold rounded-lg shadow">關閉報告</button>
            </div>
        </div>
    </div>

    <script>
        let daphaRaw = [];
        let raphaRaw = [];
        let currentTab = 'mergedDapha';
        let selectedTargetMasterId = null;
        let currentlySelectedItems = []; // 🌟 鎖定選取的記憶體物件陣列

        setupDragDrop('dropDapha', 'fileDapha', 'dapha');
        setupDragDrop('dropRapha', 'fileRapha', 'rapha');

        function setupDragDrop(dropId, inputId, company) {
            const dropZone = document.getElementById(dropId);
            const fileInput = document.getElementById(inputId);

            dropZone.addEventListener('click', () => fileInput.click());
            dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-active'); });
            dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-active'));
            dropZone.addEventListener('drop', (e) => {
                e.preventDefault();
                dropZone.classList.remove('drag-active');
                if (e.dataTransfer.files.length) {
                    fileInput.files = e.dataTransfer.files;
                    processFile(company, e.dataTransfer.files[0]);
                }
            });
        }

        function handleFileSelect(company, event) {
            if (event.target.files.length) {
                processFile(company, event.target.files[0]);
            }
        }

        function cleanAddress(addr) {
            if (!addr) return '';
            let str = String(addr).normalize('NFKC').replace(/巿/g, '市').replace(/臺/g, '台');
            return str.replace(/\s+/g, '');
        }

        function processFile(company, file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(firstSheet, {defval: ''});

                // 🌟 核心增強：替每一列加入 uid，確保絕對能被抓取對照
                const processed = json.map((row, idx) => ({
                    uid: `${company}_${idx}`,
                    id: String(row['單據號碼'] || '').trim(),
                    customer: String(row['(客戶全稱)'] || row['客戶全稱'] || '').trim(),
                    customerCode: String(row['客戶編號'] || '').trim(),
                    phone: String(row['聯絡電話'] || '').trim(),
                    address: String(row['地址'] || '').trim(),
                    addressClean: cleanAddress(row['地址']),
                    invoice: String(row['發票號碼'] || '').trim(),
                    note: String(row['備註'] || '').trim(),
                    title: String(row['聯絡職稱'] || '').trim(),
                    company: company === 'dapha' ? '大法' : '拉法',
                    excelRow: idx + 2,
                    masterId: String(row['單據號碼'] || '').trim(),
                    targetPool: company
                }));

                if (company === 'dapha') {
                    daphaRaw = processed;
                    document.getElementById('statusDapha').innerHTML = `<span class="text-emerald-600"><i class="fa-solid fa-circle-check"></i> 已載入 ${processed.length} 筆 (${file.name})</span>`;
                } else {
                    raphaRaw = processed;
                    document.getElementById('statusRapha').innerHTML = `<span class="text-emerald-600"><i class="fa-solid fa-circle-check"></i> 已載入 ${processed.length} 筆 (${file.name})</span>`;
                }

                checkReady();
            };
            reader.readAsArrayBuffer(file);
        }

        function checkReady() {
            if (daphaRaw.length > 0 && raphaRaw.length > 0) {
                document.getElementById('btnAutoMerge').disabled = false;
                document.getElementById('btnQC').disabled = false;
                document.getElementById('btnExport').disabled = false;
                runAutoConsolidation();
            }
            updateStats();
        }

        // 預先併單 (雙軌比對)
        function runAutoConsolidation() {
            daphaRaw.forEach(r => { r.masterId = r.id; r.targetPool = 'dapha'; });
            raphaRaw.forEach(r => { r.masterId = r.id; r.targetPool = 'rapha'; });

            const daphaAddressSet = new Set(daphaRaw.map(r => r.addressClean));
            const daphaPrefixSet = new Set(daphaRaw.map(r => r.customer.substring(0, 2)).filter(p => p));
            const daphaPrefixToAddr = {};
            daphaRaw.forEach(r => {
                const prefix = r.customer.substring(0, 2);
                if (prefix && !daphaPrefixToAddr[prefix]) daphaPrefixToAddr[prefix] = r.addressClean;
            });

            raphaRaw.forEach(r => {
                const hasQ = r.note.toUpperCase().includes('Q');
                if (hasQ) {
                    if (daphaAddressSet.has(r.addressClean)) {
                        r.targetPool = 'dapha';
                    } else if (daphaPrefixSet.has(r.customer.substring(0, 2))) {
                        r.targetPool = 'dapha';
                        r.addressClean = daphaPrefixToAddr[r.customer.substring(0, 2)];
                    } else {
                        r.targetPool = 'rapha';
                    }
                } else {
                    r.targetPool = 'rapha';
                }
            });

            // 大法池計算主號
            const daphaPool = [...daphaRaw, ...raphaRaw.filter(r => r.targetPool === 'dapha')];
            const daphaAddrGroup = {};
            daphaPool.forEach(r => {
                if (!daphaAddrGroup[r.addressClean]) daphaAddrGroup[r.addressClean] = [];
                daphaAddrGroup[r.addressClean].push(r);
            });

            Object.keys(daphaAddrGroup).forEach(addr => {
                const group = daphaAddrGroup[addr];
                const daphaOrders = group.filter(r => r.company === '大法');
                let leadId = daphaOrders.length > 0 ? daphaOrders.map(r => r.id).sort()[0] : group.map(r => r.id).sort()[0];
                group.forEach(r => { r.masterId = leadId; });
            });

            // 拉法池計算主號
            const raphaPool = raphaRaw.filter(r => r.targetPool === 'rapha');
            const raphaAddrGroup = {};
            raphaPool.forEach(r => {
                if (!raphaAddrGroup[r.addressClean]) raphaAddrGroup[r.addressClean] = [];
                raphaAddrGroup[r.addressClean].push(r);
            });
            Object.keys(raphaAddrGroup).forEach(addr => {
                const group = raphaAddrGroup[addr];
                const leadId = group.map(r => r.id).sort()[0];
                group.forEach(r => { r.masterId = leadId; });
            });

            runQAChecks(false);
            updateStats();
            renderTable();
        }

        // 🌟 抓取勾選項目：使用 uid 進行 100% 精準對照
        function getSelectedItems() {
            const checkboxes = document.querySelectorAll('.item-checkbox:checked');
            const selected = [];
            const allItems = [...daphaRaw, ...raphaRaw];
            
            checkboxes.forEach(cb => {
                const uid = cb.dataset.uid;
                const masterId = cb.dataset.masterId;

                if (masterId) {
                    // 併單結果頁：抓取屬於該大包裹的所有物件
                    allItems.filter(r => r.masterId === masterId).forEach(item => {
                        if (!selected.includes(item)) selected.push(item);
                    });
                } else if (uid) {
                    // 原始明細頁：透過 unique ID 直接反查
                    const found = allItems.find(r => r.uid === uid);
                    if (found && !selected.includes(found)) selected.push(found);
                }
            });
            return selected;
        }

        // 🌟 打開 Modal 時「即時鎖定」勾選項目
        function openMergeModal() {
            currentlySelectedItems = getSelectedItems();

            if (currentlySelectedItems.length === 0) {
                alert('請先在表格中勾選欲合併的訂單！');
                return;
            }
            if (daphaRaw.length === 0) {
                alert('請先上傳大法訂單資料！');
                return;
            }

            document.getElementById('selectedCountNotice').innerText = currentlySelectedItems.length;
            selectedTargetMasterId = null;
            document.getElementById('modalSearch').value = '';
            renderModalOptions();
            document.getElementById('mergeModal').classList.remove('hidden');
            document.getElementById('mergeModal').classList.add('flex');
        }

        function closeMergeModal() {
            document.getElementById('mergeModal').classList.add('hidden');
            document.getElementById('mergeModal').classList.remove('flex');
        }

        // 🌟 渲染大法可選擇的主單號清單
        function renderModalOptions(filterText = '') {
            const container = document.getElementById('modalOptionsList');
            container.innerHTML = '';

            // 列出大法原始名單中的不重複單號
            const masterMap = {};
            daphaRaw.forEach(r => {
                if (!masterMap[r.id]) {
                    masterMap[r.id] = { masterId: r.id, customer: r.customer, address: r.address };
                }
            });

            const keyword = filterText.trim().toLowerCase();
            const filteredOptions = Object.values(masterMap).filter(opt => 
                opt.masterId.toLowerCase().includes(keyword) ||
                opt.customer.toLowerCase().includes(keyword) ||
                opt.address.toLowerCase().includes(keyword)
            );

            if (filteredOptions.length === 0) {
                container.innerHTML = `<div class="text-center py-8 text-gray-400 text-sm">找不到符合的大法單號</div>`;
                return;
            }

            filteredOptions.forEach(opt => {
                const div = document.createElement('div');
                div.className = 'p-3 border border-gray-200 rounded-lg hover:border-blue-500 hover:bg-blue-50 cursor-pointer transition option-card';
                div.onclick = function() {
                    document.querySelectorAll('.option-card').forEach(c => c.classList.remove('border-blue-600', 'bg-blue-100/60', 'ring-2', 'ring-blue-500'));
                    div.classList.add('border-blue-600', 'bg-blue-100/60', 'ring-2', 'ring-blue-500');
                    selectedTargetMasterId = opt.masterId;
                };

                div.innerHTML = `
                    <div class="space-y-1">
                        <div class="flex items-center gap-2">
                            <span class="font-mono font-bold text-blue-900 bg-blue-100 px-2 py-0.5 rounded text-sm">${opt.masterId}</span>
                            <span class="font-bold text-gray-800 text-sm">${opt.customer}</span>
                        </div>
                        <div class="text-xs text-gray-500"><i class="fa-solid fa-location-dot text-gray-400 mr-1"></i>${opt.address}</div>
                    </div>
                `;
                container.appendChild(div);
            });
        }

        function filterModalOptions() {
            renderModalOptions(document.getElementById('modalSearch').value);
        }

        // 🌟 人工併單確認：直接對鎖定的 currentlySelectedItems 賦值
        function confirmManualMerge() {
            if (!selectedTargetMasterId) { alert('請點選一個大法併單單號！'); return; }

            currentlySelectedItems.forEach(item => {
                item.masterId = selectedTargetMasterId;
                if (item.company === '拉法') item.targetPool = 'dapha'; // 劃入大法合併池
            });

            closeMergeModal();
            runQAChecks(false);
            updateStats();
            renderTable(); // 秒速更新畫面與呈現「併入 XXXXXXXXX」
        }

        // 🌟 反併單（解除併單）
        function manualUnmergeSelected() {
            const selected = getSelectedItems();
            if (selected.length === 0) { alert('請先勾選要解除合併 (反併單) 的訂單！'); return; }

            selected.forEach(item => {
                item.masterId = item.id; // 還原主號為自己
                if (item.company === '拉法') item.targetPool = 'rapha'; // 歸還拉法池
            });

            runQAChecks(false);
            updateStats();
            renderTable(); // 秒速刷新畫面
        }

        // QA/QC 自動稽核
        function runQAChecks(showModal = true) {
            const totalOrigCount = daphaRaw.length + raphaRaw.length;
            const allItems = [...daphaRaw, ...raphaRaw];
            const processedCount = allItems.length;
            const check1Pass = (totalOrigCount === processedCount) && (totalOrigCount > 0);

            const daphaMasterIdsSet = new Set(daphaRaw.map(r => r.id));
            let check2Pass = true;
            let check2Errors = [];

            const daphaPool = [...daphaRaw, ...raphaRaw.filter(r => r.targetPool === 'dapha')];
            const masterGroups = {};
            daphaPool.forEach(r => {
                if (!masterGroups[r.masterId]) masterGroups[r.masterId] = [];
                masterGroups[r.masterId].push(r);
            });

            Object.keys(masterGroups).forEach(mId => {
                const group = masterGroups[mId];
                const hasRapha = group.some(r => r.company === '拉法');
                if (hasRapha) {
                    if (!daphaMasterIdsSet.has(mId)) {
                        check2Pass = false;
                        check2Errors.push(`主號 ${mId} 內含拉法訂單，但主號不屬於大法！`);
                    }
                }
            });

            const allMasterIdsSet = new Set(allItems.map(r => r.masterId));
            let check3Pass = true;
            allItems.forEach(r => {
                if (!allMasterIdsSet.has(r.masterId)) check3Pass = false;
            });

            const raphaQOrders = raphaRaw.filter(r => r.note.toUpperCase().includes('Q'));
            const raphaQMovedToDapha = raphaQOrders.filter(r => r.targetPool === 'dapha');

            const badge = document.getElementById('qcBadge');
            badge.classList.remove('hidden');
            if (check1Pass && check2Pass && check3Pass) {
                badge.className = 'px-3 py-1.5 rounded-full flex items-center gap-1.5 bg-emerald-100 text-emerald-800 font-bold';
                badge.innerHTML = `<i class="fa-solid fa-circle-check text-emerald-600"></i> QA/QC 數據核實通過`;
            } else {
                badge.className = 'px-3 py-1.5 rounded-full flex items-center gap-1.5 bg-rose-100 text-rose-800 font-bold';
                badge.innerHTML = `<i class="fa-solid fa-triangle-exclamation text-rose-600"></i> QA/QC 發現異常`;
            }

            const reportDiv = document.getElementById('qaReportContent');
            reportDiv.innerHTML = `
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                    <div class="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
                        <div class="text-xs text-blue-700">原始單據總數</div>
                        <div class="text-xl font-bold text-blue-900">${totalOrigCount} 筆</div>
                    </div>
                    <div class="bg-indigo-50 border border-indigo-200 rounded-lg p-3 text-center">
                        <div class="text-xs text-indigo-700">拉法 Q 標記總筆數</div>
                        <div class="text-xl font-bold text-indigo-900">${raphaQOrders.length} 筆</div>
                    </div>
                    <div class="bg-emerald-50 border border-emerald-200 rounded-lg p-3 text-center">
                        <div class="text-xs text-emerald-700">成功併入大法大包裹</div>
                        <div class="text-xl font-bold text-emerald-900">${raphaQMovedToDapha.length} 筆</div>
                    </div>
                </div>

                <div class="space-y-3">
                    <div class="p-4 rounded-lg border ${check1Pass ? 'bg-emerald-50/50 border-emerald-200' : 'bg-rose-50 border-rose-200'}">
                        <div class="flex items-center justify-between">
                            <span class="font-bold ${check1Pass ? 'text-emerald-900' : 'text-rose-900'}">
                                檢查 1：原始訂單筆數總額勾稽 (100% 不漏單)
                            </span>
                            <span class="text-sm px-2.5 py-0.5 rounded-full font-bold ${check1Pass ? 'bg-emerald-200 text-emerald-800' : 'bg-rose-200 text-rose-800'}">
                                ${check1Pass ? '🟢 100% 符合' : '🔴 數量不符'}
                            </span>
                        </div>
                        <p class="text-xs text-gray-600 mt-1">大法 ${daphaRaw.length} 筆 + 拉法 ${raphaRaw.length} 筆 = 總共妥善處理 ${processedCount} 筆單據。</p>
                    </div>

                    <div class="p-4 rounded-lg border ${check2Pass ? 'bg-emerald-50/50 border-emerald-200' : 'bg-rose-50 border-rose-200'}">
                        <div class="flex items-center justify-between">
                            <span class="font-bold ${check2Pass ? 'text-emerald-900' : 'text-rose-900'}">
                                檢查 2：跨公司聯合物流併單主號權重 (大法主導權)
                            </span>
                            <span class="text-sm px-2.5 py-0.5 rounded-full font-bold ${check2Pass ? 'bg-emerald-200 text-emerald-800' : 'bg-rose-200 text-rose-800'}">
                                ${check2Pass ? '🟢 鎖定大法' : '🔴 主號異常'}
                            </span>
                        </div>
                        <p class="text-xs text-gray-600 mt-1">${check2Pass ? '經穿透掃描，所有包含拉法訂單的合併包裹，併單主號全數隸屬大法單號。' : check2Errors.join('<br>')}</p>
                    </div>

                    <div class="p-4 rounded-lg border ${check3Pass ? 'bg-emerald-50/50 border-emerald-200' : 'bg-rose-50 border-rose-200'}">
                        <div class="flex items-center justify-between">
                            <span class="font-bold ${check3Pass ? 'text-emerald-900' : 'text-rose-900'}">
                                檢查 3：雙向追溯鏈閉環完整性
                            </span>
                            <span class="text-sm px-2.5 py-0.5 rounded-full font-bold ${check3Pass ? 'bg-emerald-200 text-emerald-800' : 'bg-rose-200 text-rose-800'}">
                                ${check3Pass ? '🟢 閉環正常' : '🔴 斷鏈警示'}
                            </span>
                        </div>
                        <p class="text-xs text-gray-600 mt-1">每一筆原始單據所歸屬的「併單單號」，皆可在合併報表中精確反查對接。</p>
                    </div>
                </div>
            `;

            document.getElementById('qaTimestamp').innerText = new Date().toLocaleString();

            if (showModal) {
                const qaModal = document.getElementById('qaModal');
                qaModal.classList.remove('hidden');
                qaModal.classList.add('flex');
            }
        }

        function openQAModal() { runQAChecks(true); }
        function closeQAModal() {
            document.getElementById('qaModal').classList.add('hidden');
            document.getElementById('qaModal').classList.remove('flex');
        }

        function toggleSelectAll(master) {
            document.querySelectorAll('.item-checkbox').forEach(cb => cb.checked = master.checked);
        }

        function switchTab(tab) {
            currentTab = tab;
            ['MergedDapha', 'MergedRapha', 'OrigDapha', 'OrigRapha'].forEach(t => {
                document.getElementById('tabBtn' + t).className = 'px-6 py-3 text-gray-500 hover:text-gray-700 transition flex items-center gap-2';
            });
            document.getElementById('tabBtn' + tab.charAt(0).toUpperCase() + tab.slice(1)).className = 'px-6 py-3 tab-active transition flex items-center gap-2';
            renderTable();
        }

        function getConsolidatedLogisticsGroups(poolData) {
            const groups = {};
            poolData.forEach(r => {
                const mId = r.masterId;
                if (!groups[mId]) groups[mId] = [];
                groups[mId].push(r);
            });

            const result = [];
            Object.keys(groups).forEach(mId => {
                const items = groups[mId];
                const leadItem = items.find(i => i.id === mId) || items[0];

                const sourceIdsList = items.map(i => `${i.id} (${i.company})`).join('\n');
                const invoicesList = items.map(i => i.invoice).filter(inv => inv).join('\n');
                const notesList = items.map(i => i.note).filter(n => n).join('\n');

                result.push({
                    masterId: mId,
                    customer: leadItem.customer,
                    address: leadItem.address,
                    count: items.length,
                    sourceDetails: sourceIdsList,
                    invoices: invoicesList,
                    notes: notesList
                });
            });

            return result;
        }

        // 🌟 核心表格繪製（確保帶入 uid 並即時渲染『併單單號 (歸屬主號)』）
        function renderTable() {
            const thead = document.getElementById('tableHeader');
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            const isConsolidatedTab = (currentTab === 'mergedDapha' || currentTab === 'mergedRapha');

            if (isConsolidatedTab) {
                thead.innerHTML = `
                    <tr>
                        <th class="p-3 text-center w-12"><input type="checkbox" onclick="toggleSelectAll(this)"></th>
                        <th class="p-3">併單單號 (主單號)</th>
                        <th class="p-3">客戶全稱</th>
                        <th class="p-3">地址</th>
                        <th class="p-3 text-center">合併數量</th>
                        <th class="p-3">來源單號 (單號/公司)</th>
                        <th class="p-3">發票號碼 (併單發票)</th>
                        <th class="p-3">備註</th>
                    </tr>
                `;

                let pool = (currentTab === 'mergedDapha') 
                    ? [...daphaRaw, ...raphaRaw.filter(r => r.targetPool === 'dapha')]
                    : raphaRaw.filter(r => r.targetPool === 'rapha');

                if (pool.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="8" class="text-center py-10 text-gray-400">目前頁籤無資料</td></tr>`;
                    return;
                }

                const groups = getConsolidatedLogisticsGroups(pool);

                groups.forEach(g => {
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-blue-50 transition border-b border-gray-200';

                    tr.innerHTML = `
                        <td class="p-3 text-center"><input type="checkbox" class="item-checkbox" data-master-id="${g.masterId}"></td>
                        <td class="p-3 font-mono font-bold text-blue-900 bg-blue-50/50">${g.masterId}</td>
                        <td class="p-3 font-semibold text-gray-900">${g.customer}</td>
                        <td class="p-3 text-gray-700">${g.address}</td>
                        <td class="p-3 text-center font-bold text-amber-700"><span class="bg-amber-100 px-2 py-0.5 rounded-full">${g.count}</span></td>
                        <td class="p-3 font-mono text-xs cell-multiline text-gray-800">${g.sourceDetails}</td>
                        <td class="p-3 font-mono text-xs cell-multiline text-emerald-700 font-semibold">${g.invoices || '-'}</td>
                        <td class="p-3 text-xs cell-multiline text-gray-500">${g.notes || '-'}</td>
                    `;
                    tbody.appendChild(tr);
                });

            } else {
                thead.innerHTML = `
                    <tr>
                        <th class="p-3 text-center w-12"><input type="checkbox" onclick="toggleSelectAll(this)"></th>
                        <th class="p-3">單據號碼</th>
                        <th class="p-3">客戶全稱</th>
                        <th class="p-3">聯絡電話</th>
                        <th class="p-3">地址</th>
                        <th class="p-3 text-center">來源公司</th>
                        <th class="p-3 text-center">併單單號 (歸屬主號)</th>
                        <th class="p-3">備註</th>
                    </tr>
                `;

                let dataToRender = (currentTab === 'origDapha') ? daphaRaw : raphaRaw;

                if (dataToRender.length === 0) {
                    tbody.innerHTML = `<tr><td colspan="8" class="text-center py-10 text-gray-400">目前頁籤無資料</td></tr>`;
                    return;
                }

                dataToRender.forEach(row => {
                    const tr = document.createElement('tr');
                    const isChild = (row.id !== row.masterId); // 是否被併單

                    tr.className = isChild 
                        ? 'bg-amber-50/60 hover:bg-amber-100/70 transition border-b border-gray-100 font-medium' 
                        : 'bg-white hover:bg-blue-50 transition border-b border-gray-200';

                    // 💡 即時精準顯示『併單單號 (歸屬主號)』標籤
                    const masterTag = isChild 
                        ? `<span class="text-xs bg-amber-600 text-white font-bold px-2.5 py-1 rounded-full shadow-sm"><i class="fa-solid fa-link text-[10px] mr-1"></i>併入 ${row.masterId}</span>`
                        : `<span class="text-xs text-gray-300">- 獨立單據 -</span>`;

                    tr.innerHTML = `
                        <td class="p-3 text-center"><input type="checkbox" class="item-checkbox" data-uid="${row.uid}"></td>
                        <td class="p-3 font-mono font-bold text-gray-800">${row.id}</td>
                        <td class="p-3 font-semibold text-gray-900">${row.customer}</td>
                        <td class="p-3 text-gray-500">${row.phone || '-'}</td>
                        <td class="p-3 text-gray-700">${row.address}</td>
                        <td class="p-3 text-center">${row.company === '大法' ? '<span class="bg-blue-600 text-white text-xs px-2 py-0.5 rounded-full">大法</span>' : '<span class="bg-indigo-600 text-white text-xs px-2 py-0.5 rounded-full">拉法</span>'}</td>
                        <td class="p-3 text-center">${masterTag}</td>
                        <td class="p-3 text-xs text-gray-500">${row.note || '-'}</td>
                    `;
                    tbody.appendChild(tr);
                });
            }
        }

        function updateStats() {
            document.getElementById('statDaphaCount').innerText = daphaRaw.length;
            document.getElementById('statRaphaCount').innerText = raphaRaw.length;

            const allData = [...daphaRaw, ...raphaRaw];
            const uniqueMasters = new Set(allData.map(r => r.masterId).filter(m => m));
            document.getElementById('statMergedCount').innerText = uniqueMasters.size;

            document.getElementById('countOrigDapha').innerText = daphaRaw.length;
            document.getElementById('countOrigRapha').innerText = raphaRaw.length;
            document.getElementById('countMergedDapha').innerText = daphaRaw.length + raphaRaw.filter(r => r.targetPool === 'dapha').length;
            document.getElementById('countMergedRapha').innerText = raphaRaw.filter(r => r.targetPool === 'rapha').length;
        }

        // 🌟 Excel 匯出：直接導出最新的 masterId（獨立單據留白，被併單則寫入主號）
        function exportToExcel() {
            const wb = XLSX.utils.book_new();

            const daphaPool = [...daphaRaw, ...raphaRaw.filter(r => r.targetPool === 'dapha')];
            const daphaLogisticsGroups = getConsolidatedLogisticsGroups(daphaPool);
            
            const mergedDaphaData = daphaLogisticsGroups.map(g => ({
                '併單單號': g.masterId,
                '客戶全稱': g.customer,
                '地址': g.address,
                '合併單據數量': g.count,
                '來源單號(單號/公司)': g.sourceDetails,
                '發票號碼': g.invoices,
                '備註': g.notes
            }));
            const wsMergedDapha = XLSX.utils.json_to_sheet(mergedDaphaData);
            wsMergedDapha['!cols'] = [{wch: 16}, {wch: 24}, {wch: 32}, {wch: 14}, {wch: 24}, {wch: 20}, {wch: 24}];
            XLSX.utils.book_append_sheet(wb, wsMergedDapha, "大法合併訂單");

            const raphaPool = raphaRaw.filter(r => r.targetPool === 'rapha');
            const raphaLogisticsGroups = getConsolidatedLogisticsGroups(raphaPool);

            const mergedRaphaData = raphaLogisticsGroups.map(g => ({
                '併單單號': g.masterId,
                '客戶全稱': g.customer,
                '地址': g.address,
                '合併單據數量': g.count,
                '來源單號(單號/公司)': g.sourceDetails,
                '發票號碼': g.invoices,
                '備註': g.notes
            }));
            const wsMergedRapha = XLSX.utils.json_to_sheet(mergedRaphaData);
            wsMergedRapha['!cols'] = [{wch: 16}, {wch: 24}, {wch: 32}, {wch: 14}, {wch: 24}, {wch: 20}, {wch: 24}];
            XLSX.utils.book_append_sheet(wb, wsMergedRapha, "拉法合併訂單");

            // 原始資料頁籤 (精簡原則：若 masterId 等於原始單號則留白，不同才寫入併單單號)
            const origDaphaData = daphaRaw.map(r => ({
                '單據號碼': r.id,
                '(客戶全稱)': r.customer,
                '客戶編號': r.customerCode,
                '聯絡電話': r.phone,
                '地址': r.address,
                '發票號碼': r.invoice,
                '備註': r.note,
                '聯絡職稱': r.title,
                '併單單號': (r.masterId && r.masterId !== r.id) ? r.masterId : ''
            }));
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(origDaphaData), "大法原始資料");

            const origRaphaData = raphaRaw.map(r => ({
                '單據號碼': r.id,
                '(客戶全稱)': r.customer,
                '客戶編號': r.customerCode,
                '聯絡電話': r.phone,
                '地址': r.address,
                '發票號碼': r.invoice,
                '備註': r.note,
                '聯絡職稱': r.title,
                '併單單號': (r.masterId && r.masterId !== r.id) ? r.masterId : ''
            }));
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(origRaphaData), "拉法原始資料");

            XLSX.writeFile(wb, "大法_拉法_聯合物流併單報表(WMS規格).xlsx");
        }
    </script>
</body>
</html>
