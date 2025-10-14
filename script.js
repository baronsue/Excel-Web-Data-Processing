// 全局变量
let workbook = null;
let worksheet = null;
let currentData = [];
let originalData = [];
let chart = null;
let tables = []; // 存储多个表
let currentTableIndex = 0; // 当前选中的表索引
let mergePreviewData = null; // 合并预览数据

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    // 主题初始化
    initializeTheme();
});

function initializeEventListeners() {
    const fileInput = document.getElementById('fileInput');
    const uploadArea = document.getElementById('uploadArea');
    
    // 文件输入事件
    fileInput.addEventListener('change', handleFileSelect);
    
    // 拖拽事件
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleFileDrop);
    uploadArea.addEventListener('click', () => fileInput.click());
}

// 文件处理函数
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        processFile(file);
    }
}

function handleDragOver(event) {
    event.preventDefault();
    event.currentTarget.classList.add('dragover');
}

function handleDragLeave(event) {
    event.currentTarget.classList.remove('dragover');
}

function handleFileDrop(event) {
    event.preventDefault();
    event.currentTarget.classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function processFile(file) {
    showLoading(true);
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            
            // 处理所有工作表
            const tableData = {
                name: file.name.replace(/\.[^/.]+$/, ""), // 移除文件扩展名
                fileName: file.name,
                sheets: []
            };
            
            workbook.SheetNames.forEach((sheetName, index) => {
                const sheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                tableData.sheets.push({
                    name: sheetName,
                    data: sheetData,
                    originalData: JSON.parse(JSON.stringify(sheetData))
                });
            });
            
            // 添加到表列表
            tables.push(tableData);
            
            // 设置当前表为第一个工作表
            if (tableData.sheets.length > 0) {
                currentTableIndex = tables.length - 1;
                currentData = tableData.sheets[0].data;
                originalData = tableData.sheets[0].originalData;
            }
            
            displayTablesList();
            displayData();
            showToolbar();
            showTablesSection();
            showAnalysisSection();
            updateStatistics();
            updateMergeSelectors();
            
        } catch (error) {
            alert('文件处理失败: ' + error.message);
        } finally {
            showLoading(false);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function displayData() {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    // 清空表格
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    
    if (currentData.length === 0) return;
    
    // 创建表头
    const headerRow = document.createElement('tr');
    const headers = currentData[0] || [];
    
    headers.forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header || `列${index + 1}`;
        th.contentEditable = true;
        th.addEventListener('blur', () => updateHeader(index, th.textContent));
        headerRow.appendChild(th);
    });
    
    // 添加选择列
    const selectTh = document.createElement('th');
    selectTh.innerHTML = '<input type="checkbox" id="selectAll" onchange="toggleSelectAll()">';
    headerRow.appendChild(selectTh);
    tableHead.appendChild(headerRow);
    
    // 创建数据行
    for (let i = 1; i < currentData.length; i++) {
        const row = document.createElement('tr');
        const rowData = currentData[i] || [];
        
        headers.forEach((header, colIndex) => {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.value = rowData[colIndex] || '';
            input.addEventListener('change', () => updateCell(i, colIndex, input.value));
            td.appendChild(input);
            row.appendChild(td);
        });
        
        // 添加选择框
        const selectTd = document.createElement('td');
        selectTd.innerHTML = `<input type="checkbox" class="row-select" onchange="updateRowSelection()">`;
        row.appendChild(selectTd);
        tableBody.appendChild(row);
    }
    
    updateTableInfo();
}

function updateTableInfo() {
    document.getElementById('rowCount').textContent = `${currentData.length - 1} 行`;
    document.getElementById('colCount').textContent = `${currentData[0] ? currentData[0].length : 0} 列`;
}

// 数据编辑函数
function updateCell(row, col, value) {
    if (!currentData[row]) {
        currentData[row] = [];
    }
    currentData[row][col] = value;
}

function updateHeader(col, value) {
    if (currentData[0]) {
        currentData[0][col] = value;
    }
}

// 工具栏功能
function addColumn() {
    const colName = prompt('请输入新列名:', `列${currentData[0].length + 1}`);
    if (colName) {
        // 更新表头
        if (currentData[0]) {
            currentData[0].push(colName);
        }
        
        // 为每行添加空值
        for (let i = 1; i < currentData.length; i++) {
            if (!currentData[i]) {
                currentData[i] = [];
            }
            currentData[i].push('');
        }
        
        displayData();
    }
}

function addRow() {
    const newRow = new Array(currentData[0] ? currentData[0].length : 0).fill('');
    currentData.push(newRow);
    displayData();
}

function deleteSelected() {
    const checkboxes = document.querySelectorAll('.row-select:checked');
    const rowsToDelete = Array.from(checkboxes).map(cb => 
        Array.from(cb.closest('tr').parentNode.children).indexOf(cb.closest('tr'))
    );
    
    if (rowsToDelete.length === 0) {
        alert('请先选择要删除的行');
        return;
    }
    
    if (confirm(`确定要删除选中的 ${rowsToDelete.length} 行吗？`)) {
        // 从后往前删除，避免索引变化
        rowsToDelete.sort((a, b) => b - a).forEach(index => {
            currentData.splice(index, 1);
        });
        displayData();
    }
}

function sortData() {
    const colIndex = prompt('请输入要排序的列号 (从1开始):');
    if (colIndex && !isNaN(colIndex)) {
        const index = parseInt(colIndex) - 1;
        if (index >= 0 && index < currentData[0].length) {
            const header = currentData[0];
            const dataRows = currentData.slice(1);
            
            dataRows.sort((a, b) => {
                const valA = a[index] || '';
                const valB = b[index] || '';
                return valA.toString().localeCompare(valB.toString());
            });
            
            currentData = [header, ...dataRows];
            displayData();
        }
    }
}

// 选择功能
function toggleSelectAll() {
    const selectAll = document.getElementById('selectAll');
    const checkboxes = document.querySelectorAll('.row-select');
    
    checkboxes.forEach(cb => {
        cb.checked = selectAll.checked;
    });
    
    updateRowSelection();
}

function updateRowSelection() {
    const checkboxes = document.querySelectorAll('.row-select');
    const checkedBoxes = document.querySelectorAll('.row-select:checked');
    
    checkboxes.forEach((cb, index) => {
        const row = cb.closest('tr');
        if (cb.checked) {
            row.classList.add('selected-row');
        } else {
            row.classList.remove('selected-row');
        }
    });
    
    // 更新全选状态
    const selectAll = document.getElementById('selectAll');
    if (selectAll) {
        selectAll.checked = checkedBoxes.length === checkboxes.length;
        selectAll.indeterminate = checkedBoxes.length > 0 && checkedBoxes.length < checkboxes.length;
    }
}

// 导出功能
function exportExcel() {
    if (!currentData || currentData.length === 0) {
        alert('没有数据可导出');
        return;
    }
    
    try {
        // 创建工作簿
        const wb = XLSX.utils.book_new();
        
        // 转换数据为工作表
        const ws = XLSX.utils.aoa_to_sheet(currentData);
        
        // 添加工作表到工作簿
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        
        // 导出文件
        const fileName = `excel_data_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        
        alert('文件导出成功！');
    } catch (error) {
        alert('导出失败: ' + error.message);
    }
}

// 分析功能
function updateStatistics() {
    if (!currentData || currentData.length <= 1) return;
    
    const stats = document.getElementById('statistics');
    const headers = currentData[0] || [];
    const dataRows = currentData.slice(1);
    
    let html = '';
    
    // 基本统计
    html += `<div class="stat-item">
        <span class="stat-label">总行数:</span>
        <span class="stat-value">${dataRows.length}</span>
    </div>`;
    
    html += `<div class="stat-item">
        <span class="stat-label">总列数:</span>
        <span class="stat-value">${headers.length}</span>
    </div>`;
    
    // 数值列统计
    headers.forEach((header, colIndex) => {
        const values = dataRows.map(row => row[colIndex]).filter(val => val !== '' && val !== null && val !== undefined);
        const numericValues = values.filter(val => !isNaN(parseFloat(val))).map(val => parseFloat(val));
        
        if (numericValues.length > 0) {
            const sum = numericValues.reduce((a, b) => a + b, 0);
            const avg = sum / numericValues.length;
            const min = Math.min(...numericValues);
            const max = Math.max(...numericValues);
            
            html += `<div class="stat-item">
                <span class="stat-label">${header} (数值):</span>
                <span class="stat-value">平均: ${avg.toFixed(2)}</span>
            </div>`;
        } else {
            html += `<div class="stat-item">
                <span class="stat-label">${header} (文本):</span>
                <span class="stat-value">${values.length} 个值</span>
            </div>`;
        }
    });
    
    stats.innerHTML = html;
}

function generateChart() {
    if (!currentData || currentData.length <= 1) {
        alert('没有数据可生成图表');
        return;
    }
    
    const colIndex = prompt('请输入要生成图表的列号 (从1开始):');
    if (colIndex && !isNaN(colIndex)) {
        const index = parseInt(colIndex) - 1;
        if (index >= 0 && index < currentData[0].length) {
            const header = currentData[0][index];
            const values = currentData.slice(1).map(row => row[index]).filter(val => val !== '' && val !== null);
            
            // 统计频率
            const frequency = {};
            values.forEach(val => {
                frequency[val] = (frequency[val] || 0) + 1;
            });
            
            const ctx = document.getElementById('chartCanvas').getContext('2d');
            
            if (chart) {
                chart.destroy();
            }
            
            chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: Object.keys(frequency),
                    datasets: [{
                        label: header,
                        data: Object.values(frequency),
                        backgroundColor: 'rgba(102, 126, 234, 0.8)',
                        borderColor: 'rgba(102, 126, 234, 1)',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    plugins: {
                        title: {
                            display: true,
                            text: `${header} 分布图`
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
        }
    }
}

// 筛选功能
function applyFilter() {
    const column = document.getElementById('filterColumn').value;
    const operator = document.getElementById('filterOperator').value;
    const value = document.getElementById('filterValue').value;
    
    if (!column || !value) {
        alert('请选择列并输入筛选值');
        return;
    }
    
    const colIndex = parseInt(column);
    const filteredData = [currentData[0]]; // 保留表头
    
    for (let i = 1; i < currentData.length; i++) {
        const cellValue = currentData[i][colIndex] || '';
        let include = false;
        
        switch (operator) {
            case 'contains':
                include = cellValue.toString().toLowerCase().includes(value.toLowerCase());
                break;
            case 'equals':
                include = cellValue.toString() === value;
                break;
            case 'greater':
                include = parseFloat(cellValue) > parseFloat(value);
                break;
            case 'less':
                include = parseFloat(cellValue) < parseFloat(value);
                break;
        }
        
        if (include) {
            filteredData.push(currentData[i]);
        }
    }
    
    currentData = filteredData;
    displayData();
    updateStatistics();
}

function clearFilter() {
    currentData = JSON.parse(JSON.stringify(originalData));
    displayData();
    updateStatistics();
    
    // 清空筛选控件
    document.getElementById('filterColumn').value = '';
    document.getElementById('filterValue').value = '';
}

// 初始化筛选列选项
function initializeFilterColumns() {
    const select = document.getElementById('filterColumn');
    select.innerHTML = '<option value="">选择列</option>';
    
    if (currentData && currentData[0]) {
        currentData[0].forEach((header, index) => {
            const option = document.createElement('option');
            option.value = index;
            option.textContent = header || `列${index + 1}`;
            select.appendChild(option);
        });
    }
}

// 多表管理功能
function displayTablesList() {
    const tablesList = document.getElementById('tablesList');
    tablesList.innerHTML = '';
    
    tables.forEach((table, tableIndex) => {
        table.sheets.forEach((sheet, sheetIndex) => {
            const tableCard = document.createElement('div');
            tableCard.className = 'table-card';
            if (tableIndex === currentTableIndex && sheetIndex === 0) {
                tableCard.classList.add('active');
            }
            
            const isActive = tableIndex === currentTableIndex && sheetIndex === 0;
            const rowCount = sheet.data.length - 1;
            const colCount = sheet.data[0] ? sheet.data[0].length : 0;
            
            tableCard.innerHTML = `
                <div class="table-card-header">
                    <div class="table-card-title">${table.name} - ${sheet.name}</div>
                    <div class="table-card-actions">
                        <button class="tool-btn" onclick="switchToTable(${tableIndex}, ${sheetIndex})" title="切换到该表">👁️</button>
                        <button class="tool-btn" onclick="deleteTable(${tableIndex}, ${sheetIndex})" title="删除该表">🗑️</button>
                    </div>
                </div>
                <div class="table-card-info">
                    <span>${rowCount} 行</span>
                    <span>${colCount} 列</span>
                </div>
                <div class="table-card-preview">
                    ${generateTablePreview(sheet.data)}
                </div>
            `;
            
            tablesList.appendChild(tableCard);
        });
    });
}

function generateTablePreview(data) {
    if (!data || data.length === 0) return '<p>无数据</p>';
    
    const maxRows = Math.min(3, data.length);
    const maxCols = Math.min(5, data[0] ? data[0].length : 0);
    
    let html = '<table>';
    
    for (let i = 0; i < maxRows; i++) {
        html += '<tr>';
        for (let j = 0; j < maxCols; j++) {
            const cellValue = data[i] && data[i][j] ? data[i][j] : '';
            if (i === 0) {
                html += `<th>${cellValue}</th>`;
            } else {
                html += `<td>${cellValue}</td>`;
            }
        }
        html += '</tr>';
    }
    
    html += '</table>';
    return html;
}

function switchToTable(tableIndex, sheetIndex) {
    if (tableIndex >= 0 && tableIndex < tables.length && 
        sheetIndex >= 0 && sheetIndex < tables[tableIndex].sheets.length) {
        
        currentTableIndex = tableIndex;
        const sheet = tables[tableIndex].sheets[sheetIndex];
        currentData = sheet.data;
        originalData = sheet.originalData;
        
        displayData();
        displayTablesList();
        updateStatistics();
        updateMergeSelectors();
    }
}

function deleteTable(tableIndex, sheetIndex) {
    if (tables[tableIndex].sheets.length === 1) {
        // 如果只有一个工作表，删除整个表
        if (confirm(`确定要删除表 "${tables[tableIndex].name}" 吗？`)) {
            tables.splice(tableIndex, 1);
            if (tables.length === 0) {
                // 如果没有表了，重置状态
                currentData = [];
                originalData = [];
                currentTableIndex = 0;
                hideAllSections();
            } else {
                // 切换到第一个表
                switchToTable(0, 0);
            }
        }
    } else {
        // 如果有多个工作表，只删除当前工作表
        if (confirm(`确定要删除工作表 "${tables[tableIndex].sheets[sheetIndex].name}" 吗？`)) {
            tables[tableIndex].sheets.splice(sheetIndex, 1);
            // 切换到第一个工作表
            switchToTable(tableIndex, 0);
        }
    }
    
    displayTablesList();
    updateMergeSelectors();
}

function addNewTable() {
    document.getElementById('fileInput').click();
}

// 表合并功能
function updateMergeSelectors() {
    const leftTableSelect = document.getElementById('leftTableSelect');
    const rightTableSelect = document.getElementById('rightTableSelect');
    
    // 清空选项
    leftTableSelect.innerHTML = '<option value="">选择左表</option>';
    rightTableSelect.innerHTML = '<option value="">选择右表</option>';
    
    // 添加表选项
    tables.forEach((table, tableIndex) => {
        table.sheets.forEach((sheet, sheetIndex) => {
            const optionText = `${table.name} - ${sheet.name}`;
            const optionValue = `${tableIndex}-${sheetIndex}`;
            
            const leftOption = document.createElement('option');
            leftOption.value = optionValue;
            leftOption.textContent = optionText;
            leftTableSelect.appendChild(leftOption);
            
            const rightOption = document.createElement('option');
            rightOption.value = optionValue;
            rightOption.textContent = optionText;
            rightTableSelect.appendChild(rightOption);
        });
    });
    
    // 更新关联列选择器
    updateKeyColumnSelectors();
}

function updateKeyColumnSelectors() {
    const leftTableSelect = document.getElementById('leftTableSelect');
    const rightTableSelect = document.getElementById('rightTableSelect');
    const leftKeyColumn = document.getElementById('leftKeyColumn');
    const rightKeyColumn = document.getElementById('rightKeyColumn');
    
    // 清空关联列选项
    leftKeyColumn.innerHTML = '<option value="">选择关联列</option>';
    rightKeyColumn.innerHTML = '<option value="">选择关联列</option>';
    
    // 左表关联列
    if (leftTableSelect.value) {
        const [tableIndex, sheetIndex] = leftTableSelect.value.split('-').map(Number);
        const sheet = tables[tableIndex].sheets[sheetIndex];
        if (sheet.data && sheet.data[0]) {
            sheet.data[0].forEach((header, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = header || `列${index + 1}`;
                leftKeyColumn.appendChild(option);
            });
        }
    }
    
    // 右表关联列
    if (rightTableSelect.value) {
        const [tableIndex, sheetIndex] = rightTableSelect.value.split('-').map(Number);
        const sheet = tables[tableIndex].sheets[sheetIndex];
        if (sheet.data && sheet.data[0]) {
            sheet.data[0].forEach((header, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = header || `列${index + 1}`;
                rightKeyColumn.appendChild(option);
            });
        }
    }
}

function showMergeSection() {
    document.getElementById('mergeSection').style.display = 'block';
    updateMergeSelectors();
}

// 显示/隐藏函数
function showLoading(show) {
    document.getElementById('loading').style.display = show ? 'flex' : 'none';
}

function showToolbar() {
    document.getElementById('toolbar').style.display = 'flex';
}

// 主题切换
function initializeTheme() {
    try {
        const saved = localStorage.getItem('theme') || 'standard';
        if (saved === 'glass') {
            document.body.classList.add('theme-glass');
            const btn = document.getElementById('themeToggleBtn');
            if (btn) btn.textContent = '🎨 切换标准主题';
        }
    } catch (e) {}
}

function toggleTheme() {
    const isGlass = document.body.classList.toggle('theme-glass');
    try {
        localStorage.setItem('theme', isGlass ? 'glass' : 'standard');
    } catch (e) {}
    const btn = document.getElementById('themeToggleBtn');
    if (btn) btn.textContent = isGlass ? '🎨 切换标准主题' : '🎨 切换玻璃主题';
}

function showTablesSection() {
    document.getElementById('tablesSection').style.display = 'block';
}

function showAnalysisSection() {
    document.getElementById('analysisSection').style.display = 'block';
    document.getElementById('tableSection').style.display = 'block';
    initializeFilterColumns();
}

function hideAllSections() {
    document.getElementById('tablesSection').style.display = 'none';
    document.getElementById('mergeSection').style.display = 'none';
    document.getElementById('mergePreviewSection').style.display = 'none';
    document.getElementById('tableSection').style.display = 'none';
    document.getElementById('analysisSection').style.display = 'none';
    document.getElementById('toolbar').style.display = 'none';
}

// 表合并核心功能
function previewMerge() {
    const leftTableSelect = document.getElementById('leftTableSelect');
    const rightTableSelect = document.getElementById('rightTableSelect');
    const leftKeyColumn = document.getElementById('leftKeyColumn');
    const rightKeyColumn = document.getElementById('rightKeyColumn');
    const joinTypeRadio = document.querySelector('input[name="joinType"]:checked');
    
    // 验证输入
    if (!leftTableSelect.value || !rightTableSelect.value || 
        !leftKeyColumn.value || !rightKeyColumn.value) {
        alert('请选择要合并的表和关联列');
        return;
    }
    
    if (!joinTypeRadio) {
        alert('请选择合并类型');
        return;
    }
    
    if (leftTableSelect.value === rightTableSelect.value) {
        alert('不能选择相同的表进行合并');
        return;
    }
    
    const [leftTableIndex, leftSheetIndex] = leftTableSelect.value.split('-').map(Number);
    const [rightTableIndex, rightSheetIndex] = rightTableSelect.value.split('-').map(Number);
    
    // 验证表索引
    if (leftTableIndex < 0 || leftTableIndex >= tables.length ||
        rightTableIndex < 0 || rightTableIndex >= tables.length) {
        alert('选择的表不存在');
        return;
    }
    
    if (leftSheetIndex < 0 || leftSheetIndex >= tables[leftTableIndex].sheets.length ||
        rightSheetIndex < 0 || rightSheetIndex >= tables[rightTableIndex].sheets.length) {
        alert('选择的工作表不存在');
        return;
    }
    
    const leftTable = tables[leftTableIndex].sheets[leftSheetIndex];
    const rightTable = tables[rightTableIndex].sheets[rightSheetIndex];
    
    // 验证表数据
    if (!leftTable.data || leftTable.data.length === 0) {
        alert('左表没有数据');
        return;
    }
    
    if (!rightTable.data || rightTable.data.length === 0) {
        alert('右表没有数据');
        return;
    }
    
    const leftKeyCol = parseInt(leftKeyColumn.value);
    const rightKeyCol = parseInt(rightKeyColumn.value);
    const joinType = joinTypeRadio.value;
    
    // 验证关联列索引
    if (leftKeyCol < 0 || leftKeyCol >= leftTable.data[0].length) {
        alert('左表关联列索引无效');
        return;
    }
    
    if (rightKeyCol < 0 || rightKeyCol >= rightTable.data[0].length) {
        alert('右表关联列索引无效');
        return;
    }
    
    try {
        // 执行合并
        const mergedData = performJoin(leftTable.data, rightTable.data, leftKeyCol, rightKeyCol, joinType);
        
        // 显示预览
        displayMergePreview(mergedData, leftTable, rightTable, joinType);
    } catch (error) {
        console.error('合并过程中发生错误:', error);
        alert('合并过程中发生错误: ' + error.message);
    }
}

function performJoin(leftData, rightData, leftKeyCol, rightKeyCol, joinType) {
    if (!leftData || !rightData || leftData.length === 0 || rightData.length === 0) {
        return [];
    }
    
    const leftHeaders = leftData[0] || [];
    const rightHeaders = rightData[0] || [];
    const leftRows = leftData.slice(1);
    const rightRows = rightData.slice(1);
    
    // 创建左表和右表的索引映射
    const leftIndex = new Map();
    const rightIndex = new Map();
    
    // 构建左表索引
    leftRows.forEach((row, index) => {
        const key = row[leftKeyCol];
        if (key !== undefined && key !== null && key !== '') {
            if (!leftIndex.has(key)) {
                leftIndex.set(key, []);
            }
            leftIndex.get(key).push(index);
        }
    });
    
    // 构建右表索引
    rightRows.forEach((row, index) => {
        const key = row[rightKeyCol];
        if (key !== undefined && key !== null && key !== '') {
            if (!rightIndex.has(key)) {
                rightIndex.set(key, []);
            }
            rightIndex.get(key).push(index);
        }
    });
    
    // 合并后的表头
    const mergedHeaders = [...leftHeaders];
    rightHeaders.forEach((header, index) => {
        if (index !== rightKeyCol) {
            mergedHeaders.push(header);
        }
    });
    
    const mergedRows = [];
    const processedKeys = new Set();
    
    // LEFT JOIN 或 INNER JOIN
    if (joinType === 'left' || joinType === 'inner') {
        leftRows.forEach(leftRow => {
            const key = leftRow[leftKeyCol];
            const rightMatches = rightIndex.get(key) || [];
            
            if (rightMatches.length > 0) {
                processedKeys.add(key);
                rightMatches.forEach(rightRowIndex => {
                    const rightRow = rightRows[rightRowIndex];
                    const mergedRow = [...leftRow];
                    rightHeaders.forEach((header, colIndex) => {
                        if (colIndex !== rightKeyCol) {
                            mergedRow.push(rightRow[colIndex] || '');
                        }
                    });
                    mergedRows.push(mergedRow);
                });
            } else if (joinType === 'left') {
                // LEFT JOIN: 左表有但右表没有的行
                const mergedRow = [...leftRow];
                rightHeaders.forEach((header, colIndex) => {
                    if (colIndex !== rightKeyCol) {
                        mergedRow.push('');
                    }
                });
                mergedRows.push(mergedRow);
            }
        });
    }
    
    // RIGHT JOIN
    if (joinType === 'right') {
        rightRows.forEach(rightRow => {
            const key = rightRow[rightKeyCol];
            const leftMatches = leftIndex.get(key) || [];
            
            if (leftMatches.length > 0) {
                processedKeys.add(key);
                leftMatches.forEach(leftRowIndex => {
                    const leftRow = leftRows[leftRowIndex];
                    const mergedRow = [...leftRow];
                    rightHeaders.forEach((header, colIndex) => {
                        if (colIndex !== rightKeyCol) {
                            mergedRow.push(rightRow[colIndex] || '');
                        }
                    });
                    mergedRows.push(mergedRow);
                });
            } else {
                // RIGHT JOIN: 右表有但左表没有的行
                const mergedRow = new Array(leftHeaders.length).fill('');
                mergedRow[leftKeyCol] = key;
                rightHeaders.forEach((header, colIndex) => {
                    if (colIndex !== rightKeyCol) {
                        mergedRow.push(rightRow[colIndex] || '');
                    }
                });
                mergedRows.push(mergedRow);
            }
        });
    }
    
    // FULL OUTER JOIN
    if (joinType === 'full') {
        // 处理所有匹配的行
        leftRows.forEach(leftRow => {
            const key = leftRow[leftKeyCol];
            const rightMatches = rightIndex.get(key) || [];
            
            if (rightMatches.length > 0) {
                processedKeys.add(key);
                rightMatches.forEach(rightRowIndex => {
                    const rightRow = rightRows[rightRowIndex];
                    const mergedRow = [...leftRow];
                    rightHeaders.forEach((header, colIndex) => {
                        if (colIndex !== rightKeyCol) {
                            mergedRow.push(rightRow[colIndex] || '');
                        }
                    });
                    mergedRows.push(mergedRow);
                });
            } else {
                // 左表有但右表没有的行
                const mergedRow = [...leftRow];
                rightHeaders.forEach((header, colIndex) => {
                    if (colIndex !== rightKeyCol) {
                        mergedRow.push('');
                    }
                });
                mergedRows.push(mergedRow);
            }
        });
        
        // 处理右表有但左表没有的行
        rightRows.forEach(rightRow => {
            const key = rightRow[rightKeyCol];
            if (!processedKeys.has(key)) {
                const mergedRow = new Array(leftHeaders.length).fill('');
                mergedRow[leftKeyCol] = key;
                rightHeaders.forEach((header, colIndex) => {
                    if (colIndex !== rightKeyCol) {
                        mergedRow.push(rightRow[colIndex] || '');
                    }
                });
                mergedRows.push(mergedRow);
            }
        });
    }
    
    return [mergedHeaders, ...mergedRows];
}

function displayMergePreview(mergedData, leftTable, rightTable, joinType) {
    const previewSection = document.getElementById('mergePreviewSection');
    const previewHead = document.getElementById('mergePreviewHead');
    const previewBody = document.getElementById('mergePreviewBody');
    
    previewSection.style.display = 'block';
    
    // 清空预览表格
    previewHead.innerHTML = '';
    previewBody.innerHTML = '';
    
    if (mergedData.length === 0) {
        previewBody.innerHTML = '<tr><td colspan="100%">没有匹配的数据</td></tr>';
        return;
    }
    
    // 创建表头
    const headerRow = document.createElement('tr');
    mergedData[0].forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header || `列${index + 1}`;
        headerRow.appendChild(th);
    });
    previewHead.appendChild(headerRow);
    
    // 创建数据行
    for (let i = 1; i < Math.min(mergedData.length, 101); i++) { // 最多显示100行
        const row = document.createElement('tr');
        const rowData = mergedData[i] || [];
        
        rowData.forEach((cell, colIndex) => {
            const td = document.createElement('td');
            td.textContent = cell || '';
            row.appendChild(td);
        });
        previewBody.appendChild(row);
    }
    
    // 更新预览信息
    const rowCount = mergedData.length - 1;
    const colCount = mergedData[0] ? mergedData[0].length : 0;
    
    document.getElementById('previewRowCount').textContent = `${rowCount} 行`;
    document.getElementById('previewColCount').textContent = `${colCount} 列`;
    
    // 计算匹配行数
    const leftKeyCol = parseInt(document.getElementById('leftKeyColumn').value);
    const rightKeyCol = parseInt(document.getElementById('rightKeyColumn').value);
    
    // 根据join类型计算匹配行数
    let matchedRows = 0;
    const joinType = document.querySelector('input[name="joinType"]:checked').value;
    
    if (joinType === 'inner') {
        // INNER JOIN: 只计算完全匹配的行
        matchedRows = mergedData.slice(1).filter(row => 
            row[leftKeyCol] !== '' && row[leftKeyCol] !== null && row[leftKeyCol] !== undefined &&
            row[leftKeyCol + (rightKeyCol < leftKeyCol ? 0 : rightKeyCol - 1)] !== '' &&
            row[leftKeyCol + (rightKeyCol < leftKeyCol ? 0 : rightKeyCol - 1)] !== null &&
            row[leftKeyCol + (rightKeyCol < leftKeyCol ? 0 : rightKeyCol - 1)] !== undefined
        ).length;
    } else if (joinType === 'left') {
        // LEFT JOIN: 计算左表有匹配的行数
        const leftTableSelect = document.getElementById('leftTableSelect');
        const [leftTableIndex, leftSheetIndex] = leftTableSelect.value.split('-').map(Number);
        const leftTable = tables[leftTableIndex].sheets[leftSheetIndex];
        const leftRows = leftTable.data.slice(1);
        
        const rightTableSelect = document.getElementById('rightTableSelect');
        const [rightTableIndex, rightSheetIndex] = rightTableSelect.value.split('-').map(Number);
        const rightTable = tables[rightTableIndex].sheets[rightSheetIndex];
        const rightRows = rightTable.data.slice(1);
        
        // 创建右表索引
        const rightIndex = new Map();
        rightRows.forEach(row => {
            const key = row[rightKeyCol];
            if (key !== undefined && key !== null && key !== '') {
                rightIndex.set(key, true);
            }
        });
        
        // 计算左表中有匹配的行数
        matchedRows = leftRows.filter(row => {
            const key = row[leftKeyCol];
            return key !== undefined && key !== null && key !== '' && rightIndex.has(key);
        }).length;
    } else if (joinType === 'right') {
        // RIGHT JOIN: 计算右表有匹配的行数
        const leftTableSelect = document.getElementById('leftTableSelect');
        const [leftTableIndex, leftSheetIndex] = leftTableSelect.value.split('-').map(Number);
        const leftTable = tables[leftTableIndex].sheets[leftSheetIndex];
        const leftRows = leftTable.data.slice(1);
        
        const rightTableSelect = document.getElementById('rightTableSelect');
        const [rightTableIndex, rightSheetIndex] = rightTableSelect.value.split('-').map(Number);
        const rightTable = tables[rightTableIndex].sheets[rightSheetIndex];
        const rightRows = rightTable.data.slice(1);
        
        // 创建左表索引
        const leftIndex = new Map();
        leftRows.forEach(row => {
            const key = row[leftKeyCol];
            if (key !== undefined && key !== null && key !== '') {
                leftIndex.set(key, true);
            }
        });
        
        // 计算右表中有匹配的行数
        matchedRows = rightRows.filter(row => {
            const key = row[rightKeyCol];
            return key !== undefined && key !== null && key !== '' && leftIndex.has(key);
        }).length;
    } else if (joinType === 'full') {
        // FULL OUTER JOIN: 计算所有匹配的行数
        const leftTableSelect = document.getElementById('leftTableSelect');
        const [leftTableIndex, leftSheetIndex] = leftTableSelect.value.split('-').map(Number);
        const leftTable = tables[leftTableIndex].sheets[leftSheetIndex];
        const leftRows = leftTable.data.slice(1);
        
        const rightTableSelect = document.getElementById('rightTableSelect');
        const [rightTableIndex, rightSheetIndex] = rightTableSelect.value.split('-').map(Number);
        const rightTable = tables[rightTableIndex].sheets[rightSheetIndex];
        const rightRows = rightTable.data.slice(1);
        
        // 创建索引
        const leftIndex = new Map();
        const rightIndex = new Map();
        
        leftRows.forEach(row => {
            const key = row[leftKeyCol];
            if (key !== undefined && key !== null && key !== '') {
                leftIndex.set(key, true);
            }
        });
        
        rightRows.forEach(row => {
            const key = row[rightKeyCol];
            if (key !== undefined && key !== null && key !== '') {
                rightIndex.set(key, true);
            }
        });
        
        // 计算匹配的key数量
        const matchedKeys = new Set();
        leftIndex.forEach((_, key) => {
            if (rightIndex.has(key)) {
                matchedKeys.add(key);
            }
        });
        
        matchedRows = matchedKeys.size;
    }
    
    document.getElementById('matchedRows').textContent = `匹配: ${matchedRows} 行`;
    
    // 保存预览数据
    mergePreviewData = mergedData;
}

function executeMerge() {
    if (!mergePreviewData) {
        alert('请先预览合并结果');
        return;
    }
    
    if (confirm('确定要执行合并吗？这将创建一个新的工作表。')) {
        try {
            // 创建新的合并表
            const leftTableSelect = document.getElementById('leftTableSelect');
            const rightTableSelect = document.getElementById('rightTableSelect');
            const joinTypeRadio = document.querySelector('input[name="joinType"]:checked');
            
            if (!leftTableSelect.value || !rightTableSelect.value || !joinTypeRadio) {
                alert('合并参数不完整，请重新预览');
                return;
            }
            
            const [leftTableIndex, leftSheetIndex] = leftTableSelect.value.split('-').map(Number);
            const [rightTableIndex, rightSheetIndex] = rightTableSelect.value.split('-').map(Number);
            
            // 验证表索引
            if (leftTableIndex < 0 || leftTableIndex >= tables.length ||
                rightTableIndex < 0 || rightTableIndex >= tables.length) {
                alert('选择的表不存在');
                return;
            }
            
            const leftTable = tables[leftTableIndex];
            const rightTable = tables[rightTableIndex];
            const joinType = joinTypeRadio.value;
            
            // 生成唯一的表名
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            const mergedTableName = `${leftTable.name}_${rightTable.name}_${joinType.toUpperCase()}_${timestamp}`;
            
            // 添加到表列表
            const newTable = {
                name: mergedTableName,
                fileName: `${mergedTableName}.xlsx`,
                sheets: [{
                    name: 'Merged_Data',
                    data: mergePreviewData,
                    originalData: JSON.parse(JSON.stringify(mergePreviewData))
                }]
            };
            
            tables.push(newTable);
            
            // 切换到新创建的表
            switchToTable(tables.length - 1, 0);
            
            // 隐藏预览区域
            document.getElementById('mergePreviewSection').style.display = 'none';
            
            // 清空预览数据
            mergePreviewData = null;
            
            // 更新合并选择器
            updateMergeSelectors();
            
            alert('合并完成！新表已创建并切换到合并结果。');
        } catch (error) {
            console.error('执行合并时发生错误:', error);
            alert('执行合并时发生错误: ' + error.message);
        }
    }
}

// 页面加载完成后初始化筛选列
document.addEventListener('DOMContentLoaded', function() {
    // 延迟初始化，确保DOM完全加载
    setTimeout(initializeFilterColumns, 100);
    
    // 添加表选择器事件监听
    document.getElementById('leftTableSelect').addEventListener('change', updateKeyColumnSelectors);
    document.getElementById('rightTableSelect').addEventListener('change', updateKeyColumnSelectors);
});
