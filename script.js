/* global XLSX, Papa */

// 状态存储
const state = {
  tableA: { file: null, workbook: null, sheets: [], selectedSheet: null, header: [], rows: [] },
  tableB: { file: null, workbook: null, sheets: [], selectedSheet: null, header: [], rows: [] },
  join: { type: 'inner', keysA: [], keysB: [], nullFill: '' },
  result: { header: [], rows: [], stats: null },
  history: [],
  loading: false,
  errors: [],
  warnings: [],
  // 单表处理状态
  mode: 'dual', // 'dual' 或 'single'
  singleTable: { 
    file: null, 
    workbook: null, 
    sheets: [], 
    selectedSheets: [],
    singleJoin: {
      keysA: [],
      keysB: [],
      type: 'inner',
      nullFill: ''
    }
  },
  processedData: { header: [], rows: [], stats: null },
  originalData: { header: [], rows: [], stats: null }, // 用于恢复筛选前的数据
  // 操作历史栈
  operationHistory: [],
  historyIndex: -1,
  maxHistorySize: 50,
};

// DOM
const $ = (sel) => document.querySelector(sel);
const fileA = $('#fileA');
const fileB = $('#fileB');
const dropzoneA = $('#dropzoneA');
const dropzoneB = $('#dropzoneB');
const fileInfoA = $('#fileInfoA');
const fileInfoB = $('#fileInfoB');
const sheetA = $('#sheetA');
const sheetB = $('#sheetB');
const hasHeader = $('#hasHeader');
const keysA = $('#keysA');
const keysB = $('#keysB');
const joinType = $('#joinType');
const nullFill = $('#nullFill');
const runJoin = $('#runJoin');
const preview = $('#preview');
const stats = $('#stats');
const dataStats = $('#dataStats');
const exportCSV = $('#exportCSV');
const exportXLSX = $('#exportXLSX');
const exportJSON = $('#exportJSON');
const resetApp = $('#resetApp');
const previewA = $('#previewA');
const previewB = $('#previewB');
const saveHistory = $('#saveHistory');
const historySection = $('#historySection');
const historyList = $('#historyList');
const searchTable = $('#searchTable');
const clearSearchBtn = $('#clearSearchBtn');

// 单表处理相关DOM
const dualTableMode = $('#dualTableMode');
const singleTableMode = $('#singleTableMode');
const dualTableSection = $('#dualTableSection');
const singleTableSection = $('#singleTableSection');
const singleTableSheets = $('#singleTableSheets');
const singleTableOperations = $('#singleTableOperations');
const fileSingle = $('#fileSingle');
const dropzoneSingle = $('#dropzoneSingle');
const fileInfoSingle = $('#fileInfoSingle');
const hasHeaderSingle = $('#hasHeaderSingle');
const sheetsGrid = $('#sheetsGrid');
const mergeAllSheets = $('#mergeAllSheets');
const processSelectedSheets = $('#processSelectedSheets');
const exportSelectedSheets = $('#exportSelectedSheets');
const singleTableKeys = $('#singleTableKeys');
const singleKeysA = $('#singleKeysA');
const singleKeysB = $('#singleKeysB');
const singleJoinType = $('#singleJoinType');
const singleNullFill = $('#singleNullFill');
const dataOperations = $('#dataOperations');
const undoOperation = $('#undoOperation');
const redoOperation = $('#redoOperation');
const operationHistory = $('#operationHistory');
const clearFilterBtn = $('#clearFilterBtn');

// 操作历史管理
function saveOperationSnapshot(operationName) {
  // 如果在历史中间进行新操作，丢弃后面的历史
  if (state.historyIndex < state.operationHistory.length - 1) {
    state.operationHistory = state.operationHistory.slice(0, state.historyIndex + 1);
  }
  
  // 保存当前数据快照
  const snapshot = {
    name: operationName,
    timestamp: new Date(),
    data: {
      header: [...state.processedData.header],
      rows: state.processedData.rows.map(row => [...row]),
      stats: { ...state.processedData.stats }
    }
  };
  
  state.operationHistory.push(snapshot);
  state.historyIndex++;
  
  // 限制历史大小
  if (state.operationHistory.length > state.maxHistorySize) {
    state.operationHistory.shift();
    state.historyIndex--;
  }
  
  updateUndoRedoButtons();
}

function undoLastOperation() {
  if (state.historyIndex > 0) {
    state.historyIndex--;
    const snapshot = state.operationHistory[state.historyIndex];
    
    // 恢复数据
    state.processedData = {
      header: [...snapshot.data.header],
      rows: snapshot.data.rows.map(row => [...row]),
      stats: { ...snapshot.data.stats }
    };
    
    renderTable(state.processedData.header, state.processedData.rows);
    renderStats(state.processedData.stats);
    updateOperationSelects();
    updateUndoRedoButtons();
    
    showNotification(`已撤销操作: ${state.operationHistory[state.historyIndex + 1].name}`, 'info');
  }
}

function redoLastOperation() {
  if (state.historyIndex < state.operationHistory.length - 1) {
    state.historyIndex++;
    const snapshot = state.operationHistory[state.historyIndex];
    
    // 恢复数据
    state.processedData = {
      header: [...snapshot.data.header],
      rows: snapshot.data.rows.map(row => [...row]),
      stats: { ...snapshot.data.stats }
    };
    
    renderTable(state.processedData.header, state.processedData.rows);
    renderStats(state.processedData.stats);
    updateOperationSelects();
    updateUndoRedoButtons();
    
    showNotification(`已重做操作: ${snapshot.name}`, 'info');
  }
}

function updateUndoRedoButtons() {
  // 更新按钮状态
  undoOperation.disabled = state.historyIndex <= 0;
  redoOperation.disabled = state.historyIndex >= state.operationHistory.length - 1;
  
  // 更新历史文本
  if (state.operationHistory.length > 0) {
    const currentOp = state.operationHistory[state.historyIndex];
    operationHistory.textContent = `当前: ${currentOp.name} (${state.historyIndex + 1}/${state.operationHistory.length})`;
  } else {
    operationHistory.textContent = '';
  }
}

function clearOperationHistory() {
  state.operationHistory = [];
  state.historyIndex = -1;
  updateUndoRedoButtons();
}

// 工具函数
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsText(file);
  });
}

function inferIsCSV(file) {
  const name = file?.name?.toLowerCase() || '';
  return name.endsWith('.csv');
}

function setOptions(selectEl, options) {
  selectEl.innerHTML = '';
  for (const opt of options) {
    const o = document.createElement('option');
    o.value = o.textContent = opt;
    selectEl.appendChild(o);
  }
}

function buildChips(container, columns, active = [], onToggle) {
  container.innerHTML = '';
  for (const col of columns) {
    const chip = document.createElement('button');
    chip.className = 'chip' + (active.includes(col) ? ' active' : '');
    chip.type = 'button';
    chip.textContent = col;
    chip.addEventListener('click', () => onToggle(col));
    container.appendChild(chip);
  }
}

function normalizeRow(row, header) {
  const obj = {};
  for (let i = 0; i < header.length; i++) {
    obj[header[i]] = row[i];
  }
  return obj;
}

function denormalizeRow(obj, header) {
  return header.map((h) => (h in obj ? obj[h] : undefined));
}

function computeJoinedHeader(headerA, headerB, keysA, keysB) {
  const setKeysA = new Set(keysA);
  const setKeysB = new Set(keysB);
  const result = [...headerA];
  for (const colB of headerB) {
    if (setKeysB.has(colB) && setKeysA.has(colB)) continue;
    if (headerA.includes(colB)) {
      result.push(colB + '_B');
    } else {
      result.push(colB);
    }
  }
  return result;
}

function makeKeyGetter(keys) {
  if (!keys || keys.length === 0) return () => '__ALL__';
  return (obj) => keys.map((k) => String(obj[k])).join('\u0001');
}

function indexByKeys(rows, header, keys) {
  const getKey = makeKeyGetter(keys);
  const index = new Map();
  for (const row of rows) {
    const obj = normalizeRow(row, header);
    const k = getKey(obj);
    if (!index.has(k)) index.set(k, []);
    index.get(k).push(obj);
  }
  return index;
}

function joinRows({ headerA, headerB, rowsA, rowsB, keysA, keysB, type, nullFill }) {
  const getKeyA = makeKeyGetter(keysA);
  const getKeyB = makeKeyGetter(keysB);
  const indexB = indexByKeys(rowsB, headerB, keysB);
  const joinedHeader = computeJoinedHeader(headerA, headerB, keysA, keysB);

  const results = [];
  let matchedCount = 0;
  let onlyLeft = 0;
  let onlyRight = 0;

  const usedRightKeys = new Set();

  // 先遍历 A
  for (const rowA of rowsA) {
    const objA = normalizeRow(rowA, headerA);
    const keyA = getKeyA(objA);
    const matches = indexB.get(keyA) || [];
    if (matches.length === 0) {
      if (type === 'left' || type === 'full') {
        const merged = { ...objA };
        for (const colB of headerB) {
          if (keysB.includes(colB) && keysA.includes(colB)) continue;
          const targetCol = headerA.includes(colB) ? colB + '_B' : colB;
          merged[targetCol] = nullFill;
        }
        results.push(denormalizeRow(merged, joinedHeader));
        onlyLeft++;
      }
      continue;
    }
    for (const objB of matches) {
      const merged = { ...objA };
      for (const colB of headerB) {
        if (keysB.includes(colB) && keysA.includes(colB)) continue;
        const targetCol = headerA.includes(colB) ? colB + '_B' : colB;
        merged[targetCol] = objB[colB];
      }
      results.push(denormalizeRow(merged, joinedHeader));
      matchedCount++;
    }
    usedRightKeys.add(keyA);
  }

  // RIGHT / FULL 需要加入 B 中未匹配的
  if (type === 'right' || type === 'full') {
    const indexA = indexByKeys(rowsA, headerA, keysA);
    for (const rowB of rowsB) {
      const objB = normalizeRow(rowB, headerB);
      const keyB = getKeyB(objB);
      const matches = indexA.get(keyB) || [];
      if (matches.length === 0) {
        const merged = {};
        for (const colA of headerA) {
          merged[colA] = keysA.includes(colA) && keysB.includes(colA) ? objB[colA] : nullFill;
        }
        for (const colB of headerB) {
          if (keysB.includes(colB) && keysA.includes(colB)) continue;
          const targetCol = headerA.includes(colB) ? colB + '_B' : colB;
          merged[targetCol] = objB[colB];
        }
        results.push(denormalizeRow(merged, joinedHeader));
        onlyRight++;
      }
    }
  }

  return { header: joinedHeader, rows: results, stats: { matchedCount, onlyLeft, onlyRight, total: results.length } };
}

// 渲染
function renderTable(header, rows) {
  const thead = ['<thead><tr>', ...header.map((h) => `<th>${escapeHtml(h)}</th>`), '</tr></thead>'].join('');
  const bodyRows = rows.map((r) => '<tr>' + r.map((c) => `<td>${escapeHtml(c)}</td>`).join('') + '</tr>').join('');
  const tbody = `<tbody>${bodyRows}</tbody>`;
  preview.innerHTML = thead + tbody;
}

function escapeHtml(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderStats(st) {
  stats.innerHTML = `
    <span>总行数: <strong>${st.total}</strong></span>
    <span>匹配: <strong>${st.matchedCount}</strong></span>
    <span>仅左: <strong>${st.onlyLeft}</strong></span>
    <span>仅右: <strong>${st.onlyRight}</strong></span>
  `;
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function analyzeDataTypes(rows, header) {
  const types = {};
  for (const col of header) {
    const samples = rows.slice(0, Math.min(100, rows.length)).map(r => normalizeRow(r, header)[col]);
    let hasNumber = false, hasString = false, hasEmpty = false;
    for (const val of samples) {
      if (val === '' || val === null || val === undefined) hasEmpty = true;
      else if (!isNaN(Number(val))) hasNumber = true;
      else hasString = true;
    }
    if (hasNumber && !hasString) types[col] = '数字';
    else if (hasString) types[col] = '文本';
    else types[col] = hasEmpty ? '空' : '混合';
  }
  return types;
}

function renderDataStats() {
  const cards = [];
  
  if (state.mode === 'dual') {
    // 双表模式统计
    if (state.tableA.header.length > 0) {
      const types = analyzeDataTypes(state.tableA.rows, state.tableA.header);
      cards.push(`
        <div class="stat-card">
          <h3>左表 (A)</h3>
          <div class="stat-value">${state.tableA.rows.length}</div>
          <div class="stat-detail">${state.tableA.header.length} 列 | ${state.tableA.file ? formatBytes(state.tableA.file.size) : ''}</div>
        </div>
      `);
    }
    
    if (state.tableB.header.length > 0) {
      const types = analyzeDataTypes(state.tableB.rows, state.tableB.header);
      cards.push(`
        <div class="stat-card">
          <h3>右表 (B)</h3>
          <div class="stat-value">${state.tableB.rows.length}</div>
          <div class="stat-detail">${state.tableB.header.length} 列 | ${state.tableB.file ? formatBytes(state.tableB.file.size) : ''}</div>
        </div>
      `);
    }
    
    if (state.result.header.length > 0) {
      cards.push(`
        <div class="stat-card">
          <h3>合并结果</h3>
          <div class="stat-value">${state.result.rows.length}</div>
          <div class="stat-detail">${state.result.header.length} 列</div>
        </div>
      `);
    }
  } else if (state.mode === 'single') {
    // 单表模式统计
    if (state.singleTable.file) {
      cards.push(`
        <div class="stat-card">
          <h3>上传文件</h3>
          <div class="stat-value">${state.singleTable.sheets.length}</div>
          <div class="stat-detail">工作表数量 | ${formatBytes(state.singleTable.file.size)}</div>
        </div>
      `);
    }
    
    if (state.singleTable.selectedSheets.length > 0) {
      const totalRows = state.singleTable.sheets
        .filter(sheet => state.singleTable.selectedSheets.includes(sheet.name))
        .reduce((sum, sheet) => sum + sheet.rows.length, 0);
      
      cards.push(`
        <div class="stat-card">
          <h3>选中工作表</h3>
          <div class="stat-value">${state.singleTable.selectedSheets.length}</div>
          <div class="stat-detail">共 ${totalRows} 行数据</div>
        </div>
      `);
    }
    
    if (state.processedData.header.length > 0) {
      cards.push(`
        <div class="stat-card">
          <h3>处理结果</h3>
          <div class="stat-value">${state.processedData.rows.length}</div>
          <div class="stat-detail">${state.processedData.header.length} 列</div>
        </div>
      `);
    }
  }
  
  dataStats.innerHTML = cards.join('');
  
  // 显示验证结果
  validateData();
  showValidationResults();
}

// 显示加载状态
function showLoading(message = '处理中...') {
  state.loading = true;
  const loadingDiv = document.getElementById('loadingIndicator') || createLoadingIndicator();
  loadingDiv.innerHTML = `
    <div class="loading-content">
      <div class="loading-spinner"></div>
      <div class="loading-text">${message}</div>
    </div>
  `;
  loadingDiv.style.display = 'flex';
}

function hideLoading() {
  state.loading = false;
  const loadingDiv = document.getElementById('loadingIndicator');
  if (loadingDiv) {
    loadingDiv.style.display = 'none';
  }
}

function createLoadingIndicator() {
  const loadingDiv = document.createElement('div');
  loadingDiv.id = 'loadingIndicator';
  loadingDiv.className = 'loading-indicator';
  document.body.appendChild(loadingDiv);
  return loadingDiv;
}

// 通知系统
function showNotification(message, type = 'info', duration = 3000) {
  const notification = document.createElement('div');
  notification.className = `notification notification-${type}`;
  notification.innerHTML = `
    <div class="notification-content">
      <span class="notification-icon">${getNotificationIcon(type)}</span>
      <span class="notification-message">${message}</span>
      <button class="notification-close" onclick="this.parentElement.parentElement.remove()">×</button>
    </div>
  `;
  
  const container = document.getElementById('notificationContainer') || createNotificationContainer();
  container.appendChild(notification);
  
  // 自动移除
  setTimeout(() => {
    if (notification.parentElement) {
      notification.remove();
    }
  }, duration);
}

function getNotificationIcon(type) {
  const icons = {
    success: '✅',
    error: '❌',
    warning: '⚠️',
    info: 'ℹ️'
  };
  return icons[type] || icons.info;
}

function createNotificationContainer() {
  const container = document.createElement('div');
  container.id = 'notificationContainer';
  container.className = 'notification-container';
  document.body.appendChild(container);
  return container;
}

// 模式切换
function switchMode(mode) {
  state.mode = mode;
  
  // 更新按钮状态
  dualTableMode.classList.toggle('active', mode === 'dual');
  singleTableMode.classList.toggle('active', mode === 'single');
  
  // 显示/隐藏相应区域
  dualTableSection.style.display = mode === 'dual' ? 'block' : 'none';
  singleTableSection.style.display = mode === 'single' ? 'block' : 'none';
  singleTableSheets.style.display = mode === 'single' && state.singleTable.sheets.length > 0 ? 'block' : 'none';
  
  // 数据操作模块在任何模式下都可能显示（当有数据时）
  // 但切换模式时先隐藏，等新数据生成时再显示
  if (state.processedData.rows.length === 0) {
    dataOperations.style.display = 'none';
  }
  
  // 保存模式设置
  localStorage.setItem('excel-join-mode', mode);
}

// 键盘快捷键
function setupKeyboardShortcuts() {
  document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + Enter: 执行合并
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      if (!state.loading) {
        if (state.mode === 'dual') {
          runJoin.click();
        } else {
          mergeAllSheets.click();
        }
      }
    }
    
    // Ctrl/Cmd + R: 重置 (需要确认)
    if ((e.ctrlKey || e.metaKey) && e.key === 'r') {
      e.preventDefault();
      if (confirm('确定要重置所有数据吗？')) {
        resetApp.click();
      }
    }
    
    // Ctrl/Cmd + S: 保存到历史
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
      e.preventDefault();
      if (state.result.rows.length > 0) {
        saveHistory.click();
      }
    }
    
    // Escape: 清除搜索
    if (e.key === 'Escape') {
      if (searchTable.value) {
        searchTable.value = '';
        clearSearchBtn.click();
      }
    }
  });
}

// 模板系统
function saveAsTemplate() {
  if (state.join.keysA.length === 0 || state.join.keysB.length === 0) {
    showNotification('请先选择键列', 'warning');
    return;
  }
  
  const templateName = prompt('请输入模板名称：');
  if (!templateName) return;
  
  const template = {
    id: Date.now(),
    name: templateName,
    timestamp: new Date().toLocaleString('zh-CN'),
    joinType: state.join.type,
    keysA: state.join.keysA,
    keysB: state.join.keysB,
    nullFill: state.join.nullFill,
    hasHeader: hasHeader.checked,
  };
  
  const templates = JSON.parse(localStorage.getItem('join-templates') || '[]');
  templates.unshift(template);
  if (templates.length > 10) templates.splice(10);
  
  localStorage.setItem('join-templates', JSON.stringify(templates));
  showNotification(`模板 "${templateName}" 已保存`, 'success');
  renderTemplates();
}

function loadTemplate(template) {
  joinType.value = template.joinType;
  state.join.type = template.joinType;
  nullFill.value = template.nullFill;
  state.join.nullFill = template.nullFill;
  hasHeader.checked = template.hasHeader;
  
  // 如果当前表格有相同的列名，则自动选择
  const availableKeysA = state.tableA.header.filter(h => template.keysA.includes(h));
  const availableKeysB = state.tableB.header.filter(h => template.keysB.includes(h));
  
  if (availableKeysA.length > 0 && availableKeysB.length > 0) {
    state.join.keysA = availableKeysA;
    state.join.keysB = availableKeysB;
    renderKeyChips();
    showNotification(`已应用模板 "${template.name}"`, 'success');
  } else {
    showNotification('当前表格列名与模板不匹配', 'warning');
  }
  
  persistSettings();
}

function renderTemplates() {
  const templates = JSON.parse(localStorage.getItem('join-templates') || '[]');
  const container = document.getElementById('templatesList');
  if (!container) return;
  
  if (templates.length === 0) {
    container.innerHTML = '<p style="color: var(--text-muted); text-align: center; padding: 20px;">暂无保存的模板</p>';
    return;
  }
  
  container.innerHTML = templates.map(template => `
    <div class="template-item">
      <div class="template-info">
        <div class="template-name">${template.name}</div>
        <div class="template-meta">
          ${template.joinType.toUpperCase()} JOIN | 
          ${template.keysA.length} 键列 | 
          ${template.timestamp}
        </div>
      </div>
      <div class="template-actions">
        <button class="btn small" onclick="loadTemplate(${JSON.stringify(template).replace(/"/g, '&quot;')})">应用</button>
        <button class="btn small ghost" onclick="deleteTemplate(${template.id})">删除</button>
      </div>
    </div>
  `).join('');
}

function deleteTemplate(id) {
  const templates = JSON.parse(localStorage.getItem('join-templates') || '[]');
  const filtered = templates.filter(t => t.id !== id);
  localStorage.setItem('join-templates', JSON.stringify(filtered));
  renderTemplates();
  showNotification('模板已删除', 'info');
}

// 批量操作
function batchProcess() {
  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.multiple = true;
  fileInput.accept = '.xlsx,.xls,.csv';
  
  fileInput.onchange = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length < 2) {
      showNotification('请选择至少2个文件进行批量处理', 'warning');
      return;
    }
    
    showLoading(`正在处理 ${files.length} 个文件...`);
    
    try {
      const results = [];
      for (let i = 0; i < files.length - 1; i++) {
        for (let j = i + 1; j < files.length; j++) {
          const fileA = files[i];
          const fileB = files[j];
          
          // 解析文件
          const tableA = await parseFile(fileA);
          const tableB = await parseFile(fileB);
          
          if (tableA && tableB) {
            // 自动检测键列（选择第一个列作为键列）
            const keysA = [tableA.header[0]];
            const keysB = [tableB.header[0]];
            
            const result = joinRows({
              headerA: tableA.header,
              headerB: tableB.header,
              rowsA: tableA.rows,
              rowsB: tableB.rows,
              keysA,
              keysB,
              type: 'inner',
              nullFill: '',
            });
            
            results.push({
              fileA: fileA.name,
              fileB: fileB.name,
              result: result,
            });
          }
        }
      }
      
      // 导出批量结果
      const wb = XLSX.utils.book_new();
      results.forEach((item, index) => {
        const ws = XLSX.utils.aoa_to_sheet([item.result.header, ...item.result.rows]);
        XLSX.utils.book_append_sheet(wb, ws, `合并${index + 1}_${item.fileA}_${item.fileB}`);
      });
      
      XLSX.writeFile(wb, `批量合并结果_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showNotification(`批量处理完成！共生成 ${results.length} 个合并结果`, 'success');
      
    } catch (error) {
      console.error('批量处理失败:', error);
      showNotification('批量处理失败: ' + error.message, 'error');
    } finally {
      hideLoading();
    }
  };
  
  fileInput.click();
}

async function parseFile(file) {
  try {
    const isCSV = file.name.toLowerCase().endsWith('.csv');
    if (isCSV) {
      const text = await readFileAsText(file);
      const parsed = Papa.parse(text, { skipEmptyLines: true });
      const rows = parsed.data;
      const header = rows.shift();
      return { header, rows };
    } else {
      const ab = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(ab, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
      const header = aoa[0];
      const rows = aoa.slice(1);
      return { header, rows };
    }
  } catch (error) {
    console.error('解析文件失败:', error);
    return null;
  }
}

// 单表处理功能
async function parseSingleTable(file) {
  if (!file) return;
  
  const isCSV = inferIsCSV(file);
  if (isCSV) {
    const text = await readFileAsText(file);
    const parsed = Papa.parse(text, { skipEmptyLines: true });
    const rows = parsed.data;
    const header = hasHeaderSingle.checked ? rows.shift() : rows[0].map((_, i) => `col_${i + 1}`);
    
    state.singleTable.file = file;
    state.singleTable.sheets = [{ name: 'CSV', header, rows }];
    state.singleTable.selectedSheets = []; // 默认不勾选
  } else {
    const ab = await readFileAsArrayBuffer(file);
    const wb = XLSX.read(ab, { type: 'array' });
    const sheetNames = wb.SheetNames || [];
    
    state.singleTable.file = file;
    state.singleTable.workbook = wb;
    state.singleTable.sheets = [];
    state.singleTable.selectedSheets = [];
    
    // 解析所有工作表
    for (const sheetName of sheetNames) {
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
      if (aoa && aoa.length > 0) {
        const header = hasHeaderSingle.checked ? aoa[0] : aoa[0].map((_, i) => `col_${i + 1}`);
        const rows = hasHeaderSingle.checked ? aoa.slice(1) : aoa.slice(0);
        state.singleTable.sheets.push({ name: sheetName, header, rows });
        // 不再自动添加到 selectedSheets，让用户手动选择
      }
    }
  }
  
  renderSheetsGrid();
  showFileInfo(file, fileInfoSingle);
}

function renderSheetsGrid() {
  if (state.singleTable.sheets.length === 0) {
    sheetsGrid.innerHTML = '<p style="text-align: center; color: var(--text-muted); padding: 20px;">暂无工作表数据</p>';
    return;
  }
  
  sheetsGrid.innerHTML = state.singleTable.sheets.map(sheet => `
    <div class="sheet-card ${state.singleTable.selectedSheets.includes(sheet.name) ? 'selected' : ''}">
      <div class="sheet-header">
        <span class="sheet-name">${sheet.name}</span>
        <input type="checkbox" class="sheet-checkbox" 
               ${state.singleTable.selectedSheets.includes(sheet.name) ? 'checked' : ''}
               onchange="toggleSheet('${sheet.name}')" />
      </div>
      <div class="sheet-info">
        <span>行数: ${sheet.rows.length}</span>
        <span>列数: ${sheet.header.length}</span>
      </div>
      <div class="sheet-preview">
        <table>
          <thead>
            <tr>${sheet.header.slice(0, 5).map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>
          </thead>
          <tbody>
            ${sheet.rows.slice(0, 3).map(row => 
              `<tr>${row.slice(0, 5).map(cell => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`
            ).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `).join('');
}

function toggleSheet(sheetName) {
  const index = state.singleTable.selectedSheets.indexOf(sheetName);
  if (index > -1) {
    state.singleTable.selectedSheets.splice(index, 1);
  } else {
    state.singleTable.selectedSheets.push(sheetName);
  }
  renderSheetsGrid();
  updateOperationSelects(); // 更新操作选择框
  renderDataStats(); // 更新统计信息
  
  // 如果选择了恰好2个工作表，显示键列选择区域并渲染键列chips
  if (state.singleTable.selectedSheets.length === 2) {
    singleTableKeys.style.display = 'block';
    renderSingleTableKeyChips();
  } else {
    singleTableKeys.style.display = 'none';
  }
}

// 渲染单表模式的键列选择chips
function renderSingleTableKeyChips() {
  if (state.singleTable.selectedSheets.length !== 2) return;
  
  const sheet1 = state.singleTable.sheets.find(s => s.name === state.singleTable.selectedSheets[0]);
  const sheet2 = state.singleTable.sheets.find(s => s.name === state.singleTable.selectedSheets[1]);
  
  if (!sheet1 || !sheet2) return;
  
  // 清空现有chips
  singleKeysA.innerHTML = '';
  singleKeysB.innerHTML = '';
  
  // 渲染第一个工作表的键列chips
  for (const col of sheet1.header) {
    const chip = document.createElement('button');
    chip.className = 'chip' + (state.singleTable.singleJoin.keysA.includes(col) ? ' active' : '');
    chip.type = 'button';
    chip.textContent = col;
    chip.addEventListener('click', () => {
      toggleKey(state.singleTable.singleJoin.keysA, col);
      renderSingleTableKeyChips();
    });
    singleKeysA.appendChild(chip);
  }
  
  // 渲染第二个工作表的键列chips
  for (const col of sheet2.header) {
    const chip = document.createElement('button');
    chip.className = 'chip' + (state.singleTable.singleJoin.keysB.includes(col) ? ' active' : '');
    chip.type = 'button';
    chip.textContent = col;
    chip.addEventListener('click', () => {
      toggleKey(state.singleTable.singleJoin.keysB, col);
      renderSingleTableKeyChips();
    });
    singleKeysB.appendChild(chip);
  }
}

function mergeAllSheetsData() {
  if (state.singleTable.selectedSheets.length === 0) {
    showNotification('请选择至少一个工作表', 'warning');
    return;
  }
  
  showLoading('正在合并工作表...');
  
  try {
    const selectedSheets = state.singleTable.sheets.filter(sheet => 
      state.singleTable.selectedSheets.includes(sheet.name)
    );
    
    if (selectedSheets.length === 0) {
      showNotification('没有选中的工作表', 'warning');
      return;
    }
    
    let mergedHeader, mergedRows;
    
    // 如果选择了恰好2个工作表且设置了键列，使用JOIN合并
    if (selectedSheets.length === 2 && 
        state.singleTable.singleJoin.keysA.length > 0 && 
        state.singleTable.singleJoin.keysB.length > 0) {
      
      if (state.singleTable.singleJoin.keysA.length !== state.singleTable.singleJoin.keysB.length) {
        showNotification('键列数量不匹配，请检查选择', 'error');
        return;
      }
      
      const result = joinRows({
        headerA: selectedSheets[0].header,
        headerB: selectedSheets[1].header,
        rowsA: selectedSheets[0].rows,
        rowsB: selectedSheets[1].rows,
        keysA: state.singleTable.singleJoin.keysA,
        keysB: state.singleTable.singleJoin.keysB,
        type: state.singleTable.singleJoin.type,
        nullFill: state.singleTable.singleJoin.nullFill,
      });
      
      mergedHeader = result.header;
      mergedRows = result.rows;
      
      showNotification(`JOIN合并完成: ${result.stats.matched}条匹配，${result.stats.onlyLeft}条仅左表，${result.stats.onlyRight}条仅右表`, 'success');
      
    } else {
      // 否则使用简单的行合并
      const allHeaders = new Set();
      selectedSheets.forEach(sheet => {
        sheet.header.forEach(h => allHeaders.add(h));
      });
      mergedHeader = Array.from(allHeaders);
      
      mergedRows = [];
      selectedSheets.forEach(sheet => {
        sheet.rows.forEach(row => {
          const rowObj = {};
          sheet.header.forEach((h, i) => {
            rowObj[h] = row[i];
          });
          const mergedRow = mergedHeader.map(h => rowObj[h] || '');
          mergedRows.push(mergedRow);
        });
      });
      
      showNotification(`成功合并 ${selectedSheets.length} 个工作表，共 ${mergedRows.length} 行数据`, 'success');
    }
    
    state.processedData = {
      header: mergedHeader,
      rows: mergedRows,
      stats: {
        total: mergedRows.length,
        sheets: selectedSheets.length,
        originalRows: selectedSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0)
      }
    };
    
    // 保存原始数据副本，用于恢复筛选
    state.originalData = {
      header: [...mergedHeader],
      rows: mergedRows.map(row => [...row]),
      stats: { ...state.processedData.stats }
    };
    
    renderTable(mergedHeader, mergedRows);
    renderStats(state.processedData.stats);
    
    // 清空操作历史并保存初始快照
    clearOperationHistory();
    saveOperationSnapshot('初始数据');
    
    // 显示数据操作模块
    dataOperations.style.display = 'block';
    updateOperationSelects();
    
  } catch (error) {
    console.error('合并失败:', error);
    showNotification('合并失败: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
}

// 数据筛选功能
function applyDataFilter() {
  const column = $('#filterColumn').value;
  const condition = $('#filterCondition').value;
  const value = $('#filterValue').value;
  
  if (!column) {
    showNotification('请选择筛选列', 'warning');
    return;
  }
  
  if (state.processedData.rows.length === 0) {
    showNotification('请先合并工作表', 'warning');
    return;
  }
  
  // 如果是第一次筛选，保存当前数据到originalData
  if (state.originalData.rows.length === 0) {
    state.originalData = {
      header: [...state.processedData.header],
      rows: state.processedData.rows.map(row => [...row]),
      stats: { ...state.processedData.stats }
    };
  }
  
  const columnIndex = state.processedData.header.indexOf(column);
  if (columnIndex === -1) {
    showNotification('找不到指定的列', 'error');
    return;
  }
  
  // 始终从原始数据筛选，以支持反复筛选
  const sourceRows = state.originalData.rows.length > 0 ? state.originalData.rows : state.processedData.rows;
  
  let filteredRows;
  switch (condition) {
    case 'equals':
      filteredRows = sourceRows.filter(row => 
        String(row[columnIndex]).toLowerCase() === value.toLowerCase()
      );
      break;
    case 'contains':
      filteredRows = sourceRows.filter(row => 
        String(row[columnIndex]).toLowerCase().includes(value.toLowerCase())
      );
      break;
    case 'starts_with':
      filteredRows = sourceRows.filter(row => 
        String(row[columnIndex]).toLowerCase().startsWith(value.toLowerCase())
      );
      break;
    case 'ends_with':
      filteredRows = sourceRows.filter(row => 
        String(row[columnIndex]).toLowerCase().endsWith(value.toLowerCase())
      );
      break;
    case 'greater':
      filteredRows = sourceRows.filter(row => 
        Number(row[columnIndex]) > Number(value)
      );
      break;
    case 'less':
      filteredRows = sourceRows.filter(row => 
        Number(row[columnIndex]) < Number(value)
      );
      break;
    case 'not_empty':
      filteredRows = sourceRows.filter(row => 
        row[columnIndex] && String(row[columnIndex]).trim() !== ''
      );
      break;
    default:
      filteredRows = sourceRows;
  }
  
  state.processedData.rows = filteredRows;
  state.processedData.stats.total = filteredRows.length;
  
  renderTable(state.processedData.header, filteredRows);
  renderStats(state.processedData.stats);
  
  // 保存操作快照
  saveOperationSnapshot(`筛选: ${column} ${condition} ${value || '(非空)'}`);
  
  showNotification(`筛选完成，剩余 ${filteredRows.length} 行数据（原始数据 ${sourceRows.length} 行）`, 'success');
}

// 清除筛选，还原数据
function clearDataFilter() {
  if (state.originalData.rows.length === 0) {
    showNotification('当前没有应用筛选', 'info');
    return;
  }
  
  // 恢复原始数据
  state.processedData = {
    header: [...state.originalData.header],
    rows: state.originalData.rows.map(row => [...row]),
    stats: { ...state.originalData.stats }
  };
  
  // 清空筛选输入
  $('#filterColumn').value = '';
  $('#filterCondition').value = 'equals';
  $('#filterValue').value = '';
  
  renderTable(state.processedData.header, state.processedData.rows);
  renderStats(state.processedData.stats);
  showNotification(`已清除筛选，恢复 ${state.processedData.rows.length} 行数据`, 'success');
}

// 数据排序功能
function applyDataSort() {
  const column = $('#sortColumn').value;
  const order = $('#sortOrder').value;
  
  if (!column) {
    showNotification('请选择排序列', 'warning');
    return;
  }
  
  if (state.processedData.rows.length === 0) {
    showNotification('请先合并工作表', 'warning');
    return;
  }
  
  const columnIndex = state.processedData.header.indexOf(column);
  if (columnIndex === -1) {
    showNotification('找不到指定的列', 'error');
    return;
  }
  
  const sortedRows = [...state.processedData.rows].sort((a, b) => {
    const aVal = a[columnIndex];
    const bVal = b[columnIndex];
    
    // 尝试数字比较
    const aNum = Number(aVal);
    const bNum = Number(bVal);
    
    if (!isNaN(aNum) && !isNaN(bNum)) {
      return order === 'asc' ? aNum - bNum : bNum - aNum;
    }
    
    // 字符串比较
    const aStr = String(aVal).toLowerCase();
    const bStr = String(bVal).toLowerCase();
    return order === 'asc' ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
  });
  
  state.processedData.rows = sortedRows;
  renderTable(state.processedData.header, sortedRows);
  
  // 保存操作快照
  saveOperationSnapshot(`排序: ${column} (${order === 'asc' ? '升序' : '降序'})`);
  
  showNotification(`排序完成`, 'success');
}

// 数据清洗功能
function applyDataCleanup() {
  const operation = $('#cleanupOperation').value;
  
  if (state.processedData.rows.length === 0) {
    showNotification('请先合并工作表', 'warning');
    return;
  }
  
  let cleanedRows = [...state.processedData.rows];
  let removedCount = 0;
  
  switch (operation) {
    case 'remove_duplicates':
      const seen = new Set();
      const originalLength = cleanedRows.length;
      cleanedRows = cleanedRows.filter(row => {
        const key = row.join('|');
        if (seen.has(key)) {
          removedCount++;
          return false;
        }
        seen.add(key);
        return true;
      });
      removedCount = originalLength - cleanedRows.length;
      break;
      
    case 'remove_empty_rows':
      const originalLength2 = cleanedRows.length;
      cleanedRows = cleanedRows.filter(row => 
        !row.every(cell => !cell || String(cell).trim() === '')
      );
      removedCount = originalLength2 - cleanedRows.length;
      break;
      
    case 'trim_whitespace':
      cleanedRows = cleanedRows.map(row => 
        row.map(cell => typeof cell === 'string' ? cell.trim() : cell)
      );
      break;
      
    case 'fill_empty':
      const fillValue = $('#fillValue').value;
      cleanedRows = cleanedRows.map(row => 
        row.map(cell => !cell || String(cell).trim() === '' ? fillValue : cell)
      );
      break;
  }
  
  state.processedData.rows = cleanedRows;
  state.processedData.stats.total = cleanedRows.length;
  
  renderTable(state.processedData.header, cleanedRows);
  renderStats(state.processedData.stats);
  
  // 保存操作快照
  const operationNames = {
    'remove_duplicates': '去除重复行',
    'remove_empty_rows': '删除空行',
    'trim_whitespace': '去除首尾空格',
    'fill_empty': '填充空值'
  };
  saveOperationSnapshot(`清洗: ${operationNames[operation] || operation}`);
  
  if (removedCount > 0) {
    showNotification(`清洗完成，${operation === 'remove_duplicates' ? '去重' : '删除'}了 ${removedCount} 行数据`, 'success');
  } else {
    showNotification('清洗完成', 'success');
  }
}

// 处理选中工作表
function processSelectedSheetsData() {
  if (state.singleTable.selectedSheets.length === 0) {
    showNotification('请选择至少一个工作表', 'warning');
    return;
  }
  
  showLoading('正在处理选中工作表...');
  
  try {
    const selectedSheets = state.singleTable.sheets.filter(sheet => 
      state.singleTable.selectedSheets.includes(sheet.name)
    );
    
    if (selectedSheets.length === 0) {
      showNotification('没有选中的工作表', 'warning');
      return;
    }
    
    let mergedHeader, mergedRows;
    
    // 如果选择了恰好2个工作表且设置了键列，使用JOIN合并
    if (selectedSheets.length === 2 && 
        state.singleTable.singleJoin.keysA.length > 0 && 
        state.singleTable.singleJoin.keysB.length > 0) {
      
      if (state.singleTable.singleJoin.keysA.length !== state.singleTable.singleJoin.keysB.length) {
        showNotification('键列数量不匹配，请检查选择', 'error');
        return;
      }
      
      const result = joinRows({
        headerA: selectedSheets[0].header,
        headerB: selectedSheets[1].header,
        rowsA: selectedSheets[0].rows,
        rowsB: selectedSheets[1].rows,
        keysA: state.singleTable.singleJoin.keysA,
        keysB: state.singleTable.singleJoin.keysB,
        type: state.singleTable.singleJoin.type,
        nullFill: state.singleTable.singleJoin.nullFill,
      });
      
      mergedHeader = result.header;
      mergedRows = result.rows;
      
      showNotification(`JOIN合并完成: ${result.stats.matched}条匹配，${result.stats.onlyLeft}条仅左表，${result.stats.onlyRight}条仅右表`, 'success');
      
    } else {
      // 否则使用简单的行合并
      const allHeaders = new Set();
      selectedSheets.forEach(sheet => {
        sheet.header.forEach(h => allHeaders.add(h));
      });
      mergedHeader = Array.from(allHeaders);
      
      mergedRows = [];
      selectedSheets.forEach(sheet => {
        sheet.rows.forEach(row => {
          const rowObj = {};
          sheet.header.forEach((h, i) => {
            rowObj[h] = row[i];
          });
          const mergedRow = mergedHeader.map(h => rowObj[h] || '');
          mergedRows.push(mergedRow);
        });
      });
      
      showNotification(`成功处理 ${selectedSheets.length} 个工作表，共 ${mergedRows.length} 行数据`, 'success');
    }
    
    state.processedData = {
      header: mergedHeader,
      rows: mergedRows,
      stats: {
        total: mergedRows.length,
        sheets: selectedSheets.length,
        originalRows: selectedSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0)
      }
    };
    
    // 保存原始数据副本，用于恢复筛选
    state.originalData = {
      header: [...mergedHeader],
      rows: mergedRows.map(row => [...row]),
      stats: { ...state.processedData.stats }
    };
    
    renderTable(mergedHeader, mergedRows);
    renderStats(state.processedData.stats);
    updateOperationSelects();
    
    // 清空操作历史并保存初始快照
    clearOperationHistory();
    saveOperationSnapshot('初始数据');
    
    // 显示数据操作模块
    dataOperations.style.display = 'block';
    
  } catch (error) {
    console.error('处理失败:', error);
    showNotification('处理失败: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
}

// 导出选中工作表
function exportSelectedSheetsData() {
  if (state.singleTable.selectedSheets.length === 0) {
    showNotification('请选择至少一个工作表', 'warning');
    return;
  }
  
  try {
    const selectedSheets = state.singleTable.sheets.filter(sheet => 
      state.singleTable.selectedSheets.includes(sheet.name)
    );
    
    if (selectedSheets.length === 0) {
      showNotification('没有选中的工作表', 'warning');
      return;
    }
    
    // 创建新的工作簿
    const wb = XLSX.utils.book_new();
    
    // 为每个选中的工作表创建一个工作表
    selectedSheets.forEach(sheet => {
      const wsData = [sheet.header, ...sheet.rows];
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    
    // 如果有处理后的数据，也添加一个合并后的工作表
    if (state.processedData.rows.length > 0) {
      const mergedWsData = [state.processedData.header, ...state.processedData.rows];
      const mergedWs = XLSX.utils.aoa_to_sheet(mergedWsData);
      XLSX.utils.book_append_sheet(wb, mergedWs, '合并结果');
    }
    
    const fileName = `工作表导出_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
    showNotification(`成功导出 ${selectedSheets.length} 个工作表`, 'success');
    
  } catch (error) {
    console.error('导出失败:', error);
    showNotification('导出失败: ' + error.message, 'error');
  }
}

function persistSettings() {
  const data = {
    join: {
      type: joinType.value,
      nullFill: nullFill.value,
      keysA: state.join.keysA,
      keysB: state.join.keysB,
    },
    hasHeader: hasHeader.checked,
  };
  localStorage.setItem('excel-join-settings', JSON.stringify(data));
}

function restoreSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem('excel-join-settings'));
    if (!saved) return;
    joinType.value = saved.join?.type || 'inner';
    nullFill.value = saved.join?.nullFill || '';
    hasHeader.checked = saved.hasHeader ?? true;
    state.join.keysA = saved.join?.keysA || [];
    state.join.keysB = saved.join?.keysB || [];
  } catch {}
}

// 清理不匹配的键列
function cleanInvalidKeys() {
  // 只在双表模式下清理键列
  if (state.mode === 'dual') {
    // 清理左表键列中不存在的列名
    state.join.keysA = state.join.keysA.filter(key => state.tableA.header.includes(key));
    // 清理右表键列中不存在的列名
    state.join.keysB = state.join.keysB.filter(key => state.tableB.header.includes(key));
  }
}

// 数据验证
function validateData() {
  state.errors = [];
  state.warnings = [];
  
  // 根据模式进行不同的验证
  if (state.mode === 'dual') {
    // 双表模式验证
    if (!state.tableA.file) {
      state.errors.push('请上传左表文件');
    }
    if (!state.tableB.file) {
      state.errors.push('请上传右表文件');
    }
    
    // 检查数据完整性
    if (state.tableA.rows.length === 0) {
      state.warnings.push('左表没有数据行');
    }
    if (state.tableB.rows.length === 0) {
      state.warnings.push('右表没有数据行');
    }
    
    // 检查键列选择
    if (state.join.keysA.length === 0) {
      state.errors.push('请选择左表键列');
    }
    if (state.join.keysB.length === 0) {
      state.errors.push('请选择右表键列');
    }
    if (state.join.keysA.length !== state.join.keysB.length) {
      state.errors.push(`键列数量不匹配：左表 ${state.join.keysA.length} 个，右表 ${state.join.keysB.length} 个`);
    }
    
    // 检查数据质量
    if (state.tableA.rows.length > 0) {
      const emptyRowsA = state.tableA.rows.filter(row => row.every(cell => !cell || cell.toString().trim() === '')).length;
      if (emptyRowsA > 0) {
        state.warnings.push(`左表有 ${emptyRowsA} 行空数据`);
      }
    }
    
    if (state.tableB.rows.length > 0) {
      const emptyRowsB = state.tableB.rows.filter(row => row.every(cell => !cell || cell.toString().trim() === '')).length;
      if (emptyRowsB > 0) {
        state.warnings.push(`右表有 ${emptyRowsB} 行空数据`);
      }
    }
    
    // 检查键列数据质量
    if (state.join.keysA.length > 0 && state.tableA.rows.length > 0) {
      const keyIndexA = state.join.keysA.map(key => state.tableA.header.indexOf(key));
      const nullKeysA = state.tableA.rows.filter(row => 
        keyIndexA.some(idx => !row[idx] || row[idx].toString().trim() === '')
      ).length;
      if (nullKeysA > 0) {
        state.warnings.push(`左表键列有 ${nullKeysA} 行空值`);
      }
    }
    
    if (state.join.keysB.length > 0 && state.tableB.rows.length > 0) {
      const keyIndexB = state.join.keysB.map(key => state.tableB.header.indexOf(key));
      const nullKeysB = state.tableB.rows.filter(row => 
        keyIndexB.some(idx => !row[idx] || row[idx].toString().trim() === '')
      ).length;
      if (nullKeysB > 0) {
        state.warnings.push(`右表键列有 ${nullKeysB} 行空值`);
      }
    }
  } else if (state.mode === 'single') {
    // 单表模式验证
    if (!state.singleTable.file) {
      state.errors.push('请上传单表文件');
    }
    
    if (state.singleTable.sheets.length === 0) {
      state.warnings.push('没有可用的工作表');
    }
    
    if (state.singleTable.selectedSheets.length === 0) {
      state.warnings.push('请选择要处理的工作表');
    }
    
    // 检查单表数据质量
    if (state.processedData.rows.length > 0) {
      const emptyRows = state.processedData.rows.filter(row => 
        row.every(cell => !cell || cell.toString().trim() === '')
      ).length;
      if (emptyRows > 0) {
        state.warnings.push(`处理后的数据有 ${emptyRows} 行空数据`);
      }
    }
  }
  
  return state.errors.length === 0;
}

// 显示验证结果
function showValidationResults() {
  const container = document.getElementById('validationResults') || createValidationContainer();
  container.innerHTML = '';
  
  if (state.errors.length > 0) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'validation-errors';
    errorDiv.innerHTML = `
      <h4>❌ 错误 (${state.errors.length})</h4>
      <ul>${state.errors.map(err => `<li>${err}</li>`).join('')}</ul>
    `;
    container.appendChild(errorDiv);
  }
  
  if (state.warnings.length > 0) {
    const warningDiv = document.createElement('div');
    warningDiv.className = 'validation-warnings';
    warningDiv.innerHTML = `
      <h4>⚠️ 警告 (${state.warnings.length})</h4>
      <ul>${state.warnings.map(warn => `<li>${warn}</li>`).join('')}</ul>
    `;
    container.appendChild(warningDiv);
  }
  
  if (state.errors.length === 0 && state.warnings.length === 0) {
    const successDiv = document.createElement('div');
    successDiv.className = 'validation-success';
    successDiv.innerHTML = '<h4>✅ 数据验证通过</h4>';
    container.appendChild(successDiv);
  }
}

function createValidationContainer() {
  const container = document.createElement('div');
  container.id = 'validationResults';
  container.className = 'validation-container';
  
  // 插入到数据统计卡片中
  const dataStatsCard = document.querySelector('#dataStats').parentElement;
  dataStatsCard.appendChild(container);
  
  return container;
}

// 解析
async function parseFileToTable(file, side) {
  if (!file) return;
  const isCSV = inferIsCSV(file);
  if (isCSV) {
    const text = await readFileAsText(file);
    const parsed = Papa.parse(text, { skipEmptyLines: true });
    const rows = parsed.data;
    const header = hasHeader.checked ? rows.shift() : rows[0].map((_, i) => `col_${i + 1}`);
    const table = side === 'A' ? state.tableA : state.tableB;
    table.header = header;
    table.rows = rows;
    table.sheets = ['CSV'];
    table.selectedSheet = 'CSV';
    return;
  }

  const ab = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(ab, { type: 'array' });
  const sheetNames = wb.SheetNames || [];
  const table = side === 'A' ? state.tableA : state.tableB;
  table.workbook = wb;
  table.sheets = sheetNames;
  table.selectedSheet = sheetNames[0] || null;
  if (table.selectedSheet) extractSheet(table, table.selectedSheet);
}

function extractSheet(table, sheetName) {
  const ws = table.workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!aoa || aoa.length === 0) {
    table.header = [];
    table.rows = [];
    return;
  }
  const header = hasHeader.checked ? aoa[0] : aoa[0].map((_, i) => `col_${i + 1}`);
  const rows = hasHeader.checked ? aoa.slice(1) : aoa.slice(0);
  table.header = header;
  table.rows = rows;
}

function showFileInfo(file, infoEl) {
  if (!file) {
    infoEl.classList.remove('show');
    return;
  }
  infoEl.textContent = `📁 ${file.name} (${formatBytes(file.size)})`;
  infoEl.classList.add('show');
}

// 拖拽上传
function setupDropzone(dropzone, fileInput, side) {
  // 只在点击非input区域时触发文件选择
  dropzone.addEventListener('click', (e) => {
    if (e.target !== fileInput) {
      fileInput.click();
    }
  });
  
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
  
  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });
  
  dropzone.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      // 手动设置files到input
      const dt = new DataTransfer();
      dt.items.add(files[0]);
      fileInput.files = dt.files;
      
      const table = side === 'A' ? state.tableA : state.tableB;
      table.file = files[0];
      await parseFileToTable(files[0], side);
      const sheetSelect = side === 'A' ? sheetA : sheetB;
      const fileInfo = side === 'A' ? fileInfoA : fileInfoB;
      setOptions(sheetSelect, table.sheets);
      showFileInfo(files[0], fileInfo);
      cleanInvalidKeys(); // 清理不匹配的键列
      renderKeyChips();
      renderDataStats();
      persistSettings();
    }
  });
}

// 事件绑定
fileA.addEventListener('change', async (e) => {
  state.tableA.file = e.target.files[0] || null;
  await parseFileToTable(state.tableA.file, 'A');
  setOptions(sheetA, state.tableA.sheets);
  showFileInfo(state.tableA.file, fileInfoA);
  cleanInvalidKeys(); // 清理不匹配的键列
  renderKeyChips();
  renderDataStats();
  persistSettings();
});

fileB.addEventListener('change', async (e) => {
  state.tableB.file = e.target.files[0] || null;
  await parseFileToTable(state.tableB.file, 'B');
  setOptions(sheetB, state.tableB.sheets);
  showFileInfo(state.tableB.file, fileInfoB);
  cleanInvalidKeys(); // 清理不匹配的键列
  renderKeyChips();
  renderDataStats();
  persistSettings();
});

sheetA.addEventListener('change', () => {
  state.tableA.selectedSheet = sheetA.value;
  extractSheet(state.tableA, state.tableA.selectedSheet);
  cleanInvalidKeys(); // 清理不匹配的键列
  renderKeyChips();
  renderDataStats();
});

sheetB.addEventListener('change', () => {
  state.tableB.selectedSheet = sheetB.value;
  extractSheet(state.tableB, state.tableB.selectedSheet);
  cleanInvalidKeys(); // 清理不匹配的键列
  renderKeyChips();
  renderDataStats();
});

hasHeader.addEventListener('change', () => {
  if (state.tableA.workbook && state.tableA.selectedSheet) extractSheet(state.tableA, state.tableA.selectedSheet);
  if (state.tableB.workbook && state.tableB.selectedSheet) extractSheet(state.tableB, state.tableB.selectedSheet);
  cleanInvalidKeys(); // 清理不匹配的键列
  renderKeyChips();
  renderDataStats();
  persistSettings();
});

joinType.addEventListener('change', () => { state.join.type = joinType.value; persistSettings(); });
nullFill.addEventListener('input', () => { state.join.nullFill = nullFill.value; persistSettings(); });

function renderKeyChips() {
  // 清空现有筹片
  keysA.innerHTML = '';
  keysB.innerHTML = '';
  
  // 渲染左表键列筹片
  for (const col of state.tableA.header) {
    const chip = document.createElement('button');
    chip.className = 'chip' + (state.join.keysA.includes(col) ? ' active' : '');
    chip.type = 'button';
    chip.textContent = col;
    chip.addEventListener('click', () => {
      toggleKey(state.join.keysA, col);
      persistSettings();
      renderKeyChips(); // 重新渲染以更新状态
    });
    keysA.appendChild(chip);
  }
  
  // 渲染右表键列筹片
  for (const col of state.tableB.header) {
    const chip = document.createElement('button');
    chip.className = 'chip' + (state.join.keysB.includes(col) ? ' active' : '');
    chip.type = 'button';
    chip.textContent = col;
    chip.addEventListener('click', () => {
      toggleKey(state.join.keysB, col);
      persistSettings();
      renderKeyChips(); // 重新渲染以更新状态
    });
    keysB.appendChild(chip);
  }
}

function toggleKey(arr, col) {
  const i = arr.indexOf(col);
  if (i >= 0) arr.splice(i, 1); else arr.push(col);
}

runJoin.addEventListener('click', async () => {
  // 数据验证
  if (!validateData()) {
    showValidationResults();
    return;
  }

  try {
    showLoading('正在合并数据...');
    
    // 使用 setTimeout 让 UI 更新
    await new Promise(resolve => setTimeout(resolve, 100));
    
    const result = joinRows({
      headerA: state.tableA.header,
      headerB: state.tableB.header,
      rowsA: state.tableA.rows,
      rowsB: state.tableB.rows,
      keysA: state.join.keysA,
      keysB: state.join.keysB,
      type: state.join.type,
      nullFill: state.join.nullFill,
    });

    state.result = result;
    renderTable(result.header, result.rows);
    renderStats(result.stats);
    renderDataStats();
    
    // 显示成功消息
    showNotification('合并完成！', 'success');
    
  } catch (error) {
    console.error('合并失败:', error);
    showNotification('合并失败: ' + error.message, 'error');
  } finally {
    hideLoading();
  }
});

// 预览单表
previewA.addEventListener('click', () => {
  if (state.tableA.header.length === 0) {
    alert('请先上传左表');
    return;
  }
  const previewRows = state.tableA.rows.slice(0, 100);
  renderTable(state.tableA.header, previewRows);
  stats.textContent = `预览左表前 ${previewRows.length} 行（共 ${state.tableA.rows.length} 行）`;
});

previewB.addEventListener('click', () => {
  if (state.tableB.header.length === 0) {
    alert('请先上传右表');
    return;
  }
  const previewRows = state.tableB.rows.slice(0, 100);
  renderTable(state.tableB.header, previewRows);
  stats.textContent = `预览右表前 ${previewRows.length} 行（共 ${state.tableB.rows.length} 行）`;
});

// 搜索过滤
searchTable.addEventListener('input', () => {
  const keyword = searchTable.value.toLowerCase();
  const rows = preview.querySelectorAll('tbody tr');
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(keyword) ? '' : 'none';
  });
});

clearSearchBtn.addEventListener('click', () => {
  searchTable.value = '';
  const rows = preview.querySelectorAll('tbody tr');
  rows.forEach(row => row.style.display = '');
});

// 导出
exportCSV.addEventListener('click', () => {
  const table = $('#preview');
  if (!table || !table.querySelector('tbody tr')) return;
  const rows = [];
  table.querySelectorAll('tr').forEach((tr) => {
    const row = [];
    tr.querySelectorAll('th,td').forEach((cell) => row.push(cell.textContent));
    rows.push(row);
  });
  const csv = Papa.unparse(rows);
  downloadFile(csv, 'joined.csv', 'text/csv;charset=utf-8;');
});

exportXLSX.addEventListener('click', () => {
  const table = $('#preview');
  if (!table || !table.querySelector('tbody tr')) return;
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, ws, 'JOIN');
  XLSX.writeFile(wb, 'joined.xlsx');
});

exportJSON.addEventListener('click', () => {
  if (state.result.rows.length === 0) {
    alert('没有可导出的数据');
    return;
  }
  const data = state.result.rows.map(row => normalizeRow(row, state.result.header));
  const json = JSON.stringify(data, null, 2);
  downloadFile(json, 'joined.json', 'application/json;charset=utf-8;');
});

function downloadFile(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// 历史记录
function loadHistory() {
  try {
    const saved = localStorage.getItem('excel-join-history');
    state.history = saved ? JSON.parse(saved) : [];
    renderHistory();
  } catch {}
}

function saveToHistory() {
  if (state.result.rows.length === 0) {
    alert('请先执行合并操作');
    return;
  }
  
  const record = {
    id: Date.now(),
    timestamp: new Date().toLocaleString('zh-CN'),
    joinType: state.join.type,
    keysA: state.join.keysA,
    keysB: state.join.keysB,
    fileA: state.tableA.file?.name || '未知',
    fileB: state.tableB.file?.name || '未知',
    resultRows: state.result.rows.length,
  };
  
  state.history.unshift(record);
  if (state.history.length > 10) state.history = state.history.slice(0, 10);
  
  localStorage.setItem('excel-join-history', JSON.stringify(state.history));
  renderHistory();
  alert('已保存到历史记录');
}

function renderHistory() {
  if (state.history.length === 0) {
    historySection.style.display = 'none';
    return;
  }
  
  historySection.style.display = 'block';
  historyList.innerHTML = state.history.map(record => `
    <div class="history-item">
      <div class="history-info">
        <div class="history-title">${record.joinType.toUpperCase()} JOIN</div>
        <div class="history-meta">
          ${record.fileA} ⇔ ${record.fileB} | 
          结果: ${record.resultRows} 行 | 
          ${record.timestamp}
        </div>
      </div>
      <div class="history-actions">
        <button class="btn small ghost" onclick="deleteHistory(${record.id})">删除</button>
      </div>
    </div>
  `).join('');
}

window.deleteHistory = function(id) {
  state.history = state.history.filter(r => r.id !== id);
  localStorage.setItem('excel-join-history', JSON.stringify(state.history));
  renderHistory();
};

saveHistory.addEventListener('click', saveToHistory);

resetApp.addEventListener('click', () => {
  if (confirm('确定要重置所有数据吗？')) {
    localStorage.removeItem('excel-join-settings');
    localStorage.removeItem('excel-join-history');
    location.reload();
  }
});

// 事件绑定
dualTableMode.addEventListener('click', () => switchMode('dual'));
singleTableMode.addEventListener('click', () => switchMode('single'));

// 单表处理事件
setupDropzone(dropzoneSingle, fileSingle, 'single');
fileSingle.addEventListener('change', async (e) => {
  state.singleTable.file = e.target.files[0] || null;
  await parseSingleTable(state.singleTable.file);
  if (state.singleTable.sheets.length > 0) {
    singleTableSheets.style.display = 'block';
    singleTableOperations.style.display = 'block';
    updateOperationSelects();
  }
});

hasHeaderSingle.addEventListener('change', async () => {
  if (state.singleTable.file) {
    await parseSingleTable(state.singleTable.file);
    if (state.singleTable.sheets.length > 0) {
      updateOperationSelects();
    }
  }
});

mergeAllSheets.addEventListener('click', mergeAllSheetsData);
processSelectedSheets.addEventListener('click', processSelectedSheetsData);
exportSelectedSheets.addEventListener('click', exportSelectedSheetsData);

// 单表JOIN设置事件
singleJoinType.addEventListener('change', () => {
  state.singleTable.singleJoin.type = singleJoinType.value;
});
singleNullFill.addEventListener('input', () => {
  state.singleTable.singleJoin.nullFill = singleNullFill.value;
});

// 操作按钮事件
$('#applyFilter').addEventListener('click', applyDataFilter);
clearFilterBtn.addEventListener('click', clearDataFilter);
$('#applySort').addEventListener('click', applyDataSort);
$('#applyCleanup').addEventListener('click', applyDataCleanup);

// 撤销/重做按钮事件
undoOperation.addEventListener('click', undoLastOperation);
redoOperation.addEventListener('click', redoLastOperation);

// 清洗操作变化时显示/隐藏填充值输入
$('#cleanupOperation').addEventListener('change', (e) => {
  const fillValueInput = $('#fillValue');
  fillValueInput.style.display = e.target.value === 'fill_empty' ? 'block' : 'none';
});

// 更新操作选择框
function updateOperationSelects() {
  let header = [];
  
  // 优先使用处理后的数据
  if (state.processedData.header.length > 0) {
    header = state.processedData.header;
  } else if (state.singleTable.selectedSheets.length > 0) {
    // 如果没有处理后的数据，从选中的工作表中获取所有列名
    const allHeaders = new Set();
    state.singleTable.sheets
      .filter(sheet => state.singleTable.selectedSheets.includes(sheet.name))
      .forEach(sheet => {
        sheet.header.forEach(h => allHeaders.add(h));
      });
    header = Array.from(allHeaders);
  }
  
  if (header.length > 0) {
    const options = header.map(h => `<option value="${h}">${h}</option>`).join('');
    
    $('#filterColumn').innerHTML = '<option value="">选择列...</option>' + options;
    $('#sortColumn').innerHTML = '<option value="">选择列...</option>' + options;
    $('#pivotRows').innerHTML = '<option value="">选择行字段...</option>' + options;
    $('#pivotColumns').innerHTML = '<option value="">选择列字段...</option>' + options;
    $('#pivotValues').innerHTML = '<option value="">选择值字段...</option>' + options;
  }
}

// 全局函数
window.toggleSheet = toggleSheet;

// 初始化
setupDropzone(dropzoneA, fileA, 'A');
setupDropzone(dropzoneB, fileB, 'B');
setupKeyboardShortcuts();
restoreSettings();
loadHistory();
renderTemplates();

// 恢复模式设置
const savedMode = localStorage.getItem('excel-join-mode') || 'dual';
switchMode(savedMode);

state.join.type = joinType.value;
state.join.nullFill = nullFill.value;
renderKeyChips();
renderDataStats();

// 显示欢迎消息
setTimeout(() => {
  showNotification('欢迎使用 Excel 表格合并工具！按 Ctrl+Enter 快速执行合并', 'info', 5000);
}, 1000);
