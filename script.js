/* global XLSX, Papa */

// 状态存储
const state = {
  tableA: { file: null, workbook: null, sheets: [], selectedSheet: null, header: [], rows: [] },
  tableB: { file: null, workbook: null, sheets: [], selectedSheet: null, header: [], rows: [] },
  join: { type: 'inner', keysA: [], keysB: [], nullFill: '' },
  result: { header: [], rows: [], stats: null },
  history: [],
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
const clearFilter = $('#clearFilter');

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
  
  dataStats.innerHTML = cards.join('');
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
  dropzone.addEventListener('click', () => fileInput.click());
  
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
      fileInput.files = files;
      const table = side === 'A' ? state.tableA : state.tableB;
      table.file = files[0];
      await parseFileToTable(files[0], side);
      const sheetSelect = side === 'A' ? sheetA : sheetB;
      const fileInfo = side === 'A' ? fileInfoA : fileInfoB;
      setOptions(sheetSelect, table.sheets);
      showFileInfo(files[0], fileInfo);
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
  renderKeyChips();
  renderDataStats();
  persistSettings();
});

fileB.addEventListener('change', async (e) => {
  state.tableB.file = e.target.files[0] || null;
  await parseFileToTable(state.tableB.file, 'B');
  setOptions(sheetB, state.tableB.sheets);
  showFileInfo(state.tableB.file, fileInfoB);
  renderKeyChips();
  renderDataStats();
  persistSettings();
});

sheetA.addEventListener('change', () => {
  state.tableA.selectedSheet = sheetA.value;
  extractSheet(state.tableA, state.tableA.selectedSheet);
  renderKeyChips();
  renderDataStats();
});

sheetB.addEventListener('change', () => {
  state.tableB.selectedSheet = sheetB.value;
  extractSheet(state.tableB, state.tableB.selectedSheet);
  renderKeyChips();
  renderDataStats();
});

hasHeader.addEventListener('change', () => {
  if (state.tableA.workbook && state.tableA.selectedSheet) extractSheet(state.tableA, state.tableA.selectedSheet);
  if (state.tableB.workbook && state.tableB.selectedSheet) extractSheet(state.tableB, state.tableB.selectedSheet);
  renderKeyChips();
  renderDataStats();
  persistSettings();
});

joinType.addEventListener('change', () => { state.join.type = joinType.value; persistSettings(); });
nullFill.addEventListener('input', () => { state.join.nullFill = nullFill.value; persistSettings(); });

function renderKeyChips() {
  buildChips(keysA, state.tableA.header, state.join.keysA, (col) => {
    toggleKey(state.join.keysA, col);
    persistSettings();
    renderKeyChips();
  });
  buildChips(keysB, state.tableB.header, state.join.keysB, (col) => {
    toggleKey(state.join.keysB, col);
    persistSettings();
    renderKeyChips();
  });
}

function toggleKey(arr, col) {
  const i = arr.indexOf(col);
  if (i >= 0) arr.splice(i, 1); else arr.push(col);
}

runJoin.addEventListener('click', () => {
  if (state.tableA.header.length === 0 || state.tableB.header.length === 0) {
    alert('请先上传左右两张表');
    return;
  }

  if (state.join.keysA.length !== state.join.keysB.length || state.join.keysA.length === 0) {
    alert('请为左右表选择相同数量的键列');
    return;
  }

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

clearFilter.addEventListener('click', () => {
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

// 初始化
setupDropzone(dropzoneA, fileA, 'A');
setupDropzone(dropzoneB, fileB, 'B');
restoreSettings();
loadHistory();
state.join.type = joinType.value;
state.join.nullFill = nullFill.value;
renderKeyChips();
renderDataStats();
