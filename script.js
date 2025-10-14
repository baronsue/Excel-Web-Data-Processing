/* global XLSX, Papa */

// 状态存储
const state = {
  tableA: { file: null, workbook: null, sheets: [], selectedSheet: null, header: [], rows: [] },
  tableB: { file: null, workbook: null, sheets: [], selectedSheet: null, header: [], rows: [] },
  join: { type: 'inner', keysA: [], keysB: [], nullFill: '' },
};

// DOM
const $ = (sel) => document.querySelector(sel);
const fileA = $('#fileA');
const fileB = $('#fileB');
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
const exportCSV = $('#exportCSV');
const exportXLSX = $('#exportXLSX');
const resetApp = $('#resetApp');

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

function uniqueColumns(cols) {
  return Array.from(new Set(cols));
}

function computeJoinedHeader(headerA, headerB, keysA, keysB) {
  // 避免重复列名：如果两个表同名且不是 join 键，则为右表列加后缀 "_B"
  const setKeysA = new Set(keysA);
  const setKeysB = new Set(keysB);
  const result = [...headerA];
  for (const colB of headerB) {
    if (setKeysB.has(colB) && setKeysA.has(colB)) {
      // join 键列保持 A 的列名，跳过 B 的重复键
      continue;
    }
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
        // 先填 A 的列
        for (const colA of headerA) {
          merged[colA] = keysA.includes(colA) && keysB.includes(colA) ? objB[colA] : nullFill;
        }
        // 再填 B 的列
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
  stats.textContent = `合并行: ${st.total} | 匹配: ${st.matchedCount} | 仅左: ${st.onlyLeft} | 仅右: ${st.onlyRight}`;
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

// 事件绑定
fileA.addEventListener('change', async (e) => {
  state.tableA.file = e.target.files[0] || null;
  await parseFileToTable(state.tableA.file, 'A');
  setOptions(sheetA, state.tableA.sheets);
  renderKeyChips();
  persistSettings();
});

fileB.addEventListener('change', async (e) => {
  state.tableB.file = e.target.files[0] || null;
  await parseFileToTable(state.tableB.file, 'B');
  setOptions(sheetB, state.tableB.sheets);
  renderKeyChips();
  persistSettings();
});

sheetA.addEventListener('change', () => {
  state.tableA.selectedSheet = sheetA.value;
  extractSheet(state.tableA, state.tableA.selectedSheet);
  renderKeyChips();
});

sheetB.addEventListener('change', () => {
  state.tableB.selectedSheet = sheetB.value;
  extractSheet(state.tableB, state.tableB.selectedSheet);
  renderKeyChips();
});

hasHeader.addEventListener('change', () => {
  if (state.tableA.workbook && state.tableA.selectedSheet) extractSheet(state.tableA, state.tableA.selectedSheet);
  if (state.tableB.workbook && state.tableB.selectedSheet) extractSheet(state.tableB, state.tableB.selectedSheet);
  renderKeyChips();
  persistSettings();
});

joinType.addEventListener('change', () => { state.join.type = joinType.value; persistSettings(); });
nullFill.addEventListener('input', () => { state.join.nullFill = nullFill.value; persistSettings(); });

function renderKeyChips() {
  buildChips(keysA, state.tableA.header, state.join.keysA, (col) => {
    toggleKey(state.join.keysA, col);
    buildChips(keysA, state.tableA.header, state.join.keysA, () => {});
    persistSettings();
  });
  buildChips(keysB, state.tableB.header, state.join.keysB, (col) => {
    toggleKey(state.join.keysB, col);
    buildChips(keysB, state.tableB.header, state.join.keysB, () => {});
    persistSettings();
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

  renderTable(result.header, result.rows);
  renderStats(result.stats);
});

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

resetApp.addEventListener('click', () => {
  localStorage.removeItem('excel-join-settings');
  location.reload();
});

// 初始化
restoreSettings();
state.join.type = joinType.value;
state.join.nullFill = nullFill.value;
renderKeyChips();

