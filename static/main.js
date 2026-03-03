// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let stockData = [];
let reportDate = '';
let excelUploaded = false;

// ---------------------------------------------------------------------------
// DOM refs
// ---------------------------------------------------------------------------
const $ = (sel) => document.querySelector(sel);
const screenshotInput = $('#screenshot-input');
const excelInput = $('#excel-input');
const screenshotName = $('#screenshot-name');
const excelName = $('#excel-name');
const parseBtn = $('#parse-btn');
const editBody = $('#edit-body');
const editSection = $('#edit-section');
const previewWrapper = $('#preview-wrapper');
const captureArea = $('#capture-area');
const bannerDate = $('#banner-date');
const cardsGrid = $('#cards-grid');
const captureFooter = $('#capture-footer');
const dateInput = $('#date-input');
const manualNames = $('#manual-names');

// ---------------------------------------------------------------------------
// File input display + drag & drop
// ---------------------------------------------------------------------------
screenshotInput.addEventListener('change', () => {
  screenshotName.textContent = screenshotInput.files[0]?.name || '';
});
excelInput.addEventListener('change', () => {
  excelName.textContent = excelInput.files[0]?.name || '';
});

function setupDragDrop(boxId, inputEl, nameEl) {
  const box = document.getElementById(boxId);
  if (!box) return;
  ['dragenter', 'dragover'].forEach(evt => {
    box.addEventListener(evt, (e) => { e.preventDefault(); box.classList.add('dragover'); });
  });
  ['dragleave', 'drop'].forEach(evt => {
    box.addEventListener(evt, () => box.classList.remove('dragover'));
  });
  box.addEventListener('drop', (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) {
      const dt = new DataTransfer();
      dt.items.add(file);
      inputEl.files = dt.files;
      nameEl.textContent = file.name;
    }
  });
}
setupDragDrop('screenshot-box', screenshotInput, screenshotName);
setupDragDrop('excel-box', excelInput, excelName);

// ---------------------------------------------------------------------------
// Upload & Parse
// ---------------------------------------------------------------------------
parseBtn.addEventListener('click', async () => {
  if (!excelInput.files[0]) {
    showToast('请先上传 Excel 报价表');
    return;
  }

  const hasScreenshot = !!screenshotInput.files[0];
  const manualText = manualNames.value.trim();

  if (!hasScreenshot && !manualText) {
    showToast('请上传热门询价截图或手动输入标的名称');
    return;
  }

  parseBtn.disabled = true;
  parseBtn.innerHTML = '<span class="spinner"></span> 解析中...';

  const fd = new FormData();
  if (hasScreenshot) fd.append('screenshot', screenshotInput.files[0]);
  fd.append('excel', excelInput.files[0]);
  if (manualText) fd.append('manual_names', manualText);

  try {
    const res = await fetch('/upload', { method: 'POST', body: fd });
    const json = await res.json();
    if (json.error) { showToast(json.error); return; }

    reportDate = json.date || new Date().toISOString().slice(0, 10);
    dateInput.value = reportDate;
    stockData = json.stocks || [];
    excelUploaded = true;

    if (stockData.length === 0 && !json.has_tesseract && !manualText) {
      showToast('OCR 不可用（未安装 tesseract），请手动输入标的名称');
    } else if (stockData.length === 0) {
      showToast('未匹配到标的，请检查名称或手动添加');
    } else {
      const matched = stockData.filter(s => s.matched).length;
      const ocrCount = (json.ocr_names || []).length;
      const manualCount = (json.manual_names || []).length;
      let msg = `共 ${stockData.length} 个标的，${matched} 个已匹配`;
      if (ocrCount > 0) msg += `（OCR ${ocrCount} 个）`;
      if (manualCount > 0) msg += `（手动 ${manualCount} 个）`;
      showToast(msg);
    }

    renderEditTable();
    renderPreview();
    editSection.classList.remove('hidden');
    previewWrapper.classList.remove('hidden');

    setTimeout(() => {
      editSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 100);
  } catch (e) {
    showToast('解析失败: ' + e.message);
  } finally {
    parseBtn.disabled = false;
    parseBtn.textContent = '开始解析';
  }
});

// ---------------------------------------------------------------------------
// Edit Table with Drag & Drop Reorder
// ---------------------------------------------------------------------------
let dragSrcIdx = null;

function renderEditTable() {
  editBody.innerHTML = '';
  stockData.forEach((s, i) => {
    const tr = document.createElement('tr');
    tr.draggable = true;
    tr.dataset.idx = i;

    tr.innerHTML = `
      <td class="drag-handle" title="拖拽排序">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor" opacity="0.4">
          <circle cx="4" cy="3" r="1.3"/><circle cx="10" cy="3" r="1.3"/>
          <circle cx="4" cy="7" r="1.3"/><circle cx="10" cy="7" r="1.3"/>
          <circle cx="4" cy="11" r="1.3"/><circle cx="10" cy="11" r="1.3"/>
        </svg>
      </td>
      <td>${i + 1}</td>
      <td><input type="text" value="${esc(s.name)}" data-idx="${i}" data-field="name" /></td>
      <td><input type="text" value="${esc(s.code)}" data-idx="${i}" data-field="code" /></td>
      <td><input type="number" step="0.01" value="${s.call_1m != null ? s.call_1m : ''}" data-idx="${i}" data-field="call_1m" /></td>
      <td><input type="number" step="0.01" value="${s.call_2m != null ? s.call_2m : ''}" data-idx="${i}" data-field="call_2m" /></td>
      <td><span class="status-tag ${s.matched ? 'ok' : 'miss'}">${s.matched ? '已匹配' : '未匹配'}</span></td>
      <td>
        <div class="td-actions">
          <button class="btn btn-outline btn-sm" onclick="lookupRow(${i})">查询</button>
          <button class="btn btn-danger btn-sm" onclick="removeRow(${i})">删除</button>
        </div>
      </td>
    `;

    tr.addEventListener('dragstart', onDragStart);
    tr.addEventListener('dragover', onDragOver);
    tr.addEventListener('dragenter', onDragEnter);
    tr.addEventListener('dragleave', onDragLeave);
    tr.addEventListener('drop', onDrop);
    tr.addEventListener('dragend', onDragEnd);

    editBody.appendChild(tr);
  });

  editBody.querySelectorAll('input').forEach(inp => {
    inp.addEventListener('input', (e) => {
      const idx = +e.target.dataset.idx;
      const field = e.target.dataset.field;
      let val = e.target.value.trim();
      if (field === 'call_1m' || field === 'call_2m') {
        val = val === '' ? null : parseFloat(val);
      }
      stockData[idx][field] = val;
      renderPreview();
    });
  });
}

function onDragStart(e) {
  dragSrcIdx = +this.dataset.idx;
  this.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', dragSrcIdx);
}

function onDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
}

function onDragEnter(e) {
  e.preventDefault();
  const tr = e.currentTarget;
  if (+tr.dataset.idx === dragSrcIdx) return;

  tr.classList.remove('drag-above', 'drag-below');
  const rect = tr.getBoundingClientRect();
  const mid = rect.top + rect.height / 2;
  if (e.clientY < mid) {
    tr.classList.add('drag-above');
  } else {
    tr.classList.add('drag-below');
  }
}

function onDragLeave(e) {
  e.currentTarget.classList.remove('drag-above', 'drag-below');
}

function onDrop(e) {
  e.preventDefault();
  e.stopPropagation();
  const targetIdx = +this.dataset.idx;
  this.classList.remove('drag-above', 'drag-below');

  if (dragSrcIdx === null || dragSrcIdx === targetIdx) return;

  const rect = this.getBoundingClientRect();
  const mid = rect.top + rect.height / 2;
  let insertIdx = e.clientY < mid ? targetIdx : targetIdx + 1;

  const [item] = stockData.splice(dragSrcIdx, 1);
  if (insertIdx > dragSrcIdx) insertIdx--;
  stockData.splice(insertIdx, 0, item);

  renderEditTable();
  renderPreview();
}

function onDragEnd() {
  dragSrcIdx = null;
  editBody.querySelectorAll('tr').forEach(tr => {
    tr.classList.remove('dragging', 'drag-above', 'drag-below');
  });
}

function moveRow(fromIdx, direction) {
  const toIdx = fromIdx + direction;
  if (toIdx < 0 || toIdx >= stockData.length) return;
  [stockData[fromIdx], stockData[toIdx]] = [stockData[toIdx], stockData[fromIdx]];
  renderEditTable();
  renderPreview();
}

function addRow() {
  stockData.push({ name: '', code: '', call_1m: null, call_2m: null, matched: false });
  renderEditTable();
  renderPreview();
  const rows = editBody.querySelectorAll('tr');
  if (rows.length) rows[rows.length - 1].querySelector('input')?.focus();
}

async function batchAppend() {
  const input = prompt('输入标的名称，多个用逗号、空格或换行分隔：');
  if (!input || !input.trim()) return;

  if (!excelUploaded) {
    showToast('请先上传并解析 Excel 报价表');
    return;
  }

  try {
    const res = await fetch('/batch_lookup', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ names: input }),
    });
    const json = await res.json();
    if (json.error) { showToast(json.error); return; }

    const newStocks = json.stocks || [];
    if (newStocks.length === 0) {
      showToast('未匹配到任何标的');
      return;
    }

    const existingNames = new Set(stockData.map(s => s.name));
    let added = 0;
    for (const s of newStocks) {
      if (!existingNames.has(s.name)) {
        stockData.push(s);
        existingNames.add(s.name);
        added++;
      }
    }

    renderEditTable();
    renderPreview();
    const matched = newStocks.filter(s => s.matched).length;
    showToast(`追加 ${added} 个标的（${matched} 个已匹配，${newStocks.length - added} 个重复已跳过）`);
  } catch (e) {
    showToast('批量查询失败: ' + e.message);
  }
}

function removeRow(idx) {
  stockData.splice(idx, 1);
  renderEditTable();
  renderPreview();
}

function clearAll() {
  if (stockData.length === 0) return;
  stockData = [];
  renderEditTable();
  renderPreview();
  showToast('已清空全部标的');
}

async function lookupRow(idx) {
  const name = stockData[idx].name;
  if (!name) { showToast('请先输入标的名称'); return; }

  try {
    const res = await fetch('/lookup', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name }),
    });
    const json = await res.json();
    if (json.matched) {
      stockData[idx] = { ...stockData[idx], ...json };
      renderEditTable();
      renderPreview();
      showToast(`已匹配: ${json.name} ${json.code}`);
    } else {
      showToast(`未找到匹配的标的: ${name}`);
    }
  } catch (e) {
    showToast('查询失败');
  }
}

// ---------------------------------------------------------------------------
// Date change
// ---------------------------------------------------------------------------
dateInput.addEventListener('change', () => {
  reportDate = dateInput.value;
  renderPreview();
});

// ---------------------------------------------------------------------------
// Preview Rendering
// ---------------------------------------------------------------------------
function renderPreview() {
  bannerDate.textContent = reportDate;
  captureFooter.textContent = `场外衍生品 · 香草期权报价 · ${reportDate}`;

  cardsGrid.innerHTML = '';
  stockData.forEach((s, idx) => {
    const card = document.createElement('div');
    card.className = 'stock-card';
    const fee1 = (s.call_1m != null && !isNaN(s.call_1m)) ? Number(s.call_1m).toFixed(2) + '%' : '-';
    const fee2 = (s.call_2m != null && !isNaN(s.call_2m)) ? Number(s.call_2m).toFixed(2) + '%' : '-';
    const feeClass1 = (s.call_1m != null && !isNaN(s.call_1m)) ? 'fee' : 'fee-na';
    const feeClass2 = (s.call_2m != null && !isNaN(s.call_2m)) ? 'fee' : 'fee-na';

    card.innerHTML = `
      <div class="card-header">
        <span class="card-idx">${String(idx + 1).padStart(2, '0')}</span>
        <span class="card-name">${esc(s.name)}</span>
        <span class="card-code">${esc(s.code)}</span>
      </div>
      <table class="card-table">
        <thead><tr><th>结构</th><th>周期</th><th>期权费</th></tr></thead>
        <tbody>
          <tr>
            <td class="struct">平值看涨</td>
            <td class="period">1 个月</td>
            <td class="${feeClass1}">${fee1}</td>
          </tr>
          <tr>
            <td class="struct">平值看涨</td>
            <td class="period">2 个月</td>
            <td class="${feeClass2}">${fee2}</td>
          </tr>
        </tbody>
      </table>
    `;
    cardsGrid.appendChild(card);
  });
}

// ---------------------------------------------------------------------------
// Export Image
// ---------------------------------------------------------------------------
async function exportImage() {
  if (stockData.length === 0) {
    showToast('没有数据可导出');
    return;
  }

  showToast('正在生成图片...');

  try {
    const canvas = await html2canvas(captureArea, {
      backgroundColor: '#fdf6ec',
      scale: 2,
      useCORS: true,
      logging: false,
    });

    const link = document.createElement('a');
    link.download = `热门询价标的_${reportDate}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    showToast('图片已保存');
  } catch (e) {
    showToast('图片生成失败: ' + e.message);
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function esc(str) {
  if (!str) return '';
  const d = document.createElement('div');
  d.textContent = str;
  return d.innerHTML;
}

function showToast(msg) {
  let t = document.querySelector('.toast');
  if (!t) {
    t = document.createElement('div');
    t.className = 'toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 3000);
}
