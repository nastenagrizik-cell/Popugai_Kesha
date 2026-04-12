const fileInput = document.getElementById('fileInput');
const fileStatus = document.getElementById('fileStatus');
const configPanel = document.getElementById('configPanel');
const resultsPanel = document.getElementById('resultsPanel');
const rowQuestionsSel = document.getElementById('rowQuestions');
const colQuestionSel = document.getElementById('colQuestion');
const showBaseChk = document.getElementById('showBase');
const showFullScaleChk = document.getElementById('showFullScale');
const showTop2Chk = document.getElementById('showTop2');
const showBottom2Chk = document.getElementById('showBottom2');
const significanceModeSel = document.getElementById('significanceMode');
const sigLevelSel = document.getElementById('sigLevel');
const generateBtn = document.getElementById('generateBtn');
const exportBtn = document.getElementById('exportBtn');
const codingBtn = document.getElementById('codingBtn');
const codingModal = document.getElementById('codingModal');
const codingCloseBtn = document.getElementById('codingCloseBtn');
const codingSaveBtn = document.getElementById('codingSaveBtn');
const codingQuestionTitle = document.getElementById('codingQuestionTitle');
const codingTableWrap = document.getElementById('codingTableWrap');

let datasetId = null;
let columns = [];
let lastRenderedQuestions = [];
let currentCodingQuestionKey = null;
let currentCodingItems = [];

function api(path) { return path; }

fileInput.addEventListener('change', handleUpload);
generateBtn.addEventListener('click', handleGenerate);
exportBtn.addEventListener('click', handleExport);
codingBtn.addEventListener('click', openCodingForSelectedQuestion);
codingCloseBtn.addEventListener('click', closeCodingModal);
codingSaveBtn.addEventListener('click', saveCoding);

async function handleUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  try {
    fileStatus.textContent = 'Загружаем файл…';
    const fd = new FormData();
    fd.append('file', file);
    const r = await fetch(api('/api/upload'), { method: 'POST', body: fd });
    if (!r.ok) {
      const err = await r.json().catch(() => ({}));
      throw new Error(err.detail || r.statusText);
    }
    const data = await r.json();
    datasetId = data.dataset_id;
    columns = data.columns || [];
    fileStatus.innerHTML = `✅ Загружено: <strong>${data.n_rows}</strong> респондентов, <strong>${columns.length}</strong> переменных`;
    populateSelectors();
    configPanel.classList.remove('hidden');
    resultsPanel.classList.add('hidden');
    exportBtn.disabled = true;
    codingBtn.disabled = false;
  } catch (err) {
    console.error(err);
    alert('Ошибка загрузки файла: ' + err.message);
    fileStatus.textContent = '';
  }
}

function populateSelectors() {
  rowQuestionsSel.innerHTML = '';
  colQuestionSel.innerHTML = '<option value="">Нет (Только Total)</option>';
  for (const c of columns) {
    const text = c.key === c.label ? c.key : `${c.key} — ${c.label}`;
    rowQuestionsSel.add(new Option(text, c.key));
    colQuestionSel.add(new Option(text, c.key));
  }
}

function getColumnMeta(key) {
  return columns.find(c => c.key === key) || null;
}

async function handleGenerate() {
  if (!datasetId) return alert('Сначала загрузите файл.');
  const rowCols = Array.from(rowQuestionsSel.selectedOptions).map(o => o.value);
  const colCol = colQuestionSel.value ? [colQuestionSel.value] : [];
  if (!rowCols.length) return alert('Выберите хотя бы один вопрос для строк.');

  const significanceMode = colCol.length ? significanceModeSel.value : 'none';

  const body = {
    dataset_id: datasetId,
    row_cols: rowCols,
    col_cols: colCol,
    include_top2: showTop2Chk.checked,
    include_bottom2: showBottom2Chk.checked,
    show_full_scale: showFullScaleChk.checked,
    show_base: showBaseChk.checked,
    significance_mode: significanceMode,
    sig_level: Number(sigLevelSel.value),
  };

  try {
    const r = await fetch(api('/api/crosstab'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!r.ok) {
      const err = await r.json().catch(() => ({}));
      throw new Error(err.detail || r.statusText);
    }
    const data = await r.json();
    renderAllQuestions(data);
    resultsPanel.classList.remove('hidden');
    exportBtn.disabled = false;
    resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    console.error(err);
    alert('Ошибка расчета кросс-таблицы: ' + err.message);
  }
}

function renderAllQuestions(data) {
  const container = document.getElementById('resultTable');
  container.innerHTML = '';
  lastRenderedQuestions = [];
  const questions = data.questions || [];
  if (!questions.length) return;

  const allColValues = questions.find(q => (q.col_values || []).length)?.col_values || [];

  const table = document.createElement('table');
  table.className = 'results-master-table';

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  const headCells = [
    { text: 'Показатель', cls: 'sticky-col label-cell' },
    { text: 'N', cls: 'n-cell' },
    { text: 'Total (%)', cls: 'value-cell' },
  ];

  for (const item of headCells) {
    const th = document.createElement('th');
    th.textContent = item.text;
    th.className = item.cls;
    headerRow.appendChild(th);
  }

  allColValues.forEach((cv, idx) => {
    const th = document.createElement('th');
    th.textContent = `${String.fromCharCode(65 + idx)}: ${cv} (%)`;
    th.className = 'value-cell';
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  questions.forEach((q, qIndex) => {
    const titleTr = document.createElement('tr');
    titleTr.className = 'group-title-row';
    const titleTd = document.createElement('td');
    titleTd.colSpan = 3 + allColValues.length;
    titleTd.textContent = `📋 ${q.question_label}`;
    titleTr.appendChild(titleTd);
    tbody.appendChild(titleTr);

    (q.rows || []).forEach(row => {
      const tr = document.createElement('tr');
      if (row.kind === 'base') tr.classList.add('base-row');
      if (row.kind === 'top2') tr.classList.add('top2-row');
      if (row.kind === 'bottom2') tr.classList.add('bottom2-row');

      const label = document.createElement('td');
      label.textContent = row.label;
      label.className = 'sticky-col label-cell';
      tr.appendChild(label);

      const nCell = document.createElement('td');
      nCell.textContent = row.n != null ? String(row.n) : '';
      nCell.className = 'n-cell';
      tr.appendChild(nCell);

      const totalCell = document.createElement('td');
      totalCell.textContent = row.cells && row.cells['__total'] != null ? String(row.cells['__total']) : '';
      totalCell.className = 'value-cell';
      tr.appendChild(totalCell);

      for (const cv of allColValues) {
        const td = document.createElement('td');
        td.className = 'value-cell';

        const sig = (row.sig || {})[cv] || {};
        const wrap = document.createElement('div');
        wrap.style.display = 'inline-flex';
        wrap.style.alignItems = 'center';
        wrap.style.justifyContent = 'flex-end';
        wrap.style.gap = '4px';
        wrap.style.width = '100%';

        const val = document.createElement('span');
        const cellValue = row.cells ? row.cells[cv] : null;
        val.textContent = cellValue != null ? String(cellValue) : '';
        if (sig.direction === 'up') val.className = 'sig-up';
        if (sig.direction === 'down') val.className = 'sig-down';
        wrap.appendChild(val);

        if (sig.marker) {
          const mark = document.createElement('span');
          mark.textContent = sig.marker;
          mark.className = sig.direction === 'up' ? 'sig-up' : 'sig-down';
          wrap.appendChild(mark);
        }

        if (sig.letters && sig.letters.length) {
          const letters = document.createElement('span');
          letters.textContent = sig.letters.join(' ');
          letters.className = 'sig-letters';
          wrap.appendChild(letters);
        }

        td.appendChild(wrap);
        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    });

    if (qIndex < questions.length - 1) {
      const divTr = document.createElement('tr');
      divTr.className = 'group-divider-row';
      const divTd = document.createElement('td');
      divTd.colSpan = 3 + allColValues.length;
      divTr.appendChild(divTd);
      tbody.appendChild(divTr);
    }

    lastRenderedQuestions.push(q);
  });

  table.appendChild(tbody);
  container.appendChild(table);
}

async function openCodingForSelectedQuestion() {
  const selected = Array.from(rowQuestionsSel.selectedOptions).map(o => o.value);
  if (selected.length !== 1) {
    alert('Для ручной кодировки выбери ровно один вопрос в списке строк.');
    return;
  }
  const meta = getColumnMeta(selected[0]);
  if (meta && meta.can_code === false) {
    alert('Этот вопрос похож на открытый, поэтому кодировка для него скрыта.');
    return;
  }
  await openCoding(selected[0], meta ? meta.label : selected[0]);
}

async function openCoding(questionKey, questionLabel) {
  if (!datasetId) return;
  try {
    const r = await fetch(api('/api/coding-preview'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dataset_id: datasetId, question_key: questionKey }),
    });
    if (!r.ok) {
      const err = await r.json().catch(() => ({}));
      throw new Error(err.detail || r.statusText);
    }
    const data = await r.json();
    currentCodingQuestionKey = questionKey;
    currentCodingItems = data.items || [];
    codingQuestionTitle.textContent = questionLabel || data.question_label || questionKey;
    renderCodingTable(currentCodingItems);
    codingModal.classList.remove('hidden');
    codingModal.classList.add('flex');
  } catch (err) {
    console.error(err);
    alert('Не удалось открыть кодировку: ' + err.message);
  }
}

function renderCodingTable(items) {
  const wrap = document.createElement('div');
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  ['Исходный ответ', 'N', 'Код'].forEach(t => {
    const th = document.createElement('th');
    th.textContent = t;
    hr.appendChild(th);
  });
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  items.forEach((item, idx) => {
    const tr = document.createElement('tr');
    const tdRaw = document.createElement('td');
    tdRaw.textContent = item.raw_value;
    tr.appendChild(tdRaw);
    const tdN = document.createElement('td');
    tdN.textContent = String(item.n);
    tr.appendChild(tdN);
    const tdSel = document.createElement('td');
    const sel = document.createElement('select');
    sel.className = 'coding-select';
    sel.dataset.index = String(idx);
    [
      { value: '', label: 'Не учитывать' },
      { value: '1', label: '1' },
      { value: '2', label: '2' },
      { value: '3', label: '3' },
      { value: '4', label: '4' },
      { value: '5', label: '5' },
    ].forEach(o => {
      const op = document.createElement('option');
      op.value = o.value;
      op.textContent = o.label;
      sel.appendChild(op);
    });
    sel.value = item.current_code != null ? String(item.current_code) : '';
    tdSel.appendChild(sel);
    tr.appendChild(tdSel);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrap.appendChild(table);
  codingTableWrap.innerHTML = '';
  codingTableWrap.appendChild(wrap);
}

async function saveCoding() {
  if (!datasetId || !currentCodingQuestionKey) return;
  try {
    const selects = codingTableWrap.querySelectorAll('select[data-index]');
    const mapping = {};
    selects.forEach(sel => {
      const idx = Number(sel.dataset.index);
      const item = currentCodingItems[idx];
      mapping[item.raw_value] = sel.value === '' ? null : Number(sel.value);
    });

    const r = await fetch(api('/api/coding-save'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dataset_id: datasetId, question_key: currentCodingQuestionKey, mapping }),
    });
    if (!r.ok) {
      const err = await r.json().catch(() => ({}));
      throw new Error(err.detail || r.statusText);
    }
    closeCodingModal();
    alert('Кодировка сохранена. Теперь заново построй таблицу.');
  } catch (err) {
    console.error(err);
    alert('Ошибка сохранения кодировки: ' + err.message);
  }
}

function closeCodingModal() {
  codingModal.classList.add('hidden');
  codingModal.classList.remove('flex');
}

function handleExport() {
  if (!lastRenderedQuestions.length) return;
  const wb = XLSX.utils.book_new();
  const rows = [];

  lastRenderedQuestions.forEach((q, idx) => {
    const colValues = q.col_values || [];
    rows.push([q.question_label || `Question ${idx + 1}`]);
    const header = ['Показатель', 'N', 'Total (%)', ...colValues.map((v, i) => `${String.fromCharCode(65 + i)}: ${v} (%)`)];
    rows.push(header);

    (q.rows || []).forEach(row => {
      const cells = row.cells || {};
      const sig = row.sig || {};
      const arr = [row.label ?? '', row.n ?? '', cells['__total'] ?? ''];
      colValues.forEach(cv => {
        let txt = cells[cv] != null ? String(cells[cv]) : '';
        const meta = sig[cv] || {};
        if (meta.marker) txt += meta.marker;
        if (meta.letters && meta.letters.length) txt += ' ' + meta.letters.join(' ');
        arr.push(txt.trim());
      });
      rows.push(arr);
    });
    rows.push([]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const colCount = Math.max(...rows.map(r => r.length), 0);
  ws['!cols'] = Array.from({ length: colCount }, (_, i) => {
    if (i === 0) return { wch: 42 };
    if (i === 1) return { wch: 10 };
    return { wch: 14 };
  });

  XLSX.utils.book_append_sheet(wb, ws, 'CrossTab');
  XLSX.writeFile(wb, 'crosstab.xlsx');
}
