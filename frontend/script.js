const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileChip = document.getElementById('fileChip');
const fileName = document.getElementById('fileName');
const clearFile = document.getElementById('clearFile');
const analyzeBtn = document.getElementById('analyzeBtn');
const uploadForm = document.getElementById('uploadForm');
const bankSelect = document.getElementById('bankSelect');
const pdfPassword = document.getElementById('pdfPassword');
const errorLine = document.getElementById('errorLine');

const uploadCard = document.getElementById('uploadCard');
const loadingState = document.getElementById('loadingState');
const loadingText = document.getElementById('loadingText');
const results = document.getElementById('results');

const resultBank = document.getElementById('resultBank');
const resultCount = document.getElementById('resultCount');
const stamp = document.getElementById('stamp');
const monthTabs = document.getElementById('monthTabs');
const ledgerBody = document.getElementById('ledgerBody');
const downloadBtn = document.getElementById('downloadBtn');
const startOverBtn = document.getElementById('startOverBtn');

let selectedFile = null;

const LOADING_MESSAGES = [
  'Reading pages…',
  'Separating debit and credit columns…',
  'Grouping transactions by month…',
  'Tallying charges and interest…',
];

/* ---------- File selection ---------- */

function setFile(file) {
  if (!file) return;
  selectedFile = file;
  fileName.textContent = file.name;
  fileChip.hidden = false;
  analyzeBtn.disabled = false;
  hideError();
}

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

['dragenter', 'dragover'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
  });
});

dropzone.addEventListener('drop', (e) => {
  const file = e.dataTransfer.files[0];
  if (file) setFile(file);
});

clearFile.addEventListener('click', () => {
  selectedFile = null;
  fileInput.value = '';
  fileChip.hidden = true;
  analyzeBtn.disabled = true;
});

/* ---------- Submit / analyze ---------- */

uploadForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!selectedFile) return;

  hideError();
  uploadCard.hidden = true;
  results.hidden = true;
  loadingState.hidden = false;
  cycleLoadingMessages();

  const formData = new FormData();
  formData.append('statement', selectedFile);
  formData.append('bank', bankSelect.value);
  if (pdfPassword.value) {
    formData.append('password', pdfPassword.value);
  }

  try {
    const resp = await fetch('/api/analyze', { method: 'POST', body: formData });
    const data = await resp.json();

    if (!resp.ok) {
      throw new Error(data.error || 'Something went wrong while analyzing the statement.');
    }

    renderResults(data);
  } catch (err) {
    uploadCard.hidden = false;
    showError(err.message || 'Could not reach the server. Is the backend running?');
  } finally {
    loadingState.hidden = true;
  }
});

function cycleLoadingMessages() {
  let i = 0;
  loadingText.textContent = LOADING_MESSAGES[0];
  const interval = setInterval(() => {
    if (loadingState.hidden) {
      clearInterval(interval);
      return;
    }
    i = (i + 1) % LOADING_MESSAGES.length;
    loadingText.textContent = LOADING_MESSAGES[i];
  }, 1400);
}

/* ---------- Results rendering ---------- */

function money(n) {
  const val = Number(n) || 0;
  return val.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function amountCell(value) {
  const v = Number(value) || 0;
  if (v === 0) return `<span class="amt-zero">—</span>`;
  return money(v);
}

function renderResults(data) {
  resultBank.textContent = data.bank;
  resultCount.textContent = `${data.transaction_count.toLocaleString('en-IN')} transactions across ${data.summary.length - 1} months`;
  downloadBtn.href = data.download_url;

  monthTabs.innerHTML = '';
  data.summary.forEach((row, idx) => {
    const tab = document.createElement('button');
    tab.type = 'button';
    tab.className = 'month-tab' + (idx === 0 ? ' active' : '');
    tab.textContent = row.Month;
    tab.addEventListener('click', () => {
      document.querySelectorAll('.month-tab').forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');
      scrollRowIntoView(idx);
    });
    monthTabs.appendChild(tab);
  });

  ledgerBody.innerHTML = '';
  data.summary.forEach((row) => {
    const tr = document.createElement('tr');
    const isTotal = row.Month === 'TOTAL (FY)';
    if (isTotal) tr.classList.add('total-row');

    tr.innerHTML = `
      <td>${row.Month}</td>
      <td class="amt-credit">${amountCell(row['Credit'])}</td>
      <td class="amt-debit">${amountCell(row['Debit'])}</td>
      <td class="amt-credit">${amountCell(row['Credit Charge'])}</td>
      <td class="amt-debit">${amountCell(row['Debit Charge'])}</td>
      <td class="amt-credit">${amountCell(row['Credit Interest'])}</td>
      <td class="amt-debit">${amountCell(row['Debit Interest'])}</td>
      <td>${row['Transaction Count']}</td>
    `;
    ledgerBody.appendChild(tr);
  });

  results.hidden = false;
  stamp.classList.remove('pressed');
  requestAnimationFrame(() => {
    setTimeout(() => stamp.classList.add('pressed'), 120);
  });
}

function scrollRowIntoView(idx) {
  const row = ledgerBody.children[idx];
  if (row) row.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

startOverBtn.addEventListener('click', () => {
  results.hidden = true;
  uploadCard.hidden = false;
  selectedFile = null;
  fileInput.value = '';
  fileChip.hidden = true;
  analyzeBtn.disabled = true;
  hideError();
});

/* ---------- Errors ---------- */

function showError(msg) {
  errorLine.textContent = msg;
  errorLine.hidden = false;
}
function hideError() {
  errorLine.hidden = true;
  errorLine.textContent = '';
}
