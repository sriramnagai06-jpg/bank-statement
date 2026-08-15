/* ============================================================
   Statement Ledger — Frontend Logic (Bug-fixed & Device-compatible)
   ============================================================ */

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
const warningLine = document.getElementById('warningLine');

const uploadCard = document.getElementById('uploadCard');
const loadingState = document.getElementById('loadingState');
const loadingText = document.getElementById('loadingText');
const cancelBtn = document.getElementById('cancelBtn');
const results = document.getElementById('results');

const resultBank = document.getElementById('resultBank');
const resultCount = document.getElementById('resultCount');
const stamp = document.getElementById('stamp');
const monthTabs = document.getElementById('monthTabs');
const ledgerBody = document.getElementById('ledgerBody');
const downloadBtn = document.getElementById('downloadBtn');
const startOverBtn = document.getElementById('startOverBtn');
const tableWrap = document.getElementById('tableWrap');

// New Copy-Paste Elements
const tabUploadFile = document.getElementById('tabUploadFile');
const tabPasteText = document.getElementById('tabPasteText');
const fileInputContainer = document.getElementById('fileInputContainer');
const pasteInputContainer = document.getElementById('pasteInputContainer');
const pasteInput = document.getElementById('pasteInput');
const passwordFieldContainer = document.getElementById('passwordFieldContainer');

let selectedFiles = [];
let activeMode = 'file'; // 'file' or 'paste'
let currentAbortController = null;
let loadingInterval = null;

const MAX_FILE_SIZE_MB = 50;
const WARN_FILE_SIZE_MB = 20;

const LOADING_MESSAGES = [
  'Reading pages…',
  'Separating debit and credit columns…',
  'Grouping transactions by month…',
  'Tallying charges and interest…',
];

/* ---------- Utility ---------- */

function formatBytes(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

/**
 * Check if a file is a PDF or Excel statement.
 * Many mobile browsers report an empty or wrong MIME type, so we check the file extension.
 */
function isValidExtensionOrMime(file) {
  if (!file) return false;
  var name = (file.name || '').toLowerCase();
  var type = (file.type || '').toLowerCase();

  // Extension check
  if (name.endsWith('.pdf') || name.endsWith('.xlsx') || name.endsWith('.xls')) return true;

  // MIME type check
  if (type.indexOf('pdf') !== -1 || type.indexOf('excel') !== -1 || type.indexOf('officedocument.spreadsheetml') !== -1) return true;

  return false;
}

/**
 * Validate magic header bytes for PDF (%PDF-) or Excel/Zip (PK\x03\x04).
 */
function validateMagicBytes(file) {
  return new Promise(function (resolve) {
    if (!file || file.size < 4) {
      resolve(false);
      return;
    }
    var reader = new FileReader();
    var slice = file.slice(0, 1024);
    reader.onload = function (e) {
      try {
        var buffer = new Uint8Array(e.target.result);
        var foundPDF = false;
        var foundXLSX = false;
        // PDF Magic: %PDF-
        for (var i = 0; i <= buffer.length - 5; i++) {
          if (
            buffer[i] === 0x25 &&     // %
            buffer[i + 1] === 0x50 && // P
            buffer[i + 2] === 0x44 && // D
            buffer[i + 3] === 0x46 && // F
            buffer[i + 4] === 0x2D    // -
          ) {
            foundPDF = true;
            break;
          }
        }
        // XLSX/Zip Magic: PK\x03\x04
        for (var i = 0; i <= buffer.length - 4; i++) {
          if (
            buffer[i] === 0x50 &&     // P
            buffer[i + 1] === 0x4B && // K
            buffer[i + 2] === 0x03 && // \x03
            buffer[i + 3] === 0x04    // \x04
          ) {
            foundXLSX = true;
            break;
          }
        }
        resolve(foundPDF || foundXLSX || file.name.toLowerCase().endsWith('.xls'));
      } catch (err) {
        resolve(true); // Fallback if FileReader fails
      }
    };
    reader.onerror = function () {
      resolve(true);
    };
    reader.readAsArrayBuffer(slice);
  });
}

/**
 * Auto-detect bank from file name or raw text content and auto-fill bankSelect dropdown.
 */
async function autoDetectBankForFile(file) {
  if (!file) return;

  var name = (file.name || '').toLowerCase();
  
  // 1. Check filename patterns
  var detected = detectBankFromText(name);
  
  // 2. If not detected in filename and it's a PDF, inspect first 4KB of PDF text
  if (!detected && file.size > 0 && name.endsWith('.pdf')) {
    try {
      var rawText = await readRawPDFText(file);
      detected = detectBankFromText(rawText.toLowerCase());
    } catch (e) {
      // Ignore reading errors
    }
  }

  if (detected && bankSelect) {
    bankSelect.value = detected;
    bankSelect.classList.add('auto-filled');
    setTimeout(function() {
      bankSelect.classList.remove('auto-filled');
    }, 1500);
  }
}

function detectBankFromText(text) {
  if (!text) return null;

  if (text.indexOf('kvb') !== -1 || text.indexOf('karur') !== -1) return 'kvb';
  if (text.indexOf('canara') !== -1) return 'canara';
  if (text.indexOf('sbi') !== -1 || text.indexOf('state bank') !== -1) return 'sbi';
  if (text.indexOf('hdfc') !== -1) return 'hdfc';
  if (text.indexOf('icici') !== -1) return 'icici';
  if (text.indexOf('axis') !== -1) return 'axis';
  if (text.indexOf('pnb') !== -1 || text.indexOf('punjab national') !== -1) return 'pnb';
  if (text.indexOf('bob') !== -1 || text.indexOf('bank of baroda') !== -1 || text.indexOf('baroda') !== -1) return 'bob';
  if (text.indexOf('kotak') !== -1) return 'kotak';
  if (text.indexOf('indusind') !== -1) return 'indusind';
  if (text.indexOf('union bank') !== -1 || text.indexOf('union_bank') !== -1) return 'union';
  if (text.indexOf('idfc') !== -1) return 'idfc';
  if (text.indexOf('yes bank') !== -1 || text.indexOf('yes_bank') !== -1) return 'yes';
  if (text.indexOf('bank of india') !== -1) return 'boi';
  if (text.indexOf('central bank') !== -1) return 'cbi';
  if (text.indexOf('indian overseas') !== -1) return 'iob';
  if (text.indexOf('uco') !== -1) return 'uco';
  if (text.indexOf('federal') !== -1) return 'federal';
  if (text.indexOf('south indian') !== -1) return 'southindian';
  if (text.indexOf('indian bank') !== -1) return 'indian';

  return null;
}

function readRawPDFText(file) {
  return new Promise(function(resolve) {
    var reader = new FileReader();
    var slice = file.slice(0, 4096);
    reader.onload = function(e) {
      try {
        var str = String.fromCharCode.apply(null, new Uint8Array(e.target.result));
        resolve(str);
      } catch (err) {
        resolve('');
      }
    };
    reader.onerror = function() { resolve(''); };
    reader.readAsArrayBuffer(slice);
  });
}

/* ---------- File selection & Validation ---------- */

async function setFiles(files) {
  if (!files || files.length === 0) return;

  hideError();
  hideWarning();

  if (files.length > 2) {
    showError('Maximum of 2 statement files can be analyzed together.');
    return;
  }

  var validList = [];
  var totalSize = 0;

  for (var i = 0; i < files.length; i++) {
    var file = files[i];

    if (file.size === 0) {
      showError('The selected file "' + file.name + '" is empty (0 bytes). Please select a valid bank statement.');
      return;
    }

    var isValidMagic = await validateMagicBytes(file);
    var isValidMeta = isValidExtensionOrMime(file);

    if (!isValidMagic && !isValidMeta) {
      var ext = file.name && file.name.indexOf('.') !== -1 ? file.name.split('.').pop().toUpperCase() : 'unsupported';
      showError('Selected file (' + file.name + ') is a ' + ext + ' file. Please select a valid bank statement PDF (.pdf) or Excel (.xlsx) file.');
      return;
    }

    const fileSizeMB = file.size / (1024 * 1024);
    if (fileSizeMB > MAX_FILE_SIZE_MB) {
      showError('File "' + file.name + '" is too large (' + formatBytes(file.size) + '). Maximum allowed size is ' + MAX_FILE_SIZE_MB + ' MB.');
      return;
    }

    totalSize += file.size;
    validList.push(file);
  }

  selectedFiles = validList;
  if (selectedFiles.length === 1) {
    fileName.textContent = `${selectedFiles[0].name} (${formatBytes(selectedFiles[0].size)})`;
    await autoDetectBankForFile(selectedFiles[0]);
  } else {
    var names = selectedFiles.map(function (f) { return f.name; }).join(', ');
    fileName.textContent = `${selectedFiles.length} files: ${names} (${formatBytes(totalSize)})`;
    await autoDetectBankForFile(selectedFiles[0]);
  }

  fileChip.hidden = false;
  updateAnalyzeButtonState();

  if (totalSize / (1024 * 1024) > WARN_FILE_SIZE_MB) {
    showWarning('Large files (' + formatBytes(totalSize) + '). Upload may take longer.');
  }
}

// Keyboard activation for accessibility (Label handles click natively)
dropzone.addEventListener('keydown', function (e) {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});

// File selected via input
fileInput.addEventListener('change', function () {
  if (fileInput.files && fileInput.files.length > 0) {
    setFiles(fileInput.files);
  }
});

// Drag events — prevent defaults and stop propagation to avoid page scroll
['dragenter', 'dragover'].forEach(function (evt) {
  dropzone.addEventListener(evt, function (e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach(function (evt) {
  dropzone.addEventListener(evt, function (e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove('dragover');
  });
});

dropzone.addEventListener('drop', function (e) {
  e.preventDefault();
  e.stopPropagation();
  var files = e.dataTransfer && e.dataTransfer.files;
  if (files && files.length > 0) {
    setFiles(files);
  }
});

// Prevent body-level drag from opening the file
document.body.addEventListener('dragover', function (e) {
  e.preventDefault();
});
document.body.addEventListener('drop', function (e) {
  e.preventDefault();
});

// Clear file
clearFile.addEventListener('click', function () {
  selectedFiles = [];
  fileInput.value = '';
  fileChip.hidden = true;
  updateAnalyzeButtonState();
  hideError();
  hideWarning();
});

function updateAnalyzeButtonState() {
  if (activeMode === 'file') {
    analyzeBtn.disabled = selectedFiles.length === 0;
  } else {
    analyzeBtn.disabled = !pasteInput.value.trim();
  }
}

// Tab Switching Behavior
if (tabUploadFile && tabPasteText) {
  tabUploadFile.addEventListener('click', function() {
    if (activeMode === 'file') return;
    activeMode = 'file';
    tabUploadFile.classList.add('active');
    tabPasteText.classList.remove('active');
    tabUploadFile.setAttribute('aria-selected', 'true');
    tabPasteText.setAttribute('aria-selected', 'false');
    fileInputContainer.hidden = false;
    pasteInputContainer.hidden = true;
    if (passwordFieldContainer) passwordFieldContainer.style.display = 'flex';
    if (fileChip && selectedFiles.length > 0) fileChip.hidden = false;
    updateAnalyzeButtonState();
  });

  tabPasteText.addEventListener('click', function() {
    if (activeMode === 'paste') return;
    activeMode = 'paste';
    tabPasteText.classList.add('active');
    tabUploadFile.classList.remove('active');
    tabPasteText.setAttribute('aria-selected', 'true');
    tabUploadFile.setAttribute('aria-selected', 'false');
    fileInputContainer.hidden = true;
    pasteInputContainer.hidden = false;
    if (passwordFieldContainer) passwordFieldContainer.style.display = 'none';
    if (fileChip) fileChip.hidden = true;
    updateAnalyzeButtonState();
  });
}

if (pasteInput) {
  pasteInput.addEventListener('input', function() {
    updateAnalyzeButtonState();
    updatePasteCounter();
  });
  pasteInput.addEventListener('keyup', function() {
    updateAnalyzeButtonState();
    updatePasteCounter();
  });
  pasteInput.addEventListener('change', function() {
    updateAnalyzeButtonState();
    updatePasteCounter();
  });

  // --- Native paste handler (NO preventDefault) ---
  // Let the browser handle the actual text insertion (it handles large data
  // much better than manual clipboard reading). We just update state after.
  pasteInput.addEventListener('paste', function (e) {
    // Schedule multiple checks to catch async clipboard writes
    setTimeout(function() { updateAnalyzeButtonState(); updatePasteCounter(); }, 50);
    setTimeout(function() { updateAnalyzeButtonState(); updatePasteCounter(); }, 300);
    setTimeout(function() { updateAnalyzeButtonState(); updatePasteCounter(); }, 800);
  });

  // Drag & drop text
  pasteInput.addEventListener('drop', function () {
    setTimeout(function() { updateAnalyzeButtonState(); updatePasteCounter(); }, 150);
    setTimeout(function() { updateAnalyzeButtonState(); updatePasteCounter(); }, 500);
  });

  // --- Polling fallback for Windows Clipboard History (Win+V) ---
  var _pasteLastVal = '';
  var _pastePoller = null;

  pasteInput.addEventListener('focus', function () {
    _pasteLastVal = pasteInput.value;
    _pastePoller = setInterval(function () {
      if (pasteInput.value !== _pasteLastVal) {
        _pasteLastVal = pasteInput.value;
        updateAnalyzeButtonState();
        updatePasteCounter();
      }
    }, 250);
  });

  pasteInput.addEventListener('blur', function () {
    if (_pastePoller) { clearInterval(_pastePoller); _pastePoller = null; }
    updateAnalyzeButtonState();
    updatePasteCounter();
  });
}

/** Convert an HTML table string (from Excel clipboard) to tab-separated text */
function htmlTableToTSV(html) {
  try {
    var parser = new DOMParser();
    var doc = parser.parseFromString(html, 'text/html');
    var rows = doc.querySelectorAll('tr');
    var lines = [];
    rows.forEach(function(row) {
      var cells = row.querySelectorAll('td, th');
      var parts = [];
      cells.forEach(function(cell) {
        parts.push((cell.textContent || '').trim());
      });
      if (parts.some(function(p) { return p.length > 0; })) {
        lines.push(parts.join('\t'));
      }
    });
    return lines.join('\n');
  } catch (err) {
    return '';
  }
}

/** Show a live row/char counter below the paste textarea */
function updatePasteCounter() {
  if (!pasteInput) return;
  var counterEl = document.getElementById('pasteCounter');
  if (!counterEl) return;
  var val = pasteInput.value;
  if (!val || !val.trim()) {
    counterEl.textContent = '';
    return;
  }
  var lines = val.split('\n').filter(function(l) { return l.trim().length > 0; });
  var chars = val.length;
  counterEl.textContent = '✓ ' + lines.length.toLocaleString('en-IN') + ' lines captured (' + chars.toLocaleString('en-IN') + ' chars)';
}


/* ---------- Submit / analyze ---------- */

uploadForm.addEventListener('submit', async function (e) {
  e.preventDefault();
  if (activeMode === 'file') {
    if (selectedFiles.length === 0) return;
  } else {
    if (!pasteInput.value.trim()) return;
  }

  hideError();
  hideWarning();
  uploadCard.hidden = true;
  results.hidden = true;
  loadingState.hidden = false;
  cycleLoadingMessages();

  // Create an AbortController for cancellable upload
  currentAbortController = new AbortController();

  var formData = new FormData();
  formData.append('bank', bankSelect.value);
  
  if (activeMode === 'file') {
    selectedFiles.forEach(function (file) {
      formData.append('statement', file);
    });
    if (pdfPassword.value) {
      formData.append('password', pdfPassword.value);
    }
  } else {
    formData.append('text_content', pasteInput.value.trim());
  }

  try {
    var resp = await fetch('/api/analyze', {
      method: 'POST',
      body: formData,
      signal: currentAbortController.signal,
    });

    var textBody = await resp.text();
    var data;
    try {
      data = JSON.parse(textBody);
    } catch (parseErr) {
      console.error('Non-JSON response:', textBody);
      var cleanMsg = textBody ? textBody.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim() : '';
      if (cleanMsg.length > 120) cleanMsg = cleanMsg.substring(0, 120) + '...';
      throw new Error(cleanMsg || 'Server returned an invalid response. Please try again.');
    }

    if (!resp.ok) {
      throw new Error((data && data.error) || 'Something went wrong while analyzing the statement.');
    }

    renderResults(data);
  } catch (err) {
    uploadCard.hidden = false;

    if (err.name === 'AbortError') {
      // User cancelled — do nothing
      return;
    }

    // Specific network error messages
    if (!navigator.onLine) {
      showError('You are offline. Please check your internet connection and try again.');
    } else if (err.message === 'Failed to fetch' || err.name === 'TypeError') {
      showError('Could not reach the server. Please check if the backend is running and try again.');
    } else {
      showError(err.message || 'Could not reach the server. Is the backend running?');
    }
  } finally {
    loadingState.hidden = true;
    stopLoadingMessages();
    currentAbortController = null;
  }
});

// Cancel button
cancelBtn.addEventListener('click', function () {
  if (currentAbortController) {
    currentAbortController.abort();
  }
  loadingState.hidden = true;
  uploadCard.hidden = false;
  stopLoadingMessages();
});

function cycleLoadingMessages() {
  var i = 0;
  loadingText.textContent = LOADING_MESSAGES[0];
  stopLoadingMessages(); // Clear any existing interval
  loadingInterval = setInterval(function () {
    if (loadingState.hidden) {
      stopLoadingMessages();
      return;
    }
    i = (i + 1) % LOADING_MESSAGES.length;
    loadingText.textContent = LOADING_MESSAGES[i];
  }, 1400);
}

function stopLoadingMessages() {
  if (loadingInterval) {
    clearInterval(loadingInterval);
    loadingInterval = null;
  }
}

/* ---------- Results rendering ---------- */

function money(n) {
  var val = Number(n) || 0;
  return val.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function amountCell(value) {
  var v = Number(value) || 0;
  if (v === 0) return '<span class="amt-zero">\u2014</span>';
  return money(v);
}

function renderResults(data) {
  resultBank.textContent = data.bank || 'Unknown Bank';
  
  var monthCount = (data.summary && data.summary.length > 1) ? data.summary.length - 1 : 0;
  var txnCount = data.transaction_count || 0;
  resultCount.textContent = txnCount.toLocaleString('en-IN') + ' transactions across ' + monthCount + ' months';
  
  // Set download URL — handle both relative and absolute paths
  if (data.download_url) {
    downloadBtn.href = data.download_url;
    downloadBtn.style.display = '';
  } else {
    downloadBtn.style.display = 'none';
  }

  // Populate Executive Summary Metrics Cards
  var allTxns = data.transactions || [];
  var totalDebit = 0;
  var totalCredit = 0;
  allTxns.forEach(function(t) {
    totalDebit += (Number(t.debit) || 0);
    totalCredit += (Number(t.credit) || 0);
  });

  var metricOpeningBal = document.getElementById('metricOpeningBal');
  var metricTotalDebit = document.getElementById('metricTotalDebit');
  var metricTotalCredit = document.getElementById('metricTotalCredit');
  var metricClosingBal = document.getElementById('metricClosingBal');
  var metricTxnCount = document.getElementById('metricTxnCount');
  var metricReconStatus = document.getElementById('metricReconStatus');

  if (metricOpeningBal) metricOpeningBal.innerHTML = data.opening_balance != null ? money(data.opening_balance) : '<span class="amt-zero">—</span>';
  if (metricTotalDebit) metricTotalDebit.innerHTML = money(totalDebit);
  if (metricTotalCredit) metricTotalCredit.innerHTML = money(totalCredit);
  if (metricClosingBal) metricClosingBal.innerHTML = data.closing_balance != null ? money(data.closing_balance) : '<span class="amt-zero">—</span>';
  if (metricTxnCount) metricTxnCount.textContent = txnCount.toLocaleString('en-IN');
  if (metricReconStatus) {
    var isPass = (data.reconciliation_status || '').toLowerCase() === 'pass';
    metricReconStatus.innerHTML = isPass ? '<span style="color:#166534; font-weight:700;">PASS ✓</span>' : '<span style="color:#991b1b; font-weight:700;">REVIEW ⚠️</span>';
  }

  // Build month tabs with proper ARIA
  monthTabs.innerHTML = '';
  if (data.summary && data.summary.length > 0) {
    data.summary.forEach(function (row, idx) {
      var tab = document.createElement('button');
      tab.type = 'button';
      tab.className = 'month-tab' + (idx === 0 ? ' active' : '');
      tab.textContent = row.Month;
      tab.setAttribute('role', 'tab');
      tab.setAttribute('aria-selected', idx === 0 ? 'true' : 'false');
      tab.setAttribute('id', 'tab-' + idx);
      tab.addEventListener('click', function () {
        var allTabs = document.querySelectorAll('.month-tab');
        for (var t = 0; t < allTabs.length; t++) {
          allTabs[t].classList.remove('active');
          allTabs[t].setAttribute('aria-selected', 'false');
        }
        tab.classList.add('active');
        tab.setAttribute('aria-selected', 'true');
        scrollRowIntoView(idx);
      });
      monthTabs.appendChild(tab);
    });
  }

  // Build table body
  ledgerBody.innerHTML = '';
  if (data.summary && data.summary.length > 0) {
    data.summary.forEach(function (row) {
      var tr = document.createElement('tr');
      var isTotal = row.Month === 'TOTAL (FY)';
      if (isTotal) tr.classList.add('total-row');

      tr.innerHTML =
        '<td>' + (row.Month || '') + '</td>' +
        '<td class="amt-credit">' + amountCell(row['Credit']) + '</td>' +
        '<td class="amt-debit">' + amountCell(row['Debit']) + '</td>' +
        '<td class="amt-credit">' + amountCell(row['Credit Charge']) + '</td>' +
        '<td class="amt-debit">' + amountCell(row['Debit Charge']) + '</td>' +
        '<td class="amt-credit">' + amountCell(row['Credit Interest']) + '</td>' +
        '<td class="amt-debit">' + amountCell(row['Debit Interest']) + '</td>' +
        '<td class="amt-credit">' + amountCell(row['Inter-Company Transactions']) + '</td>' +
        '<td>' + (row['Transaction Count'] || 0) + '</td>';
      ledgerBody.appendChild(tr);
    });
  }

  // Render Reconciliation & Matching Details
  var reportsSection = document.getElementById('reportsSection');
  var reconBadgeContainer = document.getElementById('reconBadgeContainer');
  var reconBadge = document.getElementById('reconBadge');
  var warningBox = document.getElementById('warningBox');
  var warningList = document.getElementById('warningList');
  var unreconciledBox = document.getElementById('unreconciledBox');
  var unreconciledList = document.getElementById('unreconciledList');
  var matchedPairsBox = document.getElementById('matchedPairsBox');
  var matchedPairsList = document.getElementById('matchedPairsList');

  // Reconciliation status badge
  if ((data.reconciliation_status || '').toLowerCase() === 'pass') {
    reconBadge.textContent = 'PASS';
    reconBadgeContainer.style.backgroundColor = '#f0fdf4';
    reconBadgeContainer.style.color = '#166534';
    reconBadgeContainer.style.border = '1px solid #86efac';
  } else {
    reconBadge.textContent = 'REVIEW REQUIRED';
    reconBadgeContainer.style.backgroundColor = '#fef2f2';
    reconBadgeContainer.style.color = '#991b1b';
    reconBadgeContainer.style.border = '1px solid #fca5a5';
  }

  // Warnings
  if (data.verification_warnings && data.verification_warnings.length > 0) {
    warningList.innerHTML = '';
    data.verification_warnings.forEach(function (warn) {
      var li = document.createElement('li');
      li.textContent = warn;
      warningList.appendChild(li);
    });
    warningBox.hidden = false;
  } else {
    warningBox.hidden = true;
  }

  // Unreconciled logs
  if (data.unreconciled_transactions && data.unreconciled_transactions.length > 0) {
    unreconciledList.innerHTML = '';
    data.unreconciled_transactions.forEach(function (log) {
      var li = document.createElement('li');
      li.textContent = log;
      unreconciledList.appendChild(li);
    });
    unreconciledBox.hidden = false;
  } else {
    unreconciledBox.hidden = true;
  }

  // Inter-company matched pairs
  if (data.inter_company_matches && data.inter_company_matches.length > 0) {
    matchedPairsList.innerHTML = '';
    data.inter_company_matches.forEach(function (pair) {
      var li = document.createElement('li');
      li.textContent = pair;
      matchedPairsList.appendChild(li);
    });
    matchedPairsBox.hidden = false;
  } else {
    matchedPairsBox.hidden = true;
  }

  if (reportsSection) {
    reportsSection.hidden = false;
  }

  // Render full transactions table
  renderTransactionsTable(data);

  // Render separate inter-company section
  renderInterCompanyTable(data);

  // Show results
  results.hidden = false;
  
  // Animate stamp
  stamp.classList.remove('pressed');
  requestAnimationFrame(function () {
    setTimeout(function () {
      stamp.classList.add('pressed');
    }, 120);
  });

  // Haptic feedback on supported devices
  try {
    if (navigator.vibrate) {
      navigator.vibrate(50);
    }
  } catch (e) {
    // Ignore — vibrate not available
  }

  // Check table scroll and update fade indicator
  updateScrollIndicator();
}

/* ---------- All Transactions Table ---------- */

var _allTxnsData = [];

function renderTransactionsTable(data) {
  var section = document.getElementById('txnDetailSection');
  var tbody   = document.getElementById('txnBody');
  var tfoot   = document.getElementById('txnFoot');
  var countEl = document.getElementById('txnDetailCount');
  var searchInput = document.getElementById('txnSearchInput');
  if (!section || !tbody || !tfoot) return;

  _allTxnsData = data.transactions || [];
  if (_allTxnsData.length === 0) {
    section.hidden = true;
    return;
  }

  function redrawTable(filterQuery) {
    tbody.innerHTML = '';
    tfoot.innerHTML = '';

    var q = (filterQuery || '').toLowerCase().trim();
    var filtered = q ? _allTxnsData.filter(function(t) {
      return (t.narration || '').toLowerCase().indexOf(q) !== -1 ||
             (t.date || '').toLowerCase().indexOf(q) !== -1 ||
             (t.type || '').toLowerCase().indexOf(q) !== -1;
    }) : _allTxnsData;

    var totalDebit  = 0;
    var totalCredit = 0;

    filtered.forEach(function (t, idx) {
      var type = t.type || 'Other';
      var isCredit = type.startsWith('Credit');
      var isDebit  = type.startsWith('Debit');
      var tr = document.createElement('tr');
      tr.innerHTML =
        '<td style="text-align:center; font-family:var(--font-mono); color:var(--slate);">' + (t.s_no || (idx + 1)) + '</td>' +
        '<td class="txn-date">'    + (t.date || '') + '</td>' +
        '<td class="txn-narr">'    + (t.narration || '').replace(/</g,'&lt;') + '</td>' +
        '<td class="txn-type ' + (isCredit ? 'type-credit' : isDebit ? 'type-debit' : '') + '">' + type + '</td>' +
        '<td class="amt-debit">'   + ((t.debit  || 0) > 0 ? money(t.debit)  : '<span class="amt-zero">—</span>') + '</td>' +
        '<td class="amt-credit">'  + ((t.credit || 0) > 0 ? money(t.credit) : '<span class="amt-zero">—</span>') + '</td>' +
        '<td class="amt-bal">'     + (t.balance != null ? money(t.balance) : '<span class="amt-zero">—</span>') + '</td>';
      tbody.appendChild(tr);
      totalDebit  += (Number(t.debit)  || 0);
      totalCredit += (Number(t.credit) || 0);
    });

    // TOTAL row in tfoot
    var tftr = document.createElement('tr');
    tftr.className = 'total-row';
    tftr.innerHTML =
      '<td colspan="4"><strong>TOTAL</strong></td>' +
      '<td class="amt-debit"><strong>'  + money(totalDebit)  + '</strong></td>' +
      '<td class="amt-credit"><strong>' + money(totalCredit) + '</strong></td>' +
      '<td></td>';
    tfoot.appendChild(tftr);

    if (countEl) {
      countEl.textContent = filtered.length.toLocaleString('en-IN') + (q ? ' of ' + _allTxnsData.length.toLocaleString('en-IN') : '') + ' txns';
    }
  }

  if (searchInput) {
    searchInput.value = '';
    searchInput.oninput = function() {
      redrawTable(searchInput.value);
    };
  }

  redrawTable('');
  section.hidden = false;
}

/* ---------- Inter-Company Transactions Table ---------- */

function renderInterCompanyTable(data) {
  var section = document.getElementById('interCompanySection');
  var tbody   = document.getElementById('icBody');
  var tfoot   = document.getElementById('icFoot');
  var countBadge = document.getElementById('icCountBadge');
  var debitBadge = document.getElementById('icDebitBadge');
  var creditBadge = document.getElementById('icCreditBadge');
  if (!section || !tbody || !tfoot) return;

  tbody.innerHTML = '';
  tfoot.innerHTML = '';

  var icTxns = data.inter_company_transactions || [];
  var icSummary = data.inter_company_summary || { count: 0, total_debit: 0, total_credit: 0 };

  if (countBadge) countBadge.textContent = icTxns.length + ' txns';
  if (debitBadge) debitBadge.textContent = 'Dr: ' + money(icSummary.total_debit || 0);
  if (creditBadge) creditBadge.textContent = 'Cr: ' + money(icSummary.total_credit || 0);

  if (icTxns.length === 0) {
    var emptyTr = document.createElement('tr');
    emptyTr.innerHTML = '<td colspan="8" style="text-align:center; padding:1.5rem; color:#64748b; font-style:italic;">No inter-company transfer transactions detected in this statement.</td>';
    tbody.appendChild(emptyTr);
    section.hidden = false;
    return;
  }

  var totalDebit = 0;
  var totalCredit = 0;

  icTxns.forEach(function (t, idx) {
    var type = t.type || 'Other';
    var isCredit = type.startsWith('Credit');
    var isDebit  = type.startsWith('Debit');
    var tr = document.createElement('tr');
    tr.innerHTML =
      '<td style="text-align:center; font-family:var(--font-mono); color:var(--slate);">' + (idx + 1) + '</td>' +
      '<td class="txn-date">'    + (t.date || '') + '</td>' +
      '<td class="txn-narr">'    + (t.narration || '').replace(/</g,'&lt;') + '</td>' +
      '<td class="txn-type ' + (isCredit ? 'type-credit' : isDebit ? 'type-debit' : '') + '">' + type + '</td>' +
      '<td class="amt-debit">'   + ((t.debit  || 0) > 0 ? money(t.debit)  : '<span class="amt-zero">—</span>') + '</td>' +
      '<td class="amt-credit">'  + ((t.credit || 0) > 0 ? money(t.credit) : '<span class="amt-zero">—</span>') + '</td>' +
      '<td class="amt-bal">'     + (t.balance != null ? money(t.balance) : '<span class="amt-zero">—</span>') + '</td>' +
      '<td><span class="txn-type" style="background:#f0fdf4; color:#166534; font-weight:600;">Inter-Company</span></td>';
    tbody.appendChild(tr);
    totalDebit  += (Number(t.debit)  || 0);
    totalCredit += (Number(t.credit) || 0);
  });

  // TOTAL row in tfoot
  var tftr = document.createElement('tr');
  tftr.className = 'total-row';
  tftr.innerHTML =
    '<td colspan="4"><strong>TOTAL INTER-COMPANY</strong></td>' +
    '<td class="amt-debit"><strong>'  + money(totalDebit)  + '</strong></td>' +
    '<td class="amt-credit"><strong>' + money(totalCredit) + '</strong></td>' +
    '<td colspan="2"></td>';
  tfoot.appendChild(tftr);

  section.hidden = false;
}

function scrollRowIntoView(idx) {
  var row = ledgerBody.children[idx];
  if (!row) return;
  
  try {
    row.scrollIntoView({ behavior: 'smooth', block: 'center' });
  } catch (e) {
    row.scrollIntoView(true);
  }
}

/* ---------- Table scroll indicator ---------- */

function updateScrollIndicator() {
  if (!tableWrap) return;
  var canScroll = tableWrap.scrollWidth > tableWrap.clientWidth;
  var isAtEnd = tableWrap.scrollLeft + tableWrap.clientWidth >= tableWrap.scrollWidth - 2;
  
  if (canScroll && !isAtEnd) {
    tableWrap.classList.add('can-scroll-right');
  } else {
    tableWrap.classList.remove('can-scroll-right');
  }
}

if (tableWrap) {
  tableWrap.addEventListener('scroll', updateScrollIndicator, { passive: true });
  window.addEventListener('resize', updateScrollIndicator, { passive: true });
}

/* ---------- Start over ---------- */

startOverBtn.addEventListener('click', function () {
  results.hidden = true;
  uploadCard.hidden = false;
  selectedFiles = [];
  fileInput.value = '';
  if (pasteInput) pasteInput.value = '';
  fileChip.hidden = true;
  updateAnalyzeButtonState();
  hideError();
  hideWarning();
  
  var reportsSection = document.getElementById('reportsSection');
  if (reportsSection) reportsSection.hidden = true;

  // Clear transactions table
  var txnSection = document.getElementById('txnDetailSection');
  if (txnSection) txnSection.hidden = true;
  var txnBody = document.getElementById('txnBody');
  if (txnBody) txnBody.innerHTML = '';
  var txnFoot = document.getElementById('txnFoot');
  if (txnFoot) txnFoot.innerHTML = '';

  // Clear inter-company table
  var icSection = document.getElementById('interCompanySection');
  if (icSection) icSection.hidden = true;
  var icBody = document.getElementById('icBody');
  if (icBody) icBody.innerHTML = '';
  var icFoot = document.getElementById('icFoot');
  if (icFoot) icFoot.innerHTML = '';
  
  // Scroll back to top
  window.scrollTo({ top: 0, behavior: 'smooth' });
});


/* ---------- Errors & Warnings ---------- */

function showError(msg) {
  errorLine.textContent = msg;
  errorLine.hidden = false;
}

function hideError() {
  errorLine.hidden = true;
  errorLine.textContent = '';
}

function showWarning(msg) {
  if (warningLine) {
    warningLine.textContent = msg;
    warningLine.hidden = false;
  }
}

function hideWarning() {
  if (warningLine) {
    warningLine.hidden = true;
    warningLine.textContent = '';
  }
}

/* ---------- Online/Offline handling ---------- */

window.addEventListener('online', function () {
  hideError();
});

window.addEventListener('offline', function () {
  showError('You are offline. Please check your internet connection.');
});
