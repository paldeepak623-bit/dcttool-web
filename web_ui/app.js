/**
 * DCTtool v2.0.0 Enterprise - Application Logic
 * Full 75 UP Districts with live search, clean launch defaults,
 * live audit & reconciliation tracking (Match / Not Match),
 * and high-speed execution pipeline.
 */

// All 75 Official Uttar Pradesh Districts
const UP_DISTRICTS = [
  "AGRA", "ALIGARH", "AMBEDKAR NAGAR", "AMETHI", "AMROHA", "AURAIYA", "AYODHYA", "AZAMGARH",
  "BAGHPAT", "BAHRAICH", "BALLIA", "BALRAMPUR", "BANDA", "BARABANKI", "BAREILLY", "BASTI",
  "BHADOHI", "BIJNOR", "BUDAUN", "BULANDSHAHR", "CHANDAULI", "CHITRAKOOT", "DEORIA", "ETAH",
  "ETAWAH", "FARRUKHABAD", "FATEHPUR", "FIROZABAD", "GAUTAM BUDDHA NAGAR", "GHAZIABAD", "GHAZIPUR",
  "GONDA", "GORAKHPUR", "HAMIRPUR", "HAPUR", "HARDOI", "HATHRAS", "JALAUN", "JAUNPUR",
  "JHANSI", "KANNAUJ", "KANPUR DEHAT", "KANPUR NAGAR", "KASGANJ", "KAUSHAMBI", "KHERI",
  "KUSHINAGAR", "LALITPUR", "LUCKNOW", "MAHARAJGANJ", "MAHOBA", "MAINPURI", "MATHURA", "MAU",
  "MEERUT", "MIRZAPUR", "MORADABAD", "MUZAFFARNAGAR", "PILIBHIT", "PRATAPGARH", "PRAYAGRAJ",
  "RAE BARELI", "RAMPUR", "SAHARANPUR", "SAMBHAL", "SANT KABIR NAGAR", "SHAHJAHANPUR", "SHAMLI",
  "SHRAVASTI", "SIDDHARTHNAGAR", "SITAPUR", "SONBHADRA", "SULTANPUR", "UNNAO", "VARANASI"
];

// Active State
let currentLoadedFile = null;
let currentAudit = null;
let isExecuting = false;

// Populate District Dropdowns on Load
function initDistrictDropdowns() {
  const d1 = document.getElementById('districtSelect');
  const d2 = document.getElementById('adminDistrictSelect');

  const populate = (sel) => {
    if (!sel) return;
    const currentVal = sel.value;
    sel.innerHTML = '<option value="">-- Select District --</option>';
    UP_DISTRICTS.forEach(d => {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = d;
      sel.appendChild(opt);
    });
    if (currentVal) sel.value = currentVal;
  };

  populate(d1);
  populate(d2);
}

// Live Filter for District Dropdown
function filterDistrictOptions(query) {
  const d1 = document.getElementById('districtSelect');
  if (!d1) return;
  const q = (query || '').trim().toUpperCase();

  d1.innerHTML = '<option value="">-- Select District --</option>';
  const filtered = q ? UP_DISTRICTS.filter(d => d.includes(q)) : UP_DISTRICTS;
  filtered.forEach(d => {
    const opt = document.createElement('option');
    opt.value = d;
    opt.textContent = d;
    d1.appendChild(opt);
  });

  // If exact or single match found, auto-select it
  if (filtered.length === 1) {
    d1.value = filtered[0];
    syncDistrict(filtered[0]);
  } else if (filtered.length > 0 && q.length >= 3) {
    d1.value = filtered[0];
    syncDistrict(filtered[0]);
  }
}

function syncDistrict(val) {
  const d1 = document.getElementById('districtSelect');
  const d2 = document.getElementById('adminDistrictSelect');
  if (d1 && d1.value !== val) d1.value = val;
  if (d2 && d2.value !== val) d2.value = val;

  const targetPacsVal = document.getElementById('cockpitPacsDisplay');
  if (targetPacsVal) {
    if (val) {
      targetPacsVal.textContent = `[Cite: 1] (${val.charAt(0) + val.slice(1).toLowerCase()})`;
    } else {
      targetPacsVal.textContent = currentLoadedFile ? '[Cite: Active PACS]' : '[No File Loaded]';
    }
  }

  const dashDist = document.getElementById('dashDistrictVal');
  if (dashDist) {
    dashDist.textContent = val || '-- Not Selected --';
  }
}

// Tab Navigation
function switchTab(tabId) {
  const navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(item => {
    if (item.dataset.tab === tabId) {
      item.classList.add('active');
    } else {
      item.classList.remove('active');
    }
  });

  const views = document.querySelectorAll('.tab-view');
  views.forEach(view => {
    if (view.id === `view-${tabId}`) {
      view.classList.add('active');
    } else {
      view.classList.remove('active');
    }
  });
}

// Bind Sidebar Nav Items
document.querySelectorAll('.sidebar-nav .nav-item').forEach(button => {
  button.addEventListener('click', () => {
    const tab = button.dataset.tab;
    if (tab) switchTab(tab);
  });
});

// Theme Toggle
const themeSwitch = document.getElementById('themeSwitch');
if (themeSwitch) {
  themeSwitch.addEventListener('change', (e) => {
    if (e.target.checked) {
      document.body.classList.add('dark-theme');
    } else {
      document.body.classList.remove('dark-theme');
    }
  });
}

// Password Visibility Toggle
function togglePwd() {
  const pwdField = document.getElementById('pwdField');
  const pwdBtn = document.getElementById('pwdEyeBtn');
  if (pwdField) {
    if (pwdField.type === 'password') {
      pwdField.type = 'text';
      if (pwdBtn) pwdBtn.textContent = '🔒';
    } else {
      pwdField.type = 'password';
      if (pwdBtn) pwdBtn.textContent = '👁️';
    }
  }
}

// Template Selection Logic
const templateCards = document.querySelectorAll('.tpl-card');
const selectAllChk = document.getElementById('selectAllChk');

function updateTemplateCount() {
  let count = 0;
  templateCards.forEach(card => {
    const chk = card.querySelector('input[type="checkbox"]');
    if (chk && chk.checked) {
      count++;
      card.classList.add('selected');
    } else {
      card.classList.remove('selected');
    }
  });

  if (selectAllChk) {
    selectAllChk.checked = count === templateCards.length;
    const parentLabel = selectAllChk.parentElement;
    if (parentLabel && parentLabel.childNodes[1]) {
      parentLabel.childNodes[1].nodeValue = ` Select all (${count}/${templateCards.length})`;
    }
  }

  // Update docked action bar count
  const metaCount = document.getElementById('dockedMetaCount');
  if (metaCount) {
    metaCount.textContent = `${currentLoadedFile ? '1 file loaded' : '0 files loaded'} | ${count} templates selected`;
  }
}

templateCards.forEach(card => {
  card.addEventListener('click', (e) => {
    const chk = card.querySelector('input[type="checkbox"]');
    if (e.target !== chk) {
      chk.checked = !chk.checked;
    }
    updateTemplateCount();
  });
});

if (selectAllChk) {
  selectAllChk.addEventListener('change', (e) => {
    const isChecked = e.target.checked;
    templateCards.forEach(card => {
      const chk = card.querySelector('input[type="checkbox"]');
      if (chk) chk.checked = isChecked;
    });
    updateTemplateCount();
  });
}

function clearAllTemplates() {
  templateCards.forEach(card => {
    const chk = card.querySelector('input[type="checkbox"]');
    if (chk) chk.checked = false;
  });
  updateTemplateCount();
}

// Currency Formatter (Indian Rupees)
function formatINR(val) {
  const num = parseFloat(val) || 0.0;
  return '₹' + num.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Update Audit UI Across Dashboard and Setup Pages
function updateAuditUI(audit) {
  if (!audit) return;
  currentAudit = audit;

  const shareVal = (audit.share_balance !== undefined) ? audit.share_balance : (audit.total_share_balance || 0);
  const stPrinVal = (audit.st_principal !== undefined) ? audit.st_principal : (audit.stloans_principal || 0);
  const ltPrinVal = (audit.lt_principal !== undefined) ? audit.lt_principal : 0;
  const stIntVal = (audit.st_interest !== undefined) ? audit.st_interest : (audit.stloans_interest || 0);
  const ltIntVal = (audit.lt_interest !== undefined) ? audit.lt_interest : 0;

  const shareStr = formatINR(shareVal);
  const stPrinStr = formatINR(stPrinVal);
  const ltPrinStr = formatINR(ltPrinVal);
  const stIntStr = formatINR(stIntVal);
  const ltIntStr = formatINR(ltIntVal);

  // 1. Dashboard Elements
  const dashShare = document.getElementById('dashShareBalance');
  if (dashShare) dashShare.textContent = shareStr;

  const dashStP = document.getElementById('dashStPrincipal');
  if (dashStP) dashStP.textContent = stPrinStr;

  const dashLtP = document.getElementById('dashLtPrincipal');
  if (dashLtP) dashLtP.textContent = ltPrinStr;

  const dashPAmount = document.getElementById('dashPrincipalAmount');
  if (dashPAmount) dashPAmount.textContent = stPrinStr;

  const dashPBadge = document.getElementById('dashPrincipalBadge');
  if (dashPBadge) {
    if (audit.has_stloans && audit.has_lt) {
      if (audit.principal_match) {
        dashPBadge.className = 'badge-match';
        dashPBadge.innerHTML = 'MATCH ✔';
      } else {
        dashPBadge.className = 'badge-mismatch';
        dashPBadge.innerHTML = 'NOT MATCH ❌';
      }
    } else {
      dashPBadge.className = 'badge-pending';
      dashPBadge.innerHTML = 'PENDING';
    }
  }

  const dashStI = document.getElementById('dashStInterest');
  if (dashStI) dashStI.textContent = stIntStr;

  const dashLtI = document.getElementById('dashLtInterest');
  if (dashLtI) dashLtI.textContent = ltIntStr;

  const dashIAmount = document.getElementById('dashInterestAmount');
  if (dashIAmount) dashIAmount.textContent = stIntStr;

  const dashIBadge = document.getElementById('dashInterestBadge');
  if (dashIBadge) {
    if (audit.has_stloans && audit.has_lt) {
      if (audit.interest_match) {
        dashIBadge.className = 'badge-match';
        dashIBadge.innerHTML = 'MATCH ✔';
      } else {
        dashIBadge.className = 'badge-mismatch';
        dashIBadge.innerHTML = 'NOT MATCH ❌';
      }
    } else {
      dashIBadge.className = 'badge-pending';
      dashIBadge.innerHTML = 'PENDING';
    }
  }

  // 2. Setup Page Audit Elements
  const setupShare = document.getElementById('setupShareBalance');
  if (setupShare) setupShare.textContent = shareStr;

  const setupStP = document.getElementById('setupStPrincipal');
  if (setupStP) setupStP.textContent = stPrinStr;

  const setupLtP = document.getElementById('setupLtPrincipal');
  if (setupLtP) setupLtP.textContent = ltPrinStr;

  const setupPAmount = document.getElementById('setupPrincipalAmount');
  if (setupPAmount) setupPAmount.textContent = stPrinStr;

  const setupPBadge = document.getElementById('setupPrincipalBadge');
  if (setupPBadge) {
    if (audit.has_stloans && audit.has_lt) {
      if (audit.principal_match) {
        setupPBadge.className = 'badge-match';
        setupPBadge.innerHTML = 'MATCH ✔';
      } else {
        setupPBadge.className = 'badge-mismatch';
        setupPBadge.innerHTML = 'NOT MATCH ❌';
      }
    } else {
      setupPBadge.className = 'badge-pending';
      setupPBadge.innerHTML = 'PENDING';
    }
  }

  const setupStI = document.getElementById('setupStInterest');
  if (setupStI) setupStI.textContent = stIntStr;

  const setupLtI = document.getElementById('setupLtInterest');
  if (setupLtI) setupLtI.textContent = ltIntStr;

  const setupIAmount = document.getElementById('setupInterestAmount');
  if (setupIAmount) setupIAmount.textContent = stIntStr;

  const setupIBadge = document.getElementById('setupInterestBadge');
  if (setupIBadge) {
    if (audit.has_stloans && audit.has_lt) {
      if (audit.interest_match) {
        setupIBadge.className = 'badge-match';
        setupIBadge.innerHTML = 'MATCH ✔';
      } else {
        setupIBadge.className = 'badge-mismatch';
        setupIBadge.innerHTML = 'NOT MATCH ❌';
      }
    } else {
      setupIBadge.className = 'badge-pending';
      setupIBadge.innerHTML = 'PENDING';
    }
  }
}

// Apply Loaded Workbook Data to UI Across Entire App
function applyLoadedFileData(data) {
  if (!data || !data.file_name) return;
  currentLoadedFile = data.file_name;

  // Transition Card 1 from Empty State to Loaded State
  const emptyDrop = document.getElementById('dropzoneEmpty');
  const loadedDrop = document.getElementById('dropzoneLoaded');
  if (emptyDrop) emptyDrop.style.display = 'none';
  if (loadedDrop) loadedDrop.style.display = 'flex';

  const nameEl = document.getElementById('loadedFileName');
  if (nameEl) nameEl.textContent = data.file_name;

  const pathEl = document.getElementById('loadedFilePath');
  if (pathEl) pathEl.textContent = data.file_path;

  // Show Remove button on Card 1
  const btnRemove = document.getElementById('btnRemoveFile');
  if (btnRemove) btnRemove.style.display = 'inline-flex';

  // Update Dashboard Quad
  const dashFile = document.getElementById('dashFileVal');
  if (dashFile) dashFile.textContent = data.file_name;

  const dashSize = document.getElementById('dashFileSize');
  if (dashSize) dashSize.textContent = data.file_size || 'Excel Workbook';

  // Update Docked Bar
  const dockedName = document.getElementById('dockedFileName');
  if (dockedName) dockedName.textContent = data.file_name;

  // Update Checklist & Progress State for Fresh File Execution
  const chkMS = document.getElementById('chkMasterSheet');
  if (chkMS) chkMS.innerHTML = `<span class="check-icon">✔</span> Master sheet loaded <span class="time">Ready</span>`;

  const chkSys = document.getElementById('chkSystemReady');
  if (chkSys) chkSys.innerHTML = `<span class="check-icon">✔</span> Ready for execution <span class="time">Ready</span>`;

  const progFill = document.getElementById('execProgFill');
  const progPct = document.getElementById('execProgPercent');
  if (progFill) progFill.style.width = '0%';
  if (progPct) progPct.textContent = '0%';

  // Auto-align detected district if found
  if (data.detected_district) {
    syncDistrict(data.detected_district);
  }

  // Update target PACS with newly detected PACS metadata
  const pacsInp = document.getElementById('pacsInput');
  if (pacsInp) {
    pacsInp.value = data.pacs_id || data.detected_pacs || '';
  }
  const adminPacs = document.getElementById('adminPacsInput');
  if (adminPacs && data.pacs_id) {
    adminPacs.value = data.pacs_id;
  }
  const cockpitPacs = document.getElementById('cockpitPacsDisplay');
  if (cockpitPacs) {
    cockpitPacs.textContent = data.pacs_name || data.pacs_id || currentLoadedFile || '[Active PACS]';
  }

  // Apply audit reconciliation if already calculated or poll in background
  if (data.audit) {
    updateAuditUI(data.audit);

    const memEl = document.getElementById('dashMemCount');
    if (memEl) memEl.textContent = `${(data.audit.membership_count || 0).toLocaleString()} Members`;

    const shareEl = document.getElementById('dashShareAccCount');
    if (shareEl) shareEl.textContent = `${(data.audit.share_members_count || 0).toLocaleString()} Accounts`;

    const stEl = document.getElementById('dashStloansCount');
    if (stEl) stEl.textContent = `${(data.audit.stloans_count || 0).toLocaleString()} Accounts`;

    const ltEl = document.getElementById('dashLtCount');
    if (ltEl) ltEl.textContent = `${(data.audit.lt_count || 0).toLocaleString()} Records`;

    const landEl = document.getElementById('dashLandCount');
    if (landEl) landEl.textContent = `${(data.audit.land_count || 0).toLocaleString()} Records`;
  } else {
    pollBackgroundAudit();
  }

  updateTemplateCount();
  appendLog(`✔ Master sheet loaded successfully: ${data.file_name} (${data.file_size || ''})`, 'green');
}

let auditPollTimer = null;
function pollBackgroundAudit(attempts = 0) {
  if (auditPollTimer) clearTimeout(auditPollTimer);
  if (attempts > 25) return;
  auditPollTimer = setTimeout(async () => {
    try {
      const res = await fetch('/api/get_audit');
      if (res.ok) {
        const d = await res.json();
        if (d.audit && (d.audit.has_membership || d.audit.has_stloans || d.audit.membership_count > 0 || d.audit.land_count > 0)) {
          updateAuditUI(d.audit);
          const memEl = document.getElementById('dashMemCount');
          if (memEl) memEl.textContent = `${(d.audit.membership_count || 0).toLocaleString()} Members`;
          const shareEl = document.getElementById('dashShareAccCount');
          if (shareEl) shareEl.textContent = `${(d.audit.share_members_count || 0).toLocaleString()} Accounts`;
          const stEl = document.getElementById('dashStloansCount');
          if (stEl) stEl.textContent = `${(d.audit.stloans_count || 0).toLocaleString()} Accounts`;
          const ltEl = document.getElementById('dashLtCount');
          if (ltEl) ltEl.textContent = `${(d.audit.lt_count || 0).toLocaleString()} Records`;
          const landEl = document.getElementById('dashLandCount');
          if (landEl) landEl.textContent = `${(d.audit.land_count || 0).toLocaleString()} Records`;
          return;
        }
      }
    } catch (e) {}
    pollBackgroundAudit(attempts + 1);
  }, 600);
}

// File Browser with Native Dialog & Universal HTML Fallback
async function browseFile() {
  try {
    const res = await fetch('/api/browse_file', { method: 'POST' });
    if (res.ok) {
      const data = await res.json();
      if (data.success) {
        applyLoadedFileData(data);
        return;
      }
      if (data.cancelled) {
        return;
      }
    }
  } catch (e) {
    console.warn("Native browse error, triggering HTML file input fallback:", e);
  }

  // Trigger HTML File Input as reliable fallback only if native failed
  const htmlInput = document.getElementById('htmlFileInput');
  if (htmlInput) {
    htmlInput.value = '';
    htmlInput.click();
  }
}

// HTML File Input Handler (Upload to Server for parsing)
async function handleHTMLFileInput(input) {
  if (!input.files || input.files.length === 0) return;
  const file = input.files[0];
  window.currentMasterFile = file;
  const formData = new FormData();
  formData.append('file', file);

  appendLog(`⏳ Uploading & analyzing master sheet: ${file.name}...`, 'blue');
  try {
    const res = await fetch('/api/upload_file', {
      method: 'POST',
      body: formData
    });
    if (res.ok) {
      const data = await res.json();
      if (data.success) {
        applyLoadedFileData(data);
      } else {
        alert("Failed to analyze master sheet: " + (data.message || "Unknown error"));
      }
    } else {
      alert(`Server response status ${res.status}: Failed to upload master sheet.`);
    }
  } catch (err) {
    console.error("Upload error:", err);
    alert("Error uploading file to engine: " + err.message);
  } finally {
    // Reset input value so consecutive files or identical names can always be selected
    input.value = '';
  }
}

// Drag & Drop Handlers for Dropzone
function handleDragOver(e) {
  e.preventDefault();
  e.stopPropagation();
  const box = document.getElementById('dropzoneEmpty');
  if (box) box.classList.add('dragover');
}

function handleDragLeave(e) {
  e.preventDefault();
  e.stopPropagation();
  const box = document.getElementById('dropzoneEmpty');
  if (box) box.classList.remove('dragover');
}

function handleDropFile(e) {
  e.preventDefault();
  e.stopPropagation();
  const box = document.getElementById('dropzoneEmpty');
  if (box) box.classList.remove('dragover');

  if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
    const input = document.getElementById('htmlFileInput');
    if (input) {
      input.files = e.dataTransfer.files;
      handleHTMLFileInput(input);
    }
  }
}

async function getSampleTemplates() {
  try {
    const res = await fetch('/api/get_samples', { method: 'POST' });
    if (res.ok) {
      const data = await res.json();
      alert(`Exported ${data.count} sample templates to:\n${data.dest}`);
    }
  } catch (e) {
    alert("Sample templates are located in DCT-V2.0/dcttool/templates");
  }
}

async function openOutputFolder() {
  try {
    await fetch('/api/open_output', { method: 'POST' });
  } catch (e) {
    alert("Output files are generated in Desktop folder: 'DCT -1'");
  }
}

async function removeMasterSheet() {
  try {
    await fetch('/api/remove_file', { method: 'POST' });
  } catch (e) {
    console.warn("Remove file api error:", e);
  }
  currentLoadedFile = null;

  const emptyDrop = document.getElementById('dropzoneEmpty');
  const loadedDrop = document.getElementById('dropzoneLoaded');
  const btnRemove = document.getElementById('btnRemoveFile');
  if (emptyDrop) emptyDrop.style.display = 'flex';
  if (loadedDrop) loadedDrop.style.display = 'none';
  if (btnRemove) btnRemove.style.display = 'none';

  const htmlInput = document.getElementById('htmlFileInput');
  if (htmlInput) htmlInput.value = '';

  const nameEl = document.getElementById('loadedFileName');
  if (nameEl) nameEl.textContent = 'No file selected';

  const pathEl = document.getElementById('loadedFilePath');
  if (pathEl) pathEl.textContent = '-';

  const dashFile = document.getElementById('dashFileVal');
  if (dashFile) dashFile.textContent = 'No File Loaded';

  const dashSize = document.getElementById('dashFileSize');
  if (dashSize) dashSize.textContent = 'Awaiting upload';

  const dockedName = document.getElementById('dockedFileName');
  if (dockedName) dockedName.textContent = 'No file loaded';

  const chkMS = document.getElementById('chkMasterSheet');
  if (chkMS) chkMS.innerHTML = `<span class="check-icon">⏳</span> Master sheet input <span class="time" id="timeMasterSheet">-</span>`;

  const chkSys = document.getElementById('chkSystemReady');
  if (chkSys) chkSys.innerHTML = `<span class="check-icon">⏳</span> Ready for execution <span class="time" id="timeReady">-</span>`;

  // Reset Records Breakdown
  const memEl = document.getElementById('dashMemCount');
  if (memEl) memEl.textContent = '0 Members';
  const shareEl = document.getElementById('dashShareAccCount');
  if (shareEl) shareEl.textContent = '0 Accounts';
  const stEl = document.getElementById('dashStloansCount');
  if (stEl) stEl.textContent = '0 Accounts';
  const ltEl = document.getElementById('dashLtCount');
  if (ltEl) ltEl.textContent = '0 Records';
  const landEl = document.getElementById('dashLandCount');
  if (landEl) landEl.textContent = '0 Records';

  // Reset Audit
  updateAuditUI({
    share_balance: 0,
    share_members_count: 0,
    st_principal: 0,
    st_interest: 0,
    lt_principal: 0,
    lt_interest: 0,
    has_membership: false,
    has_stloans: false,
    has_lt: false,
  });

  appendLog('ℹ Master sheet removed by user.', 'blue');
}

function clearConsoleLogs() {
  const setupLog = document.getElementById('setupLogStream');
  if (setupLog) {
    setupLog.innerHTML = `
      <div class="log-line blue">⚡ Logs cleared by user</div>
      <div class="log-line white">💡 Ready for next action</div>
    `;
  }
}

function appendLog(msg, color = 'white') {
  const setupLog = document.getElementById('setupLogStream');
  if (setupLog) {
    const line = document.createElement('div');
    line.className = `log-line ${color}`;
    line.textContent = msg;
    setupLog.appendChild(line);
    setupLog.scrollTop = setupLog.scrollHeight;
  }
}

// Portal Select Change Handler
const portalSelect = document.getElementById('portalSelect');
if (portalSelect) {
  portalSelect.addEventListener('change', (e) => {
    const u = e.target.value;
    const dashPortal = document.getElementById('dashPortalVal');
    if (dashPortal) dashPortal.textContent = u || '-- Not Selected --';

    const chkPortal = document.getElementById('chkPortalConn');
    if (chkPortal) {
      chkPortal.innerHTML = `<span class="check-icon">✔</span> Portal gateway selected <span class="time">Ready</span>`;
    }
  });
}

// ---------------------------------------------------------------------------
// HOSPITAL ECG MONITOR & AUDIO ENGINE (SCREENSHOT 3 REPLICA)
// ---------------------------------------------------------------------------
let ecgAnimationId = null;
let ecgAudioCtx = null;
let ecgIsMuted = false;
let ecgSweepX = 0;
let ecgHistory = [];

function initEcgAudio() {
  if (!ecgAudioCtx) {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (AudioCtx) ecgAudioCtx = new AudioCtx();
    } catch (e) {
      console.warn("AudioContext init error:", e);
    }
  }
  if (ecgAudioCtx && ecgAudioCtx.state === 'suspended') {
    ecgAudioCtx.resume();
  }
}

function playEcgBeep() {
  if (ecgIsMuted) return;
  initEcgAudio();
  if (!ecgAudioCtx) return;

  try {
    const now = ecgAudioCtx.currentTime;
    const osc = ecgAudioCtx.createOscillator();
    const gain = ecgAudioCtx.createGain();

    osc.type = 'sine';
    osc.frequency.setValueAtTime(880, now); // Authentic clinical ECG frequency (A5)

    gain.gain.setValueAtTime(0.08, now);
    gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.09); // Short 90ms clinical beep

    osc.connect(gain);
    gain.connect(ecgAudioCtx.destination);

    osc.start(now);
    osc.stop(now + 0.09);
  } catch (err) {
    console.warn("ECG Beep error:", err);
  }
}

function playCompletionChime() {
  if (ecgIsMuted) return;
  initEcgAudio();
  if (!ecgAudioCtx) return;

  try {
    const now = ecgAudioCtx.currentTime;
    // Two-tone medical success chime (587Hz -> 880Hz)
    const osc1 = ecgAudioCtx.createOscillator();
    const gain1 = ecgAudioCtx.createGain();
    osc1.type = 'sine';
    osc1.frequency.setValueAtTime(587.33, now);
    gain1.gain.setValueAtTime(0.08, now);
    gain1.gain.exponentialRampToValueAtTime(0.0001, now + 0.16);
    osc1.connect(gain1);
    gain1.connect(ecgAudioCtx.destination);
    osc1.start(now);
    osc1.stop(now + 0.16);

    const osc2 = ecgAudioCtx.createOscillator();
    const gain2 = ecgAudioCtx.createGain();
    osc2.type = 'sine';
    osc2.frequency.setValueAtTime(880, now + 0.13);
    gain2.gain.setValueAtTime(0.08, now + 0.13);
    gain2.gain.exponentialRampToValueAtTime(0.0001, now + 0.35);
    osc2.connect(gain2);
    gain2.connect(ecgAudioCtx.destination);
    osc2.start(now + 0.13);
    osc2.stop(now + 0.35);
  } catch (err) {
    console.warn("Chime error:", err);
  }
}

function toggleEcgMute(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  ecgIsMuted = !ecgIsMuted;
  const btn = document.getElementById('ecgMuteBtn');
  if (btn) {
    btn.textContent = ecgIsMuted ? '🔇' : '🔊';
    btn.title = ecgIsMuted ? 'Unmute Hospital Monitor Beep' : 'Mute Hospital Monitor Beep';
  }
}

function getEcgY(tNorm) {
  // P-Q-R-S-T Cardiac curve
  if (tNorm < 0.15) {
    return 0;
  } else if (tNorm < 0.25) {
    const p = (tNorm - 0.15) / 0.10;
    return Math.sin(p * Math.PI) * 0.18; // P wave
  } else if (tNorm < 0.36) {
    return 0; // PR segment
  } else if (tNorm < 0.39) {
    const q = (tNorm - 0.36) / 0.03;
    return -Math.sin(q * Math.PI) * 0.14; // Q dip
  } else if (tNorm < 0.44) {
    const r = (tNorm - 0.39) / 0.05;
    return Math.sin(r * Math.PI) * 0.94; // R peak spike!
  } else if (tNorm < 0.48) {
    const s = (tNorm - 0.44) / 0.04;
    return -Math.sin(s * Math.PI) * 0.30; // S dip
  } else if (tNorm < 0.60) {
    return 0; // ST segment
  } else if (tNorm < 0.80) {
    const t = (tNorm - 0.60) / 0.20;
    return Math.sin(t * Math.PI) * 0.28; // T wave
  } else {
    return 0;
  }
}

function startEcgMonitor() {
  const box = document.getElementById('ecgMonitorBox');
  const checklist = document.getElementById('statusChecklist');
  if (box) box.style.display = 'block';
  if (checklist) checklist.style.display = 'none';

  const stepEl = document.getElementById('ecgStatusStep');
  if (stepEl) {
    stepEl.textContent = '⚡ Starting automated template pipeline...';
    stepEl.style.color = '#94a3b8';
  }
  const liveEl = document.getElementById('ecgLiveTime');
  if (liveEl) {
    liveEl.textContent = '● PULSE ACTIVE';
    liveEl.style.color = '#10b981';
  }

  initEcgAudio();

  const canvas = document.getElementById('ecgCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const w = canvas.parentElement ? canvas.parentElement.clientWidth : 340;
  const h = 96;
  canvas.width = w > 50 ? w : 340;
  canvas.height = h;

  ecgHistory = new Array(canvas.width).fill(h * 0.58);
  ecgSweepX = 0;
  let cyclePos = 0;
  const cycleWidth = 140; // ~72 BPM
  let lastRTime = 0;

  if (ecgAnimationId) cancelAnimationFrame(ecgAnimationId);

  const speed = 2.4;

  function drawGrid() {
    ctx.fillStyle = '#040812';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Minor grid (12px)
    ctx.lineWidth = 0.5;
    ctx.strokeStyle = 'rgba(0, 180, 255, 0.09)';
    ctx.beginPath();
    for (let x = 0; x < canvas.width; x += 12) {
      ctx.moveTo(x, 0); ctx.lineTo(x, canvas.height);
    }
    for (let y = 0; y < canvas.height; y += 12) {
      ctx.moveTo(0, y); ctx.lineTo(canvas.width, y);
    }
    ctx.stroke();

    // Major grid (48px)
    ctx.lineWidth = 1;
    ctx.strokeStyle = 'rgba(0, 180, 255, 0.22)';
    ctx.beginPath();
    for (let x = 0; x < canvas.width; x += 48) {
      ctx.moveTo(x, 0); ctx.lineTo(x, canvas.height);
    }
    for (let y = 0; y < canvas.height; y += 48) {
      ctx.moveTo(0, y); ctx.lineTo(canvas.width, y);
    }
    ctx.stroke();
  }

  function renderFrame(timestamp) {
    if (!isExecuting) return;

    const prevX = ecgSweepX;
    ecgSweepX = (ecgSweepX + speed) % canvas.width;

    const stepCount = Math.max(1, Math.round(speed));
    for (let i = 0; i < stepCount; i++) {
      const curX = Math.round((prevX + i) % canvas.width);
      cyclePos = (cyclePos + 1) % cycleWidth;
      const tNorm = cyclePos / cycleWidth;
      const val = getEcgY(tNorm);

      const baselineY = canvas.height * 0.58;
      const amp = canvas.height * 0.48;
      const y = baselineY - (val * amp);
      ecgHistory[curX] = y;

      if (tNorm >= 0.40 && tNorm <= 0.43) {
        if (!lastRTime || (timestamp - lastRTime) > 600) {
          lastRTime = timestamp;
          playEcgBeep();

          const heart = document.getElementById('ecgHeartIcon');
          if (heart) {
            heart.style.transform = 'scale(1.4)';
            setTimeout(() => { if (heart) heart.style.transform = 'scale(1)'; }, 140);
          }
          const bpmEl = document.getElementById('ecgBpmDisplay');
          if (bpmEl) {
            const bpm = 76 + Math.floor(Math.sin(timestamp / 3000) * 4) + Math.floor(Math.random() * 2);
            bpmEl.textContent = `${bpm} BPM`;
          }
        }
      }
    }

    drawGrid();

    ctx.save();
    ctx.lineWidth = 2.4;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#00f0ff';
    ctx.shadowColor = '#00f0ff';
    ctx.shadowBlur = 9;

    const eraseGap = 24;
    const sweepHead = Math.round(ecgSweepX);

    const startTail = (sweepHead + eraseGap) % canvas.width;
    if (startTail > sweepHead) {
      ctx.beginPath();
      let started = false;
      for (let x = startTail; x < canvas.width; x++) {
        if (!started) { ctx.moveTo(x, ecgHistory[x]); started = true; }
        else { ctx.lineTo(x, ecgHistory[x]); }
      }
      ctx.stroke();
    }

    ctx.beginPath();
    let started2 = false;
    for (let x = 0; x < sweepHead; x++) {
      if (!started2) { ctx.moveTo(x, ecgHistory[x]); started2 = true; }
      else { ctx.lineTo(x, ecgHistory[x]); }
    }
    ctx.stroke();

    const headY = ecgHistory[sweepHead] || (canvas.height * 0.58);
    ctx.fillStyle = '#ffffff';
    ctx.shadowColor = '#00f0ff';
    ctx.shadowBlur = 14;
    ctx.beginPath();
    ctx.arc(sweepHead, headY, 3.2, 0, Math.PI * 2);
    ctx.fill();

    ctx.restore();

    ecgAnimationId = requestAnimationFrame(renderFrame);
  }

  ecgAnimationId = requestAnimationFrame(renderFrame);
}

function updateEcgStep(prog) {
  const stepEl = document.getElementById('ecgStatusStep');
  if (!stepEl) return;
  if (prog < 35) {
    stepEl.textContent = '🌐 Syncing village & portal data...';
  } else if (prog < 55) {
    stepEl.textContent = '📑 Processing Short Term Loans...';
  } else if (prog < 75) {
    stepEl.textContent = '👥 Validating Membership & Land Details...';
  } else if (prog < 90) {
    stepEl.textContent = '💳 Reconciling Loan Transactions...';
  } else {
    stepEl.textContent = '💾 Finalizing & Exporting Templates...';
  }
}

function stopEcgMonitor(isSuccess = true) {
  if (ecgAnimationId) {
    cancelAnimationFrame(ecgAnimationId);
    ecgAnimationId = null;
  }

  const canvas = document.getElementById('ecgCanvas');
  if (canvas) {
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#040812';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.lineWidth = 0.5;
    ctx.strokeStyle = 'rgba(0, 180, 255, 0.08)';
    ctx.beginPath();
    for (let x = 0; x < canvas.width; x += 12) { ctx.moveTo(x, 0); ctx.lineTo(x, canvas.height); }
    for (let y = 0; y < canvas.height; y += 12) { ctx.moveTo(0, y); ctx.lineTo(canvas.width, y); }
    ctx.stroke();

    ctx.save();
    ctx.lineWidth = 2.2;
    ctx.strokeStyle = isSuccess ? '#10b981' : '#ef4444';
    ctx.shadowColor = isSuccess ? '#10b981' : '#ef4444';
    ctx.shadowBlur = 8;
    ctx.beginPath();
    const midY = canvas.height * 0.58;
    ctx.moveTo(0, midY);
    ctx.lineTo(canvas.width, midY);
    ctx.stroke();
    ctx.restore();
  }

  const stepEl = document.getElementById('ecgStatusStep');
  if (stepEl) {
    stepEl.textContent = isSuccess ? '✔ Templates generated successfully!' : '❌ Automation stopped';
    stepEl.style.color = isSuccess ? '#10b981' : '#ef4444';
  }
  const liveEl = document.getElementById('ecgLiveTime');
  if (liveEl) {
    liveEl.textContent = isSuccess ? 'COMPLETE' : 'STOPPED';
    liveEl.style.color = isSuccess ? '#10b981' : '#ef4444';
  }

  if (isSuccess) {
    playCompletionChime();
  }

  setTimeout(() => {
    const box = document.getElementById('ecgMonitorBox');
    const checklist = document.getElementById('statusChecklist');
    if (box) box.style.display = 'none';
    if (checklist) checklist.style.display = 'block';
  }, 2400);
}

// Process Execution & Live Logs
async function startExecution() {
  if (isExecuting) return;

  if (!currentLoadedFile) {
    alert("Kripya pahle 'Browse master sheet' button se Master Excel sheet select karein!");
    return;
  }

  isExecuting = true;

  // Launch Live Hospital ECG Monitor & Sound
  startEcgMonitor();

  const btnSetup = document.getElementById('btnStartSetup');
  const btnCard = document.getElementById('btnStartCard');
  if (btnSetup) {
    btnSetup.disabled = true;
    btnSetup.innerHTML = '⏳ Processing templates...';
  }
  if (btnCard) {
    btnCard.disabled = true;
    btnCard.innerHTML = '⏳ Processing templates...';
  }

  // Update Execution status card
  const execStatusDot = document.getElementById('execStatusDot');
  const execStatusText = document.getElementById('execStatusText');
  const execStatusSummary = document.getElementById('execStatusSummary');
  const execProgFill = document.getElementById('execProgFill');
  const execProgPercent = document.getElementById('execProgPercent');

  if (execStatusDot) execStatusDot.className = 'badge-dot-yellow';
  if (execStatusText) execStatusText.textContent = 'Processing...';
  if (execStatusSummary) execStatusSummary.textContent = 'Executing 28 business rules and template population...';
  if (execProgFill) execProgFill.style.width = '20%';
  if (execProgPercent) execProgPercent.textContent = '20%';

  const consoleBox = document.getElementById('processConsole');
  if (consoleBox) {
    consoleBox.innerHTML = '';
    consoleBox.classList.add('has-logs');
  }

  const pacsInput = document.getElementById('pacsInput');
  const distSelect = document.getElementById('districtSelect');
  const portSelect = document.getElementById('portalSelect');
  const loginInput = document.getElementById('loginInput');
  const pwdField = document.getElementById('pwdField');

  // Collect template states
  const tplStates = {};
  document.querySelectorAll('.tpl-card').forEach(card => {
    const key = card.dataset.tpl;
    const chk = card.querySelector('input[type="checkbox"]');
    if (key && chk) {
      tplStates[key] = chk.checked;
    }
  });

  const payload = {
    file_name: currentLoadedFile || (window.currentMasterFile ? window.currentMasterFile.name : ""),
    file_path: (typeof currentLoadedFilePath !== 'undefined' ? currentLoadedFilePath : ""),
    pacs_id: pacsInput ? pacsInput.value.trim() : "",
    district: distSelect ? distSelect.value : "",
    portal_url: portSelect ? portSelect.value : "",
    login_id: loginInput ? loginInput.value.trim() : "",
    password: pwdField ? pwdField.value : "",
    templates: tplStates
  };

  const startMsg = `⚡ Starting automation pipeline for: ${currentLoadedFile}...`;
  appendLog(startMsg, "blue");
  appendConsoleLog(startMsg, "blue");

  const portalMsg = `🌐 Synchronizing registered village and land details from portal...`;
  appendLog(portalMsg, "white");
  appendConsoleLog(portalMsg, "white");

  // Animate progress smoothly while backend is processing
  let currentProg = 20;
  const progressTicker = setInterval(() => {
    if (currentProg < 88) {
      currentProg += Math.floor(Math.random() * 8) + 4;
      if (currentProg > 88) currentProg = 88;
      if (execProgFill) execProgFill.style.width = `${currentProg}%`;
      if (execProgPercent) execProgPercent.textContent = `${currentProg}%`;
      updateEcgStep(currentProg);
    }
  }, 1500);

  try {
    let res = await fetch('/api/start_process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    // Seamless auto-recovery: If Render restarted or session expired, auto-reupload file in background!
    if (res.status === 400 && window.currentMasterFile) {
      try {
        const cloneRes = res.clone();
        const errJson = await cloneRes.json();
        if (errJson && errJson.error && (errJson.error.includes("expired") || errJson.error.includes("not found"))) {
          appendLog("🔄 Restoring master sheet upload to cloud container...", "blue");
          appendConsoleLog("🔄 Restoring master sheet upload to cloud container...", "blue");
          const reFormData = new FormData();
          reFormData.append('file', window.currentMasterFile);
          const upRes = await fetch('/api/upload_file', { method: 'POST', body: reFormData });
          if (upRes.ok) {
            const upData = await upRes.json();
            if (upData.success) {
              appendLog("✔ Master sheet restored! Starting automation...", "green");
              appendConsoleLog("✔ Master sheet restored! Starting automation...", "green");
              payload.file_name = upData.file_name;
              payload.file_path = upData.file_path;
              res = await fetch('/api/start_process', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
              });
            }
          }
        }
      } catch (reErr) {
        console.warn("Auto-recover retry error:", reErr);
      }
    }

    clearInterval(progressTicker);

    if (res.ok) {
      const data = await res.json();
      if (data.success) {
        // Render backend logs
        if (data.logs && data.logs.length > 0) {
          data.logs.forEach(line => {
            const color = (line.includes('✔') || line.includes('SUCCESS') || line.includes('MATCH ✔')) ? 'green' : 
                          (line.includes('❌') || line.includes('ERROR') || line.includes('NOT MATCH') || line.includes('failed')) ? 'red' : 
                          (line.includes('⚡') || line.includes('Writing') || line.includes('[PORTAL]') || line.includes('[MANDAL]')) ? 'blue' : 'white';
            appendLog(line, color);
            appendConsoleLog(line, color);
          });
        }

        if (data.audit) {
          updateAuditUI(data.audit);
        }

        if (execProgFill) execProgFill.style.width = '100%';
        if (execProgPercent) execProgPercent.textContent = '100%';
        if (execStatusDot) execStatusDot.className = 'badge-dot-green';
        if (execStatusText) execStatusText.textContent = 'Completed ✔';
        if (execStatusSummary) execStatusSummary.textContent = 'All selected templates generated successfully!';

        const chkReady = document.getElementById('chkSystemReady');
        if (chkReady) chkReady.innerHTML = `<span class="check-icon">✔</span> Process completed <span class="time">Done</span>`;

        const declC = document.getElementById('cockpitDeclCount');
        if (declC) declC.textContent = 'Completed';
        const loanC = document.getElementById('cockpitLoanCount');
        if (loanC) loanC.textContent = 'Completed';

        // Stop ECG monitor with Success State
        stopEcgMonitor(true);

        appendLog("✨ AUTOMATION EXECUTION COMPLETED SUCCESSFULLY ✨", "green");
        appendConsoleLog("✨ AUTOMATION EXECUTION COMPLETED SUCCESSFULLY ✨", "green");
        if (data.export_dir) {
          appendLog(`📂 Templates exported to: ${data.export_dir}`, "blue");
          appendConsoleLog(`📂 Templates exported to: ${data.export_dir}`, "blue");
        }

        if (data.download_url) {
          appendLog(`📥 Downloading split templates ZIP: ${data.export_dir || 'Templates.zip'}...`, "green");
          appendConsoleLog(`📥 Downloading split templates ZIP: ${data.export_dir || 'Templates.zip'}...`, "green");
          const dlLink = document.createElement('a');
          dlLink.href = data.download_url;
          dlLink.download = data.export_dir || 'DCT_Split_Templates.zip';
          document.body.appendChild(dlLink);
          dlLink.click();
          document.body.removeChild(dlLink);
        }

        alert(`✅ Master Process Completed Successfully!\n\nAll validated templates have been processed and your download has started.`);
      } else {
        // data.success == false
        stopEcgMonitor(false);
        if (execProgFill) execProgFill.style.width = '0%';
        if (execProgPercent) execProgPercent.textContent = '0%';
        if (execStatusDot) execStatusDot.className = 'badge-dot-red';
        if (execStatusText) execStatusText.textContent = 'Failed ❌';
        if (execStatusSummary) execStatusSummary.textContent = data.error || 'Execution failed';

        const errMsg = `❌ Error: ${data.error || 'Execution failed'}`;
        appendLog(errMsg, "red");
        appendConsoleLog(errMsg, "red");

        if (data.logs && data.logs.length > 0) {
          data.logs.forEach(line => {
            appendLog(line, line.includes('ERROR') ? 'red' : 'white');
            appendConsoleLog(line, line.includes('ERROR') ? 'red' : 'white');
          });
        }

        alert(`❌ Template Processing Failed:\n\n${data.error || 'Unknown error occurred. Please check logs.'}`);
      }
    } else {
      clearInterval(progressTicker);
      stopEcgMonitor(false);
      let httpErrMsg = `Server returned HTTP status ${res.status}`;
      try {
        const errJson = await res.json();
        if (errJson && errJson.error) {
          httpErrMsg = errJson.error;
        }
        if (errJson && errJson.logs && errJson.logs.length > 0) {
          errJson.logs.forEach(line => {
            appendLog(line, "red");
            appendConsoleLog(line, "red");
          });
        }
      } catch (_) {}
      appendLog(`❌ ${httpErrMsg}`, "red");
      appendConsoleLog(`❌ ${httpErrMsg}`, "red");
      alert(`❌ Error:\n\n${httpErrMsg}`);
    }
  } catch (err) {
    clearInterval(progressTicker);
    stopEcgMonitor(false);
    console.error("API error:", err);
    appendLog(`❌ Connection Error: ${err.message}`, "red");
    appendConsoleLog(`❌ Connection Error: ${err.message}`, "red");
    alert(`❌ Connection Error: ${err.message}`);
  } finally {
    clearInterval(progressTicker);
    isExecuting = false;
    if (btnSetup) {
      btnSetup.disabled = false;
      btnSetup.innerHTML = '▶ Start master process';
    }
    if (btnCard) {
      btnCard.disabled = false;
      btnCard.innerHTML = '▶ Start Automation Process';
    }
  }
}

function appendConsoleLog(text, color = 'white') {
  const consoleBox = document.getElementById('processConsole');
  if (!consoleBox) return;
  const line = document.createElement('div');
  line.className = `log-line ${color}`;
  line.textContent = text;
  consoleBox.appendChild(line);
  consoleBox.scrollTop = consoleBox.scrollHeight;
}

function clearConsole() {
  const consoleBox = document.getElementById('processConsole');
  if (!consoleBox) return;
  consoleBox.innerHTML = `
    <div class="empty-state">
      <div class="empty-screen-icon">
        <span class="screen-top-dots">•••</span>
        <span class="screen-prompt">&gt; _</span>
      </div>
      <h3 class="empty-title">Ready for Execution</h3>
      <p class="empty-desc">Click 'Start master process' to run automated template generation with full reconciliation audit.</p>
    </div>
  `;
}

function exportConsoleLogs() {
  const consoleBox = document.getElementById('processConsole');
  if (!consoleBox) return;
  const text = consoleBox.innerText;
  const blob = new Blob([text], { type: 'text/plain' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `DCTtool_Execution_Log_${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
  a.click();
}

function saveSettings() {
  const adminPacs = document.getElementById('adminPacsInput');
  const adminPort = document.getElementById('adminPortalSelect');
  const adminDist = document.getElementById('adminDistrictSelect');

  if (adminPacs) {
    const p1 = document.getElementById('pacsInput');
    if (p1) p1.value = adminPacs.value.trim();
  }

  if (adminPort && adminPort.value) {
    const p2 = document.getElementById('portalSelect');
    if (p2) {
      p2.value = adminPort.value;
      const dp = document.getElementById('dashPortalVal');
      if (dp) dp.textContent = adminPort.value;
    }
  }

  if (adminDist && adminDist.value) {
    syncDistrict(adminDist.value);
  }

  alert("Admin automation preferences saved successfully!");
  switchTab('setup');
}

// Initialise on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  initDistrictDropdowns();
  updateTemplateCount();
});
