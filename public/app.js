const alertEl = document.getElementById('alert');

const cardClients = document.getElementById('card-clients');
const clientSelect = document.getElementById('client-select');
const addClientBtn = document.getElementById('add-client-btn');
const deleteClientBtn = document.getElementById('delete-client-btn');
const activeClientNameEl = document.getElementById('active-client-name');
const addClientForm = document.getElementById('add-client-form');
const newClientName = document.getElementById('new-client-name');
const newClientGa4 = document.getElementById('new-client-ga4');
const newClientSheet = document.getElementById('new-client-sheet');
const cancelClientBtn = document.getElementById('cancel-client-btn');

const stepSignin = document.getElementById('step-signin');
const stepGenerate = document.getElementById('step-generate');
const userBar = document.getElementById('user-bar');
const headerEmailEl = document.getElementById('header-email');

const currentMonthSelect = document.getElementById('current-month-select');
const currentYearSelect = document.getElementById('current-year-select');
const comparisonMonthSelect = document.getElementById('comparison-month-select');
const comparisonYearSelect = document.getElementById('comparison-year-select');

const generateBtn = document.getElementById('generate-btn');
const statusEl = document.getElementById('status');
const warningsEl = document.getElementById('warnings');

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

let clients = [];
let activeClientId = null;

function showAlert(message, type) {
  alertEl.textContent = message;
  alertEl.className = `alert ${type}`;
  alertEl.hidden = false;
}

function showStatus(message, type) {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.hidden = false;
}

function handleUrlParams() {
  const params = new URLSearchParams(window.location.search);
  if (params.get('authError')) {
    showAlert(params.get('authError'), 'error');
  } else if (params.get('authSuccess')) {
    showAlert('Successfully connected to Google Analytics.', 'success');
  }
  if (params.toString()) {
    window.history.replaceState({}, document.title, '/');
  }
}

// --- Client management ---

async function loadClients() {
  const res = await fetch('/api/clients');
  clients = await res.json();
  cardClients.hidden = false;
  renderClientSelect();
}

function renderClientSelect() {
  clientSelect.innerHTML = '';
  clients.forEach((c) => {
    const opt = document.createElement('option');
    opt.value = c.id;
    opt.textContent = c.name;
    clientSelect.appendChild(opt);
  });

  if (clients.length === 0) {
    activeClientId = null;
    activeClientNameEl.textContent = 'No clients yet — add one to get started';
    deleteClientBtn.disabled = true;
    updateGenerateButtonState();
    return;
  }

  const stillExists = clients.some((c) => c.id === activeClientId);
  if (!stillExists) {
    activeClientId = clients[0].id;
  }
  clientSelect.value = activeClientId;
  deleteClientBtn.disabled = false;
  updateActiveClientLabel();
}

function updateActiveClientLabel() {
  const client = clients.find((c) => c.id === activeClientId);
  activeClientNameEl.textContent = client ? client.name : '—';
  updateGenerateButtonState();
}

function toggleAddClientForm(show) {
  addClientForm.hidden = !show;
  if (show) {
    newClientName.value = '';
    newClientGa4.value = '';
    newClientSheet.value = '';
    newClientName.focus();
  }
}

async function saveClient(e) {
  e.preventDefault();
  const name = newClientName.value.trim();
  const ga4PropertyId = newClientGa4.value.trim();
  const sheetId = newClientSheet.value.trim();

  if (!name || !ga4PropertyId) {
    showAlert('Client name and GA4 Property ID are required.', 'error');
    return;
  }

  try {
    const res = await fetch('/api/clients', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, ga4PropertyId, sheetId }),
    });
    const data = await res.json();
    if (!res.ok) {
      showAlert(data.error || 'Failed to save client.', 'error');
      return;
    }
    toggleAddClientForm(false);
    activeClientId = data.id;
    await loadClients();
    showAlert(`Client "${data.name}" added.`, 'success');
  } catch (e) {
    showAlert('Network error while saving the client.', 'error');
  }
}

async function deleteActiveClient() {
  const client = clients.find((c) => c.id === activeClientId);
  if (!client) return;
  if (!window.confirm(`Remove client "${client.name}"? This cannot be undone.`)) return;

  try {
    const res = await fetch(`/api/clients/${client.id}`, { method: 'DELETE' });
    if (!res.ok && res.status !== 204) {
      const data = await res.json().catch(() => ({}));
      showAlert(data.error || 'Failed to remove client.', 'error');
      return;
    }
    activeClientId = null;
    await loadClients();
    showAlert(`Client "${client.name}" removed.`, 'success');
  } catch (e) {
    showAlert('Network error while removing the client.', 'error');
  }
}

// --- Period pickers ---

function populatePeriodSelectors() {
  [currentMonthSelect, comparisonMonthSelect].forEach((select) => {
    select.innerHTML = '';
    MONTH_NAMES.forEach((name, i) => {
      const opt = document.createElement('option');
      opt.value = String(i + 1).padStart(2, '0');
      opt.textContent = name;
      select.appendChild(opt);
    });
  });

  const now = new Date();
  const currentYear = now.getFullYear();
  const years = [currentYear, currentYear - 1, currentYear - 2];
  [currentYearSelect, comparisonYearSelect].forEach((select) => {
    select.innerHTML = '';
    years.forEach((year) => {
      const opt = document.createElement('option');
      opt.value = String(year);
      opt.textContent = String(year);
      select.appendChild(opt);
    });
  });

  // Current Period defaults to the last completed month (e.g. today is July 2026 -> June 2026).
  let lastCompletedMonth = now.getMonth(); // 0-11, i.e. "last month" as a 1-based index
  let lastCompletedYear = currentYear;
  if (lastCompletedMonth === 0) {
    lastCompletedMonth = 12;
    lastCompletedYear = currentYear - 1;
  }
  currentMonthSelect.value = String(lastCompletedMonth).padStart(2, '0');
  currentYearSelect.value = String(lastCompletedYear);

  // Comparison Period defaults to the month before the Current Period.
  let comparisonMonth = lastCompletedMonth - 1;
  let comparisonYear = lastCompletedYear;
  if (comparisonMonth === 0) {
    comparisonMonth = 12;
    comparisonYear = lastCompletedYear - 1;
  }
  comparisonMonthSelect.value = String(comparisonMonth).padStart(2, '0');
  comparisonYearSelect.value = String(comparisonYear);

  updateGenerateButtonState();
}

function updateGenerateButtonState() {
  const periodsComplete =
    currentMonthSelect.value && currentYearSelect.value && comparisonMonthSelect.value && comparisonYearSelect.value;
  generateBtn.disabled = !(periodsComplete && activeClientId);
}

function periodLabel(monthSelect, yearSelect) {
  const monthIdx = Number(monthSelect.value) - 1;
  return `${MONTH_NAMES[monthIdx]} ${yearSelect.value}`;
}

function showWarnings(warnings) {
  if (!warnings || warnings.length === 0) {
    warningsEl.hidden = true;
    warningsEl.innerHTML = '';
    return;
  }
  warningsEl.innerHTML = warnings.map((w) => `<p>${w}</p>`).join('');
  warningsEl.hidden = false;
}

async function checkAuthStatus() {
  const res = await fetch('/auth/status');
  const { authenticated, needsReauth, email } = await res.json();

  userBar.hidden = !authenticated;
  headerEmailEl.textContent = authenticated ? email || '' : '';

  if (authenticated && !needsReauth) {
    stepSignin.hidden = true;
    stepGenerate.hidden = false;
    populatePeriodSelectors();
  } else {
    stepSignin.hidden = false;
    stepGenerate.hidden = true;
    if (authenticated && needsReauth) {
      showAlert(
        'Please reconnect your Google account to grant access to Google Sheets for SEO reporting.',
        'error'
      );
    }
  }
}

async function generateReport() {
  if (!activeClientId) {
    showAlert('Please select or add a client first.', 'error');
    return;
  }

  const currentMonth = `${currentYearSelect.value}-${currentMonthSelect.value}`;
  const comparisonMonth = `${comparisonYearSelect.value}-${comparisonMonthSelect.value}`;
  const currentLabel = periodLabel(currentMonthSelect, currentYearSelect);
  const comparisonLabel = periodLabel(comparisonMonthSelect, comparisonYearSelect);

  generateBtn.disabled = true;
  showWarnings([]);
  showStatus('Fetching data from Google Analytics and building your report...', '');

  try {
    const res = await fetch('/api/generate-report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ clientId: activeClientId, currentMonth, comparisonMonth }),
    });

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      if (res.status === 401) {
        showStatus(data.error || 'Your session expired. Please sign in again.', 'error');
        await checkAuthStatus();
        return;
      }
      showStatus(data.error || 'Something went wrong generating the report.', 'error');
      return;
    }

    const warningsHeader = res.headers.get('X-Report-Warnings');
    const warnings = warningsHeader ? JSON.parse(decodeURIComponent(warningsHeader)) : [];

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `WebrocketAI-Report-${currentLabel.replace(' ', '-')}-vs-${comparisonLabel.replace(' ', '-')}.pptx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    showWarnings(warnings);
    showStatus(`Report for ${currentLabel} vs. ${comparisonLabel} generated and downloaded successfully.`, 'success');
  } catch (e) {
    showStatus('Network error while generating the report. Please try again.', 'error');
  } finally {
    updateGenerateButtonState();
  }
}

generateBtn.addEventListener('click', generateReport);
currentMonthSelect.addEventListener('change', updateGenerateButtonState);
currentYearSelect.addEventListener('change', updateGenerateButtonState);
comparisonMonthSelect.addEventListener('change', updateGenerateButtonState);
comparisonYearSelect.addEventListener('change', updateGenerateButtonState);

clientSelect.addEventListener('change', () => {
  activeClientId = clientSelect.value;
  updateActiveClientLabel();
});
addClientBtn.addEventListener('click', () => toggleAddClientForm(true));
cancelClientBtn.addEventListener('click', () => toggleAddClientForm(false));
addClientForm.addEventListener('submit', saveClient);
deleteClientBtn.addEventListener('click', deleteActiveClient);

handleUrlParams();
loadClients();
checkAuthStatus();
