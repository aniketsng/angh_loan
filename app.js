const STORAGE_KEY = 'ledger-loan-manager-v1';
const state = { loans: [], payments: [], view: 'overview', search: '', filter: 'all', selectedLoanId: null, excelHandle: null, excelStatus: 'not-connected', excelHandleRemembered: true };
const currency = new Intl.NumberFormat('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 });
const shortDate = new Intl.DateTimeFormat('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
const viewRoot = document.querySelector('#view-root');
const loanDialog = document.querySelector('#loan-dialog');
const paymentDialog = document.querySelector('#payment-dialog');
const excelDialog = document.querySelector('#excel-dialog');
let toastTimer;
let installPrompt = null;

function localDateValue(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]);
}

function money(value) {
  return currency.format(Math.round(Number(value) || 0));
}

function dateLabel(value) {
  if (!value) return 'No date';
  const date = new Date(`${value}T12:00:00`);
  return Number.isNaN(date.getTime()) ? 'No date' : shortDate.format(date);
}

function monthDate(startDate, monthsAhead) {
  const [year, month, day] = startDate.split('-').map(Number);
  const targetMonth = month - 1 + monthsAhead;
  const targetYear = year + Math.floor(targetMonth / 12);
  const normalizedMonth = ((targetMonth % 12) + 12) % 12;
  const lastDay = new Date(targetYear, normalizedMonth + 1, 0).getDate();
  return `${targetYear}-${String(normalizedMonth + 1).padStart(2, '0')}-${String(Math.min(day, lastDay)).padStart(2, '0')}`;
}

function calculateEmi(principal, annualRate, termMonths) {
  const monthlyRate = annualRate / 1200;
  if (monthlyRate === 0) return principal / termMonths;
  const factor = (1 + monthlyRate) ** termMonths;
  return principal * monthlyRate * factor / (factor - 1);
}

function scheduleFor(loan) {
  const monthlyRate = loan.rate / 1200;
  const emi = loan.emi;
  let balance = loan.principal;
  return Array.from({ length: loan.term }, (_, index) => {
    const interest = balance * monthlyRate;
    const principalPaid = Math.min(balance, emi - interest);
    const installment = {
      number: index + 1,
      dueDate: monthDate(loan.startDate, index + 1),
      amount: index === loan.term - 1 ? balance + interest : emi,
      interest,
      principal: index === loan.term - 1 ? balance : principalPaid,
      balance: index === loan.term - 1 ? 0 : Math.max(0, balance - principalPaid)
    };
    balance = installment.balance;
    return installment;
  });
}

function paymentsFor(loan) {
  return state.payments.filter((payment) => payment.loanId === loan.id);
}

function paidCount(loan) {
  return new Set(paymentsFor(loan).map((payment) => payment.installment)).size;
}

function nextInstallment(loan) {
  const paid = new Set(paymentsFor(loan).map((payment) => payment.installment));
  return scheduleFor(loan).find((installment) => !paid.has(installment.number)) || null;
}

function loanStats(loan) {
  const schedule = scheduleFor(loan);
  const payments = paymentsFor(loan);
  const paidNumbers = new Set(payments.map((payment) => payment.installment));
  const paidInstallments = schedule.filter((installment) => paidNumbers.has(installment.number));
  const paidPrincipal = paidInstallments.reduce((sum, installment) => sum + installment.principal, 0);
  const collected = payments.reduce((sum, payment) => sum + payment.amount, 0);
  const outstanding = Math.max(0, loan.principal - paidPrincipal);
  const next = nextInstallment(loan);
  return { schedule, payments, paidNumbers, paidPrincipal, collected, outstanding, next };
}

function workbookRecords(loans, payments) {
  const loanRows = loans.map((loan) => {
    const loanPayments = payments.filter((payment) => payment.loanId === loan.id);
    const paid = new Set(loanPayments.map((payment) => payment.installment));
    const schedule = scheduleFor(loan);
    const paidPrincipal = schedule
      .filter((installment) => paid.has(installment.number))
      .reduce((sum, installment) => sum + installment.principal, 0);
    const next = schedule.find((installment) => !paid.has(installment.number));
    const status = !next ? 'Paid off' : next.dueDate < localDateValue() ? 'Overdue' : 'Active';
    return [
      loan.id, loan.borrower, loan.type, loan.principal, loan.rate, loan.term, loan.emi,
      loan.startDate, paid.size, Math.max(0, loan.principal - paidPrincipal), status, loan.notes || ''
    ];
  });
  const paymentRows = payments.map((payment) => [
    payment.id,
    payment.loanId,
    loans.find((loan) => loan.id === payment.loanId)?.borrower || '',
    payment.installment,
    payment.amount,
    payment.paidDate,
    payment.note || ''
  ]);
  return { loans: loanRows, payments: paymentRows };
}

function updateExcelStatus() {
  const status = document.querySelector('#excel-status');
  const button = document.querySelector('#excel-connect');
  const labels = {
    'not-connected': ['Excel not connected', 'Connect Excel'],
    linked: [state.excelHandleRemembered ? 'Excel backup linked' : 'Excel linked for this session', 'Change workbook'],
    syncing: ['Updating Excel...', 'Change workbook'],
    'needs-permission': ['Excel access needed', 'Reconnect Excel'],
    error: ['Excel backup needs attention', 'Reconnect Excel'],
    unsupported: ['Excel sync unavailable', 'Excel unavailable']
  };
  const [label, buttonLabel] = labels[state.excelStatus];
  status.textContent = label;
  status.className = `excel-status ${state.excelStatus === 'linked' ? 'linked' : state.excelStatus === 'error' || state.excelStatus === 'needs-permission' ? 'error' : state.excelStatus}`;
  button.textContent = buttonLabel;
  button.disabled = state.excelStatus === 'unsupported';
}

async function syncExcel(loans, payments, requestPermission = false) {
  if (!state.excelHandle) return true;
  state.excelStatus = 'syncing';
  updateExcelStatus();
  try {
    await window.excelStore.sync(state.excelHandle, workbookRecords(loans, payments), requestPermission);
    state.excelStatus = 'linked';
    updateExcelStatus();
    return true;
  } catch (error) {
    console.error('Could not update the Excel backup.', error);
    state.excelStatus = error.message.includes('Reconnect') ? 'needs-permission' : 'error';
    updateExcelStatus();
    return false;
  }
}

async function save(loans = state.loans, payments = state.payments) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ loans, payments }));
  } catch (error) {
    console.error('Could not save loan manager data.', error);
    showToast('Could not save changes in this browser. Check available storage and try again.');
    return false;
  }
  await syncExcel(loans, payments);
  return true;
}

function showSaveToast(message) {
  if (state.excelHandle && state.excelStatus !== 'linked') {
    showToast(`${message} Excel backup needs attention; this device copy is saved.`);
  } else if (!state.excelHandle) {
    showToast(`${message} Link Excel once to enable automatic workbook backups.`);
  } else {
    showToast(message);
  }
}

async function connectExcel() {
  if (!window.excelStore || typeof window.showSaveFilePicker !== 'function') {
    state.excelStatus = 'unsupported';
    updateExcelStatus();
    showToast('Automatic Excel backups require a current Chromium-based browser.');
    return;
  }
  let handle;
  try {
    handle = await window.showSaveFilePicker({
      suggestedName: 'Ledger Loan Manager.xlsx',
      types: [{ description: 'Excel workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
    });
  } catch (error) {
    if (error.name === 'AbortError') return;
    console.error('Could not open the Excel workbook picker.', error);
    showToast('Could not choose an Excel workbook. Try again.');
    return;
  }

  let workbookBackup;
  try {
    workbookBackup = await window.excelStore.readWorkbook(handle);
  } catch (error) {
    console.error('Could not read the selected Excel workbook.', error);
    showToast('Could not read this workbook. Nothing was changed; choose it again or select another file.');
    return;
  }

  const shouldRestoreBackup = !state.loans.length && !state.payments.length && workbookBackup
    && (workbookBackup.loans.length || workbookBackup.payments.length);
  state.excelHandle = handle;
  state.excelHandleRemembered = true;
  try {
    await window.excelStore.saveHandle(handle);
  } catch (error) {
    console.error('Could not remember the Excel workbook link.', error);
    state.excelHandleRemembered = false;
  }
  if (shouldRestoreBackup) {
    state.loans = workbookBackup.loans;
    state.payments = workbookBackup.payments;
    render();
    if (!await save()) return;
    showSaveToast('Previous loans and repayments restored from Excel.');
    return;
  }
  const synced = await syncExcel(state.loans, state.payments, true);
  if (synced) {
    showToast(state.excelHandleRemembered
      ? 'Excel backup connected. New records will sync automatically.'
      : 'Excel backup connected for this session. The browser could not remember the link.');
  } else {
    showToast('The workbook was selected, but the initial backup failed. Reconnect and try again.');
  }
}

async function restoreExcel() {
  if (!state.excelHandle) {
    excelDialog.showModal();
    return;
  }
  try {
    const permission = await state.excelHandle.requestPermission({ mode: 'readwrite' });
    if (permission !== 'granted') {
      state.excelStatus = 'needs-permission';
      updateExcelStatus();
      showToast('Excel access was not granted. Tap Reconnect Excel to try again.');
      return;
    }
    if (await syncExcel(state.loans, state.payments)) {
      showToast('Excel backup reconnected and updated.');
    } else {
      showToast('Excel access returned, but the workbook could not be updated. Choose another workbook.');
    }
  } catch (error) {
    console.error('Could not reconnect the Excel workbook.', error);
    state.excelStatus = 'error';
    updateExcelStatus();
    showToast('Could not reconnect Excel. Choose the workbook again to restore automatic backups.');
  }
}

async function initializeExcel() {
  if (!window.excelStore || typeof window.showSaveFilePicker !== 'function') {
    state.excelStatus = 'unsupported';
    updateExcelStatus();
    return;
  }
  try {
    state.excelHandle = await window.excelStore.loadHandle();
    if (!state.excelHandle) {
      updateExcelStatus();
      return;
    }
    const permission = await state.excelHandle.queryPermission({ mode: 'readwrite' });
    if (permission !== 'granted') {
      state.excelStatus = 'needs-permission';
      updateExcelStatus();
      return;
    }
    await syncExcel(state.loans, state.payments);
  } catch (error) {
    console.error('Could not restore the Excel workbook link.', error);
    state.excelStatus = 'error';
    updateExcelStatus();
    showToast('Could not restore the Excel backup link. Reconnect your workbook to resume syncing.');
  }
}

async function registerProgressiveApp() {
  if (location.protocol === 'file:' || !('serviceWorker' in navigator)) return;
  try {
    await navigator.serviceWorker.register('./service-worker.js');
  } catch (error) {
    console.error('Could not enable offline app support.', error);
    showToast('Offline support could not start. Open the app from localhost or a secure website.');
    return;
  }
  try {
    if (navigator.storage?.persist) await navigator.storage.persist();
  } catch (error) {
    console.error('Could not request persistent browser storage.', error);
  }
}

function load() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    if (saved && Array.isArray(saved.loans) && Array.isArray(saved.payments)) {
      state.loans = saved.loans;
      state.payments = saved.payments;
    }
  } catch (error) {
    showToast('Saved data could not be read. You can still add new loans.');
  }
}

function showToast(message) {
  const toast = document.querySelector('#toast');
  toast.textContent = message;
  toast.classList.add('visible');
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toast.classList.remove('visible'), 2800);
}

function getLoan(id) {
  return state.loans.find((loan) => loan.id === id);
}

function statusFor(loan) {
  const installment = nextInstallment(loan);
  if (!installment) return { label: 'Paid off', className: 'paid' };
  if (installment.dueDate < localDateValue()) return { label: 'Overdue', className: 'overdue' };
  return { label: 'Active', className: 'active' };
}

function avatar(name, type) {
  const initials = String(name || '?').trim().split(/\s+/).slice(0, 2).map((part) => part[0] || '').join('').toUpperCase();
  return `<span class="avatar ${type === 'Business' ? 'orange' : ''}">${escapeHtml(initials || '?')}</span>`;
}

function pageHeading(eyebrow, title, description, buttonLabel = 'New loan', action = 'add-loan') {
  return `<div class="page-heading"><div><p class="eyebrow">${eyebrow}</p><h1>${title}</h1><p class="page-description">${description}</p></div><button class="button button-primary" type="button" data-action="${action}"><span class="button-plus">+</span>${buttonLabel}</button></div>`;
}

function emptyState(title, message, actionLabel = 'Add your first loan', action = 'add-loan') {
  return `<div class="empty-state"><span class="empty-icon">▤</span><h3>${title}</h3><p>${message}</p><button class="button button-primary button-small" type="button" data-action="${action}">${actionLabel}</button></div>`;
}

function loanPreview(loan) {
  const stats = loanStats(loan);
  const status = statusFor(loan);
  return `<div class="loan-preview-row">
    <div class="borrower-cell">${avatar(loan.borrower, loan.type)}<div class="loan-detail"><div class="borrower-name">${escapeHtml(loan.borrower)}</div><div class="borrower-meta">${escapeHtml(loan.type)} loan · ${paidCount(loan)}/${loan.term} EMIs</div></div></div>
    <div class="amount-cell"><div class="amount-primary">${money(stats.outstanding)}</div><div class="amount-secondary">outstanding</div></div>
    <span class="status ${status.className}">${status.label}</span>
  </div>`;
}

function duePreview(loan) {
  const installment = nextInstallment(loan);
  if (!installment) return '';
  const overdue = installment.dueDate < localDateValue();
  return `<div class="due-row"><div class="due-info"><div class="due-name">${escapeHtml(loan.borrower)}</div><div class="due-meta">${overdue ? 'Overdue' : 'Due'} ${dateLabel(installment.dueDate)} · EMI ${installment.number}</div><div class="due-amount">${money(installment.amount)}</div></div><button class="small-icon-button" type="button" data-action="record-payment" data-id="${escapeHtml(loan.id)}" aria-label="Record payment for ${escapeHtml(loan.borrower)}" title="Record payment">↗</button></div>`;
}

function renderOverview() {
  const today = localDateValue();
  const allStats = state.loans.map((loan) => loanStats(loan));
  const principal = state.loans.reduce((sum, loan) => sum + loan.principal, 0);
  const outstanding = allStats.reduce((sum, stats) => sum + stats.outstanding, 0);
  const collected = allStats.reduce((sum, stats) => sum + stats.collected, 0);
  const principalRepaid = Math.max(0, principal - outstanding);
  const repaymentProgress = principal ? Math.min(100, Math.round(principalRepaid / principal * 100)) : 0;
  const overdue = state.loans.filter((loan) => {
    const next = nextInstallment(loan);
    return next && next.dueDate < today;
  });
  const dueThisMonth = state.loans.filter((loan) => {
    const next = nextInstallment(loan);
    return next && next.dueDate.slice(0, 7) === today.slice(0, 7);
  });
  const recentLoans = [...state.loans].sort((a, b) => b.createdAt.localeCompare(a.createdAt)).slice(0, 5);
  const upcoming = [...state.loans].filter((loan) => {
    const next = nextInstallment(loan);
    return next && next.dueDate <= monthDate(today, 1);
  }).sort((a, b) => nextInstallment(a).dueDate.localeCompare(nextInstallment(b).dueDate)).slice(0, 5);

  viewRoot.innerHTML = `${pageHeading('YOUR FINANCES', 'Loan overview', 'A clearer view of your lending portfolio.')}
    <section class="overview-hero">
      <div class="hero-copy"><p class="hero-eyebrow">TOTAL OUTSTANDING</p><div class="hero-amount">${money(outstanding)}</div><p class="hero-caption">Principal remaining across <strong>${state.loans.length} loan${state.loans.length === 1 ? '' : 's'}</strong></p><div class="hero-issued"><span>Principal issued</span><strong>${money(principal)}</strong></div></div>
      <div class="hero-progress"><div class="hero-ring" style="--progress:${repaymentProgress}%"><div class="hero-ring-center"><strong>${repaymentProgress}%</strong><span>repaid</span></div></div><p>${money(principalRepaid)} paid toward principal</p></div>
    </section>
    <div class="metric-grid overview-metrics">
      <article class="metric-card"><div class="metric-top"><span>Collected to date</span><span class="metric-icon">↙</span></div><div class="metric-value">${money(collected)}</div><div class="metric-sub">${state.payments.length} repayment${state.payments.length === 1 ? '' : 's'} recorded</div></article>
      <article class="metric-card"><div class="metric-top"><span>EMIs due this month</span><span class="metric-icon orange">◷</span></div><div class="metric-value">${dueThisMonth.length}</div><div class="metric-sub">Installments in the current month</div></article>
      <article class="metric-card"><div class="metric-top"><span>Overdue loans</span><span class="metric-icon ${overdue.length ? 'red' : ''}">!</span></div><div class="metric-value">${overdue.length}</div><div class="metric-sub">${overdue.length ? 'Needs follow-up' : 'All payments are on track'}</div></article>
    </div>
    <div class="dashboard-grid">
      <section class="panel"><div class="panel-heading"><div><h2>Recent loans</h2><p>Your latest loan records</p></div><button class="text-button" type="button" data-action="view-loans">View all ↗</button></div>
        ${recentLoans.length ? `<div class="panel-body">${recentLoans.map(loanPreview).join('')}</div>` : emptyState('A fresh start', 'Add a loan to see balances and repayment progress here.')}
      </section>
      <section class="panel"><div class="panel-heading"><div><h2>Upcoming repayments</h2><p>Next EMI due for each loan</p></div><button class="text-button" type="button" data-action="view-payments">All repayments ↗</button></div>
        ${upcoming.length ? `<div class="panel-body">${upcoming.map(duePreview).join('')}</div>` : `<div class="panel-body">${emptyState('Nothing due yet', state.loans.length ? 'All loans are paid off.' : 'Upcoming EMIs will show here once you add a loan.', state.loans.length ? '' : 'Add a loan')}</div>`}
      </section>
    </div>`;
}

function filteredLoans() {
  const query = state.search.trim().toLowerCase();
  return [...state.loans].filter((loan) => {
    const matchesQuery = !query || [loan.borrower, loan.type, loan.notes].some((value) => String(value || '').toLowerCase().includes(query));
    const status = statusFor(loan);
    const matchesFilter = state.filter === 'all'
      || (state.filter === 'active' && status.label !== 'Paid off')
      || (state.filter === 'paid' && status.label === 'Paid off')
      || (state.filter === 'overdue' && status.label === 'Overdue');
    return matchesQuery && matchesFilter;
  }).sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

function renderLoanTable(loans) {
  if (!loans.length) {
    return `<div class="panel">${emptyState(state.loans.length ? 'No matching loans' : 'No loans added yet', state.loans.length ? 'Try another search or status filter.' : 'Create a loan to start tracking its EMI schedule.')}</div>`;
  }
  return `<div class="panel table-wrap"><table class="data-table"><thead><tr><th>Borrower</th><th>Principal</th><th>Outstanding</th><th>EMI / progress</th><th>Next due</th><th>Status</th><th></th></tr></thead><tbody>${loans.map((loan) => {
    const stats = loanStats(loan);
    const next = stats.next;
    const status = statusFor(loan);
    return `<tr>
      <td><div class="borrower-cell">${avatar(loan.borrower, loan.type)}<div><div class="table-primary">${escapeHtml(loan.borrower)}</div><div class="table-secondary">${escapeHtml(loan.type)} loan</div></div></div></td>
      <td><div class="table-primary">${money(loan.principal)}</div><div class="table-secondary">${loan.rate}% p.a.</div></td>
      <td><div class="table-primary">${money(stats.outstanding)}</div><div class="table-secondary">${money(stats.collected)} collected</div></td>
      <td><div class="table-primary">${money(loan.emi)}</div><div class="table-secondary">${paidCount(loan)} of ${loan.term} paid</div></td>
      <td>${next ? `<div class="table-primary">${dateLabel(next.dueDate)}</div><div class="table-secondary">EMI ${next.number}</div>` : '<span class="table-secondary">Complete</span>'}</td>
      <td><span class="status ${status.className}">${status.label}</span></td>
      <td><div class="table-actions"><button class="text-button" type="button" data-action="open-loan" data-id="${escapeHtml(loan.id)}">Details</button></div></td>
    </tr>`;
  }).join('')}</tbody></table></div>`;
}

function renderLoans() {
  const loans = filteredLoans();
  viewRoot.innerHTML = `${pageHeading('PORTFOLIO', 'All loans', 'Manage borrowers, EMI progress, and outstanding balances.')}
    <div class="toolbar"><input id="loan-search" class="search-field" type="search" value="${escapeHtml(state.search)}" placeholder="Search borrower or loan type..." aria-label="Search loans" /><select id="loan-filter" class="filter-select" aria-label="Filter by status"><option value="all" ${state.filter === 'all' ? 'selected' : ''}>All statuses</option><option value="active" ${state.filter === 'active' ? 'selected' : ''}>Active</option><option value="overdue" ${state.filter === 'overdue' ? 'selected' : ''}>Overdue</option><option value="paid" ${state.filter === 'paid' ? 'selected' : ''}>Paid off</option></select></div>
    ${renderLoanTable(loans)}`;
}

function renderLoanDetail(loan) {
  if (!loan) {
    state.view = 'loans';
    render();
    return;
  }
  const stats = loanStats(loan);
  const status = statusFor(loan);
  const progress = Math.min(100, paidCount(loan) / loan.term * 100);
  const schedule = stats.schedule;
  viewRoot.innerHTML = `<div class="page-heading"><div><button class="text-button" type="button" data-action="view-loans">← All loans</button><p class="eyebrow" style="margin-top:12px">${escapeHtml(loan.type.toUpperCase())} LOAN</p><h1>${escapeHtml(loan.borrower)}</h1><p class="page-description">Started ${dateLabel(loan.startDate)} · <span class="status ${status.className}">${status.label}</span></p></div><div class="table-actions"><button class="button button-quiet" type="button" data-action="edit-loan" data-id="${escapeHtml(loan.id)}">Edit details</button>${stats.next ? `<button class="button button-primary" type="button" data-action="record-payment" data-id="${escapeHtml(loan.id)}">Record EMI <span>↗</span></button>` : ''}<button class="button button-danger" type="button" data-action="delete-loan" data-id="${escapeHtml(loan.id)}">Delete loan</button></div></div>
    <div class="detail-view">
      <section class="panel detail-card"><p class="eyebrow">LOAN SUMMARY</p><div class="detail-label">Outstanding principal</div><div class="detail-amount">${money(stats.outstanding)}</div><div class="detail-label">of ${money(loan.principal)} original principal</div>
        <div class="progress-track"><div class="progress-fill" style="width:${progress}%"></div></div><div class="progress-legend"><span>${paidCount(loan)} of ${loan.term} EMIs paid</span><span>${Math.round(progress)}%</span></div>
        <div class="detail-stats"><div class="detail-stat"><span>Monthly EMI</span><strong>${money(loan.emi)}</strong></div><div class="detail-stat"><span>Interest rate</span><strong>${loan.rate}% p.a.</strong></div><div class="detail-stat"><span>Amount collected</span><strong>${money(stats.collected)}</strong></div><div class="detail-stat"><span>Loan term</span><strong>${loan.term} months</strong></div><div class="detail-stat"><span>Disbursed on</span><strong>${dateLabel(loan.startDate)}</strong></div><div class="detail-stat"><span>Est. total interest</span><strong>${money(loan.emi * loan.term - loan.principal)}</strong></div></div>
        ${loan.notes ? `<div class="detail-stat" style="margin-top:20px"><span>Notes</span><strong>${escapeHtml(loan.notes)}</strong></div>` : ''}
      </section>
      <section class="panel"><div class="panel-heading"><div><h2>Repayment schedule</h2><p>Reducing-balance EMI breakdown</p></div></div><div class="schedule-list">${schedule.map((installment) => {
        const payment = stats.payments.find((entry) => entry.installment === installment.number);
        const overdue = !payment && installment.dueDate < localDateValue();
        const installmentStatus = payment ? 'paid' : overdue ? 'overdue' : 'upcoming';
        const label = payment ? `Paid ${dateLabel(payment.paidDate)}` : overdue ? 'Overdue' : `Due ${dateLabel(installment.dueDate)}`;
        return `<div class="schedule-row"><span class="schedule-number">${installment.number}</span><div><div class="schedule-main">${money(installment.amount)}</div><div class="schedule-sub">Principal ${money(installment.principal)} · Interest ${money(installment.interest)}</div></div><span class="status ${installmentStatus}">${label}</span></div>`;
      }).join('')}</div></section>
    </div>`;
}

function renderPayments() {
  const payments = [...state.payments].sort((a, b) => b.paidDate.localeCompare(a.paidDate));
  const collected = payments.reduce((sum, payment) => sum + payment.amount, 0);
  viewRoot.innerHTML = `${pageHeading('CASH FLOW', 'Repayments', 'A complete record of EMI payments received.', 'Record repayment', 'add-payment')}
    <div class="metric-grid">
      <article class="metric-card"><div class="metric-top"><span>Collected so far</span><span class="metric-icon">↙</span></div><div class="metric-value">${money(collected)}</div><div class="metric-sub">${payments.length} recorded payment${payments.length === 1 ? '' : 's'}</div></article>
      <article class="metric-card"><div class="metric-top"><span>Outstanding principal</span><span class="metric-icon orange">◷</span></div><div class="metric-value">${money(state.loans.reduce((sum, loan) => sum + loanStats(loan).outstanding, 0))}</div><div class="metric-sub">Across ${state.loans.length} loan${state.loans.length === 1 ? '' : 's'}</div></article>
      <article class="metric-card"><div class="metric-top"><span>EMIs remaining</span><span class="metric-icon">▤</span></div><div class="metric-value">${state.loans.reduce((sum, loan) => sum + loan.term - paidCount(loan), 0)}</div><div class="metric-sub">Scheduled installments</div></article>
      <article class="metric-card"><div class="metric-top"><span>Overdue loans</span><span class="metric-icon ${state.loans.some((loan) => statusFor(loan).label === 'Overdue') ? 'red' : ''}">!</span></div><div class="metric-value">${state.loans.filter((loan) => statusFor(loan).label === 'Overdue').length}</div><div class="metric-sub">Next installment is past due</div></article>
    </div>
    <section class="panel"><div class="panel-heading"><div><h2>Payment history</h2><p>Most recent repayments first</p></div></div>
      ${payments.length ? `<div>${payments.map((payment) => {
        const loan = getLoan(payment.loanId);
        return `<div class="payment-list-row">${avatar(loan?.borrower || 'Loan', loan?.type)}<div><div class="table-primary">${escapeHtml(loan?.borrower || 'Loan removed')}</div><div class="table-secondary">EMI ${payment.installment}${payment.note ? ` · ${escapeHtml(payment.note)}` : ''}</div></div><span class="payment-date">${dateLabel(payment.paidDate)}</span><div class="amount-cell"><div class="amount-primary">${money(payment.amount)}</div><span class="status paid">Received</span></div></div>`;
      }).join('')}</div>` : emptyState('No repayments recorded', 'Once you record an EMI, your payment history will appear here.', state.loans.length ? 'Record a repayment' : 'Add your first loan', state.loans.length ? 'add-payment' : 'add-loan')}
    </section>`;
}

function render() {
  const sectionName = state.view === 'overview' ? 'Overview' : state.view === 'loans' ? (state.selectedLoanId ? 'Loan details' : 'All loans') : 'Repayments';
  document.querySelector('#current-section').textContent = sectionName;
  document.querySelector('#loan-count').textContent = state.loans.length || '';
  document.querySelectorAll('.nav-item').forEach((item) => item.classList.toggle('active', item.dataset.view === (state.selectedLoanId ? 'loans' : state.view)));
  if (state.view === 'overview') renderOverview();
  if (state.view === 'loans') state.selectedLoanId ? renderLoanDetail(getLoan(state.selectedLoanId)) : renderLoans();
  if (state.view === 'payments') renderPayments();
}

function setView(view) {
  state.view = view;
  state.selectedLoanId = null;
  if (view !== 'loans') state.search = '';
  render();
}

function openLoanDialog(loan = null) {
  const form = document.querySelector('#loan-form');
  form.reset();
  document.querySelector('#loan-id').value = loan?.id || '';
  document.querySelector('#loan-dialog-eyebrow').textContent = loan ? 'EDIT LOAN' : 'NEW LOAN';
  document.querySelector('#loan-dialog-title').textContent = loan ? 'Edit loan details' : 'Add a loan';
  document.querySelector('#loan-submit').innerHTML = loan ? 'Save changes <span>↗</span>' : 'Save loan <span>↗</span>';
  if (loan) {
    document.querySelector('#borrower').value = loan.borrower;
    document.querySelector('#loan-type').value = loan.type;
    document.querySelector('#principal').value = loan.principal;
    document.querySelector('#rate').value = loan.rate;
    document.querySelector('#term').value = loan.term;
    document.querySelector('#start-date').value = loan.startDate;
    document.querySelector('#loan-notes').value = loan.notes || '';
    document.querySelectorAll('#loan-form input, #loan-form select').forEach((field) => {
      if (!['borrower', 'loan-type', 'loan-notes'].includes(field.id)) field.disabled = true;
    });
  } else {
    document.querySelectorAll('#loan-form input, #loan-form select').forEach((field) => { field.disabled = false; });
    document.querySelector('#start-date').value = localDateValue();
  }
  loanDialog.showModal();
}

function openPaymentDialog(loanId = '') {
  if (!state.loans.length) {
    showToast('Add a loan before recording a repayment.');
    return;
  }
  const select = document.querySelector('#payment-loan');
  const eligibleLoans = state.loans.filter((loan) => nextInstallment(loan));
  if (!eligibleLoans.length) {
    showToast('All loan installments have been paid.');
    return;
  }
  select.innerHTML = eligibleLoans.map((loan) => `<option value="${escapeHtml(loan.id)}" ${loan.id === loanId ? 'selected' : ''}>${escapeHtml(loan.borrower)} · ${escapeHtml(loan.type)}</option>`).join('');
  document.querySelector('#payment-date').value = localDateValue();
  document.querySelector('#payment-note').value = '';
  updatePaymentPreview();
  paymentDialog.showModal();
}

function updatePaymentPreview() {
  const loan = getLoan(document.querySelector('#payment-loan').value);
  const next = loan && nextInstallment(loan);
  document.querySelector('#payment-preview').innerHTML = next
    ? `<span>EMI ${next.number} · Due ${dateLabel(next.dueDate)}</span><strong>${money(next.amount)}</strong>`
    : '<span>No unpaid installments</span>';
}

async function deleteLoan(id) {
  const loan = getLoan(id);
  if (!loan || !window.confirm(`Delete the loan for ${loan.borrower} and its repayment history? This cannot be undone.`)) return;
  const loans = state.loans.filter((entry) => entry.id !== id);
  const payments = state.payments.filter((payment) => payment.loanId !== id);
  if (!await save(loans, payments)) return;
  state.loans = loans;
  state.payments = payments;
  state.selectedLoanId = null;
  render();
  showSaveToast('Loan and its repayment history deleted.');
}

document.addEventListener('click', async (event) => {
  const button = event.target.closest('button');
  if (!button) return;
  if (button.dataset.view) {
    setView(button.dataset.view);
    return;
  }
  const { action, id } = button.dataset;
  if (action === 'add-loan') openLoanDialog();
  if (action === 'add-payment') openPaymentDialog();
  if (action === 'record-payment') openPaymentDialog(id);
  if (action === 'edit-loan') openLoanDialog(getLoan(id));
  if (action === 'view-loans') setView('loans');
  if (action === 'view-payments') setView('payments');
  if (action === 'open-loan') {
    state.view = 'loans';
    state.selectedLoanId = id;
    render();
  }
  if (action === 'delete-loan') await deleteLoan(id);
  if (button.dataset.close) document.querySelector(`#${button.dataset.close}`).close();
});

document.querySelector('#excel-connect').addEventListener('click', () => {
  if (state.excelHandle && state.excelStatus === 'needs-permission') {
    restoreExcel();
    return;
  }
  excelDialog.showModal();
});
document.querySelector('#excel-select-file').addEventListener('click', async () => {
  excelDialog.close();
  await connectExcel();
});

viewRoot.addEventListener('input', (event) => {
  if (event.target.id === 'loan-search') {
    state.search = event.target.value;
    const position = event.target.selectionStart;
    renderLoans();
    const search = document.querySelector('#loan-search');
    search.focus();
    search.setSelectionRange(position, position);
  }
});

viewRoot.addEventListener('change', (event) => {
  if (event.target.id === 'loan-filter') {
    state.filter = event.target.value;
    renderLoans();
  }
});

document.querySelector('#loan-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const id = document.querySelector('#loan-id').value;
  const borrower = document.querySelector('#borrower').value.trim();
  const type = document.querySelector('#loan-type').value;
  const notes = document.querySelector('#loan-notes').value.trim();
  if (!borrower) {
    showToast('Enter the borrower name.');
    document.querySelector('#borrower').focus();
    return;
  }
  const existing = id ? getLoan(id) : null;
  if (existing) {
    const loans = state.loans.map((loan) => loan.id === id ? { ...loan, borrower, type, notes } : loan);
    if (!await save(loans, state.payments)) return;
    state.loans = loans;
  } else {
    const principal = Number(document.querySelector('#principal').value);
    const rate = Number(document.querySelector('#rate').value);
    const term = Number(document.querySelector('#term').value);
    const startDate = document.querySelector('#start-date').value;
    if (!Number.isFinite(principal) || principal <= 0 || !Number.isFinite(rate) || rate < 0 || rate > 100 || !Number.isInteger(term) || term < 1 || term > 600 || !startDate) {
      showToast('Enter a valid principal, interest rate, term, and start date.');
      return;
    }
    const loan = {
      id: crypto.randomUUID(),
      borrower,
      type,
      principal,
      rate,
      term,
      startDate,
      notes,
      emi: calculateEmi(principal, rate, term),
      createdAt: new Date().toISOString()
    };
    const loans = [...state.loans, loan];
    if (!await save(loans, state.payments)) return;
    state.loans = loans;
  }
  loanDialog.close();
  render();
  showSaveToast(existing ? 'Loan details updated.' : 'Loan added successfully.');
});

document.querySelector('#payment-loan').addEventListener('change', updatePaymentPreview);
document.querySelector('#payment-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const loan = getLoan(document.querySelector('#payment-loan').value);
  const installment = loan && nextInstallment(loan);
  if (!loan || !installment) {
    showToast('No unpaid installment is available for this loan.');
    return;
  }
  const paidDate = document.querySelector('#payment-date').value;
  if (!paidDate) {
    showToast('Choose a payment date.');
    return;
  }
  const payment = {
    id: crypto.randomUUID(),
    loanId: loan.id,
    installment: installment.number,
    amount: installment.amount,
    paidDate,
    note: document.querySelector('#payment-note').value.trim()
  };
  const payments = [...state.payments, payment];
  if (!await save(state.loans, payments)) return;
  state.payments = payments;
  paymentDialog.close();
  render();
  showSaveToast(`${money(installment.amount)} repayment recorded.`);
});

document.querySelectorAll('[data-close]').forEach((button) => {
  button.addEventListener('click', () => document.querySelector(`#${button.dataset.close}`).close());
});

document.querySelectorAll('.dialog').forEach((dialog) => {
  dialog.addEventListener('click', (event) => {
    if (event.target === dialog) dialog.close();
  });
});

window.addEventListener('beforeinstallprompt', (event) => {
  event.preventDefault();
  installPrompt = event;
  document.querySelector('#install-app').hidden = false;
});

document.querySelector('#install-app').addEventListener('click', async () => {
  if (!installPrompt) return;
  await installPrompt.prompt();
  const { outcome } = await installPrompt.userChoice;
  installPrompt = null;
  document.querySelector('#install-app').hidden = true;
  if (outcome === 'accepted') showToast('Ledger is installing on your device.');
});

window.addEventListener('appinstalled', () => {
  installPrompt = null;
  document.querySelector('#install-app').hidden = true;
  showToast('Ledger was added to your home screen.');
});

document.querySelector('#today-label').textContent = new Intl.DateTimeFormat('en-IN', { weekday: 'short', day: 'numeric', month: 'short' }).format(new Date());
load();
render();
updateExcelStatus();
initializeExcel();
registerProgressiveApp();
