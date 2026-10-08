/**
 * Cookie Sync - Universal & Outlook Session Migrator (v2.1.0)
 * Cross-Browser Extension (Chrome & Firefox)
 * By Aprajit Sarkar (https://github.com/AprajitSarkar)
 */

// Universal Browser API Bridge
const ext = typeof browser !== 'undefined' ? browser : chrome;

// Authentication Domains for Microsoft & Outlook
const MICROSOFT_AUTH_DOMAINS = [
  "live.com",
  "office.com",
  "office365.com",
  "microsoft.com",
  "microsoftonline.com",
  "outlook.com",
  "msauth.net",
  "msftauth.net",
  "passport.net",
  "windows.net"
];

// State
let currentActiveTab = null;
let currentTabDomain = null;
let currentBaseDomain = null;
let isOutlookSessionActive = false;
let detectedActiveEmail = null;
let lastExportedSession = null;
let scannedCookies = [];
let trackedEmailsList = [];

// Initialize Popup
document.addEventListener('DOMContentLoaded', async () => {
  detectBrowser();
  setupTabs();
  setupEventHandlers();
  await loadTrackedEmails();
  await checkActiveTab();
  await scanActiveCookies();
  autoCheckClipboardOnOpen();
});

// Detect Browser Environment
function detectBrowser() {
  const isFirefox = navigator.userAgent.toLowerCase().includes('firefox');
  const browserBadge = document.getElementById('browser-badge');
  if (isFirefox) {
    browserBadge.textContent = 'Firefox';
    browserBadge.style.color = '#FF7139';
    browserBadge.style.borderColor = 'rgba(255, 113, 57, 0.4)';
    browserBadge.style.background = 'rgba(255, 113, 57, 0.12)';
  } else {
    browserBadge.textContent = 'Chrome / Edge';
    browserBadge.style.color = '#38BDF8';
    browserBadge.style.borderColor = 'rgba(56, 189, 248, 0.4)';
    browserBadge.style.background = 'rgba(56, 189, 248, 0.12)';
  }
}

// Setup Navigation Tabs (4 Tabs)
function setupTabs() {
  const tabButtons = document.querySelectorAll('.nav-tab');
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.getAttribute('data-target')));
  });
}

function switchTab(targetId) {
  const tabButtons = document.querySelectorAll('.nav-tab');
  const panels = document.querySelectorAll('.tab-panel');

  tabButtons.forEach(b => {
    const isTarget = b.getAttribute('data-target') === targetId;
    b.classList.toggle('active', isTarget);
    b.setAttribute('aria-selected', isTarget ? 'true' : 'false');
  });

  panels.forEach(p => {
    p.classList.toggle('active', p.id === targetId);
  });

  if (targetId === 'panel-inspect') {
    renderCookiesList();
  } else if (targetId === 'panel-emails') {
    renderEmailsList();
  }
}

// Extract base domain e.g. "linkedin.com" from "www.linkedin.com"
function getBaseDomain(hostname) {
  if (!hostname) return '';
  const parts = hostname.split('.').filter(Boolean);
  if (parts.length <= 2) return hostname;
  const lastTwo = parts.slice(-2).join('.');
  if (['co.uk', 'com.au', 'co.nz', 'co.in', 'org.uk'].includes(lastTwo)) {
    return parts.slice(-3).join('.');
  }
  return parts.slice(-2).join('.');
}

// Check Current Active Tab & Determine Mode
async function checkActiveTab() {
  const statusCard = document.getElementById('status-card');
  const statusLabel = document.getElementById('status-label');
  const statusSub = document.getElementById('status-sub');
  const modeTag = document.getElementById('mode-tag');
  const exportBtnLabel = document.getElementById('export-btn-label');
  const openOutlookBtn = document.getElementById('open-outlook-btn');
  const statDomainsCount = document.getElementById('stat-domains-count');

  try {
    const tabs = await queryTabs({ active: true, currentWindow: true });
    if (tabs && tabs.length > 0) {
      currentActiveTab = tabs[0];
      const url = currentActiveTab.url || '';

      if (/^(chrome|edge|about|moz-extension|chrome-extension):/i.test(url)) {
        statusCard.className = 'status-card';
        statusLabel.textContent = 'System / Empty Tab';
        modeTag.textContent = 'Standby';
        statusSub.textContent = 'Switch to any website or Outlook tab to capture';
        openOutlookBtn.style.display = 'inline-block';
        statDomainsCount.textContent = 'Global';
        switchTab('panel-import');
        return;
      }

      let parsedUrl = null;
      try { parsedUrl = new URL(url); } catch (e) {}

      if (parsedUrl) {
        currentTabDomain = parsedUrl.hostname;
        currentBaseDomain = getBaseDomain(currentTabDomain);
      }

      isOutlookSessionActive = isOutlookUrl(url);

      if (isOutlookSessionActive) {
        statusCard.className = 'status-card outlook-active';
        statusLabel.textContent = 'Outlook Active';
        modeTag.textContent = 'Outlook Mode';
        modeTag.className = 'mode-tag mode-outlook';
        statusSub.textContent = currentTabDomain || 'outlook.live.com';
        exportBtnLabel.textContent = 'Copy Outlook Session';
        statDomainsCount.textContent = 'Outlook 365';
        openOutlookBtn.style.display = 'none';
        switchTab('panel-export');
        
        // Try quick extraction of active account email
        quickDetectActiveAccountEmail();
      } else {
        statusCard.className = 'status-card connected';
        const displayDomain = currentBaseDomain || currentTabDomain || 'Active Site';
        statusLabel.textContent = displayDomain;
        modeTag.textContent = 'Universal';
        modeTag.className = 'mode-tag';
        statusSub.textContent = currentTabDomain;
        exportBtnLabel.textContent = `Copy ${displayDomain} Session`;
        statDomainsCount.textContent = displayDomain;
        openOutlookBtn.style.display = 'inline-block';
        switchTab('panel-export');
      }
    }
  } catch (err) {
    statusCard.className = 'status-card';
    statusLabel.textContent = 'Ready';
    statusSub.textContent = 'Click below to manage session';
    openOutlookBtn.style.display = 'inline-block';
  }
}

function isOutlookUrl(url) {
  if (!url) return false;
  return /outlook\.(live|office|office365)\.com|login\.(live|microsoftonline)\.com|live\.com|office\.com/i.test(url);
}

function isMicrosoftDomain(domain) {
  if (!domain) return false;
  const d = domain.toLowerCase();
  return MICROSOFT_AUTH_DOMAINS.some(target => d.includes(target));
}

// Quick email detection to update UI preview
async function quickDetectActiveAccountEmail() {
  if (!currentActiveTab || !currentActiveTab.id) return;
  try {
    const res = await sendTabMessage(currentActiveTab.id, { action: "GET_STORAGE" });
    if (res && res.accountEmail) {
      updateActiveAccountEmailUI(res.accountEmail);
    }
  } catch (e) {}
}

function updateActiveAccountEmailUI(email) {
  if (!email) return;
  detectedActiveEmail = email.toLowerCase().trim();
  const card = document.getElementById('active-account-card');
  const label = document.getElementById('active-account-email');
  if (card && label) {
    label.textContent = detectedActiveEmail;
    card.style.display = 'flex';
  }
}

// Auto-check clipboard when popup opens
async function autoCheckClipboardOnOpen() {
  const textarea = document.getElementById('import-session-textarea');
  if (textarea.value.trim()) return;

  try {
    if (navigator.clipboard && navigator.clipboard.readText) {
      const text = await navigator.clipboard.readText();
      if (text && text.trim().startsWith('{') && (text.includes('CookieSync') || text.includes('OutlookSessionSync') || text.includes('cookies'))) {
        textarea.value = text.trim();
        showToast('✓ Detected copied session in clipboard!');
      }
    }
  } catch (e) {}
}

// Setup Interactive Handlers
function setupEventHandlers() {
  const importTextarea = document.getElementById('import-session-textarea');

  // Open Outlook button
  document.getElementById('open-outlook-btn').addEventListener('click', () => {
    createTab('https://outlook.live.com/mail/0/');
  });

  // Export Session
  document.getElementById('export-session-btn').addEventListener('click', handleExportSession);

  // Download JSON
  document.getElementById('download-json-btn').addEventListener('click', handleDownloadJson);

  // Toggle JSON Preview
  document.getElementById('toggle-json-preview-btn').addEventListener('click', () => {
    const box = document.getElementById('raw-preview-box');
    const isHidden = box.style.display === 'none';
    box.style.display = isHidden ? 'block' : 'none';
    document.getElementById('toggle-json-preview-btn').textContent = isHidden ? 'Hide Payload' : 'View Payload';
  });

  // Copy Preview
  document.getElementById('copy-preview-btn').addEventListener('click', () => {
    const textarea = document.getElementById('export-json-textarea');
    if (textarea.value) {
      navigator.clipboard.writeText(textarea.value).then(() => {
        showToast('JSON payload copied!');
      });
    }
  });

  // Paste & Login Button
  document.getElementById('paste-and-login-btn').addEventListener('click', handlePasteAndLogin);

  // Quick Paste Button
  const quickPasteBtn = document.getElementById('quick-paste-btn');
  if (quickPasteBtn) {
    quickPasteBtn.addEventListener('click', async () => {
      try {
        let text = '';
        if (navigator.clipboard && navigator.clipboard.readText) {
          text = await navigator.clipboard.readText();
        }
        if (!text) {
          importTextarea.focus();
          document.execCommand('paste');
          text = importTextarea.value.trim();
        }
        if (text) {
          importTextarea.value = text;
          showToast('Pasted from clipboard!');
          if (text.startsWith('{')) executeImport(text);
        } else {
          importTextarea.focus();
          showToast('Please press Ctrl+V to paste.', true);
        }
      } catch (err) {
        importTextarea.focus();
        showToast('Please press Ctrl+V to paste.', true);
      }
    });
  }

  // Auto-import on Paste into Textarea
  if (importTextarea) {
    importTextarea.addEventListener('paste', () => {
      setTimeout(() => {
        const text = importTextarea.value.trim();
        if (text.startsWith('{')) {
          showToast('Detected session! Importing...');
          executeImport(text);
        }
      }, 100);
    });
  }

  // Manual Import
  document.getElementById('manual-import-btn').addEventListener('click', () => {
    const text = importTextarea.value.trim();
    if (!text) return showToast('Please paste the session JSON first!', true);
    executeImport(text);
  });

  // File Upload
  const fileInput = document.getElementById('file-input');
  document.getElementById('upload-file-btn').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', handleFileUpload);

  // Refresh Emails
  document.getElementById('refresh-emails-btn').addEventListener('click', async () => {
    await loadTrackedEmails();
    renderEmailsList();
    showToast('Accounts list refreshed!');
  });

  // Filter Emails
  document.getElementById('emails-filter-input').addEventListener('input', (e) => {
    renderEmailsList(e.target.value);
  });

  // Clear All Emails
  document.getElementById('clear-emails-btn').addEventListener('click', handleClearAllEmails);

  // Copy Active Detected Account Email
  const copyActiveBtn = document.getElementById('btn-copy-active-email');
  if (copyActiveBtn) {
    copyActiveBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (!detectedActiveEmail) return showToast('No account email detected yet!', true);
      navigator.clipboard.writeText(detectedActiveEmail).then(() => {
        showToast(`✓ Copied active email: ${detectedActiveEmail}!`);
      });
    });
  }

  // Inspect Refresh
  document.getElementById('refresh-inspect-btn').addEventListener('click', async () => {
    await scanActiveCookies();
    renderCookiesList();
    showToast('Cookies refreshed!');
  });

  // Inspect Filter
  document.getElementById('cookie-filter-input').addEventListener('input', (e) => {
    renderCookiesList(e.target.value);
  });

  // Clear Cookies
  document.getElementById('clear-cookies-btn').addEventListener('click', handleClearActiveCookies);
}

// -------------------------------------------------------------
// EMAILS MANAGEMENT (MRU: Most Recently Used Ordering)
// -------------------------------------------------------------
async function loadTrackedEmails() {
  try {
    const res = await getStorageData(['trackedEmails']);
    trackedEmailsList = res.trackedEmails || [];
    updateEmailsBadge();
  } catch (e) {
    trackedEmailsList = [];
  }
}

async function recordTrackedEmail(email, source = "Outlook") {
  if (!email || !email.includes('@')) return;
  const cleanEmail = email.trim().toLowerCase();

  await loadTrackedEmails();

  // Find if already present
  const existingIdx = trackedEmailsList.findIndex(item => item.email.toLowerCase() === cleanEmail);
  let count = 1;

  if (existingIdx !== -1) {
    count = (trackedEmailsList[existingIdx].count || 1) + 1;
    // Remove from current position
    trackedEmailsList.splice(existingIdx, 1);
  }

  // Add to TOP (Index 0) - Most Recently Used
  trackedEmailsList.unshift({
    email: cleanEmail,
    lastUsed: Date.now(),
    count: count,
    source: source
  });

  await setStorageData({ trackedEmails: trackedEmailsList });
  updateEmailsBadge();
  renderEmailsList();
}

function updateEmailsBadge() {
  const badge = document.getElementById('emails-badge');
  const countPill = document.getElementById('emails-count-pill');
  const count = trackedEmailsList.length;

  if (badge) {
    badge.textContent = count;
    badge.style.display = count > 0 ? 'inline-block' : 'none';
  }
  if (countPill) {
    countPill.textContent = `${count} ${count === 1 ? 'Account' : 'Accounts'}`;
  }
}

function formatRelativeTime(timestamp) {
  if (!timestamp) return 'Recently';
  const now = Date.now();
  const diffSec = Math.max(0, Math.floor((now - timestamp) / 1000));

  if (diffSec < 45) return 'Just now';

  const diffMin = Math.floor(diffSec / 60);
  if (diffMin === 1) return '1 minute ago';
  if (diffMin < 60) return `${diffMin} minutes ago`;

  const diffHour = Math.floor(diffMin / 60);
  if (diffHour === 1) return '1 hour ago';
  if (diffHour < 24) return `${diffHour} hours ago`;

  const diffDay = Math.floor(diffHour / 24);
  if (diffDay === 1) return '1 day ago';
  if (diffDay < 7) return `${diffDay} days ago`;

  const diffWeeks = Math.floor(diffDay / 7);
  if (diffWeeks === 1) return '1 week ago';
  if (diffWeeks < 5) return `${diffWeeks} weeks ago`;

  const diffMonths = Math.floor(diffDay / 30);
  if (diffMonths === 1) return '1 month ago';
  if (diffMonths < 12) return `${diffMonths} months ago`;

  return new Date(timestamp).toLocaleDateString([], { month: 'short', day: 'numeric', year: 'numeric' });
}

function formatExactTime(timestamp) {
  if (!timestamp) return '';
  return new Date(timestamp).toLocaleString([], {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function renderEmailsList(filter = '') {
  const listEl = document.getElementById('emails-list');
  listEl.innerHTML = '';

  const q = filter.trim().toLowerCase();
  const filtered = trackedEmailsList.filter(item => !q || item.email.toLowerCase().includes(q));

  if (!filtered.length) {
    listEl.innerHTML = '<div class="empty-state">No matching account emails found.</div>';
    return;
  }

  filtered.forEach((item, index) => {
    const card = document.createElement('div');
    card.className = 'email-card' + (index === 0 ? ' latest-item' : '');

    // Initials for avatar
    const namePart = item.email.split('@')[0] || '';
    const initial = namePart.substring(0, 2).toUpperCase();
    const relTime = formatRelativeTime(item.lastUsed);
    const exactTime = formatExactTime(item.lastUsed);

    card.innerHTML = `
      <div class="email-card-main">
        <div class="email-avatar" title="${item.email}">${initial}</div>
        <div class="email-details">
          <div class="email-addr-row">
            <span class="email-addr" title="${item.email}">${item.email}</span>
            ${index === 0 ? '<span class="latest-pill">Latest</span>' : ''}
          </div>
          <div class="email-meta" title="Synced at: ${exactTime}">
            <span class="email-time">
              <svg viewBox="0 0 20 20" fill="currentColor" width="11" height="11" class="time-clock-icon">
                <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm1-12a1 1 0 10-2 0v4a1 1 0 00.293.707l2.828 2.829a1 1 0 101.415-1.415L11 9.586V6z" clip-rule="evenodd" />
              </svg>
              ${relTime}
            </span>
            <span class="email-count-badge">${item.count || 1}x synced</span>
            ${item.source ? `<span class="email-source-badge">${item.source}</span>` : ''}
          </div>
        </div>
      </div>
      <div class="email-actions">
        <button class="btn-action-icon btn-copy-email" title="Copy email: ${item.email}" data-email="${item.email}" aria-label="Copy Email">
          <svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14">
            <path d="M8 3a1 1 0 011-1h2a1 1 0 110 2H9a1 1 0 01-1-1z" />
            <path d="M6 3a2 2 0 00-2 2v11a2 2 0 002 2h8a2 2 0 002-2V5a2 2 0 00-2-2 3 3 0 01-3 3H9a3 3 0 01-3-3z" />
          </svg>
        </button>
        <button class="btn-action-icon btn-delete-email" title="Delete email: ${item.email}" data-email="${item.email}" aria-label="Delete Email">
          <svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14">
            <path fill-rule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 012 0v6a1 1 0 11-2 0V8zm5-1a1 1 0 00-1 1v6a1 1 0 102 0V8a1 1 0 00-1-1z" clip-rule="evenodd" />
          </svg>
        </button>
      </div>
    `;

    // Copy event with visual feedback
    const copyBtn = card.querySelector('.btn-copy-email');
    copyBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const mail = copyBtn.getAttribute('data-email');
      navigator.clipboard.writeText(mail).then(() => {
        copyBtn.classList.add('btn-copied-success');
        copyBtn.innerHTML = `
          <svg viewBox="0 0 20 20" fill="#10B981" width="14" height="14">
            <path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd" />
          </svg>
        `;
        showToast(`✓ Copied ${mail}!`);
        setTimeout(() => {
          copyBtn.classList.remove('btn-copied-success');
          copyBtn.innerHTML = `
            <svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14">
              <path d="M8 3a1 1 0 011-1h2a1 1 0 110 2H9a1 1 0 01-1-1z" />
              <path d="M6 3a2 2 0 00-2 2v11a2 2 0 002 2h8a2 2 0 002-2V5a2 2 0 00-2-2 3 3 0 01-3 3H9a3 3 0 01-3-3z" />
            </svg>
          `;
        }, 1500);
      });
    });

    // Delete event
    const delBtn = card.querySelector('.btn-delete-email');
    delBtn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const mail = delBtn.getAttribute('data-email');
      await deleteTrackedEmail(mail);
    });

    listEl.appendChild(card);
  });
}

async function deleteTrackedEmail(email) {
  trackedEmailsList = trackedEmailsList.filter(item => item.email.toLowerCase() !== email.toLowerCase());
  await setStorageData({ trackedEmails: trackedEmailsList });
  updateEmailsBadge();
  renderEmailsList();
  showToast(`Removed ${email}`);
}

async function handleClearAllEmails() {
  if (!trackedEmailsList.length) return;
  if (!confirm('Are you sure you want to clear all tracked account emails?')) return;
  trackedEmailsList = [];
  await setStorageData({ trackedEmails: [] });
  updateEmailsBadge();
  renderEmailsList();
  showToast('Email history cleared!');
}

// -------------------------------------------------------------
// COOKIE SCANNER (Outlook vs. Universal Active Website)
// -------------------------------------------------------------
async function scanActiveCookies() {
  try {
    let matched = [];
    const seen = new Set();

    if (isOutlookSessionActive) {
      const allBrowserCookies = await getAllCookies({});
      if (allBrowserCookies && allBrowserCookies.length) {
        for (const c of allBrowserCookies) {
          if (isMicrosoftDomain(c.domain)) {
            const key = `${c.domain}|${c.name}|${c.path}`;
            if (!seen.has(key)) {
              seen.add(key);
              matched.push(c);
            }
          }
        }
      }

      // Check cookies for authoritative Microsoft email (DefaultAnchorMailbox / SignInName)
      const anchorCookie = matched.find(c => c.name === 'DefaultAnchorMailbox' || c.name === 'SignInName');
      if (anchorCookie && anchorCookie.value && anchorCookie.value.includes('@')) {
        const val = decodeURIComponent(anchorCookie.value);
        const m = val.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        if (m) {
          updateActiveAccountEmailUI(m[0]);
        }
      }
    } else if (currentActiveTab && currentActiveTab.url) {
      const urlCookies = await getAllCookies({ url: currentActiveTab.url });
      for (const c of (urlCookies || [])) {
        const key = `${c.domain}|${c.name}|${c.path}`;
        if (!seen.has(key)) {
          seen.add(key);
          matched.push(c);
        }
      }

      if (currentBaseDomain) {
        const baseCookies = await getAllCookies({ domain: currentBaseDomain });
        for (const c of (baseCookies || [])) {
          const key = `${c.domain}|${c.name}|${c.path}`;
          if (!seen.has(key)) {
            seen.add(key);
            matched.push(c);
          }
        }
      }
    }

    scannedCookies = matched;
    document.getElementById('stat-cookies-count').textContent = scannedCookies.length;
    document.getElementById('inspect-summary').textContent = `${scannedCookies.length} cookies found for ${isOutlookSessionActive ? 'Outlook' : (currentBaseDomain || 'current site')}`;
  } catch (err) {
    console.error("Cookie scan failed:", err);
  }
}

// -------------------------------------------------------------
// EXPORT SESSION LOGIC (Universal & Outlook)
// -------------------------------------------------------------
async function handleExportSession() {
  const exportBtn = document.getElementById('export-session-btn');
  const btnLabel = document.getElementById('export-btn-label');
  const originalLabel = btnLabel.textContent;

  exportBtn.disabled = true;
  btnLabel.textContent = 'Exporting Session...';

  try {
    await scanActiveCookies();

    let storageData = { localStorage: {}, sessionStorage: {} };
    let sourceUrl = currentActiveTab ? currentActiveTab.url : 'https://outlook.live.com/mail/0/';
    let accountEmail = detectedActiveEmail;

    if (currentActiveTab && currentActiveTab.id) {
      // Method A: Content script messaging
      let gotStorage = false;
      try {
        const response = await sendTabMessage(currentActiveTab.id, { action: "GET_STORAGE" });
        if (response) {
          if (response.localStorage) {
            storageData.localStorage = response.localStorage;
            storageData.sessionStorage = response.sessionStorage || {};
            gotStorage = true;
          }
          if (response.accountEmail) {
            accountEmail = response.accountEmail;
            updateActiveAccountEmailUI(accountEmail);
          }
        }
      } catch (e) {}

      // Method B: executeScript fallback
      if (!gotStorage) {
        try {
          const scriptRes = await executeScriptInTab(currentActiveTab.id, () => {
            const ls = {};
            const ss = {};
            try {
              for (let i = 0; i < localStorage.length; i++) {
                const k = localStorage.key(i);
                ls[k] = localStorage.getItem(k);
              }
            } catch (e) {}
            try {
              for (let i = 0; i < sessionStorage.length; i++) {
                const k = sessionStorage.key(i);
                ss[k] = sessionStorage.getItem(k);
              }
            } catch (e) {}

            let foundEmail = null;
            const sec = document.querySelector('#mectrl_currentAccount_secondary, [id*="currentAccount_secondary"]');
            if (sec && sec.textContent && sec.textContent.includes('@')) {
              foundEmail = sec.textContent.trim().toLowerCase();
            }
            return { ls, ss, accountEmail: foundEmail, title: document.title };
          });

          if (scriptRes && scriptRes[0] && scriptRes[0].result) {
            storageData.localStorage = scriptRes[0].result.ls || {};
            storageData.sessionStorage = scriptRes[0].result.ss || {};
            if (scriptRes[0].result.accountEmail && !accountEmail) {
              accountEmail = scriptRes[0].result.accountEmail;
              updateActiveAccountEmailUI(accountEmail);
            }
          }
        } catch (err) {
          console.warn("Script execution fallback warning:", err);
        }
      }
    }

    // Record email in MRU tracked emails
    if (accountEmail) {
      await recordTrackedEmail(accountEmail, isOutlookSessionActive ? "Outlook" : (currentBaseDomain || "Universal"));
    }

    const lsKeysCount = Object.keys(storageData.localStorage).length;
    const ssKeysCount = Object.keys(storageData.sessionStorage).length;
    const totalStorage = lsKeysCount + ssKeysCount;
    document.getElementById('stat-storage-count').textContent = totalStorage;

    const isFirefox = navigator.userAgent.toLowerCase().includes('firefox');
    const sessionPackage = {
      app: "CookieSync",
      version: "2.1.0",
      type: isOutlookSessionActive ? "outlook" : "universal",
      exportedAt: new Date().toISOString(),
      accountEmail: accountEmail || null,
      sourceBrowser: isFirefox ? "Firefox" : "Chrome",
      sourceUrl: sourceUrl,
      targetDomain: isOutlookSessionActive ? "live.com" : (currentBaseDomain || currentTabDomain),
      cookieCount: scannedCookies.length,
      storageCount: totalStorage,
      cookies: scannedCookies,
      storage: storageData
    };

    lastExportedSession = sessionPackage;

    const jsonString = JSON.stringify(sessionPackage, null, 2);
    let clipboardSuccess = false;
    try {
      await navigator.clipboard.writeText(jsonString);
      clipboardSuccess = true;
    } catch (e) {
      const ta = document.getElementById('export-json-textarea');
      ta.value = jsonString;
      ta.select();
      clipboardSuccess = document.execCommand('copy');
    }

    document.getElementById('export-json-textarea').value = jsonString;
    document.getElementById('download-json-btn').disabled = false;
    document.getElementById('toggle-json-preview-btn').disabled = false;

    const scopeDesc = accountEmail ? `(${accountEmail})` : (isOutlookSessionActive ? 'Outlook' : (currentBaseDomain || 'active page'));
    showToast(`✓ Copied ${scannedCookies.length} cookies for ${scopeDesc}!`);
    btnLabel.textContent = 'Session Copied!';
    setTimeout(() => {
      btnLabel.textContent = originalLabel;
      exportBtn.disabled = false;
    }, 2000);

  } catch (err) {
    console.error("Export failed:", err);
    showToast(`Export error: ${err.message}`, true);
    btnLabel.textContent = originalLabel;
    exportBtn.disabled = false;
  }
}

function handleDownloadJson() {
  if (!lastExportedSession) return;
  const domainSlug = (lastExportedSession.accountEmail || lastExportedSession.targetDomain || 'session').replace(/[^a-z0-9]/gi, '_');
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(lastExportedSession, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `cookie_sync_${domainSlug}_${new Date().toISOString().slice(0, 10)}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
  showToast("Session file downloaded!");
}

// -------------------------------------------------------------
// IMPORT SESSION LOGIC (Purge Old Cookies & Data First, Then Inject)
// -------------------------------------------------------------
async function handlePasteAndLogin() {
  const pasteBtn = document.getElementById('paste-and-login-btn');
  const btnLabel = document.getElementById('import-btn-label');
  const originalLabel = btnLabel.textContent;
  const textarea = document.getElementById('import-session-textarea');

  pasteBtn.disabled = true;
  btnLabel.textContent = 'Pasting & Logging In...';

  try {
    let text = textarea.value.trim();

    if (!text && navigator.clipboard && navigator.clipboard.readText) {
      try {
        text = await navigator.clipboard.readText();
        text = (text || '').trim();
      } catch (err) {}
    }

    if (!text) {
      textarea.focus();
      try {
        document.execCommand('paste');
        text = textarea.value.trim();
      } catch (e) {}
    }

    if (!text) {
      showToast('Clipboard empty! Press Ctrl+V in the box to paste your session.', true);
      textarea.focus();
      pasteBtn.disabled = false;
      btnLabel.textContent = originalLabel;
      return;
    }

    textarea.value = text;
    await executeImport(text);

  } catch (err) {
    console.error("Import error:", err);
    showToast(`Import failed: ${err.message}`, true);
  } finally {
    pasteBtn.disabled = false;
    btnLabel.textContent = originalLabel;
  }
}

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (event) => {
    const text = event.target.result;
    document.getElementById('import-session-textarea').value = text;
    showToast('File loaded! Automatically importing...');
    executeImport(text);
  };
  reader.readAsText(file);
}

// Execute the Full Import Sequence
async function executeImport(jsonString) {
  let sessionData;
  try {
    sessionData = JSON.parse(jsonString);
  } catch (e) {
    showToast('Invalid JSON format! Please copy a valid session package.', true);
    return;
  }

  if (!sessionData.cookies || !Array.isArray(sessionData.cookies)) {
    showToast('Invalid session package: No cookies found in payload.', true);
    return;
  }

  const progressCard = document.getElementById('import-progress-card');
  const progressFill = document.getElementById('progress-bar-fill');
  const progressMsg = document.getElementById('progress-message');

  progressCard.style.display = 'block';
  progressFill.style.width = '10%';
  progressMsg.textContent = 'Preparing clean environment...';

  try {
    const isOutlook = sessionData.type === 'outlook' || (sessionData.cookies && sessionData.cookies.some(c => isMicrosoftDomain(c.domain)));
    const targetDomain = sessionData.targetDomain || (isOutlook ? 'live.com' : (currentBaseDomain || ''));

    // Record imported account email into MRU list
    let importedEmail = sessionData.accountEmail;
    if (!importedEmail && sessionData.cookies) {
      const anchorCookie = sessionData.cookies.find(c => c.name === 'DefaultAnchorMailbox' || c.name === 'SignInName');
      if (anchorCookie && anchorCookie.value && anchorCookie.value.includes('@')) {
        const val = decodeURIComponent(anchorCookie.value);
        const m = val.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        if (m) importedEmail = m[0];
      }
    }
    if (importedEmail) {
      await recordTrackedEmail(importedEmail, isOutlook ? "Outlook Import" : "Universal Import");
    }

    // Step 1: Pre-register storage in extension storage for document_start injection
    const storageApi = (typeof browser !== 'undefined' && browser.storage) ? browser.storage.local : chrome.storage?.local;
    if (storageApi && sessionData.storage) {
      await new Promise(resolve => {
        storageApi.set({
          pendingStorage: sessionData.storage,
          pendingDomain: targetDomain,
          pendingTimestamp: Date.now()
        }, () => resolve());
      });
    }

    // Step 2: Clear existing cookies and storage first
    progressMsg.textContent = `Purging existing cookies and data for ${targetDomain || 'target site'}...`;
    progressFill.style.width = '25%';

    if (isOutlook) {
      await purgeOutlookCookiesInternal();
    } else if (targetDomain) {
      await purgeDomainCookiesInternal(targetDomain);
    }

    if (currentActiveTab && currentActiveTab.id) {
      try {
        await sendTabMessage(currentActiveTab.id, { action: "CLEAR_PAGE_DATA" });
      } catch (e) {}
    }

    progressFill.style.width = '45%';

    // Step 3: Set all cookies from session package
    progressMsg.textContent = `Injecting ${sessionData.cookies.length} cookies...`;
    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < sessionData.cookies.length; i++) {
      const cookie = sessionData.cookies[i];
      const res = await setCookieSafe(cookie);
      if (res.success) {
        successCount++;
      } else {
        failCount++;
      }
      const pct = 45 + Math.floor((i / sessionData.cookies.length) * 40);
      progressFill.style.width = `${pct}%`;
    }

    progressFill.style.width = '88%';
    progressMsg.textContent = `Injected ${successCount} cookies!`;

    // Step 4: Determine Target URL & Launch/Refresh
    let targetUrl = sessionData.sourceUrl || "https://outlook.live.com/mail/0/";
    if (!sessionData.sourceUrl) {
      targetUrl = isOutlook ? "https://outlook.live.com/mail/0/" : `https://${targetDomain}/`;
    }

    progressFill.style.width = '95%';
    progressMsg.textContent = `Launching session for ${importedEmail || targetDomain || 'target'}...`;

    let matchPattern = "*://*." + (targetDomain || 'live.com') + "/*";
    if (isOutlook) matchPattern = "*://*.live.com/*";

    const openTabs = await queryTabs({ url: [matchPattern] });

    if (openTabs && openTabs.length > 0) {
      const tab = openTabs[0];
      try {
        await sendTabMessage(tab.id, { action: "INJECT_STORAGE", data: sessionData.storage });
      } catch (e) {}
      await reloadTab(tab.id);
      await updateTab(tab.id, { active: true });
      progressFill.style.width = '100%';
      progressMsg.textContent = 'Session refreshed & authenticated!';
    } else {
      await createTab(targetUrl);
      progressFill.style.width = '100%';
      progressMsg.textContent = 'Session opened & authenticated!';
    }

    const emailNotice = importedEmail ? ` (${importedEmail})` : '';
    showToast(`✓ Logged in${emailNotice}! ${successCount} cookies applied.`);

    setTimeout(() => {
      progressCard.style.display = 'none';
      progressFill.style.width = '0%';
    }, 4000);

  } catch (err) {
    console.error("Import execution failed:", err);
    progressMsg.textContent = `Failed: ${err.message}`;
    showToast(`Error: ${err.message}`, true);
  }
}

// -------------------------------------------------------------
// BULLETPROOF COOKIE SETTER (Cross-Browser Resilient)
// -------------------------------------------------------------
async function setCookieSafe(cookie) {
  const cookiesApi = (typeof browser !== 'undefined' && browser.cookies) ? browser.cookies : chrome.cookies;
  const isSecure = Boolean(cookie.secure);
  
  const rawDomain = cookie.domain || "";
  const domainClean = rawDomain.replace(/^\./, "");
  const pathClean = cookie.path || "/";
  const url = `https://${domainClean}${pathClean.startsWith("/") ? pathClean : "/" + pathClean}`;

  const details = {
    url: url,
    name: cookie.name,
    value: cookie.value || "",
    path: cookie.path || "/",
    secure: isSecure,
    httpOnly: Boolean(cookie.httpOnly)
  };

  if (!cookie.hostOnly && domainClean) {
    details.domain = domainClean;
  }

  // SameSite handling
  if (cookie.sameSite) {
    const s = String(cookie.sameSite).toLowerCase();
    if (s === 'lax') details.sameSite = 'lax';
    else if (s === 'strict') details.sameSite = 'strict';
    else if (s === 'no_restriction' || s === 'none') {
      details.sameSite = 'no_restriction';
      details.secure = true;
    }
  }

  // Expiration handling
  if (cookie.expirationDate && !cookie.session) {
    details.expirationDate = cookie.expirationDate;
  }

  try {
    const res = await new Promise((resolve) => {
      cookiesApi.set(details, (c) => {
        const lastErr = ext.runtime.lastError;
        if (lastErr || !c) {
          resolve({ success: false, error: lastErr?.message });
        } else {
          resolve({ success: true, cookie: c });
        }
      });
    });

    if (res.success) return res;

    // Resilient Fallback
    const fallbackDetails = {
      url: url,
      name: cookie.name,
      value: cookie.value || "",
      path: "/",
      secure: isSecure
    };
    return await new Promise((resolve) => {
      cookiesApi.set(fallbackDetails, (c) => {
        resolve({ success: !!c });
      });
    });

  } catch (err) {
    return { success: false, error: err.message };
  }
}

// -------------------------------------------------------------
// COOKIE PURGING HELPERS
// -------------------------------------------------------------
async function purgeDomainCookiesInternal(domain) {
  const cookiesApi = (typeof browser !== 'undefined' && browser.cookies) ? browser.cookies : chrome.cookies;
  const cleanDomain = domain.replace(/^\./, '');
  const allCookies = await getAllCookies({});
  let count = 0;

  for (const c of (allCookies || [])) {
    if (c.domain.includes(cleanDomain)) {
      const url = `http${c.secure ? 's' : ''}://${c.domain.replace(/^\./, '')}${c.path}`;
      await new Promise(r => cookiesApi.remove({ url, name: c.name }, () => r()));
      count++;
    }
  }
  return count;
}

async function purgeOutlookCookiesInternal() {
  const cookiesApi = (typeof browser !== 'undefined' && browser.cookies) ? browser.cookies : chrome.cookies;
  const allCookies = await getAllCookies({});
  let count = 0;

  for (const c of (allCookies || [])) {
    if (isMicrosoftDomain(c.domain)) {
      const url = `http${c.secure ? 's' : ''}://${c.domain.replace(/^\./, '')}${c.path}`;
      await new Promise(r => cookiesApi.remove({ url, name: c.name }, () => r()));
      count++;
    }
  }
  return count;
}

async function handleClearActiveCookies() {
  const targetName = isOutlookSessionActive ? 'Outlook / Microsoft' : (currentBaseDomain || 'current website');
  if (!confirm(`Are you sure you want to clear all cookies for ${targetName}?`)) return;

  let cleared = 0;
  if (isOutlookSessionActive) {
    cleared = await purgeOutlookCookiesInternal();
  } else if (currentBaseDomain) {
    cleared = await purgeDomainCookiesInternal(currentBaseDomain);
  }

  showToast(`Purged ${cleared} cookies for ${targetName}!`);
  await scanActiveCookies();
  renderCookiesList();
}

// -------------------------------------------------------------
// INSPECT LIST RENDERER
// -------------------------------------------------------------
function renderCookiesList(filter = '') {
  const listEl = document.getElementById('cookie-list');
  listEl.innerHTML = '';

  const q = filter.trim().toLowerCase();
  const filtered = scannedCookies.filter(c => !q || c.name.toLowerCase().includes(q) || c.domain.toLowerCase().includes(q));

  if (!filtered.length) {
    listEl.innerHTML = '<div class="empty-state">No matching cookies found.</div>';
    return;
  }

  filtered.slice(0, 100).forEach(c => {
    const row = document.createElement('div');
    row.className = 'cookie-row';
    row.innerHTML = `
      <span class="cookie-name" title="${c.name}">${c.name}</span>
      <span class="cookie-domain" title="${c.domain}">${c.domain}</span>
    `;
    listEl.appendChild(row);
  });
}

// -------------------------------------------------------------
// CHROME / FIREFOX API HELPERS
// -------------------------------------------------------------
function queryTabs(queryInfo) {
  return new Promise((resolve) => {
    const tabsApi = (typeof browser !== 'undefined' && browser.tabs) ? browser.tabs : chrome.tabs;
    tabsApi.query(queryInfo, (tabs) => resolve(tabs || []));
  });
}

function getAllCookies(details) {
  return new Promise((resolve) => {
    const cookiesApi = (typeof browser !== 'undefined' && browser.cookies) ? browser.cookies : chrome.cookies;
    cookiesApi.getAll(details, (cookies) => resolve(cookies || []));
  });
}

function sendTabMessage(tabId, message) {
  return new Promise((resolve, reject) => {
    const tabsApi = (typeof browser !== 'undefined' && browser.tabs) ? browser.tabs : chrome.tabs;
    tabsApi.sendMessage(tabId, message, (response) => {
      const err = ext.runtime.lastError;
      if (err) reject(err);
      else resolve(response);
    });
  });
}

function executeScriptInTab(tabId, func) {
  return new Promise((resolve, reject) => {
    const scriptingApi = (typeof browser !== 'undefined' && browser.scripting) ? browser.scripting : chrome.scripting;
    if (scriptingApi && scriptingApi.executeScript) {
      scriptingApi.executeScript({
        target: { tabId: tabId },
        func: func
      }, (results) => {
        const err = ext.runtime.lastError;
        if (err) reject(err);
        else resolve(results);
      });
    } else {
      const tabsApi = (typeof browser !== 'undefined' && browser.tabs) ? browser.tabs : chrome.tabs;
      tabsApi.executeScript(tabId, { code: `(${func.toString()})()` }, (results) => {
        resolve([{ result: results ? results[0] : null }]);
      });
    }
  });
}

function createTab(url) {
  return new Promise((resolve) => {
    const tabsApi = (typeof browser !== 'undefined' && browser.tabs) ? browser.tabs : chrome.tabs;
    tabsApi.create({ url }, (tab) => resolve(tab));
  });
}

function reloadTab(tabId) {
  return new Promise((resolve) => {
    const tabsApi = (typeof browser !== 'undefined' && browser.tabs) ? browser.tabs : chrome.tabs;
    tabsApi.reload(tabId, () => resolve());
  });
}

function updateTab(tabId, updateProps) {
  return new Promise((resolve) => {
    const tabsApi = (typeof browser !== 'undefined' && browser.tabs) ? browser.tabs : chrome.tabs;
    tabsApi.update(tabId, updateProps, (tab) => resolve(tab));
  });
}

function getStorageData(keys) {
  return new Promise((resolve) => {
    const storageApi = (typeof browser !== 'undefined' && browser.storage) ? browser.storage.local : chrome.storage?.local;
    if (storageApi) {
      storageApi.get(keys, (res) => resolve(res || {}));
    } else {
      resolve({});
    }
  });
}

function setStorageData(data) {
  return new Promise((resolve) => {
    const storageApi = (typeof browser !== 'undefined' && browser.storage) ? browser.storage.local : chrome.storage?.local;
    if (storageApi) {
      storageApi.set(data, () => resolve());
    } else {
      resolve();
    }
  });
}

// Toast Utility
function showToast(message, isError = false) {
  const toast = document.getElementById('toast');
  const toastMsg = document.getElementById('toast-msg');
  const toastIcon = document.getElementById('toast-icon');

  toastMsg.textContent = message;
  toastIcon.textContent = isError ? '✕' : '✓';
  toast.className = 'toast' + (isError ? ' toast-error' : '') + ' show';

  setTimeout(() => {
    toast.className = 'toast';
  }, 2600);
}
