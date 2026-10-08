/**
 * Cookie Sync - Universal & Outlook Session Migrator (v2.0.0)
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
let lastExportedSession = null;
let scannedCookies = [];

// Initialize Popup
document.addEventListener('DOMContentLoaded', async () => {
  detectBrowser();
  setupTabs();
  setupEventHandlers();
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

// Setup Navigation Tabs
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
  }
}

// Extract base domain e.g. "linkedin.com" from "www.linkedin.com"
function getBaseDomain(hostname) {
  if (!hostname) return '';
  const parts = hostname.split('.').filter(Boolean);
  if (parts.length <= 2) return hostname;
  // Handle two-part TLDs e.g. co.uk, com.au
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

      // Check if it's a browser system tab
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
// COOKIE SCANNER (Outlook vs. Universal Active Website)
// -------------------------------------------------------------
async function scanActiveCookies() {
  try {
    let matched = [];
    const seen = new Set();

    if (isOutlookSessionActive) {
      // Outlook Mode: Scan across all Microsoft auth domains
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
    } else if (currentActiveTab && currentActiveTab.url) {
      // Universal Mode: Query cookies for active URL and domains
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
    // 1. Scan latest cookies
    await scanActiveCookies();

    // 2. Extract localStorage & sessionStorage from active tab
    let storageData = { localStorage: {}, sessionStorage: {} };
    let sourceUrl = currentActiveTab ? currentActiveTab.url : 'https://outlook.live.com/mail/0/';

    if (currentActiveTab && currentActiveTab.id) {
      // Method A: Content script messaging
      let gotStorage = false;
      try {
        const response = await sendTabMessage(currentActiveTab.id, { action: "GET_STORAGE" });
        if (response && response.localStorage) {
          storageData.localStorage = response.localStorage;
          storageData.sessionStorage = response.sessionStorage || {};
          gotStorage = true;
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
            return { ls, ss, title: document.title };
          });

          if (scriptRes && scriptRes[0] && scriptRes[0].result) {
            storageData.localStorage = scriptRes[0].result.ls || {};
            storageData.sessionStorage = scriptRes[0].result.ss || {};
          }
        } catch (err) {
          console.warn("Script execution fallback warning:", err);
        }
      }
    }

    const lsKeysCount = Object.keys(storageData.localStorage).length;
    const ssKeysCount = Object.keys(storageData.sessionStorage).length;
    const totalStorage = lsKeysCount + ssKeysCount;
    document.getElementById('stat-storage-count').textContent = totalStorage;

    // 3. Construct Complete Session Package
    const isFirefox = navigator.userAgent.toLowerCase().includes('firefox');
    const sessionPackage = {
      app: "CookieSync",
      version: "2.0.0",
      type: isOutlookSessionActive ? "outlook" : "universal",
      exportedAt: new Date().toISOString(),
      sourceBrowser: isFirefox ? "Firefox" : "Chrome",
      sourceUrl: sourceUrl,
      targetDomain: isOutlookSessionActive ? "live.com" : (currentBaseDomain || currentTabDomain),
      cookieCount: scannedCookies.length,
      storageCount: totalStorage,
      cookies: scannedCookies,
      storage: storageData
    };

    lastExportedSession = sessionPackage;

    // 4. Write to Clipboard
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

    // 5. Update UI
    document.getElementById('export-json-textarea').value = jsonString;
    document.getElementById('download-json-btn').disabled = false;
    document.getElementById('toggle-json-preview-btn').disabled = false;

    const scopeName = isOutlookSessionActive ? 'Outlook' : (currentBaseDomain || 'active page');
    showToast(`✓ Copied ${scannedCookies.length} cookies & ${totalStorage} tokens for ${scopeName}!`);
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

// Download Session JSON
function handleDownloadJson() {
  if (!lastExportedSession) return;
  const domainSlug = (lastExportedSession.targetDomain || 'session').replace(/[^a-z0-9]/gi, '_');
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

    // Step 2: STRICT USER REQUIREMENT: Purge existing cookies and data of the target website first!
    progressMsg.textContent = `Purging existing cookies and data for ${targetDomain || 'target site'}...`;
    progressFill.style.width = '25%';

    if (isOutlook) {
      await purgeOutlookCookiesInternal();
    } else if (targetDomain) {
      await purgeDomainCookiesInternal(targetDomain);
    }

    // Also clear open tab's localStorage & sessionStorage if currently on that site
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
    progressMsg.textContent = `Launching session for ${targetDomain || 'target'}...`;

    // Check if a tab for this domain is already open
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

    showToast(`✓ Logged in! ${successCount} cookies applied.`);

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
      details.secure = true; // SameSite=None requires Secure
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

    // Resilient Fallback: Try with relaxed attributes if strict setting failed
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
      // Manifest V2 fallback for older Firefox
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
