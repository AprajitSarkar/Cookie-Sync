/**
 * Cookie Sync - Universal Content Script
 * Injected at document_start across all websites.
 * Facilitates storage extraction, pre-injection at document_start, authenticated email extraction, and session cleanup.
 */

(function() {
  const hostname = window.location.hostname.toLowerCase();

  // Helper: Extract authenticated account email (Outlook / Microsoft / General)
  function extractAuthenticatedEmail() {
    try {
      // 1. Direct Microsoft MeControl ID (exact element shown in user's account dropdown)
      const meControlSec = document.querySelector('#mectrl_currentAccount_secondary, [id*="currentAccount_secondary"], .mectrl_account_secondary');
      if (meControlSec && meControlSec.textContent) {
        const text = meControlSec.textContent.trim().toLowerCase();
        if (text.includes('@') && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text)) {
          return text;
        }
      }

      // 2. Microsoft Account Manager Trigger button (aria-label)
      const navBtn = document.querySelector('#O365_NavHeader button[aria-label*="@"], #mectrl_main_trigger[aria-label*="@"]');
      if (navBtn) {
        const aria = navBtn.getAttribute('aria-label') || '';
        const match = aria.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        if (match) return match[0].toLowerCase();
      }

      // 3. MeControl Popup Dialog Container (strictly checks for "Sign out" & "My Microsoft account")
      const mePopup = document.querySelector('#mectrl_popup, [id*="mectrl_account_menu"], div[role="dialog"]');
      if (mePopup && (mePopup.textContent.includes('Sign out') || mePopup.textContent.includes('My Microsoft account'))) {
        const emailMatch = mePopup.textContent.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        if (emailMatch) return emailMatch[0].toLowerCase();
      }

      // 4. Inspect localStorage for authoritative Microsoft UPN (ClientConfigure / BootDiagnostics)
      try {
        for (let i = 0; i < localStorage.length; i++) {
          const k = localStorage.key(i);
          if (k && (k.includes('ClientConfigure') || k.includes('BootDiagnostics') || k.includes('MSAL') || k.includes('user'))) {
            const v = localStorage.getItem(k);
            if (v && v.includes('"upn"')) {
              const m = v.match(/"upn"\s*:\s*"([^"]+@[^"]+)"/i);
              if (m) return m[1].toLowerCase();
            }
          }
        }
      } catch (e) {}

      // 5. Inspect sessionStorage for MSAL authenticated user
      try {
        for (let i = 0; i < sessionStorage.length; i++) {
          const k = sessionStorage.key(i);
          if (k && k.toLowerCase().includes('msal')) {
            const v = sessionStorage.getItem(k);
            if (v && v.includes('"username"')) {
              const m = v.match(/"username"\s*:\s*"([^"]+@[^"]+)"/i);
              if (m) return m[1].toLowerCase();
            }
          }
        }
      } catch (e) {}

    } catch (err) {
      console.warn("[Cookie Sync] Email extraction warning:", err);
    }
    return null;
  }

  // 1. Instant document_start storage pre-injection
  try {
    const storageApi = (typeof browser !== 'undefined' && browser.storage) ? browser.storage.local : chrome.storage?.local;
    if (storageApi) {
      storageApi.get(["pendingStorage", "pendingDomain", "pendingTimestamp"], (res) => {
        if (res && res.pendingStorage && res.pendingTimestamp) {
          if (Date.now() - res.pendingTimestamp < 180000) {
            const targetDomain = (res.pendingDomain || "").toLowerCase();
            const matchesDomain = !targetDomain || 
              hostname.includes(targetDomain) || 
              targetDomain.includes(hostname) ||
              (/outlook|live\.com|office\.com|microsoft/i.test(targetDomain) && /outlook|live\.com|office\.com|microsoft/i.test(hostname));

            if (matchesDomain) {
              const data = res.pendingStorage;
              let lsCount = 0;
              let ssCount = 0;

              if (data.localStorage && typeof data.localStorage === 'object') {
                for (const [k, v] of Object.entries(data.localStorage)) {
                  try {
                    window.localStorage.setItem(k, typeof v === 'string' ? v : JSON.stringify(v));
                    lsCount++;
                  } catch (e) {}
                }
              }

              if (data.sessionStorage && typeof data.sessionStorage === 'object') {
                for (const [k, v] of Object.entries(data.sessionStorage)) {
                  try {
                    window.sessionStorage.setItem(k, typeof v === 'string' ? v : JSON.stringify(v));
                    ssCount++;
                  } catch (e) {}
                }
              }

              console.log(`[Cookie Sync] Pre-injected ${lsCount} localStorage & ${ssCount} sessionStorage keys at document_start for ${hostname}!`);
              storageApi.remove(["pendingStorage", "pendingDomain", "pendingTimestamp"]);
            }
          }
        }
      });
    }
  } catch (err) {
    console.warn("[Cookie Sync] Storage pre-injection warning:", err);
  }

  // 2. Runtime message listener for active popup communication
  try {
    const runtimeApi = (typeof browser !== 'undefined' && browser.runtime) ? browser.runtime : chrome.runtime;
    if (runtimeApi && runtimeApi.onMessage) {
      runtimeApi.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === "GET_STORAGE") {
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

          const accountEmail = extractAuthenticatedEmail();

          sendResponse({
            localStorage: ls,
            sessionStorage: ss,
            accountEmail: accountEmail,
            title: document.title,
            url: window.location.href,
            domain: window.location.hostname
          });
          return true;
        }

        if (request.action === "INJECT_STORAGE") {
          const data = request.data || {};
          let lsCount = 0;
          let ssCount = 0;

          if (data.localStorage && typeof data.localStorage === 'object') {
            for (const [k, v] of Object.entries(data.localStorage)) {
              try {
                localStorage.setItem(k, typeof v === 'string' ? v : JSON.stringify(v));
                lsCount++;
              } catch (e) {}
            }
          }

          if (data.sessionStorage && typeof data.sessionStorage === 'object') {
            for (const [k, v] of Object.entries(data.sessionStorage)) {
              try {
                sessionStorage.setItem(k, typeof v === 'string' ? v : JSON.stringify(v));
                ssCount++;
              } catch (e) {}
            }
          }

          sendResponse({ success: true, lsCount, ssCount });
          return true;
        }

        if (request.action === "CLEAR_PAGE_DATA") {
          try {
            localStorage.clear();
            sessionStorage.clear();
          } catch (e) {}
          sendResponse({ success: true });
          return true;
        }
      });
    }
  } catch (err) {
    console.warn("[Cookie Sync] Message listener init warning:", err);
  }
})();
