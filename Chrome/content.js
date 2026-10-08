/**
 * Cookie Sync - Universal Content Script
 * Injected at document_start across all websites.
 * Facilitates storage extraction, pre-injection at document_start, and session cleanup.
 */

(function() {
  const hostname = window.location.hostname.toLowerCase();

  // 1. Instant document_start storage pre-injection
  try {
    const storageApi = (typeof browser !== 'undefined' && browser.storage) ? browser.storage.local : chrome.storage?.local;
    if (storageApi) {
      storageApi.get(["pendingStorage", "pendingDomain", "pendingTimestamp"], (res) => {
        if (res && res.pendingStorage && res.pendingTimestamp) {
          // Valid if set within the last 3 minutes
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

              // Clear pending storage once consumed
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
        // Extract localStorage & sessionStorage
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
          sendResponse({
            localStorage: ls,
            sessionStorage: ss,
            title: document.title,
            url: window.location.href,
            domain: window.location.hostname
          });
          return true;
        }

        // Inject storage keys
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

        // Purge client-side page storage
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
