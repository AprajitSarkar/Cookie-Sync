# 🍪 Cookie Sync

> **Universal & Outlook Session Migrator for Google Chrome & Mozilla Firefox**  
> *Developed by [Aprajit Sarkar](https://github.com/AprajitSarkar)*

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Manifest V3](https://img.shields.io/badge/Manifest-V3-success.svg)](https://developer.chrome.com/docs/extensions/mv3/intro/)
[![Firefox AMO](https://img.shields.io/badge/Firefox-Add--ons-FF7139.svg)](https://addons.mozilla.org/)
[![Chrome Web Store](https://img.shields.io/badge/Chrome-Web%20Store-38BDF8.svg)](https://chrome.google.com/webstore)

---

## 🚀 Overview

**Cookie Sync** (also known as *Outlook Account Sync*) is an ultra-fast, cross-browser session and cookie migration extension built for **Google Chrome**, **Mozilla Firefox**, Microsoft Edge, and Chromium-based browsers.

Transfer logged-in sessions seamlessly across browsers or profiles **in 1 click via clipboard**—eliminating the need to re-type passwords, perform 2-Factor Authentication (2FA), or resolve security prompts.

Whether you're migrating an **Outlook 365 / Personal mailbox** or an active session on **LinkedIn, GitHub, Twitter/X, or any web application**, Cookie Sync handles the extraction, sanitization, and deep injection automatically.

---

## ✨ Key Features

- ⚡ **Deep Outlook / Microsoft Auth Migration**:
  - Automatically identifies Microsoft Exchange and MSAL authentication domains (`live.com`, `office.com`, `office365.com`, `microsoft.com`, `microsoftonline.com`, `outlook.com`, `msauth.net`, `passport.net`).
  - Gathers all session cookies, authentication tokens, `localStorage`, and `sessionStorage`.
- 🌐 **Universal Website Support**:
  - Works on **any website** (LinkedIn, GitHub, Slack, social networks, admin dashboards).
  - Intelligently scopes cookies by base domain and host.
- 🧹 **Automatic Clean Purge Before Injection**:
  - Prevents session conflict and authentication loops by clearing existing cookies and stored client data for the target site *before* injecting the new session.
- 📋 **1-Click Clipboard Synchronization**:
  - Click **Copy Session** to create a lightweight JSON payload.
  - Click **Paste from Clipboard & Login** in the target browser to instantly authenticate.
- 🛡️ **Zero-Server / Privacy First**:
  - 100% client-side operation. No credentials, tokens, or cookies ever touch an external server.
- 🔍 **Built-In Cookie Inspector**:
  - Live inspection, real-time search/filter, and instant cookie purging for any domain.

---

## 📂 Repository Structure

```
Outlook Account Sync / Cookie-Sync
├── Chrome/                     # Unpacked Chrome Manifest V3 Extension
│   ├── manifest.json
│   ├── popup.html
│   ├── popup.js
│   ├── popup.css
│   ├── content.js
│   ├── background.js
│   └── icons/
├── Firefox/                    # Unpacked Firefox Manifest V3 Extension (AMO ready)
│   ├── manifest.json
│   ├── popup.html
│   ├── popup.js
│   ├── popup.css
│   ├── content.js
│   ├── background.js
│   └── icons/
├── Cookie-Sync-Chrome.zip      # Production package for Chrome Web Store
├── Cookie-Sync-Firefox.zip     # Production package for Firefox Add-ons (AMO)
├── package_extensions.ps1      # Packaging script
├── LICENSE                     # MIT License
└── README.md
```

---

## 📥 Installation Guide

### For Google Chrome / Brave / Edge:
1. Download or clone this repository.
2. Open Chrome and navigate to `chrome://extensions/`.
3. Enable **Developer mode** (toggle in the top-right corner).
4. Click **Load unpacked** and select the [`Chrome`](Chrome/) folder.
5. Alternatively, upload [`Cookie-Sync-Chrome.zip`](Cookie-Sync-Chrome.zip) directly to the **Chrome Web Store Developer Dashboard**.

### For Mozilla Firefox:
1. Open Firefox and navigate to `about:debugging#/runtime/this-firefox`.
2. Click **Load Temporary Add-on...**.
3. Select `manifest.json` inside the [`Firefox`](Firefox/) folder.
4. For permanent installation or distribution, upload [`Cookie-Sync-Firefox.zip`](Cookie-Sync-Firefox.zip) to the **[Firefox Add-on Developer Hub (AMO)](https://addons.mozilla.org/developers/)**.

---

## 🛠️ Usage Instructions

### 1. Exporting a Session (e.g., from Firefox):
1. Log in to **Microsoft Outlook** or your desired website in Firefox.
2. Click the **Cookie Sync** extension icon in your browser toolbar.
3. The extension detects the active site (e.g. *Outlook Mode* or *Universal Mode: linkedin.com*).
4. Click **Copy Session to Clipboard** (or **Save File**).

### 2. Importing a Session (e.g., into Chrome):
1. Open Chrome.
2. Click the **Cookie Sync** extension icon.
3. Click **Paste from Clipboard & Login** (or paste manually with `Ctrl+V`).
4. The extension automatically:
   - Cleans up stale cookies and storage on the target site.
   - Applies the authenticated cookies and storage keys.
   - Launches / refreshes the page with you fully logged in!

---

## 🏷️ Keywords & Topics
`cookie-sync` • `outlook-sync` • `session-migrator` • `cookie-manager` • `chrome-extension` • `firefox-addon` • `browser-extension` • `webextension` • `cross-browser` • `manifest-v3` • `productivity` • `privacy-first`

---

## 📄 License
This project is licensed under the [MIT License](LICENSE) - see the LICENSE file for details.

Developed with ❤️ by **[Aprajit Sarkar](https://github.com/AprajitSarkar)**.
