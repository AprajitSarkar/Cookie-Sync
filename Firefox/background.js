/**
 * Cookie Sync - Firefox Background Script
 * Handles lifecycle, extension state, and session routing for Mozilla Firefox.
 */

const ext = typeof browser !== 'undefined' ? browser : chrome;

ext.runtime.onInstalled.addListener((details) => {
  console.log(`[Cookie Sync] Firefox Extension installed/updated (${details.reason}) - Version: 2.0.0`);
});

// Listener for background actions or relay
ext.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'PING') {
    sendResponse({ status: 'PONG', version: '2.0.0' });
    return true;
  }
});
