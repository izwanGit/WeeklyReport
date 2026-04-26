// Background service worker — no logic needed,
// all work is done in popup.js.
// This file must exist because manifest.json declares it.
chrome.runtime.onInstalled.addListener(() => {
  console.log("MyGenie Cookie Bridge installed.");
});
