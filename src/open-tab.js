document.getElementById("openTabBtn").addEventListener("click", () => {
  chrome.tabs.create({ url: chrome.runtime.getURL("tab.html") });
});
