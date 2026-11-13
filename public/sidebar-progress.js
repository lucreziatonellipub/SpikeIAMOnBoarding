(() => {
  "use strict";

  const detailsByThread = new Map();
  const versions = new Map();
  const pending = new Map();
  // Resolve relative to this asset to support Chainlit mounted under a prefix.
  const endpoint = new URL("../onboarding/progress", document.currentScript.src);
  let scheduled = false;
  let retryAfter = 0;
  let version = 0;
  let stream = null;

  // Match chat rows only: #thread-options also starts with "thread-".
  const rows = () => [...document.querySelectorAll('#thread-history li[id^="thread-"]')];
  const threadId = (row) => row.id.slice("thread-".length);

  function normalizeProgress(value) {
    return Number.isFinite(value) ? Math.max(0, Math.min(100, Math.round(value))) : 0;
  }

  function render(row, details) {
    const progress = normalizeProgress(details.progress);
    if (details.name) {
      const title = row.querySelector('[data-sidebar="menu-button"] .truncate');
      if (title) {
        // React owns its text nodes. Display the live title through CSS instead
        // of replacing nodes that React may still need during navigation/send.
        title.dataset.onboardingTitle = details.name;
        title.title = details.name;
      }
    }
    let indicator = row.querySelector(":scope > .onboarding-progress");
    if (!indicator) {
      indicator = document.createElement("div");
      indicator.className = "onboarding-progress";
      const label = document.createElement("span");
      label.className = "onboarding-progress-label";
      const track = document.createElement("div");
      track.className = "onboarding-progress-track";
      const fill = document.createElement("div");
      fill.className = "onboarding-progress-fill";
      track.append(fill);
      indicator.append(label, track);
      // Outside React's link/button: preserve its title and rename/menu behavior.
      row.append(indicator);
    }
    const italian = (details.lang || document.documentElement.lang || navigator.language).startsWith("it");
    const signature = `${progress}:${italian}`;
    if (indicator.dataset.progress === signature) return;
    indicator.dataset.progress = signature;
    const label = `${italian ? "Completamento" : "Completion"}: ${progress}%`;
    indicator.querySelector(".onboarding-progress-label").textContent = label;
    const track = indicator.querySelector(".onboarding-progress-track");
    track.setAttribute("role", "progressbar");
    track.setAttribute("aria-label", italian ? "Completamento questionario" : "Questionnaire completion");
    track.setAttribute("aria-valuemin", "0");
    track.setAttribute("aria-valuemax", "100");
    track.setAttribute("aria-valuenow", String(progress));
    indicator.querySelector(".onboarding-progress-fill").style.width = `${progress}%`;
  }

  async function loadProgress(ids) {
    const requestVersion = ++version;
    ids.forEach((id) => {
      pending.set(id, requestVersion);
      versions.set(id, requestVersion);
    });
    try {
      const response = await fetch(endpoint, {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ thread_ids: ids }),
      });
      if (!response.ok) throw new Error(`Progress request failed: ${response.status}`);
      const values = await response.json();
      for (const id of ids) {
        // Live events and later requests take precedence over a stale response.
        if (versions.get(id) === requestVersion) {
          detailsByThread.set(id, values[id] ?? { progress: 0 });
        }
      }
      retryAfter = 0;
    } catch (error) {
      console.warn("Could not load questionnaire progress", error);
      retryAfter = Date.now() + 5000;
      setTimeout(scheduleRender, 5000);
    } finally {
      ids.forEach((id) => {
        if (pending.get(id) === requestVersion) pending.delete(id);
      });
      scheduleRender();
    }
  }

  function updateRows() {
    scheduled = false;
    const chatRows = rows();
    if (chatRows.length) connectStream();
    const missing = [];
    chatRows.forEach((row) => {
      const id = threadId(row);
      if (detailsByThread.has(id)) {
        render(row, detailsByThread.get(id));
      } else if (!pending.has(id)) {
        missing.push(id);
      }
    });
    if (missing.length && Date.now() >= retryAfter) loadProgress(missing.slice(0, 1000));
  }

  function scheduleRender() {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(updateRows);
  }

  function connectStream() {
    if (stream || document.hidden) return;
    // A direct connection also works when Chainlit is embedded in an iframe;
    // Chainlit's window_message is forwarded to the parent window instead.
    stream = new EventSource(`${endpoint.href}/events`, { withCredentials: true });
    stream.onmessage = (event) => {
      let data;
      try { data = JSON.parse(event.data); } catch { return; }
      if (!data || typeof data.thread_id !== "string" || !Number.isFinite(data.progress)) return;
      versions.set(data.thread_id, ++version);
      detailsByThread.set(data.thread_id, data);
      scheduleRender();
    };
    // EventSource reconnects automatically. Reload saved values after a gap.
    stream.onopen = () => {
      const ids = rows().map(threadId);
      if (ids.length) loadProgress(ids.slice(0, 1000));
    };
  }

  window.addEventListener("pagehide", () => {
    stream?.close();
    stream = null;
  });
  window.addEventListener("pageshow", scheduleRender);
  // Background tabs must not exhaust the browser's per-origin HTTP connection
  // pool and starve Chainlit's socket/polling requests in the active tab.
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      stream?.close();
      stream = null;
    } else {
      scheduleRender();
    }
  });

  function start() {
    // Handles pagination, new chats, navigation and sidebar remounts.
    new MutationObserver(scheduleRender).observe(document.body, { childList: true, subtree: true });
    scheduleRender();
  }
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }
})();
