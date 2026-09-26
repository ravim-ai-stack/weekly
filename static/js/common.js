// static/js/common.js - shared helpers used across all four pages.

function showToast(message, type) {
  const toast = document.getElementById("toast");
  if (!toast) return;
  toast.textContent = message;
  toast.className = `toast ${type}`;
  toast.classList.remove("hidden");
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => toast.classList.add("hidden"), 5000);
}

const Draft = {
  KEY: "weeklyReportDraft",

  save(data) {
    sessionStorage.setItem(Draft.KEY, JSON.stringify(data));
  },

  load() {
    const raw = sessionStorage.getItem(Draft.KEY);
    return raw ? JSON.parse(raw) : null;
  },

  update(partial) {
    const current = Draft.load() || {};
    Draft.save({ ...current, ...partial });
  },

  clear() {
    sessionStorage.removeItem(Draft.KEY);
  },

  requireOrRedirect(redirectTo, requiredKeys) {
    const data = Draft.load();
    if (!data || requiredKeys.some((k) => data[k] === undefined || data[k] === null)) {
      window.location.href = redirectTo;
      return null;
    }
    return data;
  },
};

const PageLoading = {
  show(message) {
    const overlay = document.getElementById("page_loading");
    const text = document.getElementById("page_loading_text");
    if (!overlay) return;
    if (text) text.textContent = message || "Working...";
    overlay.classList.remove("hidden");
  },

  hide() {
    const overlay = document.getElementById("page_loading");
    if (!overlay) return;
    overlay.classList.add("hidden");
  },
};

async function sendJSON(url, method, body) {
  const resp = await fetch(url, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {}),
  });
  let data;
  try {
    data = await resp.json();
  } catch (e) {
    data = { success: false, error: "Unexpected server response." };
  }
  if (!resp.ok && data.success === undefined) {
    data.success = false;
  }
  return data;
}

async function postJSON(url, body) {
  return sendJSON(url, "POST", body);
}

async function deleteJSON(url, body) {
  return sendJSON(url, "DELETE", body);
}
