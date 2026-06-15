const state = {
  q: "",
  accountTypes: [],
  states: [],
  msas: [],
  contractStatuses: [],
  sortBy: "account_name",
  theme: document.documentElement.dataset.theme || "dark",
  offset: 0,
  limit: 50,
  total: 0,
  selectedId: "",
  activeTab: "timeline",
  focusedRecord: null,
  accounts: [],
  detail: null,
  attachments: null,
  attachmentsLoading: false,
  attachmentsError: "",
  timelineType: "all",
  timelineQuery: "",
  attachmentCategory: "all",
  attachmentModule: "all",
  attachmentAvailability: "all",
  attachmentQuery: "",
  savedViews: [],
};

const els = {
  metrics: document.querySelector("#metrics"),
  search: document.querySelector("#search"),
  exportExcel: document.querySelector("#exportExcel"),
  exportContacts: document.querySelector("#exportContacts"),
  themeToggle: document.querySelector("#themeToggle"),
  clearFilters: document.querySelector("#clearFilters"),
  sortAccounts: document.querySelector("#sortAccounts"),
  savedViewName: document.querySelector("#savedViewName"),
  saveView: document.querySelector("#saveView"),
  savedViewsList: document.querySelector("#savedViewsList"),
  accountList: document.querySelector("#accountList"),
  resultCount: document.querySelector("#resultCount"),
  resultMeta: document.querySelector("#resultMeta"),
  currentView: document.querySelector("#currentView"),
  loadMore: document.querySelector("#loadMore"),
  detail: document.querySelector("#detail"),
};

const filterConfig = {
  account_type: {
    key: "accountTypes",
    label: "types",
    emptyLabel: "All types",
    summary: document.querySelector("#accountTypeSummary"),
    options: document.querySelector("#accountTypeOptions"),
  },
  state: {
    key: "states",
    label: "states",
    emptyLabel: "All states",
    summary: document.querySelector("#stateSummary"),
    options: document.querySelector("#stateOptions"),
  },
  msa: {
    key: "msas",
    label: "MSAs",
    emptyLabel: "All MSAs",
    summary: document.querySelector("#msaSummary"),
    options: document.querySelector("#msaOptions"),
  },
  contract_status: {
    key: "contractStatuses",
    label: "statuses",
    emptyLabel: "All statuses",
    summary: document.querySelector("#contractStatusSummary"),
    options: document.querySelector("#contractStatusOptions"),
  },
};

const THEME_STORAGE_KEY = "zoho-theme";
const SAVED_VIEWS_STORAGE_KEY = "zoho-saved-views";
const SORT_LABELS = {
  account_name: "Name",
  updated_desc: "Recently updated",
  owner: "Owner",
  contract_status: "Contract status",
  msa: "MSA",
};

function fmtNumber(value) {
  return Number(value || 0).toLocaleString();
}

function text(value, fallback = "—") {
  return value === undefined || value === null || value === "" ? fallback : value;
}

function compactDateTime(value, fallback = "—") {
  if (!value) return fallback;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function detailedDateTime(value, fallback = "—") {
  if (!value) return fallback;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function debounce(fn, wait = 250) {
  let timeout;
  return (...args) => {
    clearTimeout(timeout);
    timeout = setTimeout(() => fn(...args), wait);
  };
}

function setTheme(theme, { persist = true } = {}) {
  state.theme = theme;
  document.documentElement.dataset.theme = theme;
  els.themeToggle.textContent = theme === "dark" ? "Light mode" : "Dark mode";
  els.themeToggle.setAttribute("aria-pressed", String(theme === "dark"));
  els.themeToggle.setAttribute("aria-label", theme === "dark" ? "Switch to light mode" : "Switch to dark mode");
  if (persist) {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  }
}

function currentViewSnapshot() {
  return {
    q: state.q,
    accountTypes: [...state.accountTypes],
    states: [...state.states],
    msas: [...state.msas],
    contractStatuses: [...state.contractStatuses],
    sortBy: state.sortBy,
  };
}

function snapshotsMatch(a, b) {
  return JSON.stringify(a) === JSON.stringify(b);
}

function persistSavedViews() {
  localStorage.setItem(SAVED_VIEWS_STORAGE_KEY, JSON.stringify(state.savedViews));
}

function loadSavedViews() {
  try {
    const raw = localStorage.getItem(SAVED_VIEWS_STORAGE_KEY);
    state.savedViews = raw ? JSON.parse(raw) : [];
  } catch {
    state.savedViews = [];
  }
}

function renderSavedViews() {
  if (!state.savedViews.length) {
    els.savedViewsList.innerHTML = `<p class="filter-help">Save a search and filter combination to reuse it later.</p>`;
    return;
  }

  const current = currentViewSnapshot();
  els.savedViewsList.innerHTML = state.savedViews
    .map((view) => {
      const active = snapshotsMatch(view.filters, current);
      const chips = [
        view.filters.q ? `Search: ${view.filters.q}` : "",
        view.filters.accountTypes.length ? `${view.filters.accountTypes.length} type${view.filters.accountTypes.length === 1 ? "" : "s"}` : "",
        view.filters.states.length ? `${view.filters.states.length} state${view.filters.states.length === 1 ? "" : "s"}` : "",
        view.filters.msas.length ? `${view.filters.msas.length} MSA${view.filters.msas.length === 1 ? "" : "s"}` : "",
        view.filters.contractStatuses.length ? `${view.filters.contractStatuses.length} status${view.filters.contractStatuses.length === 1 ? "" : "es"}` : "",
      ].filter(Boolean).join(" · ");
      return `
        <article class="saved-view ${active ? "active" : ""}">
          <button class="saved-view-apply" type="button" data-saved-view-id="${escapeHtml(view.id)}">
            <strong>${escapeHtml(view.name)}</strong>
            <span>${escapeHtml(chips || "No filters")}</span>
          </button>
          <button class="saved-view-delete" type="button" data-saved-view-delete="${escapeHtml(view.id)}" aria-label="Delete saved view ${escapeHtml(view.name)}">Delete</button>
        </article>
      `;
    })
    .join("");

  document.querySelectorAll("[data-saved-view-id]").forEach((button) => {
    button.addEventListener("click", () => {
      applySavedView(button.dataset.savedViewId);
    });
  });

  document.querySelectorAll("[data-saved-view-delete]").forEach((button) => {
    button.addEventListener("click", () => {
      deleteSavedView(button.dataset.savedViewDelete);
    });
  });
}

function applySavedView(id) {
  const view = state.savedViews.find((item) => item.id === id);
  if (!view) return;
  state.q = view.filters.q || "";
  state.accountTypes = [...(view.filters.accountTypes || [])];
  state.states = [...(view.filters.states || [])];
  state.msas = [...(view.filters.msas || [])];
  state.contractStatuses = [...(view.filters.contractStatuses || [])];
  state.sortBy = view.filters.sortBy || "account_name";
  state.offset = 0;
  els.search.value = state.q;
  els.sortAccounts.value = state.sortBy;
  updateExportLink();
  renderSavedViews();
  Promise.all([loadFilters(), loadAccounts()]);
}

function deleteSavedView(id) {
  state.savedViews = state.savedViews.filter((view) => view.id !== id);
  persistSavedViews();
  renderSavedViews();
}

function saveCurrentView() {
  const name = els.savedViewName.value.trim();
  if (!name) {
    els.savedViewName.focus();
    return;
  }
  const existing = state.savedViews.find((view) => view.name.toLowerCase() === name.toLowerCase());
  const nextView = {
    id: existing?.id || `view_${Date.now()}`,
    name,
    filters: currentViewSnapshot(),
  };
  if (existing) {
    state.savedViews = state.savedViews.map((view) => (view.id === existing.id ? nextView : view));
  } else {
    state.savedViews = [nextView, ...state.savedViews];
  }
  persistSavedViews();
  els.savedViewName.value = "";
  renderSavedViews();
}

function initTheme() {
  setTheme(state.theme, { persist: false });
  const media = window.matchMedia("(prefers-color-scheme: dark)");
  const applySystemTheme = (event) => {
    if (localStorage.getItem(THEME_STORAGE_KEY)) return;
    setTheme(event.matches ? "dark" : "light", { persist: false });
  };
  if (typeof media.addEventListener === "function") {
    media.addEventListener("change", applySystemTheme);
  } else if (typeof media.addListener === "function") {
    media.addListener(applySystemTheme);
  }
}

async function api(path) {
  const response = await fetch(path);
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || "Request failed");
  }
  return payload;
}

function accountParams() {
  const params = new URLSearchParams();
  if (state.q) params.set("q", state.q);
  state.accountTypes.forEach((value) => params.append("account_type", value));
  state.states.forEach((value) => params.append("state", value));
  state.msas.forEach((msa) => params.append("msa", msa));
  state.contractStatuses.forEach((value) => params.append("contract_status", value));
  params.set("sort", state.sortBy);
  params.set("offset", String(state.offset));
  params.set("limit", String(state.limit));
  return params;
}

function updateExportLink() {
  const params = accountParams();
  params.delete("offset");
  params.delete("limit");
  const query = params.toString();
  els.exportExcel.href = query ? `/export/accounts.xlsx?${query}` : "/export/accounts.xlsx";
  els.exportContacts.href = query ? `/export/contacts.xlsx?${query}` : "/export/contacts.xlsx";
}

function clearFilterValue(key, value = null) {
  if (key === "q") {
    state.q = "";
    els.search.value = "";
    return;
  }

  if (key === "sortBy") {
    state.sortBy = "account_name";
    els.sortAccounts.value = state.sortBy;
    return;
  }

  if (!Array.isArray(state[key])) return;
  state[key] = value === null ? [] : state[key].filter((item) => item !== value);
}

function currentViewBadges() {
  const badges = [];

  if (state.q) {
    badges.push({ label: `Search: ${state.q}`, key: "q" });
  }

  [
    ["accountTypes", "Type"],
    ["states", "State"],
    ["msas", "MSA"],
    ["contractStatuses", "Contract"],
  ].forEach(([key, label]) => {
    state[key].forEach((value) => {
      badges.push({ label: `${label}: ${value}`, key, value });
    });
  });

  if (state.sortBy !== "account_name") {
    badges.push({ label: `Sort: ${SORT_LABELS[state.sortBy] || state.sortBy}`, key: "sortBy" });
  }

  return badges;
}

function renderCurrentView() {
  const badges = currentViewBadges();
  if (!badges.length) {
    els.currentView.innerHTML = "";
    els.currentView.hidden = true;
    return;
  }

  els.currentView.hidden = false;
  els.currentView.innerHTML = `
    <div class="current-view-head">
      <strong>Current view</strong>
      <button type="button" class="text-button" data-clear-view="all">Clear all</button>
    </div>
    <div class="current-view-badges">
      ${badges
        .map(
          (badge) => `
            <button
              type="button"
              class="current-view-badge"
              data-clear-key="${escapeHtml(badge.key)}"
              data-clear-value="${escapeHtml(badge.value || "")}"
              aria-label="Remove ${escapeHtml(badge.label)}"
            >
              <span>${escapeHtml(badge.label)}</span>
              <strong>×</strong>
            </button>
          `,
        )
        .join("")}
    </div>
  `;

  els.currentView.querySelector('[data-clear-view="all"]')?.addEventListener("click", () => {
    els.search.value = "";
    Object.assign(state, {
      q: "",
      accountTypes: [],
      states: [],
      msas: [],
      contractStatuses: [],
      sortBy: "account_name",
      offset: 0,
    });
    els.sortAccounts.value = state.sortBy;
    resetAndLoad();
  });

  els.currentView.querySelectorAll("[data-clear-key]").forEach((button) => {
    button.addEventListener("click", () => {
      clearFilterValue(button.dataset.clearKey, button.dataset.clearValue || null);
      resetAndLoad();
    });
  });
}

async function loadMetadata() {
  const meta = await api("/api/metadata");
  const labels = [
    ["accounts", "Accounts"],
    ["contacts", "Contacts"],
    ["deals", "Deals"],
    ["cases", "Cases"],
    ["accountNotes", "Acct notes"],
    ["msaMappings", "MSAs"],
  ];
  els.metrics.innerHTML = labels
    .map(([key, label]) => `<div class="metric"><strong>${fmtNumber(meta.kpis[key])}</strong><span>${label}</span></div>`)
    .join("");
}

function filterParams(field) {
  const params = accountParams();
  params.delete("offset");
  params.delete("limit");
  params.set("field", field);
  return params;
}

function filterSummary(field) {
  const config = filterConfig[field];
  const values = state[config.key];
  if (!values.length) return config.emptyLabel;
  if (values.length <= 2) return values.join(", ");
  return `${values.length} ${config.label} selected`;
}

function renderFilter(field, values) {
  const config = filterConfig[field];
  const selected = new Set(state[config.key]);
  config.summary.textContent = filterSummary(field);
  config.options.innerHTML = values.length
    ? values
      .map((item) => `
        <label class="filter-option">
          <input
            type="checkbox"
            data-filter="${escapeHtml(field)}"
            value="${escapeHtml(item.value)}"
            ${selected.has(item.value) ? "checked" : ""}
          />
          <span class="filter-option-text">${escapeHtml(item.value)}</span>
          <span class="filter-option-count">${fmtNumber(item.count)}</span>
        </label>
      `)
      .join("")
    : `<p class="filter-empty">No matching options</p>`;

  config.options.querySelectorAll('input[type="checkbox"]').forEach((input) => {
    input.addEventListener("change", () => {
      const current = new Set(state[config.key]);
      if (input.checked) current.add(input.value);
      else current.delete(input.value);
      state[config.key] = Array.from(current).sort((a, b) => a.localeCompare(b));
      resetAndLoad();
    });
  });
}

async function loadFilters() {
  const fields = Object.keys(filterConfig);
  const payloads = await Promise.all(
    fields.map((field) => api(`/api/filter-values?${filterParams(field)}`)),
  );
  fields.forEach((field, index) => {
    renderFilter(field, payloads[index].values);
  });
}

async function loadAccounts({ append = false } = {}) {
  const payload = await api(`/api/accounts?${accountParams()}`);
  state.total = payload.total;
  state.accounts = append ? state.accounts.concat(payload.accounts) : payload.accounts;
  renderAccounts();
}

function chipsFor(account) {
  return [account.type, account.facilityType, account.state, account.msa, account.contractStatus, account.products]
    .filter(Boolean)
    .slice(0, 4);
}

function renderAccounts() {
  els.resultCount.textContent = `${fmtNumber(state.total)} accounts`;
  const shown = Math.min(state.offset + state.limit, state.total);
  els.resultMeta.textContent = state.total ? `Showing ${fmtNumber(shown)}` : "";
  els.loadMore.style.display = shown < state.total ? "block" : "none";
  renderCurrentView();

  if (!state.accounts.length) {
    els.accountList.innerHTML = `<div class="empty-state"><h2>No matches</h2><p>Try a broader search or clear a filter.</p></div>`;
    return;
  }

  els.accountList.innerHTML = state.accounts
    .map((account) => {
      const location = [account.city, account.state].filter(Boolean).join(", ");
      const contract = [account.contractStatus, account.contractType].filter(Boolean).join(" · ");
      const chips = chipsFor(account).map((chip) => `<span class="chip">${escapeHtml(chip)}</span>`).join("");
      const metaItems = [
        ["Owner", account.owner],
        ["Contract", contract],
        ["Updated", compactDateTime(account.modifiedTime)],
        ["Beds", account.beds],
      ]
        .filter(([, value]) => hasValue(value))
        .map(([label, value]) => `
          <div class="account-meta-item">
            <span>${escapeHtml(label)}</span>
            <strong>${escapeHtml(text(value))}</strong>
          </div>
        `)
        .join("");
      return `
        <button class="account-row ${account.id === state.selectedId ? "active" : ""}" type="button" data-id="${escapeHtml(account.id)}">
          <div class="account-row-main">
            <strong>${escapeHtml(text(account.name, "Unnamed account"))}</strong>
            <span class="muted">${escapeHtml(text(location, "No location"))}</span>
          </div>
          <span class="chips">${chips}</span>
          <div class="account-row-meta">${metaItems}</div>
        </button>
      `;
    })
    .join("");

  document.querySelectorAll(".account-row").forEach((button) => {
    button.addEventListener("click", () => selectAccount(button.dataset.id));
  });
}

async function selectAccount(id) {
  state.selectedId = id;
  state.timelineType = "all";
  state.timelineQuery = "";
  state.focusedRecord = null;
  state.attachments = null;
  state.attachmentsError = "";
  state.attachmentsLoading = true;
  state.attachmentCategory = "all";
  state.attachmentModule = "all";
  state.attachmentAvailability = "all";
  state.attachmentQuery = "";
  state.detail = await api(`/api/accounts/${encodeURIComponent(id)}`);
  renderAccounts();
  renderDetail();
  try {
    state.attachments = await api(`/api/accounts/${encodeURIComponent(id)}/attachments`);
  } catch (error) {
    state.attachmentsError = error.message;
  } finally {
    state.attachmentsLoading = false;
    renderDetail();
  }
}

function field(label, value) {
  return `<div class="field"><span>${escapeHtml(label)}</span><strong>${escapeHtml(text(value))}</strong></div>`;
}

function hasValue(value) {
  return value !== undefined && value !== null && String(value).trim() !== "";
}

function disclosureSection(title, body, { open = false, meta = "" } = {}) {
  return `
    <details class="detail-section" ${open ? "open" : ""}>
      <summary>
        <span>${escapeHtml(title)}</span>
        ${meta ? `<small>${escapeHtml(meta)}</small>` : ""}
      </summary>
      <div class="detail-section-body">${body}</div>
    </details>
  `;
}

function bestContactMethod(contact) {
  if (hasValue(contact?.email)) return contact.email;
  if (hasValue(contact?.phone)) return contact.phone;
  if (hasValue(contact?.mobile)) return contact.mobile;
  return "";
}

function contactPriority(contact) {
  const haystack = [contact.title, contact.contact_type, contact.role].filter(Boolean).join(" ").toLowerCase();
  const name = String(contact.contact_name || "").toLowerCase();
  let score = 0;

  if (haystack.includes("main contact") || haystack.includes("primary")) score += 100;
  if (["sched", "admission", "intake", "referral", "liaison", "coordinator"].some((token) => haystack.includes(token))) score += 40;
  if (hasValue(contact.email)) score += 18;
  if (hasValue(contact.phone) || hasValue(contact.mobile)) score += 12;
  if (hasValue(contact.title)) score += 8;
  if (hasValue(contact.contact_type)) score += 4;
  if (!name.includes("inc.") && !name.includes("llc") && !name.includes("center")) score += 6;

  return score;
}

function contactSignal(contacts) {
  const reachable = contacts.filter((contact) => hasValue(contact.email) || hasValue(contact.phone) || hasValue(contact.mobile));
  const best = [...reachable].sort((a, b) => contactPriority(b) - contactPriority(a) || String(a.contact_name || "").localeCompare(String(b.contact_name || "")))[0]
    || [...contacts].sort((a, b) => contactPriority(b) - contactPriority(a) || String(a.contact_name || "").localeCompare(String(b.contact_name || "")))[0]
    || null;
  return {
    value: reachable.length ? `${reachable.length} reachable` : "No reachable contacts",
    meta: `${contacts.length} total`,
    detail: best
      ? `${text(best.contact_name, "Unnamed contact")} · ${text(best.title || best.contact_type || best.role, "No role")} · ${text(bestContactMethod(best), "No direct method")}`
      : "No contacts on this account",
  };
}

function latestActivitySignal(account, related) {
  const events = [
    { label: "Account updated", at: account.modifiedTime || account.createdTime, detail: text(account.owner, "Account owner") },
    ...related.notes.map((note) => ({
      label: "Latest note",
      at: note.modified_time || note.created_time,
      detail: text(note.note_owner || note.note_title, "Note activity"),
    })),
    ...related.deals.map((deal) => ({
      label: "Latest deal",
      at: deal.modified_time,
      detail: text(deal.deal_name, "Deal activity"),
    })),
    ...related.cases.map((item) => ({
      label: "Latest case",
      at: item.modified_time,
      detail: text(item.subject || item.case_number, "Case activity"),
    })),
    ...related.contacts.map((contact) => ({
      label: "Latest contact update",
      at: contact.modified_time,
      detail: text(contact.contact_name, "Contact activity"),
    })),
  ].filter((event) => hasValue(event.at));

  if (!events.length) {
    return {
      value: "No recent activity",
      meta: "No dated events",
      detail: "This record has no activity timestamps yet",
    };
  }

  events.sort((a, b) => new Date(b.at) - new Date(a.at));
  const latest = events[0];
  return {
    value: compactDateTime(latest.at),
    meta: latest.label,
    detail: latest.detail,
  };
}

function documentSignal(account) {
  if (state.attachmentsLoading) {
    return {
      value: "Loading documents…",
      meta: "Checking attachments",
      detail: "Attachment signals update after the record loads",
    };
  }

  if (state.attachmentsError) {
    return {
      value: "Documents unavailable",
      meta: "Attachment lookup failed",
      detail: state.attachmentsError,
    };
  }

  const attachments = state.attachments?.attachments || [];
  const contractLike = attachments.filter((attachment) => /contract|agreement|signed|w-?9/i.test(attachment.original_filename || ""));
  return {
    value: attachments.length ? `${attachments.length} files` : "No documents",
    meta: contractLike.length ? `${contractLike.length} contract-related` : "No contract-like docs found",
    detail: hasValue(account.contractStatus)
      ? `Status: ${account.contractStatus}`
      : "No contract status on this account",
  };
}

function attentionSignal(account, related) {
  const issues = [];
  if (!hasValue(account.billingEmail) && !hasValue(account.phone) && !hasValue(account.billingPhone)) {
    issues.push("no facility contact method");
  }
  if (!related.contacts.some((contact) => hasValue(contact.email) || hasValue(contact.phone) || hasValue(contact.mobile))) {
    issues.push("no reachable contact");
  }
  if (!hasValue(account.contractStatus)) {
    issues.push("missing contract status");
  }
  if (!related.notes.length) {
    issues.push("no notes");
  }

  return {
    value: issues.length ? `${issues.length} attention item${issues.length === 1 ? "" : "s"}` : "Looks complete",
    meta: issues.length ? "Needs review" : "No obvious blockers",
    detail: issues.length ? issues.slice(0, 3).join(" · ") : "Core relationship signals are present",
    urgent: issues.length > 0,
  };
}

function summaryCard(title, signal, { tone = "default" } = {}) {
  return `
    <section class="summary-card ${tone === "warning" ? "warning" : ""}">
      <span>${escapeHtml(title)}</span>
      <strong>${escapeHtml(signal.value)}</strong>
      <small>${escapeHtml(signal.meta)}</small>
      <p>${escapeHtml(signal.detail)}</p>
    </section>
  `;
}

function summaryHeader(account, related) {
  const cards = [
    summaryCard("Contact coverage", contactSignal(related.contacts)),
    summaryCard("Recent activity", latestActivitySignal(account, related)),
    summaryCard("Documents", documentSignal(account)),
  ];
  const attention = attentionSignal(account, related);
  cards.push(summaryCard("Attention", attention, { tone: attention.urgent ? "warning" : "default" }));
  return `<section class="summary-strip">${cards.join("")}</section>`;
}

function feeValue(fee) {
  const raw = text(fee.value);
  if (raw === "—" || fee.kind !== "currency") {
    return raw;
  }
  const cleaned = String(raw).replace(/[$,]/g, "").trim();
  if (!cleaned || Number.isNaN(Number(cleaned))) {
    return raw;
  }
  return Number(cleaned).toLocaleString(undefined, {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function fmtFileSize(value) {
  const size = Number(value || 0);
  if (!size) return "—";
  const units = ["B", "KB", "MB", "GB"];
  let scaled = size;
  let unit = units[0];
  for (let index = 1; index < units.length && scaled >= 1024; index += 1) {
    scaled /= 1024;
    unit = units[index];
  }
  return unit === "B" ? `${size} B` : `${scaled.toFixed(1)} ${unit}`;
}

function attachmentKind(attachment) {
  const extension = (attachment.file_extension || "").replace(".", "").toUpperCase();
  return extension || "FILE";
}

function attachmentDateValue(attachment) {
  const raw = attachment.modified_time || attachment.created_time || "";
  const parsed = Date.parse(raw);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function attachmentCategory(attachment) {
  const filename = String(attachment.original_filename || "").toLowerCase();
  const module = String(attachment.parent_module || "").toLowerCase();
  if (/w-?9|tax/i.test(filename)) return "w9";
  if (/signed|completion_certificate|completion certificate|executed/i.test(filename)) return "signed";
  if (/contract|agreement|msa|service agreement|client service/i.test(filename) || module.includes("zohosign")) return "contract";
  return "other";
}

function attachmentCategoryLabel(category) {
  return {
    all: "All documents",
    contract: "Contracts",
    signed: "Signed docs",
    w9: "W-9 / tax",
    other: "Other docs",
  }[category] || "Documents";
}

function attachmentMatchesFilters(attachment) {
  const query = state.attachmentQuery.trim().toLowerCase();
  if (state.attachmentCategory !== "all" && attachmentCategory(attachment) !== state.attachmentCategory) {
    return false;
  }
  if (state.attachmentModule !== "all" && (attachment.parent_module || "Unknown") !== state.attachmentModule) {
    return false;
  }
  if (state.attachmentAvailability !== "all" && attachment.availability !== state.attachmentAvailability) {
    return false;
  }
  if (!query) return true;
  const haystack = [
    attachment.original_filename,
    attachment.parent_module,
    attachment.mapping_confidence,
    attachment.record_status,
  ].filter(Boolean).join(" ").toLowerCase();
  return haystack.includes(query);
}

function attachmentFiltersBar(attachments) {
  const moduleCounts = attachments.reduce((counts, attachment) => {
    const module = attachment.parent_module || "Unknown";
    counts[module] = (counts[module] || 0) + 1;
    return counts;
  }, {});
  const availabilityCounts = attachments.reduce((counts, attachment) => {
    counts[attachment.availability] = (counts[attachment.availability] || 0) + 1;
    return counts;
  }, {});
  const categoryCounts = attachments.reduce((counts, attachment) => {
    const category = attachmentCategory(attachment);
    counts[category] = (counts[category] || 0) + 1;
    return counts;
  }, {});
  const categoryOptions = ["all", "contract", "signed", "w9", "other"];
  const availabilityOptions = [
    ["all", "All availability"],
    ["available", "Available only"],
    ["unmatched", "Missing matches"],
  ];
  const modules = Object.keys(moduleCounts).sort((a, b) => moduleCounts[b] - moduleCounts[a] || a.localeCompare(b));

  return `
    <section class="attachment-filters">
      <div class="attachment-filter-row attachment-filter-row-categories">
        ${categoryOptions.map((category) => {
          const count = category === "all" ? attachments.length : (categoryCounts[category] || 0);
          return `<button class="attachment-pill ${state.attachmentCategory === category ? "active" : ""}" type="button" data-attachment-category="${category}">${escapeHtml(attachmentCategoryLabel(category))} <span>${fmtNumber(count)}</span></button>`;
        }).join("")}
      </div>
      <div class="attachment-filter-row">
        <label class="attachment-select">
          <span>Source</span>
          <select id="attachmentModuleFilter">
            <option value="all">All modules</option>
            ${modules.map((module) => `<option value="${escapeHtml(module)}" ${state.attachmentModule === module ? "selected" : ""}>${escapeHtml(module)} (${fmtNumber(moduleCounts[module])})</option>`).join("")}
          </select>
        </label>
        <label class="attachment-select">
          <span>Availability</span>
          <select id="attachmentAvailabilityFilter">
            ${availabilityOptions.map(([value, label]) => {
              const count = value === "all" ? attachments.length : (availabilityCounts[value] || 0);
              return `<option value="${value}" ${state.attachmentAvailability === value ? "selected" : ""}>${escapeHtml(label)} (${fmtNumber(count)})</option>`;
            }).join("")}
          </select>
        </label>
        <label class="attachment-search">
          <span>Find document</span>
          <input id="attachmentQuery" type="search" value="${escapeHtml(state.attachmentQuery)}" placeholder="Filename, module, status" />
        </label>
      </div>
    </section>
  `;
}

function priorityAttachmentsSection(attachments) {
  const priorityCategories = ["signed", "contract", "w9"];
  const groups = priorityCategories
    .map((category) => ({
      category,
      label: attachmentCategoryLabel(category),
      items: attachments
        .filter((attachment) => attachmentCategory(attachment) === category)
        .sort((a, b) => attachmentDateValue(b) - attachmentDateValue(a) || String(a.original_filename || "").localeCompare(String(b.original_filename || "")))
        .slice(0, 6),
    }))
    .filter((group) => group.items.length);

  if (!groups.length) return "";

  return `
    <section class="attachment-priority">
      <header>
        <h3>Priority documents</h3>
        <p class="muted">Contract-related files surfaced first so common document hunts take fewer clicks.</p>
      </header>
      <div class="attachment-priority-groups">
        ${groups.map((group) => `
          <section class="attachment-priority-group">
            <div class="attachment-priority-head">
              <strong>${escapeHtml(group.label)}</strong>
              <span>${fmtNumber(group.items.length)}</span>
            </div>
            <div class="attachment-priority-list">
              ${group.items.map((attachment) => {
                const openUrl = `${attachment.download_url}?disposition=inline`;
                const canDownload = attachment.availability === "available";
                return `
                  <article class="attachment-row compact">
                    <div class="attachment-main">
                      <strong title="${escapeHtml(attachment.original_filename)}">${escapeHtml(text(attachment.original_filename, "Unnamed attachment"))}</strong>
                      <span>${escapeHtml(attachmentKind(attachment))} · ${escapeHtml(text(attachment.modified_time || attachment.created_time, "No date"))}</span>
                    </div>
                    <div class="attachment-meta">
                      <span class="status ${escapeHtml(attachment.availability)}">${escapeHtml(availabilityLabel(attachment.availability))}</span>
                    </div>
                    <div class="attachment-actions">
                      ${canDownload ? `<a href="${escapeHtml(openUrl)}" target="_blank" rel="noreferrer">Open</a><a href="${escapeHtml(attachment.download_url)}">Download</a>` : `<span class="muted">No file</span>`}
                    </div>
                  </article>
                `;
              }).join("")}
            </div>
          </section>
        `).join("")}
      </div>
    </section>
  `;
}

function serviceFeesPanel(fees) {
  const configuredCount = fees.filter((fee) => hasValue(fee.value)).length;
  const midpoint = Math.ceil(fees.length / 2);
  const columns = [fees.slice(0, midpoint), fees.slice(midpoint)];
  return disclosureSection(
    "Service fee schedule",
    `
      <section class="service-fees" aria-label="Service fees">
        <header>
          <h3>RCM Summary - Service Fees</h3>
        </header>
        <div class="service-fee-grid">
          ${columns
            .map(
              (column) => `
                <div class="service-fee-column">
                  ${column
                    .map(
                      (fee) => `
                        <div class="fee-row">
                          <span>${escapeHtml(fee.label)}</span>
                          <strong>${escapeHtml(feeValue(fee))}</strong>
                        </div>
                      `,
                    )
                    .join("")}
                </div>
              `,
            )
            .join("")}
        </div>
      </section>
    `,
    {
      meta: configuredCount ? `${configuredCount} configured` : "No configured fees",
    },
  );
}

function isFocusedRecord(tab, recordId) {
  return Boolean(
    state.focusedRecord
    && state.focusedRecord.tab === tab
    && state.focusedRecord.id === recordId,
  );
}

function table(rows, columns, { tab = "", rowIdKey = "record_id" } = {}) {
  if (!rows.length) {
    return `<p class="muted">No records found.</p>`;
  }
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>${columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("")}</tr>
        </thead>
        <tbody>
          ${rows
            .map(
              (row) => `
                <tr data-record-id="${escapeHtml(row[rowIdKey] || "")}" class="${isFocusedRecord(tab, row[rowIdKey]) ? "source-focus" : ""}">
                  ${columns.map((column) => `<td>${escapeHtml(text(row[column.key], ""))}</td>`).join("")}
                </tr>
              `,
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;
}

function notesPanel(notes) {
  if (!notes.length) {
    return `<p class="muted">No notes found.</p>`;
  }
  return notes
    .map(
      (note) => `
        <section class="note ${isFocusedRecord("notes", note.record_id) ? "source-focus" : ""}" data-record-id="${escapeHtml(note.record_id || "")}">
          <header>
            <strong>${escapeHtml(text(note.note_title, "Untitled note"))}</strong>
            <span>${escapeHtml(text(note.modified_time || note.created_time, ""))}</span>
          </header>
          <p>${escapeHtml(text(note.note_content, ""))}</p>
          <p class="muted">${escapeHtml([note.parent_module, note.note_owner].filter(Boolean).join(" · "))}</p>
        </section>
      `,
    )
    .join("");
}

function rawPanel(raw) {
  const items = Object.entries(raw).filter(([, value]) => value !== "");
  return `
    <div class="raw-grid">
      ${items
        .map(([key, value]) => `<div class="raw-item"><span>${escapeHtml(key)}</span><strong>${escapeHtml(value)}</strong></div>`)
        .join("")}
    </div>
  `;
}

function groupedAttachments(attachments) {
  return attachments.reduce((groups, attachment) => {
    const module = attachment.parent_module || "Unknown";
    if (!groups[module]) groups[module] = [];
    groups[module].push(attachment);
    return groups;
  }, {});
}

function availabilityLabel(status) {
  return {
    available: "Available",
    unmatched: "Unmatched",
    missing_zip: "Missing zip",
    missing_zip_entry: "Missing entry",
    bad_zip: "Bad zip",
  }[status] || text(status, "Unknown");
}

function attachmentsPanel() {
  if (state.attachmentsLoading) {
    return `<p class="muted">Loading attachments...</p>`;
  }
  if (state.attachmentsError) {
    return `<div class="empty-state inline"><h2>Attachments unavailable</h2><p>${escapeHtml(state.attachmentsError)}</p></div>`;
  }
  const payload = state.attachments;
  const attachments = payload?.attachments || [];
  if (!attachments.length) {
    return `<p class="muted">No attachments found for this account.</p>`;
  }

  const filteredAttachments = attachments
    .filter(attachmentMatchesFilters)
    .sort((a, b) => attachmentDateValue(b) - attachmentDateValue(a) || String(a.original_filename || "").localeCompare(String(b.original_filename || "")));
  const groups = groupedAttachments(filteredAttachments);
  const modules = Object.keys(groups).sort((a, b) => groups[b].length - groups[a].length || a.localeCompare(b));
  return `
    <div class="attachment-summary">
      <strong>${fmtNumber(payload.total)} attachments</strong>
      <span>${payload.groups.map((group) => `${escapeHtml(group.module)} (${fmtNumber(group.count)})`).join(" · ")}</span>
      <span>${fmtNumber(filteredAttachments.length)} shown</span>
    </div>
    ${attachmentFiltersBar(attachments)}
    ${priorityAttachmentsSection(filteredAttachments)}
    <div class="attachment-groups">
      ${!filteredAttachments.length ? `<div class="empty-state inline"><h2>No matching documents</h2><p>Try a broader attachment filter or clear the search.</p></div>` : ""}
      ${modules
        .map((module, index) => `
          <details class="attachment-group" ${index < 3 ? "open" : ""}>
            <summary>${escapeHtml(module)} <span>${fmtNumber(groups[module].length)}</span></summary>
            <div class="attachment-list">
              ${groups[module]
                .map((attachment) => {
                  const openUrl = `${attachment.download_url}?disposition=inline`;
                  const canDownload = attachment.availability === "available";
                  return `
                    <article class="attachment-row ${isFocusedRecord("attachments", attachment.attachment_record_id) ? "source-focus" : ""}" data-record-id="${escapeHtml(attachment.attachment_record_id || "")}">
                      <div class="attachment-main">
                        <strong title="${escapeHtml(attachment.original_filename)}">${escapeHtml(text(attachment.original_filename, "Unnamed attachment"))}</strong>
                        <span>${escapeHtml(attachmentKind(attachment))} · ${escapeHtml(fmtFileSize(attachment.file_size))} · ${escapeHtml(text(attachment.modified_time || attachment.created_time, "No date"))}</span>
                      </div>
                      <div class="attachment-meta">
                        <span class="status ${escapeHtml(attachment.availability)}">${escapeHtml(availabilityLabel(attachment.availability))}</span>
                        <span>${escapeHtml(text(attachment.mapping_confidence, ""))}</span>
                      </div>
                      <div class="attachment-actions">
                        ${
                          canDownload
                            ? `<a href="${escapeHtml(openUrl)}" target="_blank" rel="noreferrer">Open</a><a href="${escapeHtml(attachment.download_url)}">Download</a>`
                            : `<span class="muted">No file</span>`
                        }
                      </div>
                    </article>
                  `;
                })
                .join("")}
            </div>
          </details>
        `)
        .join("")}
    </div>
  `;
}

const TIMELINE_TYPE_META = {
  account: { label: "Account" },
  note: { label: "Note" },
  deal: { label: "Deal" },
  case: { label: "Case" },
  contact: { label: "Contact" },
  email: { label: "Email" },
  meeting: { label: "Meeting" },
  task: { label: "Task" },
  attachment: { label: "Attachment" },
};

function attachmentTimelineEvents() {
  const attachments = state.attachments?.attachments || [];
  return attachments.map((attachment) => {
    const summary = [
      attachment.parent_module || "",
      attachmentKind(attachment),
      fmtFileSize(attachment.file_size),
    ].filter(Boolean).join(" · ");
    return {
      id: `attachment:${attachment.attachment_record_id}`,
      type: "attachment",
      at: attachment.modified_time || attachment.created_time || "",
      title: text(attachment.original_filename, "Unnamed attachment"),
      summary,
      detail: attachment.mapping_confidence || "",
      badge: attachment.availability === "available" ? "Available" : availabilityLabel(attachment.availability),
      status: attachmentCategoryLabel(attachmentCategory(attachment)),
      sourceId: attachment.attachment_record_id,
      sourceTab: "attachments",
      downloadUrl: attachment.download_url,
      openUrl: `${attachment.download_url}?disposition=inline`,
      availability: attachment.availability,
    };
  });
}

function combinedTimelineEvents() {
  const base = state.detail?.timeline || [];
  const attachments = state.attachmentsLoading || state.attachmentsError ? [] : attachmentTimelineEvents();
  return [...base, ...attachments].sort((a, b) => {
    const left = a.at || "";
    const right = b.at || "";
    if (left !== right) return left < right ? 1 : -1;
    return String(a.title || "").localeCompare(String(b.title || ""));
  });
}

function timelineSourceLabel(tab) {
  return {
    contacts: "Contacts",
    deals: "Deals",
    cases: "Cases",
    notes: "Notes",
    attachments: "Attachments",
  }[tab] || "Details";
}

function jumpToTimelineSource(sourceTab, sourceId) {
  if (!sourceTab || !sourceId) return;
  state.activeTab = sourceTab;
  state.focusedRecord = { tab: sourceTab, id: sourceId };
  renderDetail();
}

function focusSelectedRecord() {
  if (!state.focusedRecord) return;
  const escape = globalThis.CSS?.escape ? globalThis.CSS.escape(state.focusedRecord.id) : state.focusedRecord.id.replace(/"/g, '\\"');
  const target = els.detail.querySelector(`[data-record-id="${escape}"]`);
  if (!target) return;
  target.scrollIntoView({ behavior: "smooth", block: "center" });
}

function timelineFilters(events) {
  const counts = events.reduce((lookup, event) => {
    lookup[event.type] = (lookup[event.type] || 0) + 1;
    return lookup;
  }, {});
  const types = Object.keys(counts).sort((a, b) => counts[b] - counts[a] || a.localeCompare(b));
  return `
    <section class="timeline-toolbar">
      <div class="timeline-pills">
        <button class="timeline-pill ${state.timelineType === "all" ? "active" : ""}" type="button" data-timeline-type="all">All <span>${fmtNumber(events.length)}</span></button>
        ${types.map((type) => `
          <button class="timeline-pill ${state.timelineType === type ? "active" : ""}" type="button" data-timeline-type="${escapeHtml(type)}">${escapeHtml(TIMELINE_TYPE_META[type]?.label || type)} <span>${fmtNumber(counts[type])}</span></button>
        `).join("")}
      </div>
      <label class="timeline-search">
        <span>Find activity</span>
        <input id="timelineQuery" type="search" value="${escapeHtml(state.timelineQuery)}" placeholder="Search title, detail, status" />
      </label>
    </section>
  `;
}

function timelineMatches(event) {
  if (state.timelineType !== "all" && event.type !== state.timelineType) {
    return false;
  }
  const query = state.timelineQuery.trim().toLowerCase();
  if (!query) return true;
  const haystack = [
    event.title,
    event.summary,
    event.detail,
    event.badge,
    event.status,
    event.owner,
  ].filter(Boolean).join(" ").toLowerCase();
  return haystack.includes(query);
}

function timelineDateLabel(value) {
  if (!value) return "Undated";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Undated";
  return date.toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
  });
}

function timelinePanel() {
  const allEvents = combinedTimelineEvents();
  if (!allEvents.length) {
    return `<div class="empty-state inline"><h2>No timeline activity yet</h2><p>This account does not have notes, outreach, meetings, tasks, or document activity to show.</p></div>`;
  }

  const filtered = allEvents.filter(timelineMatches);
  const groups = filtered.reduce((lookup, event) => {
    const bucket = timelineDateLabel(event.at);
    if (!lookup[bucket]) lookup[bucket] = [];
    lookup[bucket].push(event);
    return lookup;
  }, {});

  const orderedGroups = Object.entries(groups);
  return `
    <section class="timeline-summary">
      <strong>${fmtNumber(allEvents.length)} activity items</strong>
      <span>${state.attachmentsLoading ? "Attachments still loading into the timeline…" : "Notes, outreach, operations, and documents in one feed."}</span>
    </section>
    ${timelineFilters(allEvents)}
    ${!filtered.length ? `<div class="empty-state inline"><h2>No matching activity</h2><p>Try a different type or clear the timeline search.</p></div>` : ""}
    <section class="timeline">
      ${orderedGroups.map(([label, events]) => `
        <section class="timeline-group">
          <header class="timeline-group-head">${escapeHtml(label)}</header>
          <div class="timeline-group-items">
            ${events.map((event) => `
              <article class="timeline-item">
                <div class="timeline-item-head">
                  <div class="timeline-item-title">
                    <span class="timeline-type">${escapeHtml(TIMELINE_TYPE_META[event.type]?.label || event.type)}</span>
                    <strong>${escapeHtml(text(event.title, "Untitled activity"))}</strong>
                  </div>
                  <time>${escapeHtml(detailedDateTime(event.at, "Undated"))}</time>
                </div>
                ${event.summary ? `<p class="timeline-summary-line">${escapeHtml(event.summary)}</p>` : ""}
                ${event.detail ? `<p class="timeline-detail-line">${escapeHtml(event.detail)}</p>` : ""}
                ${(event.badge || event.status || event.owner) ? `
                  <div class="timeline-meta">
                    ${event.badge ? `<span class="timeline-badge">${escapeHtml(event.badge)}</span>` : ""}
                    ${event.status ? `<span class="timeline-status">${escapeHtml(event.status)}</span>` : ""}
                    ${event.owner ? `<span>${escapeHtml(event.owner)}</span>` : ""}
                  </div>
                ` : ""}
                <div class="timeline-actions">
                  ${event.sourceTab && event.sourceId ? `<button type="button" class="timeline-action" data-timeline-source-tab="${escapeHtml(event.sourceTab)}" data-timeline-source-id="${escapeHtml(event.sourceId)}">View in ${escapeHtml(timelineSourceLabel(event.sourceTab))}</button>` : ""}
                  ${event.type === "attachment" && event.availability === "available"
                    ? `<a class="timeline-link" href="${escapeHtml(event.openUrl)}" target="_blank" rel="noreferrer">Open</a><a class="timeline-link" href="${escapeHtml(event.downloadUrl)}">Download</a>`
                    : ""}
                </div>
              </article>
            `).join("")}
          </div>
        </section>
      `).join("")}
    </section>
  `;
}

function renderPanel() {
  const related = state.detail.related;
  if (state.activeTab === "timeline") {
    return timelinePanel();
  }
  if (state.activeTab === "contacts") {
    return table(related.contacts, [
      { key: "contact_name", label: "Name" },
      { key: "title", label: "Title" },
      { key: "email", label: "Email" },
      { key: "phone", label: "Phone" },
      { key: "mobile", label: "Mobile" },
      { key: "contact_type", label: "Type" },
    ], { tab: "contacts" });
  }
  if (state.activeTab === "deals") {
    return table(related.deals, [
      { key: "deal_name", label: "Deal" },
      { key: "stage", label: "Stage" },
      { key: "amount", label: "Amount" },
      { key: "closing_date", label: "Close date" },
      { key: "contact_name", label: "Contact" },
    ], { tab: "deals" });
  }
  if (state.activeTab === "cases") {
    return table(related.cases, [
      { key: "case_number", label: "Case" },
      { key: "status", label: "Status" },
      { key: "case_origin", label: "Origin" },
      { key: "subject", label: "Subject" },
      { key: "description", label: "Description" },
    ], { tab: "cases" });
  }
  if (state.activeTab === "notes") {
    return notesPanel(related.notes);
  }
  if (state.activeTab === "attachments") {
    return attachmentsPanel();
  }
  return rawPanel(state.detail.account.raw);
}

function renderDetail() {
  if (!state.detail) return;
  const account = state.detail.account;
  const address = [
    account.address.street,
    account.address.street2,
    [account.address.city, account.address.state, account.address.zip].filter(Boolean).join(", "),
  ].filter(Boolean).join(" · ");
  const tabs = [
    ["timeline", `Timeline (${fmtNumber((state.detail.timeline?.length || 0) + (state.attachments?.total || 0))})`],
    ["contacts", `Contacts (${state.detail.related.contacts.length})`],
    ["deals", `Deals (${state.detail.related.deals.length})`],
    ["cases", `Cases (${state.detail.related.cases.length})`],
    ["notes", `Notes (${state.detail.related.notes.length})`],
    ["attachments", `Attachments (${state.attachments?.total ?? (state.attachmentsLoading ? "..." : 0)})`],
    ["raw", "Raw fields"],
  ];
  const primaryFields = [
    field("Owner", account.owner),
    field("Phone", account.phone || account.billingPhone),
    field("Billing email", account.billingEmail),
    field("Contract", [account.contractStatus, account.contractType].filter(Boolean).join(" · ")),
    field("MSA", account.msa),
    field("Territories", account.territories),
  ];
  const secondaryFields = [
    hasValue(account.totalBeds || account.certifiedMedicareBeds)
      ? field("Beds", account.totalBeds || account.certifiedMedicareBeds)
      : "",
    hasValue(account.chainAccount) ? field("Chain", account.chainAccount) : "",
    hasValue(account.facilityNpi) ? field("NPI", account.facilityNpi) : "",
  ].filter(Boolean).join("");
  const overviewMeta = [
    `${state.detail.related.contacts.length} contacts`,
    `${state.detail.related.notes.length} notes`,
    `${state.attachments?.total ?? (state.attachmentsLoading ? "..." : 0)} attachments`,
  ].join(" · ");

  els.detail.innerHTML = `
    <header class="detail-header">
      <h2>${escapeHtml(text(account.name, "Unnamed account"))}</h2>
      <div class="chips">
        ${[account.type, account.facilityType, account.contractStatus, account.productsAvailable].filter(Boolean).map((chip) => `<span class="chip">${escapeHtml(chip)}</span>`).join("")}
      </div>
      <p class="muted">${escapeHtml(text(address, "No address"))}</p>
      <p class="detail-meta">${escapeHtml(overviewMeta)}</p>
    </header>

    ${summaryHeader(account, state.detail.related)}

    <section class="detail-grid detail-grid-primary">
      ${primaryFields.join("")}
    </section>

    ${secondaryFields
      ? disclosureSection(
        "More facility details",
        `<section class="detail-grid detail-grid-secondary">${secondaryFields}</section>`,
        { meta: "Beds, chain, identifiers" },
      )
      : ""}

    ${serviceFeesPanel(account.serviceFees || [])}

    <nav class="tabs">
      ${tabs.map(([key, label]) => `<button class="tab ${state.activeTab === key ? "active" : ""}" type="button" data-tab="${key}">${escapeHtml(label)}</button>`).join("")}
    </nav>

    <section class="panel">${renderPanel()}</section>
  `;

  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeTab = button.dataset.tab;
      state.focusedRecord = null;
      renderDetail();
    });
  });

  document.querySelectorAll("[data-timeline-type]").forEach((button) => {
    button.addEventListener("click", () => {
      state.timelineType = button.dataset.timelineType;
      renderDetail();
    });
  });

  document.querySelectorAll("[data-timeline-source-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      jumpToTimelineSource(button.dataset.timelineSourceTab, button.dataset.timelineSourceId);
    });
  });

  const timelineQuery = document.querySelector("#timelineQuery");
  if (timelineQuery) {
    timelineQuery.addEventListener("input", debounce(() => {
      state.timelineQuery = timelineQuery.value;
      renderDetail();
    }, 150));
  }

  document.querySelectorAll("[data-attachment-category]").forEach((button) => {
    button.addEventListener("click", () => {
      state.attachmentCategory = button.dataset.attachmentCategory;
      renderDetail();
    });
  });

  const moduleFilter = document.querySelector("#attachmentModuleFilter");
  if (moduleFilter) {
    moduleFilter.addEventListener("change", () => {
      state.attachmentModule = moduleFilter.value;
      renderDetail();
    });
  }

  const availabilityFilter = document.querySelector("#attachmentAvailabilityFilter");
  if (availabilityFilter) {
    availabilityFilter.addEventListener("change", () => {
      state.attachmentAvailability = availabilityFilter.value;
      renderDetail();
    });
  }

  const attachmentQuery = document.querySelector("#attachmentQuery");
  if (attachmentQuery) {
    attachmentQuery.addEventListener("input", debounce(() => {
      state.attachmentQuery = attachmentQuery.value;
      renderDetail();
    }, 150));
  }

  if (state.focusedRecord && state.activeTab === state.focusedRecord.tab) {
    requestAnimationFrame(() => {
      focusSelectedRecord();
    });
  }
}

function resetAndLoad() {
  state.offset = 0;
  updateExportLink();
  renderSavedViews();
  Promise.all([loadFilters(), loadAccounts()]);
}

els.search.addEventListener("input", debounce(() => {
  state.q = els.search.value.trim();
  resetAndLoad();
}));

els.clearFilters.addEventListener("click", () => {
  els.search.value = "";
  Object.assign(state, { q: "", accountTypes: [], states: [], msas: [], contractStatuses: [], sortBy: "account_name", offset: 0 });
  els.sortAccounts.value = state.sortBy;
  updateExportLink();
  renderSavedViews();
  Promise.all([loadFilters(), loadAccounts()]);
});

els.sortAccounts.addEventListener("change", () => {
  state.sortBy = els.sortAccounts.value;
  resetAndLoad();
});

els.saveView.addEventListener("click", () => {
  saveCurrentView();
});

els.savedViewName.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    saveCurrentView();
  }
});

els.themeToggle.addEventListener("click", () => {
  setTheme(state.theme === "dark" ? "light" : "dark");
});

els.loadMore.addEventListener("click", () => {
  state.offset += state.limit;
  loadAccounts({ append: true });
});

async function init() {
  try {
    initTheme();
    loadSavedViews();
    renderSavedViews();
    els.sortAccounts.value = state.sortBy;
    await loadMetadata();
    await loadFilters();
    updateExportLink();
    await loadAccounts();
  } catch (error) {
    els.detail.innerHTML = `<div class="empty-state"><h2>Something needs attention</h2><p>${escapeHtml(error.message)}</p></div>`;
  }
}

init();
