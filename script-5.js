/* =========================================================================
   WORKSHOP2U — FRONT-END LOGIC (full feature set)
   Talks to a Google Apps Script Web App (Code.gs). Set CONFIG.API_URL below.
   ========================================================================= */

const CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbzQO0bFEfT3dNEhX2ijR9hytr78pfTEC5k64Z4m0Kpnj_TufQzCohKLb0lkH3GQSECu/exec",
  WHATSAPP_NUMBER: "60137137100", // digits only, country code first — used by the WhatsApp button
  // One Google review link per workshop, shown after a customer submits
  // their internal review. Get yours from business.google.com → your
  // listing → "Get more reviews" (gives a short g.page/r/... link), or via
  // https://search.google.com/local/writereview?placeid=YOUR_PLACE_ID
  // using Google's Place ID Finder if you don't have Business Profile
  // access. Leave a workshop blank/unset to simply not show the button for it.
  GOOGLE_REVIEW_LINKS: {
    "Melaka": "https://search.google.com/local/writereview?placeid=ChIJx12eMwDl0TERGeSikxRbNjU",
    "Negeri Sembilan": "https://search.google.com/local/writereview?placeid=ChIJnT6mkSTnzTERheFsqeYbjNA",
    "Johor": "https://search.google.com/local/writereview?placeid=ChIJazvBzr5x2jERYqMdQsUujLU"
  }
};

/* Escapes user-submitted free text before it's inserted into the page —
   applied to fields that come from the public (the booking form, review
   comments) so a visitor can't inject HTML/script into what other people
   or staff see rendered back. */
/* Draws a 0–5 rating as always-five stars: full (amber), half (amber/grey
   split) or empty (grey). Ratings are rounded to the nearest half, so
   4 → ★★★★☆, 4.5 → ★★★★ + a half star, 5 → ★★★★★. The colouring is done in
   CSS (.star-display in style.css). */
function renderStars(rating) {
  const r = Math.min(5, Math.max(0, Math.round((Number(rating) || 0) * 2) / 2));
  const full = Math.floor(r);
  const half = r - full >= 0.5 ? 1 : 0;
  const empty = 5 - full - half;
  return `<span class="star-display" role="img" aria-label="${r} out of 5 stars" title="${r} / 5">`
    + '<span class="star full">★</span>'.repeat(full)
    + (half ? '<span class="star half">★</span>' : "")
    + '<span class="star empty">★</span>'.repeat(empty)
    + `</span>`;
}

function escapeHtml(str) {
  return String(str || "").replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

const cssv = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const chartInk = () => cssv("--ink"), chartMuted = () => cssv("--mu"), chartGrid = () => cssv("--grid");
const SESSION_KEY = "w2u_session";
const THEME_KEY = "w2u_theme";
let pendingOtpUsername = null;
let currentAnalyticsChart = null;
let currentServiceTypeChart = null;
let currentMemberChart = null;
let currentDailyRevenueChart = null;

/* -------------------------------------------------------------------------
   Bright / dark mode toggle — persisted in localStorage, defaults to dark
   (the site's original look) so nothing changes for returning visitors
   until they choose "bright" for themselves.
------------------------------------------------------------------------- */
function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  document.getElementById("themeToggle").textContent = theme === "light" ? "🌙" : "☀️";
  localStorage.setItem(THEME_KEY, theme);
}
document.getElementById("themeToggle").addEventListener("click", () => {
  const current = document.documentElement.getAttribute("data-theme") || "light";
  applyTheme(current === "light" ? "dark" : "light");
});
applyTheme(localStorage.getItem(THEME_KEY) || "light");

/* -------------------------------------------------------------------------
   Low-level API helper. POSTs as text/plain to avoid a CORS preflight,
   which is the standard workaround for calling Apps Script cross-origin.
------------------------------------------------------------------------- */
async function apiCall(action, payload = {}, isRetry = false, requestId = null) {
  if (CONFIG.API_URL.includes("REPLACE_WITH_YOUR_DEPLOYMENT_ID")) {
    return { success: false, message: "Backend not configured yet. Set CONFIG.API_URL in script.js." };
  }
  // One id per logical action, reused if we retry below — NOT a new one per
  // attempt. This lets the server recognize "this is the same request you
  // already handled" if the first attempt actually succeeded server-side
  // but its response was merely slow or lost, so a retry can't duplicate a
  // mutation (an extra row, an extra email, etc.).
  if (!requestId) requestId = Date.now() + "-" + Math.random().toString(36).slice(2);
  // Hard safety net: without this, a connection that stalls (rather than
  // cleanly failing) leaves the fetch pending forever — which is exactly
  // what an infinite "Checking your session..." or a login stuck on
  // "Logging in..." looks like. Aborting after 20s guarantees this
  // function always settles one way or another.
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 35000);
  try {
    const res = await fetch(CONFIG.API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      cache: "no-store",
      signal: controller.signal,
      // Apps Script serves its response via a redirect to a
      // script.googleusercontent.com URL that's partly derived from the
      // request itself. Two identical requests (e.g. re-opening the same
      // tab, or checking a vehicle you just edited in the Sheet) can end up
      // hitting the exact same URL — which the browser's HTTP cache is then
      // free to serve from cache instead of asking Google again, returning
      // stale data even right after you've edited the Sheet. Appending a
      // unique value to every request guarantees a fresh URL every time.
      body: JSON.stringify({ action, ...payload, requestId, _: Date.now() + "-" + Math.random().toString(36).slice(2) })
    });
    clearTimeout(timeoutId);
    return await res.json();
  } catch (err) {
    clearTimeout(timeoutId);
    console.error("API error:", err);
    // Apps Script Web Apps serve their response via a redirect to
    // script.googleusercontent.com; if that hop gets blocked (ad blockers,
    // privacy extensions, network filters) the fetch fails here even though
    // the server-side action already completed. One retry clears up most
    // transient cases; if it still fails, surface a clear message.
    if (!isRetry) {
      await new Promise(r => setTimeout(r, 900));
      return apiCall(action, payload, true, requestId);
    }
    if (err.name === "AbortError") {
      return { success: false, message: "The server took too long to respond (over 35 seconds). Please check your connection and try again." };
    }
    return { success: false, message: "Network error contacting server. If this keeps happening, check whether a browser extension, antivirus, or network filter is blocking script.googleusercontent.com, then try again." };
  }
}

function getSession() { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch { return null; } }
function setSession(data) { localStorage.setItem(SESSION_KEY, JSON.stringify(data)); }
function clearSession() { localStorage.removeItem(SESSION_KEY); }
function showMsg(el, text, ok) { el.textContent = text; el.className = "form-msg show " + (ok ? "ok" : "err"); }
function setBusy(btn, busyText) {
  if (!btn) return;
  if (btn.dataset.originalText === undefined) btn.dataset.originalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = busyText || "Please wait...";
}
function clearBusy(btn) {
  if (!btn) return;
  btn.disabled = false;
  if (btn.dataset.originalText !== undefined) btn.textContent = btn.dataset.originalText;
}
function money(n) { return "RM " + Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function siteUrl() { return window.location.origin + window.location.pathname; }

/* -------------------------------------------------------------------------
   Nav (mobile toggle)
------------------------------------------------------------------------- */

/* -------------------------------------------------------------------------
   FEATURE 13 — booking form CAPTCHA (math challenge) + honeypot
------------------------------------------------------------------------- */
let captchaAnswer = 0;
function newCaptcha() {
  const a = Math.floor(Math.random() * 9) + 1;
  const b = Math.floor(Math.random() * 9) + 1;
  captchaAnswer = a + b;
  document.getElementById("captchaQuestion").textContent = `What is ${a} + ${b}?`;
  document.getElementById("captchaInput").value = "";
}
newCaptcha();

/* -------------------------------------------------------------------------
   FEATURE 10 — WhatsApp float button
------------------------------------------------------------------------- */
document.getElementById("whatsappFloat").href =
  `https://wa.me/${CONFIG.WHATSAPP_NUMBER}?text=${encodeURIComponent("Hi Workshop2U, I'd like to ask about your services.")}`;

/* -------------------------------------------------------------------------
   Public booking form
------------------------------------------------------------------------- */
const bookingForm = document.getElementById("bookingForm");
bookingForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = document.getElementById("bkSubmit");
  const msg = document.getElementById("bkMsg");

  if (Number(document.getElementById("captchaInput").value) !== captchaAnswer) {
    showMsg(msg, "That answer doesn't look right — please try the sum again.", false);
    newCaptcha();
    return;
  }

  btn.disabled = true; btn.textContent = "Submitting...";
  const booking = {
    name: document.getElementById("bkName").value.trim(),
    phone: document.getElementById("bkPhone").value.trim(),
    email: document.getElementById("bkEmail").value.trim(),
    workshop: document.getElementById("bkWorkshop").value,
    vehicleType: document.getElementById("bkVehicleType").value.trim(),
    plate: document.getElementById("bkPlate").value.trim(),
    serviceType: document.getElementById("bkService").value,
    date: document.getElementById("bkDate").value,
    time: document.getElementById("bkTime").value,
    notes: document.getElementById("bkNotes").value.trim(),
    website: document.getElementById("bkWebsite").value // honeypot — must stay empty
  };

  const result = await apiCall("bookAppointment", { booking });
  btn.disabled = false; btn.textContent = "Request Appointment";

  if (result.success) {
    showMsg(msg, "Thanks! Your appointment request has been received — we'll confirm by email shortly.", true);
    bookingForm.reset();
    newCaptcha();
  } else {
    showMsg(msg, result.message || "Something went wrong. Please try again.", false);
  }
});

/* -------------------------------------------------------------------------
   Login  (+ FEATURE 4 OTP step, + FEATURE 5 lockout messaging, + FEATURE 6 forgot password)
------------------------------------------------------------------------- */
const loginForm = document.getElementById("loginForm");
loginForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = document.getElementById("loginSubmit");
  const msg = document.getElementById("loginMsg");
  btn.disabled = true; btn.textContent = "Logging in...";

  const username = document.getElementById("loginUser").value.trim();
  const password = document.getElementById("loginPass").value;
  const result = await apiCall("login", { username, password });

  btn.disabled = false; btn.textContent = "Authenticate Access";

  if (result.success && result.otpRequired) {
    pendingOtpUsername = result.username;
    document.getElementById("loginPanel").classList.add("hidden");
    document.getElementById("otpPanel").classList.remove("hidden");
    document.getElementById("otpHint").textContent = "We emailed a 6-digit code to the account " + result.username + ".";
    msg.className = "form-msg";
  } else if (result.success) {
    setSession(result.user);
    loginForm.reset();
    msg.className = "form-msg";
    enterDashboard(result.user);
  } else {
    showMsg(msg, result.message || "Invalid username or password.", false);
  }
});

const otpForm = document.getElementById("otpForm");
otpForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const msg = document.getElementById("otpMsg");
  setBusy(btn, "Verifying...");
  const code = document.getElementById("otpCode").value.trim();
  const result = await apiCall("verifyOtp", { username: pendingOtpUsername, code });
  clearBusy(btn);
  if (result.success) {
    setSession(result.user);
    otpForm.reset();
    document.getElementById("otpPanel").classList.add("hidden");
    enterDashboard(result.user);
  } else {
    showMsg(msg, result.message || "Incorrect code.", false);
  }
});
document.getElementById("otpBack").addEventListener("click", () => {
  document.getElementById("otpPanel").classList.add("hidden");
  document.getElementById("loginPanel").classList.remove("hidden");
  otpForm.reset();
});

document.getElementById("showForgotPassword").addEventListener("click", (e) => {
  e.preventDefault();
  document.getElementById("loginPanel").classList.add("hidden");
  document.getElementById("forgotPanel").classList.remove("hidden");
});
document.getElementById("forgotBack").addEventListener("click", () => {
  document.getElementById("forgotPanel").classList.add("hidden");
  document.getElementById("loginPanel").classList.remove("hidden");
});
document.getElementById("forgotRequestForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const msg = document.getElementById("forgotRequestMsg");
  setBusy(btn, "Sending...");
  const usernameOrEmail = document.getElementById("forgotIdentifier").value.trim();
  const result = await apiCall("requestPasswordReset", { usernameOrEmail });
  clearBusy(btn);
  showMsg(msg, result.message || "If that account exists, a reset code has been emailed to it.", true);
  document.getElementById("resetUsername").value = usernameOrEmail;
  document.getElementById("forgotResetForm").classList.remove("hidden");
});
document.getElementById("forgotResetForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const msg = document.getElementById("forgotResetMsg");
  setBusy(btn, "Updating...");
  const username = document.getElementById("resetUsername").value.trim();
  const code = document.getElementById("resetCode").value.trim();
  const newPassword = document.getElementById("resetNewPassword").value;
  const result = await apiCall("resetPassword", { username, code, newPassword });
  clearBusy(btn);
  if (result.success) {
    showMsg(msg, "Password updated! You can now log in.", true);
    setTimeout(() => {
      document.getElementById("forgotPanel").classList.add("hidden");
      document.getElementById("loginPanel").classList.remove("hidden");
      document.getElementById("forgotResetForm").classList.add("hidden");
      document.getElementById("forgotRequestForm").reset();
      document.getElementById("forgotResetForm").reset();
    }, 1500);
  } else {
    showMsg(msg, result.message || "Could not reset password.", false);
  }
});

/* -------------------------------------------------------------------------
   Entering dashboards
------------------------------------------------------------------------- */
function enterDashboard(user) {
  if (window.portalEnter) window.portalEnter(user);
  document.getElementById("loginPanel").classList.add("hidden");
  document.getElementById("otpPanel").classList.add("hidden");
  document.getElementById("forgotPanel").classList.add("hidden");
  document.getElementById("authChecking").classList.add("hidden");

  if (user.role === "member") {
    document.getElementById("memberDashboard").classList.remove("hidden");
    document.getElementById("staffDashboard").classList.add("hidden");
    document.getElementById("memName").textContent = user.fullName;
    loadMemberProfile(user);
    loadMemberVehicles();
    loadMemberHistory();
  } else {
    document.getElementById("staffDashboard").classList.remove("hidden");
    document.getElementById("memberDashboard").classList.add("hidden");
    document.getElementById("staffName").textContent = user.fullName;
    document.getElementById("staffRoleBadge").textContent = user.role.toUpperCase();
    setupStaffScope(user);
    // Fetch what's needed for the visible stat-row + the default landing
    // tab (Dashboard) on login — everything else lazy-loads the first time
    // its tab is opened, see staffTabLoaders below.
    loadStaffHistory();
    loadStaffMembers();
    loadAnalytics();
    loadedStaffTabs.add("staffHistoryPanel");
    loadedStaffTabs.add("staffMembersPanel");
    loadedStaffTabs.add("staffAnalyticsPanel");
    if (user.role === "webmaster") {
      document.getElementById("tabAccounts").classList.remove("hidden");
      document.getElementById("tabAuditLog").classList.remove("hidden");
    }
  }
}

function setupStaffScope(user) {
  const wsFilter = document.getElementById("staffWorkshopFilter");
  const memFilter = document.getElementById("staffMemberWorkshopFilter");
  const bkFilter = document.getElementById("staffBookingWorkshopFilter");
  const revFilter = document.getElementById("staffReviewWorkshopFilter");
  const svcTypeFilter = document.getElementById("svcTypeWorkshopFilter");
  const memberChartFilter = document.getElementById("memberWorkshopFilter");
  const ahWorkshop = document.getElementById("ahWorkshop");
  const amWorkshop = document.getElementById("amWorkshop");

  if (user.role === "admin") {
    [wsFilter, memFilter, bkFilter, revFilter, svcTypeFilter, memberChartFilter, document.getElementById("dailyRevenueWorkshopFilter")].forEach(sel => { sel.innerHTML = `<option value="${user.workshop}">${user.workshop}</option>`; sel.disabled = true; });
    ahWorkshop.value = user.workshop; ahWorkshop.disabled = true;
    amWorkshop.value = user.workshop; amWorkshop.disabled = true;
    document.getElementById("staffScope").textContent = user.workshop;
  } else {
    document.getElementById("staffScope").textContent = "All Locations";
  }
}

function logoutAll() {
  if (window.portalLeave) window.portalLeave();
  const session = getSession();
  if (session) apiCall("logout", { token: session.token });
  clearSession();
  document.getElementById("memberDashboard").classList.add("hidden");
  document.getElementById("staffDashboard").classList.add("hidden");
  document.getElementById("loginPanel").classList.remove("hidden");
}
document.getElementById("menuLogout").addEventListener("click", logoutAll);

/* -------------------------------------------------------------------------
   Dashboard tab switching  (+ lazy-load each staff tab's data on first open,
   so login only fires 2 requests instead of hitting every tab at once)
------------------------------------------------------------------------- */
const staffTabLoaders = {
  staffHistoryPanel: loadStaffHistory,
  staffMembersPanel: loadStaffMembers,
  staffBookingsPanel: loadBookings,
  staffAnalyticsPanel: loadAnalytics,
  staffAccountsPanel: loadAccounts,
  staffAuditLogPanel: loadAuditLog,
  staffReviewsPanel: loadReviews
};
const loadedStaffTabs = new Set();

document.querySelectorAll(".dash-tab").forEach(tab => {
  tab.addEventListener("click", () => {
    const group = tab.closest(".dash-tabs");
    const dashEl = document.getElementById(tab.dataset.tab).closest("#memberDashboard, #staffDashboard");
    group.querySelectorAll(".dash-tab").forEach(t => t.classList.remove("active"));
    dashEl.querySelectorAll(".dash-panel").forEach(p => p.classList.remove("active"));
    tab.classList.add("active");
    document.getElementById(tab.dataset.tab).classList.add("active");

    const loader = staffTabLoaders[tab.dataset.tab];
    if (loader && (tab.dataset.tab === "staffAnalyticsPanel" || !loadedStaffTabs.has(tab.dataset.tab))) {
      loader();
      loadedStaffTabs.add(tab.dataset.tab);
    }
  });
});

/* -------------------------------------------------------------------------
   FEATURE 12 — export helpers (CSV + print/PDF)
------------------------------------------------------------------------- */
function tableToCSV(table) {
  const rows = [...table.querySelectorAll("tr")].filter(tr => !tr.classList.contains("empty-row"));
  return rows.map(tr => [...tr.children].filter(td => !td.classList.contains("no-export")).map(td => {
    const text = td.textContent.replace(/"/g, '""');
    return `"${text}"`;
  }).join(",")).join("\n");
}
function downloadCSV(tableId, filename) {
  const table = document.getElementById(tableId);
  const csv = tableToCSV(table);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
}
document.getElementById("memExportCsv").addEventListener("click", () => downloadCSV("memHistoryTable", "my-service-history.csv"));
document.getElementById("memPrint").addEventListener("click", () => window.print());
document.getElementById("staffExportCsv").addEventListener("click", () => downloadCSV("staffHistoryTable", "service-history.csv"));
document.getElementById("staffPrint").addEventListener("click", () => window.print());

/* -------------------------------------------------------------------------
   Member dashboard data  (+ FEATURE 9 vehicles, + FEATURE 1 next-due display)
------------------------------------------------------------------------- */
function loadMemberProfile(user) {
  document.getElementById("memProfName").value = user.fullName || "";
  document.getElementById("memProfAddress").value = user.address || "";
  document.getElementById("memProfPhone").value = user.phone || "";
  document.getElementById("memProfEmail").value = user.email || "";
}

async function loadMemberVehicles() {
  const session = getSession();
  const result = await apiCall("getVehicles", { token: session.token });
  const select = document.getElementById("memVehicleFilter");
  const listEl = document.getElementById("memVehicleList");
  select.innerHTML = `<option value="All">All my vehicles</option>`;
  if (!result.success || !result.vehicles.length) {
    listEl.innerHTML = `<p>No vehicles on file yet — they'll appear automatically after your first service.</p>`;
    return;
  }
  result.vehicles.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v.plate;
    opt.textContent = `${v.plate} — ${v.vehicleType}`;
    select.appendChild(opt);
  });
  listEl.innerHTML = result.vehicles.map(v => `
    <div class="feature-card" style="margin-bottom:14px">
      <h3>${v.plate}</h3><p>${v.vehicleType}</p>
    </div>`).join("");
}
document.getElementById("memVehicleFilter").addEventListener("change", loadMemberHistory);

async function loadMemberHistory() {
  const session = getSession();
  const tbody = document.getElementById("memHistoryBody");
  tbody.innerHTML = `<tr class="empty-row"><td colspan="7">Loading...</td></tr>`;
  const plate = document.getElementById("memVehicleFilter").value;
  const result = await apiCall("getHistory", { token: session.token, plate });

  if (!result.success) { tbody.innerHTML = `<tr class="empty-row"><td colspan="7">${result.message || "Could not load history."}</td></tr>`; return; }
  const rows = result.history || [];
  tbody.innerHTML = rows.length ? rows.map(r => `
      <tr><td>${r.date}</td><td>${r.vehicleType}</td><td>${r.plate}</td><td>${r.workshop}</td><td>${r.serviceType}</td><td>${money(r.price)}</td><td>${r.notes || ""}</td></tr>`).join("")
    : `<tr class="empty-row"><td colspan="7">No service history yet.</td></tr>`;

  document.getElementById("memTotalServices").textContent = rows.length;
  document.getElementById("memTotalSpent").textContent = money(rows.reduce((s, r) => s + Number(r.price || 0), 0));
  document.getElementById("memLastDate").textContent = rows.length ? rows[0].date : "—";
  document.getElementById("memNextDue").textContent = result.nextServiceDue || "—";
}

document.getElementById("memPasswordForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const session = getSession();
  const msg = document.getElementById("memPasswordMsg");
  setBusy(btn, "Updating...");
  const oldPassword = document.getElementById("memOldPass").value;
  const newPassword = document.getElementById("memNewPass").value;
  const result = await apiCall("changePassword", { token: session.token, oldPassword, newPassword });
  clearBusy(btn);
  if (result.success) { showMsg(msg, "Password updated successfully.", true); e.target.reset(); }
  else showMsg(msg, result.message || "Could not update password.", false);
});

/* -------------------------------------------------------------------------
   Staff dashboard — Service History
------------------------------------------------------------------------- */
let currentStaffHistoryRows = [];
async function loadStaffHistory() {
  const session = getSession();
  const tbody = document.getElementById("staffHistoryBody");
  tbody.innerHTML = `<tr class="empty-row"><td colspan="10">Loading...</td></tr>`;
  const workshop = document.getElementById("staffWorkshopFilter").value;
  const search = document.getElementById("staffSearch").value.trim().toLowerCase();

  const result = await apiCall("getHistory", { token: session.token, workshop });
  if (!result.success) { tbody.innerHTML = `<tr class="empty-row"><td colspan="10">${result.message || "Could not load history."}</td></tr>`; return; }
  let rows = result.history || [];
  if (search) rows = rows.filter(r => (r.plate || "").toLowerCase().includes(search) || (r.name || "").toLowerCase().includes(search));
  currentStaffHistoryRows = rows;

  tbody.innerHTML = rows.length ? rows.map(historyRowHtml).join("")
    : `<tr class="empty-row"><td colspan="10">No service records found.</td></tr>`;

  document.getElementById("staffTotalServices").textContent = rows.length;
  document.getElementById("staffTotalRevenue").textContent = money(rows.reduce((s, r) => s + Number(r.price || 0), 0));

}
const historyRowHtml = r => `
      <tr><td>${r.date}</td><td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.address)}</td><td>${escapeHtml(r.vehicleType)}</td><td>${escapeHtml(r.plate)}</td><td>${r.workshop}</td><td>${r.serviceType}</td><td>${money(r.price)}</td><td>${escapeHtml(r.notes)}</td>
      <td class="no-export">${r.id ? `<button class="btn btn-ghost btn-small" data-act="editHistory" data-id="${r.id}">Edit</button>` : ""}</td></tr>`;
/* One click listener for every Edit button (including rows added later without a reload). */
document.getElementById("staffHistoryBody").addEventListener("click", e => {
  const b = e.target.closest("button[data-act='editHistory']");
  if (b) startEditHistory(b.dataset.id);
});
const flashRow = tr => { if (tr) { tr.classList.add("row-new"); setTimeout(() => tr.classList.remove("row-new"), 2500); } };
/* Adds ONE new row to the list that is already on screen (newest date first),
   instead of reloading the whole list. Respects the current workshop filter / search. */
function appendHistoryRow(rec, id) {
  const wf = document.getElementById("staffWorkshopFilter").value;
  const search = document.getElementById("staffSearch").value.trim().toLowerCase();
  const row = { id: id || "", date: rec.date, username: rec.username, name: rec.name, address: rec.address, vehicleType: rec.vehicleType, plate: rec.plate, workshop: rec.workshop, serviceType: rec.serviceType, price: rec.price, notes: rec.notes };
  if (wf !== "All" && row.workshop !== wf) return;
  if (search && !((row.plate || "").toLowerCase().includes(search) || (row.name || "").toLowerCase().includes(search))) return;
  const tbody = document.getElementById("staffHistoryBody");
  if (!currentStaffHistoryRows.length) tbody.innerHTML = "";
  let idx = currentStaffHistoryRows.findIndex(r => new Date(r.date) < new Date(row.date));
  if (idx < 0) idx = currentStaffHistoryRows.length;
  currentStaffHistoryRows.splice(idx, 0, row);
  const html = historyRowHtml(row);
  if (tbody.rows[idx]) tbody.rows[idx].insertAdjacentHTML("beforebegin", html); else tbody.insertAdjacentHTML("beforeend", html);
  flashRow(tbody.rows[idx]);
  document.getElementById("staffTotalServices").textContent = currentStaffHistoryRows.length;
  document.getElementById("staffTotalRevenue").textContent = money(currentStaffHistoryRows.reduce((t, r) => t + Number(r.price || 0), 0));
}

/* Reuses the "Add Service Record" form for editing — switches it into edit
   mode (tracked via a data attribute on the form), pre-fills every field
   from the row the staff member clicked Edit on, and reveals the form
   (it now lives inline inside the Service History panel itself). */
function startEditHistory(id) {
  const row = currentStaffHistoryRows.find(r => r.id === id);
  if (!row) return;

  document.getElementById("ahDate").value = row.date;
  document.getElementById("ahUsername").value = row.username || "";
  document.getElementById("ahName").value = row.name || "";
  document.getElementById("ahAddress").value = row.address || "";
  document.getElementById("ahVehicleType").value = row.vehicleType || "";
  document.getElementById("ahPlate").value = row.plate || "";
  if (!document.getElementById("ahWorkshop").disabled) document.getElementById("ahWorkshop").value = row.workshop;
  document.getElementById("ahServiceType").value = row.serviceType;
  document.getElementById("ahPrice").value = row.price;
  document.getElementById("ahNotes").value = row.notes || "";

  const form = document.getElementById("addHistoryForm");
  form.dataset.editId = id;
  document.getElementById("addHistoryFormTitle").textContent = "Edit Service Record";
  document.getElementById("addHistorySubmitBtn").textContent = "Update Service Record";
  document.getElementById("addHistoryNote").textContent = "Editing a record does not re-send the invoice email or create a new review link.";
  document.getElementById("cancelEditHistory").textContent = "Cancel Edit";
  document.getElementById("addHistoryMsg").className = "form-msg";

  const panel = document.getElementById("staffAddHistoryPanel");
  panel.classList.remove("hidden");
  panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

/* Also doubles as the "+ Add Service Record" open button's reset state —
   resets the form back to plain add-mode and hides the panel again. */
/* Resets the form to plain "add a new record" state — used both when
   cancelling and when opening via "+ Add Service Record", so that clicking
   the + button while mid-edit (instead of Cancel) can't leave a stale
   editId behind and silently overwrite the row being edited. */
function resetToAddMode() {
  const form = document.getElementById("addHistoryForm");
  delete form.dataset.editId;
  form.reset();
  if (getSession().role === "admin") document.getElementById("ahWorkshop").value = getSession().workshop;
  document.getElementById("addHistoryFormTitle").textContent = "Add Service Record";
  document.getElementById("addHistorySubmitBtn").textContent = "Save Service Record";
  document.getElementById("addHistoryNote").textContent = "Saving a record automatically emails the customer a PDF invoice and a feedback link, and updates their vehicle list.";
  document.getElementById("cancelEditHistory").textContent = "Cancel";
  document.getElementById("addHistoryMsg").className = "form-msg";
}
function exitEditHistory() {
  resetToAddMode();
  document.getElementById("staffAddHistoryPanel").classList.add("hidden");
}
document.getElementById("cancelEditHistory").addEventListener("click", exitEditHistory);
document.getElementById("openAddHistory").addEventListener("click", () => {
  resetToAddMode();
  document.getElementById("staffAddHistoryPanel").classList.remove("hidden");
  document.getElementById("ahDate").focus();
  document.getElementById("staffAddHistoryPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
});
document.getElementById("staffRefresh").addEventListener("click", loadStaffHistory);
document.getElementById("staffWorkshopFilter").addEventListener("change", loadStaffHistory);
document.getElementById("staffSearch").addEventListener("input", () => { clearTimeout(window._searchDebounce); window._searchDebounce = setTimeout(loadStaffHistory, 300); });

document.getElementById("addHistoryForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const session = getSession();
  const msg = document.getElementById("addHistoryMsg");
  const editId = e.target.dataset.editId;
  setBusy(btn, editId ? "Updating..." : "Saving...");
  const record = {
    date: document.getElementById("ahDate").value,
    username: document.getElementById("ahUsername").value.trim(),
    name: document.getElementById("ahName").value.trim(),
    address: document.getElementById("ahAddress").value.trim(),
    vehicleType: document.getElementById("ahVehicleType").value.trim(),
    plate: document.getElementById("ahPlate").value.trim(),
    workshop: document.getElementById("ahWorkshop").value,
    serviceType: document.getElementById("ahServiceType").value,
    price: document.getElementById("ahPrice").value,
    notes: document.getElementById("ahNotes").value.trim()
  };
  const result = editId
    ? await apiCall("updateHistory", { token: session.token, historyId: editId, record })
    : await apiCall("addHistory", { token: session.token, record, siteUrl: siteUrl() });
  clearBusy(btn);
  if (result.success) {
    if (editId) {
      showMsg(msg, "Service record updated.", true);
      exitEditHistory();
    } else {
      showMsg(msg, "Service record saved — invoice emailed to the customer if we have their email on file.", true);
      e.target.reset();
      if (getSession().role === "admin") document.getElementById("ahWorkshop").value = getSession().workshop;
      appendHistoryRow(record, result.id);   // add just this row; the rest of the list stays as is
    }
    if (editId) loadStaffHistory();
    loadAnalytics();
  } else {
    showMsg(msg, result.message || "Could not save record.", false);
  }
});

/* -------------------------------------------------------------------------
   Staff dashboard — Members
------------------------------------------------------------------------- */
async function loadStaffMembers() {
  const session = getSession();
  const tbody = document.getElementById("staffMembersBody");
  tbody.innerHTML = `<tr class="empty-row"><td colspan="8">Loading...</td></tr>`;
  const workshop = document.getElementById("staffMemberWorkshopFilter").value;
  const result = await apiCall("getMembers", { token: session.token, workshop });
  if (!result.success) { tbody.innerHTML = `<tr class="empty-row"><td colspan="8">${result.message || "Could not load members."}</td></tr>`; return; }
  const rows = result.members || [];
  tbody.innerHTML = rows.length ? rows.map(memberRowHtml).join("")
    : `<tr class="empty-row"><td colspan="8">No members found.</td></tr>`;
  document.getElementById("staffTotalMembers").textContent = rows.length;
}
const memberRowHtml = m => `
      <tr><td>${escapeHtml(m.username)}</td><td>${escapeHtml(m.fullName)}</td><td>${escapeHtml(m.address || "")}</td><td>${escapeHtml(m.phone || "")}</td><td>${escapeHtml(m.email || "")}</td><td>${m.workshop}</td>
      <td><span class="tag ${m.status === 'Active' ? 'active' : 'inactive'}">${m.status}</span></td>
      <td><button class="btn btn-ghost btn-small" data-act="toggleMember" data-user="${escapeHtml(m.username)}" data-status="${m.status}">${m.status === 'Active' ? 'Deactivate' : 'Activate'}</button></td></tr>`;
function appendMemberRow(m) {
  const wf = document.getElementById("staffMemberWorkshopFilter").value;
  if (wf !== "All" && m.workshop !== wf) return;
  const tbody = document.getElementById("staffMembersBody");
  if (tbody.querySelector(".empty-row")) tbody.innerHTML = "";
  tbody.insertAdjacentHTML("beforeend", memberRowHtml(m));
  flashRow(tbody.lastElementChild);
  document.getElementById("staffTotalMembers").textContent = tbody.rows.length;
}
{
  const tbody = document.getElementById("staffMembersBody");
  tbody.addEventListener("click", async e => {
    const btn = e.target.closest("button[data-act='toggleMember']");
    if (!btn) return;
    const session = getSession();
    {
      const newStatus = btn.dataset.status === "Active" ? "Inactive" : "Active";
      if (newStatus === "Inactive" && !confirm(`Deactivate ${btn.dataset.user}? They won't be able to log in until reactivated. Their service history is kept either way.`)) return;
      setBusy(btn, "...");
      const result = await apiCall("updateMemberStatus", { token: session.token, username: btn.dataset.user, status: newStatus });
      clearBusy(btn);
      if (result.success) {
        // update just this row in place
        const tr = btn.closest("tr");
        btn.dataset.status = newStatus;
        btn.textContent = newStatus === "Active" ? "Deactivate" : "Activate";
        const tag = tr.querySelector(".tag");
        tag.textContent = newStatus;
        tag.className = "tag " + (newStatus === "Active" ? "active" : "inactive");
      } else alert(result.message || "Could not update member.");
    }
  });
}
document.getElementById("staffMembersRefresh").addEventListener("click", loadStaffMembers);
document.getElementById("staffMemberWorkshopFilter").addEventListener("change", loadStaffMembers);

document.getElementById("addMemberForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const session = getSession();
  const msg = document.getElementById("addMemberMsg");
  setBusy(btn, "Creating...");
  const member = {
    username: document.getElementById("amUsername").value.trim(),
    password: document.getElementById("amPassword").value,
    fullName: document.getElementById("amName").value.trim(),
    phone: document.getElementById("amPhone").value.trim(),
    address: document.getElementById("amAddress").value.trim(),
    email: document.getElementById("amEmail").value.trim(),
    workshop: document.getElementById("amWorkshop").value,
    role: "member"
  };
  const result = await apiCall("addMember", { token: session.token, member });
  clearBusy(btn);
  if (result.success) {
    showMsg(msg, "Member account created.", true);
    e.target.reset();
    if (getSession().role === "admin") document.getElementById("amWorkshop").value = getSession().workshop;
    appendMemberRow({ username: member.username, fullName: member.fullName, address: member.address, phone: member.phone, email: member.email, workshop: member.workshop, status: "Active" });
  } else {
    showMsg(msg, result.message || "Could not create member.", false);
  }
});
document.getElementById("openAddMember").addEventListener("click", () => {
  document.getElementById("staffAddMemberPanel").classList.remove("hidden");
  document.getElementById("amUsername").focus();
  document.getElementById("staffAddMemberPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
});
document.getElementById("cancelAddMember").addEventListener("click", () => {
  document.getElementById("addMemberForm").reset();
  if (getSession().role === "admin") document.getElementById("amWorkshop").value = getSession().workshop;
  document.getElementById("addMemberMsg").className = "form-msg";
  document.getElementById("staffAddMemberPanel").classList.add("hidden");
});

document.getElementById("staffPasswordForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const session = getSession();
  const msg = document.getElementById("staffPasswordMsg");
  setBusy(btn, "Updating...");
  const oldPassword = document.getElementById("staffOldPass").value;
  const newPassword = document.getElementById("staffNewPass").value;
  const result = await apiCall("changePassword", { token: session.token, oldPassword, newPassword });
  clearBusy(btn);
  if (result.success) { showMsg(msg, "Password updated successfully.", true); e.target.reset(); }
  else showMsg(msg, result.message || "Could not update password.", false);
});

/* -------------------------------------------------------------------------
   FEATURE 2 — Booking approval workflow
------------------------------------------------------------------------- */
async function loadBookings() {
  const session = getSession();
  const tbody = document.getElementById("bookingsBody");
  tbody.innerHTML = `<tr class="empty-row"><td colspan="9">Loading...</td></tr>`;
  const workshop = document.getElementById("staffBookingWorkshopFilter").value;
  const result = await apiCall("getBookings", { token: session.token, workshop });
  if (!result.success) { tbody.innerHTML = `<tr class="empty-row"><td colspan="9">${result.message || "Could not load bookings."}</td></tr>`; return; }
  const rows = result.bookings || [];
  tbody.innerHTML = rows.length ? rows.map(b => `
      <tr>
        <td>${b.date} ${b.time}</td><td>${escapeHtml(b.name)}</td><td>${escapeHtml(b.phone)}</td><td>${escapeHtml(b.vehicleType)} (${escapeHtml(b.plate)})</td>
        <td>${b.workshop}</td><td>${b.serviceType}</td>
        <td><span class="tag ${b.status === 'Confirmed' ? 'active' : (b.status === 'Rejected' ? 'inactive' : '')}">${b.status}</span></td>
        <td>${escapeHtml(b.notes)}</td>
        <td>
          <button class="btn btn-ghost btn-small" data-act="Confirmed" data-id="${b.id}">Confirm</button>
          <button class="btn btn-ghost btn-small" data-act="Rejected" data-id="${b.id}">Reject</button>
          <button class="btn btn-ghost btn-small" data-act="reschedule" data-id="${b.id}">Reschedule</button>
        </td>
      </tr>`).join("")
    : `<tr class="empty-row"><td colspan="9">No booking requests found.</td></tr>`;

  tbody.querySelectorAll("button[data-act]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const id = btn.dataset.id;
      let result, shown;
      setBusy(btn, "...");
      if (btn.dataset.act === "reschedule") {
        const newDate = prompt("New date (YYYY-MM-DD):");
        if (!newDate) { clearBusy(btn); return; }
        const newTime = prompt("New time (HH:MM):");
        result = await apiCall("updateBookingStatus", { token: session.token, bookingId: id, status: "Rescheduled", newDate, newTime });
        shown = { status: "Rescheduled", when: newDate + " " + (newTime || "") };
      } else {
        result = await apiCall("updateBookingStatus", { token: session.token, bookingId: id, status: btn.dataset.act });
        shown = { status: btn.dataset.act };
      }
      clearBusy(btn);
      if (!result || !result.success) { alert((result && result.message) || "Could not update booking."); return; }
      const tr = btn.closest("tr");            // change only this row — the list stays put
      if (shown.when) tr.children[0].textContent = shown.when.trim();
      const tag = tr.children[6].querySelector(".tag");
      tag.textContent = shown.status;
      tag.className = "tag " + (shown.status === "Confirmed" ? "active" : (shown.status === "Rejected" ? "inactive" : ""));
      flashRow(tr);
    });
  });
}
document.getElementById("staffBookingWorkshopFilter").addEventListener("change", loadBookings);
document.getElementById("bookingsRefresh").addEventListener("click", loadBookings);

/* -------------------------------------------------------------------------
   FEATURE 7 — Analytics dashboard (Chart.js)
------------------------------------------------------------------------- */
/* The month dropdown is populated once and left alone after that — if it
   were rebuilt on every loadAnalytics() call (which also runs whenever the
   *other* two dropdowns change), it would silently reset the daily-revenue
   chart back to the current month even if the staff member had deliberately
   picked an earlier one. Starts at September 2026 (the site's go-live
   month) through whichever month it actually is today, so it grows on its
   own with no further edits needed. */
function populateMonthDropdown() {
  const sel = document.getElementById("dailyRevenueMonthFilter");
  if (sel.options.length) return; // already populated — leave the user's selection alone
  const start = new Date(2026, 8, 1); // September 2026
  const now = new Date();
  const cursor = new Date(start.getFullYear(), start.getMonth(), 1);
  const options = [];
  while (cursor <= now) {
    const value = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`;
    const label = cursor.toLocaleString("en-US", { month: "long", year: "numeric" });
    options.push(`<option value="${value}">${label}</option>`);
    cursor.setMonth(cursor.getMonth() + 1);
  }
  sel.innerHTML = options.join("");
  sel.selectedIndex = options.length - 1; // default to the current month
}

/* Builds (or rebuilds) a pie chart in canvasId. Returns the new Chart
   instance, or null if there's nothing to draw (caller shows an empty-state
   message instead). Tooltips show both quantity and percentage, per the
   request ("percentage ... and quantity"). */
/* Draws "count (percent)" on each slice that is big enough to hold it. Slices too
   small for the text are left blank — hovering them shows the same info in the tooltip. */
const pieSliceLabels = {
  id: "pieSliceLabels",
  afterDatasetsDraw(chart) {
    const ctx = chart.ctx, meta = chart.getDatasetMeta(0), data = chart.data.datasets[0].data;
    const total = data.reduce((a, b) => a + Number(b), 0);
    meta.data.forEach((arc, i) => {
      const v = Number(data[i]);
      if (!v || !total) return;
      const text = `${v} (${Math.round((v / total) * 100)}%)`;
      const { startAngle, endAngle, innerRadius, outerRadius, x, y } = arc.getProps(["startAngle", "endAngle", "innerRadius", "outerRadius", "x", "y"], true);
      const mid = (startAngle + endAngle) / 2, r = (innerRadius + outerRadius) * 0.62;
      ctx.save();
      ctx.font = "600 12px Inter, sans-serif";
      const w = ctx.measureText(text).width;
      const chord = 2 * r * Math.sin(Math.min(endAngle - startAngle, Math.PI) / 2);  // room across the slice
      if (chord >= w + 8 && outerRadius > 45) {
        ctx.fillStyle = "#fff"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
        ctx.shadowColor = "rgba(0,0,0,.55)"; ctx.shadowBlur = 3;
        ctx.fillText(text, x + Math.cos(mid) * r, y + Math.sin(mid) * r);
      }
      ctx.restore();
    });
  }
};

function buildPieChart(canvasId, labels, counts, colorMap) {
  const total = counts.reduce((a, b) => a + b, 0);
  if (!labels.length || total === 0) return null;
  const canvas = document.getElementById(canvasId);
  return new Chart(canvas.getContext("2d"), {
    type: "pie",
    plugins: [pieSliceLabels],
    data: { labels, datasets: [{ data: counts, backgroundColor: labels.map(l => colorMap[l] || "#98a1ad") }] },
    options: {
      plugins: {
        legend: { position: "bottom", labels: { color: chartInk() } },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${ctx.raw} (${((ctx.raw / total) * 100).toFixed(1)}%)` } }
      }
    }
  });
}

const SERVICE_TYPE_COLORS = { "Normal Service": "#f2a71b", "Repair": "#c94a3d", "Troubleshooting": chartMuted(), "Car Wash": "#2f8f5b", "A/C Service": "#7a8ba3" };
const WORKSHOP_COLORS = { "Melaka": "#f2a71b", "Negeri Sembilan": chartMuted(), "Johor": "#2f8f5b" };
const STATUS_COLORS = { "Active": "#2f8f5b", "Inactive": "#c94a3d" };

async function loadAnalytics() {
  const session = getSession();
  populateMonthDropdown();

  const serviceTypeWorkshop = document.getElementById("svcTypeWorkshopFilter").value;
  const memberWorkshop = document.getElementById("memberWorkshopFilter").value;
  const month = document.getElementById("dailyRevenueMonthFilter").value;
  const dailyWorkshop = document.getElementById("dailyRevenueWorkshopFilter").value;

  const result = await apiCall("getAnalytics", { token: session.token, serviceTypeWorkshop, memberWorkshop, month, dailyWorkshop });
  if (!result.success) return;

  document.getElementById("anaTotalRevenue").textContent = money(result.totalRevenue);
  document.getElementById("anaTotalJobs").textContent = result.totalJobs;

  // ---- existing: monthly revenue by workshop ----
  const datasets = result.workshops.map(w => ({ label: w, data: result.revenueSeries[w], backgroundColor: WORKSHOP_COLORS[w] || "#98a1ad" }));
  if (currentAnalyticsChart) currentAnalyticsChart.destroy();
  currentAnalyticsChart = new Chart(document.getElementById("analyticsChart").getContext("2d"), {
    type: "bar",
    data: { labels: result.months, datasets },
    options: {
      responsive: true,
      // .chart-panel is a fixed white surface in both themes (same as
      // stat boxes and tables), so these need to be dark ink colors, not
      // the light colors that would suit a dark background.
      plugins: { legend: { labels: { color: chartInk() } }, title: { display: true, text: "Monthly revenue by workshop (RM)", color: chartInk() } },
      scales: {
        x: { ticks: { color: chartMuted() }, grid: { color: chartGrid() } },
        y: { ticks: { color: chartMuted() }, grid: { color: chartGrid() } }
      }
    }
  });

  // ---- pie 1: service type breakdown ----
  if (currentServiceTypeChart) { currentServiceTypeChart.destroy(); currentServiceTypeChart = null; }
  const svc = result.serviceTypeBreakdown;
  currentServiceTypeChart = buildPieChart("serviceTypeChart", svc.labels, svc.counts, SERVICE_TYPE_COLORS);
  document.getElementById("serviceTypeChart").classList.toggle("hidden", !currentServiceTypeChart);
  document.getElementById("serviceTypeEmptyMsg").classList.toggle("hidden", !!currentServiceTypeChart);

  // ---- pie 2: members — by workshop when scope is "All", else Active/Inactive for the one chosen workshop ----
  if (currentMemberChart) { currentMemberChart.destroy(); currentMemberChart = null; }
  const mem = result.memberBreakdown;
  const memColors = mem.mode === "status" ? STATUS_COLORS : WORKSHOP_COLORS;
  currentMemberChart = buildPieChart("memberChart", mem.labels, mem.counts, memColors);
  document.getElementById("memberChart").classList.toggle("hidden", !currentMemberChart);
  document.getElementById("memberEmptyMsg").classList.toggle("hidden", !!currentMemberChart);
  document.getElementById("memberChartTitle").textContent = mem.mode === "status" ? `Member Status — ${mem.scope[0]}` : "Members by Workshop";

  // ---- bar + line overlay: daily revenue for the selected month, split by
  // service type. The bar is total daily revenue; each line is one service
  // type's slice of that same total, using the same colors as the Service
  // Type pie chart above for consistency. ----
  const dr = result.dailyRevenue;
  const daysInMonth = dr.days.length;
  const dayLabels = Array.from({ length: daysInMonth }, (_, i) => i + 1);
  const serviceTypeLines = Object.keys(dr.byServiceType).map(type => ({
    type: "line",
    label: type,
    data: dr.byServiceType[type],
    borderColor: SERVICE_TYPE_COLORS[type] || "#98a1ad",
    backgroundColor: SERVICE_TYPE_COLORS[type] || "#98a1ad",
    borderWidth: 2,
    pointRadius: 2,
    tension: 0.25,
    yAxisID: "y"
  }));
  if (currentDailyRevenueChart) currentDailyRevenueChart.destroy();
  currentDailyRevenueChart = new Chart(document.getElementById("dailyRevenueChart").getContext("2d"), {
    type: "bar",
    data: {
      labels: dayLabels,
      datasets: [
        { type: "bar", label: "Total revenue", data: dr.days, backgroundColor: "#2f6fed", order: 2 },
        ...serviceTypeLines.map(l => ({ ...l, order: 1 }))
      ]
    },
    options: {
      plugins: { legend: { display: true, position: "bottom", labels: { color: chartInk() } } },
      scales: {
        x: { ticks: { color: chartMuted() }, grid: { color: chartGrid() } },
        y: {
          ticks: { color: chartMuted(), callback: value => "RM " + value.toLocaleString() },
          grid: { color: chartGrid() },
          title: { display: true, text: "Revenue (RM)", color: chartMuted() }
        }
      }
    }
  });
}
document.getElementById("svcTypeWorkshopFilter").addEventListener("change", loadAnalytics);
document.getElementById("memberWorkshopFilter").addEventListener("change", loadAnalytics);
document.getElementById("dailyRevenueMonthFilter").addEventListener("change", loadAnalytics);
document.getElementById("dailyRevenueWorkshopFilter").addEventListener("change", loadAnalytics);

/* -------------------------------------------------------------------------
   FEATURE 8 (display) — customer reviews, for staff to actually see them.
   Reviews only exist here once a customer has submitted one — see
   getReviewContext_/submitReview_ in Code.gs, which write to the Reviews
   tab only at the moment of submission, not when the service was logged.
------------------------------------------------------------------------- */
async function loadReviews() {
  const session = getSession();
  const tbody = document.getElementById("reviewsBody");
  tbody.innerHTML = `<tr class="empty-row"><td colspan="5">Loading...</td></tr>`;
  const workshop = document.getElementById("staffReviewWorkshopFilter").value;
  const result = await apiCall("getReviews", { token: session.token, workshop });
  if (!result.success) { tbody.innerHTML = `<tr class="empty-row"><td colspan="5">${result.message || "Could not load reviews."}</td></tr>`; return; }
  const rows = result.reviews || [];
  tbody.innerHTML = rows.length ? rows.map(r => `
      <tr><td>${r.submittedAt}</td><td>${escapeHtml(r.name)}</td><td>${r.workshop}</td><td>${renderStars(r.rating)}</td><td>${escapeHtml(r.comment)}</td></tr>`).join("")
    : `<tr class="empty-row"><td colspan="5">No reviews submitted yet.</td></tr>`;
  document.getElementById("revAvgRating").innerHTML = rows.length
    ? `${result.avgRating.toFixed(1)} / 5 ${renderStars(result.avgRating)}`
    : "—";
  document.getElementById("revCount").textContent = result.count;
}
document.getElementById("staffReviewWorkshopFilter").addEventListener("change", loadReviews);
document.getElementById("reviewsRefresh").addEventListener("click", loadReviews);
async function loadAccounts() {
  const session = getSession();
  const tbody = document.getElementById("accountsBody");
  tbody.innerHTML = `<tr class="empty-row"><td colspan="6">Loading...</td></tr>`;
  const result = await apiCall("getAccounts", { token: session.token });
  if (!result.success) { tbody.innerHTML = `<tr class="empty-row"><td colspan="6">${result.message || "Could not load accounts."}</td></tr>`; return; }
  const rows = result.accounts || [];
  tbody.innerHTML = rows.map(a => `
    <tr><td>${a.username}</td><td>${a.role}</td><td>${a.workshop}</td><td>${a.fullName}</td>
      <td><span class="tag ${a.status === 'Active' ? 'active' : 'inactive'}">${a.status}</span></td>
      <td>
        <button class="btn btn-ghost btn-small" data-act="toggle" data-user="${a.username}">${a.status === 'Active' ? 'Deactivate' : 'Activate'}</button>
        <button class="btn btn-ghost btn-small" data-act="reset" data-user="${a.username}">Reset Password</button>
      </td></tr>`).join("");

  tbody.querySelectorAll("button[data-act]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const username = btn.dataset.user;
      if (btn.dataset.act === "toggle") {
        const row = rows.find(r => r.username === username);
        const newStatus = row.status === "Active" ? "Inactive" : "Active";
        setBusy(btn, "...");
        await apiCall("updateAccount", { token: session.token, username, updates: { status: newStatus } });
        loadAccounts();
      } else {
        const newPass = prompt("Enter a new temporary password for " + username + ":");
        if (newPass) {
          setBusy(btn, "...");
          await apiCall("updateAccount", { token: session.token, username, updates: { password: newPass } });
          clearBusy(btn);
          alert("Password reset.");
        }
      }
    });
  });
}

document.getElementById("addStaffForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = e.target.querySelector('button[type="submit"]');
  const session = getSession();
  const msg = document.getElementById("addStaffMsg");
  setBusy(btn, "Creating...");
  const member = {
    username: document.getElementById("asUsername").value.trim(),
    password: document.getElementById("asPassword").value,
    fullName: document.getElementById("asName").value.trim(),
    email: document.getElementById("asEmail").value.trim(),
    role: document.getElementById("asRole").value,
    workshop: document.getElementById("asWorkshop").value
  };
  const result = await apiCall("addMember", { token: session.token, member });
  clearBusy(btn);
  if (result.success) { showMsg(msg, "Staff account created.", true); e.target.reset(); loadAccounts(); }
  else showMsg(msg, result.message || "Could not create account.", false);
});

async function loadAuditLog() {
  const session = getSession();
  const tbody = document.getElementById("auditLogBody");
  tbody.innerHTML = `<tr class="empty-row"><td colspan="5">Loading...</td></tr>`;
  const result = await apiCall("getAuditLog", { token: session.token });
  if (!result.success) { tbody.innerHTML = `<tr class="empty-row"><td colspan="5">${result.message || "Could not load audit log."}</td></tr>`; return; }
  const rows = result.log || [];
  tbody.innerHTML = rows.length ? rows.map(r => `<tr><td>${r.timestamp}</td><td>${r.username}</td><td>${r.role}</td><td>${r.action}</td><td>${r.details}</td></tr>`).join("")
    : `<tr class="empty-row"><td colspan="5">No activity recorded yet.</td></tr>`;
}
document.getElementById("auditLogRefresh").addEventListener("click", loadAuditLog);

/* -------------------------------------------------------------------------
   FEATURE 8 — Review submission (triggered via ?review=TOKEN in the URL)
------------------------------------------------------------------------- */
let selectedRating = 0;
function initReviewFlow() {
  const params = new URLSearchParams(window.location.search);
  const token = params.get("review");
  if (!token) return;

  // Open the popup and make the star form usable immediately — purely from
  // the token in the URL, no network wait. The workshop name and the
  // invalid/already-submitted check both come from the server, but neither
  // needs to block the customer from being able to start rating right away;
  // waiting on that round-trip first was the reason the popup used to take
  // 5-10 seconds to appear (and could get missed entirely if the customer
  // clicked elsewhere on the site before it showed up).
  document.getElementById("reviewModal").classList.remove("hidden");
  document.getElementById("reviewForm").dataset.token = token;

  apiCall("getReviewContext", { reviewToken: token }).then(result => {
    if (!result.success) {
      document.getElementById("reviewFormSection").classList.add("hidden");
      const invalid = document.getElementById("reviewInvalidMsg");
      invalid.classList.remove("hidden");
      invalid.innerHTML = `<p>${escapeHtml(result.message)}</p>`;
      return;
    }
    document.getElementById("reviewWorkshopName").textContent = result.workshop;
    document.getElementById("reviewForm").dataset.workshop = result.workshop;
  });
}

document.querySelectorAll(".star-btn").forEach(star => {
  star.addEventListener("click", () => {
    selectedRating = Number(star.dataset.value);
    document.querySelectorAll(".star-btn").forEach(s => s.classList.toggle("selected", Number(s.dataset.value) <= selectedRating));
  });
});
document.getElementById("reviewForm").addEventListener("submit", (e) => {
  e.preventDefault();
  const msg = document.getElementById("reviewMsg");
  if (!selectedRating) { showMsg(msg, "Please select a star rating first.", false); return; }
  const token = e.target.dataset.token;
  const workshop = e.target.dataset.workshop;
  const comment = document.getElementById("reviewComment").value.trim();

  // Optimistic UI: reviews are low-stakes, and a customer watching a
  // spinner for 30+ seconds (see the network-timing thread with the
  // developer if that's happening) is far more likely to abandon than the
  // rare background-save hiccup is to matter. So we show "thank you"
  // immediately — the actual save happens right after, just not blocking
  // what the customer sees. If it genuinely fails, we quietly switch back
  // to the form with nothing lost, and a clear message.
  document.getElementById("reviewFormSection").classList.add("hidden");
  showReviewThankYou(workshop, comment);

  apiCall("submitReview", { reviewToken: token, rating: selectedRating, comment }).then(result => {
    if (!result.success) {
      document.getElementById("reviewThankYouSection").classList.add("hidden");
      document.getElementById("reviewFormSection").classList.remove("hidden");
      showMsg(document.getElementById("reviewMsg"), result.message || "We couldn't save that — please try submitting again.", false);
    }
  });
});

/* After a successful internal review, invite the customer to post the same
   feedback on Google too — with a one-click copy of what they just wrote,
   so they don't have to retype it over there. */
/* Looks up the Google review link for a workshop. Tolerant of case/extra
   spaces in the workshop name, and adds https:// if the link was pasted
   without it (otherwise the browser treats it as a path on your own site).
   Returns "" — so no Google button is shown — if none is configured. */
function getGoogleReviewLink(workshop) {
  const links = CONFIG.GOOGLE_REVIEW_LINKS || {};
  const wanted = String(workshop || "").trim().toLowerCase();
  const key = Object.keys(links).find(k => k.trim().toLowerCase() === wanted);
  let link = key ? String(links[key] || "").trim() : "";
  if (!link) {
    console.warn(`No Google review link set for "${workshop}" — add one to CONFIG.GOOGLE_REVIEW_LINKS in script.js to show the "Submit Google Review" button.`);
    return "";
  }
  if (!/^https?:\/\//i.test(link)) link = "https://" + link;
  return link;
}

function showReviewThankYou(workshop, comment) {
  const googleLink = getGoogleReviewLink(workshop);
  const hasGoogleLink = googleLink.length > 0;

  let html = `<h3>Thank you for your feedback!</h3>`;

  if (comment) {
    html += `
      <div class="form-panel" style="padding:14px 16px;margin-bottom:14px;">
        <p style="margin:0;font-style:italic;">"${escapeHtml(comment)}"</p>
      </div>`;
  }

  if (hasGoogleLink) {
    html += `<p>Mind sharing this on Google too? It really helps other drivers find us.</p>`;
    if (comment) html += `<button class="btn btn-ghost btn-block" id="copyReviewBtn" type="button" style="margin-bottom:10px;">Copy My Review</button>`;
    html += `<a class="btn btn-amber btn-block" href="${escapeHtml(googleLink)}" target="_blank" rel="noopener">Submit Google Review</a>`;
  }

  const section = document.getElementById("reviewThankYouSection");
  section.innerHTML = html;
  section.classList.remove("hidden");

  const copyBtn = document.getElementById("copyReviewBtn");
  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(comment);
        copyBtn.textContent = "Copied!";
        setTimeout(() => { copyBtn.textContent = "Copy My Review"; }, 1800);
      } catch (err) {
        alert("Couldn't copy automatically — please select and copy the text above.");
      }
    });
  }
}
document.getElementById("reviewClose").addEventListener("click", () => {
  document.getElementById("reviewModal").classList.add("hidden");
  const url = new URL(window.location);
  url.searchParams.delete("review");
  window.history.replaceState({}, "", url);
});

/* -------------------------------------------------------------------------
   Restore session on page load
------------------------------------------------------------------------- */
(async function init() {
  initReviewFlow();

  let bailedOut = false;
  document.getElementById("authCheckingSkip").addEventListener("click", () => {
    bailedOut = true;
    clearSession();
    document.getElementById("authChecking").classList.add("hidden");
    document.getElementById("loginPanel").classList.remove("hidden");
  });

  const session = getSession();
  if (!session) return; // loginPanel is already visible — see the inline script in index.html
  const result = await apiCall("validateSession", { token: session.token });
  if (bailedOut) return; // user already gave up and reset the UI manually — don't override that
  if (result.success) {
    enterDashboard(session);
  } else {
    clearSession();
    document.getElementById("authChecking").classList.add("hidden");
    document.getElementById("loginPanel").classList.remove("hidden");
  }
})();
