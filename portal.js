/* Portal shell: sidebar, profile menu, page title, theme sync. Business logic stays in script.js */
(function () {
  const b = document.body, $ = s => document.querySelector(s);
  const setTitle = t => { $("#pageTitle").textContent = t; };
  const syncThemeUi = () => {
    const dark = document.documentElement.getAttribute("data-theme") === "dark";
    $("#menuTheme").textContent = dark ? "Switch to bright view" : "Switch to dark view";
  };

  window.portalEnter = function (u) {
    const member = u.role === "member";
    b.classList.add("in-app", member ? "role-member" : "role-staff");
    b.classList.remove(member ? "role-staff" : "role-member");
    const name = u.fullName || u.username || "";
    $("#pfName").textContent = name;
    $("#pfMail").textContent = u.email || u.username || "";
    $("#pfRole").textContent = String(u.role || "").toUpperCase();
    $("#pfAv").textContent = name.split(/\s+/).slice(0, 2).map(w => w[0] || "").join("").toUpperCase();
    const first = document.querySelector((member ? ".tabs-member" : ".tabs-staff") + " .dash-tab.active");
    setTitle(first ? first.textContent : "Dashboard");
    if (innerWidth < 800) b.classList.add("collapsed");
    syncThemeUi();
  };
  window.portalLeave = function () {
    b.classList.remove("in-app", "role-member", "role-staff");
    $("#menu").classList.add("hidden");
  };

  document.addEventListener("click", e => {
    const tab = e.target.closest(".dash-tab");
    if (tab) { setTitle(tab.textContent); if (innerWidth < 800) b.classList.add("collapsed"); }
    if (!e.target.closest(".pfw")) $("#menu").classList.add("hidden");
  });
  $("#pf").addEventListener("click", () => $("#menu").classList.toggle("hidden"));
  $("#sideToggle").addEventListener("click", () => b.classList.toggle("collapsed"));
  $("#menuTheme").addEventListener("click", () => { $("#themeToggle").click(); $("#menu").classList.add("hidden"); });
  document.querySelector('[data-open="Password"]').addEventListener("click", () => {
    $("#menu").classList.add("hidden");
    const t = document.querySelector(b.classList.contains("role-member") ? '[data-tab="memPasswordPanel"]' : '[data-tab="staffPasswordPanel"]');
    if (t) t.click();
  });
  // booking form lives in a modal so signed-in people can still request an appointment
  $("#menuBook").addEventListener("click", () => { $("#menu").classList.add("hidden"); $("#bookingModal").classList.remove("hidden"); });
  $("#bookingClose").addEventListener("click", () => $("#bookingModal").classList.add("hidden"));
  // keep charts readable when the theme flips
  $("#themeToggle").addEventListener("click", () => {
    syncThemeUi();
    if (b.classList.contains("in-app") && $("#staffAnalyticsPanel").classList.contains("active") && !$("#staffDashboard").classList.contains("hidden")) loadAnalytics();
  });
  syncThemeUi();
})();
