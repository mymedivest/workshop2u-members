/**
 * ==========================================================================
 * WORKSHOP2U — GOOGLE APPS SCRIPT BACKEND (full feature set)
 * ==========================================================================
 * Deploy as a Web App bound to the Google Sheet. Exposes one POST endpoint
 * routed by an "action" field. Run initializeSheet() once to create tabs.
 *
 * TABS (all created by initializeSheet):
 *   Credentials    Username | PasswordHash | Role | WorkshopLocation |
 *                  FullName | Address | Phone | Email | Status | CreatedDate
 *   History        Timestamp | Date | Username | Name | CustomerAddress |
 *                  VehicleType | VehiclePlateNumber | WorkshopLocation |
 *                  ServiceType | Price | Notes | ReviewToken
 *   Bookings       BookingID | Timestamp | Name | Phone | Email |
 *                  VehicleType | VehiclePlateNumber | PreferredDate |
 *                  PreferredTime | WorkshopLocation | ServiceType | Notes | Status
 *   Sessions       Token | Username | Role | WorkshopLocation | FullName |
 *                  CreatedAt | ExpiresAt
 *   Vehicles       Username | PlateNumber | VehicleType | Nickname
 *   LoginAttempts  Username | FailCount | LockUntil
 *   OtpCodes       Username | Code | CreatedAt | ExpiresAt
 *   PasswordResets Token | Username | CreatedAt | ExpiresAt
 *   Reviews        ReviewToken | Username | Workshop | Rating | Comment |
 *                  Status | CreatedAt | SubmittedAt
 *   AuditLog       Timestamp | Username | Role | Action | Details
 *   InvoiceQueue   HistoryID | Date | Name | Address | VehicleType |
 *                  PlateNumber | Workshop | ServiceType | Price | Notes |
 *                  CustomerEmail | ReviewToken | SiteUrl | QueuedAt
 *   Reminders      PlateNumber | Username | LastReminderSent
 * ==========================================================================
 */

// ---- Configuration ----------------------------------------------------------
const SESSION_LENGTH_HOURS = 12;
const OTP_LENGTH_MINUTES = 5;
const RESET_CODE_LENGTH_MINUTES = 30;
const MAX_LOGIN_ATTEMPTS = 5;
const LOCKOUT_MINUTES = 15;
const REMINDER_INTERVAL_DAYS = 180;   // remind customer ~6 months after last service
const REMINDER_LOOKAHEAD_DAYS = 7;    // send reminder when due date is within this window
const REMINDER_COOLDOWN_DAYS = 30;    // don't re-send a reminder more than once a month
const WORKSHOP_NOTIFY_EMAILS = {
  "Melaka": "melaka@workshop2u.com.my",
  "Negeri Sembilan": "ns@workshop2u.com.my",
  "Johor": "johor@workshop2u.com.my"
};

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) { return ss_().getSheetByName(name); }
function tz_() { return Session.getScriptTimeZone() || "Asia/Kuala_Lumpur"; }

// ---- Entry points -------------------------------------------------------------
// Actions that change data. A retry of any of these (e.g. the client's
// automatic retry after a slow/lost response) must not re-run the action —
// see MUTATION_CACHE_SECONDS below.
const MUTATING_ACTIONS = ["addHistory", "updateHistory", "addMember", "updateMemberStatus",
  "updateAccount", "bookAppointment", "updateBookingStatus", "submitReview",
  "changePassword", "resetPassword", "requestPasswordReset", "verifyOtp"];
const MUTATION_CACHE_SECONDS = 180; // comfortably longer than the client's worst-case retry window

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {
    return jsonOut_({ success: false, message: "Invalid request." });
  }
  const action = body.action;

  // Duplicate-request guard. The client sends one requestId per logical
  // action, reused across its own automatic retry (not regenerated), so a
  // retry of a slow-but-actually-successful request returns the cached
  // result instead of executing the mutation a second time (duplicate rows,
  // duplicate emails, etc.) — this is the general-purpose version of the
  // same fix already applied to login OTP and password reset.
  //
  // The check-cache / run-mutation / write-cache sequence is wrapped in a
  // lock. Without the lock, two near-simultaneous deliveries of the exact
  // same request (the client's own retry, or Apps Script's own
  // infrastructure occasionally re-delivering a slow request before its
  // first response is even back) can both see an empty cache and both
  // proceed to run the mutation, because "check" and "write" aren't atomic
  // on their own — that's how the same service record was saved twice
  // despite this guard already existing. The lock makes that sequence
  // atomic: whichever request gets there first runs and caches the result
  // before the second one is even allowed to check.
  const isMutation = MUTATING_ACTIONS.indexOf(action) !== -1;
  const cacheKey = isMutation && body.requestId ? "reqid_" + action + "_" + body.requestId : null;
  const lock = cacheKey ? LockService.getScriptLock() : null;

  if (lock) {
    try {
      lock.waitLock(30000);
    } catch (lockErr) {
      return jsonOut_({ success: false, message: "Server is busy right now — please try again in a moment." });
    }
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      lock.releaseLock();
      return jsonOut_(JSON.parse(cached));
    }
  }

  let result;
  try {
    switch (action) {
      case "login": result = login_(body); break;
      case "verifyOtp": result = verifyOtp_(body); break;
      case "validateSession": result = validateSession_(body.token); break;
      case "logout": result = logout_(body.token); break;
      case "requestPasswordReset": result = requestPasswordReset_(body); break;
      case "resetPassword": result = resetPassword_(body); break;
      case "changePassword": result = changePassword_(body); break;

      case "getHistory": result = getHistory_(body); break;
      case "addHistory": result = addHistory_(body); break;
      case "updateHistory": result = updateHistory_(body); break;

      case "getMembers": result = getMembers_(body); break;
      case "addMember": result = addMember_(body); break;
      case "updateMemberStatus": result = updateMemberStatus_(body); break;
      case "getVehicles": result = getVehicles_(body); break;

      case "getAccounts": result = getAccounts_(body); break;
      case "updateAccount": result = updateAccount_(body); break;

      case "bookAppointment": result = bookAppointment_(body); break;
      case "getBookings": result = getBookings_(body); break;
      case "updateBookingStatus": result = updateBookingStatus_(body); break;

      case "getAnalytics": result = getAnalytics_(body); break;
      case "getAuditLog": result = getAuditLog_(body); break;

      case "getReviewContext": result = getReviewContext_(body); break;
      case "submitReview": result = submitReview_(body); break;
      case "getReviews": result = getReviews_(body); break;

      default: result = { success: false, message: "Unknown action." };
    }
  } catch (err) {
    result = { success: false, message: "Server error: " + err.message };
  } finally {
    if (lock) {
      try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), MUTATION_CACHE_SECONDS); } catch (cacheErr) { /* best effort */ }
      lock.releaseLock();
    }
  }
  return jsonOut_(result);
}

function doGet(e) {
  return jsonOut_({ success: true, message: "Workshop2U API is running." });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---- Password hashing -----------------------------------------------------
function hashPassword_(plain) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain, Utilities.Charset.UTF_8);
  return digest.map(b => ("0" + (b & 0xFF).toString(16)).slice(-2)).join("");
}
function randomCode_(digits) {
  let code = "";
  for (let i = 0; i < digits; i++) code += Math.floor(Math.random() * 10);
  return code;
}

// ---- Audit log --------------------------------------------------------------
function logAudit_(username, role, action, details) {
  try {
    sheet_("AuditLog").appendRow([new Date(), username || "(guest)", role || "-", action, details || ""]);
  } catch (e) { /* never block main flow because of logging */ }
}

// ---- Credentials sheet helpers ---------------------------------------------
function credentialColumns_() {
  return ["Username","PasswordHash","Role","WorkshopLocation","FullName","Address","Phone","Email","Status","CreatedDate"];
}
function findCredentialRow_(username) {
  const sh = sheet_("Credentials");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) return { rowIndex: i + 1, values: data[i] };
  }
  return null;
}
function findCredentialByEmail_(email) {
  const sh = sheet_("Credentials");
  const data = sh.getDataRange().getValues();
  const cols = credentialColumns_();
  const emailIdx = cols.indexOf("Email");
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailIdx]).toLowerCase() === String(email).toLowerCase()) return { rowIndex: i + 1, values: data[i] };
  }
  return null;
}
function rowToCredential_(values) {
  const cols = credentialColumns_();
  const obj = {};
  cols.forEach((c, i) => obj[c] = values[i]);
  return obj;
}

// =============================================================================
// FEATURE 5 — RATE LIMITING & LOCKOUT
// =============================================================================
function findAttemptRow_(username) {
  const sh = sheet_("LoginAttempts");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) return { rowIndex: i + 1, values: data[i] };
  }
  return null;
}
function isLockedOut_(username) {
  const found = findAttemptRow_(username);
  if (!found) return { locked: false };
  const lockUntil = found.values[2];
  if (lockUntil && new Date(lockUntil).getTime() > Date.now()) {
    const mins = Math.ceil((new Date(lockUntil).getTime() - Date.now()) / 60000);
    return { locked: true, minutesLeft: mins };
  }
  return { locked: false };
}
function registerFailedAttempt_(username) {
  const sh = sheet_("LoginAttempts");
  const found = findAttemptRow_(username);
  if (!found) {
    sh.appendRow([username, 1, ""]);
    return;
  }
  let failCount = Number(found.values[1] || 0) + 1;
  let lockUntil = "";
  if (failCount >= MAX_LOGIN_ATTEMPTS) {
    lockUntil = new Date(Date.now() + LOCKOUT_MINUTES * 60 * 1000);
    failCount = 0;
  }
  sh.getRange(found.rowIndex, 2, 1, 2).setValues([[failCount, lockUntil]]);
}
function resetLoginAttempts_(username) {
  const found = findAttemptRow_(username);
  if (found) sheet_("LoginAttempts").getRange(found.rowIndex, 2, 1, 2).setValues([[0, ""]]);
}

// =============================================================================
// AUTH — login, OTP (feature 4), sessions
// =============================================================================
function login_(body) {
  const username = (body.username || "").trim();
  const password = body.password || "";
  if (!username || !password) return { success: false, message: "Username and password are required." };

  const lock = isLockedOut_(username);
  if (lock.locked) {
    return { success: false, message: "Too many failed attempts. Try again in " + lock.minutesLeft + " minute(s)." };
  }

  const found = findCredentialRow_(username);
  if (!found) { registerFailedAttempt_(username); return { success: false, message: "Invalid username or password." }; }
  const cred = rowToCredential_(found.values);

  if (String(cred.Status).toLowerCase() !== "active") {
    return { success: false, message: "This account is inactive. Please contact your workshop." };
  }
  if (hashPassword_(password) !== cred.PasswordHash) {
    registerFailedAttempt_(username);
    logAudit_(username, cred.Role, "LOGIN_FAILED", "Wrong password");
    return { success: false, message: "Invalid username or password." };
  }
  resetLoginAttempts_(username);

  // Members log straight in. Staff (admin/manager/webmaster) get a one-time
  // code emailed to them as a second factor.
  if (cred.Role === "member") {
    return { success: true, user: createSession_(cred) };
  }

  if (!cred.Email) {
    // No email on file — can't OTP, so log in directly rather than lock them out.
    logAudit_(username, cred.Role, "LOGIN_SUCCESS_NO_OTP", "No email on file, OTP skipped");
    return { success: true, user: createSession_(cred) };
  }

  // Duplicate-request guard: Apps Script Web Apps can occasionally receive
  // the same POST twice in quick succession (a platform quirk with how the
  // /exec endpoint serves its response). Without this guard that would send
  // two different OTP codes for one login click. We use a short-lived cache
  // lock so only the first request in any 20-second window actually
  // generates and emails a code — the (near-simultaneous) second request
  // just gets told a code is already on its way.
  const cache = CacheService.getScriptCache();
  const lockKey = "otp_lock_" + username.toLowerCase();
  if (cache.get(lockKey)) {
    return { success: true, otpRequired: true, username: cred.Username, message: "A code was already sent — check your email." };
  }
  cache.put(lockKey, "1", 20);

  const code = randomCode_(6);
  const sh = sheet_("OtpCodes");
  const data = sh.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).toLowerCase() === username.toLowerCase()) { rowIndex = i + 1; break; }
  const now = new Date();
  const expires = new Date(now.getTime() + OTP_LENGTH_MINUTES * 60 * 1000);
  if (rowIndex > -1) sh.getRange(rowIndex, 1, 1, 4).setValues([[username, code, now, expires]]);
  else sh.appendRow([username, code, now, expires]);

  try {
    MailApp.sendEmail({
      to: cred.Email,
      subject: "Workshop2U — Your login verification code",
      body: "Your one-time login code is: " + code + "\n\nThis code expires in " + OTP_LENGTH_MINUTES + " minutes.\n\nIf you didn't try to log in, you can ignore this email."
    });
  } catch (e) { /* best effort */ }

  logAudit_(username, cred.Role, "LOGIN_OTP_SENT", "");
  return { success: true, otpRequired: true, username: cred.Username };
}

function verifyOtp_(body) {
  const username = (body.username || "").trim();
  const code = (body.code || "").trim();
  if (!username || !code) return { success: false, message: "Enter the code sent to your email." };

  const sh = sheet_("OtpCodes");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase()) {
      const [ , storedCode, , expiresAt ] = data[i];
      if (new Date(expiresAt).getTime() < Date.now()) {
        sh.deleteRow(i + 1);
        return { success: false, message: "Code expired. Please log in again." };
      }
      if (String(storedCode) !== code) {
        logAudit_(username, "-", "OTP_FAILED", "");
        return { success: false, message: "Incorrect code." };
      }
      sh.deleteRow(i + 1);
      const found = findCredentialRow_(username);
      if (!found) return { success: false, message: "Account not found." };
      const cred = rowToCredential_(found.values);
      logAudit_(username, cred.Role, "LOGIN_SUCCESS", "OTP verified");
      return { success: true, user: createSession_(cred) };
    }
  }
  return { success: false, message: "No pending code for this account. Please log in again." };
}

function createSession_(cred) {
  const token = Utilities.getUuid();
  const now = new Date();
  const expires = new Date(now.getTime() + SESSION_LENGTH_HOURS * 60 * 60 * 1000);
  sheet_("Sessions").appendRow([token, cred.Username, cred.Role, cred.WorkshopLocation, cred.FullName, now, expires]);
  return {
    token, username: cred.Username, role: cred.Role, workshop: cred.WorkshopLocation,
    fullName: cred.FullName, address: cred.Address, phone: cred.Phone, email: cred.Email
  };
}

function getSessionRow_(token) {
  const sh = sheet_("Sessions");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (data[i][0] === token) return { rowIndex: i + 1, values: data[i] };
  return null;
}
function validateSession_(token) {
  if (!token) return { success: false, message: "No session." };
  const found = getSessionRow_(token);
  if (!found) return { success: false, message: "Session not found." };
  const [ , username, role, workshop, fullName, , expiresAt ] = found.values;
  if (new Date(expiresAt).getTime() < Date.now()) {
    sheet_("Sessions").deleteRow(found.rowIndex);
    return { success: false, message: "Session expired." };
  }
  return { success: true, user: { token, username, role, workshop, fullName } };
}
function logout_(token) {
  const found = getSessionRow_(token);
  if (found) { logAudit_(found.values[1], found.values[2], "LOGOUT", ""); sheet_("Sessions").deleteRow(found.rowIndex); }
  return { success: true };
}
function requireAuth_(token, allowedRoles) {
  const check = validateSession_(token);
  if (!check.success) return { ok: false, error: check.message };
  if (allowedRoles && allowedRoles.indexOf(check.user.role) === -1) return { ok: false, error: "You do not have permission to do this." };
  return { ok: true, user: check.user };
}

// =============================================================================
// FEATURE 6 — SELF-SERVICE PASSWORD RESET (code-based, no external link needed)
// =============================================================================
function requestPasswordReset_(body) {
  const identifier = (body.usernameOrEmail || "").trim();
  if (!identifier) return { success: false, message: "Enter your username or email." };

  let found = findCredentialRow_(identifier) || findCredentialByEmail_(identifier);
  // Always respond the same way whether or not the account exists, to avoid
  // leaking which usernames/emails are registered.
  const genericMsg = "If that account exists, a reset code has been emailed to it.";
  if (!found) return { success: true, message: genericMsg };

  const cred = rowToCredential_(found.values);
  if (!cred.Email) return { success: true, message: genericMsg };

  const cache = CacheService.getScriptCache();
  const lockKey = "reset_lock_" + cred.Username.toLowerCase();
  if (cache.get(lockKey)) return { success: true, message: genericMsg };
  cache.put(lockKey, "1", 20);

  const code = randomCode_(6);
  const now = new Date();
  const expires = new Date(now.getTime() + RESET_CODE_LENGTH_MINUTES * 60 * 1000);
  sheet_("PasswordResets").appendRow([code, cred.Username, now, expires]);

  try {
    MailApp.sendEmail({
      to: cred.Email,
      subject: "Workshop2U — Password reset code",
      body: "Your password reset code is: " + code + "\n\nEnter this code on the Reset Password screen along with your username (" + cred.Username + ") and a new password. This code expires in " + RESET_CODE_LENGTH_MINUTES + " minutes."
    });
  } catch (e) { /* best effort */ }

  logAudit_(cred.Username, cred.Role, "PASSWORD_RESET_REQUESTED", "");
  return { success: true, message: genericMsg };
}

function resetPassword_(body) {
  const username = (body.username || "").trim();
  const code = (body.code || "").trim();
  const newPassword = body.newPassword || "";
  if (!username || !code || !newPassword) return { success: false, message: "All fields are required." };
  if (newPassword.length < 6) return { success: false, message: "New password must be at least 6 characters." };

  const sh = sheet_("PasswordResets");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === code && String(data[i][1]).toLowerCase() === username.toLowerCase()) {
      if (new Date(data[i][3]).getTime() < Date.now()) { sh.deleteRow(i + 1); return { success: false, message: "Code expired. Please request a new one." }; }
      sh.deleteRow(i + 1);
      const found = findCredentialRow_(username);
      if (!found) return { success: false, message: "Account not found." };
      const cols = credentialColumns_();
      sheet_("Credentials").getRange(found.rowIndex, cols.indexOf("PasswordHash") + 1).setValue(hashPassword_(newPassword));
      resetLoginAttempts_(username);
      logAudit_(username, "-", "PASSWORD_RESET_COMPLETE", "");
      return { success: true };
    }
  }
  return { success: false, message: "Invalid or expired code." };
}

function changePassword_(body) {
  const auth = requireAuth_(body.token, null);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const found = findCredentialRow_(user.username);
  if (!found) return { success: false, message: "Account not found." };
  const cred = rowToCredential_(found.values);
  if (hashPassword_(body.oldPassword || "") !== cred.PasswordHash) return { success: false, message: "Current password is incorrect." };
  if (!body.newPassword || body.newPassword.length < 6) return { success: false, message: "New password must be at least 6 characters." };

  const cols = credentialColumns_();
  sheet_("Credentials").getRange(found.rowIndex, cols.indexOf("PasswordHash") + 1).setValue(hashPassword_(body.newPassword));
  logAudit_(user.username, user.role, "PASSWORD_CHANGED", "");
  return { success: true };
}

// =============================================================================
// HISTORY  (+ FEATURE 9 vehicles, + FEATURE 1 reminders, + FEATURE 3 invoice email,
//            + review-request email)
// =============================================================================
function historyColumns_() {
  return ["Timestamp","Date","Username","Name","CustomerAddress","VehicleType","VehiclePlateNumber","WorkshopLocation","ServiceType","Price","Notes","ReviewToken","HistoryID"];
}
function formatDate_(value) {
  if (!value) return "";
  const d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, tz_(), "yyyy-MM-dd");
}

/* Fills in a HistoryID for any History row that doesn't have one yet, and
   labels the column header if it's blank. Uses a lock so two people loading
   the page at the same moment can't hand the same row two different IDs. */
function ensureHistoryIds_() {
  const sh = sheet_("History");
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const idCol = historyColumns_().indexOf("HistoryID") + 1;
  if (sh.getMaxColumns() < idCol) sh.insertColumnsAfter(sh.getMaxColumns(), idCol - sh.getMaxColumns());

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const idRange = sh.getRange(2, idCol, lastRow - 1, 1);
    const ids = idRange.getValues();          // re-read inside the lock
    let changed = false;
    ids.forEach(row => { if (!row[0]) { row[0] = Utilities.getUuid(); changed = true; } });
    if (changed) idRange.setValues(ids);
    if (!sh.getRange(1, idCol).getValue()) sh.getRange(1, idCol).setValue("HistoryID");
  } finally {
    lock.releaseLock();
  }
}

function getHistory_(body) {
  const auth = requireAuth_(body.token, null);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const sh = sheet_("History");
  let data = sh.getDataRange().getValues();
  // Rows saved before HistoryID existed have no ID, so they'd have no Edit
  // button. The first time staff load the history after upgrading, give
  // every such row an ID (one-off — after that this check finds nothing).
  if (user.role !== "member" && data.slice(1).some(r => !r[12])) {
    ensureHistoryIds_();
    data = sh.getDataRange().getValues();
  }
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const record = {
      id: r[12], date: formatDate_(r[1]), username: r[2], name: r[3], address: r[4],
      vehicleType: r[5], plate: r[6], workshop: r[7], serviceType: r[8], price: r[9], notes: r[10]
    };
    if (user.role === "member" && record.username !== user.username) continue;
    if (user.role === "admin" && record.workshop !== user.workshop) continue;
    if ((user.role === "manager" || user.role === "webmaster") && body.workshop && body.workshop !== "All" && record.workshop !== body.workshop) continue;
    if (body.plate && body.plate !== "All" && record.plate !== body.plate) continue;
    rows.push(record);
  }
  rows.sort((a, b) => new Date(b.date) - new Date(a.date));

  const result = { success: true, history: rows };
  if (user.role === "member" && rows.length) {
    const lastDate = new Date(rows[0].date);
    const nextDue = new Date(lastDate.getTime() + REMINDER_INTERVAL_DAYS * 24 * 60 * 60 * 1000);
    result.nextServiceDue = formatDate_(nextDue);
  }
  return result;
}

function addHistory_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;
  const rec = body.record || {};

  const workshop = user.role === "admin" ? user.workshop : rec.workshop;
  if (!rec.date || !rec.name || !rec.vehicleType || !rec.plate || !workshop || !rec.serviceType || rec.price === undefined) {
    return { success: false, message: "Missing required fields." };
  }

  const reviewToken = Utilities.getUuid();
  const historyId = Utilities.getUuid();
  sheet_("History").appendRow([
    new Date(), rec.date, rec.username || "", rec.name, rec.address || "",
    rec.vehicleType, rec.plate, workshop, rec.serviceType, Number(rec.price) || 0, rec.notes || "", reviewToken, historyId
  ]);

  // Feature 9 — keep a distinct vehicle registry per member.
  if (rec.username) addVehicleIfNew_(rec.username, rec.plate, rec.vehicleType);

  logAudit_(user.username, user.role, "HISTORY_ADDED", rec.plate + " @ " + workshop);

  // Feature 3 — queue the invoice email (PDF generation + send) instead of
  // doing it here. Building the PDF and emailing it can genuinely take a
  // while, and doing that synchronously meant a slow-but-working save could
  // exceed the browser's timeout, trigger an automatic retry, and duplicate
  // the whole record. processInvoiceQueue() (run every minute by a time
  // trigger — see README) does the actual sending shortly after.
  try {
    let customerEmail = rec.email || "";
    if (!customerEmail && rec.username) {
      const credRow = findCredentialRow_(rec.username);
      if (credRow) customerEmail = rowToCredential_(credRow.values).Email;
    }
    if (customerEmail) {
      sheet_("InvoiceQueue").appendRow([
        historyId, rec.date, rec.name, rec.address || "", rec.vehicleType, rec.plate,
        workshop, rec.serviceType, Number(rec.price) || 0, rec.notes || "",
        customerEmail, reviewToken, body.siteUrl || "", new Date()
      ]);
    }
  } catch (e) { /* best effort — don't fail the save if queuing the invoice fails */ }

  return { success: true, id: historyId };
}

/* Sends every invoice currently in the queue, then clears those rows. Meant
   to run on a time-based trigger every minute or so (see README) — each run
   only has to deal with whatever was added since the last run, typically
   nothing or one record. A lock keeps two overlapping runs from racing. */
function processInvoiceQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // another run is already processing; skip this tick
  try {
    const sh = sheet_("InvoiceQueue");
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;

    const sentRows = [];
    for (let i = 1; i < data.length; i++) {
      const [historyId, date, name, address, vehicleType, plate, workshop, serviceType, price, notes, email, reviewToken, siteUrl] = data[i];
      try {
        const rec = { date, name, address, vehicleType, plate, serviceType, price, notes };
        const pdfBlob = generateReceiptPdf_(rec, workshop);
        const reviewLink = (siteUrl ? siteUrl : "") + "?review=" + reviewToken;
        MailApp.sendEmail({
          to: email,
          subject: "Workshop2U — Your service invoice",
          body: "Hi " + name + ",\n\nThanks for servicing your vehicle (" + plate + ") with us at " + workshop + ".\n" +
                "Service: " + serviceType + "\nAmount: RM " + Number(price).toFixed(2) + "\n\n" +
                "Your invoice is attached as a PDF.\n\n" +
                (siteUrl ? "We'd love your feedback — rate your visit here:\n" + reviewLink + "\n\n" : "") +
                "Thank you for choosing Workshop2U.",
          attachments: [pdfBlob]
        });
        sentRows.push(i + 1);
      } catch (e) {
        // Leave this row in the queue — it'll be retried on the next run
        // rather than silently lost. If a row keeps failing (e.g. a bad
        // email address), it'll stay queued; worth an occasional glance.
      }
    }
    // Delete from the bottom up so row numbers of not-yet-deleted rows don't shift.
    sentRows.sort((a, b) => b - a).forEach(rowNum => sh.deleteRow(rowNum));
  } finally {
    lock.releaseLock();
  }
}

/* Editing an existing record — deliberately does NOT re-send the invoice
   email, re-generate a review token, or touch the Reviews tab. Those only
   happen once, at creation, in addHistory_ above. This just corrects the
   stored fields (e.g. a typo'd price or plate number). */
function updateHistory_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;
  const rec = body.record || {};
  if (!body.historyId) return { success: false, message: "Missing record id." };
  if (!rec.date || !rec.name || !rec.vehicleType || !rec.plate || !rec.serviceType || rec.price === undefined) {
    return { success: false, message: "Missing required fields." };
  }

  const sh = sheet_("History");
  const data = sh.getDataRange().getValues();
  const cols = historyColumns_();
  const idIdx = cols.indexOf("HistoryID");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] !== body.historyId) continue;
    const currentWorkshop = data[i][cols.indexOf("WorkshopLocation")];
    if (user.role === "admin" && currentWorkshop !== user.workshop) return { success: false, message: "Not your workshop." };
    const workshop = user.role === "admin" ? user.workshop : (rec.workshop || currentWorkshop);

    sh.getRange(i + 1, cols.indexOf("Date") + 1, 1, 1).setValue(rec.date);
    sh.getRange(i + 1, cols.indexOf("Name") + 1, 1, 1).setValue(rec.name);
    sh.getRange(i + 1, cols.indexOf("CustomerAddress") + 1, 1, 1).setValue(rec.address || "");
    sh.getRange(i + 1, cols.indexOf("VehicleType") + 1, 1, 1).setValue(rec.vehicleType);
    sh.getRange(i + 1, cols.indexOf("VehiclePlateNumber") + 1, 1, 1).setValue(rec.plate);
    sh.getRange(i + 1, cols.indexOf("WorkshopLocation") + 1, 1, 1).setValue(workshop);
    sh.getRange(i + 1, cols.indexOf("ServiceType") + 1, 1, 1).setValue(rec.serviceType);
    sh.getRange(i + 1, cols.indexOf("Price") + 1, 1, 1).setValue(Number(rec.price) || 0);
    sh.getRange(i + 1, cols.indexOf("Notes") + 1, 1, 1).setValue(rec.notes || "");
    if (rec.username !== undefined) sh.getRange(i + 1, cols.indexOf("Username") + 1, 1, 1).setValue(rec.username);

    logAudit_(user.username, user.role, "HISTORY_EDITED", body.historyId + " (" + rec.plate + " @ " + workshop + ")");
    return { success: true };
  }
  return { success: false, message: "Record not found." };
}

/* Builds the invoice as HTML and converts it straight to a PDF blob —
   Apps Script's Blob conversion service supports this directly. This
   replaced an earlier version that created a real Google Doc, saved it,
   exported it, then deleted it (DocumentApp + DriveApp, 4+ separate API
   round trips) purely to get PDF bytes we never needed to keep as a Doc.
   Same output, no Drive usage, and meaningfully faster per invoice — which
   matters because generateReceiptPdf_ runs once per row inside
   processInvoiceQueue_'s trigger run, so a faster PDF here means the queue
   drains faster and a shorter trigger interval (see README) is practical. */
function generateReceiptPdf_(rec, workshop) {
  const esc = escapeHtml_;
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"></head>
<body style="font-family:Arial,Helvetica,sans-serif;color:#14181d;padding:32px;">
  <h1 style="margin:0 0 2px;letter-spacing:.04em;">WORKSHOP2U</h1>
  <h2 style="margin:0 0 22px;color:#5b6472;font-weight:normal;">Service Invoice</h2>
  <p style="margin:0 0 14px;"><strong>Workshop:</strong> ${esc(workshop)}<br><strong>Date:</strong> ${esc(rec.date)}</p>
  <p style="margin:0 0 14px;"><strong>Customer:</strong> ${esc(rec.name)}<br>
     <strong>Address:</strong> ${esc(rec.address || "-")}<br>
     <strong>Vehicle:</strong> ${esc(rec.vehicleType)} (${esc(rec.plate)})</p>
  <p style="margin:0 0 4px;"><strong>Service type:</strong> ${esc(rec.serviceType)}</p>
  <p style="margin:0 0 14px;font-size:1.3em;"><strong>Amount: RM ${Number(rec.price).toFixed(2)}</strong></p>
  ${rec.notes ? `<p style="margin:0 0 14px;"><strong>Notes:</strong> ${esc(rec.notes)}</p>` : ""}
  <p style="margin-top:26px;color:#5b6472;">Thank you for choosing Workshop2U — We pick up, fix &amp; deliver.</p>
</body></html>`;
  const pdfName = "Invoice-" + rec.plate + "-" + rec.date + ".pdf";
  return Utilities.newBlob(html, "text/html", pdfName).getAs("application/pdf").setName(pdfName);
}

function escapeHtml_(str) {
  return String(str === undefined || str === null ? "" : str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Feature 1 helper — called by addHistory_
function addVehicleIfNew_(username, plate, vehicleType) {
  const sh = sheet_("Vehicles");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase() && String(data[i][1]).toLowerCase() === plate.toLowerCase()) return;
  }
  sh.appendRow([username, plate, vehicleType, ""]);
}

function getVehicles_(body) {
  const auth = requireAuth_(body.token, null);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;
  const username = user.role === "member" ? user.username : (body.username || "");
  if (!username) return { success: false, message: "No username specified." };

  const sh = sheet_("Vehicles");
  const data = sh.getDataRange().getValues();
  const vehicles = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase()) {
      vehicles.push({ plate: data[i][1], vehicleType: data[i][2], nickname: data[i][3] });
    }
  }
  return { success: true, vehicles };
}

// =============================================================================
// FEATURE 1 (continued) — scheduled reminder job. Attach a daily time trigger
// to this function from Apps Script: Triggers > Add Trigger > sendServiceReminders.
// =============================================================================
function sendServiceReminders() {
  const history = sheet_("History").getDataRange().getValues();
  const latestByVehicle = {}; // key: username|plate -> {date, name, vehicleType, plate}
  for (let i = 1; i < history.length; i++) {
    const r = history[i];
    const username = r[2], name = r[3], vehicleType = r[5], plate = r[6];
    if (!username || !plate) continue;
    const key = username + "|" + plate;
    const d = new Date(r[1]);
    if (!latestByVehicle[key] || d > latestByVehicle[key].date) {
      latestByVehicle[key] = { date: d, name, vehicleType, plate, username };
    }
  }

  const reminderSheet = sheet_("Reminders");
  const reminderData = reminderSheet.getDataRange().getValues();
  const lastSent = {}; // key: plate|username -> Date
  for (let i = 1; i < reminderData.length; i++) {
    lastSent[reminderData[i][0] + "|" + reminderData[i][1]] = new Date(reminderData[i][2]);
  }

  const now = new Date();
  Object.keys(latestByVehicle).forEach(key => {
    const v = latestByVehicle[key];
    const nextDue = new Date(v.date.getTime() + REMINDER_INTERVAL_DAYS * 24 * 60 * 60 * 1000);
    const daysUntilDue = (nextDue.getTime() - now.getTime()) / (24 * 60 * 60 * 1000);
    if (daysUntilDue > REMINDER_LOOKAHEAD_DAYS || daysUntilDue < -30) return; // out of window

    const cooldownKey = v.plate + "|" + v.username;
    const prevSent = lastSent[cooldownKey];
    if (prevSent && (now.getTime() - prevSent.getTime()) < REMINDER_COOLDOWN_DAYS * 24 * 60 * 60 * 1000) return;

    const credRow = findCredentialRow_(v.username);
    if (!credRow) return;
    const cred = rowToCredential_(credRow.values);
    if (!cred.Email) return;

    try {
      MailApp.sendEmail({
        to: cred.Email,
        subject: "Workshop2U — Your vehicle's service is due soon",
        body: "Hi " + v.name + ",\n\nYour vehicle " + v.vehicleType + " (" + v.plate + ") is due for its next service around " +
              formatDate_(nextDue) + ".\n\nBook your pickup today: https://workshop2u.com.my\n\nThank you for choosing Workshop2U."
      });
      reminderSheet.appendRow([v.plate, v.username, now]);
      logAudit_(v.username, "member", "REMINDER_SENT", v.plate);
    } catch (e) { /* best effort */ }
  });
}

// =============================================================================
// MEMBERS
// =============================================================================
function getMembers_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const data = sheet_("Credentials").getDataRange().getValues();
  const members = [];
  for (let i = 1; i < data.length; i++) {
    const cred = rowToCredential_(data[i]);
    if (cred.Role !== "member") continue;
    if (user.role === "admin" && cred.WorkshopLocation !== user.workshop) continue;
    if ((user.role === "manager" || user.role === "webmaster") && body.workshop && body.workshop !== "All" && cred.WorkshopLocation !== body.workshop) continue;
    members.push({ username: cred.Username, fullName: cred.FullName, address: cred.Address, phone: cred.Phone, email: cred.Email, workshop: cred.WorkshopLocation, status: cred.Status });
  }
  return { success: true, members };
}

function addMember_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;
  const m = body.member || {};

  if (!m.username || !m.password || !m.fullName) return { success: false, message: "Username, password and full name are required." };
  if (findCredentialRow_(m.username)) return { success: false, message: "That username already exists." };

  let role = m.role || "member";
  let workshop = m.workshop || user.workshop;
  if (user.role === "admin") { role = "member"; workshop = user.workshop; }

  sheet_("Credentials").appendRow([m.username, hashPassword_(m.password), role, workshop, m.fullName, m.address || "", m.phone || "", m.email || "", "Active", new Date()]);
  logAudit_(user.username, user.role, "ACCOUNT_CREATED", m.username + " (" + role + ")");
  return { success: true };
}

/**
 * Deactivate/reactivate a MEMBER (customer) account — this is the "remove
 * member" action available directly from the Members tab to admin/manager/
 * webmaster, not just the webmaster-only account manager. We deliberately
 * deactivate rather than delete the row, so existing History/Vehicles/
 * Reviews records tied to that username stay intact. Only touches accounts
 * with Role = "member" — staff accounts must still go through the
 * webmaster's "Manage Staff Accounts" tab (updateAccount_).
 */
function updateMemberStatus_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const found = findCredentialRow_(body.username);
  if (!found) return { success: false, message: "Member not found." };
  const cred = rowToCredential_(found.values);
  if (cred.Role !== "member") return { success: false, message: "This action only applies to member accounts." };
  if (user.role === "admin" && cred.WorkshopLocation !== user.workshop) return { success: false, message: "Not your workshop." };

  const cols = credentialColumns_();
  sheet_("Credentials").getRange(found.rowIndex, cols.indexOf("Status") + 1).setValue(body.status);
  logAudit_(user.username, user.role, "MEMBER_STATUS_UPDATED", body.username + " -> " + body.status);
  return { success: true };
}

// =============================================================================
// WEBMASTER — account management  (+ FEATURE 11 audit log)
// =============================================================================
function getAccounts_(body) {
  const auth = requireAuth_(body.token, ["webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const data = sheet_("Credentials").getDataRange().getValues();
  const accounts = [];
  for (let i = 1; i < data.length; i++) {
    const cred = rowToCredential_(data[i]);
    if (cred.Role === "member") continue; // members belong on the Members tab, not here
    accounts.push({ username: cred.Username, role: cred.Role, workshop: cred.WorkshopLocation, fullName: cred.FullName, status: cred.Status });
  }
  return { success: true, accounts };
}

function updateAccount_(body) {
  const auth = requireAuth_(body.token, ["webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const found = findCredentialRow_(body.username);
  if (!found) return { success: false, message: "Account not found." };
  const sh = sheet_("Credentials");
  const cols = credentialColumns_();
  const updates = body.updates || {};

  if (updates.status) sh.getRange(found.rowIndex, cols.indexOf("Status") + 1).setValue(updates.status);
  if (updates.password) sh.getRange(found.rowIndex, cols.indexOf("PasswordHash") + 1).setValue(hashPassword_(updates.password));
  if (updates.role) sh.getRange(found.rowIndex, cols.indexOf("Role") + 1).setValue(updates.role);
  if (updates.workshop) sh.getRange(found.rowIndex, cols.indexOf("WorkshopLocation") + 1).setValue(updates.workshop);

  logAudit_(auth.user.username, "webmaster", "ACCOUNT_UPDATED", body.username + " -> " + JSON.stringify(updates));
  return { success: true };
}

function getAuditLog_(body) {
  const auth = requireAuth_(body.token, ["webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const data = sheet_("AuditLog").getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    rows.push({ timestamp: Utilities.formatDate(new Date(data[i][0]), tz_(), "yyyy-MM-dd HH:mm"), username: data[i][1], role: data[i][2], action: data[i][3], details: data[i][4] });
  }
  rows.reverse();
  return { success: true, log: rows.slice(0, 300) };
}

// =============================================================================
// BOOKINGS  (+ FEATURE 2 approval workflow, + FEATURE 13 spam guard)
// =============================================================================
function bookAppointment_(body) {
  const b = body.booking || {};

  // Feature 13 — honeypot field. Real users never fill this hidden input;
  // bots that auto-fill every field will, so we silently "succeed" without saving.
  if (b.website) return { success: true };

  if (!b.name || !b.phone || !b.email || !b.workshop || !b.date || !b.time) {
    return { success: false, message: "Please complete all required fields." };
  }

  const bookingId = Utilities.getUuid();
  sheet_("Bookings").appendRow([bookingId, new Date(), b.name, b.phone, b.email, b.vehicleType || "", b.plate || "", b.date, b.time, b.workshop, b.serviceType || "", b.notes || "", "New"]);
  logAudit_("(guest)", "-", "BOOKING_CREATED", b.name + " / " + b.plate + " @ " + b.workshop);

  try {
    MailApp.sendEmail({
      to: b.email,
      subject: "Workshop2U — Appointment Request Received",
      body: "Hi " + b.name + ",\n\nThanks for booking with Workshop2U (" + b.workshop + ").\n" +
            "Requested date/time: " + b.date + " " + b.time + "\nService: " + (b.serviceType || "-") + "\n\n" +
            "We will contact you shortly to confirm.\n\nWorkshop2U"
    });
  } catch (e) { /* ignore */ }
  notifyManagers_(b.workshop, "New Booking Request — " + b.workshop,
    "A new appointment request was received:\n\n" +
    bookingSummary_({ name: b.name, phone: b.phone, email: b.email, vehicleType: b.vehicleType, plate: b.plate, workshop: b.workshop, date: b.date, time: b.time, serviceType: b.serviceType, notes: b.notes }) +
    "\n\nPlease log in to the portal > Bookings to confirm, reject or reschedule.\n\nWorkshop2U");

  return { success: true };
}

// Turns a Bookings "PreferredTime" cell into plain "HH:mm" (Malaysia time).
// Sheets stores a typed time like 09:00 as a Date in 1899, which must never be shown raw.
function bookingTime_(raw, display) {
  const m = String(display || "").match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?/i);
  if (m) {
    let h = parseInt(m[1], 10);
    const ap = (m[3] || "").toUpperCase();
    if (ap === "PM" && h < 12) h += 12;
    if (ap === "AM" && h === 12) h = 0;
    return (h < 10 ? "0" : "") + h + ":" + m[2];
  }
  if (raw instanceof Date) return Utilities.formatDate(raw, "Asia/Kuala_Lumpur", "HH:mm");
  return String(raw || "");
}

// ---- Booking notifications to the workshop's manager(s) ----------------------
const BOOKING_TZ = "Asia/Kuala_Lumpur";

/* Emails of active manager/admin accounts whose WorkshopLocation matches the
   booking's workshop (e.g. a Melaka booking -> the Melaka manager), plus the
   shared WORKSHOP_NOTIFY_EMAILS address if one is configured. */
function managerEmails_(workshop) {
  const out = {};
  try {
    const data = sheet_("Credentials").getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const c = rowToCredential_(data[i]);
      const role = String(c.Role || "").toLowerCase();
      if (role !== "manager" && role !== "admin") continue;
      if (String(c.Status || "").toLowerCase() !== "active") continue;
      if (!c.Email || c.WorkshopLocation !== workshop) continue;
      out[String(c.Email).trim().toLowerCase()] = String(c.Email).trim();
    }
  } catch (e) { /* best effort */ }
  const shared = WORKSHOP_NOTIFY_EMAILS[workshop];
  if (shared) out[String(shared).trim().toLowerCase()] = shared;
  return Object.keys(out).map(k => out[k]);
}

function notifyManagers_(workshop, subject, body) {
  managerEmails_(workshop).forEach(to => {
    try { MailApp.sendEmail({ to: to, subject: subject, body: body }); } catch (e) { /* ignore */ }
  });
}

function bookingSummary_(d) {
  return "Customer: " + d.name + "\nPhone: " + (d.phone || "-") + "\nEmail: " + (d.email || "-") +
         "\nVehicle: " + (d.vehicleType || "-") + " (" + (d.plate || "-") + ")" +
         "\nWorkshop: " + d.workshop + "\nRequested date/time: " + d.date + " " + d.time +
         "\nService: " + (d.serviceType || "-") + (d.notes ? "\nNotes: " + d.notes : "");
}

/* Builds the real appointment moment (Malaysia time) from the date + time cells. */
function bookingMoment_(dateCell, timeRaw, timeShown) {
  const d = formatDate_(dateCell), t = bookingTime_(timeRaw, timeShown);
  if (!d || !/^\d{2}:\d{2}$/.test(t)) return null;
  try { return Utilities.parseDate(d + " " + t, BOOKING_TZ, "yyyy-MM-dd HH:mm"); } catch (e) { return null; }
}

/* ---------------------------------------------------------------------------
   24-HOUR BOOKING REMINDERS. Run installBookingReminderTrigger() ONCE from the
   Apps Script editor; it creates an hourly trigger for sendBookingReminders().
   Only CONFIRMED bookings get a reminder, once each (tracked in the Bookings
   sheet's column N "ReminderSent"; rescheduling clears it so a new reminder
   goes out for the new time).
--------------------------------------------------------------------------- */
function installBookingReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "sendBookingReminders") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("sendBookingReminders").timeBased().everyHours(1).create();
}

function sendBookingReminders() {
  const sh = sheet_("Bookings");
  const range = sh.getDataRange();
  const data = range.getValues(), shown = range.getDisplayValues();
  if (data.length < 2) return;
  if (!String(data[0][13] || "").trim()) sh.getRange(1, 14).setValue("ReminderSent");
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[12]).toLowerCase() !== "confirmed" || r[13]) continue;
    const when = bookingMoment_(r[7], r[8], shown[i][8]);
    if (!when) continue;
    const hours = (when.getTime() - now.getTime()) / 3600000;
    if (hours <= 0 || hours > 24) continue;
    const d = { name: r[2], phone: r[3], email: r[4], vehicleType: r[5], plate: r[6], date: formatDate_(r[7]), time: bookingTime_(r[8], shown[i][8]), workshop: r[9], serviceType: r[10], notes: r[11] };
    try {
      if (d.email) MailApp.sendEmail({
        to: d.email,
        subject: "Workshop2U — Reminder: your appointment is within 24 hours",
        body: "Hi " + d.name + ",\n\nThis is a friendly reminder of your upcoming appointment with Workshop2U (" + d.workshop + ").\n" +
              "Date/time: " + d.date + " " + d.time + "\nService: " + (d.serviceType || "-") + "\n\nSee you then!\n\nWorkshop2U"
      });
    } catch (e) { /* ignore */ }
    notifyManagers_(d.workshop, "Reminder — booking within 24 hours (" + d.workshop + ")",
      "Upcoming confirmed appointment:\n\n" + bookingSummary_(d) + "\n\nWorkshop2U");
    sh.getRange(i + 1, 14).setValue(new Date());
    logAudit_("(system)", "-", "BOOKING_REMINDER_SENT", r[0] + " @ " + d.workshop);
  }
}

function bookingColumns_() {
  return ["BookingID","Timestamp","Name","Phone","Email","VehicleType","VehiclePlateNumber","PreferredDate","PreferredTime","WorkshopLocation","ServiceType","Notes","Status"];
}

function getBookings_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const range = sheet_("Bookings").getDataRange();
  const data = range.getValues();
  const shown = range.getDisplayValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const booking = { id: r[0], name: r[2], phone: r[3], email: r[4], vehicleType: r[5], plate: r[6], date: formatDate_(r[7]) || r[7], time: bookingTime_(r[8], shown[i][8]), workshop: r[9], serviceType: r[10], notes: r[11], status: r[12] };
    if (user.role === "admin" && booking.workshop !== user.workshop) continue;
    if ((user.role === "manager" || user.role === "webmaster") && body.workshop && body.workshop !== "All" && booking.workshop !== body.workshop) continue;
    rows.push(booking);
  }
  rows.sort((a, b) => new Date(b.date) - new Date(a.date));
  return { success: true, bookings: rows };
}

function updateBookingStatus_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const sh = sheet_("Bookings");
  const data = sh.getDataRange().getValues();
  const shown = sh.getDataRange().getDisplayValues();
  const cols = bookingColumns_();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.bookingId) {
      const workshop = data[i][cols.indexOf("WorkshopLocation")];
      if (user.role === "admin" && workshop !== user.workshop) return { success: false, message: "Not your workshop." };

      if (body.status) sh.getRange(i + 1, cols.indexOf("Status") + 1).setValue(body.status);
      if (body.newDate) sh.getRange(i + 1, cols.indexOf("PreferredDate") + 1).setValue(body.newDate);
      if (body.newTime) sh.getRange(i + 1, cols.indexOf("PreferredTime") + 1).setValue(body.newTime);
      if (body.newDate || body.newTime) sh.getRange(i + 1, 14).clearContent(); // new time -> new reminder

      const email = data[i][cols.indexOf("Email")];
      const name = data[i][cols.indexOf("Name")];
      const bkWorkshop = data[i][cols.indexOf("WorkshopLocation")];
      const bkService = data[i][cols.indexOf("ServiceType")] || "-";
      // Show the date/time the booking now has (new values if it was just rescheduled).
      const bkDate = formatDate_(body.newDate || data[i][cols.indexOf("PreferredDate")]);
      const bkTime = body.newTime ? String(body.newTime) : bookingTime_(data[i][cols.indexOf("PreferredTime")], shown[i][cols.indexOf("PreferredTime")]);
      const bkStatus = body.status || data[i][cols.indexOf("Status")];
      try {
        if (email) {
          MailApp.sendEmail({
            to: email,
            subject: "Workshop2U — Booking update",
            body: "Hi " + name + ",\n\nThanks for booking with Workshop2U (" + bkWorkshop + ").\n" +
                  "Requested date/time: " + bkDate + " " + bkTime + "\nService: " + bkService + "\n\n" +
                  "Your booking status is now: " + bkStatus + "\n\nWorkshop2U"
          });
        }
      } catch (e) { /* ignore */ }
      notifyManagers_(bkWorkshop, "Booking " + bkStatus + " — " + bkWorkshop,
        "Booking status is now: " + bkStatus + " (updated by " + user.username + ")\n\n" +
        bookingSummary_({ name: name, phone: data[i][cols.indexOf("Phone")], email: email, vehicleType: data[i][cols.indexOf("VehicleType")], plate: data[i][cols.indexOf("VehiclePlateNumber")], workshop: bkWorkshop, date: bkDate, time: bkTime, serviceType: bkService, notes: data[i][cols.indexOf("Notes")] }) +
        "\n\nWorkshop2U");

      logAudit_(user.username, user.role, "BOOKING_UPDATED", body.bookingId + " -> " + body.status);
      return { success: true };
    }
  }
  return { success: false, message: "Booking not found." };
}

// =============================================================================
// FEATURE 7 — ANALYTICS DASHBOARD
// =============================================================================
function getAnalytics_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const workshops = ["Melaka", "Negeri Sembilan", "Johor"];
  const roleScope = user.role === "admin" ? [user.workshop] : workshops;

  // Narrows the user's role-based scope by a dropdown's chosen workshop.
  // An admin can never see past their own workshop regardless of what's
  // requested; "All" (or nothing requested) keeps the full role scope.
  function narrowScope(requested) {
    if (user.role === "admin") return [user.workshop];
    if (!requested || requested === "All") return workshops;
    return workshops.indexOf(requested) !== -1 ? [requested] : workshops;
  }

  const data = sheet_("History").getDataRange().getValues();

  // ---- existing: monthly revenue/jobs by workshop, last 12 months ----
  const monthsSet = {};
  const revenue = {}; roleScope.forEach(w => revenue[w] = {});
  const jobs = {}; roleScope.forEach(w => jobs[w] = {});
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const workshop = r[7];
    if (roleScope.indexOf(workshop) === -1) continue;
    const d = new Date(r[1]);
    if (isNaN(d.getTime())) continue;
    const monthKey = Utilities.formatDate(d, tz_(), "yyyy-MM");
    monthsSet[monthKey] = true;
    revenue[workshop][monthKey] = (revenue[workshop][monthKey] || 0) + Number(r[9] || 0);
    jobs[workshop][monthKey] = (jobs[workshop][monthKey] || 0) + 1;
  }
  const months = Object.keys(monthsSet).sort().slice(-12);
  const series = {}, jobSeries = {};
  roleScope.forEach(w => {
    series[w] = months.map(m => revenue[w][m] || 0);
    jobSeries[w] = months.map(m => jobs[w][m] || 0);
  });
  const totalRevenue = Object.values(series).flat().reduce((a, b) => a + b, 0);
  const totalJobs = Object.values(jobSeries).flat().reduce((a, b) => a + b, 0);

  // ---- service type breakdown (pie chart 1) ----
  const svcScope = narrowScope(body.serviceTypeWorkshop);
  const svcCounts = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (svcScope.indexOf(r[7]) === -1) continue;
    const type = r[8] || "Unspecified";
    svcCounts[type] = (svcCounts[type] || 0) + 1;
  }
  const serviceTypeBreakdown = { labels: Object.keys(svcCounts), counts: Object.values(svcCounts), scope: svcScope };

  // ---- member breakdown (pie chart 2) — "All" shows the split *by
  // workshop* (matches "percentage of members ... for all sites" literally);
  // narrowing to one specific workshop switches to Active vs Inactive for
  // that workshop instead, since a by-workshop pie scoped to a single
  // workshop would just be one pointless 100%-full slice. ----
  const memScope = narrowScope(body.memberWorkshop);
  const credData = sheet_("Credentials").getDataRange().getValues();
  const isSingleWorkshop = memScope.length === 1;
  const byWorkshopCounts = {}; memScope.forEach(w => byWorkshopCounts[w] = 0);
  let activeCount = 0, inactiveCount = 0;
  for (let i = 1; i < credData.length; i++) {
    const cred = rowToCredential_(credData[i]);
    if (cred.Role !== "member") continue;
    if (memScope.indexOf(cred.WorkshopLocation) === -1) continue;
    byWorkshopCounts[cred.WorkshopLocation]++;
    if (String(cred.Status).toLowerCase() === "active") activeCount++; else inactiveCount++;
  }
  const memberBreakdown = isSingleWorkshop
    ? { mode: "status", labels: ["Active", "Inactive"], counts: [activeCount, inactiveCount], scope: memScope }
    : { mode: "workshop", labels: Object.keys(byWorkshopCounts), counts: Object.values(byWorkshopCounts), scope: memScope };

  // ---- daily revenue for one selected month (bar + per-service-type line
  // overlay) — always within the user's full role scope; no separate
  // workshop dropdown for this one, per the request (only a month
  // selector). byServiceType only ever has these 5 fixed keys (matching the
  // Add Service Record dropdown) so the front end can always draw exactly
  // 5 lines, even for a type with no records that month (all zeros). ----
  const SERVICE_TYPES = ["Normal Service", "Repair", "Troubleshooting", "Car Wash", "A/C Service"];
  const monthKey = /^\d{4}-\d{2}$/.test(body.month || "") ? body.month : Utilities.formatDate(new Date(), tz_(), "yyyy-MM");
  const [yy, mm] = monthKey.split("-").map(Number);
  const daysInMonth = new Date(yy, mm, 0).getDate();
  const dailyRevenue = new Array(daysInMonth).fill(0);
  const byServiceType = {}; SERVICE_TYPES.forEach(t => byServiceType[t] = new Array(daysInMonth).fill(0));
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (roleScope.indexOf(r[7]) === -1) continue;
    const d = new Date(r[1]);
    if (isNaN(d.getTime())) continue;
    if (Utilities.formatDate(d, tz_(), "yyyy-MM") !== monthKey) continue;
    const day = d.getDate();
    const price = Number(r[9] || 0);
    dailyRevenue[day - 1] += price;
    if (byServiceType[r[8]]) byServiceType[r[8]][day - 1] += price;
  }

  return {
    success: true, months, revenueSeries: series, jobSeries, totalRevenue, totalJobs, workshops: roleScope,
    serviceTypeBreakdown, memberBreakdown,
    dailyRevenue: { month: monthKey, days: dailyRevenue, byServiceType }
  };
}

// =============================================================================
// FEATURE 8 — REVIEWS
// The ReviewToken lives on the History row itself (set when the service was
// recorded), so we validate against History rather than pre-creating a row
// in Reviews. The Reviews tab now only ever gets a new row at the moment a
// customer actually submits feedback — nothing is written there before that.
// =============================================================================
function findHistoryRowByReviewToken_(token) {
  const data = sheet_("History").getDataRange().getValues();
  const cols = historyColumns_();
  const tokenIdx = cols.indexOf("ReviewToken");
  for (let i = 1; i < data.length; i++) {
    if (data[i][tokenIdx] === token) {
      return { username: data[i][2], name: data[i][3], workshop: data[i][7] };
    }
  }
  return null;
}

/* Finds the Reviews row for a token, if any. Columns: 0 ReviewToken,
   6 Status, 7 CreatedAt. */
function findReviewRow_(token) {
  const data = sheet_("Reviews").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) return { rowIndex: i + 1, status: data[i][6], createdAt: data[i][7] };
  }
  return null;
}
// Only a row whose Status is "Submitted" counts. An older version of this
// code left "Pending" placeholder rows behind when a service record was
// saved; those mean "link sent, nothing submitted yet", not "already done".
function reviewAlreadySubmitted_(token) {
  const found = findReviewRow_(token);
  return !!found && found.status === "Submitted";
}

function getReviewContext_(body) {
  const token = body.reviewToken;
  const match = findHistoryRowByReviewToken_(token);
  if (!match) return { success: false, message: "This review link is invalid or has expired." };
  if (reviewAlreadySubmitted_(token)) return { success: false, message: "You've already submitted feedback for this visit. Thank you!" };
  return { success: true, workshop: match.workshop };
}

function submitReview_(body) {
  const token = body.reviewToken;
  const rating = Number(body.rating);
  if (!rating || rating < 1 || rating > 5) return { success: false, message: "Please select a rating." };

  const match = findHistoryRowByReviewToken_(token);
  if (!match) return { success: false, message: "This review link is invalid or has expired." };
  if (reviewAlreadySubmitted_(token)) return { success: false, message: "You've already submitted feedback for this visit." };

  const now = new Date();
  const existing = findReviewRow_(token);   // only ever a leftover "Pending" row here
  const rowValues = [token, match.username, match.name, match.workshop, rating, body.comment || "", "Submitted", (existing && existing.createdAt) || now, now];
  if (existing) sheet_("Reviews").getRange(existing.rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  else sheet_("Reviews").appendRow(rowValues);
  logAudit_(match.username, "member", "REVIEW_SUBMITTED", "Rating " + rating + " @ " + match.workshop);
  return { success: true };
}

/** Feature 8 (display) — list submitted reviews for the staff dashboard. */
function getReviews_(body) {
  const auth = requireAuth_(body.token, ["admin", "manager", "webmaster"]);
  if (!auth.ok) return { success: false, message: auth.error };
  const user = auth.user;

  const data = sheet_("Reviews").getDataRange().getValues();
  const reviews = [];
  let ratingSum = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] !== "Submitted") continue;   // skip leftover "Pending" placeholders
    const r = { token: data[i][0], username: data[i][1], name: data[i][2], workshop: data[i][3], rating: data[i][4], comment: data[i][5], submittedAt: formatDate_(data[i][8]) };
    if (user.role === "admin" && r.workshop !== user.workshop) continue;
    if ((user.role === "manager" || user.role === "webmaster") && body.workshop && body.workshop !== "All" && r.workshop !== body.workshop) continue;
    reviews.push(r);
    ratingSum += Number(r.rating) || 0;
  }
  reviews.sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
  const avgRating = reviews.length ? (ratingSum / reviews.length) : 0;
  return { success: true, reviews, avgRating, count: reviews.length };
}

// =============================================================================
// ONE-TIME SETUP
// =============================================================================
function initializeSheet() {
  const ss = ss_();
  const sheets = {
    "Credentials": credentialColumns_(),
    "History": historyColumns_(),
    "Bookings": bookingColumns_(),
    "Sessions": ["Token","Username","Role","WorkshopLocation","FullName","CreatedAt","ExpiresAt"],
    "Vehicles": ["Username","PlateNumber","VehicleType","Nickname"],
    "LoginAttempts": ["Username","FailCount","LockUntil"],
    "OtpCodes": ["Username","Code","CreatedAt","ExpiresAt"],
    "PasswordResets": ["Token","Username","CreatedAt","ExpiresAt"],
    "Reviews": ["ReviewToken","Username","Name","Workshop","Rating","Comment","Status","CreatedAt","SubmittedAt"],
    "AuditLog": ["Timestamp","Username","Role","Action","Details"],
    "Reminders": ["PlateNumber","Username","LastReminderSent"],
    "InvoiceQueue": ["HistoryID","Date","Name","Address","VehicleType","PlateNumber","Workshop","ServiceType","Price","Notes","CustomerEmail","ReviewToken","SiteUrl","QueuedAt"]
  };
  Object.keys(sheets).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(sheets[name]); sh.setFrozenRows(1); }
  });

  const credSheet = ss.getSheetByName("Credentials");
  if (credSheet.getLastRow() === 1) {
    credSheet.appendRow(["webmaster", hashPassword_("ChangeMe123!"), "webmaster", "All", "Site Webmaster", "", "", "", "Active", new Date()]);
    Logger.log("Created default webmaster login -> username: webmaster / password: ChangeMe123! (change this immediately, and add a real email so OTP works)");
  }
}

/** Run daily (Triggers > Add Trigger) to purge expired sessions/codes. */
function cleanExpiredSessions() {
  purgeExpired_("Sessions", 6);
  purgeExpired_("OtpCodes", 3);
  purgeExpired_("PasswordResets", 3);
}
function purgeExpired_(sheetName, expiryColIndex) {
  const sh = sheet_(sheetName);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (new Date(data[i][expiryColIndex]).getTime() < Date.now()) sh.deleteRow(i + 1);
  }
}
