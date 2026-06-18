const STORAGE_KEY = "availability-composer-settings-v2";

const GRAPH_SCOPES = ["User.Read", "Calendars.Read"];
const GRAPH_URL = "https://graph.microsoft.com/v1.0";
const GOOGLE_SCOPE = "https://www.googleapis.com/auth/calendar.readonly";
const GOOGLE_EVENTS_URL = "https://www.googleapis.com/calendar/v3/calendars/primary/events";

const el = {
  googleClientId: document.getElementById("googleClientId"),
  googleSignInBtn: document.getElementById("googleSignInBtn"),
  googleSignOutBtn: document.getElementById("googleSignOutBtn"),
  clientId: document.getElementById("clientId"),
  tenantId: document.getElementById("tenantId"),
  signInBtn: document.getElementById("signInBtn"),
  signOutBtn: document.getElementById("signOutBtn"),
  icsText: document.getElementById("icsText"),
  icsFile: document.getElementById("icsFile"),
  loadIcsTextBtn: document.getElementById("loadIcsTextBtn"),
  loadIcsFileBtn: document.getElementById("loadIcsFileBtn"),
  clearIcsBtn: document.getElementById("clearIcsBtn"),
  authBadge: document.getElementById("authBadge"),
  redirectUriText: document.getElementById("redirectUriText"),
  durationMin: document.getElementById("durationMin"),
  spanDays: document.getElementById("spanDays"),
  maxOptions: document.getElementById("maxOptions"),
  dayStart: document.getElementById("dayStart"),
  dayEnd: document.getElementById("dayEnd"),
  slotStep: document.getElementById("slotStep"),
  leadHours: document.getElementById("leadHours"),
  weekdaysOnly: document.getElementById("weekdaysOnly"),
  generateBtn: document.getElementById("generateBtn"),
  statusText: document.getElementById("statusText"),
  resultText: document.getElementById("resultText"),
  resultList: document.getElementById("resultList"),
  copyBtn: document.getElementById("copyBtn"),
};

const nowUrl = new URL(window.location.href);
nowUrl.hash = "";
el.redirectUriText.textContent = nowUrl.toString();

let msalClient = null;
let activeAccount = null;
let googleTokenClient = null;
let googleTokenClientId = "";
let googleAccessToken = "";
let sourceMode = "google";
let icsBusyBlocks = null;

init();

function init() {
  applySavedSettings();
  wireInputs();
  updateAuthUi();
}

function wireInputs() {
  const inputsToPersist = [
    el.googleClientId,
    el.clientId,
    el.tenantId,
    el.durationMin,
    el.spanDays,
    el.maxOptions,
    el.dayStart,
    el.dayEnd,
    el.slotStep,
    el.leadHours,
    el.weekdaysOnly,
  ];

  for (const input of inputsToPersist) {
    input.addEventListener("change", () => {
      saveSettings();
      if (input === el.clientId || input === el.tenantId) {
        msalClient = null;
        activeAccount = null;
      }
      if (input === el.googleClientId) {
        googleTokenClient = null;
        googleTokenClientId = "";
        googleAccessToken = "";
      }
      updateAuthUi();
    });
  }

  el.googleSignInBtn.addEventListener("click", onGoogleSignIn);
  el.googleSignOutBtn.addEventListener("click", onGoogleSignOut);

  el.signInBtn.addEventListener("click", onSignIn);
  el.signOutBtn.addEventListener("click", onSignOut);

  el.loadIcsTextBtn.addEventListener("click", onLoadIcsText);
  el.loadIcsFileBtn.addEventListener("click", onLoadIcsFile);
  el.clearIcsBtn.addEventListener("click", onClearIcs);

  el.generateBtn.addEventListener("click", onGenerate);
  el.copyBtn.addEventListener("click", onCopy);
}

function applySavedSettings() {
  const defaults = {
    googleClientId: "",
    clientId: "",
    tenantId: "common",
    durationMin: "60",
    spanDays: "7",
    maxOptions: "5",
    dayStart: "09:00",
    dayEnd: "17:00",
    slotStep: "15",
    leadHours: "2",
    weekdaysOnly: true,
  };

  let saved = {};
  try {
    saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
  } catch {
    saved = {};
  }

  el.googleClientId.value = saved.googleClientId ?? defaults.googleClientId;
  el.clientId.value = saved.clientId ?? defaults.clientId;
  el.tenantId.value = saved.tenantId ?? defaults.tenantId;
  el.durationMin.value = String(saved.durationMin ?? defaults.durationMin);
  el.spanDays.value = String(saved.spanDays ?? defaults.spanDays);
  el.maxOptions.value = String(saved.maxOptions ?? defaults.maxOptions);
  el.dayStart.value = saved.dayStart ?? defaults.dayStart;
  el.dayEnd.value = saved.dayEnd ?? defaults.dayEnd;
  el.slotStep.value = String(saved.slotStep ?? defaults.slotStep);
  el.leadHours.value = String(saved.leadHours ?? defaults.leadHours);
  el.weekdaysOnly.checked = saved.weekdaysOnly ?? defaults.weekdaysOnly;
}

function saveSettings() {
  const payload = {
    googleClientId: el.googleClientId.value.trim(),
    clientId: el.clientId.value.trim(),
    tenantId: el.tenantId.value,
    durationMin: Number(el.durationMin.value),
    spanDays: Number(el.spanDays.value),
    maxOptions: Number(el.maxOptions.value),
    dayStart: el.dayStart.value,
    dayEnd: el.dayEnd.value,
    slotStep: Number(el.slotStep.value),
    leadHours: Number(el.leadHours.value),
    weekdaysOnly: el.weekdaysOnly.checked,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function updateStatus(text, mode = "") {
  el.statusText.textContent = text;
  el.statusText.classList.remove("ok", "warn");
  if (mode) {
    el.statusText.classList.add(mode);
  }
}

function updateAuthUi() {
  const hasGoogleClientId = Boolean(el.googleClientId.value.trim());
  const hasMsClientId = Boolean(el.clientId.value.trim());
  const signedInGoogle = Boolean(googleAccessToken);
  const signedInOutlook = Boolean(activeAccount);
  const hasIcs = Array.isArray(icsBusyBlocks);

  el.googleSignInBtn.disabled = !hasGoogleClientId;
  el.googleSignOutBtn.disabled = !signedInGoogle;
  el.signInBtn.disabled = !hasMsClientId;
  el.signOutBtn.disabled = !signedInOutlook;
  el.generateBtn.disabled = !(signedInGoogle || signedInOutlook || hasIcs);

  if (sourceMode === "google" && signedInGoogle) {
    el.authBadge.textContent = "Google connected";
    el.authBadge.classList.add("ok");
    updateStatus("Ready with live Google Calendar. Click Find my times.", "ok");
    return;
  }

  if (sourceMode === "graph" && signedInOutlook) {
    el.authBadge.textContent = activeAccount.username;
    el.authBadge.classList.add("ok");
    updateStatus("Ready with live Outlook Calendar. Click Find my times.", "ok");
    return;
  }

  if (sourceMode === "ics" && hasIcs) {
    el.authBadge.textContent = "Using .ics import";
    el.authBadge.classList.add("ok");
    updateStatus("Ready in .ics mode. Click Find my times.", "ok");
    return;
  }

  el.authBadge.classList.remove("ok");
  if (!hasGoogleClientId && !hasMsClientId && !hasIcs) {
    el.authBadge.textContent = "No source configured";
    updateStatus("Add a Google client ID, Outlook client ID, or load .ics data.", "warn");
    return;
  }

  el.authBadge.textContent = "Not connected";
  updateStatus("Connect Google, Outlook, or .ics data, then generate availability.");
}

function createMsalClient() {
  if (msalClient) {
    return msalClient;
  }

  const clientId = el.clientId.value.trim();
  if (!clientId) {
    throw new Error("Missing Microsoft client ID");
  }

  const tenantId = el.tenantId.value;
  const redirectUri = nowUrl.toString();

  msalClient = new msal.PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
  });

  const knownAccounts = msalClient.getAllAccounts();
  if (knownAccounts.length > 0) {
    activeAccount = knownAccounts[0];
  }

  return msalClient;
}

async function onGoogleSignIn() {
  try {
    saveSettings();
    await requestGoogleAccessToken("consent");
    sourceMode = "google";
    updateAuthUi();
  } catch (error) {
    updateStatus(formatError(error), "warn");
  }
}

function onGoogleSignOut() {
  if (googleAccessToken && window.google?.accounts?.oauth2?.revoke) {
    window.google.accounts.oauth2.revoke(googleAccessToken, () => {});
  }
  googleAccessToken = "";
  googleTokenClient = null;
  googleTokenClientId = "";

  if (activeAccount) {
    sourceMode = "graph";
  } else if (Array.isArray(icsBusyBlocks)) {
    sourceMode = "ics";
  }

  updateAuthUi();
}

async function requestGoogleAccessToken(prompt = "") {
  const clientId = el.googleClientId.value.trim();
  if (!clientId) {
    throw new Error("Missing Google OAuth client ID");
  }
  if (!window.google?.accounts?.oauth2) {
    throw new Error("Google Identity script did not load. Check your network or content blocker.");
  }

  if (!googleTokenClient || googleTokenClientId !== clientId) {
    googleTokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: GOOGLE_SCOPE,
      callback: () => {},
      error_callback: () => {},
    });
    googleTokenClientId = clientId;
  }

  const tokenResponse = await new Promise((resolve, reject) => {
    googleTokenClient.callback = (resp) => {
      if (resp?.error) {
        reject(new Error(resp.error_description || resp.error));
        return;
      }
      resolve(resp);
    };

    googleTokenClient.error_callback = (resp) => {
      reject(new Error(resp?.message || "Google token request failed"));
    };

    googleTokenClient.requestAccessToken({ prompt });
  });

  if (!tokenResponse?.access_token) {
    throw new Error("Google did not return an access token.");
  }

  googleAccessToken = tokenResponse.access_token;
}

async function onSignIn() {
  try {
    saveSettings();
    const client = createMsalClient();
    const loginResult = await client.loginPopup({
      scopes: GRAPH_SCOPES,
      prompt: "select_account",
    });
    activeAccount = loginResult.account;
    sourceMode = "graph";
    updateAuthUi();
  } catch (error) {
    updateStatus(formatError(error), "warn");
  }
}

async function onSignOut() {
  try {
    if (!msalClient || !activeAccount) {
      return;
    }
    await msalClient.logoutPopup({ account: activeAccount });
    activeAccount = null;

    if (googleAccessToken) {
      sourceMode = "google";
    } else if (Array.isArray(icsBusyBlocks)) {
      sourceMode = "ics";
    }

    updateAuthUi();
  } catch (error) {
    updateStatus(formatError(error), "warn");
  }
}

async function onGenerate() {
  el.copyBtn.disabled = true;
  el.resultList.innerHTML = "";
  el.resultText.value = "";

  try {
    const settings = readSettings();
    validateSettings(settings);
    saveSettings();

    updateStatus("Checking your calendar and building suggestions...");

    const { start, end } = buildRange(settings);
    const busyBlocks = await loadBusyBlocksBySource(start, end);
    const candidates = buildCandidateSlots(start, end, busyBlocks, settings);

    if (candidates.length === 0) {
      updateStatus("No free slots in that range. Widen the date range or time window.", "warn");
      return;
    }

    const picks = chooseSpreadOptions(candidates, settings.maxOptions, start, end);
    const output = buildEmailOutput(picks, settings, start, end);

    el.resultText.value = output.text;
    renderListPreview(output.lines);
    el.copyBtn.disabled = false;
    updateStatus(`Generated ${picks.length} options. Copy and paste into your email.`, "ok");
  } catch (error) {
    updateStatus(formatError(error), "warn");
  }
}

async function loadBusyBlocksBySource(start, end) {
  if (sourceMode === "google" && googleAccessToken) {
    try {
      return await fetchGoogleBusyBlocks(googleAccessToken, start, end);
    } catch (error) {
      if (String(error.message || "").includes("401")) {
        await requestGoogleAccessToken("");
        return fetchGoogleBusyBlocks(googleAccessToken, start, end);
      }
      throw error;
    }
  }

  if (sourceMode === "graph" && activeAccount) {
    const token = await acquireToken();
    return fetchBusyBlocks(token, start, end);
  }

  if (sourceMode === "ics" && Array.isArray(icsBusyBlocks)) {
    return filterBusyBlocksByRange(icsBusyBlocks, start, end);
  }

  if (googleAccessToken) {
    sourceMode = "google";
    return fetchGoogleBusyBlocks(googleAccessToken, start, end);
  }

  if (activeAccount) {
    sourceMode = "graph";
    const token = await acquireToken();
    return fetchBusyBlocks(token, start, end);
  }

  if (Array.isArray(icsBusyBlocks)) {
    sourceMode = "ics";
    return filterBusyBlocksByRange(icsBusyBlocks, start, end);
  }

  throw new Error("Connect Google, Outlook, or .ics data first.");
}

async function onLoadIcsFile() {
  try {
    const [file] = el.icsFile.files || [];
    if (!file) {
      throw new Error("Select an .ics file first.");
    }
    const text = await file.text();
    const blocks = parseIcsBusyBlocks(text);
    if (blocks.length === 0) {
      throw new Error("No busy events found in this .ics file.");
    }
    icsBusyBlocks = blocks;
    sourceMode = "ics";
    updateAuthUi();
    updateStatus(`Loaded ${blocks.length} busy blocks from ${file.name}.`, "ok");
  } catch (error) {
    updateStatus(formatError(error), "warn");
  }
}

function onLoadIcsText() {
  try {
    const text = (el.icsText.value || "").trim();
    if (!text) {
      throw new Error("Paste .ics text first.");
    }
    const blocks = parseIcsBusyBlocks(text);
    if (blocks.length === 0) {
      throw new Error("No busy events found in pasted .ics text.");
    }
    icsBusyBlocks = blocks;
    sourceMode = "ics";
    updateAuthUi();
    updateStatus(`Loaded ${blocks.length} busy blocks from pasted .ics data.`, "ok");
  } catch (error) {
    updateStatus(formatError(error), "warn");
  }
}

function onClearIcs() {
  icsBusyBlocks = null;
  if (googleAccessToken) {
    sourceMode = "google";
  } else if (activeAccount) {
    sourceMode = "graph";
  }
  el.icsText.value = "";
  el.icsFile.value = "";
  updateAuthUi();
}

function readSettings() {
  return {
    durationMin: Number(el.durationMin.value),
    spanDays: Number(el.spanDays.value),
    maxOptions: Number(el.maxOptions.value),
    dayStart: el.dayStart.value,
    dayEnd: el.dayEnd.value,
    slotStep: Number(el.slotStep.value),
    leadHours: Number(el.leadHours.value),
    weekdaysOnly: el.weekdaysOnly.checked,
  };
}

function validateSettings(settings) {
  if (!Number.isFinite(settings.durationMin) || settings.durationMin < 5) {
    throw new Error("Meeting length must be at least 5 minutes.");
  }
  if (!Number.isFinite(settings.spanDays) || settings.spanDays < 1) {
    throw new Error("Look ahead must be at least 1 day.");
  }
  if (!Number.isFinite(settings.maxOptions) || settings.maxOptions < 1) {
    throw new Error("Number of options must be at least 1.");
  }
  if (!Number.isFinite(settings.slotStep) || settings.slotStep < 5) {
    throw new Error("Slot interval must be at least 5 minutes.");
  }

  const startMinutes = timeToMinutes(settings.dayStart);
  const endMinutes = timeToMinutes(settings.dayEnd);

  if (endMinutes <= startMinutes) {
    throw new Error("Day end must be later than day start.");
  }
  if (settings.durationMin > endMinutes - startMinutes) {
    throw new Error("Meeting length cannot exceed your daily time window.");
  }
}

async function acquireToken() {
  const client = createMsalClient();

  if (!activeAccount) {
    const knownAccounts = client.getAllAccounts();
    if (knownAccounts.length > 0) {
      activeAccount = knownAccounts[0];
    }
  }

  if (!activeAccount) {
    throw new Error("Sign in to Outlook first.");
  }

  try {
    const tokenResult = await client.acquireTokenSilent({
      account: activeAccount,
      scopes: GRAPH_SCOPES,
    });
    return tokenResult.accessToken;
  } catch {
    const tokenResult = await client.acquireTokenPopup({
      account: activeAccount,
      scopes: GRAPH_SCOPES,
      prompt: "consent",
    });
    return tokenResult.accessToken;
  }
}

function buildRange(settings) {
  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);
  end.setDate(end.getDate() + settings.spanDays);
  return { start, end };
}

async function fetchGoogleBusyBlocks(token, start, end) {
  let pageToken = "";
  const blocks = [];

  while (true) {
    const params = new URLSearchParams({
      timeMin: start.toISOString(),
      timeMax: end.toISOString(),
      singleEvents: "true",
      orderBy: "startTime",
      maxResults: "2500",
    });

    if (pageToken) {
      params.set("pageToken", pageToken);
    }

    const url = `${GOOGLE_EVENTS_URL}?${params.toString()}`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      const bodyText = await response.text();
      throw new Error(`Google Calendar request failed (${response.status}): ${bodyText}`);
    }

    const payload = await response.json();

    for (const event of payload.items || []) {
      if (event.status === "cancelled") {
        continue;
      }
      if (event.transparency === "transparent") {
        continue;
      }

      const startDate = parseGoogleDate(event.start);
      const endDate = parseGoogleDate(event.end);

      if (!(startDate instanceof Date) || Number.isNaN(startDate.getTime())) {
        continue;
      }
      if (!(endDate instanceof Date) || Number.isNaN(endDate.getTime())) {
        continue;
      }
      if (endDate <= startDate) {
        continue;
      }

      blocks.push({
        start: startDate,
        end: endDate,
      });
    }

    pageToken = payload.nextPageToken || "";
    if (!pageToken) {
      break;
    }
  }

  blocks.sort((a, b) => a.start - b.start);
  return mergeOverlaps(blocks);
}

function parseGoogleDate(node) {
  if (!node) {
    return null;
  }
  if (node.dateTime) {
    return new Date(node.dateTime);
  }
  if (node.date) {
    return new Date(`${node.date}T00:00:00`);
  }
  return null;
}

async function fetchBusyBlocks(token, start, end) {
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const busyShowAs = new Set(["busy", "tentative", "oof", "workingElsewhere", "unknown"]);

  const params = new URLSearchParams({
    startDateTime: start.toISOString(),
    endDateTime: end.toISOString(),
    $top: "100",
    $orderby: "start/dateTime",
    $select: "start,end,showAs,isAllDay",
  });

  let url = `${GRAPH_URL}/me/calendarView?${params.toString()}`;
  const blocks = [];

  while (url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: `outlook.timezone="${timeZone}"`,
      },
    });

    if (!response.ok) {
      const bodyText = await response.text();
      throw new Error(`Graph request failed (${response.status}): ${bodyText}`);
    }

    const payload = await response.json();

    for (const event of payload.value || []) {
      if (!busyShowAs.has(event.showAs)) {
        continue;
      }

      const startDate = parseGraphDate(event.start);
      const endDate = parseGraphDate(event.end);

      if (!(startDate instanceof Date) || Number.isNaN(startDate.getTime())) {
        continue;
      }
      if (!(endDate instanceof Date) || Number.isNaN(endDate.getTime())) {
        continue;
      }
      if (endDate <= startDate) {
        continue;
      }

      blocks.push({
        start: startDate,
        end: endDate,
      });
    }

    url = payload["@odata.nextLink"] || null;
  }

  blocks.sort((a, b) => a.start - b.start);
  return mergeOverlaps(blocks);
}

function parseGraphDate(node) {
  if (!node) {
    return null;
  }

  if (node.dateTime) {
    return new Date(node.dateTime);
  }

  if (node.date) {
    return new Date(`${node.date}T00:00:00`);
  }

  return null;
}

function filterBusyBlocksByRange(blocks, start, end) {
  const out = [];
  for (const block of blocks) {
    if (block.end <= start || block.start >= end) {
      continue;
    }
    out.push({
      start: block.start < start ? new Date(start) : new Date(block.start),
      end: block.end > end ? new Date(end) : new Date(block.end),
    });
  }
  out.sort((a, b) => a.start - b.start);
  return out;
}

function parseIcsBusyBlocks(icsText) {
  const unfolded = icsText.replace(/\r?\n[ \t]/g, "");
  const lines = unfolded.split(/\r?\n/);
  const blocks = [];

  let inEvent = false;
  let event = {};

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) {
      continue;
    }

    if (line === "BEGIN:VEVENT") {
      inEvent = true;
      event = {};
      continue;
    }

    if (line === "END:VEVENT") {
      if (inEvent) {
        const block = eventToBusyBlock(event);
        if (block) {
          blocks.push(block);
        }
      }
      inEvent = false;
      event = {};
      continue;
    }

    if (!inEvent) {
      continue;
    }

    const sep = line.indexOf(":");
    if (sep === -1) {
      continue;
    }

    const keyPart = line.slice(0, sep);
    const value = line.slice(sep + 1);
    const [nameRaw, ...paramParts] = keyPart.split(";");
    const name = nameRaw.toUpperCase();
    const params = {};

    for (const part of paramParts) {
      const [k, v] = part.split("=");
      if (k && v) {
        params[k.toUpperCase()] = v;
      }
    }

    if (name === "DTSTART") {
      event.dtStart = { value, params };
    } else if (name === "DTEND") {
      event.dtEnd = { value, params };
    } else if (name === "TRANSP") {
      event.transparency = value.toUpperCase();
    } else if (name === "STATUS") {
      event.status = value.toUpperCase();
    }
  }

  blocks.sort((a, b) => a.start - b.start);
  return mergeOverlaps(blocks);
}

function eventToBusyBlock(event) {
  if (!event.dtStart) {
    return null;
  }
  if (event.status === "CANCELLED") {
    return null;
  }
  if (event.transparency === "TRANSPARENT") {
    return null;
  }

  const start = parseIcsDateValue(event.dtStart.value, event.dtStart.params);
  let end = event.dtEnd ? parseIcsDateValue(event.dtEnd.value, event.dtEnd.params) : null;

  if (!start || Number.isNaN(start.getTime())) {
    return null;
  }

  if (!end || Number.isNaN(end.getTime())) {
    end = new Date(start.getTime() + 30 * 60 * 1000);
  }

  if (end <= start) {
    return null;
  }

  return { start, end };
}

function parseIcsDateValue(value, params = {}) {
  const normalized = String(value || "").trim();
  const isDateOnly = params.VALUE === "DATE" || /^\d{8}$/.test(normalized);

  if (isDateOnly) {
    const y = Number(normalized.slice(0, 4));
    const m = Number(normalized.slice(4, 6));
    const d = Number(normalized.slice(6, 8));
    return new Date(y, m - 1, d, 0, 0, 0, 0);
  }

  const utcMatch = normalized.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$/);
  if (utcMatch) {
    const [, y, m, d, hh, mm, ss] = utcMatch;
    return new Date(Date.UTC(Number(y), Number(m) - 1, Number(d), Number(hh), Number(mm), Number(ss)));
  }

  const localMatch = normalized.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$/);
  if (localMatch) {
    const [, y, m, d, hh, mm, ss] = localMatch;
    return new Date(Number(y), Number(m) - 1, Number(d), Number(hh), Number(mm), Number(ss));
  }

  const shortMatch = normalized.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})$/);
  if (shortMatch) {
    const [, y, m, d, hh, mm] = shortMatch;
    return new Date(Number(y), Number(m) - 1, Number(d), Number(hh), Number(mm), 0);
  }

  return new Date(normalized);
}

function mergeOverlaps(items) {
  if (items.length < 2) {
    return items;
  }

  const merged = [items[0]];

  for (let i = 1; i < items.length; i += 1) {
    const prev = merged[merged.length - 1];
    const cur = items[i];

    if (cur.start <= prev.end) {
      if (cur.end > prev.end) {
        prev.end = cur.end;
      }
    } else {
      merged.push(cur);
    }
  }

  return merged;
}

function buildCandidateSlots(rangeStart, rangeEnd, busyBlocks, settings) {
  const candidates = [];
  const dayStartMin = timeToMinutes(settings.dayStart);
  const dayEndMin = timeToMinutes(settings.dayEnd);
  const durationMs = settings.durationMin * 60 * 1000;
  const stepMs = settings.slotStep * 60 * 1000;
  const earliest = new Date(Date.now() + settings.leadHours * 60 * 60 * 1000);

  let cursor = new Date(rangeStart);
  cursor.setHours(0, 0, 0, 0);

  while (cursor < rangeEnd) {
    const day = cursor.getDay();
    if (!settings.weekdaysOnly || (day >= 1 && day <= 5)) {
      const windowStart = new Date(cursor);
      windowStart.setMinutes(dayStartMin, 0, 0);

      const windowEnd = new Date(cursor);
      windowEnd.setMinutes(dayEndMin, 0, 0);

      let slot = new Date(Math.max(windowStart.getTime(), earliest.getTime()));
      slot = ceilToStep(slot, stepMs);

      while (slot < windowEnd) {
        const slotEnd = new Date(slot.getTime() + durationMs);

        if (slotEnd > windowEnd || slotEnd > rangeEnd) {
          break;
        }

        if (!hasOverlap(slot, slotEnd, busyBlocks)) {
          candidates.push(new Date(slot));
        }

        slot = new Date(slot.getTime() + stepMs);
      }
    }

    cursor.setDate(cursor.getDate() + 1);
  }

  return candidates;
}

function ceilToStep(date, stepMs) {
  const time = date.getTime();
  const rounded = Math.ceil(time / stepMs) * stepMs;
  return new Date(rounded);
}

function hasOverlap(slotStart, slotEnd, busyBlocks) {
  for (const busy of busyBlocks) {
    if (busy.end <= slotStart) {
      continue;
    }
    if (busy.start >= slotEnd) {
      break;
    }
    if (busy.start < slotEnd && busy.end > slotStart) {
      return true;
    }
  }
  return false;
}

function chooseSpreadOptions(candidates, desiredCount, rangeStart, rangeEnd) {
  if (candidates.length <= desiredCount) {
    return [...candidates];
  }

  const selected = [];
  const seenDay = new Set();
  const seenBucket = new Set();

  const spanMs = rangeEnd.getTime() - rangeStart.getTime();
  const safeSpan = spanMs > 0 ? spanMs : 1;

  while (selected.length < desiredCount) {
    let bestCandidate = null;
    let bestScore = -Infinity;

    for (const slot of candidates) {
      if (selected.some((picked) => picked.getTime() === slot.getTime())) {
        continue;
      }

      const soonness = 1 - (slot.getTime() - rangeStart.getTime()) / safeSpan;

      let distanceScore = 0.7;
      let nearestMs = Infinity;
      for (const picked of selected) {
        const dist = Math.abs(slot.getTime() - picked.getTime());
        if (dist < nearestMs) {
          nearestMs = dist;
        }
      }

      if (selected.length > 0) {
        distanceScore = Math.min(1, nearestMs / safeSpan);
      }

      const dayKey = toDayKey(slot);
      const bucketKey = toTimeBucket(slot);
      const dayBonus = seenDay.has(dayKey) ? 0 : 0.24;
      const bucketBonus = seenBucket.has(bucketKey) ? 0 : 0.18;
      const closePenalty = nearestMs < 2 * 60 * 60 * 1000 ? 0.24 : 0;

      const score =
        0.35 * soonness +
        0.45 * distanceScore +
        dayBonus +
        bucketBonus -
        closePenalty;

      if (score > bestScore) {
        bestScore = score;
        bestCandidate = slot;
      }
    }

    if (!bestCandidate) {
      break;
    }

    selected.push(bestCandidate);
    seenDay.add(toDayKey(bestCandidate));
    seenBucket.add(toTimeBucket(bestCandidate));
  }

  selected.sort((a, b) => a - b);
  return selected;
}

function toDayKey(date) {
  return `${date.getFullYear()}-${date.getMonth()}-${date.getDate()}`;
}

function toTimeBucket(date) {
  const h = date.getHours();
  if (h < 11) {
    return "morning";
  }
  if (h < 14) {
    return "midday";
  }
  if (h < 17) {
    return "afternoon";
  }
  return "evening";
}

function buildEmailOutput(slots, settings) {
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const durationMs = settings.durationMin * 60 * 1000;

  const dayFmt = new Intl.DateTimeFormat(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
  });

  const timeFmt = new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  });

  const lines = slots.map((slot, i) => {
    const end = new Date(slot.getTime() + durationMs);
    return `${i + 1}) ${dayFmt.format(slot)}: ${timeFmt.format(slot)} - ${timeFmt.format(end)}`;
  });

  const intro =
    `Here are ${slots.length} options for a ${settings.durationMin}-minute meeting ` +
    `in the next ${settings.spanDays} day(s) (${timeZone}):`;

  const text = `${intro}\n\n${lines.join("\n")}`;

  return {
    lines,
    text,
  };
}

function renderListPreview(lines) {
  el.resultList.innerHTML = "";
  for (const line of lines) {
    const li = document.createElement("li");
    li.textContent = line;
    el.resultList.appendChild(li);
  }
}

async function onCopy() {
  if (!el.resultText.value.trim()) {
    return;
  }

  try {
    await navigator.clipboard.writeText(el.resultText.value);
    updateStatus("Copied to clipboard.", "ok");
  } catch {
    updateStatus("Copy failed. Select the text and copy manually.", "warn");
  }
}

function timeToMinutes(hhmm) {
  const [hRaw, mRaw] = hhmm.split(":");
  const h = Number(hRaw);
  const m = Number(mRaw);
  return h * 60 + m;
}

function formatError(error) {
  if (!error) {
    return "Something went wrong.";
  }

  if (typeof error === "string") {
    return error;
  }

  if (error.errorMessage) {
    return error.errorMessage;
  }

  if (error.message) {
    return error.message;
  }

  return "Something went wrong.";
}
