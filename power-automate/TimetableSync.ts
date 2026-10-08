/* Paste this entire file into an Excel Office Script. No external libraries or fetch calls. */
interface Property { name: string; params: Record<string, string>; value: string }
interface Definition { properties: Property[] }
interface CalendarDate { wall: string; utc: number; zone: string }
interface Session {
  key: string; family: string; identity: string; start: string; end: string;
  title: string; location: string; description: string; activity: string; module: string;
  cancelled: boolean;
}
interface OutlookEvent {
  id: string; subject: string; start: string; end: string; body: string;
  location: string; showAs: string; hasAttendees: boolean; isAllDay: boolean; isRecurring: boolean; timeZone: string;
  categories: string[]; importance: string; isReminderOn: boolean; reminderMinutesBeforeStart: number;
}
interface SavedEvent {
  id: string; family: string; identity: string; start: string; end: string;
  lastShowAs: string; override: string; missingSince: string; fingerprint: string;
}
interface State { version: number; sourceId: string; lastFeedRunId: string; lastSuccess: string; events: Record<string, SavedEvent> }
interface Fields {
  subject: string; start: string; end: string; timeZone: string; body: string;
  location: string; showAs: string; categories: string[]; importance: string;
  isReminderOn: boolean; reminderMinutesBeforeStart: number; isAllDay: boolean;
}
interface Operation { key: string; id: string; fields: Fields; observedShowAs: string }
interface Removal { key: string; id: string; reason: string }
interface Hold { key: string; reason: string }
interface Plan {
  version: number; sourceId: string; feedRunId: string; checkedAt: string;
  creates: Operation[]; updates: Operation[]; deletes: Removal[]; held: Hold[];
  nextState: State; previousState: State; sessionCount: number;
}
interface Ack { kind: string; key: string; id: string }
interface Result {
  status: string; planJson: string; stateJson: string; sessionCount: number;
  creates: Operation[]; updates: Operation[]; deletes: Removal[]; held: Hold[];
}

function main(
  workbook: ExcelScript.Workbook,
  mode: string,
  icsContent: string,
  existingEventsJson: string,
  stateJson: string,
  sourceId: string,
  feedRunId: string,
  fetchedAtUtc: string,
  scanComplete: boolean,
  planJson: string = "",
  acknowledgementsJson: string = "[]",
  rulesJson: string = "{}",
  nowUtc: string = ""
): Result {
  // Use Microsoft's runtime clock for freshness, not a timestamp captured before queueing.
  const now = new Date();
  if (!Number.isFinite(now.getTime())) throw new Error("Invalid current time.");
  if (mode === "failure") {
    if (workbook) {
      const sheet = workbook.getWorksheet("Preview") || workbook.addWorksheet("Preview");
      sheet.getUsedRange()?.clear();
      const applyFailed = planJson === "apply";
      sheet.getRange("A1:B3").setValues([["Status", applyFailed ? "Applying changes failed. Some operations may have completed; inspect the flow run." : "Source download or calendar scan failed. Existing appointments preserved."], ["Checked at UTC", now.toISOString()], ["Calendar writes", applyFailed ? "Stopped; inspect successful operations" : "Not attempted"]]);
    }
    return {status: "failed", planJson: "", stateJson: "", sessionCount: 0, creates: [], updates: [], deletes: [], held: []};
  }
  if (mode === "commit") return commitPlan(planJson, acknowledgementsJson, stateJson, now);
  if (mode !== "plan") throw new Error("Mode must be plan or commit.");
  const scanReference = nowUtc ? Date.parse(nowUtc) : now.getTime();
  if (!Number.isFinite(scanReference) || londonWall(scanReference).substring(0, 10) !== londonWall(now.getTime()).substring(0, 10))
    throw new Error("Calendar scan crossed a date boundary. Obtain a complete new scan.");
  if (!scanComplete) throw new Error("An incomplete Outlook scan cannot be reconciled.");
  if (!sourceId || sourceId.length > 100 || !feedRunId) throw new Error("Source ID and unique download run ID are required.");
  const fetched = Date.parse(fetchedAtUtc);
  if (!Number.isFinite(fetched) || now.getTime() - fetched > 900000 || fetched - now.getTime() > 60000)
    throw new Error("Feed freshness could not be verified. Existing events are preserved.");
  const state = readState(stateJson, sourceId);
  if (state.lastFeedRunId === feedRunId) throw new Error("This download has already been committed. Obtain a new feed file.");
  const outlook = JSON.parse(existingEventsJson) as OutlookEvent[];
  if (!Array.isArray(outlook) || outlook.length > 5000) throw new Error("Invalid Outlook scan.");
  const rules = JSON.parse(rulesJson) as Record<string, string>;
  Object.keys(rules).forEach((selector) => { if (!["free", "busy"].includes(rules[selector])) throw new Error("Rules must use free or busy."); });
  const start = londonWall(now.getTime()).substring(0, 10) + "T00:00:00";
  const end = new Date(Date.parse(start + "Z") + 365 * 86400000).toISOString().substring(0, 19);
  const sourceText = icsContent.startsWith('"') ? JSON.parse(icsContent) as string : icsContent;
  if (typeof sourceText !== "string") throw new Error("Calendar input must be text.");
  const sessions = parseCalendar(sourceText, sourceId, start, end);
  const plan = reconcile(sessions, outlook, state, rules, sourceId, feedRunId, now, start, end);
  writePreview(workbook, plan);
  console.log("Timetable plan: " + sessions.length + " sessions, " + plan.creates.length + " creates, " + plan.updates.length + " updates, " + plan.deletes.length + " deletes, " + plan.held.length + " held. Calendar not changed by this script.");
  return { status: "planned", planJson: JSON.stringify(plan), stateJson: "", sessionCount: sessions.length,
    creates: plan.creates, updates: plan.updates, deletes: plan.deletes, held: plan.held };
}

function readState(raw: string, sourceId: string): State {
  if (!raw || raw === "{}") return {version: 1, sourceId, lastFeedRunId: "", lastSuccess: "", events: {}};
  const state = JSON.parse(raw) as State;
  if (state.version !== 1 || state.sourceId !== sourceId || !state.events || typeof state.events !== "object")
    throw new Error("State does not belong to this source.");
  return state;
}
function property(ev: Definition, name: string): Property | undefined {
  const values = ev.properties.filter((p) => p.name === name);
  if (values.length > 1 && !["EXDATE", "RDATE", "CATEGORIES"].includes(name)) throw new Error("Duplicate " + name + " property.");
  return values[0];
}
function value(ev: Definition, name: string): string { return property(ev, name)?.value || ""; }
function decode(raw: string): string {
  return raw.replace(/\\([nN,;\\])/g, (match: string, escaped: string) => /[nN]/.test(escaped) ? "\n" : escaped);
}
function html(raw: string): string {
  return raw.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function lastSunday(year: number, month: number): number {
  const last = new Date(Date.UTC(year, month + 1, 0)); return last.getUTCDate() - last.getUTCDay();
}
function londonOffset(utc: number): number {
  const year = new Date(utc).getUTCFullYear();
  return utc >= Date.UTC(year, 2, lastSunday(year, 2), 1) && utc < Date.UTC(year, 9, lastSunday(year, 9), 1) ? 3600000 : 0;
}
function londonWall(utc: number): string { return new Date(utc + londonOffset(utc)).toISOString().substring(0, 19); }
function wallInstant(wall: string, zone: string): number {
  const nominal = Date.parse(wall + "Z");
  if (!Number.isFinite(nominal) || new Date(nominal).toISOString().substring(0, 19) !== wall) throw new Error("Invalid calendar date.");
  if (zone === "UTC") return nominal;
  const candidates = [nominal, nominal - 3600000].filter((n) => londonWall(n) === wall);
  if (candidates.length !== 1) throw new Error("Ambiguous or nonexistent London time requires review.");
  return candidates[0];
}
function calendarDate(p: Property): CalendarDate {
  if (p.params.VALUE && p.params.VALUE !== "DATE-TIME") throw new Error("All-day source dates are unsupported.");
  const m = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/.exec(p.value);
  if (!m) throw new Error("A valid iCalendar date-time is required.");
  const zone = m[7] ? "UTC" : (p.params.TZID || "Europe/London");
  if (!["UTC", "Europe/London", "GMT Standard Time"].includes(zone)) throw new Error("Unsupported source timezone: " + zone);
  const wall = m[1] + "-" + m[2] + "-" + m[3] + "T" + m[4] + ":" + m[5] + ":" + m[6];
  return {wall, utc: wallInstant(wall, zone), zone};
}
function parseDefinitions(raw: string): Definition[] {
  const text = raw.replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n[ \t]/g, "").trim();
  if (text.length > 5000000 || !text.startsWith("BEGIN:VCALENDAR\n") || !text.endsWith("END:VCALENDAR"))
    throw new Error("Feed is incomplete, oversized, or not iCalendar.");
  const stack: string[] = []; const definitions: Definition[] = []; let current: Definition | undefined; let validVersion = false;
  text.split("\n").forEach((line) => {
    if (line.startsWith("BEGIN:")) {
      const component = line.substring(6); stack.push(component);
      if (component === "VEVENT") {
        if (current || stack.length !== 2) throw new Error("Invalid event nesting.");
        current = {properties: []};
      }
    } else if (line.startsWith("END:")) {
      const component = line.substring(4);
      if (stack.pop() !== component) throw new Error("Invalid calendar nesting.");
      if (component === "VEVENT") { if (!current) throw new Error("Invalid event."); definitions.push(current); current = undefined; }
    } else if (stack.length === 1) {
      if (line === "VERSION:2.0") validVersion = true;
      if (line.startsWith("METHOD:") && line !== "METHOD:PUBLISH") throw new Error("Meeting messages are not timetable snapshots.");
    } else if (current && stack[stack.length - 1] === "VEVENT") {
      const colon = line.indexOf(":"); if (colon < 1) throw new Error("Malformed event property.");
      const parts = line.substring(0, colon).split(";"); const name = parts.shift()!.toUpperCase(); const params: Record<string, string> = {};
      parts.forEach((part) => { const index = part.indexOf("="); if (index < 1) throw new Error("Malformed parameter."); params[part.substring(0, index).toUpperCase()] = part.substring(index + 1).replace(/^"|"$/g, ""); });
      current.properties.push({name, params, value: line.substring(colon + 1)});
    }
  });
  if (!validVersion || stack.length || current || !definitions.length || definitions.length > 1000) throw new Error("Empty or invalid event feed refused.");
  return definitions;
}
function occurrenceDates(ev: Definition): CalendarDate[] {
  const initial = property(ev, "DTSTART"); if (!initial) throw new Error("Missing DTSTART.");
  const start = calendarDate(initial); const dates: CalendarDate[] = [start]; const rawRule = value(ev, "RRULE");
  if (rawRule) {
    const rule: Record<string, string> = {};
    rawRule.split(";").forEach((part) => { const pair = part.split("="); if (pair.length !== 2 || rule[pair[0]]) throw new Error("Invalid recurrence rule."); rule[pair[0]] = pair[1]; });
    if (!["DAILY", "WEEKLY"].includes(rule.FREQ)) throw new Error("Only daily and weekly teaching recurrences are supported. Unsupported rules stop the flow.");
    Object.keys(rule).forEach((key) => { if (!["FREQ", "INTERVAL", "COUNT", "UNTIL", "BYDAY", "WKST"].includes(key)) throw new Error("Unsupported recurrence field: " + key); });
    const interval = rule.INTERVAL ? Number(rule.INTERVAL) : 1; const count = rule.COUNT ? Number(rule.COUNT) : 10000;
    if (!Number.isInteger(interval) || interval < 1 || interval > 365 || !Number.isInteger(count) || count < 1 || count > 10000 || (rule.COUNT && rule.UNTIL)) throw new Error("Invalid recurrence bounds.");
    if (!rule.COUNT && !rule.UNTIL) throw new Error("An unbounded source recurrence requires review.");
    const until = rule.UNTIL ? calendarDate({name: "UNTIL", params: rule.UNTIL.endsWith("Z") ? {} : initial.params, value: rule.UNTIL}).utc : Infinity;
    const dayNames = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
    const startDay = new Date(start.wall + "Z").getUTCDay(); const weekdays = rule.BYDAY ? rule.BYDAY.split(",").map((d) => dayNames.indexOf(d)) : [startDay];
    const weekStart = rule.WKST ? dayNames.indexOf(rule.WKST) : 1;
    if (weekStart < 0 || weekdays.some((d) => d < 0) || (rule.BYDAY && !weekdays.includes(startDay))) throw new Error("Unsupported or inconsistent BYDAY/WKST.");
    const initialMidnight = Date.parse(start.wall.substring(0, 10) + "T00:00:00Z");
    const startWeek = initialMidnight - ((startDay - weekStart + 7) % 7) * 86400000;
    let generated = 1; let completed = count === 1;
    for (let day = 1; day <= 100000 && !completed; day++) {
      const candidateWall = new Date(Date.parse(start.wall + "Z") + day * 86400000).toISOString().substring(0, 19);
      const instant = wallInstant(candidateWall, start.zone);
      if (instant > until) { completed = true; break; }
      const candidateMidnight = initialMidnight + day * 86400000;
      const weekday = new Date(candidateMidnight).getUTCDay();
      const weekIndex = Math.floor((candidateMidnight - startWeek) / (7 * 86400000));
      const selected = rule.FREQ === "DAILY" ? day % interval === 0 && (!rule.BYDAY || weekdays.includes(weekday)) : weekIndex % interval === 0 && weekdays.includes(weekday);
      if (selected) { dates.push({wall: candidateWall, utc: instant, zone: start.zone}); generated++; if (generated >= count) completed = true; }
    }
    if (!completed) throw new Error("Recurrence exceeds expansion safety bound.");
    if (start.utc > until) throw new Error("Recurrence ends before it starts.");
  }
  ev.properties.filter((p) => p.name === "RDATE").forEach((p) => p.value.split(",").forEach((raw) => dates.push(calendarDate({name: p.name, params: p.params, value: raw}))));
  const excluded: number[] = [];
  ev.properties.filter((p) => p.name === "EXDATE").forEach((p) => p.value.split(",").forEach((raw) => excluded.push(calendarDate({name: p.name, params: p.params, value: raw}).utc)));
  return dates.filter((d, index) => !excluded.includes(d.utc) && dates.findIndex((other) => other.utc === d.utc) === index);
}
function parseCalendar(raw: string, sourceId: string, windowStart: string, windowEnd: string): Session[] {
  const definitions = parseDefinitions(raw); const sessions: Record<string, Session> = {}; const exceptionSeen: Definition[] = [];
  const masters = definitions.filter((ev) => !property(ev, "RECURRENCE-ID"));
  const exceptions = definitions.filter((ev) => property(ev, "RECURRENCE-ID"));
  definitions.forEach((ev) => { if (!value(ev, "UID")) throw new Error("Missing UID."); const rid = property(ev, "RECURRENCE-ID"); if (rid?.params.RANGE) throw new Error("RANGE exceptions require review."); });
  masters.forEach((master) => {
    const ending = property(master, "DTEND"); const beginning = property(master, "DTSTART");
    if (!beginning || !ending || property(master, "DURATION")) throw new Error("Explicit DTSTART/DTEND are required.");
    const initial = calendarDate(beginning); const initialEnd = calendarDate(ending);
    if (initial.zone !== initialEnd.zone) throw new Error("Mismatched start/end zones.");
    const wallDuration = Date.parse(initialEnd.wall + "Z") - Date.parse(initial.wall + "Z");
    if (wallDuration <= 0 || wallDuration > 86400000) throw new Error("Invalid teaching-session duration.");
    occurrenceDates(master).forEach((date) => {
      const matching = exceptions.filter((ev) => value(ev, "UID") === value(master, "UID") && calendarDate(property(ev, "RECURRENCE-ID")!).utc === date.utc);
      if (matching.length > 1) throw new Error("Duplicate recurrence exceptions.");
      const exception = matching[0]; if (exception) exceptionSeen.push(exception);
      const ev: Definition = exception ? {properties: master.properties.filter((p) => !exception.properties.some((q) => p.name === q.name)).concat(exception.properties)} : master;
      const changedStart = exception ? property(exception, "DTSTART") : undefined;
      const changedEnd = exception ? property(exception, "DTEND") : undefined;
      const start = changedStart ? calendarDate(changedStart) : date;
      const endWall = changedEnd ? calendarDate(changedEnd) : {wall: new Date(Date.parse(start.wall + "Z") + wallDuration).toISOString().substring(0, 19), zone: start.zone, utc: 0};
      const endingUtc = changedEnd ? endWall.utc : wallInstant(endWall.wall, endWall.zone);
      if (endingUtc <= start.utc || endingUtc - start.utc > 86400000) throw new Error("Invalid event duration.");
      const startLondon = londonWall(start.utc); const endLondon = londonWall(endingUtc);
      if (startLondon < windowStart || startLondon >= windowEnd) return;
      const uid = value(master, "UID"); const cmis = /^\d+ID(\d+):.*@timetabling\.port\.ac\.uk$/.exec(uid);
      const family = cmis ? "cmis:" + cmis[1] : "ical:" + uid;
      const identity = cmis ? londonWall(date.utc).substring(0, 10) : new Date(date.utc).toISOString();
      const key = encodeURIComponent(JSON.stringify([sourceId, family, identity]));
      const description = decode(value(ev, "DESCRIPTION")).trim(); const lines = description.split("\n").map((line) => line.trim()).filter((line) => !!line);
      const moduleMatch = /\b([MI]\d{5})(?:\/\w+)?\b/.exec(description);
      const activities = ["tutorial", "workshop", "lecture", "problem based learning", "self directive studies", "self directed studies", "computer aided teaching", "drop-in", "induction wk event"];
      const categories: string[] = [];
      ev.properties.filter((p) => p.name === "CATEGORIES").forEach((p) => {
        p.value.split(/(?<!\\),/).forEach((item) => categories.push(decode(item).trim()));
      });
      const activity = lines.concat(categories).find((line) => activities.includes(line.toLowerCase())) || "Other";
      const module = moduleMatch ? moduleMatch[1] : "";
      const name = module && lines.length > 1 ? lines[1] : decode(value(ev, "SUMMARY")) || "Class";
      const title = [module, name, activity === "Other" ? "" : activity].filter((part) => !!part).join(" - ");
      const session: Session = {key, family, identity, start: startLondon, end: endLondon, title, module, activity,
        location: decode(value(ev, "LOCATION")), description, cancelled: value(ev, "STATUS").toUpperCase() === "CANCELLED"};
      if (sessions[key] && JSON.stringify(sessions[key]) !== JSON.stringify(session)) throw new Error("Conflicting occurrence identity.");
      sessions[key] = session;
      if (Object.keys(sessions).length > 5000) throw new Error("Expanded calendar exceeds safety limit.");
    });
  });
  if (exceptions.some((ev) => !exceptionSeen.includes(ev))) throw new Error("An orphan or excluded recurrence exception requires review.");
  return Object.keys(sessions).map((key) => sessions[key]).sort((a, b) => a.start.localeCompare(b.start) || a.key.localeCompare(b.key));
}
function marker(source: string, key: string): string { return "Timetable Sync [v1:" + encodeURIComponent(source) + ":" + key + "]"; }
function eventMarker(body: string, source: string): string {
  const prefix = "Timetable Sync [v1:" + encodeURIComponent(source) + ":";
  const index = body.indexOf(prefix); if (index < 0) return "";
  const ending = body.indexOf("]", index + prefix.length);
  if (ending < 0) throw new Error("Incomplete ownership marker.");
  const key = body.substring(index + prefix.length, ending);
  const decoded = JSON.parse(decodeURIComponent(key)) as string[];
  if (decoded.length !== 3 || decoded[0] !== source) throw new Error("Invalid ownership marker.");
  return key;
}
function defaults(session: Session, rules: Record<string, string>): string {
  const selectors = ["event:" + session.key, "family:" + session.family, "module-activity:" + session.module.toLowerCase() + "|" + session.activity.toLowerCase(), "module:" + session.module.toLowerCase(), "activity:" + session.activity.toLowerCase(), "default"];
  const chosen = selectors.find((selector) => !!rules[selector]);
  return chosen ? rules[chosen] : (["self directive studies", "self directed studies"].includes(session.activity.toLowerCase()) ? "free" : "busy");
}
function fields(session: Session, showAs: string, current: OutlookEvent | undefined, sourceId: string): Fields {
  return {subject: session.title, start: session.start, end: session.end, timeZone: "GMT Standard Time",
    body: "<p>" + html(session.description).replace(/\n/g, "<br/>") + "</p><p><small>" + html(marker(sourceId, session.key)) + "</small></p>",
    location: session.location, showAs, categories: current?.categories || [], importance: current?.importance || "normal",
    isReminderOn: current?.isReminderOn ?? true, reminderMinutesBeforeStart: current?.reminderMinutesBeforeStart ?? 15, isAllDay: false};
}
function normalizedWall(raw: string, zone: string): string {
  if (/Z$|[+-]\d{2}:\d{2}$/.test(raw)) {
    const instant = Date.parse(raw); if (!Number.isFinite(instant)) throw new Error("Invalid Outlook timestamp."); return londonWall(instant);
  }
  const wall = raw.substring(0, 19);
  if (!["UTC", "Europe/London", "GMT Standard Time"].includes(zone)) throw new Error("Outlook timestamps need an explicit supported timezone.");
  return londonWall(wallInstant(wall, zone));
}
function reconcile(sessions: Session[], outlook: OutlookEvent[], state: State, rules: Record<string, string>, sourceId: string, feedRunId: string, now: Date, start: string, end: string): Plan {
  const nextState = JSON.parse(JSON.stringify(state)) as State;
  const plan: Plan = {version: 1, sourceId, feedRunId, checkedAt: now.toISOString(), creates: [], updates: [], deletes: [], held: [], nextState, previousState: state, sessionCount: sessions.length};
  const managed: Record<string, OutlookEvent> = {}; const byKey: Record<string, Session> = {};
  outlook.forEach((ev) => {
    if (!ev.id || typeof ev.body !== "string" || typeof ev.hasAttendees !== "boolean") throw new Error("Outlook event data is incomplete.");
    const key = eventMarker(ev.body, sourceId); if (!key) return;
    if (managed[key]) throw new Error("Duplicate owned appointment detected; review before applying.");
    if (ev.hasAttendees || ev.isAllDay || ev.isRecurring) throw new Error("Managed event became a meeting, all-day event or recurring series. Review required.");
    managed[key] = ev;
    if (!state.events[key]) {
      const identity = JSON.parse(decodeURIComponent(key)) as string[];
      nextState.events[key] = {id: ev.id, family: identity[1], identity: identity[2], start: normalizedWall(ev.start, ev.timeZone), end: normalizedWall(ev.end, ev.timeZone), lastShowAs: ev.showAs, override: ev.showAs, missingSince: "", fingerprint: ""};
    }
  });
  sessions.forEach((session) => { byKey[session.key] = session; });
  const missingKeys = Object.keys(state.events).filter((key) => state.events[key].start >= start && state.events[key].start < end && (!byKey[key] || byKey[key].cancelled));
  const newKeys = sessions.filter((s) => !s.cancelled && !managed[s.key] && !state.events[s.key]);
  const ambiguous = newKeys.filter((s) => missingKeys.some((key) => state.events[key].family === s.family));
  sessions.forEach((session) => {
    if (session.cancelled) return;
    const current = managed[session.key]; const saved = state.events[session.key];
    if (!current && saved?.id) {
      plan.held.push({key: session.key, reason: "Managed appointment is missing or moved outside the scan; automatic recreation held."}); return;
    }
    if (ambiguous.includes(session)) { plan.held.push({key: session.key, reason: "Possible cross-date move; identity must be reviewed."}); return; }
    const explicit = rules["event:" + session.key];
    const manualChange = current && saved && current.showAs !== saved.lastShowAs;
    const override = manualChange ? current!.showAs : (saved?.override || (current && !saved ? current.showAs : ""));
    const showAs = explicit || override || defaults(session, rules);
    if (!["free", "tentative", "busy", "oof", "workingElsewhere", "unknown"].includes(showAs)) throw new Error("Invalid availability value.");
    const desired = fields(session, showAs, current, sourceId);
    const operation: Operation = {key: session.key, id: current?.id || "", fields: desired, observedShowAs: current?.showAs || showAs};
    if (!current) plan.creates.push(operation);
    const fingerprint = JSON.stringify([session.title, session.description, session.location, session.start, session.end]);
    if (current && (current.subject !== desired.subject || normalizedWall(current.start, current.timeZone) !== desired.start || normalizedWall(current.end, current.timeZone) !== desired.end || current.location !== desired.location || current.showAs !== desired.showAs || saved?.fingerprint !== fingerprint)) plan.updates.push(operation);
    nextState.events[session.key] = {id: current?.id || "", family: session.family, identity: session.identity, start: session.start, end: session.end, lastShowAs: showAs, override, missingSince: "", fingerprint};
  });
  missingKeys.forEach((key) => {
    const saved = state.events[key]; if (saved.start < start || saved.start >= end) return;
    if (ambiguous.some((session) => session.family === saved.family)) { plan.held.push({key, reason: "Possible cross-date move; deletion held."}); return; }
    const current = managed[key];
    if (!current) { plan.held.push({key, reason: "Missing managed appointment requires review; state retained."}); return; }
    const candidate = nextState.events[key];
    if (!candidate.missingSince) candidate.missingSince = now.toISOString();
    if (saved.missingSince && !Number.isFinite(Date.parse(saved.missingSince))) throw new Error("Invalid deletion confirmation state.");
    if (!saved.missingSince || now.getTime() - Date.parse(saved.missingSince) < 300000) { plan.held.push({key, reason: "Missing or cancelled class awaiting another fresh download at least five minutes later."}); return; }
    plan.deletes.push({key, id: current.id, reason: byKey[key]?.cancelled ? "Source cancellation confirmed" : "Missing from two valid downloads"});
  });
  const registeredInWindow = Object.keys(state.events).filter((key) => state.events[key].start >= start && state.events[key].start < end).length;
  if (plan.deletes.length > 10 || (registeredInWindow > 0 && plan.deletes.length / registeredInWindow > 0.25)) {
    plan.deletes.forEach((op) => plan.held.push({key: op.key, reason: "Mass deletion limit exceeded; manual review required."})); plan.deletes = [];
  }
  nextState.lastFeedRunId = feedRunId; nextState.lastSuccess = now.toISOString();
  return plan;
}
function writePreview(workbook: ExcelScript.Workbook, plan: Plan): void {
  if (!workbook) return;
  const sheet = workbook.getWorksheet("Preview") || workbook.addWorksheet("Preview");
  const rows: string[][] = [["Checked at UTC", plan.checkedAt, "Sessions", String(plan.sessionCount), "", "", ""],
    ["Status", "Plan generated. Calendar writes are controlled separately by Power Automate.", "", "", "", "", ""],
    ["Action", "Title", "Start (London)", "End (London)", "Show as", "Reason", "Key"]];
  plan.creates.forEach((op) => rows.push(["Create", op.fields.subject, op.fields.start, op.fields.end, op.fields.showAs, "New class", op.key]));
  plan.updates.forEach((op) => rows.push(["Update", op.fields.subject, op.fields.start, op.fields.end, op.fields.showAs, "Source or rule changed", op.key]));
  plan.deletes.forEach((op) => rows.push(["Delete", "Managed class", plan.previousState.events[op.key].start, plan.previousState.events[op.key].end, "", op.reason, op.key]));
  plan.held.forEach((hold) => rows.push(["Held", "", "", "", "", hold.reason, hold.key]));
  sheet.getUsedRange()?.clear();
  sheet.getRangeByIndexes(0, 0, rows.length, 7).setValues(rows.map((row) => row.map((cell) => /^[=+\-@]/.test(cell) ? "'" + cell : cell)));
  sheet.getRange("A:G").getFormat().autofitColumns();
}
function commitPlan(planJson: string, acknowledgementsJson: string, currentStateJson: string, now: Date): Result {
  const plan = JSON.parse(planJson) as Plan; const acks = JSON.parse(acknowledgementsJson) as Ack[];
  if (plan.version !== 1 || !Array.isArray(acks) || !Number.isFinite(Date.parse(plan.checkedAt)) || now.getTime() - Date.parse(plan.checkedAt) > 900000 || Date.parse(plan.checkedAt) - now.getTime() > 60000) throw new Error("Invalid or stale commit.");
  const current = readState(currentStateJson, plan.sourceId);
  if (JSON.stringify(current) !== JSON.stringify(plan.previousState)) throw new Error("State changed while the flow was applying. Commit refused.");
  const expected = plan.creates.map((op) => ({kind: "create", key: op.key, id: ""})).concat(plan.updates.map((op) => ({kind: "update", key: op.key, id: op.id})), plan.deletes.map((op) => ({kind: "delete", key: op.key, id: op.id})));
  if (acks.length !== expected.length) throw new Error("Incomplete operation acknowledgements. State was not advanced.");
  expected.forEach((operation) => {
    const matching = acks.filter((ack) => ack.kind === operation.kind && ack.key === operation.key && !!ack.id);
    if (matching.length !== 1 || (operation.id && matching[0].id !== operation.id)) throw new Error("Invalid operation acknowledgement.");
    if (operation.kind === "create") plan.nextState.events[operation.key].id = matching[0].id;
    if (operation.kind === "delete") delete plan.nextState.events[operation.key];
  });
  return {status: "committed", planJson: "", stateJson: JSON.stringify(plan.nextState), sessionCount: plan.sessionCount, creates: [], updates: [], deletes: [], held: plan.held};
}
