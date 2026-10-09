const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const assert = require('node:assert/strict');
const {stripTypeScriptTypes} = require('node:module');
const code = fs.readFileSync(path.join(__dirname, 'TimetableSync.ts'), 'utf8');
const context = {Date, console: {log:()=>{}}};
vm.createContext(context);
vm.runInContext(stripTypeScriptTypes(code, {mode: 'strip'}) + '\nthis.api={main,parseCalendar,marker,londonWall};', context);
const api = context.api;
const normalize = value => JSON.parse(JSON.stringify(value));
const source = 'synthetic-school';
const now = '2026-10-01T08:00:00Z';
let previewRows = [];
const worksheet = {
  getUsedRange: () => ({clear: () => {previewRows=[];}}),
  getRangeByIndexes: () => ({setValues: rows => {previewRows=normalize(rows);}}),
  getRange: () => ({getFormat: () => ({autofitColumns: () => {}})})
};
const workbook = {getWorksheet: () => worksheet, addWorksheet: () => worksheet};
function freezeClock(utc) {
  context.Date=class extends Date { constructor(value) { super(arguments.length?value:utc); } static now() {return Date.parse(utc);} };
}
freezeClock(now);
const event = (uid, beginning, ending, extras = '', title = 'Example course') =>
  `BEGIN:VEVENT\r\nUID:${uid}\r\nDTSTART;TZID=Europe/London:${beginning}\r\nDTEND;TZID=Europe/London:${ending}\r\nSUMMARY:Room 101\r\nLOCATION:Room 101\r\nDESCRIPTION:M12345\\n${title}\\nTeaching Event\\nWorkshop\r\n${extras}END:VEVENT\r\n`;
const calendar = events => `BEGIN:VCALENDAR\r\nVERSION:2.0\r\n${events}END:VCALENDAR\r\n`;
const base = event('0ID123:example@timetabling.port.ac.uk', '20261002T100000', '20261002T120000', 'RRULE:FREQ=WEEKLY;COUNT=4;BYDAY=FR\r\n');
const parse = text => normalize(api.parseCalendar(text, source, '2026-09-01T00:00:00', '2027-09-01T00:00:00'));
const plan = (text, outlook = [], state = '{}', checked = now, run = 'download-1', rules = '{}', complete = true) => {
  freezeClock(checked); return normalize(api.main(workbook, 'plan', text, JSON.stringify(outlook), state, source, run, checked, complete, '', '[]', rules, checked));
};
const commit = (result, acks, state = '{}', checked = now) => {
  freezeClock(checked); return normalize(api.main({}, 'commit', '', '[]', state, source, '', '', false, result.planJson, JSON.stringify(acks), '{}', checked));
};
const asOutlook = (op, id = 'owned-1') => ({id, ...op.fields, hasAttendees: false, isRecurring: false});
let passed = 0;
function test(name, fn) { fn(); passed++; console.log('PASS ' + name); }

test('a successful plan writes a reviewable Preview worksheet', () => {
  const result=plan(calendar(base));
  assert.equal(previewRows[0][3],String(result.sessionCount));
  assert.equal(previewRows[3][0],'Create');
  assert.equal(previewRows[3][1],result.creates[0].fields.subject);
});

test('managed category is added to new appointments', () => {
  const result=plan(calendar(base),[], '{}',now,'category-new',JSON.stringify({calendarCategory:'Timetable'}));
  assert.ok(result.creates.every(op=>op.fields.categories.includes('Timetable')));
});
test('adding a category preserves existing labels and Free choice without repeat updates', () => {
  const initial=plan(calendar(base));
  const events=initial.creates.map((op,i)=>({...asOutlook(op,'category-'+i),categories:['Personal label'],showAs:'free'}));
  const stored=commit(initial,initial.creates.map((op,i)=>({kind:'create',key:op.key,id:'category-'+i}))).stateJson;
  const rules=JSON.stringify({calendarCategory:'Timetable'});
  const amended=plan(calendar(base),events,stored,now,'category-update',rules);
  assert.equal(amended.creates.length,0);
  assert.equal(amended.updates.length,events.length);
  amended.updates.forEach(op=>{assert.deepEqual(op.fields.categories,['Personal label','Timetable']);assert.equal(op.fields.showAs,'free');});
  const committed=commit(amended,amended.updates.map(op=>({kind:'update',key:op.key,id:op.id})),stored).stateJson;
  const updated=amended.updates.map(op=>asOutlook(op,op.id));
  const repeated=plan(calendar(base),updated,committed,now,'category-repeat',rules);
  assert.equal(repeated.updates.length,0);
});

test('weekly recurrence, title extraction and stable CMIS identity', () => {
  const sessions = parse(calendar(base)); assert.equal(sessions.length, 4);
  assert.equal(sessions[0].title, 'M12345 - Example course - Workshop');
  assert.equal(sessions[0].key, parse(calendar(base.replace('0ID123:', '7ID123:')))[0].key);
});
test('multiple BYDAY values respect COUNT and INTERVAL', () => {
  const text = event('weekly', '20261002T100000', '20261002T120000', 'RRULE:FREQ=WEEKLY;COUNT=4;INTERVAL=2;BYDAY=FR,MO\r\n');
  assert.deepEqual(parse(calendar(text)).map(s => s.start.substring(0,10)), ['2026-10-02','2026-10-12','2026-10-16','2026-10-26']);
});
test('EXDATE and RDATE alter individual occurrences', () => {
  const text = base.replace('END:VEVENT', 'EXDATE;TZID=Europe/London:20261009T100000\r\nRDATE;TZID=Europe/London:20261030T100000\r\nEND:VEVENT');
  assert.deepEqual(parse(calendar(text)).map(s => s.start.substring(0,10)), ['2026-10-02','2026-10-16','2026-10-23','2026-10-30']);
});
test('a moved occurrence retains its original identity', () => {
  const exception = event('0ID123:example@timetabling.port.ac.uk', '20261010T140000', '20261010T160000', 'RECURRENCE-ID;TZID=Europe/London:20261009T100000\r\n');
  const sessions = parse(calendar(base + exception));
  assert.equal(sessions[1].start, '2026-10-10T14:00:00'); assert.equal(sessions[1].identity, '2026-10-09');
});
test('cancelled exceptions are retained as explicit cancellations', () => {
  const exception = 'BEGIN:VEVENT\r\nUID:0ID123:example@timetabling.port.ac.uk\r\nRECURRENCE-ID;TZID=Europe/London:20261009T100000\r\nSTATUS:CANCELLED\r\nEND:VEVENT\r\n';
  assert.equal(parse(calendar(base + exception))[1].cancelled, true);
});
test('London recurrence preserves wall time across the DST boundary', () => {
  const text = event('dst', '20270326T100000', '20270326T120000', 'RRULE:FREQ=WEEKLY;COUNT=2;BYDAY=FR\r\n');
  assert.deepEqual(parse(calendar(text)).map(s => s.start), ['2027-03-26T10:00:00','2027-04-02T10:00:00']);
});
test('UTC source timestamps convert to London rather than stripping Z', () => {
  const text = base.replace(/DTSTART;TZID=Europe\/London:20261002T100000/, 'DTSTART:20261002T090000Z').replace(/DTEND;TZID=Europe\/London:20261002T120000/, 'DTEND:20261002T110000Z');
  assert.equal(parse(calendar(text))[0].start, '2026-10-02T10:00:00');
});
test('empty, malformed and unsupported feeds stop safely', () => {
  assert.throws(() => parse(calendar('')), /Empty/);
  assert.throws(() => parse('<html>error</html>'), /not iCalendar/);
  assert.throws(() => parse(calendar(base.replace('FREQ=WEEKLY','FREQ=MONTHLY'))), /Only daily and weekly/);
  assert.throws(() => parse(calendar(base.replace('COUNT=4','COUNT=4;BYSETPOS=1'))), /Unsupported recurrence/);
});
test('stale downloads and incomplete Outlook scans cannot produce a plan', () => {
  assert.throws(() => plan(calendar(base), [], '{}', now, 'x', '{}', false), /incomplete Outlook scan/);
  assert.throws(() => api.main({}, 'plan', calendar(base), '[]', '{}', source, 'x', '2026-09-30T08:00:00Z', true, '', '[]', '{}', now), /freshness/);
});
test('JSON-quoted text supports Excel’s single-line parameter form', () => {
  assert.equal(plan(JSON.stringify(calendar(base))).sessionCount,4);
});
test('queue delays cannot disguise a stale feed using an earlier caller timestamp', () => {
  freezeClock('2026-10-01T08:30:00Z');
  assert.throws(()=>api.main({},'plan',calendar(base),'[]','{}',source,'queued',now,true,'','[]','{}',now),/freshness/);
});
test('a scan crossing London midnight must be repeated', () => {
  freezeClock('2026-10-02T00:05:00Z');
  assert.throws(()=>api.main({},'plan',calendar(base),'[]','{}',source,'boundary','2026-10-02T00:05:00Z',true,'','[]','{}','2026-10-01T22:55:00Z'),/date boundary/);
});
test('creating and committing a batch records IDs and avoids repeat creates', () => {
  const first = plan(calendar(base)); assert.equal(first.creates.length, 4);
  const acks = first.creates.map((op,index) => ({kind:'create',key:op.key,id:'id-'+index}));
  const state = commit(first, acks).stateJson;
  const current = first.creates.map((op,index) => asOutlook(op,'id-'+index));
  const second = plan(calendar(base),current,state,'2026-10-01T08:05:00Z','download-2');
  assert.equal(second.creates.length,0); assert.equal(second.updates.length,0);
});
test('manual Free choice survives a source room update', () => {
  const first = plan(calendar(base)); const state = commit(first, first.creates.map((op,i)=>({kind:'create',key:op.key,id:'id-'+i}))).stateJson;
  const current = first.creates.map((op,i)=>asOutlook(op,'id-'+i)); current[0].showAs='free';
  const result = plan(calendar(base.replace('LOCATION:Room 101','LOCATION:Room 202')),current,state,'2026-10-01T08:05:00Z','download-2');
  assert.equal(result.updates[0].fields.showAs,'free'); assert.equal(result.updates[0].fields.location,'Room 202');
});
test('deletion needs two fresh committed observations at least five minutes apart', () => {
  const first = plan(calendar(base)); let state = commit(first, first.creates.map((op,i)=>({kind:'create',key:op.key,id:'id-'+i}))).stateJson;
  const current = first.creates.map((op,i)=>asOutlook(op,'id-'+i));
  const excluded = calendar(base.replace('END:VEVENT','EXDATE;TZID=Europe/London:20261009T100000\r\nEND:VEVENT'));
  const missing = plan(excluded,current,state,'2026-10-01T08:05:00Z','download-2'); assert.equal(missing.deletes.length,0);
  state = commit(missing,[],state,'2026-10-01T08:05:00Z').stateJson;
  assert.equal(plan(excluded,current,state,'2026-10-01T08:06:00Z','download-3').deletes.length,0);
  assert.equal(plan(excluded,current,state,'2026-10-01T08:10:00Z','download-4').deletes.length,1);
});
test('an unexplained cross-date move is held rather than duplicated or deleted', () => {
  const first = plan(calendar(base)); const state = commit(first,first.creates.map((op,i)=>({kind:'create',key:op.key,id:'id-'+i}))).stateJson;
  const current = first.creates.map((op,i)=>asOutlook(op,'id-'+i));
  const shifted = calendar(base.replace(/20261002/g,'20261003').replace('BYDAY=FR','BYDAY=SA'));
  const result = plan(shifted,current,state,'2026-10-01T08:05:00Z','download-2'); assert.equal(result.creates.length,0); assert.equal(result.deletes.length,0); assert.ok(result.held.length>0);
});
test('state loss and a failed create checkpoint recover existing owned events', () => {
  const first = plan(calendar(base)); const current = first.creates.map((op,i)=>asOutlook(op,'id-'+i));
  const recovered = plan(calendar(base),current,'{}','2026-10-01T08:05:00Z','download-2'); assert.equal(recovered.creates.length,0);
});
test('unrelated personal appointments are never managed', () => {
  const result = plan(calendar(base), [{...asOutlook(plan(calendar(base)).creates[0]),id:'personal',body:'Personal appointment'}]);
  assert.equal(result.creates.length,4); assert.equal(result.updates.length,0); assert.equal(result.deletes.length,0);
});
test('owned meetings and duplicate ownership markers stop the batch', () => {
  const first = plan(calendar(base)); const one = asOutlook(first.creates[0]);
  assert.throws(()=>plan(calendar(base),[{...one,hasAttendees:true}]),/meeting/);
  assert.throws(()=>plan(calendar(base),[{...one,isRecurring:true}]),/recurring series/);
  assert.throws(()=>plan(calendar(base),[one,{...one,id:'duplicate'}]),/Duplicate owned/);
});
test('past state records do not block normal future occurrences', () => {
  const first=plan(calendar(base)); const state=commit(first,first.creates.map((op,i)=>({kind:'create',key:op.key,id:'id-'+i}))).stateJson;
  const later=base.replace(/20261002/g,'20261106');
  assert.equal(plan(calendar(later),[],state,'2026-11-01T08:00:00Z','later-run').creates.length,4);
});
test('incomplete acknowledgements never advance state', () => {
  const first = plan(calendar(base)); assert.throws(()=>commit(first,[]),/Incomplete/);
});
test('changed state cannot be overwritten by an earlier plan', () => {
  const first = plan(calendar(base)); const acks=first.creates.map((op,i)=>({kind:'create',key:op.key,id:'id-'+i}));
  const committed = commit(first,acks).stateJson; assert.throws(()=>commit(first,acks,committed),/State changed/);
});
if (process.argv[2]) test('private feed expansion agrees with the Python RFC parser', () => {
  const sourceText=fs.readFileSync(process.argv[2],'utf8'); const expected=JSON.parse(fs.readFileSync(process.argv[3],'utf8'));
  const actual=api.parseCalendar(sourceText,'portsmouth-2026-2027','2026-09-01T00:00:00','2027-09-01T00:00:00');
  assert.equal(actual.length,expected.length);
  const a=actual.map(s=>[s.family,s.identity,s.start,s.end,s.title,s.location]);
  const b=expected.map(s=>[s.family,s.identity_date,s.start.substring(0,19),s.end.substring(0,19),s.title,s.location]);
  a.forEach((row,index) => row.forEach((field,column) => assert.equal(field===b[index][column],true,`Private-feed mismatch at occurrence ${index}, field ${column}; values suppressed.`)));
});
console.log(`${passed} planner checks passed.`);
