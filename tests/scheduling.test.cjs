const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const context = vm.createContext({ console, document: {querySelectorAll:()=>[],getElementById:()=>null}, localStorage:{getItem:()=>null,setItem:()=>{}}, window:{} });
vm.runInContext(fs.readFileSync('app.js','utf8').replace('init();',''),context);
vm.runInContext(`
function fixture() {
 state.managerProfiles = {};
 state.monthlyManagerConditions = {};
 state.scheduleFixedRules = [];
 state.scheduleTemplates = DEFAULT_SHIFT_TEMPLATES;
 state.scheduleWeeklyHeadcounts = normalizeWeeklyHeadcounts(null);
 for (let day=0;day<7;day++) for (const shift of SHIFT_IDS) state.scheduleWeeklyHeadcounts[day][shift] = 1;
 for (const name of ['장채민','김지원','대체']) {
   const p = ensureManagerProfile(name); p.weeklyMaximumShifts=7;
 }
 for (const day of [1,2,3]) ensureManagerProfile('장채민').prioritySlots[day]['1T']=true;
 return ['장채민','김지원','대체'];
}
`,context);
function run(code) { return vm.runInContext(code,context); }
let rows=run(`buildAutoScheduleForMonth('2026-11', fixture())`);
for(const row of rows.filter(r=>[1,2,3].includes(r.weekday)&&r.shiftId==='1T')) assert.equal(row.primaryManager,'장채민');
rows=run(`{ const names=fixture(); ensureManagerProfile('장채민').unavailableRanges=[{start:'2026-11-02',end:'2026-11-04'}]; buildAutoScheduleForMonth('2026-11',names); }`);
for(const row of rows.filter(r=>r.date>='2026-11-02'&&r.date<='2026-11-04')) assert.notEqual(row.primaryManager,'장채민');
assert.ok(rows.find(r=>r.date==='2026-11-02'&&r.shiftId==='1T').primaryManager);
rows=run(`{ const names=fixture(); ensureManagerProfile('장채민').monthlyMinimumShifts=10; buildAutoScheduleForMonth('2026-11',names); }`);
assert.ok(rows.filter(r=>r.primaryManager==='장채민'||r.secondaryManager==='장채민').length>=10);
run(`{ const names=fixture(); const p=ensureManagerProfile('장채민'); p.weeklyMaximumShifts=1; p.monthlyMinimumShifts=10; globalThis.result=buildAutoScheduleForMonth('2026-11',names); }`);
assert.ok(run(`getMonthlyMinimumWarnings(result,['장채민']).length`)===1);
assert.ok(run(`result.filter(r=>r.primaryManager==='장채민').length`) < 10);
// Legacy fixed preferences must never make blank availability cells eligible.
assert.equal(run(`{ fixture(); const p=ensureManagerProfile('김지원'); p.fixedPreference.enabled=true; p.availability[1]['1T']=false; isManagerAvailable('김지원','2026-11-02',1,'1T'); }`),false);
// Persisted minimum survives normalization / import.
assert.equal(run(`normalizeManagerProfiles({'장채민':{monthlyMinimumShifts:10}})['장채민'].monthlyMinimumShifts`),10);
// Two workers per shift, no duplicate daily assignment, weekly caps and unavailable dates.
rows=run(`{ const names=fixture(); for(let d=0;d<7;d++) state.scheduleWeeklyHeadcounts[d]['1T']=2; buildAutoScheduleForMonth('2026-11',names); }`);
const days=new Set(), weeks=new Map();
for(const row of rows) for(const name of [row.primaryManager,row.secondaryManager].filter(Boolean)) {
 const key=name+row.date; assert.ok(!days.has(key)); days.add(key);
 const week=run('getScheduleWeekKey('+JSON.stringify(row.date)+')');
 const wk=name+week; weeks.set(wk,(weeks.get(wk)||0)+1); assert.ok(weeks.get(wk)<=7);
}
console.log('Passed: starred shifts, unavailable replacement, monthly minimum, infeasible warning, blank availability, persistence, two-worker daily/weekly limits.');
