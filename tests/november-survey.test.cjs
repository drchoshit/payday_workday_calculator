const fs=require('node:fs');
const vm=require('node:vm');
const assert=require('node:assert/strict');
const context=vm.createContext({console,document:{querySelectorAll:()=>[],getElementById:()=>null},
  localStorage:{getItem:()=>null,setItem:()=>{}},window:{}});
vm.runInContext(fs.readFileSync('app.js','utf8').replace('init();',''),context);
const run=code=>vm.runInContext(code,context);
run(`
state.scheduleTemplates=cloneDefaultShiftTemplates();
state.scheduleWeeklyHeadcounts=normalizeWeeklyHeadcounts(null);
for(let d=0;d<7;d++) for(const shift of SHIFT_IDS) state.scheduleWeeklyHeadcounts[d][shift]=1;
for(const name of ['장채민','권수현','김지원']) ensureBaseManagerProfile(name).weeklyMaximumShifts=7;
state.monthlyManagerConditions['2026-12']={'권수현':{weeklyMaximumShifts:2}};
globalThis.beforeDecember=JSON.stringify(state.monthlyManagerConditions['2026-12']);
globalThis.beforeExcluded=JSON.stringify(state.managerProfiles['권수현']);
applyNovemberSurvey();
`);
assert.equal(run(`state.scheduleMonth`),'2026-11');
assert.equal(run(`Object.keys(state.monthlyManagerConditions['2026-11']).length`),10);
assert.equal(run(`JSON.stringify(state.managerProfiles['권수현'])===beforeExcluded`),true);
assert.equal(run(`JSON.stringify(state.monthlyManagerConditions['2026-12'])===beforeDecember`),true);
assert.equal(run(`'권수현' in state.monthlyManagerConditions['2026-11']`),false);
assert.equal(run(`ensureManagerProfile('김진솔').weeklyMaximumShifts`),7); // Preference 2-3T is not imported.
assert.equal(run(`ensureManagerProfile('이유민').weeklyMaximumShifts`),7); // Preference 2T is not imported.
const expected={
  김지원:{'1T':[2,3,4,5,6],'2T':[4,5]},김진솔:{'1T':[1,2,4],'2T':[1,2,4]},
  장채민:{'1T':[1,2,3]},김민서:{'1T':[5],'2T':[4]},유시우:{'3T':[2,4,6]},
  이연승:{'1T':[0,6],'2T':[2]},이영서:{'1T':[0],'2T':[1]},김지연:{'3T':[1,3,5]},
  임아영:{'1T':[0,1,2,3,4],'2T':[0]},이유민:{'1T':[0,2,3,5],'2T':[0,2,3]},
};
for(const [name,shifts] of Object.entries(expected)) for(let d=0;d<7;d++) for(const shift of ['1T','2T','3T']) {
  assert.equal(run(`ensureManagerProfile(${JSON.stringify(name)}).availability[${d}][${JSON.stringify(shift)}]`),
    (shifts[shift] || []).includes(d),`${name} ${d} ${shift}`);
}
assert.equal(run(`getManagerTimeWindow('김민서',1,'2T').start`),undefined);
assert.equal(run(`getManagerTimeWindow('김지원',4,'2T').start`),undefined);
assert.equal(run(`getManagerTimeWindow('김민서',4,'2T','2026-11-12').start`),undefined);
assert.equal(run(`getManagerTimeWindow('김진솔',3,'2T').start`),undefined);
assert.equal(run(`isManagerAvailable('김지원','2026-11-04',3,'1T')`),false);
assert.equal(run(`isManagerAvailable('이연승','2026-11-14',6,'1T')`),false);
assert.equal(run(`isManagerAvailable('임아영','2026-11-30',1,'1T')`),false);
assert.equal(run(`isManagerAvailable('이유민','2026-11-22',0,'1T')`),false);
assert.equal(run(`ensureManagerProfile('이영서').unavailableRanges.length`),0); // December exception omitted.
const rows=run(`buildAutoScheduleForMonth('2026-11',Object.keys(state.monthlyManagerConditions['2026-11']))`);
for(const row of rows.filter(r=>[1,2,3].includes(r.weekday)&&r.shiftId==='1T')) {
  if(['2026-11-02','2026-11-23'].includes(row.date)) assert.notEqual(row.primaryManager,'장채민');
  else assert.equal(row.primaryManager,'장채민',row.date);
}
for(const row of rows) for(const prefix of ['primary','secondary']) {
  if(row[prefix+'Manager']) {
    assert.equal(row[prefix+'Start'],row.shiftStart);
    assert.equal(row[prefix+'End'],row.shiftEnd);
  }
}
run(`ensureManagerProfile('김지원').availability[1]['1T']=true; applyNovemberSurvey();`);
assert.equal(run(`ensureManagerProfile('김지원').availability[1]['1T']`),true); // Migration is not reapplied.
run(`state.scheduleMonth='2026-12';`);
assert.equal(run(`ensureManagerProfile('장채민').prioritySlots[1]['1T']`),false);
assert.equal(run(`ensureManagerProfile('장채민').availability[4]['1T']`),true);
run(`state.scheduleMonth='2026-11'; ensureManagerProfile('장채민').availability[4]['1T']=true;`);
assert.equal(run(`state.managerProfiles['장채민'].prioritySlots[1]['1T']`),false);
assert.equal(run(`buildPersistPayload().monthlyManagerConditions['2026-11']['장채민'].prioritySlots[1]['1T']`),true);
// Upgrade an existing v1 browser without replacing manual caps, stars, absences, or December.
run(`
beforeDecember=JSON.stringify(state.monthlyManagerConditions['2026-12']);
state.appliedUpdates=state.appliedUpdates.filter(id=>id!=='2026-11-full-shifts-v2');
const p=state.monthlyManagerConditions['2026-11'];
p['김지원'].availability[3]['2T']=true;
p['김지원'].timeWindows={'3:2T':{start:'18:30'}};
p['김지원'].prioritySlots[5]['2T']=true;
p['김지원'].weeklyMaximumShifts=4;
p['김진솔'].availability[3]['2T']=true;
p['김진솔'].timeWindows={'1:1T':{end:'15:00'}};
p['김민서'].availability[1]['2T']=true;
p['김민서'].availability[2]['2T']=true;
p['김민서'].timeWindows={'2026-11-12:2T':{start:'16:00'}};
p['이연승'].availability[1]['2T']=true;
applyNovemberSurvey();
`);
assert.equal(run(`ensureManagerProfile('김지원').availability[3]['2T']`),false);
assert.equal(run(`ensureManagerProfile('김지원').prioritySlots[5]['2T']`),true);
assert.equal(run(`ensureManagerProfile('김지원').weeklyMaximumShifts`),4);
assert.equal(run(`ensureManagerProfile('김진솔').availability[1]['1T']`),true);
assert.equal(run(`ensureManagerProfile('김진솔').availability[2]['2T']`),true);
assert.equal(run(`ensureManagerProfile('김진솔').availability[3]['2T']`),false);
assert.equal(run(`ensureManagerProfile('김민서').availability[4]['2T']`),true);
assert.equal(run(`ensureManagerProfile('이연승').availability[2]['2T']`),true);
assert.equal(run(`Object.values(state.monthlyManagerConditions['2026-11']).every(p=>Object.keys(p.timeWindows||{}).length===0)`),true);
assert.equal(run(`JSON.stringify(state.monthlyManagerConditions['2026-12'])===beforeDecember`),true);
assert.equal(run(`ensureManagerProfile('장채민').prioritySlots[1]['1T']`),true);
assert.equal(run(`isManagerAvailable('김지원','2026-11-04',3,'1T')`),false);
run(`ensureManagerProfile('김민서').availability[1]['2T']=true; applyNovemberSurvey();`);
assert.equal(run(`ensureManagerProfile('김민서').availability[1]['2T']`),true); // Upgrade is applied only once.
console.log('Passed: full-shift-only November survey, existing-browser upgrade, preserved manual stars/caps, absence dates, fixed Mon-Wed, month isolation, one-time migrations.');
