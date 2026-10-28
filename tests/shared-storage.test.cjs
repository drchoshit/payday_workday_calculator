const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const {createServer}=require('../shared-server.cjs');
const {mergeState}=require('../shared-state.cjs');

async function main() {
  let saved={revision:0,payload:null};
  const store={async read(){return structuredClone(saved);},async write(revision,payload){
    if(revision!==saved.revision) return null;
    saved={revision:revision+1,payload:structuredClone(payload)};return {revision:saved.revision};
  }};
  const server=createServer({store,password:'test-password',allowedOrigin:'https://payday-workday-calculator.onrender.com'});
  await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
  const endpoint='http://127.0.0.1:'+server.address().port;
  try {
    assert.equal((await fetch(endpoint+'/api/state')).status,401);
    assert.equal((await fetch(endpoint+'/.git/config')).status,404);
    assert.equal((await fetch(endpoint+'/shared-server.cjs')).status,404);
    assert.equal((await fetch(endpoint+'/api/session',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:'wrong'})})).status,401);
    function computer(initial) {
      let data=structuredClone(initial),status='';const memory=new Map();
      const context=vm.createContext({console,URL,location:{origin:endpoint},fetch,
        localStorage:{getItem:key=>memory.get(key)||null,setItem:(key,value)=>memory.set(key,value)},
        document:{activeElement:null},setTimeout:()=>1,clearTimeout:()=>{},setInterval:()=>1});
      vm.runInContext(fs.readFileSync('shared-storage.js','utf8'),context);
      const client=new context.PayrollSharedStorage({snapshot:()=>structuredClone(data),restore:value=>{data=structuredClone(value);},status:value=>{status=value;}});
      return {client,memory,get data(){return data;},get status(){return status;},edit(fn){fn(data);}};
    }
    const a=computer({settings:{taxRate:3.3},monthlyManagerConditions:{'2026-11':{A:{weeklyMaximumShifts:3},B:{weeklyMaximumShifts:2}}}});
    await a.client.connect(endpoint,'test-password');await a.client.flush();
    assert.equal(saved.revision,1);
    const b=computer({settings:{},entries:[]});
    await b.client.connect(endpoint,'test-password');await b.client.flush();
    assert.deepEqual(b.data,a.data); // A different computer loads authoritative shared data.
    assert(b.memory.has('payroll-before-shared-load'));
    a.edit(data=>{data.monthlyManagerConditions['2026-11'].A.weeklyMaximumShifts=4;});
    b.edit(data=>{data.monthlyManagerConditions['2026-11'].B.weeklyMaximumShifts=5;});
    await a.client.flush();await b.client.flush();await b.client.flush();
    assert.equal(saved.payload.monthlyManagerConditions['2026-11'].A.weeklyMaximumShifts,4);
    assert.equal(saved.payload.monthlyManagerConditions['2026-11'].B.weeklyMaximumShifts,5);
    await a.client.refresh();
    assert.equal(a.data.monthlyManagerConditions['2026-11'].B.weeklyMaximumShifts,5);
    a.edit(data=>{data.settings.taxRate=4;});b.edit(data=>{data.settings.taxRate=5;});
    await a.client.flush();await b.client.flush();
    assert.equal(b.client.paused,true);
    assert.equal(saved.payload.settings.taxRate,4); // Conflicting data cannot silently overwrite.
    assert.equal(b.data.settings.taxRate,5); // Unsynced local input remains intact.
    await b.client.start({reload:true});
    assert.equal(b.data.settings.taxRate,4);
    assert.equal(mergeState({x:1,y:1},{x:2,y:1},{x:1,y:2}).y,2);
    assert.throws(()=>mergeState({x:1},{x:2},{x:3}));
    console.log('Passed: authenticated shared storage, two computers, durable-store interface, CAS, unrelated concurrent changes, conflict protection, local backup, reload, private-file protection.');
  } finally {await new Promise(resolve=>server.close(resolve));}
}
main().catch(error=>{console.error(error);process.exitCode=1;});
