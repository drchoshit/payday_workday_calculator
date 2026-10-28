(function () {
  const equal=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
  const isObject=value=>value!==null && typeof value==='object' && !Array.isArray(value);
  const clone=value=>value===undefined ? undefined : JSON.parse(JSON.stringify(value));
  function merge(base,local,remote,path='') {
    if(equal(local,base)) return clone(remote);
    if(equal(remote,base) || equal(local,remote)) return clone(local);
    if(isObject(base) && isObject(local) && isObject(remote)) {
      const result={};
      for(const key of new Set([...Object.keys(base),...Object.keys(local),...Object.keys(remote)])) {
        if(['__proto__','prototype','constructor'].includes(key)) throw new Error('Invalid field');
        const value=merge(base[key],local[key],remote[key],path+'/'+key);
        if(value!==undefined) result[key]=value;
      }
      return result;
    }
    throw new Error('다른 컴퓨터에서 같은 항목을 수정했습니다. 로컬 백업을 내보낸 뒤 공유 데이터를 다시 불러오세요. ('+path+')');
  }
  globalThis.PayrollSharedStorage=class {
    constructor({snapshot,restore,status}) {
      this.snapshot=snapshot;this.restore=restore;this.status=status;
      this.base=null;this.revision=0;this.ready=false;this.busy=false;this.paused=false;
      this.endpoint=localStorage.getItem('payroll-shared-endpoint') || location.origin;
      this.token=localStorage.getItem('payroll-shared-token') || '';
      this.cacheKey='payroll-shared-base:'+this.endpoint;
    }
    async request(method,data,route='/api/state',keepalive=false) {
      const response=await fetch(this.endpoint+route,{method,cache:'no-store',keepalive,
        headers:{'Content-Type':'application/json','Authorization':'Bearer '+this.token},
        ...(data===undefined ? {} : {body:JSON.stringify(data)})});
      if(!(response.headers.get('content-type') || '').includes('application/json')) throw new Error('공유 저장 서버가 연결되지 않았습니다. 현재는 이 컴퓨터에만 저장됩니다.');
      const body=await response.json();
      if(!response.ok && response.status!==409) throw new Error(body.error || '공유 저장 연결 실패');
      return {code:response.status,...body};
    }
    setBase(payload,revision) {
      this.base=clone(payload);this.revision=revision;
      localStorage.setItem(this.cacheKey,JSON.stringify({payload:this.base,revision}));
    }
    backup() {
      localStorage.setItem('payroll-before-shared-load',JSON.stringify(this.snapshot()));
    }
    async connect(endpoint,password) {
      const url=new URL(endpoint || location.origin);
      if(url.protocol!=='https:' && !['localhost','127.0.0.1'].includes(url.hostname)) throw new Error('HTTPS 주소를 입력하세요.');
      if(url.username || url.password || url.pathname!=='/' || url.search || url.hash) throw new Error('공유 서버의 기본 주소만 입력하세요.');
      this.ready=false;this.token='';
      this.endpoint=url.origin;this.cacheKey='payroll-shared-base:'+this.endpoint;
      const auth=await this.request('POST',{password},'/api/session');
      this.token=auth.token;
      localStorage.setItem('payroll-shared-endpoint',this.endpoint);
      localStorage.setItem('payroll-shared-token',this.token);
      this.paused=false;
      await this.start();
    }
    async start({reload=false}={}) {
      if(!this.token) {this.status('이 컴퓨터에 자동 저장 중 · 다른 컴퓨터와 공유하려면 공유 저장에 연결하세요.');return;}
      this.ready=false;
      try {
        const remote=await this.request('GET');
        this.backup();
        const cached=JSON.parse(localStorage.getItem(this.cacheKey) || 'null');
        const local=this.snapshot();
        let next=remote.payload;
        if(remote.payload && cached && !reload) next=merge(cached.payload,local,remote.payload);
        this.setBase(remote.payload,remote.revision);
        if(next) this.restore(next);
        this.ready=true;this.paused=false;
        this.changed();
        this.poll ||= setInterval(()=>this.refresh(),15000);
      } catch(error) {this.paused=true;this.status(error.message,true);}
    }
    changed() {
      if(!this.ready || this.paused) return;
      clearTimeout(this.timer);
      this.status('이 컴퓨터 저장 완료 · 공유 저장 중…');
      this.timer=setTimeout(()=>this.flush(),200);
    }
    async flush(keepalive=false) {
      clearTimeout(this.timer);
      clearTimeout(this.retry);this.retry=null;
      if(!this.ready || this.paused || this.busy) return;
      const payload=this.snapshot();
      if(equal(payload,this.base)) {this.status('공유 저장 완료 · 다른 컴퓨터에서도 확인할 수 있습니다.');return;}
      this.busy=true;
      try {
        const saved=await this.request('PUT',{revision:this.revision,payload},'/api/state',keepalive);
        if(saved.code===409) {
          const next=merge(this.base,payload,saved.payload);
          this.setBase(saved.payload,saved.revision);
          // Include edits typed while the request was in flight.
          this.restore(merge(payload,this.snapshot(),next));
        } else this.setBase(payload,saved.revision);
        this.status('공유 저장 완료 · 다른 컴퓨터에서도 확인할 수 있습니다.');
      } catch(error) {
        this.status(error.message+' 이 컴퓨터의 입력은 보관되어 있습니다.',true);
        if(error.message.startsWith('다른 컴퓨터')) this.paused=true;
        else this.retry=setTimeout(()=>this.flush(),5000);
      } finally {
        this.busy=false;
        if(!this.paused && !this.retry && !equal(this.snapshot(),this.base)) this.changed();
      }
    }
    async refresh() {
      if(!this.ready || this.busy || this.paused) return;
      if(!equal(this.snapshot(),this.base)) {this.retry=null;await this.flush();return;}
      const active=document.activeElement;
      if(active && ['INPUT','TEXTAREA','SELECT'].includes(active.tagName)) return;
      try {
        this.busy=true;
        const remote=await this.request('GET');
        // Never replace an edit made while the read was in flight.
        if(remote.revision!==this.revision && equal(this.snapshot(),this.base)) {
          this.setBase(remote.payload,remote.revision);
          if(remote.payload) this.restore(remote.payload);
          this.status('다른 컴퓨터의 저장 내용을 불러왔습니다.');
        }
      } catch(error) {this.status(error.message,true);} finally {this.busy=false;}
    }
  };
  globalThis.PayrollMergeState=merge;
})();
