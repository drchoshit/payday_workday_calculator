const http = require('node:http');
const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const { PostgresStateStore } = require('./shared-state.cjs');

function createServer({store, password, allowedOrigin = ''}) {
  const signingKey = crypto.createHash('sha256').update('payroll-session:' + password).digest();
  const failures = new Map();
  const sign = text => crypto.createHmac('sha256',signingKey).update(text).digest('hex');
  const same = (a,b) => {
    const x=Buffer.from(String(a)),y=Buffer.from(String(b));
    return x.length===y.length && crypto.timingSafeEqual(x,y);
  };
  const authenticated = req => {
    const token=(req.headers.authorization || '').replace(/^Bearer /,'');
    const [expires,signature]=token.split('.');
    return /^\d+$/.test(expires || '') && Number(expires)>Date.now() && same(signature,sign(expires));
  };
  return http.createServer(async (req,res) => {
    const url = new URL(req.url,'http://localhost');
    const origin = req.headers.origin;
    res.setHeader('Cache-Control','no-store');
    res.setHeader('X-Content-Type-Options','nosniff');
    res.setHeader('Referrer-Policy','same-origin');
    if (origin && allowedOrigin && origin === allowedOrigin) {
      res.setHeader('Access-Control-Allow-Origin',allowedOrigin);
      res.setHeader('Vary','Origin');
      res.setHeader('Access-Control-Allow-Headers','Authorization, Content-Type');
      res.setHeader('Access-Control-Allow-Methods','GET, PUT, POST, OPTIONS');
    }
    const json=(status,data) => {res.writeHead(status,{'Content-Type':'application/json; charset=utf-8'});res.end(JSON.stringify(data));};
    if (req.method==='OPTIONS') {res.writeHead(204);res.end();return;}
    const body=async () => {
      if (!(req.headers['content-type'] || '').startsWith('application/json')) throw new Error('JSON required');
      let size=0,chunks=[];
      for await (const chunk of req) {size+=chunk.length;if(size>10*1024*1024) throw new Error('Payload too large');chunks.push(chunk);}
      return JSON.parse(Buffer.concat(chunks).toString('utf8'));
    };
    try {
      if(url.pathname==='/healthz') {json(200,{ok:true});return;}
      if(url.pathname==='/api/session' && req.method==='POST') {
        const address=req.socket.remoteAddress;
        const failure=failures.get(address);
        if(failure && failure.count>=10 && Date.now()-failure.at<60000) {json(429,{error:'잠시 후 다시 로그인하세요.'});return;}
        const input=await body();
        if(!same(input.password,password)) {
          failures.set(address,{count:(failure && Date.now()-failure.at<60000 ? failure.count : 0)+1,at:Date.now()});
          json(401,{error:'공유 저장 비밀번호가 맞지 않습니다.'});return;
        }
        failures.delete(address);
        const expires=String(Date.now()+7*24*3600*1000);
        json(200,{token:expires+'.'+sign(expires)});return;
      }
      if(url.pathname==='/api/state') {
        if(!authenticated(req)) {json(401,{error:'공유 저장 로그인이 필요합니다.'});return;}
        if(req.method==='GET') {json(200,await store.read());return;}
        if(req.method==='PUT') {
          const data=await body();
          if(!Number.isSafeInteger(data.revision) || data.revision<0 || !data.payload || typeof data.payload!=='object' || Array.isArray(data.payload)) {json(400,{error:'잘못된 저장 데이터입니다.'});return;}
          const result=await store.write(data.revision,data.payload);
          json(result ? 200 : 409,result || {error:'다른 컴퓨터에서 먼저 저장했습니다.',...await store.read()});return;
        }
        json(405,{error:'Method not allowed'});return;
      }
      // Serve an explicit allowlist: never expose .git, backups, or server secrets.
      const files={'/':'index.html','/index.html':'index.html','/app.js':'app.js','/shared-storage.js':'shared-storage.js','/styles.css':'styles.css'};
      const file=files[url.pathname];
      if(!file || req.method!=='GET') {json(404,{error:'Not found'});return;}
      const content=await fs.readFile(path.join(__dirname,file));
      const type=file.endsWith('.html') ? 'text/html' : file.endsWith('.css') ? 'text/css' : 'text/javascript';
      res.writeHead(200,{'Content-Type':type+'; charset=utf-8'});res.end(content);
    } catch(error) {
      json(error instanceof SyntaxError || /JSON required|Payload too large/.test(error.message) ? 400 : 503,
        {error:'공유 저장에 실패했습니다. 이 컴퓨터의 데이터는 유지됩니다.'});
    }
  });
}

async function main() {
  if (!process.env.DATABASE_URL || !process.env.SHARED_PASSWORD) throw new Error('DATABASE_URL and SHARED_PASSWORD are required');
  const { Pool } = require('pg');
  const pool=new Pool({connectionString:process.env.DATABASE_URL,
    ssl:process.env.DATABASE_SSL==='true' ? {rejectUnauthorized:true} : undefined,
    connectionTimeoutMillis:10000,statement_timeout:15000});
  const store=new PostgresStateStore(pool);
  await store.initialize();
  const server=createServer({store,password:process.env.SHARED_PASSWORD,allowedOrigin:process.env.ALLOWED_ORIGIN || ''});
  server.listen(Number(process.env.PORT || 5500),'0.0.0.0');
}
if (require.main===module) main().catch(error => {console.error(error.message);process.exitCode=1;});
module.exports={createServer};
