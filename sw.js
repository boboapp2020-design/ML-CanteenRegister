const CACHE='canteen-v58';
const ASSETS=['./','index.html','manifest.json','icon-192.png','icon-512.png'];
/* ดึงแบบข้าม HTTP cache ของเบราว์เซอร์เสมอ กันไฟล์เก่าค้าง ผู้ใช้ไม่ต้องล้างแคชเอง */
function fresh(url){return new Request(url,{cache:'reload'});}
self.addEventListener('install',e=>{e.waitUntil(
  caches.open(CACHE).then(c=>Promise.all(ASSETS.map(u=>fetch(fresh(u)).then(r=>r.ok?c.put(u,r):null).catch(()=>null))))
    .then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('message',e=>{if(e.data==='SKIP_WAITING')self.skipWaiting();});
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET')return;
  const u=new URL(e.request.url);
  if(u.origin!==self.location.origin)return;
  const isPage=e.request.mode==='navigate'||u.pathname.endsWith('/index.html')||u.pathname.endsWith('/');
  if(isPage){
    e.respondWith(fetch(fresh(e.request.url)).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put('index.html',cp));return res;}).catch(()=>caches.match('index.html').then(r=>r||caches.match(e.request))));
    return;
  }
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put(e.request,cp));return res;}).catch(()=>caches.match('index.html'))));
});
