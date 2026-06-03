import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const R="#C8102E", G="#007A53", GOLD="#FFD700";

// ── Fixtures ──────────────────────────────────────────────────────────────────
// All times UTC. Grouped by calendar date for display.
const FIXTURES = [
  // ── May 31 (MD1) — verified KO times ──
  // South Korea vs T&T: BYU South Field, Utah ~01:00 UTC
  {id:"f01",home:"South Korea",away:"Trinidad & Tobago",ko:"2026-05-31T01:00:00Z",date:"2026-05-31",md:1},
  // Japan vs Iceland: Japan National Stadium, 15:55 JST = 06:55 UTC
  {id:"f02",home:"Japan",away:"Iceland",ko:"2026-05-31T06:55:00Z",date:"2026-05-31",md:1},
  // Ecuador vs Saudi Arabia: Red Bull Arena, NJ — 09:00 UTC est.
  {id:"f03",home:"Ecuador",away:"Saudi Arabia",ko:"2026-05-31T09:00:00Z",date:"2026-05-31",md:1},
  // Germany vs Finland: MEWA Arena Mainz, 19:45 CEST = 17:45 UTC ✓
  {id:"f04",home:"Germany",away:"Finland",ko:"2026-05-31T17:45:00Z",date:"2026-05-31",md:1},
  // USA vs Senegal: Bank of America Stadium, Charlotte — 3:30 PM ET = 19:30 UTC ✓
  {id:"f05",home:"USA",away:"Senegal",ko:"2026-05-31T19:30:00Z",date:"2026-05-31",md:1},
  // Mexico vs Australia: Rose Bowl, Pasadena — est. 3 PM PDT = 22:00 UTC
  {id:"f06",home:"Mexico",away:"Australia",ko:"2026-05-31T22:00:00Z",date:"2026-05-31",md:1},
  // Brazil vs Panama: Maracanã — est. 8 PM BRT = 23:00 UTC
  {id:"f07",home:"Brazil",away:"Panama",ko:"2026-05-31T23:00:00Z",date:"2026-05-31",md:1},
  // ── June 1 (MD2) ──
  // Poland vs Ukraine: Wrocław, est. 20:45 CEST = 18:45 UTC
  {id:"f08",home:"Poland",away:"Ukraine",ko:"2026-06-01T18:45:00Z",date:"2026-06-01",md:2},
  // Norway vs Sweden: Ullevaal, est. 20:00 CEST = 18:00 UTC
  {id:"f09",home:"Norway",away:"Sweden",ko:"2026-06-01T18:00:00Z",date:"2026-06-01",md:2},
  // Turkey vs North Macedonia: Ülker Stadyumu, est. 21:00 TRT = 18:00 UTC
  {id:"f10",home:"Turkey",away:"North Macedonia",ko:"2026-06-01T19:00:00Z",date:"2026-06-01",md:2},
  // ── June 2 (MD3) ──
  // Belgium vs Croatia: confirmed June 2 — est. 20:45 CEST = 18:45 UTC
  {id:"f11",home:"Belgium",away:"Croatia",ko:"2026-06-02T18:45:00Z",date:"2026-06-02",md:3},
  // Netherlands vs Algeria: est. 20:45 CEST = 18:45 UTC
  {id:"f12",home:"Netherlands",away:"Algeria",ko:"2026-06-02T18:45:00Z",date:"2026-06-02",md:3},
  // ── June 4 (MD4) ──
  // Spain vs Iraq: La Coruña — 21:00 CEST = 19:00 UTC ✓
  {id:"f13",home:"Spain",away:"Iraq",ko:"2026-06-04T19:00:00Z",date:"2026-06-04",md:4},
  // Mexico vs Serbia: est. 20:00 local (Mexico) = ~02:00 UTC June 5 → listed June 4 evening US time
  {id:"f14",home:"Mexico",away:"Serbia",ko:"2026-06-04T22:00:00Z",date:"2026-06-04",md:4},
  // ── June 6 (MD5) ──
  // Germany vs USA: Soldier Field, Chicago — 2:30 PM CT = 18:30 UTC ✓
  {id:"f15",home:"Germany",away:"USA",ko:"2026-06-06T18:30:00Z",date:"2026-06-06",md:5},
  // Portugal vs Chile: est. 21:00 local = 19:00 UTC
  {id:"f16",home:"Portugal",away:"Chile",ko:"2026-06-06T19:00:00Z",date:"2026-06-06",md:5},
  // Argentina vs Honduras: est. 8 PM ET = 00:00 UTC June 7
  {id:"f17",home:"Argentina",away:"Honduras",ko:"2026-06-07T00:00:00Z",date:"2026-06-06",md:5},
  // Brazil vs Egypt: est. 9 PM ET = 01:00 UTC June 7
  {id:"f18",home:"Brazil",away:"Egypt",ko:"2026-06-07T01:00:00Z",date:"2026-06-06",md:5},
  // ── June 9-10 (MD6) ──
  // Argentina vs Iceland: est. 8 PM ET = 00:00 UTC June 10
  {id:"f19",home:"Argentina",away:"Iceland",ko:"2026-06-10T00:00:00Z",date:"2026-06-09",md:6},
  // Portugal vs Nigeria: est. 8 PM local = 19:00 UTC
  {id:"f20",home:"Portugal",away:"Nigeria",ko:"2026-06-10T19:00:00Z",date:"2026-06-10",md:6},
];

const DATES=[...new Map(FIXTURES.map(f=>[f.date,{date:f.date,md:f.md,label:new Date(f.date+"T12:00:00Z").toLocaleDateString([],{weekday:"short",day:"numeric",month:"short"})}])).values()];
const ESPN_DATES=["20260531","20260601","20260602","20260604","20260606","20260609","20260610"];

// ── Flags ─────────────────────────────────────────────────────────────────────
const FLAGS={
  "South Korea":"🇰🇷","Trinidad & Tobago":"🇹🇹","Japan":"🇯🇵","Iceland":"🇮🇸",
  "Ecuador":"🇪🇨","Saudi Arabia":"🇸🇦","Mexico":"🇲🇽","Australia":"🇦🇺",
  "Germany":"🇩🇪","Finland":"🇫🇮","Poland":"🇵🇱","Ukraine":"🇺🇦",
  "USA":"🇺🇸","Senegal":"🇸🇳","Brazil":"🇧🇷","Panama":"🇵🇦",
  "Norway":"🇳🇴","Sweden":"🇸🇪","Turkey":"🇹🇷","North Macedonia":"🇲🇰",
  "Belgium":"🇧🇪","Croatia":"🇭🇷","Netherlands":"🇳🇱","Algeria":"🇩🇿",
  "Spain":"🇪🇸","Iraq":"🇮🇶","Serbia":"🇷🇸","Argentina":"🇦🇷",
  "Honduras":"🇭🇳","Egypt":"🇪🇬","Portugal":"🇵🇹","Chile":"🇨🇱",
  "Nigeria":"🇳🇬",
};
const fl=t=>FLAGS[t]||"🏳";

// ── Team colours ──────────────────────────────────────────────────────────────
const TC={
  "South Korea":"#C60C30","Trinidad & Tobago":"#CE1126","Japan":"#BC002D","Iceland":"#003897",
  "Ecuador":"#FFD100","Saudi Arabia":"#006C35","Mexico":"#006847","Australia":"#FFD700",
  "Germany":"#1a1a1a","Finland":"#003580","Poland":"#DC143C","Ukraine":"#005BBB",
  "USA":"#002868","Senegal":"#00853F","Brazil":"#009C3B","Panama":"#DA121A",
  "Norway":"#EF2B2D","Sweden":"#006AA7","Turkey":"#E30A17","North Macedonia":"#CE2028",
  "Belgium":"#EF3340","Croatia":"#FF3333","Netherlands":"#FF4713","Algeria":"#006233",
  "Spain":"#C60B1E","Iraq":"#007A3D","Serbia":"#C6363C","Argentina":"#74ACDF",
  "Honduras":"#0073CF","Egypt":"#CE1126","Portugal":"#006600","Chile":"#D52B1E","Nigeria":"#008751",
};

// ── Team name resolver for ESPN ───────────────────────────────────────────────
const normT=s=>s.toLowerCase().replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim();
const ALIASES={
  "united states":"USA","us":"USA","usmnt":"USA","korea republic":"South Korea",
  "republic of korea":"South Korea","trinidad and tobago":"Trinidad & Tobago",
  "trinidad tobago":"Trinidad & Tobago","saudi arabia":"Saudi Arabia",
  "north macedonia":"North Macedonia","fyr macedonia":"North Macedonia",
  "türkiye":"Turkey","turkiye":"Turkey","the netherlands":"Netherlands",
};
const resolveTeam=name=>{
  if(!name)return null;
  const k=normT(name);
  if(ALIASES[k])return ALIASES[k];
  // exact match against our team list
  for(const t of Object.keys(FLAGS)){if(normT(t)===k)return t;}
  // partial
  for(const t of Object.keys(FLAGS)){if(k.includes(normT(t))||normT(t).includes(k))return t;}
  return name;
};

// ── Anthropic API + web_search Live Fetch (server-side, no CORS) ──────────────
const buildLiveMap=async(setLog)=>{
  const lm={};const log=[];

  // Only query fixtures that have already kicked off (or within 2h)
  const nowMs=Date.now();
  const active=FIXTURES.filter(f=>nowMs>=new Date(f.ko).getTime()-7_200_000);
  if(!active.length){log.push("⏰ No matches started yet");setLog(log);return lm;}

  // One tight search per unique match date that has active fixtures
  const byDate=[...new Map(active.map(f=>[f.date,f.date])).keys()];
  log.push(`⚡ Searching ${active.length} fixture(s) across ${byDate.length} date(s)…`);setLog([...log]);

  // Fire one Anthropic call per date in parallel — keeps each search focused & fast
  await Promise.all(byDate.map(async date=>{
    const list=active.filter(f=>f.date===date).map(f=>`${f.home} vs ${f.away}`).join("; ");
    const prompt=
      `Search for the CURRENT live scores or final results of these international football friendlies on ${date}: ${list}.\n`+
      `Reply with ONLY a raw JSON array — no markdown, no explanation.\n`+
      `Schema per match: {"home":"","away":"","homeScore":0,"awayScore":0,"status":"NS|1H|HT|2H|FT","clock":""}\n`+
      `status: NS=not started, 1H=first half, HT=half-time, 2H=second half, FT=finished.\n`+
      `Omit matches you cannot find a score for.`;
    try{
      const resp=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514",
          max_tokens:600,
          tools:[{type:"web_search_20250305",name:"web_search"}],
          messages:[{role:"user",content:prompt}]
        })
      });
      if(!resp.ok){log.push(`❌ API HTTP ${resp.status} for ${date}`);return;}
      const data=await resp.json();
      const text=data.content.filter(b=>b.type==="text").map(b=>b.text).join("");
      const m=text.match(/\[[\s\S]*?\]/);
      if(!m){log.push(`⚠ No JSON for ${date}`);return;}
      const results=JSON.parse(m[0]);
      for(const r of results){
        const hN=resolveTeam(r.home),aN=resolveTeam(r.away);
        const fixture=active.find(f=>
          (resolveTeam(f.home)===hN&&resolveTeam(f.away)===aN)||
          (normT(f.home).slice(0,4)===normT(r.home||"").slice(0,4)&&normT(f.away).slice(0,4)===normT(r.away||"").slice(0,4))
        );
        if(!fixture){log.push(`⚠ No fixture: ${r.home} vs ${r.away}`);continue;}
        const hS=parseInt(r.homeScore)||0,aS=parseInt(r.awayScore)||0;
        const st=(r.status||"NS").toUpperCase();
        const isDone=st==="FT"||st==="AET";
        const isLive=["1H","HT","2H","ET"].includes(st);
        const scoreResult=hS>aS?"H":aS>hS?"A":"D";
        lm[fixture.id]={hS,aS,clock:r.clock||st,state:isDone?"post":isLive?"in":"pre",
          done:isDone,result:isDone?scoreResult:null,liveResult:isLive||isDone?scoreResult:null};
        log.push(`${isDone?"🏁":isLive?"⚽":"🕐"} ${fl(fixture.home)} ${fixture.home} ${hS}–${aS} ${fixture.away} ${fl(fixture.away)} [${st}]`);
      }
      log.push(`✅ ${date}: ${results.length} result(s)`);
    }catch(err){log.push(`❌ ${date}: ${err.message}`);}
  }));

  setLog(log);return lm;
};

// ── Scoring ───────────────────────────────────────────────────────────────────
const calcScore=(picks={},lm={})=>FIXTURES.reduce((s,f)=>{
  const p=picks[f.id],r=lm[f.id];if(!p||!r)return s;
  if(r.done&&r.result)return s+(p===r.result?3:0);
  if(r.state==="in"&&r.liveResult)return s+(p===r.liveResult?3:0);
  return s;
},0);

// ── Helpers ───────────────────────────────────────────────────────────────────
const norm=s=>s.trim().toLowerCase();
const valid=s=>s.trim().length>=2&&s.trim().length<=24&&/^[a-zA-Z0-9 _'!-]+$/.test(s.trim());
const fmtDate=d=>new Date(d+"T12:00:00Z").toLocaleDateString([],{weekday:"long",day:"numeric",month:"long"});
const fmtKO=ko=>{try{return new Date(ko).toLocaleTimeString([],{hour:"2-digit",minute:"2-digit",timeZoneName:"short"});}catch{return"KO";}};
const cdown=(ko,now)=>{const ms=new Date(ko)-now;if(ms<=0)return null;const d=Math.floor(ms/864e5),h=Math.floor((ms%864e5)/3.6e6),m=Math.floor((ms%3.6e6)/60000);return d>0?`${d}d ${h}h`:h>0?`${h}h ${m}m`:`${m}m`;};
const isLocked=(ko,now)=>now>=new Date(ko)-600000;
const hasStarted=(ko,now)=>now>=new Date(ko);
const gc=()=>Math.random().toString(36).slice(2,8).toUpperCase();

// ── Storage ───────────────────────────────────────────────────────────────────
const sGet=(k,sh=true)=>window.storage.get(k,sh);
const sSet=(k,v,sh=true)=>window.storage.set(k,v,sh);
const sDel=(k,sh=false)=>window.storage.delete(k,sh);
const SESS="wc26fr_sess1";
const lKey=cd=>`wc26fr_l_${cd}`;
const pKey=un=>`wc26fr_u_${norm(un).replace(/[^a-z0-9]/g,"_")}`;
async function getProfile(un){try{const r=await sGet(pKey(un),false);return r?JSON.parse(r.value):{leagues:[]};}catch{return{leagues:[]};}}
async function addLeague(un,cd,ln){const p=await getProfile(un);p.leagues=p.leagues.filter(l=>l.cd!==cd);p.leagues.unshift({cd,ln,at:Date.now()});await sSet(pKey(un),JSON.stringify(p),false);}

// ── Excel export ──────────────────────────────────────────────────────────────
const exportXLS=(members,lv,picks,nm,tn)=>{
  const sorted=[...members].map(m=>norm(m.username)===norm(nm)?{...m,picks,score:calcScore(picks,lv)}:{...m,score:calcScore(m.picks||{},lv)}).sort((a,b)=>(b.score||0)-(a.score||0));
  const hdr=["Rank","Team","Score",...FIXTURES.map(f=>`MD${f.md}: ${f.home} v ${f.away}`)];
  const rows=sorted.map((m,i)=>[i+1,m.teamName||m.username,m.score||0,...FIXTURES.map(f=>m.picks?.[f.id]??"-")]);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([hdr,...rows]),"Leaderboard");
  XLSX.writeFile(wb,"wc2026_friendlies_picks.xlsx");
};

// ── UI atoms ──────────────────────────────────────────────────────────────────
const Field=({label,children,hint})=>(<div style={{marginBottom:14}}><label style={{display:"block",marginBottom:5,fontSize:10,color:"#9ca3af",fontWeight:700,letterSpacing:1}}>{label}</label>{children}{hint&&<p style={{margin:"4px 0 0",fontSize:10,color:"#6b7280"}}>{hint}</p>}</div>);
const TInput=({value,onChange,placeholder,style={}})=>(<input value={value} onChange={onChange} placeholder={placeholder} style={{background:"#1f2937",border:"1px solid #374151",borderRadius:8,padding:"11px 14px",color:"#f9fafb",width:"100%",fontSize:15,boxSizing:"border-box",fontWeight:600,...style}}/>);
const ErrBox=({msg})=>msg?(<div style={{background:"#450a0a",border:"1px solid #ef4444",borderRadius:8,padding:"10px 14px",color:"#f87171",fontSize:13,marginTop:10}}>⚠ {msg}</div>):null;
const PBtn=({bg,color="#fff",disabled,onClick,children,border})=>(<button onClick={onClick} disabled={disabled} style={{display:"block",background:disabled?"#111827":bg,color:disabled?"#374151":color,border:border||"none",borderRadius:8,padding:"13px",cursor:disabled?"not-allowed":"pointer",fontWeight:700,fontSize:14,width:"100%",textAlign:"center"}}>{children}</button>);

const WcBadge=({size=48})=>(<div style={{width:size,height:size,borderRadius:"50%",background:"linear-gradient(135deg,#C8102E,#007A53)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.42,flexShrink:0}}>⚽</div>);

const Hero=()=>(<div style={{background:"linear-gradient(135deg,#001a00 0%,#003d00 40%,#8B0000 100%)",padding:"20px 16px 16px",position:"relative",overflow:"hidden"}}>
  <svg style={{position:"absolute",inset:0,opacity:.06,width:"100%",height:"100%"}} viewBox="0 0 400 160" preserveAspectRatio="xMidYMid slice">{Array.from({length:18}).map((_,i)=><circle key={i} cx={(i*83)%400} cy={(i*53)%160} r={(i%4)+2} fill="white"/>)}</svg>
  <div style={{position:"relative",display:"flex",alignItems:"center",gap:14,maxWidth:560,margin:"0 auto"}}>
    <WcBadge size={54}/>
    <div>
      <div style={{fontSize:10,color:"#86efac",fontWeight:700,letterSpacing:1.5,marginBottom:2}}>FIFA WORLD CUP 2026 · PRE-TOURNAMENT</div>
      <h1 style={{margin:"0 0 2px",fontSize:20,fontWeight:900}}>International Friendlies</h1>
      <p style={{margin:0,fontSize:11,opacity:.5}}>Prediction League · 3 pts per correct result · Live scoring</p>
    </div>
  </div>
</div>);

export default function App(){
  const [view,setView]=useState("loading");
  const [tab,setTab]=useState("sched");
  const [nm,setNm]=useState(""),[tn,setTn]=useState(""),[cd,setCd]=useState(""),[ln,setLn]=useState("");
  const [picks,setPicks]=useState({});
  const [members,setMem]=useState([]);
  const [uLeagues,setUL]=useState([]);
  const [lv,setLv]=useState({});
  const [apiLog,setApiLog]=useState([]);
  const [apiStatus,setApiStatus]=useState("idle");
  const [showLog,setShowLog]=useState(false);
  const [now,setNow]=useState(Date.now());
  const [ld,setLd]=useState(false),[er,setEr]=useState(""),[ ss,setSS]=useState("idle");
  const [dateFilter,setDF]=useState("all");
  const [fUN,setFUN]=useState(""),[ fTN,setFTN]=useState(""),[ fJC,setFJC]=useState("");
  const saveRef=useRef(null),savingRef=useRef(false),initRef=useRef(true);
  const picksRef=useRef(picks);picksRef.current=picks;
  const lvRef=useRef(lv);lvRef.current=lv;

  useEffect(()=>{
    (async()=>{
      try{
        const r=await sGet(SESS,false);if(!r)throw 0;
        const s=JSON.parse(r.value);if(!s.nm||!s.cd)throw 0;
        const lg=await sGet(lKey(s.cd));if(!lg)throw 0;
        const league=JSON.parse(lg.value);
        const me=league.members.find(m=>norm(m.username)===norm(s.nm));
        if(me?.picks)setPicks(me.picks);
        setNm(s.nm);setTn(s.tn||me?.teamName||"");setCd(s.cd);setLn(league.name);setMem(league.members);
        setView("app");
      }catch{setView("home");}
    })();
  },[]);

  const fetchLive=useCallback(async()=>{
    setApiStatus("loading");
    const map=await buildLiveMap(setApiLog);
    setLv(map);
    setApiStatus(Object.keys(map).length>0?"ok":"empty");
    if(Object.keys(map).length>0)setMem(prev=>prev.map(m=>({...m,score:calcScore(m.picks||{},map)})));
  },[]);

  useEffect(()=>{
    if(view!=="app")return;
    fetchLive();
    const id=setInterval(fetchLive,30000);
    return()=>clearInterval(id);
  },[view,fetchLive]);

  useEffect(()=>{const id=setInterval(()=>setNow(Date.now()),15000);return()=>clearInterval(id);},[]);

  const saveSess=async(u,t,c)=>{try{await sSet(SESS,JSON.stringify({nm:u,tn:t,cd:c}),false);}catch{}};
  const reset=()=>{setFUN("");setFTN("");setFJC("");setEr("");};
  const logout=async()=>{try{await sDel(SESS,false);}catch{}setNm("");setTn("");setCd("");setLn("");setPicks({});setMem([]);setLv({});setSS("idle");reset();setView("home");};

  const enter=async(u,t,code,league)=>{
    initRef.current=true;
    const me=league.members.find(m=>norm(m.username)===norm(u));
    if(me?.picks)setPicks(me.picks);
    setNm(u);setTn(t||me?.teamName||"");setCd(code);setLn(league.name);setMem(league.members);
    await saveSess(u,t||me?.teamName||"",code);
    try{await addLeague(u,code,league.name);}catch{}
    reset();setView("app");
    setTimeout(()=>{initRef.current=false;},1500);
  };

const createLeague = async () => {
    const u = fUN.trim(), t = fTN.trim();
    if (!valid(u)) return setEr("Username 2–24 chars.");
    if (!valid(t)) return setEr("Team name 2–24 chars.");
    setLd(true); setEr("");
    try {
      const c = gc(); // Generates your 6-character code
      const league = { 
        name: "WC2026 Friendlies", 
        creator: u, 
        code: c, 
        createdAt: Date.now(), 
        members: [{ username: u, teamName: t, picks: {}, score: 0, at: Date.now() }] 
      };
      
      // Saved directly to your browser's persistent local storage!
      localStorage.setItem(`lg_${c}`, JSON.stringify(league));
      
      // Simulate your original "enter" function logic to push state to the UI
      await enter(u, t, c, league);
    } catch (e) {
      setEr(`Failed: ${e?.message || "error"}`);
    }
    setLd(false);
  };

  const joinLeague = async () => {
    const u = fUN.trim(), t = fTN.trim(), c = fJC.trim().toUpperCase();
    if (!valid(u)) return setEr("Username 2–24 chars.");
    if (!valid(t)) return setEr("Team name 2–24 chars.");
    if (c.length < 6) return setEr("Enter the 6-character code.");
    setLd(true); setEr("");
    try {
      // Look for the league inside local storage instead of sGet
      const r = localStorage.getItem(`lg_${c}`);
      if (!r) {
        setEr(`No league "${c}".`);
        setLd(false);
        return;
      }
      
      const league = JSON.parse(r);
      const byUN = league.members.find(m => norm(m.username) === norm(u));
      const byTN = league.members.find(m => norm(m.teamName || "") === norm(t));
      
      if (byUN && byTN && norm(byUN.username) === norm(byTN.username)) {
        await enter(byUN.username, byUN.teamName, c, league);
        setLd(false);
        return;
      }
      if (byUN) {
        setEr(`Username "${u}" taken.`);
        setLd(false);
        return;
      }
      if (byTN) {
        setEr(`Team name "${t}" taken.`);
        setLd(false);
        return;
      }
      
      league.members.push({ username: u, teamName: t, picks: {}, score: 0, at: Date.now() });
      
      // Update local storage instead of sSet
      localStorage.setItem(`lg_${c}`, JSON.stringify(league));
      await enter(u, t, c, league);
    } catch (e) {
      setEr(e.message || "Error");
    }
    setLd(false);
  };
  const login=async()=>{
    const u=fUN.trim();if(!u)return setEr("Enter your username.");
    setLd(true);setEr("");
    try{
      const profile=await getProfile(u);
      if(!profile.leagues.length){setEr(`No account for "${u}".`);setLd(false);return;}
      if(profile.leagues.length===1){
        const r=await sGet(lKey(profile.leagues[0].cd));
        if(!r){setEr("League not found.");setLd(false);return;}
        const league=JSON.parse(r.value);
        const me=league.members.find(m=>norm(m.username)===norm(u));
        await enter(u,me?.teamName||"",profile.leagues[0].cd,league);
      }else{setUL(profile.leagues);setView("pickLeague");}
    }catch{setEr("Could not find account.");}
    setLd(false);
  };

  const autoSave=useCallback(async(newPicks,code,username,teamName)=>{
    if(!code||!username||savingRef.current||initRef.current)return;
    savingRef.current=true;setSS("saving");
    try{
      const r=await sGet(lKey(code));if(!r)throw 0;
      const lg=JSON.parse(r.value);
      const ex=lg.members.find(m=>norm(m.username)===norm(username))||{};
      lg.members=lg.members.filter(m=>norm(m.username)!==norm(username));
      lg.members.push({...ex,username,teamName,picks:newPicks,score:calcScore(newPicks,lvRef.current),at:Date.now()});
      await sSet(lKey(code),JSON.stringify(lg));
      setSS("saved");
      setMem(lg.members.map(m=>({...m,score:calcScore(m.picks||{},lvRef.current)})));
      setTimeout(()=>setSS("idle"),2500);
    }catch{setSS("error");}
    savingRef.current=false;
  },[]);

  useEffect(()=>{
    if(view!=="app"||!cd||!nm)return;
    clearTimeout(saveRef.current);
    saveRef.current=setTimeout(()=>autoSave(picksRef.current,cd,nm,tn),800);
    return()=>clearTimeout(saveRef.current);
  },[picks,cd,nm,tn,view]);

  const doPick=(id,v)=>{const f=FIXTURES.find(x=>x.id===id);if(f&&isLocked(f.ko,now))return;setPicks(p=>({...p,[id]:v}));};

  const picked=Object.keys(picks).length;
  const pct=Math.round((picked/FIXTURES.length)*100);
  const myScore=calcScore(picks,lv);
  const doneCount=FIXTURES.filter(f=>lv[f.id]?.done).length;
  const liveCount=FIXTURES.filter(f=>lv[f.id]?.state==="in").length;

  const SaveBadge=ss==="saving"?<span style={{fontSize:10,color:"#fbbf24",fontWeight:700}}>💾 Saving…</span>:ss==="saved"?<span style={{fontSize:10,color:"#4ade80",fontWeight:700}}>✓ Saved</span>:ss==="error"?<span style={{fontSize:10,color:"#f87171",fontWeight:700}}>⚠ Error</span>:null;
  const pg={minHeight:"100vh",background:"#060d1e",color:"#f9fafb",fontFamily:"system-ui,sans-serif"};
  const card={background:"#111827",borderRadius:12,padding:16,marginBottom:10};

  if(view==="loading")return(<div style={{...pg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14}}><WcBadge size={56}/><style>{`@keyframes sl{0%{transform:translateX(-100%)}100%{transform:translateX(250%)}}`}</style><div style={{width:60,height:3,background:"#1f2937",borderRadius:2,overflow:"hidden"}}><div style={{width:"45%",height:"100%",background:G,borderRadius:2,animation:"sl 1s infinite"}}/></div><div style={{fontSize:11,color:"#6b7280"}}>WC 2026 Friendlies</div></div>);

  const back=(to)=>(<button onClick={()=>{setEr("");setView(to);}} style={{background:"transparent",border:"none",color:"#6b7280",fontSize:13,cursor:"pointer",padding:"8px 0",display:"block",marginBottom:8}}>← Back</button>);

  if(view==="home")return(<div style={pg}><Hero/><div style={{padding:20,maxWidth:400,margin:"0 auto",boxSizing:"border-box",display:"flex",flexDirection:"column",gap:6}}>
    <div style={{background:"#111827",borderRadius:10,padding:"10px 14px",marginBottom:6,fontSize:12,color:"#86efac",border:"1px solid #007A5344",textAlign:"center"}}>⚽ {FIXTURES.length} matches · May 31 – June 10 · Live scoring</div>
    <PBtn bg={G} onClick={()=>{reset();setView("create");}}>🚀 Create a League</PBtn>
    <div style={{textAlign:"center",color:"#4b5563",margin:"2px 0",fontSize:12}}>Have a code from a friend?</div>
    <PBtn bg={R} onClick={()=>{reset();setView("join");}}>🔗 Join a League</PBtn>
    <div style={{textAlign:"center",color:"#4b5563",margin:"2px 0",fontSize:12}}>Already joined?</div>
    <PBtn bg="transparent" border="1px solid #374151" color="#60a5fa" onClick={()=>{reset();setView("login");}}>👤 Log In</PBtn>
  </div></div>);

  const authShell=(title,color,sub,body)=>(<div style={pg}><Hero/><div style={{padding:20,maxWidth:400,margin:"0 auto",boxSizing:"border-box"}}><div style={{...card,padding:20}}><p style={{margin:"0 0 4px",fontWeight:800,fontSize:16,color}}>{title}</p>{sub&&<p style={{margin:"0 0 18px",fontSize:12,color:"#9ca3af"}}>{sub}</p>}{body}</div></div></div>);

  if(view==="create")return authShell("🚀 Create a League",G,"Username is private · Team name shows on leaderboard",<>{back("home")}<Field label="USERNAME (PRIVATE)"><TInput value={fUN} onChange={e=>{setFUN(e.target.value);setEr("");}} placeholder="e.g. godfrey99"/></Field><Field label="TEAM NAME (PUBLIC)"><TInput value={fTN} onChange={e=>{setFTN(e.target.value);setEr("");}} placeholder="e.g. The Samba Kings"/></Field><PBtn bg={G} disabled={ld} onClick={createLeague}>{ld?"Creating…":"Create League & Get Code"}</PBtn><ErrBox msg={er}/></>);
  if(view==="join")return authShell("🔗 Join a League","#34d399","Username stays private",<>{back("home")}<Field label="USERNAME (PRIVATE)"><TInput value={fUN} onChange={e=>{setFUN(e.target.value);setEr("");}} placeholder="e.g. godfrey99"/></Field><Field label="TEAM NAME (PUBLIC)"><TInput value={fTN} onChange={e=>{setFTN(e.target.value);setEr("");}} placeholder="e.g. The Samba Kings"/></Field><Field label="LEAGUE CODE"><TInput value={fJC} onChange={e=>{setFJC(e.target.value.toUpperCase());setEr("");}} placeholder="ABC123" style={{textTransform:"uppercase",letterSpacing:6,fontWeight:900,fontSize:24,textAlign:"center"}}/></Field><PBtn bg={R} disabled={ld} onClick={joinLeague}>{ld?"Joining…":"Join League"}</PBtn><ErrBox msg={er}/></>);
  if(view==="login")return authShell("👤 Log In","#86efac",null,<>{back("home")}<Field label="YOUR USERNAME"><TInput value={fUN} onChange={e=>{setFUN(e.target.value);setEr("");}} placeholder="e.g. godfrey99"/></Field><PBtn bg={G} disabled={ld} onClick={login}>{ld?"Looking up…":"Log In"}</PBtn><ErrBox msg={er}/></>);
  if(view==="pickLeague")return(<div style={pg}><Hero/><div style={{padding:20,maxWidth:400,margin:"0 auto",boxSizing:"border-box"}}>{back("login")}<p style={{margin:"0 0 14px",fontWeight:800}}>Choose a league:</p>{uLeagues.map(l=>(<button key={l.cd} onClick={()=>{setLd(true);sGet(lKey(l.cd)).then(r=>{const league=JSON.parse(r.value);const me=league.members.find(m=>norm(m.username)===norm(fUN.trim()));enter(me?.username||fUN.trim(),me?.teamName||"",l.cd,league);}).catch(()=>setEr("Could not load.")).finally(()=>setLd(false));}} disabled={ld} style={{...card,width:"100%",textAlign:"left",cursor:"pointer",border:`1px solid ${G}`}}><div style={{fontWeight:700,fontSize:14,color:G,marginBottom:4}}>{l.ln}</div><span style={{fontFamily:"monospace",fontSize:20,letterSpacing:4,fontWeight:900,color:"#60a5fa"}}>{l.cd}</span></button>))}<ErrBox msg={er}/></div></div>);

  // ── APP ───────────────────────────────────────────────────────────────────
  const visibleFixtures=dateFilter==="all"?FIXTURES:FIXTURES.filter(f=>f.date===dateFilter);
  const groupedByDate=DATES.map(d=>({...d,fixtures:visibleFixtures.filter(f=>f.date===d.date)})).filter(g=>g.fixtures.length>0);

  const ApiDot=()=>{
    const[c,label]=apiStatus==="loading"?["#fbbf24","syncing…"]:apiStatus==="ok"?["#4ade80",`live · ${doneCount} done${liveCount>0?`, ${liveCount} live`:""}`]:["#ef4444","no data"];
    return(<button onClick={()=>setShowLog(p=>!p)} style={{display:"flex",alignItems:"center",gap:5,background:"transparent",border:"1px solid #1f2937",borderRadius:20,padding:"3px 9px",cursor:"pointer",color:c,fontSize:10,fontWeight:700}}><span style={{width:6,height:6,borderRadius:"50%",background:c,display:"inline-block",boxShadow:liveCount>0?`0 0 6px ${c}`:"none"}}/>ESPN {label}</button>);
  };

  // ── Fixtures Tab ──────────────────────────────────────────────────────────
  const SchedTab=()=>(
    <div style={{padding:12,maxWidth:580,margin:"0 auto",paddingBottom:80}}>
      <style>{`@keyframes livePulse{0%,100%{opacity:1}50%{opacity:.3}}`}</style>
      <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
        {[{date:"all",md:null,label:"All"},...DATES].map(d=>{const sel=dateFilter===d.date;return(<button key={d.date} onClick={()=>setDF(d.date)} style={{padding:"5px 11px",borderRadius:20,fontSize:11,fontWeight:700,cursor:"pointer",background:sel?G:"#1f2937",color:sel?"white":"#9ca3af",border:`1px solid ${sel?G:"#374151"}`}}>{d.date==="all"?"All Matches":`MD${d.md} · ${d.label}`}</button>);})}</div>
      {ss!=="idle"&&<div style={{textAlign:"center",marginBottom:8}}>{SaveBadge}</div>}
      {groupedByDate.map(grp=>(
        <div key={grp.date}>
          <div style={{display:"flex",alignItems:"center",gap:10,margin:"6px 0 10px"}}>
            <div style={{height:2,flex:1,background:`linear-gradient(90deg,${G},transparent)`}}/>
            <div style={{fontSize:11,fontWeight:800,color:G,whiteSpace:"nowrap"}}>Match Day {grp.md} · {fmtDate(grp.date)}</div>
            <div style={{height:2,flex:1,background:`linear-gradient(90deg,transparent,${G})`}}/>
          </div>
          {grp.fixtures.map(f=>{
            const p=picks[f.id],lm=lv[f.id],fL=isLocked(f.ko,now);
            const isLive=lm?.state==="in",isDone=lm?.done;
            const activeResult=isDone?lm.result:isLive?lm.liveResult:null;
            const earning=p&&activeResult!=null?(p===activeResult?3:0):null;
            const cd2=cdown(f.ko,now);
            return(
              <div key={f.id} style={{...card,padding:0,overflow:"hidden",marginBottom:12,border:`1px solid ${isLive?"#ef4444":isDone&&earning===3?"#22c55e55":isDone&&earning===0&&p?"#ef444433":p?G+"55":"#1f2937"}`,boxShadow:isLive?"0 0 14px #ef444425":"none"}}>

                {/* Status bar */}
                <div style={{padding:"6px 12px",background:"#0d1520",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:4}}>
                  <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                    <span style={{fontSize:9,background:G,color:"white",borderRadius:4,padding:"1px 6px",fontWeight:800}}>MD{f.md}</span>
                    {isLive&&<span style={{fontSize:9,background:R,color:"white",borderRadius:4,padding:"1px 7px",fontWeight:800,animation:"livePulse 1.2s infinite"}}>⚽ LIVE {lm.clock}</span>}
                    {isDone&&<span style={{fontSize:9,background:earning===3?"#14532d":earning===0&&p?"#450a0a":"#1f2937",color:earning===3?"#4ade80":earning===0&&p?"#f87171":"#9ca3af",borderRadius:4,padding:"1px 6px",fontWeight:700}}>FT {lm.hS}–{lm.aS} · {earning===3?"+3 pts ✓":earning===0&&p?"0 pts ✗":"—"}</span>}
                    {isLive&&earning!=null&&<span style={{fontSize:9,background:earning===3?"#1a3a1a":"#2a1a1a",color:earning===3?"#86efac":"#f87171",borderRadius:4,padding:"1px 6px",fontWeight:700,border:`1px solid ${earning===3?"#22c55e55":"#ef444455"}`}}>⚡ {earning===3?"+3 live":"0 live"}</span>}
                  </div>
                  <div style={{fontSize:9,color:fL&&!isDone&&!isLive?"#ef4444":"#6b7280",fontWeight:fL?700:400}}>
                    {isDone||isLive?"":fL?"🔒 Locked":cd2?`🔓 ${cd2}`:`KO ${fmtKO(f.ko)}`}
                  </div>
                </div>

                {/* Live scoreboard */}
                {lm&&(isLive||isDone)&&(
                  <div style={{padding:"10px 14px",background:isLive?"#0a0505":"#040a14",display:"flex",alignItems:"center",justifyContent:"center",gap:12,borderBottom:"1px solid #111"}}>
                    <div style={{textAlign:"right",flex:1}}>
                      <div style={{fontSize:22,marginBottom:2}}>{fl(f.home)}</div>
                      <div style={{fontWeight:800,fontSize:12,color:lm.hS>lm.aS?"#4ade80":"#e5e7eb"}}>{f.home}</div>
                    </div>
                    <div style={{textAlign:"center",minWidth:90}}>
                      <div style={{fontSize:36,fontWeight:900,color:isLive?R:"#e5e7eb",fontFamily:"monospace",letterSpacing:4,lineHeight:1}}>{lm.hS} – {lm.aS}</div>
                      <div style={{fontSize:8,color:isLive?"#f87171":"#6b7280",marginTop:3}}>{isLive?`⏱ ${lm.clock}`:"FULL TIME"}</div>
                    </div>
                    <div style={{textAlign:"left",flex:1}}>
                      <div style={{fontSize:22,marginBottom:2}}>{fl(f.away)}</div>
                      <div style={{fontWeight:800,fontSize:12,color:lm.aS>lm.hS?"#4ade80":"#e5e7eb"}}>{f.away}</div>
                    </div>
                  </div>
                )}

                {/* Pregame matchup header */}
                {!(isLive||isDone)&&(
                  <div style={{padding:"12px 14px 4px",display:"flex",alignItems:"center",justifyContent:"center",gap:10}}>
                    <div style={{textAlign:"right",flex:1}}>
                      <div style={{fontSize:20}}>{fl(f.home)}</div>
                      <div style={{fontWeight:800,fontSize:12}}>{f.home}</div>
                    </div>
                    <div style={{fontSize:11,color:"#374151",background:"#0d1520",padding:"4px 10px",borderRadius:6,fontWeight:700}}>vs</div>
                    <div style={{textAlign:"left",flex:1}}>
                      <div style={{fontSize:20}}>{fl(f.away)}</div>
                      <div style={{fontWeight:800,fontSize:12}}>{f.away}</div>
                    </div>
                  </div>
                )}

                {/* Pick buttons */}
                <div style={{padding:"10px 12px 12px"}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
                    {[["H",`${f.home} Win`,TC[f.home]||G],["D","Draw","#f59e0b"],["A",`${f.away} Win`,TC[f.away]||R]].map(([v,lbl,col])=>{
                      const sel=p===v;
                      const finalOk=isDone&&sel&&v===lm.result,finalBad=isDone&&sel&&v!==lm.result;
                      const liveOk=isLive&&sel&&v===lm.liveResult,liveBad=isLive&&sel&&v!==lm.liveResult;
                      const missed=isDone&&!sel&&lm.result===v,liveMissed=isLive&&!sel&&lm.liveResult===v;
                      return(<button key={v} onClick={()=>doPick(f.id,v)} disabled={fL}
                        style={{padding:"10px 4px",borderRadius:8,fontSize:11,fontWeight:700,cursor:fL?"not-allowed":"pointer",lineHeight:1.3,
                          border:`2px solid ${finalOk?"#22c55e":finalBad?"#ef4444":missed?"#374151":liveOk?"#86efac":liveBad?"#f87171":liveMissed?"#374151":sel?col:"#1f2937"}`,
                          background:finalOk?"#14532d":finalBad?"#450a0a":missed?"#111827":liveOk?"#0d2a0d":liveBad?"#2a0a0a":sel?col+"25":"transparent",
                          color:finalOk?"#4ade80":finalBad?"#f87171":missed?"#6b7280":liveOk?"#86efac":liveBad?"#f87171":liveMissed?"#6b7280":sel?col:"#6b7280",
                          opacity:fL&&!sel?.4:1}}>
                        {v==="H"?`${fl(f.home)} `:v==="A"?`${fl(f.away)} `:""}{finalOk?"✓ ":finalBad?"✗ ":liveOk?"⚡ ":liveBad?"⚡ ":""}{lbl}{missed?" ✓":liveMissed?" ⚡":""}
                      </button>);
                    })}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );

  // ── All Picks Tab ─────────────────────────────────────────────────────────
  const TodayTab=()=>{
    const everyone=[{username:nm,teamName:tn,picks,isMe:true},...members.filter(m=>norm(m.username)!==norm(nm)).map(m=>({...m,isMe:false}))];
    return(
      <div style={{padding:12,maxWidth:620,margin:"0 auto",paddingBottom:80}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:12,color:"#9ca3af"}}>{everyone.length} players · {doneCount}/{FIXTURES.length} done{liveCount>0?` · ${liveCount} live`:""}</div>
          <button onClick={()=>exportXLS(members,lv,picks,nm,tn)} style={{background:"#14532d",border:"1px solid #16a34a",borderRadius:8,padding:"7px 12px",color:"#86efac",fontSize:11,fontWeight:700,cursor:"pointer"}}>📊 Export</button>
        </div>
        {groupedByDate.map(grp=>(
          <div key={grp.date}>
            <div style={{display:"flex",alignItems:"center",gap:8,margin:"6px 0 8px"}}>
              <div style={{height:2,flex:1,background:`linear-gradient(90deg,${G},transparent)`}}/>
              <span style={{fontSize:10,fontWeight:800,color:G,whiteSpace:"nowrap"}}>MD{grp.md} · {fmtDate(grp.date)}</span>
              <div style={{height:2,flex:1,background:`linear-gradient(90deg,transparent,${G})`}}/>
            </div>
            {grp.fixtures.map(f=>{
              const lm=lv[f.id],started=hasStarted(f.ko,now),isLive=lm?.state==="in";
              return(
                <div key={f.id} style={{...card,padding:0,overflow:"hidden",marginBottom:10,border:isLive?"1px solid #ef444466":"1px solid #1f2937"}}>
                  <div style={{padding:"8px 12px",background:"#0d1520",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div>
                      <div style={{fontWeight:800,fontSize:12}}>{fl(f.home)} {f.home} <span style={{color:"#374151"}}>vs</span> {f.away} {fl(f.away)}</div>
                      <div style={{fontSize:9,color:"#6b7280",marginTop:1}}>{fmtKO(f.ko)}</div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      {isLive&&<div style={{fontSize:10,background:R,color:"white",borderRadius:5,padding:"2px 7px",fontWeight:800}}>⚽ {lm.hS}–{lm.aS} {lm.clock}</div>}
                      {lm?.done&&<div style={{fontSize:12,fontWeight:800,color:"#e5e7eb"}}>FT {lm.hS}–{lm.aS}</div>}
                      {!started&&<div style={{fontSize:9,color:"#6b7280"}}>{cdown(f.ko,now)?`in ${cdown(f.ko,now)}`:fmtKO(f.ko)}</div>}
                    </div>
                  </div>
                  {!started?(
                    <div>
                      <div style={{padding:"9px 12px",borderBottom:"1px solid #0d1520",display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:6,height:6,borderRadius:"50%",background:G}}/>
                        <span style={{fontSize:12,fontWeight:700,color:"#86efac",flex:1}}>{tn} (you)</span>
                        {picks[f.id]?<span style={{fontSize:11,fontWeight:700,color:G,background:"#1f2937",borderRadius:5,padding:"2px 9px"}}>{picks[f.id]==="H"?`${fl(f.home)} ${f.home} Win`:picks[f.id]==="D"?"Draw":`${fl(f.away)} ${f.away} Win`}</span>:<span style={{fontSize:11,color:"#6b7280"}}>No pick</span>}
                      </div>
                      {everyone.filter(m=>!m.isMe).length>0&&<div style={{padding:"9px 12px",color:"#6b7280",fontSize:11}}>🔒 {everyone.filter(m=>!m.isMe).length} others — revealed at KO</div>}
                    </div>
                  ):(
                    <>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",background:"#0a0f1a",padding:"4px 12px",borderBottom:"1px solid #111827"}}>
                        <div style={{fontSize:8,color:"#6b7280",fontWeight:700}}>TEAM</div>
                        {[["H",`${fl(f.home)} H`],["D","Draw"],["A",`${fl(f.away)} A`]].map(([v,l])=>{
                          const active=(lm?.done&&lm.result===v)||(isLive&&lm?.liveResult===v);
                          return(<div key={v} style={{fontSize:8,color:active?G:"#6b7280",fontWeight:700,textAlign:"center"}}>{active&&(isLive?"⚡ ":"✓ ")}{l}</div>);
                        })}
                      </div>
                      {everyone.map(m=>{
                        const mp=m.picks?.[f.id];
                        const finalOk=lm?.done&&mp&&mp===lm?.result,finalBad=lm?.done&&mp&&mp!==lm?.result;
                        const liveOk=isLive&&mp&&mp===lm?.liveResult,liveBad=isLive&&mp&&mp!==lm?.liveResult;
                        return(
                          <div key={m.username} style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",padding:"6px 12px",borderBottom:"1px solid #0a0f1a",alignItems:"center",background:finalOk?"#0a200f":finalBad?"#1a0505":liveOk?"#051a05":liveBad?"#150505":m.isMe?"#0d1a30":"transparent"}}>
                            <div style={{display:"flex",alignItems:"center",gap:6,minWidth:0}}>
                              <div style={{width:6,height:6,borderRadius:"50%",flexShrink:0,background:finalOk?"#22c55e":finalBad?"#ef4444":liveOk?"#86efac":liveBad?"#f87171":m.isMe?G:"#374151"}}/>
                              <span style={{fontSize:11,fontWeight:m.isMe?700:400,color:m.isMe?"#86efac":"#9ca3af",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.teamName||m.username}{m.isMe?" ★":""}</span>
                            </div>
                            {["H","D","A"].map(v=>(<div key={v} style={{textAlign:"center"}}>
                              {mp===v&&<span style={{fontSize:10,fontWeight:800,display:"inline-flex",alignItems:"center",justifyContent:"center",width:20,height:20,borderRadius:"50%",color:"white",background:finalOk?"#22c55e":finalBad?"#ef4444":liveOk?"#16803c":liveBad?"#b91c1c":v==="H"?TC[f.home]||G:v==="D"?"#f59e0b":TC[f.away]||R}}>{finalOk?"✓":finalBad?"✗":liveOk?"⚡":liveBad?"⚡":"●"}</span>}
                            </div>))}
                          </div>
                        );
                      })}
                    </>
                  )}
                </div>
              );
            })}
          </div>
        ))}
      </div>
    );
  };

  // ── Leaderboard Tab ───────────────────────────────────────────────────────
  const BoardTab=()=>{
    const sorted=[...members]
      .map(m=>norm(m.username)===norm(nm)?{...m,picks,score:calcScore(picks,lv)}:{...m,score:calcScore(m.picks||{},lv)})
      .sort((a,b)=>(b.score||0)-(a.score||0));
    const maxPts=FIXTURES.length*3;
    const correctFinal=FIXTURES.filter(f=>lv[f.id]?.done&&picks[f.id]&&picks[f.id]===lv[f.id]?.result).length;
    const correctLive=FIXTURES.filter(f=>lv[f.id]?.state==="in"&&picks[f.id]&&picks[f.id]===lv[f.id]?.liveResult).length;
    return(
      <div style={{padding:12,maxWidth:560,margin:"0 auto",paddingBottom:80}}>
        <div style={{...card,background:"linear-gradient(135deg,#001a00,#003d00)",border:`1px solid ${G}55`,textAlign:"center",padding:"14px 20px",marginBottom:12}}>
          <div style={{display:"flex",justifyContent:"center",alignItems:"center",gap:10,marginBottom:4}}>
            <WcBadge size={32}/>
            <div>
              <p style={{margin:0,fontSize:9,color:"#86efac",fontWeight:800,letterSpacing:2}}>INVITE CODE</p>
              <p style={{margin:0,fontFamily:"monospace",fontSize:32,fontWeight:900,letterSpacing:10,color:"#60a5fa"}}>{cd}</p>
            </div>
          </div>
          <p style={{margin:0,fontSize:10,color:"#6b7280"}}>Share this code for friends to join</p>
        </div>

        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10,flexWrap:"wrap"}}>
          <ApiDot/>
          <button onClick={fetchLive} disabled={apiStatus==="loading"} style={{fontSize:10,fontWeight:700,background:"#1f2937",border:"none",borderRadius:20,padding:"4px 12px",color:"#9ca3af",cursor:"pointer"}}>↻ Refresh</button>
          <span style={{fontSize:9,color:"#374151"}}>ESPN · 30s auto</span>
          <button onClick={()=>exportXLS(members,lv,picks,nm,tn)} style={{fontSize:10,fontWeight:700,background:"#14532d",border:"1px solid #16a34a",borderRadius:20,padding:"4px 12px",color:"#86efac",cursor:"pointer"}}>📊 Export</button>
        </div>

        {showLog&&apiLog.length>0&&<div style={{...card,padding:"10px 12px",marginBottom:10,background:"#0a0f1a",maxHeight:180,overflowY:"auto"}}>{apiLog.map((l,i)=><div key={i} style={{fontSize:9,fontFamily:"monospace",lineHeight:1.8,color:l.startsWith("✅")||l.startsWith("⚽")?"#4ade80":l.startsWith("❌")||l.startsWith("⚠")?"#f87171":"#6b7280"}}>{l}</div>)}</div>}

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:12}}>
          {[[doneCount,`/${FIXTURES.length}`,"Final",R],[liveCount,"","Live","#ef4444"],[myScore,"pts","My Score","#22c55e"],[sorted.length,"","Players","#60a5fa"]].map(([v,u,l,c])=>(
            <div key={l} style={{background:"#111827",borderRadius:10,padding:"10px 6px",textAlign:"center",border:`1px solid ${l==="Live"&&liveCount>0?"#ef444444":"#1f2937"}`,boxShadow:l==="Live"&&liveCount>0?"0 0 8px #ef444422":"none"}}>
              <div style={{fontSize:22,fontWeight:900,color:c,lineHeight:1}}>{v}<span style={{fontSize:9,color:"#6b7280",fontWeight:400}}>{u}</span></div>
              <div style={{fontSize:9,color:"#6b7280",marginTop:2}}>{l}</div>
            </div>
          ))}
        </div>

        <div style={{...card,padding:"12px",marginBottom:14,border:`1px solid ${G}55`}}>
          <div style={{fontSize:9,color:"#86efac",fontWeight:800,letterSpacing:1,marginBottom:6}}>YOUR PERFORMANCE</div>
          <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
            <div><div style={{fontSize:26,fontWeight:900,color:"#22c55e",lineHeight:1}}>{myScore}</div><div style={{fontSize:9,color:"#6b7280"}}>total pts</div></div>
            <div><div style={{fontSize:26,fontWeight:900,color:"#4ade80",lineHeight:1}}>{correctFinal}</div><div style={{fontSize:9,color:"#6b7280"}}>final ✓</div></div>
            {correctLive>0&&<div><div style={{fontSize:26,fontWeight:900,color:"#f87171",lineHeight:1}}>{correctLive}</div><div style={{fontSize:9,color:"#6b7280"}}>live ⚡</div></div>}
            <div style={{flex:1,minWidth:80,display:"flex",flexDirection:"column",justifyContent:"center"}}>
              <div style={{height:6,background:"#1f2937",borderRadius:3,marginBottom:4}}><div style={{height:6,borderRadius:3,background:G,width:`${maxPts>0?Math.round((myScore/maxPts)*100):0}%`,transition:"width .5s"}}/></div>
              <div style={{fontSize:9,color:"#6b7280"}}>{myScore}/{maxPts} pts · {picked}/{FIXTURES.length} picks</div>
            </div>
          </div>
          {liveCount>0&&<div style={{marginTop:8,padding:"6px 8px",background:"#1a0a0a",borderRadius:6,border:"1px solid #ef444433",fontSize:10,color:"#fca5a5"}}>⚡ {liveCount} match{liveCount>1?"es":""} in play — standings update with every goal</div>}
        </div>

        <p style={{margin:"4px 0 8px",fontSize:9,color:"#374151",fontWeight:800,letterSpacing:1.5}}>STANDINGS · {sorted.length} PLAYER{sorted.length!==1?"S":""}</p>
        {sorted.map((m,i)=>{
          const isMe=norm(m.username)===norm(nm),sc=m.score||0;
          const medal=i===0?"🥇":i===1?"🥈":i===2?"🥉":null;
          const mFin=FIXTURES.filter(f=>lv[f.id]?.done&&m.picks?.[f.id]&&m.picks[f.id]===lv[f.id]?.result).length;
          const mLive=FIXTURES.filter(f=>lv[f.id]?.state==="in"&&m.picks?.[f.id]&&m.picks[f.id]===lv[f.id]?.liveResult).length;
          return(
            <div key={m.username} style={{...card,display:"flex",alignItems:"center",gap:10,padding:"12px",marginBottom:8,border:isMe?`2px solid ${G}`:"1px solid #1f2937",background:isMe?"linear-gradient(90deg,#001a00,#002200)":undefined}}>
              <div style={{fontSize:medal?20:13,width:26,textAlign:"center",flexShrink:0,color:"#6b7280",fontWeight:700}}>{medal||i+1}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
                  <span style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.teamName||m.username}</span>
                  {isMe&&<span style={{fontSize:8,background:G,color:"white",borderRadius:5,padding:"1px 5px",fontWeight:800,flexShrink:0}}>YOU</span>}
                </div>
                <div style={{height:3,background:"#1f2937",borderRadius:2}}><div style={{height:3,borderRadius:2,background:isMe?G:R,width:`${maxPts>0?Math.round((sc/maxPts)*100):0}%`,transition:"width .5s"}}/></div>
                <div style={{fontSize:9,color:"#6b7280",marginTop:2}}>{mFin}/{doneCount} final{mLive>0?` · ${mLive} live ⚡`:""} · {Object.keys(m.picks||{}).length}/{FIXTURES.length} picks</div>
              </div>
              <div style={{textAlign:"right",flexShrink:0}}>
                <div style={{fontWeight:900,fontSize:28,color:isMe?"#86efac":"#f9fafb",lineHeight:1}}>{sc}</div>
                <div style={{fontSize:9,color:mLive>0?"#ef4444":"#6b7280"}}>{mLive>0?"⚡ live":"PTS"}</div>
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  return(
    <div style={{...pg,display:"flex",flexDirection:"column",minHeight:"100vh"}}>
      <div style={{background:"linear-gradient(90deg,#001a00,#003d00)",padding:"10px 14px",borderBottom:`1px solid ${G}33`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",maxWidth:620,margin:"0 auto"}}>
          <div style={{display:"flex",alignItems:"center",gap:8,minWidth:0}}>
            <WcBadge size={28}/>
            <div style={{minWidth:0}}>
              <p style={{margin:0,fontSize:9,opacity:.5,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{ln} · <span style={{fontFamily:"monospace",letterSpacing:2}}>{cd}</span></p>
              <h2 style={{margin:"1px 0 0",fontSize:13,fontWeight:800,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{tn} <span style={{fontSize:10,color:"#86efac",fontWeight:400}}>({nm})</span></h2>
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
            {SaveBadge}
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:11,color:myScore>0?"#4ade80":"#86efac",fontWeight:700}}>{myScore} pts{liveCount>0?" ⚡":""}</div>
              <div style={{fontSize:9,color:apiStatus==="ok"?"#4ade80":apiStatus==="loading"?"#fbbf24":"#6b7280"}}>{apiStatus==="ok"?`● ${liveCount>0?`${liveCount} live`:"live"}`:"○ pre"}</div>
            </div>
            <button onClick={logout} style={{background:"transparent",border:`1px solid ${G}44`,borderRadius:6,padding:"4px 8px",color:"#6b7280",fontSize:10,cursor:"pointer"}}>⇄</button>
          </div>
        </div>
        <div style={{height:2,background:"rgba(255,255,255,.06)",borderRadius:2,maxWidth:620,margin:"6px auto 0"}}>
          <div style={{height:2,background:G,borderRadius:2,width:`${pct}%`,transition:"width .4s"}}/>
        </div>
      </div>

      <div style={{flex:1,overflowY:"auto"}}>
        <div style={{maxWidth:660,margin:"0 auto"}}>
          {tab==="sched"&&<SchedTab/>}
          {tab==="today"&&<TodayTab/>}
          {tab==="board"&&<BoardTab/>}
        </div>
      </div>

      <div style={{position:"sticky",bottom:0,background:"#0d1520",borderTop:`1px solid ${G}44`,display:"flex",zIndex:100}}>
        {[["sched","📅","Fixtures",`${picked}/${FIXTURES.length}`],["today","👥","All Picks",`${doneCount} done`],["board","🏆","Standings",`${myScore} pts`]].map(([t,icon,label,sub])=>(
          <button key={t} onClick={()=>setTab(t)} style={{flex:1,padding:"8px 0",border:"none",background:"transparent",cursor:"pointer",borderTop:`2px solid ${tab===t?G:"transparent"}`}}>
            <div style={{fontSize:20}}>{icon}</div>
            <div style={{fontSize:10,fontWeight:700,color:tab===t?"#86efac":"#9ca3af",marginTop:1}}>{label}</div>
            <div style={{fontSize:8,color:tab===t?G:"#6b7280",marginTop:1}}>{sub}</div>
          </button>
        ))}
      </div>
    </div>
  );
}