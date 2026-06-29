import { useState, useEffect, useCallback } from "react";
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, ReferenceLine, AreaChart, Area } from "recharts";
import * as XLSX from "xlsx";

/* ── TOKENS ─────────────────────────────────────────────────────────────── */
const T = {
  bg:"#F4F6F9",surface:"#FFFFFF",s2:"#F0F3F7",s3:"#E5EAF2",
  border:"#DDE3EE",blue:"#0A66C2",blueL:"#EBF3FC",
  gold:"#B8870A",goldL:"#FEF9E7",
  green:"#15803D",greenL:"#F0FDF4",
  amber:"#D97706",amberL:"#FFFBEB",
  red:"#DC2626",redL:"#FEF2F2",
  purple:"#7C3AED",purpleL:"#F5F3FF",
  text:"#0F172A",sec:"#475569",muted:"#94A3B8",
  sh:"0 1px 3px rgba(0,0,0,0.07),0 0 0 1px rgba(0,0,0,0.04)",
  shM:"0 4px 14px rgba(0,0,0,0.1),0 0 0 1px rgba(0,0,0,0.04)",
};
const CAT_C = {Branding:T.blue,Consulting:T.gold,Congruence:T.green};
const CAT_B = {Branding:T.blueL,Consulting:T.goldL,Congruence:T.greenL};
const BENCH = {
  engRate:{label:"Eng Rate",avg:3,good:5,excellent:7,unit:"%"},
  comments:{label:"Comments/Post",avg:4,good:10,excellent:20,unit:""},
  saves:{label:"Saves/Post",avg:2,good:5,excellent:10,unit:""},
  followers:{label:"Followers/Post",avg:2,good:6,excellent:12,unit:""},
};
const TODAY = "2026-06-25";
const KEY = {posts:"li_p3",weekly:"li_w3",followers:"li_f3",mods:"li_m3",sentiment:"li_s3"};

const DOMAINS = [
  {key:"brandManagement",  label:"Brand Management",   icon:"🏷", color:T.blue,   bg:T.blueL},
  {key:"advertising",      label:"Advertising",         icon:"📢", color:T.purple, bg:T.purpleL},
  {key:"consulting",       label:"Consulting",          icon:"📋", color:T.gold,   bg:T.goldL},
  {key:"productMarketing", label:"Product Marketing",   icon:"🚀", color:T.green,  bg:T.greenL},
];

/* ── CONTENT CALENDAR ────────────────────────────────────────────────────── */
const CAL = [
  {w:1,date:"2026-06-15",day:"Mon",type:"Deep",cat:"Congruence",topic:"Brand memory vs awareness — the verb test",pillar:"India's Brand Moment",hashtags:"#BrandStrategy #Branding #IndianMarketing #Marketing #Advertising"},
  {w:1,date:"2026-06-18",day:"Thu",type:"Sharp",cat:"Branding",topic:"Purpose branding as insurance, not strategy",pillar:"Contrarian Takes",hashtags:"#BrandStrategy #Marketing #Branding #PurposeBranding #Advertising"},
  {w:2,date:"2026-06-22",day:"Mon",type:"Deep",cat:"Branding",topic:"Fevicol — own the concept, not the category",pillar:"Campaign Decoded",hashtags:"#BrandStrategy #Branding #IndianMarketing #Advertising #Fevicol"},
  {w:2,date:"2026-06-25",day:"Thu",type:"Sharp",cat:"Branding",topic:"India D2C brands confusing distribution with brand",pillar:"India's Brand Moment",hashtags:"#D2C #IndianMarketing #Branding #StartupIndia #BrandStrategy"},
  {w:3,date:"2026-06-29",day:"Mon",type:"Deep",cat:"Congruence",topic:"BYJU's — brand equity as a lagging indicator",pillar:"Brand Autopsy",hashtags:"#BrandStrategy #EdTech #IndianMarketing #Marketing #Branding"},
  {w:3,date:"2026-07-02",day:"Thu",type:"Sharp",cat:"Congruence",topic:"Jio — pricing was the campaign, brand was the retention",pillar:"India's Brand Moment",hashtags:"#Telecom #IndianMarketing #BrandStrategy #Marketing #Branding"},
  {w:4,date:"2026-07-06",day:"Mon",type:"Deep",cat:"Congruence",topic:"Nykaa vs Purplle — two brand bets, two different moats",pillar:"India's Brand Moment",hashtags:"#BeautyIndustry #IndianMarketing #D2C #BrandStrategy #Branding"},
  {w:4,date:"2026-07-09",day:"Thu",type:"Sharp",cat:"Branding",topic:"Why 'impressions' is the metric that broke advertising",pillar:"Contrarian Takes",hashtags:"#DigitalMarketing #ContentStrategy #LinkedIn #BrandStrategy #Marketing"},
  {w:5,date:"2026-07-13",day:"Mon",type:"Deep",cat:"Congruence",topic:"Lego — brand focus as a financial decision",pillar:"Brand Autopsy",hashtags:"#BrandStrategy #BusinessStrategy #Marketing #Branding #CampaignDecoded"},
  {w:5,date:"2026-07-17",day:"Thu",type:"Sharp",cat:"Congruence",topic:"Men's formal wear and the culture shift — Peter England",pillar:"India's Brand Moment",hashtags:"#FashionMarketing #IndianMarketing #RetailMarketing #BrandStrategy #Branding"},
  {w:6,date:"2026-07-20",day:"Mon",type:"Deep",cat:"Congruence",topic:"Old Spice — abandoning one customer to earn another",pillar:"Campaign Decoded",hashtags:"#BrandStrategy #Advertising #CampaignDecoded #Marketing #Branding"},
  {w:6,date:"2026-07-24",day:"Thu",type:"Sharp",cat:"Branding",topic:"The saree paradox — Nalli",pillar:"India's Brand Moment",hashtags:"#IndianMarketing #FashionMarketing #Heritage #BrandStrategy #Branding"},
  {w:7,date:"2026-07-27",day:"Mon",type:"Deep",cat:"Consulting",topic:"Why brand ends up on slide 34 — and what it costs",pillar:"Consulting Lens",hashtags:"#Consulting #BusinessStrategy #BrandStrategy #Marketing #Management"},
  {w:7,date:"2026-07-31",day:"Thu",type:"Sharp",cat:"Congruence",topic:"Air India — the sequence that determines if brand rehab works",pillar:"India's Brand Moment",hashtags:"#AirIndia #IndianMarketing #BrandStrategy #Marketing #Branding"},
  {w:8,date:"2026-08-03",day:"Mon",type:"Deep",cat:"Branding",topic:"Ayurvedic heritage vs the naturals wave — Chandrika",pillar:"Brand Autopsy",hashtags:"#FMCG #IndianMarketing #Branding #BrandStrategy #Ayurveda"},
  {w:8,date:"2026-08-07",day:"Thu",type:"Sharp",cat:"Branding",topic:"Why Indian brands struggle when they go global",pillar:"Contrarian Takes",hashtags:"#IndianMarketing #GlobalBrands #Branding #BrandStrategy #Marketing"},
  {w:9,date:"2026-08-10",day:"Mon",type:"Deep",cat:"Congruence",topic:"American brands and Indian Gen Z — American Eagle",pillar:"Campaign Decoded",hashtags:"#GenZ #FashionMarketing #IndianMarketing #BrandStrategy #RetailMarketing"},
  {w:9,date:"2026-08-14",day:"Thu",type:"Sharp",cat:"Branding",topic:"Amul's 58-year formula — consistency as strategy",pillar:"Campaign Decoded",hashtags:"#IndianMarketing #Branding #Advertising #BrandStrategy #Marketing"},
  {w:10,date:"2026-08-17",day:"Mon",type:"Deep",cat:"Branding",topic:"CRED — building aspiration before utility in fintech",pillar:"Brand Autopsy",hashtags:"#Fintech #IndianMarketing #BrandStrategy #Marketing #Branding"},
  {w:10,date:"2026-08-21",day:"Thu",type:"Sharp",cat:"Consulting",topic:"The floor cleaner problem — commodity branding",pillar:"Consulting Lens",hashtags:"#Consulting #FMCG #BrandStrategy #Marketing #Management"},
  {w:11,date:"2026-08-24",day:"Mon",type:"Deep",cat:"Branding",topic:"Gen Z fashion identity in India — aspiration, redefined",pillar:"India's Brand Moment",hashtags:"#GenZ #IndianMarketing #FashionMarketing #BrandStrategy #Branding"},
  {w:11,date:"2026-08-28",day:"Thu",type:"Sharp",cat:"Consulting",topic:"The most dangerous person in a brand meeting has a 2x2",pillar:"Consulting Lens",hashtags:"#Consulting #BrandStrategy #BusinessStrategy #Marketing #Management"},
  {w:12,date:"2026-08-31",day:"Mon",type:"Deep",cat:"Branding",topic:"Duolingo — when the brand lives inside the product",pillar:"Brand Autopsy",hashtags:"#ProductMarketing #BrandStrategy #AppMarketing #Marketing #Branding"},
  {w:12,date:"2026-09-04",day:"Thu",type:"Sharp",cat:"Congruence",topic:"Energy category repositioning — Glucovita Bolts",pillar:"India's Brand Moment",hashtags:"#FMCG #IndianMarketing #EnergyDrinks #BrandStrategy #Marketing"},
  {w:13,date:"2026-09-07",day:"Mon",type:"Deep",cat:"Congruence",topic:"Freshworks vs Zoho — two Indian SaaS brand bets",pillar:"India's Brand Moment",hashtags:"#SaaS #IndianMarketing #TechMarketing #BrandStrategy #Branding"},
  {w:13,date:"2026-09-11",day:"Thu",type:"Sharp",cat:"Branding",topic:"Notion vs Confluence — identity in B2B software",pillar:"Campaign Decoded",hashtags:"#ProductMarketing #B2BMarketing #SaaS #BrandStrategy #Marketing"},
  {w:14,date:"2026-09-14",day:"Mon",type:"Deep",cat:"Branding",topic:"MTR Foods — heritage in a convenience era",pillar:"Brand Autopsy",hashtags:"#FMCG #IndianMarketing #FoodMarketing #BrandStrategy #Branding"},
  {w:14,date:"2026-09-18",day:"Thu",type:"Sharp",cat:"Consulting",topic:"Category creation vs category entry",pillar:"Consulting Lens",hashtags:"#Consulting #BrandStrategy #Marketing #BusinessStrategy #Management"},
  {w:15,date:"2026-09-21",day:"Mon",type:"Deep",cat:"Congruence",topic:"Value fashion positioning — Max Fashion",pillar:"India's Brand Moment",hashtags:"#FashionMarketing #IndianMarketing #RetailMarketing #BrandStrategy #Branding"},
  {w:15,date:"2026-09-25",day:"Thu",type:"Sharp",cat:"Branding",topic:"The health claim paradox — artificial sweeteners",pillar:"Brand Autopsy",hashtags:"#FMCG #HealthMarketing #IndianMarketing #BrandStrategy #Marketing"},
  {w:16,date:"2026-09-28",day:"Mon",type:"Deep",cat:"Congruence",topic:"Freemium as brand strategy — Slack, Notion, Figma",pillar:"Consulting Lens",hashtags:"#ProductMarketing #SaaS #BrandStrategy #Marketing #Branding"},
  {w:16,date:"2026-10-01",day:"Thu",type:"Sharp",cat:"Congruence",topic:"Stripe — building a developer brand through documentation",pillar:"Consulting Lens",hashtags:"#DeveloperMarketing #TechMarketing #BrandStrategy #Marketing #Branding"},
  {w:17,date:"2026-10-05",day:"Mon",type:"Deep",cat:"Congruence",topic:"Tata brand architecture — trust as a portfolio moat",pillar:"Brand Autopsy",hashtags:"#IndianMarketing #BrandStrategy #Tata #Marketing #Branding"},
  {w:17,date:"2026-10-08",day:"Thu",type:"Sharp",cat:"Consulting",topic:"Why most tech product launches fail — positioning first",pillar:"Consulting Lens",hashtags:"#ProductMarketing #TechMarketing #Consulting #BrandStrategy #Marketing"},
];

/* ── STORAGE ─────────────────────────────────────────────────────────────── */
const sg = async k => { try { const r = localStorage.getItem(k); return r ? JSON.parse(r) : []; } catch { return []; } };
const ss = async (k,v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } };

/* ── XLSX PARSER — multi-strategy ────────────────────────────────────────── */
function pn(v) { return parseFloat(String(v??"").replace(/[%,\s]/g,""))||0; }

function fmtD(raw) {
  const s = String(raw??"").trim();
  if(/^\d{4,5}$/.test(s)){try{const d=XLSX.SSF.parse_date_code(parseInt(s));if(d)return`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.m-1]} ${d.d}`;}catch{}}
  const mo={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
  const low=s.toLowerCase();
  for(const[m]of Object.entries(mo))if(low.includes(m)){const dn=s.match(/\d+/g);if(dn){const d=dn.find(x=>+x>=1&&+x<=31);if(d)return`${m[0].toUpperCase()+m.slice(1)} ${d}`;}}
  return s.slice(-5)||s;
}

// Classify a key string into a known metric
function classifyKey(k) {
  const l = k.toLowerCase().trim();
  if(l.includes("impression")&&!l.includes("click")&&!l.includes("unique")) return "impressions";
  if((l.includes("unique")&&l.includes("impression"))||l.includes("member")||l.includes("reach")) return "reach";
  if(l.includes("reaction")||l==="likes"||l.includes("total like")) return "reactions";
  if(l.includes("comment")&&!l.includes("uncomment")) return "comments";
  if(l.includes("save")||l.includes("bookmark")) return "saves";
  if(l.includes("repost")||l.includes("reshare")) return "reposts";
  if(l.includes("follower")||(l.includes("new")&&l.includes("follow"))) return "newFollowers";
  if(l==="engagement"||l.includes("total engagement")) return "engagements";
  if(l.includes("engagement")&&l.includes("rate")) return "engRate";
  if(l.includes("date")) return "date";
  if(l.includes("seniority")) return "DEMO_seniority";
  if(l.includes("industry")) return "DEMO_industry";
  if(l.includes("function")) return "DEMO_function";
  if(l.includes("location")||l.includes("geo")) return "DEMO_location";
  if(l.includes("company")) return "DEMO_company";
  return null;
}

function parseAnyFile(buf) {
  // Returns { dayData, totals, demo, detectedType }
  const wb = XLSX.read(buf,{type:"array",cellDates:false});
  const result = {dayData:[],totals:{},demo:{},detectedType:"unknown"};

  for(const sn of wb.SheetNames){
    const ws = wb.Sheets[sn];
    const raw = XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
    if(!raw.length) continue;

    // ── LinkedIn named-sheet handlers (specific to LinkedIn's export format) ──
    const SN = sn.toUpperCase();

    // DEMOGRAPHICS sheet: col A = category type, col B = value, col C = percentage
    if(SN === "DEMOGRAPHICS"){
      const groups = {};
      for(let i=1;i<raw.length;i++){
        const r=raw[i];
        const cat=String(r[0]??"").trim();
        const val=String(r[1]??"").trim();
        const pctStr=String(r[2]??"").trim();
        if(!cat||!val) continue;
        const pct=pctStr==="< 1%"?0.5:parseFloat(pctStr)||0;
        if(!groups[cat]) groups[cat]=[];
        groups[cat].push({name:val,count:pct,pct:Math.round(pct)||1});
      }
      const catMap={"Seniority":"seniority","Industry":"industry","Job function":"function","Job title":"jobTitle","Location":"location","Company":"company","Company size":"companySize"};
      for(const[cat,items] of Object.entries(groups)){const k=catMap[cat];if(k) result.demo[k]=items;}
      result.detectedType="followers";
      continue;
    }

    // FOLLOWERS sheet: row 1 = total count kv, rows 3+ = date / new followers
    if(SN === "FOLLOWERS"){
      if(raw[0]?.[1]) result.totals.totalFollowers=pn(raw[0][1]);
      for(let i=2;i<raw.length;i++){
        const r=raw[i];
        if(!r[0]||String(r[0]).toLowerCase().includes("date")) continue;
        const d=fmtD(r[0]); const nf=pn(r[1]);
        const ex=result.dayData.find(x=>x.date===d);
        if(ex) ex.newFollowers=nf;
        else result.dayData.push({date:d,impressions:0,members:0,engagements:0,engRate:0,newFollowers:nf});
      }
      result.detectedType="followers"; continue;
    }

    // DISCOVERY sheet: summary impressions + reach
    if(SN === "DISCOVERY"){
      for(let i=0;i<raw.length;i++){
        const k=String(raw[i][0]??"").toLowerCase();
        if(k.includes("impression")) result.totals.impressions=Math.max(result.totals.impressions||0,pn(raw[i][1]));
        if(k.includes("member")||k.includes("reach")) result.totals.reach=Math.max(result.totals.reach||0,pn(raw[i][1]));
      }
      continue;
    }

    // ENGAGEMENT sheet: daily impressions + engagements
    if(SN === "ENGAGEMENT"){
      for(let i=1;i<raw.length;i++){
        const r=raw[i];
        if(!r[0]||String(r[0]).toLowerCase().includes("date")) continue;
        const d=fmtD(r[0]); const imp=pn(r[1]); const eng=pn(r[2]);
        const ex=result.dayData.find(x=>x.date===d);
        if(ex){ex.impressions=imp;ex.engagements=eng;}
        else result.dayData.push({date:d,impressions:imp,members:0,engagements:eng,engRate:0,newFollowers:0});
      }
      result.detectedType="weekly"; continue;
    }
    // ── End LinkedIn named-sheet handlers ────────────────────────────────────

    // Find header row: first row where >=2 cells match known metrics
    let hdrIdx = -1;
    for(let i=0;i<Math.min(raw.length,20);i++){
      const matches = raw[i].filter(c=>classifyKey(String(c??""))!==null).length;
      if(matches>=2){hdrIdx=i;break;}
    }

    if(hdrIdx>=0){
      const hdrs = raw[hdrIdx].map(c=>({raw:String(c??""),cls:classifyKey(String(c??""))}));
      const dataRows = raw.slice(hdrIdx+1).filter(r=>r.some(c=>c!==""));

      // Check if it has a date column → daily format
      const dateColIdx = hdrs.findIndex(h=>h.cls==="date");
      const hasDate = dateColIdx>=0;
      const metricCols = hdrs.map((h,i)=>({...h,i})).filter(h=>h.cls&&h.cls!=="date"&&!h.cls.startsWith("DEMO_"));
      const demoCols  = hdrs.map((h,i)=>({...h,i})).filter(h=>h.cls?.startsWith("DEMO_"));

      if(hasDate && dataRows.length>1){
        // Daily aggregate format
        result.detectedType = "weekly";
        for(const row of dataRows){
          const dateVal = row[dateColIdx];
          if(!dateVal) continue;
          const day = {date:fmtD(dateVal),impressions:0,members:0,engagements:0,engRate:0,newFollowers:0,reactions:0,comments:0,saves:0,reposts:0};
          for(const col of metricCols) if(col.cls in day) day[col.cls]=pn(row[col.i]);
          if(day.impressions>0||day.engagements>0) result.dayData.push(day);
        }
      }

      if(!hasDate && dataRows.length>=1){
        // Single row totals format
        result.detectedType = "post";
        for(const row of dataRows){
          let gotAny=false;
          for(const col of metricCols){
            const v=pn(row[col.i]);
            if(v>0){gotAny=true;result.totals[col.cls]=Math.max(result.totals[col.cls]||0,v);}
          }
          if(gotAny) break; // take first data row only
        }
      }

      // Demo sheets
      if(demoCols.length===0 && raw[hdrIdx].length<=2){
        // Could be a key-value demographic sheet: col0=name, col1=count
        const sname = sn.toLowerCase();
        let demoKey = null;
        if(sname.includes("seniority")) demoKey="seniority";
        else if(sname.includes("industry")) demoKey="industry";
        else if(sname.includes("function")) demoKey="function";
        else if(sname.includes("location")||sname.includes("geo")) demoKey="location";
        else if(sname.includes("company")) demoKey="company";
        if(demoKey){
          const items=[];
          for(let i=hdrIdx+1;i<raw.length;i++){const r=raw[i];const nm=String(r[0]??"").trim();if(!nm)continue;items.push({name:nm,count:pn(r[1])||1});}
          if(items.length){const tot=items.reduce((s,x)=>s+x.count,0);result.demo[demoKey]=items.map(x=>({...x,pct:tot>0?Math.round((x.count/tot)*100):0}));}
        }
      }
    }

    // Key-value scan (catch summary rows, even without proper headers)
    const kv_totals = {};
    for(const row of raw){
      if(row.length<2) continue;
      const k=String(row[0]??"").trim();
      const v=pn(row[1]);
      if(!k||!v) continue;
      const cls=classifyKey(k);
      if(cls&&!cls.startsWith("DEMO_")&&cls!=="date")
        kv_totals[cls]=Math.max(kv_totals[cls]||0,v);
    }
    // Merge kv_totals into result.totals (only fill gaps)
    for(const [k,v] of Object.entries(kv_totals))
      if(!result.totals[k]||result.totals[k]===0) result.totals[k]=v;
  }

  // If we have dayData but no totals, derive totals from dayData
  if(result.dayData.length>0){
    result.totals.impressions = result.totals.impressions||result.dayData.reduce((s,d)=>s+d.impressions,0);
    result.totals.reach = result.totals.reach||result.dayData.reduce((s,d)=>s+d.members,0);
    result.totals.newFollowers = result.totals.newFollowers||result.dayData.reduce((s,d)=>s+d.newFollowers,0);
    if(!result.totals.engagements)result.totals.engagements=result.dayData.reduce((s,d)=>s+d.engagements,0);
  }

  return result;
}

/* ── UTILS ───────────────────────────────────────────────────────────────── */
const fn = n => (!n&&n!==0)?"—":n>=1000?(n/1000).toFixed(1).replace(/\.0$/,"")+"k":String(Math.round(n));
const fp = n => n?n.toFixed(1)+"%":"—";
function engRateFrom(t,d){ const i=t?.impressions||d?.reduce((s,x)=>s+x.impressions,0)||0;const e=(t?.reactions||0)+(t?.comments||0)+(t?.saves||0)+(t?.reposts||0)||(d?.reduce((s,x)=>s+x.engagements,0)||0);return i>0?(e/i)*100:0; }
function bColor(v,b){ return v>=b.excellent?T.green:v>=b.good?T.gold:v>=b.avg?T.blue:T.red; }
function bLabel(v,b){ return v>=b.excellent?"Excellent":v>=b.good?"Good":v>=b.avg?"Average":"Below avg"; }
function calStatus(date,pubDates,pubTopics){
  const byDate=pubDates.has(date);
  if((byDate)&&date<=TODAY) return "published";
  if(date===TODAY) return "today";
  if(date<TODAY&&!byDate) return "overdue";
  return "upcoming";
}

/* ── MINI COMPONENTS ─────────────────────────────────────────────────────── */
const Card=({children,style={}})=><div style={{background:T.surface,borderRadius:12,border:`1px solid ${T.border}`,boxShadow:T.sh,...style}}>{children}</div>;
const Tag=({cat})=><span style={{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20,background:CAT_B[cat]||T.s2,color:CAT_C[cat]||T.sec}}>{cat}</span>;
const Pill=({label,color,bg})=><span style={{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20,background:bg||T.s2,color:color||T.sec}}>{label}</span>;
const SHdr=({t,s})=><div style={{marginBottom:14}}><div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:14,color:T.text}}>{t}</div>{s&&<div style={{fontSize:11.5,color:T.muted,marginTop:2}}>{s}</div>}</div>;
const CTip=({active,payload,label})=>active&&payload?.length?<div style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:8,padding:"10px 14px",boxShadow:T.shM,fontSize:12}}><div style={{fontWeight:700,color:T.text,marginBottom:4}}>{label}</div>{payload.map((p,i)=><div key={i} style={{color:p.color}}>{p.name}: <strong>{typeof p.value==="number"&&p.value<100?p.value.toFixed(1)+"%":fn(p.value)}</strong></div>)}</div>:null;

function KPI({label,value,sub,color,bv,bench}){
  const bc=bench&&bv!=null?bColor(bv,bench):null;
  return <Card style={{padding:"18px 20px",flex:1,minWidth:120}}>
    <div style={{fontSize:10.5,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>{label}</div>
    <div style={{fontSize:26,fontWeight:800,color:color||T.text,fontFamily:"'Syne',sans-serif",lineHeight:1,marginBottom:3}}>{value}</div>
    {sub&&<div style={{fontSize:11,color:T.muted}}>{sub}</div>}
    {bc&&<div style={{display:"inline-flex",alignItems:"center",gap:3,marginTop:5,background:bc===T.green?T.greenL:bc===T.gold?T.goldL:bc===T.blue?T.blueL:T.redL,borderRadius:20,padding:"2px 7px"}}>
      <div style={{width:5,height:5,borderRadius:"50%",background:bc}}/><span style={{fontSize:10,fontWeight:700,color:bc}}>{bLabel(bv,bench)}</span>
    </div>}
  </Card>;
}

function BenchBar({label,val,bench,unit=""}){
  const max=bench.excellent*1.4,pct=Math.min((val/max)*100,100);
  const col=bColor(val,bench);
  return <div style={{marginBottom:12}}>
    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
      <span style={{fontSize:12.5,fontWeight:600,color:T.sec}}>{label}</span>
      <span style={{fontSize:12,fontWeight:700,color:col}}>{val?`${val.toFixed(1)}${unit}`:"—"} <span style={{color:T.muted,fontWeight:500}}>({bLabel(val,bench)})</span></span>
    </div>
    <div style={{position:"relative",background:T.s3,borderRadius:5,height:7,overflow:"hidden"}}>
      <div style={{position:"absolute",left:0,top:0,width:`${pct}%`,background:col,height:"100%",borderRadius:5,transition:"width 0.6s ease"}}/>
      {[bench.avg,bench.good,bench.excellent].map((v,i)=><div key={i} style={{position:"absolute",left:`${(v/max)*100}%`,top:0,width:1,height:"100%",background:["#D97706","#B8870A","#15803D"][i],opacity:0.5}}/>)}
    </div>
    <div style={{display:"flex",justifyContent:"space-between",marginTop:2}}>
      {[["Avg",bench.avg],["Good",bench.good],["Excellent",bench.excellent]].map(([l,v])=><span key={l} style={{fontSize:9.5,color:T.muted}}>{l}: {v}{unit}</span>)}
    </div>
  </div>;
}

function StatusBadge({status}){
  const c={published:{bg:T.greenL,color:T.green,icon:"✓",label:"Published"},today:{bg:T.blueL,color:T.blue,icon:"→",label:"Today"},overdue:{bg:T.redL,color:T.red,icon:"!",label:"Overdue"},upcoming:{bg:T.s2,color:T.muted,icon:"·",label:"Upcoming"}}[status]||{bg:T.s2,color:T.muted,icon:"·",label:"—"};
  return <span style={{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20,background:c.bg,color:c.color,display:"inline-flex",alignItems:"center",gap:3}}><span>{c.icon}</span>{c.label}</span>;
}

function UploadZone({id,label,hint,onFiles}){
  const [drag,setDrag]=useState(false);
  return <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
    onDrop={e=>{e.preventDefault();setDrag(false);const fs=Array.from(e.dataTransfer?.files||[]).filter(f=>f.name.endsWith(".xlsx"));if(fs.length)onFiles(fs);}}
    onClick={()=>document.getElementById(id).click()}
    style={{border:`2px dashed ${drag?T.blue:T.border}`,borderRadius:10,background:drag?T.blueL:T.s2,padding:"22px 16px",textAlign:"center",cursor:"pointer",transition:"all 0.15s"}}>
    <input id={id} type="file" accept=".xlsx" multiple onChange={e=>{const fs=Array.from(e.target.files||[]);if(fs.length)onFiles(fs);e.target.value="";}} style={{display:"none"}}/>
    <div style={{fontSize:13,fontWeight:700,color:T.sec,marginBottom:3}}>{label}</div>
    <div style={{fontSize:11.5,color:T.muted}}>{hint}</div>
  </div>;
}

/* ── OVERVIEW ────────────────────────────────────────────────────────────── */
function Overview({posts,weekly,followers}){
  const totImp=posts.reduce((s,p)=>s+(p.totals?.impressions||0),0);
  const totReact=posts.reduce((s,p)=>s+(p.totals?.reactions||0),0);
  const totComm=posts.reduce((s,p)=>s+(p.totals?.comments||0),0);
  const totSaves=posts.reduce((s,p)=>s+(p.totals?.saves||0),0);
  const totFoll=posts.reduce((s,p)=>s+(p.totals?.newFollowers||0),0)+weekly.reduce((s,w)=>s+w.dayData.reduce((ss,d)=>ss+d.newFollowers,0),0);
  const avgEng=posts.length?posts.reduce((s,p)=>s+engRateFrom(p.totals,[]),0)/posts.length:0;
  const avgComm=posts.length?totComm/posts.length:0;
  const avgSav=posts.length?totSaves/posts.length:0;
  const avgFoll=posts.length?totFoll/posts.length:0;

  // Audience: merge all demographics from posts and follower reports
  const latestDemo = (() => {
    const d = {};
    for(const p of posts) if(p.demographics) Object.assign(d,p.demographics);
    const lf = [...followers].sort((a,b)=>a.month>b.month?-1:1)[0];
    if(lf?.demographics) Object.assign(d,lf.demographics);
    return d;
  })();

  const bestPost=[...posts].sort((a,b)=>(b.totals?.impressions||0)-(a.totals?.impressions||0))[0];
  const weekTrend=weekly.map(w=>({label:w.weekLabel||"?",impressions:w.dayData.reduce((s,d)=>s+d.impressions,0)}));
  const nextPosts=CAL.filter(c=>c.date>=TODAY).slice(0,3);

  if(!posts.length&&!weekly.length) return (
    <div style={{textAlign:"center",padding:"70px 20px"}}>
      <div style={{fontSize:42,marginBottom:16}}>📊</div>
      <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:20,color:T.text,marginBottom:8}}>No data yet</div>
      <div style={{fontSize:14,color:T.muted}}>Go to Reports and upload your LinkedIn exports. Data appears here automatically.</div>
    </div>
  );

  return <div>
    <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:20}}>
      <KPI label="Posts Tracked" value={posts.length||"—"} sub={`${CAL.filter(c=>c.date<=TODAY).length} due so far`} color={T.blue}/>
      <KPI label="Total Impressions" value={fn(totImp)} sub="All posts combined" color={T.blue}/>
      <KPI label="Avg Eng Rate" value={fp(avgEng)} sub="vs 3% industry avg" color={bColor(avgEng,BENCH.engRate)} bv={avgEng} bench={BENCH.engRate}/>
      <KPI label="Reactions" value={fn(totReact)} sub="Total"/>
      <KPI label="Comments" value={fn(totComm)} sub="Highest algo weight" color={T.green} bv={avgComm} bench={BENCH.comments}/>
      <KPI label="Saves" value={fn(totSaves)} sub="Quality signal" color={T.gold} bv={avgSav} bench={BENCH.saves}/>
      <KPI label="Followers Gained" value={fn(totFoll)} sub="From content" color={T.purple} bv={avgFoll} bench={BENCH.followers}/>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Performance vs Benchmarks" s="LinkedIn personal profiles, 2026"/>
        <BenchBar label="Engagement Rate" val={avgEng} bench={BENCH.engRate} unit="%"/>
        <BenchBar label="Comments / Post" val={avgComm} bench={BENCH.comments}/>
        <BenchBar label="Saves / Post" val={avgSav} bench={BENCH.saves}/>
        <BenchBar label="Followers / Post" val={avgFoll} bench={BENCH.followers}/>
      </Card>
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Impression Trend" s={weekTrend.length?`${weekTrend.length} weeks`:"Upload weekly reports to see trend"}/>
        {weekTrend.length>0?<ResponsiveContainer width="100%" height={185}>
          <AreaChart data={weekTrend} margin={{top:4,right:8,left:-24,bottom:0}}>
            <defs><linearGradient id="ag" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={T.blue} stopOpacity={0.15}/><stop offset="95%" stopColor={T.blue} stopOpacity={0}/></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/>
            <XAxis dataKey="label" tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false}/>
            <YAxis tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false} tickFormatter={fn}/>
            <Tooltip content={<CTip/>}/>
            <Area type="monotone" dataKey="impressions" name="Impressions" stroke={T.blue} strokeWidth={2.5} fill="url(#ag)" dot={{r:4,fill:T.blue,strokeWidth:0}}/>
          </AreaChart>
        </ResponsiveContainer>:<div style={{height:185,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted,fontSize:13}}>No weekly data</div>}
      </Card>
    </div>

    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      {/* Audience snapshot */}
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Audience Snapshot" s="From uploaded post and follower reports"/>
        {Object.keys(latestDemo).length>0?
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
            {Object.entries(latestDemo).slice(0,4).map(([key,vals])=>
              <div key={key}>
                <div style={{fontSize:10.5,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>{key}</div>
                {vals.slice(0,4).map((v,i)=><div key={i} style={{marginBottom:7}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}>
                    <span style={{fontSize:11.5,color:T.sec,fontWeight:500}}>{v.name}</span>
                    <span style={{fontSize:11,color:T.muted,fontWeight:600}}>{v.pct}%</span>
                  </div>
                  <div style={{background:T.s3,borderRadius:4,height:4}}><div style={{width:`${v.pct}%`,background:T.blue,height:"100%",borderRadius:4}}/></div>
                </div>)}
              </div>
            )}
          </div>
        :<div style={{color:T.muted,fontSize:13,textAlign:"center",padding:"20px 0"}}>Upload post or follower reports to see audience data</div>}
      </Card>

      {/* Next posts */}
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Coming Up" s="From your 17-week content plan"/>
        {nextPosts.map((p,i)=><div key={i} style={{display:"flex",gap:12,alignItems:"flex-start",marginBottom:i<2?14:0,paddingBottom:i<2?14:0,borderBottom:i<2?`1px solid ${T.border}`:"none"}}>
          <div style={{background:p.date===TODAY?T.blueL:T.s2,borderRadius:8,padding:"6px 10px",textAlign:"center",flexShrink:0,minWidth:42}}>
            <div style={{fontSize:11,fontWeight:700,color:p.date===TODAY?T.blue:T.sec}}>{p.day}</div>
            <div style={{fontSize:10,color:T.muted}}>{p.date.slice(5).replace("-","/")}</div>
          </div>
          <div>
            <div style={{fontSize:12.5,fontWeight:600,color:T.text,lineHeight:1.35,marginBottom:4}}>{p.topic}</div>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}><Tag cat={p.cat}/><Pill label={p.type} color={p.type==="Deep"?T.blue:T.gold} bg={p.type==="Deep"?T.blueL:T.goldL}/></div>
          </div>
        </div>)}
      </Card>
    </div>

    {/* Posts table */}
    {posts.length>0&&<Card style={{padding:"20px 22px"}}>
      <SHdr t="All Tracked Posts" s={`${posts.length} posts`}/>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5}}>
          <thead><tr style={{borderBottom:`2px solid ${T.border}`}}>
            {["Date","Topic","Cat","Impressions","Reach","Reactions","Comments","Saves","Eng Rate","Followers"].map(h=><th key={h} style={{textAlign:"left",padding:"8px 10px",color:T.muted,fontWeight:700,fontSize:10.5,textTransform:"uppercase",letterSpacing:"0.06em",whiteSpace:"nowrap"}}>{h}</th>)}
          </tr></thead>
          <tbody>{[...posts].sort((a,b)=>a.date>b.date?-1:1).map((p,i)=>{
            const er=engRateFrom(p.totals,[]);
            return <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:i%2?T.s2:"transparent"}}>
              <td style={{padding:"10px 10px",color:T.muted,whiteSpace:"nowrap"}}>{p.date||"—"}</td>
              <td style={{padding:"10px 10px",maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontWeight:500,color:T.text}}>{p.topic||"—"}</div></td>
              <td style={{padding:"10px 10px"}}><Tag cat={p.category||"Congruence"}/></td>
              <td style={{padding:"10px 10px",fontWeight:700,color:T.blue}}>{fn(p.totals?.impressions)}</td>
              <td style={{padding:"10px 10px",color:T.sec}}>{fn(p.totals?.reach)}</td>
              <td style={{padding:"10px 10px",color:T.sec}}>{fn(p.totals?.reactions)}</td>
              <td style={{padding:"10px 10px",color:T.green,fontWeight:(p.totals?.comments||0)>5?700:400}}>{fn(p.totals?.comments)}</td>
              <td style={{padding:"10px 10px",color:T.gold,fontWeight:(p.totals?.saves||0)>3?700:400}}>{fn(p.totals?.saves)}</td>
              <td style={{padding:"10px 10px"}}>{er>0?<span style={{color:bColor(er,BENCH.engRate),fontWeight:700}}>{fp(er)}</span>:<span style={{color:T.muted}}>—</span>}</td>
              <td style={{padding:"10px 10px",color:T.purple}}>{fn(p.totals?.newFollowers)}</td>
            </tr>;
          })}</tbody>
        </table>
      </div>
    </Card>}
  </div>;
}

/* ── REPORTS ─────────────────────────────────────────────────────────────── */
function Reports({posts,weekly,followers,onSavePost,onSaveWeekly,onSaveFollower,onDelete}){
  const [sub,setSub]=useState("post");
  const [proc,setProc]=useState(false);
  const [msg,setMsg]=useState({text:"",ok:true});
  const flash=(text,ok=true)=>{setMsg({text,ok});setTimeout(()=>setMsg({text:"",ok:true}),3500);};

  // Post upload
  const [pTopic,setPTopic]=useState("");
  const [pDate,setPDate]=useState(TODAY);
  const [pCat,setPCat]=useState("Congruence");
  const [pFiles,setPFiles]=useState([]);
  const [pPreview,setPPreview]=useState(null);

  // Weekly upload
  const [wLabel,setWLabel]=useState("");
  const [wFiles,setWFiles]=useState([]);
  const [wPreview,setWPreview]=useState(null);

  // Follower upload
  const [fMonth,setFMonth]=useState(TODAY.slice(0,7));
  const [fFiles,setFFiles]=useState([]);
  const [fPreview,setFPreview]=useState(null);

  const calEntry = CAL.find(c=>c.date===pDate)||CAL.find(c=>c.topic===pTopic);

  // Parse and preview post files
  const previewPost = async files => {
    setPFiles(files);
    const merged={totals:{},demo:{}};
    for(const f of files){
      const buf=new Uint8Array(await f.arrayBuffer());
      const res=parseAnyFile(buf);
      Object.assign(merged.totals,res.totals);
      Object.assign(merged.demo,res.demo);
    }
    setPPreview(merged);
  };

  const previewWeekly = async files => {
    setWFiles(files);
    let days=[];
    for(const f of files){
      const buf=new Uint8Array(await f.arrayBuffer());
      const res=parseAnyFile(buf);
      days=[...days,...res.dayData];
    }
    const seen={};for(const d of days)if(!seen[d.date]||d.impressions>seen[d.date].impressions)seen[d.date]=d;
    setWPreview(Object.values(seen));
  };

  const previewFollower = async files => {
    setFFiles(files);
    let demo={};
    for(const f of files){const buf=new Uint8Array(await f.arrayBuffer());const res=parseAnyFile(buf);Object.assign(demo,res.demo);}
    setFPreview(demo);
  };

  const savePost = async () => {
    if(!pPreview){flash("Parse files first by dropping them above","false");return;}
    setProc(true);
    const entry={id:Date.now().toString(),date:pDate,topic:pTopic||calEntry?.topic||"Post "+pDate,category:pCat||calEntry?.cat||"Congruence",totals:pPreview.totals,demographics:pPreview.demo,uploadedAt:new Date().toISOString()};
    await onSavePost(entry);
    setPFiles([]);setPPreview(null);setPTopic("");
    flash("✓ Post saved. Overview updated.");setProc(false);
  };

  const saveWeekly = async () => {
    if(!wPreview?.length){flash("No day data found in these files. Are you sure this is an aggregate export?");return;}
    setProc(true);
    const entry={id:Date.now().toString(),weekLabel:wLabel||`Week of ${wPreview[0]?.date||"?"}`,dayData:wPreview,uploadedAt:new Date().toISOString()};
    await onSaveWeekly(entry);
    setWFiles([]);setWPreview(null);setWLabel("");
    flash("✓ Weekly data saved.");setProc(false);
  };

  const saveFollower = async () => {
    if(!fPreview||!Object.keys(fPreview).length){flash("No demographic data found. Check the file format.");return;}
    setProc(true);
    const entry={id:Date.now().toString(),month:fMonth,demographics:fPreview,uploadedAt:new Date().toISOString()};
    await onSaveFollower(entry);
    setFFiles([]);setFPreview(null);
    flash("✓ Follower report saved.");setProc(false);
  };

  const Btn=({onClick,disabled,label})=><button onClick={onClick} disabled={disabled} style={{width:"100%",background:disabled?T.s3:T.blue,color:disabled?T.muted:"#fff",border:"none",borderRadius:8,padding:"11px",fontSize:13,fontWeight:700,cursor:disabled?"default":"pointer",fontFamily:"inherit",transition:"background 0.15s"}}>{label}</button>;

  const tabs=[["post","📝 Post Reports"],["weekly","📅 Weekly Reports"],["follower","👥 Follower Reports"]];

  return <div>
    <div style={{display:"flex",gap:4,marginBottom:22,borderBottom:`1px solid ${T.border}`}}>
      {tabs.map(([k,l])=><button key={k} onClick={()=>setSub(k)} style={{padding:"10px 18px",background:"none",border:"none",borderBottom:`2px solid ${sub===k?T.blue:"transparent"}`,color:sub===k?T.blue:T.sec,fontWeight:sub===k?700:500,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>)}
    </div>

    {msg.text&&<div style={{background:msg.ok!==false?T.greenL:T.redL,border:`1px solid ${msg.ok!==false?T.green:T.red}`,color:msg.ok!==false?T.green:T.red,borderRadius:8,padding:"10px 16px",fontSize:13,fontWeight:600,marginBottom:16}}>{msg.text}</div>}

    {/* POST REPORTS */}
    {sub==="post"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
      <Card style={{padding:"22px"}}>
        <SHdr t="Upload Post Analytics" s="Single post analytics export from LinkedIn"/>
        <div style={{marginBottom:12}}>
          <label style={{fontSize:11.5,fontWeight:700,color:T.sec,display:"block",marginBottom:5}}>Match to calendar (or type manually)</label>
          <select value={pTopic} onChange={e=>{setPTopic(e.target.value);const c=CAL.find(x=>x.topic===e.target.value);if(c){setPDate(c.date);setPCat(c.cat);}}} style={{width:"100%",border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 12px",fontSize:12.5,color:T.text,background:T.surface,fontFamily:"inherit",outline:"none"}}>
            <option value="">— Select from content calendar —</option>
            {CAL.filter(c=>c.date<=TODAY).map((c,i)=><option key={i} value={c.topic}>{c.date} · {c.topic.slice(0,55)}</option>)}
          </select>
          {!pTopic&&<input placeholder="Or type topic manually…" onChange={e=>setPTopic(e.target.value)} style={{width:"100%",border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 12px",fontSize:12.5,color:T.text,background:T.s2,fontFamily:"inherit",outline:"none",marginTop:8}}/>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
          <div>
            <label style={{fontSize:11.5,fontWeight:700,color:T.sec,display:"block",marginBottom:5}}>Post date</label>
            <input type="date" value={pDate} onChange={e=>setPDate(e.target.value)} style={{width:"100%",border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 12px",fontSize:12.5,color:T.text,background:T.surface,fontFamily:"inherit",outline:"none"}}/>
          </div>
          <div>
            <label style={{fontSize:11.5,fontWeight:700,color:T.sec,display:"block",marginBottom:5}}>Category</label>
            <select value={pCat} onChange={e=>setPCat(e.target.value)} style={{width:"100%",border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 12px",fontSize:12.5,color:T.text,background:T.surface,fontFamily:"inherit",outline:"none"}}>
              <option>Congruence</option><option>Branding</option><option>Consulting</option>
            </select>
          </div>
        </div>
        <div style={{marginBottom:14}}>
          <UploadZone id="pfu" label="Drop post analytics .xlsx" hint="LinkedIn single post export" onFiles={previewPost}/>
          {pFiles.length>0&&<div style={{marginTop:8,display:"flex",flexWrap:"wrap",gap:5}}>{pFiles.map((f,i)=><span key={i} style={{fontSize:10.5,background:T.blueL,color:T.blue,padding:"2px 9px",borderRadius:20,fontWeight:600}}>{f.name.slice(0,28)}</span>)}</div>}
        </div>
        {/* Preview */}
        {pPreview&&<div style={{background:T.s2,borderRadius:8,padding:"12px 14px",marginBottom:14,fontSize:12}}>
          <div style={{fontWeight:700,color:T.text,marginBottom:8}}>Parsed data preview:</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
            {Object.entries(pPreview.totals).filter(([,v])=>v>0).map(([k,v])=><div key={k} style={{background:T.surface,borderRadius:6,padding:"6px 10px"}}>
              <div style={{fontSize:14,fontWeight:800,color:T.blue}}>{fn(v)}</div>
              <div style={{fontSize:10,color:T.muted,textTransform:"capitalize"}}>{k}</div>
            </div>)}
          </div>
          {!Object.values(pPreview.totals).some(v=>v>0)&&<div style={{color:T.amber,fontSize:12}}>⚠ No metrics detected. This file may be an aggregate export — try uploading it under Weekly Reports instead.</div>}
        </div>}
        <Btn onClick={savePost} disabled={proc||!pFiles.length} label={proc?"Saving…":"Save Post Data"}/>
      </Card>

      <Card style={{padding:"22px"}}>
        <SHdr t="Saved Posts" s={`${posts.length} tracked`}/>
        <div style={{maxHeight:420,overflowY:"auto",display:"flex",flexDirection:"column",gap:8}}>
          {!posts.length&&<div style={{color:T.muted,fontSize:13,textAlign:"center",padding:"30px 0"}}>No posts saved yet</div>}
          {[...posts].sort((a,b)=>b.date>a.date?1:-1).map((p,i)=><div key={i} style={{background:T.s2,borderRadius:9,padding:"11px 13px",display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:600,color:T.text,lineHeight:1.3,marginBottom:4}}>{p.topic||"—"}</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                <Tag cat={p.category||"Congruence"}/>
                <span style={{fontSize:10.5,color:T.muted}}>{p.date}</span>
                {p.totals?.impressions>0&&<span style={{fontSize:10.5,color:T.blue,fontWeight:600}}>{fn(p.totals.impressions)} impr</span>}
                {p.totals?.reactions>0&&<span style={{fontSize:10.5,color:T.sec}}>{fn(p.totals.reactions)}R · {fn(p.totals.comments)}C · {fn(p.totals.saves)}S</span>}
              </div>
            </div>
            <button onClick={()=>onDelete("post",p.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,lineHeight:1}}>×</button>
          </div>)}
        </div>
      </Card>
    </div>}

    {/* WEEKLY REPORTS */}
    {sub==="weekly"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
      <Card style={{padding:"22px"}}>
        <SHdr t="Upload Weekly Report" s="Aggregate analytics export (date range)"/>
        <div style={{marginBottom:12}}>
          <label style={{fontSize:11.5,fontWeight:700,color:T.sec,display:"block",marginBottom:5}}>Week label</label>
          <input value={wLabel} onChange={e=>setWLabel(e.target.value)} placeholder="e.g. Week 1 · Jun 15–21" style={{width:"100%",border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 12px",fontSize:12.5,color:T.text,background:T.surface,fontFamily:"inherit",outline:"none"}}/>
        </div>
        <div style={{marginBottom:14}}>
          <UploadZone id="wfu" label="Drop weekly aggregate .xlsx" hint="LinkedIn aggregate analytics export" onFiles={previewWeekly}/>
          {wFiles.length>0&&<div style={{marginTop:8,display:"flex",flexWrap:"wrap",gap:5}}>{wFiles.map((f,i)=><span key={i} style={{fontSize:10.5,background:T.blueL,color:T.blue,padding:"2px 9px",borderRadius:20,fontWeight:600}}>{f.name.slice(0,28)}</span>)}</div>}
        </div>
        {wPreview?.length>0&&<div style={{background:T.s2,borderRadius:8,padding:"12px 14px",marginBottom:14,fontSize:12}}>
          <div style={{fontWeight:700,color:T.text,marginBottom:6}}>Parsed: {wPreview.length} days of data</div>
          <div style={{color:T.sec}}>Total impressions: <strong>{fn(wPreview.reduce((s,d)=>s+d.impressions,0))}</strong></div>
          <div style={{color:T.sec}}>Peak day: <strong>{wPreview.reduce((a,b)=>b.impressions>a.impressions?b:a).date} — {fn(Math.max(...wPreview.map(d=>d.impressions)))}</strong></div>
        </div>}
        <Btn onClick={saveWeekly} disabled={proc||!wFiles.length} label={proc?"Saving…":"Save Weekly Data"}/>
      </Card>
      <Card style={{padding:"22px"}}>
        <SHdr t="Saved Weekly Reports" s={`${weekly.length} weeks`}/>
        <div style={{maxHeight:420,overflowY:"auto",display:"flex",flexDirection:"column",gap:8}}>
          {!weekly.length&&<div style={{color:T.muted,fontSize:13,textAlign:"center",padding:"30px 0"}}>No weekly data yet</div>}
          {[...weekly].reverse().map((w,i)=><div key={i} style={{background:T.s2,borderRadius:9,padding:"11px 13px",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <div style={{fontSize:12,fontWeight:700,color:T.text,marginBottom:3}}>{w.weekLabel||"—"}</div>
              <div style={{fontSize:11,color:T.muted}}>{w.dayData.length} days · {fn(w.dayData.reduce((s,d)=>s+d.impressions,0))} impressions</div>
            </div>
            <button onClick={()=>onDelete("weekly",w.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16}}>×</button>
          </div>)}
        </div>
      </Card>
    </div>}

    {/* FOLLOWER REPORTS */}
    {sub==="follower"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
      <Card style={{padding:"22px"}}>
        <SHdr t="Upload Follower Report" s="Monthly follower demographics export"/>
        <div style={{marginBottom:12}}>
          <label style={{fontSize:11.5,fontWeight:700,color:T.sec,display:"block",marginBottom:5}}>Month</label>
          <input type="month" value={fMonth} onChange={e=>setFMonth(e.target.value)} style={{width:"100%",border:`1px solid ${T.border}`,borderRadius:7,padding:"9px 12px",fontSize:12.5,color:T.text,background:T.surface,fontFamily:"inherit",outline:"none"}}/>
        </div>
        <div style={{marginBottom:14}}>
          <UploadZone id="ffu" label="Drop follower .xlsx" hint="LinkedIn follower demographics export" onFiles={previewFollower}/>
          {fFiles.length>0&&<div style={{marginTop:8,display:"flex",flexWrap:"wrap",gap:5}}>{fFiles.map((f,i)=><span key={i} style={{fontSize:10.5,background:T.blueL,color:T.blue,padding:"2px 9px",borderRadius:20,fontWeight:600}}>{f.name.slice(0,28)}</span>)}</div>}
        </div>
        {fPreview&&<div style={{background:T.s2,borderRadius:8,padding:"12px 14px",marginBottom:14,fontSize:12}}>
          <div style={{fontWeight:700,color:T.text,marginBottom:5}}>Demographic categories found:</div>
          {Object.keys(fPreview).length?Object.keys(fPreview).map(k=><div key={k} style={{color:T.sec}}>✓ {k} ({fPreview[k].length} entries)</div>):<div style={{color:T.amber}}>⚠ No demographic data detected</div>}
        </div>}
        <Btn onClick={saveFollower} disabled={proc||!fFiles.length} label={proc?"Saving…":"Save Follower Data"}/>
      </Card>
      <Card style={{padding:"22px"}}>
        <SHdr t="Saved Follower Reports" s={`${followers.length} months`}/>
        <div style={{maxHeight:420,overflowY:"auto",display:"flex",flexDirection:"column",gap:8}}>
          {!followers.length&&<div style={{color:T.muted,fontSize:13,textAlign:"center",padding:"30px 0"}}>No follower data yet</div>}
          {[...followers].reverse().map((f,i)=><div key={i} style={{background:T.s2,borderRadius:9,padding:"11px 13px",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <div style={{fontSize:12,fontWeight:700,color:T.text,marginBottom:3}}>{f.month}</div>
              <div style={{fontSize:11,color:T.muted}}>{Object.keys(f.demographics||{}).join(", ")||"No data"}</div>
            </div>
            <button onClick={()=>onDelete("follower",f.id)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16}}>×</button>
          </div>)}
        </div>
      </Card>
    </div>}
  </div>;
}

/* ── PROGRESS ────────────────────────────────────────────────────────────── */
function Progress({posts,weekly,followers}){
  const postTrend=[...posts].sort((a,b)=>a.date>b.date?1:-1).map((p,i)=>({label:p.date?.slice(5)||`P${i+1}`,impressions:p.totals?.impressions||0,engRate:engRateFrom(p.totals,[]),comments:p.totals?.comments||0,saves:p.totals?.saves||0,followers:p.totals?.newFollowers||0}));
  const lf=[...followers].sort((a,b)=>a.month>b.month?-1:1)[0];
  return <div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Impressions per Post"/>
        {postTrend.length?<ResponsiveContainer width="100%" height={190}><BarChart data={postTrend} margin={{top:4,right:8,left:-24,bottom:0}} barSize={28}><CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/><XAxis dataKey="label" tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false}/><YAxis tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false} tickFormatter={fn}/><Tooltip content={<CTip/>}/><Bar dataKey="impressions" name="Impressions" fill={T.blue} radius={[4,4,0,0]}/></BarChart></ResponsiveContainer>:<div style={{height:190,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted,fontSize:13}}>No data yet</div>}
      </Card>
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Engagement Rate" s={`Avg: ${BENCH.engRate.avg}% · Good: ${BENCH.engRate.good}%`}/>
        {postTrend.length?<ResponsiveContainer width="100%" height={190}><LineChart data={postTrend} margin={{top:4,right:30,left:-24,bottom:0}}><CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/><XAxis dataKey="label" tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false}/><YAxis tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false} tickFormatter={v=>v+"%"}/><Tooltip content={<CTip/>}/><ReferenceLine y={BENCH.engRate.avg} stroke={T.amber} strokeDasharray="4 3" label={{value:"Avg",fill:T.amber,fontSize:9,position:"right"}}/><ReferenceLine y={BENCH.engRate.good} stroke={T.green} strokeDasharray="4 3" label={{value:"Good",fill:T.green,fontSize:9,position:"right"}}/><Line type="monotone" dataKey="engRate" name="Eng Rate %" stroke={T.gold} strokeWidth={2.5} dot={{r:5,fill:T.gold,strokeWidth:0}}/></LineChart></ResponsiveContainer>:<div style={{height:190,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted,fontSize:13}}>No data yet</div>}
      </Card>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Comments & Saves" s="Quality engagement over time"/>
        {postTrend.length?<ResponsiveContainer width="100%" height={180}><LineChart data={postTrend} margin={{top:4,right:8,left:-24,bottom:0}}><CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false}/><XAxis dataKey="label" tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false}/><YAxis tick={{fontSize:10,fill:T.muted}} axisLine={false} tickLine={false}/><Tooltip content={<CTip/>}/><Line type="monotone" dataKey="comments" name="Comments" stroke={T.green} strokeWidth={2.5} dot={{r:4,fill:T.green,strokeWidth:0}}/><Line type="monotone" dataKey="saves" name="Saves" stroke={T.gold} strokeWidth={2} dot={{r:4,fill:T.gold,strokeWidth:0}} strokeDasharray="5 3"/><Line type="monotone" dataKey="followers" name="New Followers" stroke={T.purple} strokeWidth={2} dot={{r:3,fill:T.purple,strokeWidth:0}}/></LineChart></ResponsiveContainer>:<div style={{height:180,display:"flex",alignItems:"center",justifyContent:"center",color:T.muted,fontSize:13}}>No data yet</div>}
      </Card>
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="Latest Audience" s={lf?lf.month:"Upload a follower report"}/>
        {lf&&Object.keys(lf.demographics||{}).length?Object.entries(lf.demographics).slice(0,2).map(([k,vals])=><div key={k} style={{marginBottom:14}}>
          <div style={{fontSize:10.5,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:7}}>{k}</div>
          {vals.slice(0,4).map((v,i)=><div key={i} style={{marginBottom:7}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}><span style={{fontSize:11.5,color:T.sec}}>{v.name}</span><span style={{fontSize:11,color:T.muted,fontWeight:600}}>{v.pct}%</span></div>
            <div style={{background:T.s3,borderRadius:4,height:4}}><div style={{width:`${v.pct}%`,background:T.blue,height:"100%",borderRadius:4}}/></div>
          </div>)}
        </div>):<div style={{color:T.muted,fontSize:13,textAlign:"center",padding:"30px 0"}}>No follower data yet</div>}
      </Card>
    </div>
  </div>;
}

/* ── CONTENT PLAN ────────────────────────────────────────────────────────── */
// Maps a calendar entry to the sentiment domain most relevant to it
function sentimentDomainFor(entry) {
  const p = entry.pillar?.toLowerCase()||"";
  const c = entry.cat?.toLowerCase()||"";
  if(p.includes("campaign")||p.includes("contrarian")) return "advertising";
  if(p.includes("consulting")) return "consulting";
  if(p.includes("product")||c==="consulting") return "productMarketing";
  return "brandManagement";
}

function ContentPlan({posts,mods,onSaveMods,sentiment}){
  const [aiRec,setAiRec]=useState("");
  const [aiLoading,setAiLoading]=useState(false);
  const [pendingMods,setPendingMods]=useState([]);
  const [hashOpen,setHashOpen]=useState(null);
  const hasSentiment=!!(sentiment?.domains);

  // Build published set: match by date (primary) or by topic (fuzzy secondary)
  const pubDates=new Set(posts.map(p=>String(p.date).trim()));
  const pubTopics=new Set(posts.map(p=>String(p.topic||"").toLowerCase().trim()));

  const getStatus=c=>{
    const byDate=pubDates.has(c.date);
    const byTopic=pubTopics.has(c.topic.toLowerCase().trim());
    if((byDate||byTopic)&&c.date<=TODAY) return "published";
    if(c.date===TODAY) return "today";
    if(c.date<TODAY&&!byDate&&!byTopic) return "overdue";
    return "upcoming";
  };

  const nextPost=CAL.find(c=>{const s=getStatus(c);return s==="today"||s==="overdue";})||CAL.find(c=>getStatus(c)==="upcoming");
  const pubCount=CAL.filter(c=>getStatus(c)==="published").length;
  const overdueCount=CAL.filter(c=>getStatus(c)==="overdue").length;

  // Apply mod overrides to calendar
  const getModForDate=date=>mods.find(m=>m.date===date);
  const calWithMods=CAL.map(c=>{
    const mod=getModForDate(c.date);
    if(!mod) return c;
    return {...c,topic:mod.newTopic||c.topic,hashtags:mod.newHashtags||c.hashtags,_modNote:mod.note,_modPriority:mod.priority};
  });

  const generateRec=async()=>{
    setAiLoading(true);setAiRec("");setPendingMods([]);
    const pubSummary=posts.map(p=>`${p.date} | ${p.topic} | Impr:${p.totals?.impressions||0} | Eng:${engRateFrom(p.totals,[]).toFixed(1)}% | Comm:${p.totals?.comments||0} | Saves:${p.totals?.saves||0}`).join("\n")||"No posts uploaded yet.";
    const upcoming=calWithMods.filter(c=>getStatus(c)!=="published").slice(0,8).map(c=>`${c.date}(${c.day}) | ${c.type} | ${c.cat} | ${c.pillar} | ${c.topic}`).join("\n");

    // Build sentiment block
    const sentBlock = hasSentiment ? `
LIVE SENTIMENT SIGNALS (generated ${sentiment.generatedDate}):
Brand Management: ${sentiment.domains.brandManagement?.sentiment?.toUpperCase()} — ${sentiment.domains.brandManagement?.dominantTheme||""}
  Hot debate: ${sentiment.domains.brandManagement?.hotDebate||"none"}
Advertising: ${sentiment.domains.advertising?.sentiment?.toUpperCase()} — ${sentiment.domains.advertising?.dominantTheme||""}
  Hot debate: ${sentiment.domains.advertising?.hotDebate||"none"}
Consulting: ${sentiment.domains.consulting?.sentiment?.toUpperCase()} — ${sentiment.domains.consulting?.dominantTheme||""}
  Hot debate: ${sentiment.domains.consulting?.hotDebate||"none"}
Product Marketing: ${sentiment.domains.productMarketing?.sentiment?.toUpperCase()} — ${sentiment.domains.productMarketing?.dominantTheme||""}
  Hot debate: ${sentiment.domains.productMarketing?.hotDebate||"none"}

Cross-domain pattern: ${sentiment.crossDomainInsight||""}
This week's content opportunity: ${sentiment.contentOpportunity||""}
Audience signal: ${sentiment.audienceSignal||""}` : "No sentiment data available — recommendations based on performance data only.";

    const prompt=`You are a LinkedIn content strategist for Sidharth Marri, brand strategist and marketing consultant at IIM Kozhikode. He posts Mon (Deep, 300-450w) + Thu (Sharp, 80-150w). ~2,450 connections.

PUBLISHED POSTS:
${pubSummary}

UPCOMING IN CALENDAR:
${upcoming}

TODAY: ${TODAY}
${overdueCount>0?`OVERDUE: ${overdueCount} posts missed their schedule.`:""}

${sentBlock}

Your job is to connect the sentiment signals to the content calendar and produce a sentiment-informed content recommendation.

**Next post to publish**
Exact topic, date, type. Cite which specific sentiment signal makes this the right post to publish RIGHT NOW.

**Why post this now (not next week)**
Connect the timing to the live discourse. What is being discussed in the relevant domain that this post speaks to? Be specific about the sentiment signal.

**Angle adjustment**
What specific change to the hook, framing, or angle of this post makes it more relevant to current discourse? Give a concrete before/after rewrite of the opening line if needed.

**Best time to post**
Day and time. One sentence.

**Sentiment-to-calendar mapping**
For each of the next 4 upcoming posts, rate its relevance to current sentiment on a scale of High/Medium/Low and say why in one line. Format as: [date] [topic snippet] → [High/Medium/Low]: [reason]

**Highest priority post in the next 8**
Which one is most time-sensitive given the live discourse? Why in 2 sentences.

**CALENDAR MODIFICATIONS (JSON)**
Output ONLY valid JSON after this header. Each mod must include a sentimentReason field explaining which signal drives it:
\`\`\`json
[
  {"date":"YYYY-MM-DD","type":"note","value":"Specific angle or framing note","sentimentReason":"Which sentiment signal drives this and why"},
  {"date":"YYYY-MM-DD","type":"priority","value":"high","sentimentReason":"Why current discourse makes this urgent"},
  {"date":"YYYY-MM-DD","type":"topic","value":"Revised topic if current discourse warrants a change","sentimentReason":"What is happening in the domain that motivates this change"}
]
\`\`\`

Every point must be tied to either a specific performance number or a specific sentiment signal. No generic advice.`;

    try{
      const res=await fetch("/api/claude",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1800,messages:[{role:"user",content:prompt}]})});
      const data=await res.json();
      const text=data.content?.find(c=>c.type==="text")?.text||"Could not generate.";
      setAiRec(text);
      const jsonMatch=text.match(/```json\s*([\s\S]*?)```/);
      if(jsonMatch){
        try{const parsed=JSON.parse(jsonMatch[1].trim());setPendingMods(parsed.map(m=>({...m,pending:true})));}catch{}
      }
    }catch{setAiRec("Failed to connect. Try again.");}
    setAiLoading(false);
  };

  const applyMods=async()=>{
    const newMods=[...mods.filter(m=>!pendingMods.find(p=>p.date===m.date)),...pendingMods.map(m=>({date:m.date,note:m.type==="note"?m.value:undefined,priority:m.type==="priority"?m.value:undefined,newTopic:m.type==="topic"?m.value:undefined,sentimentReason:m.sentimentReason||undefined,appliedAt:new Date().toISOString()}))];
    await onSaveMods(newMods);
    setPendingMods([]);
    setAiRec(prev=>prev.replace(/```json[\s\S]*?```/,""));
  };

  const clearMod=async date=>{
    const next=mods.filter(m=>m.date!==date);
    await onSaveMods(next);
  };

  function fmtRec(text){
    const withoutJson=text.replace(/```json[\s\S]*?```/,"").trim();
    return withoutJson.split("\n").map((line,i)=>{
      if(line.startsWith("**")&&line.endsWith("**")) return <div key={i} style={{fontWeight:700,color:T.text,fontSize:13.5,marginTop:18,marginBottom:5}}>{line.replace(/\*\*/g,"")}</div>;
      if(!line.trim()) return <div key={i} style={{height:4}}/>;
      return <div key={i} style={{fontSize:13.5,color:T.sec,lineHeight:1.7,marginBottom:3}}>{line.replace(/\*\*/g,"")}</div>;
    });
  }

  return <div>
    {/* Summary KPIs */}
    <div style={{display:"flex",gap:12,marginBottom:20,flexWrap:"wrap"}}>
      {[["Total Posts","34","in 17 weeks",T.blue],["Published",pubCount,`of ${CAL.filter(c=>c.date<=TODAY).length} due`,T.green],["Overdue",overdueCount,"missed schedule",overdueCount>0?T.red:T.muted],["Upcoming",CAL.filter(c=>c.date>TODAY).length,"remaining",T.gold],["Mods Applied",mods.length,"calendar updates",T.purple]].map(([l,v,s,c])=><KPI key={l} label={l} value={String(v)} sub={s} color={c}/>)}
    </div>

    <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16,marginBottom:16}}>
      {/* Calendar table */}
      <Card style={{padding:"20px 22px"}}>
        <SHdr t="17-Week Content Calendar" s={hasSentiment?"Live status · sentiment relevance · hashtags":"Live status · hashtags · generate sentiment report for relevance signals"}/>
        <div style={{overflowX:"auto",maxHeight:500,overflowY:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead style={{position:"sticky",top:0,background:T.surface,zIndex:1}}>
              <tr style={{borderBottom:`2px solid ${T.border}`}}>
                {["Wk","Date","Type","Topic","Category","Status",hasSentiment?"Pulse":"","#"].filter(Boolean).map(h=><th key={h} style={{textAlign:"left",padding:"8px 9px",color:T.muted,fontWeight:700,fontSize:10.5,textTransform:"uppercase",letterSpacing:"0.06em",whiteSpace:"nowrap"}}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {calWithMods.map((c,i)=>{
                const st=getStatus(c);
                const isNext=nextPost?.date===c.date;
                const mod=getModForDate(c.date);
                const rowBg=isNext?T.blueL:st==="overdue"?T.redL:st==="published"?T.greenL:i%2?"transparent":T.s2;
                // Sentiment relevance for this post
                const domKey=sentimentDomainFor(c);
                const domSent=sentiment?.domains?.[domKey];
                const sentHot=domSent&&["heated","critical","mixed"].includes(domSent.sentiment?.toLowerCase());
                const sentPos=domSent&&domSent.sentiment?.toLowerCase()==="positive";
                const sentPulse=domSent?{bg:sentHot?T.redL:sentPos?T.greenL:T.blueL,color:sentHot?T.red:sentPos?T.green:T.blue,label:sentHot?"🔥 Hot":sentPos?"↑ +ve":"● "+domSent.sentiment}:null;
                return <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:rowBg}}>
                  <td style={{padding:"9px 9px",color:T.muted,fontWeight:600}}>{c.w}</td>
                  <td style={{padding:"9px 9px",color:T.muted,whiteSpace:"nowrap"}}>{c.date.slice(5).replace("-","/")} <span style={{fontSize:9.5}}>{c.day}</span></td>
                  <td style={{padding:"9px 9px"}}><span style={{fontSize:10,fontWeight:700,color:c.type==="Deep"?T.blue:T.gold,background:c.type==="Deep"?T.blueL:T.goldL,padding:"2px 7px",borderRadius:20}}>{c.type}</span></td>
                  <td style={{padding:"9px 9px",maxWidth:220}}>
                    <div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontWeight:isNext?700:500,color:isNext?T.blue:T.text}}>{c.topic}</div>
                    {mod?.note&&<div style={{fontSize:10,color:T.purple,marginTop:2}}>💡 {mod.note}</div>}
                    {mod?.sentimentReason&&<div style={{fontSize:10,color:T.green,marginTop:2}}>🔍 {mod.sentimentReason.slice(0,60)}{mod.sentimentReason.length>60?"…":""}</div>}
                    {mod?.newTopic&&<div style={{fontSize:10,color:T.amber,marginTop:2}}>✏ {mod.newTopic.slice(0,50)}</div>}
                    {mod?.priority==="high"&&<div style={{fontSize:10,color:T.red,fontWeight:700,marginTop:2}}>⭐ High priority</div>}
                  </td>
                  <td style={{padding:"9px 9px"}}><Tag cat={c.cat}/></td>
                  <td style={{padding:"9px 9px"}}><StatusBadge status={st}/></td>
                  {hasSentiment&&<td style={{padding:"9px 9px"}}>
                    {sentPulse&&st!=="published"&&<span style={{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20,background:sentPulse.bg,color:sentPulse.color,whiteSpace:"nowrap"}}>{sentPulse.label}</span>}
                    {st==="published"&&<span style={{fontSize:10,color:T.muted}}>—</span>}
                  </td>}
                  <td style={{padding:"9px 9px",position:"relative"}}>
                    <button onClick={()=>setHashOpen(hashOpen===c.date?null:c.date)} style={{background:"none",border:`1px solid ${T.border}`,borderRadius:6,padding:"3px 8px",fontSize:10.5,cursor:"pointer",color:T.blue,fontFamily:"inherit"}}>#</button>
                    {mod&&<button onClick={()=>clearMod(c.date)} style={{background:"none",border:"none",cursor:"pointer",fontSize:10.5,color:T.muted,marginLeft:3}}>✕</button>}
                    {hashOpen===c.date&&<div style={{position:"absolute",right:0,background:T.surface,border:`1px solid ${T.border}`,borderRadius:8,padding:"10px 12px",zIndex:50,fontSize:11,color:T.sec,width:260,boxShadow:T.shM,lineHeight:1.9,marginTop:4}}>
                      {c.hashtags}
                      <button onClick={e=>{e.stopPropagation();navigator.clipboard?.writeText(c.hashtags);}} style={{display:"block",marginTop:6,fontSize:10.5,color:T.blue,background:"none",border:`1px solid ${T.blue}`,borderRadius:5,padding:"3px 8px",cursor:"pointer",fontFamily:"inherit"}}>Copy</button>
                    </div>}
                  </td>
                </tr>;
              })}
            </tbody>
          </table>
        </div>
      </Card>

      {/* AI panel */}
      <div style={{display:"flex",flexDirection:"column",gap:14}}>

        {/* Sentiment pulse widget */}
        {hasSentiment&&<Card style={{padding:"16px 18px"}}>
          <div style={{fontSize:10,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:10}}>🔍 Live Sentiment Pulse</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
            {DOMAINS.map(d=>{
              const dom=sentiment.domains?.[d.key];
              if(!dom) return null;
              const sc=SENT_COLOR(dom.sentiment);
              const sb=SENT_BG(dom.sentiment);
              return <div key={d.key} style={{background:sb,borderRadius:8,padding:"8px 10px",border:`1px solid ${sc}22`}}>
                <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:3}}>
                  <span style={{fontSize:13}}>{d.icon}</span>
                  <span style={{fontSize:10.5,fontWeight:700,color:T.text}}>{d.label.split(" ")[0]}</span>
                </div>
                <span style={{fontSize:10,fontWeight:700,color:sc,textTransform:"capitalize"}}>{dom.sentiment}</span>
                <div style={{fontSize:10,color:T.muted,marginTop:2,lineHeight:1.4,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical"}}>{dom.dominantTheme}</div>
              </div>;
            })}
          </div>
          <div style={{marginTop:10,fontSize:11,color:T.muted,lineHeight:1.5,fontStyle:"italic"}}>
            {sentiment.contentOpportunity?.slice(0,120)}{(sentiment.contentOpportunity?.length||0)>120?"…":""}
          </div>
        </Card>}

        {/* Next up card */}
        {nextPost&&<Card style={{padding:"16px 18px",borderLeft:`3px solid ${T.blue}`}}>
          <div style={{fontSize:10,fontWeight:700,color:T.blue,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>Next Up</div>
          <div style={{fontSize:13,fontWeight:700,color:T.text,lineHeight:1.4,marginBottom:6}}>{nextPost.topic}</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
            <Tag cat={nextPost.cat}/>
            <Pill label={nextPost.type} color={nextPost.type==="Deep"?T.blue:T.gold} bg={nextPost.type==="Deep"?T.blueL:T.goldL}/>
            <StatusBadge status={getStatus(nextPost)}/>
          </div>
          <div style={{fontSize:11,color:T.muted,lineHeight:1.8}}>
            <div>📅 {nextPost.date} ({nextPost.day})</div>
            <div>🏷 {nextPost.pillar}</div>
          </div>
          <div style={{marginTop:8,padding:"7px 10px",background:T.s2,borderRadius:7,fontSize:10.5,color:T.sec,lineHeight:1.7,fontStyle:"italic"}}>{nextPost.hashtags}</div>
        </Card>}

        {/* AI planner card */}
        <Card style={{padding:"16px 18px",flex:1}}>
          <div style={{marginBottom:12}}>
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:13.5,color:T.text}}>AI Content Planner</div>
            <div style={{fontSize:11.5,color:T.muted,marginTop:2}}>
              {hasSentiment
                ? <span>🔍 <strong style={{color:T.green}}>Sentiment-informed</strong> · {sentiment.generatedDate}</span>
                : <span style={{color:T.amber}}>⚠ No sentiment data — <button onClick={()=>{}} style={{background:"none",border:"none",color:T.blue,cursor:"pointer",fontSize:11.5,fontFamily:"inherit",padding:0,textDecoration:"underline"}}>generate report first</button> for best results</span>}
            </div>
          </div>

          <button onClick={generateRec} disabled={aiLoading} style={{width:"100%",background:aiLoading?T.s3:hasSentiment?T.green:T.blue,color:aiLoading?T.muted:"#fff",border:"none",borderRadius:8,padding:"11px",fontSize:13,fontWeight:700,cursor:aiLoading?"default":"pointer",fontFamily:"inherit",marginBottom:aiRec?14:0,transition:"background 0.15s"}}>
            {aiLoading?"Analysing…":hasSentiment?"What to post? (Sentiment-Informed)":"What should I post next?"}
          </button>

          {aiRec&&!aiLoading&&<div style={{maxHeight:420,overflowY:"auto"}}>
            <div style={{background:T.s2,borderRadius:9,padding:"14px 16px",marginBottom:pendingMods.length?12:0,fontSize:13}}>
              {aiRec.replace(/```json[\s\S]*?```/,"").trim().split("\n").map((line,i)=>{
                if(line.startsWith("**")&&line.endsWith("**")) return <div key={i} style={{fontWeight:700,color:T.text,fontSize:13.5,marginTop:16,marginBottom:5}}>{line.replace(/\*\*/g,"")}</div>;
                if(!line.trim()) return <div key={i} style={{height:4}}/>;
                // Highlight sentiment references
                const hasSentRef=line.includes("Heated")||line.includes("heated")||line.includes("Hot")||line.includes("sentiment")||line.includes("🔍")||line.includes("🔥");
                return <div key={i} style={{fontSize:13,color:hasSentRef?T.green:T.sec,lineHeight:1.7,marginBottom:3,fontWeight:hasSentRef?600:400}}>{line.replace(/\*\*/g,"")}</div>;
              })}
            </div>

            {/* Pending mods */}
            {pendingMods.length>0&&<div style={{background:T.purpleL,border:`1px solid ${T.purple}`,borderRadius:9,padding:"14px 16px"}}>
              <div style={{fontWeight:700,color:T.purple,fontSize:12.5,marginBottom:10}}>
                📋 {pendingMods.length} sentiment-informed calendar update{pendingMods.length>1?"s":""} ready:
              </div>
              {pendingMods.map((m,i)=><div key={i} style={{background:T.surface,borderRadius:8,padding:"10px 12px",marginBottom:8}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                  <span style={{fontWeight:700,color:T.text,fontSize:12}}>{m.date}</span>
                  <span style={{fontSize:10,fontWeight:700,textTransform:"uppercase",background:T.s2,color:T.muted,padding:"1px 6px",borderRadius:20}}>{m.type}</span>
                </div>
                <div style={{fontSize:12.5,color:T.sec,marginBottom:m.sentimentReason?5:0,lineHeight:1.5}}>{m.value}</div>
                {m.sentimentReason&&<div style={{fontSize:11,color:T.green,lineHeight:1.4,background:T.greenL,padding:"5px 8px",borderRadius:6}}>
                  🔍 {m.sentimentReason}
                </div>}
              </div>)}
              <button onClick={applyMods} style={{width:"100%",background:T.purple,color:"#fff",border:"none",borderRadius:8,padding:"10px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",marginTop:4}}>
                Apply All to Calendar
              </button>
              <button onClick={()=>setPendingMods([])} style={{width:"100%",background:"none",border:`1px solid ${T.border}`,borderRadius:8,padding:"8px",fontSize:12,fontWeight:500,cursor:"pointer",fontFamily:"inherit",marginTop:6,color:T.muted}}>
                Dismiss
              </button>
            </div>}
          </div>}
        </Card>
      </div>
    </div>
  </div>;
}

/* ── SENTIMENT REPORT ────────────────────────────────────────────────────── */
async function callClaude(systemPrompt, userPrompt, maxTokens=2500) {
  const res = await fetch("/api/claude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: maxTokens,
      system: systemPrompt,
      messages: [{ role: "user", content: userPrompt }],
    }),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`API ${res.status}${body ? ": " + body.slice(0, 200) : ""}`);
  }
  const data = await res.json();
  if (data.error) throw new Error(data.error.message || JSON.stringify(data.error));
  const text = data.content?.find(c => c.type === "text")?.text || "";
  if (!text) throw new Error("Empty response from API");
  return text;
}

const SENT_COLOR = s=>({positive:T.green,neutral:T.blue,mixed:T.amber,critical:T.amber,heated:T.red,cautious:T.gold}[s?.toLowerCase()]||T.muted);
const SENT_BG    = s=>({positive:T.greenL,neutral:T.blueL,mixed:T.amberL,critical:T.amberL,heated:T.redL,cautious:T.goldL}[s?.toLowerCase()]||T.s2);

function SentimentReport({data,onGenerate,loading,posts}){
  const hasData = data&&data.domains;
  const avgEngRate = posts.length?posts.reduce((s,p)=>s+engRateFrom(p.totals,[]),0)/posts.length:0;

  const generate = async () => {
    onGenerate(null, "loading");
    try {
      const systemPrompt = `You are a professional discourse analyst specialising in LinkedIn and business media. Your knowledge covers professional trends, debates, and sentiment in brand management, advertising, consulting, and product marketing up to mid-2026.

You MUST respond with ONLY a valid JSON object. No preamble. No explanation. No markdown fences. No trailing text. Just the raw JSON.`;

      const userPrompt = `Analyse the current state of professional discourse on LinkedIn and in business media (Marketing Week, Campaign, Forbes, HBR, The Drum, McKinsey Insights, etc.) as of ${TODAY} for these four domains:

1. Brand Management / Branding
2. Advertising / Marketing Communications
3. Management Consulting / Strategy Consulting
4. Product Marketing / Go-to-Market Strategy

Context: Sidharth Marri is a brand strategist at IIM Kozhikode (India) building a LinkedIn presence at the intersection of brand strategy and management consulting. His audience is senior professionals in Advertising Services and Consulting, concentrated in Bengaluru, Delhi, Mumbai. His engagement rate is ${avgEngRate.toFixed(1)}%. His best-performing posts have been about brand equity becoming language (Fevicol, verb test), purpose branding authenticity, and India-specific brand strategy.

For each domain provide:
- Overall professional sentiment right now (one of: positive / neutral / mixed / heated / critical / cautious)
- One sentence on why that sentiment is prevailing
- 3 specific trending topics being actively debated in professional circles this week
- One specific hot debate or polarising take circulating right now
- The single dominant theme of the week in one sentence
- A specific content angle Sidharth can own based on what is being discussed

Also provide: a cross-domain pattern, the single best content opportunity for Sidharth this week, and what his specific audience is most engaged with right now.

Return this exact JSON structure:
{
  "generatedDate": "${TODAY}",
  "weekOf": "Week of ${TODAY}",
  "domains": {
    "brandManagement": {
      "sentiment": "positive",
      "sentimentReason": "one sentence",
      "trendingTopics": ["topic1", "topic2", "topic3"],
      "hotDebate": "description of current debate",
      "dominantTheme": "single biggest conversation this week",
      "contentAngle": "specific angle for Sidharth"
    },
    "advertising": {
      "sentiment": "mixed",
      "sentimentReason": "one sentence",
      "trendingTopics": ["topic1", "topic2", "topic3"],
      "hotDebate": "description of current debate",
      "dominantTheme": "single biggest conversation this week",
      "contentAngle": "specific angle for Sidharth"
    },
    "consulting": {
      "sentiment": "neutral",
      "sentimentReason": "one sentence",
      "trendingTopics": ["topic1", "topic2", "topic3"],
      "hotDebate": "description of current debate",
      "dominantTheme": "single biggest conversation this week",
      "contentAngle": "specific angle for Sidharth"
    },
    "productMarketing": {
      "sentiment": "heated",
      "sentimentReason": "one sentence",
      "trendingTopics": ["topic1", "topic2", "topic3"],
      "hotDebate": "description of current debate",
      "dominantTheme": "single biggest conversation this week",
      "contentAngle": "specific angle for Sidharth"
    }
  },
  "crossDomainInsight": "pattern cutting across all four domains",
  "contentOpportunity": "the single best content opportunity for Sidharth this week",
  "audienceSignal": "what his specific follower base is most engaged with right now"
}`;

      const raw = await callClaude(systemPrompt, userPrompt);

      // Strip any accidental markdown fences
      const clean = raw.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();

      // Find the JSON object
      const start = clean.indexOf("{");
      const end   = clean.lastIndexOf("}");
      if (start === -1 || end === -1) throw new Error(`No JSON object found in response. Raw output: "${clean.slice(0, 300)}"`);

      const parsed = JSON.parse(clean.slice(start, end + 1));
      if (!parsed.domains) throw new Error("Response missing 'domains' key. Check raw output.");

      onGenerate(parsed, "done");
    } catch (e) {
      onGenerate({ error: String(e), generatedDate: TODAY }, "error");
    }
  };

  return <div>
    {/* Header bar */}
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:24}}>
      <div>
        <div style={{fontSize:12.5,color:T.muted}}>
          {hasData?`Last generated: ${data.generatedDate} · ${data.weekOf||""}`:loading?"Searching LinkedIn and professional media…":"Not yet generated this week"}
        </div>
        {hasData&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>Brand Management · Advertising · Consulting · Product Marketing</div>}
      </div>
      <button onClick={generate} disabled={loading} style={{background:loading?T.s3:T.blue,color:loading?T.muted:"#fff",border:"none",borderRadius:8,padding:"10px 22px",fontSize:13,fontWeight:700,cursor:loading?"default":"pointer",fontFamily:"inherit",transition:"background 0.15s"}}>
        {loading?"Searching…":hasData?"Refresh Report":"Generate Weekly Report"}
      </button>
    </div>

    {/* Loading state */}
    {loading&&<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        {DOMAINS.map(d=><Card key={d.key} style={{padding:"22px",opacity:0.5}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
            <span style={{fontSize:18}}>{d.icon}</span>
            <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:14,color:T.text}}>{d.label}</div>
          </div>
          <div style={{height:12,background:T.s3,borderRadius:6,marginBottom:8,width:"80%"}}/>
          <div style={{height:12,background:T.s3,borderRadius:6,marginBottom:8,width:"60%"}}/>
          <div style={{height:12,background:T.s3,borderRadius:6,width:"90%"}}/>
        </Card>)}
      </div>
      <div style={{textAlign:"center",padding:"16px",color:T.muted,fontSize:13}}>Searching LinkedIn, Marketing Week, Campaign, Forbes and HBR for this week's signals…</div>
    </div>}

    {/* Error state */}
    {!loading&&data?.error&&<Card style={{padding:"22px",borderLeft:`3px solid ${T.red}`}}>
      <div style={{fontWeight:700,color:T.red,marginBottom:8}}>Search failed</div>
      <div style={{fontSize:13,color:T.sec,fontFamily:"monospace"}}>{data.error}</div>
    </Card>}

    {/* Empty state */}
    {!loading&&!hasData&&!data?.error&&<div style={{textAlign:"center",padding:"70px 20px"}}>
      <div style={{fontSize:42,marginBottom:16}}>🔍</div>
      <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:20,color:T.text,marginBottom:8}}>No report yet</div>
      <div style={{fontSize:14,color:T.muted,maxWidth:420,margin:"0 auto",lineHeight:1.7}}>Click Generate to search LinkedIn and professional media for this week's sentiment across Brand Management, Advertising, Consulting, and Product Marketing.</div>
    </div>}

    {/* Domain cards */}
    {!loading&&hasData&&<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        {DOMAINS.map(d=>{
          const dom = data.domains?.[d.key];
          if(!dom) return null;
          const sc=SENT_COLOR(dom.sentiment);
          const sb=SENT_BG(dom.sentiment);
          return <Card key={d.key} style={{padding:"22px",borderLeft:`3px solid ${sc}`}}>
            {/* Domain header */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:20}}>{d.icon}</span>
                <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:14,color:T.text}}>{d.label}</div>
              </div>
              <div style={{textAlign:"right"}}>
                <span style={{fontSize:11,fontWeight:700,padding:"3px 10px",borderRadius:20,background:sb,color:sc,textTransform:"capitalize"}}>{dom.sentiment}</span>
              </div>
            </div>

            {/* Sentiment reason */}
            {dom.sentimentReason&&<div style={{fontSize:12,color:T.muted,marginBottom:10,fontStyle:"italic",lineHeight:1.5}}>{dom.sentimentReason}</div>}

            {/* Dominant theme */}
            <div style={{background:T.s2,borderRadius:8,padding:"10px 12px",marginBottom:12,fontSize:13,color:T.sec,lineHeight:1.55,fontWeight:500}}>
              "{dom.dominantTheme}"
            </div>

            {/* Trending topics */}
            <div style={{marginBottom:12}}>
              <div style={{fontSize:10,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:7}}>Trending this week</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                {dom.trendingTopics?.map((t,i)=><span key={i} style={{fontSize:11.5,padding:"4px 10px",borderRadius:20,background:d.bg,color:d.color,fontWeight:600,border:`1px solid ${d.color}22`}}>{t}</span>)}
              </div>
            </div>

            {/* Hot debate */}
            {dom.hotDebate&&<div style={{marginBottom:12,padding:"10px 12px",background:T.redL,borderRadius:8,borderLeft:`2px solid ${T.red}`}}>
              <div style={{fontSize:10,fontWeight:700,color:T.red,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:4}}>💬 Hot debate</div>
              <div style={{fontSize:12.5,color:T.sec,lineHeight:1.55}}>{dom.hotDebate}</div>
            </div>}

            {/* Content angle */}
            {dom.contentAngle&&<div style={{padding:"9px 12px",background:T.greenL,borderRadius:8,borderLeft:`2px solid ${T.green}`}}>
              <div style={{fontSize:10,fontWeight:700,color:T.green,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>→ Your angle</div>
              <div style={{fontSize:12.5,color:T.sec,lineHeight:1.5}}>{dom.contentAngle}</div>
            </div>}
          </Card>;
        })}
      </div>

      {/* Cross-domain insight */}
      {data.crossDomainInsight&&<Card style={{padding:"20px 24px",marginBottom:14}}>
        <div style={{fontSize:10.5,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:10}}>Cross-Domain Pattern</div>
        <div style={{fontSize:14,color:T.sec,lineHeight:1.7}}>{data.crossDomainInsight}</div>
      </Card>}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        {/* Content opportunity */}
        {data.contentOpportunity&&<Card style={{padding:"20px 22px",borderLeft:`3px solid ${T.gold}`}}>
          <div style={{fontSize:10.5,fontWeight:700,color:T.gold,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:10}}>🎯 Best Content Opportunity This Week</div>
          <div style={{fontSize:13.5,color:T.text,lineHeight:1.7}}>{data.contentOpportunity}</div>
        </Card>}

        {/* Audience signal */}
        {data.audienceSignal&&<Card style={{padding:"20px 22px",borderLeft:`3px solid ${T.purple}`}}>
          <div style={{fontSize:10.5,fontWeight:700,color:T.purple,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:10}}>👥 Your Audience Signal</div>
          <div style={{fontSize:13.5,color:T.text,lineHeight:1.7}}>{data.audienceSignal}</div>
        </Card>}
      </div>
    </div>}
  </div>;
}

/* ── MAIN ────────────────────────────────────────────────────────────────── */
export default function App(){
  const [tab,setTab]=useState("overview");
  const [posts,setPosts]=useState([]);
  const [weekly,setWeekly]=useState([]);
  const [followers,setFollowers]=useState([]);
  const [mods,setMods]=useState([]);
  const [sentiment,setSentiment]=useState(null);
  const [sentimentStatus,setSentimentStatus]=useState("idle"); // idle|loading|done|error
  const [loading,setLoading]=useState(true);

  useEffect(()=>{
    Promise.all([sg(KEY.posts),sg(KEY.weekly),sg(KEY.followers),sg(KEY.mods),sg(KEY.sentiment)])
      .then(([p,w,f,m,s])=>{setPosts(p);setWeekly(w);setFollowers(f);setMods(m);if(s&&!Array.isArray(s))setSentiment(s);})
      .finally(()=>setLoading(false));
  },[]);

  const savePost=useCallback(async e=>{const n=[...posts.filter(p=>p.id!==e.id),e];setPosts(n);await ss(KEY.posts,n);},[posts]);
  const saveWeekly=useCallback(async e=>{const n=[...weekly.filter(w=>w.id!==e.id),e];setWeekly(n);await ss(KEY.weekly,n);},[weekly]);
  const saveFollower=useCallback(async e=>{const n=[...followers.filter(f=>f.id!==e.id),e];setFollowers(n);await ss(KEY.followers,n);},[followers]);
  const saveMods=useCallback(async m=>{setMods(m);await ss(KEY.mods,m);},[]);
  const handleSentiment=useCallback(async(data,status)=>{
    setSentimentStatus(status);
    if(data&&status==="done"){setSentiment(data);await ss(KEY.sentiment,data);}
    if(status==="error"&&data){setSentiment(data);}
  },[]);
  const delEntry=useCallback(async(type,id)=>{
    if(type==="post"){const n=posts.filter(p=>p.id!==id);setPosts(n);await ss(KEY.posts,n);}
    if(type==="weekly"){const n=weekly.filter(w=>w.id!==id);setWeekly(n);await ss(KEY.weekly,n);}
    if(type==="follower"){const n=followers.filter(f=>f.id!==id);setFollowers(n);await ss(KEY.followers,n);}
  },[posts,weekly,followers]);

  const navTabs=[
    {id:"overview",icon:"📊",label:"Overview"},
    {id:"reports",icon:"📁",label:"Reports"},
    {id:"progress",icon:"📈",label:"Progress"},
    {id:"content",icon:"📅",label:"Content Plan"},
    {id:"sentiment",icon:"🔍",label:"Sentiment"},
  ];

  if(loading) return <div style={{background:T.bg,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"Inter,sans-serif"}}><div style={{textAlign:"center",color:T.muted,fontSize:14}}>Loading your dashboard…</div></div>;

  return <div style={{background:T.bg,minHeight:"100vh",fontFamily:"'Inter',sans-serif",color:T.text}}>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Inter:wght@400;500;600;700&display=swap');*{box-sizing:border-box;margin:0;padding:0;}::-webkit-scrollbar{width:5px;height:5px;}::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px;}`}</style>

    {/* Header */}
    <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
      <div style={{display:"flex",alignItems:"center",gap:11,padding:"13px 0"}}>
        <div style={{width:32,height:32,background:T.blue,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <span style={{color:"#fff",fontSize:16,fontWeight:800,fontFamily:"'Syne',sans-serif"}}>in</span>
        </div>
        <div>
          <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:16,color:T.text}}>LinkedIn Analytics</div>
          <div style={{fontSize:10,color:T.muted}}>Sidharth Marri · {posts.length} posts · {weekly.length} weeks · {mods.length} mods · {sentiment?.generatedDate?`Sentiment ${sentiment.generatedDate}`:"No sentiment report yet"}</div>
        </div>
      </div>
      <nav style={{display:"flex",gap:2}}>
        {navTabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"10px 16px",background:"none",border:"none",borderBottom:`2px solid ${tab===t.id?T.blue:"transparent"}`,color:tab===t.id?T.blue:T.sec,fontWeight:tab===t.id?700:500,fontSize:13,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:5}}><span>{t.icon}</span>{t.label}</button>)}
      </nav>
    </div>

    <div style={{maxWidth:1160,margin:"0 auto",padding:"26px 22px 80px"}}>
      <div style={{marginBottom:20}}>
        <div style={{fontFamily:"'Syne',sans-serif",fontWeight:800,fontSize:21,color:T.text,letterSpacing:"-0.02em"}}>{navTabs.find(t=>t.id===tab)?.label}</div>
        <div style={{fontSize:12,color:T.muted,marginTop:2}}>{{overview:"Master tracker — all performance data and benchmarks at a glance",reports:"Upload new exports — each type has its own flow and data preview",progress:"Week-over-week trends and audience evolution",content:"17-week calendar with live status, hashtags, and AI recommendations",sentiment:"Weekly sentiment pulse across Brand Management, Advertising, Consulting, and Product Marketing"}[tab]}</div>
      </div>
      {tab==="overview"&&<Overview posts={posts} weekly={weekly} followers={followers}/>}
      {tab==="reports"&&<Reports posts={posts} weekly={weekly} followers={followers} onSavePost={savePost} onSaveWeekly={saveWeekly} onSaveFollower={saveFollower} onDelete={delEntry}/>}
      {tab==="progress"&&<Progress posts={posts} weekly={weekly} followers={followers}/>}
      {tab==="content"&&<ContentPlan posts={posts} mods={mods} onSaveMods={saveMods} sentiment={sentiment}/>}
      {tab==="sentiment"&&<SentimentReport data={sentiment} loading={sentimentStatus==="loading"} onGenerate={handleSentiment} posts={posts}/>}
    </div>
  </div>;
}
