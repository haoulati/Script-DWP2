// ==UserScript==
// @name         FCLM Wash DWP2
// @namespace    https://w.amazon.com/bin/view/Users/haoulati/scripts/fclm-wash/
// @version      1.3.5
// @description  Wash Sheet with QS Plan (auto-load from SharePoint + Wiki fallback), unplanned log detection and Slack alerts
// @author       haoulati
// @match        https://fclm-portal.amazon.com/*
// @updateURL    https://raw.githubusercontent.com/haoulati/Script-DWP2/main/fclm-wash.user.js
// @downloadURL  https://raw.githubusercontent.com/haoulati/Script-DWP2/main/fclm-wash.user.js
// @run-at       document-start
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @connect      fclm-portal.amazon.com
// @connect      hooks.slack.com
// @connect      amazonfra.sharepoint.com
// @connect      w.amazon.com
// ==/UserScript==

// Hide page while loading
(function() {
  if (!window.location.pathname.startsWith('/ppa/wash')) return;
  const s = document.createElement('style');
  s.textContent = 'html{visibility:hidden!important}html.wash-ready{visibility:visible!important}body>*:not(#wc):not(script):not(style){display:none!important}';
  (document.head || document.documentElement).appendChild(s);
  document.title = 'Wash Sheet - DWP2';
})();

(function() {
  'use strict';

  // ─── CONFIG ───────────────────────────────────────────────────────────────────
  const SHAREPOINT_BASE = 'https://amazonfra.sharepoint.com/sites/DWP2_Station/Documents partages/1 Area/Projets/FCLM Wash';
  const WIKI_BASE       = 'https://w.amazon.com/bin/download/Users/haoulati/fclm-plan';

  // ─── COLOR THRESHOLDS ─────────────────────────────────────────────────────────
  const dCls  = v => v==null?'': v>=-1&&v<=3?'w-green': v<=5?'w-yellow':'w-red';
  const dBg   = v => v==null?'': v>=-1&&v<=3?'w-bg-green': v<=5?'w-bg-yellow':'w-bg-red';
  const dClr  = v => v==null?'#1565c0': v>=-1&&v<=3?'#2e7d32': v<=5?'#f57f17':'#c62828';
  const dEmoji= v => v==null?'⚪': v>=-1&&v<=3?'🟢': v<=5?'🟡':'🔴';

  // ─── DATE HELPERS ─────────────────────────────────────────────────────────────
  const FR_MONTHS = {
    'jan':'01','fév':'02','mar':'03','avr':'04','mai':'05','juin':'06',
    'juil':'07','août':'08','sep':'09','oct':'10','nov':'11','déc':'12',
    'jan.':'01','feb.':'02','apr.':'04','jun.':'06','jul.':'07',
    'aug.':'08','sep.':'09','oct.':'10','nov.':'11','dec.':'12',
  };
  function parseFrDate(str) {
    if (!str) return null;
    str = str.trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.split(' ')[0];
    const m = str.match(/^([a-zA-Zéûôîàè.]+)\s+(\d{1,2}),?\s+(\d{4})/i);
    if (m) {
      const mon = FR_MONTHS[m[1].toLowerCase().replace('.','').trim()]
               || FR_MONTHS[m[1].toLowerCase().trim()];
      if (mon) return `${m[3]}-${mon}-${m[2].padStart(2,'0')}`;
    }
    const m2 = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m2) return `${m2[3]}-${m2[2].padStart(2,'0')}-${m2[1].padStart(2,'0')}`;
    return null;
  }

  // ─── CONSTANTS ────────────────────────────────────────────────────────────────
  const BASE = 'https://fclm-portal.amazon.com';
  const CATS = ['Sort','Pick Stage','Site Support','UTR','OTR','RTS','CRETS'];

  // ✅ Process à ignorer complètement dans l'affichage
  const IGNORED_PROCESSES = new Set([
    'Tesseract Pick',
    'Z&T Cart Stager',
    'SSD Core Stage',
    'AR Stage',
    'Zancasort Stow Bag Replenishment',
    'Empty Cart Returns',
    // RTS — tous les Reverse sont ignorés, seul RTS Associate compte
    'Reverse Receive',
    'Reverse Stow',
    'Reverse Cycle',
    'Reverse Induct',
    'Reverse Pick Stage',
    'Reverse Indirect',
    'Reverse Problem Solve',
  ]);

  const CSV_MAIN_TO_CAT = {
    'sort':'Sort','pick stage':'Pick Stage','site support':'Site Support',
    'utr':'UTR','otr':'OTR','rts':'RTS',
    'reverse cycle':     'RTS',
    'crets':             'CRETS',
    'non-core support':  'CRETS',
    'off task':          'UTR',
    // ✅ Nouveaux mappings
    'reverse indirect':'RTS','reverse pick stage':'RTS',
    'reverse problem solve':'RTS','cs dsl':'Site Support',
  };

  const PPA_TO_CSV = {
    'ADTA Container Building':'ADTA Stower (100541)',
    'Container Building':'Container Building (100021)',
    'Pick to Buffer':'Pick to Buffer (100440)',
    'Labeler':'ASML Labeler (100796)',
    'ASML Induct Loader':'ASML Induct Loader (100943)',
    'Sort Problem Solve':'Sort Problem Solve (100465)',
    'ASML Pusher':'ASML Pusher (100944)',
    'Auto Divert Straightener':'Auto Divert Straightener (100477)',
    'Inbound Dock W/S':'Inbound Dock W/S (100405)',
    'Induct':'Induct (100020)',
    'Non Con Manual Handling':'Non Con Manual Handling (100570)',
    'Induct Line Loader':'Induct Line Loader',
    'Pusher':'Pusher',
    'ASL Induct Loader':'ASL Induct Loader',
    'ASL Induct':'ASL Induct',
    'ASL Pusher':'ASL Pusher',
    'ASML Induct':'ASML Induct',
    'Tesseract Straightener':'Tesseract Straightener',
    'Zancasort Container Building':'Zancasort Container Building',
    'Diverter':'Diverter',
    'Non-Scan Stow':'Non-Scan Stow',
    'Pick & Stage - Pick':'Pick and Stage (100407)',
    'Non-Scan Pick & Stage':'Non-Scan Pick & Stage (100463)',
    'Zancasort Pick':'Zancasort Pick',
    'Yard Assist and Cart Handling':'Yard Assist and Cart Handling',
    'Pick Stage Dispatch Problem Solve':'Pick Stage Problem Solve (100478)',
    'HR':'HR (100428)',
    'Sustainability':'Sustainability',
    'Safety':'Safety (100429)',
    'Trainer':'Training / Safety School (100436)',
    'Trainee':'Training / Safety School (100436)',
    'Training / Safety School (100436)':'Training / Safety School (100436)',
    'IT Super User':'IT Super User',
    '5S / Non Productive':'5S (100470)',
    'UTR OPS Supervisor / SA':'UTR Supervisor/Shift Assistant (100464)',
    'Inbound Traffic Controller':'Yard Marshall (100410/100941)',
    'Outbound Traffic Controller':'Yard Marshall (100410/100941)',
    'Stow Bag Replenishment':'Stow Bag Replenishment (100462)',
    'Learning Coordinator':'Learning Coordinator',
    'OTR Supervisor / Shift Assistant':'OTR Supervisor / Shift Assistant',
    'OTR Support':'OTR Support Functions (100469)',
    'RTS':'RTS (100412)',
    'RTS Associate':'RTS Associate (100412)',
    'Reverse Receive':'RTS (100412)',
    'Reverse Stow':'RTS (100412)',
    'Reverse Cycle':'RTS (100412)',
    // ✅ Nouveaux mappings RTS
    'Reverse Induct':'RTS (100412)',
    'Reverse Pick Stage':'RTS (100412)',
    'Reverse Indirect':'RTS (100412)',
    'Reverse Problem Solve':'RTS (100412)',
    // ✅ CS DSL
    'CS DSL':'CS DSL',
    'Off Task':'UTR Supervisor/Shift Assistant (100464)',
  };

  const CSV_TO_CAT = {
    'ADTA Stower (100541)':'Sort','Container Building (100021)':'Sort',
    'Pick to Buffer (100440)':'Sort','ASML Labeler (100796)':'Sort',
    'ASML Induct Loader (100943)':'Sort','Sort Problem Solve (100465)':'Sort',
    'ASML Pusher (100944)':'Sort','Auto Divert Straightener (100477)':'Sort',
    'Inbound Dock W/S (100405)':'Sort','Induct (100020)':'Sort',
    'Non Con Manual Handling (100570)':'Sort','Induct Line Loader':'Sort',
    'Pusher':'Sort','ASL Induct Loader':'Sort','ASL Induct':'Sort',
    'ASL Pusher':'Sort','ASML Induct':'Sort','Tesseract Straightener':'Sort',
    'Zancasort Container Building':'Sort','Diverter':'Sort','Non-Scan Stow':'Sort',
    'Pick and Stage (100407)':'Pick Stage','Non-Scan Pick & Stage (100463)':'Pick Stage',
    'Zancasort Pick':'Pick Stage',
    'Yard Assist and Cart Handling':'Pick Stage',
    'Pick Stage Problem Solve (100478)':'Pick Stage',
    'HR (100428)':'Site Support','Sustainability':'Site Support',
    'Safety (100429)':'Site Support',
    'Training / Safety School (100436)':'Site Support',
    'IT Super User':'Site Support',
    'CS DSL':'Site Support',
    '5S (100470)':'UTR','UTR Supervisor/Shift Assistant (100464)':'UTR',
    'Yard Marshall (100410/100941)':'UTR','Stow Bag Replenishment (100462)':'UTR',
    'Learning Coordinator':'UTR',
    'OTR Supervisor / Shift Assistant':'OTR','OTR Support Functions (100469)':'OTR','OTR Support':'OTR',
    'RTS (100412)':'RTS','RTS Associate (100412)':'RTS',
    'Reverse Receive':'RTS','Reverse Stow':'RTS','Reverse Cycle':'RTS',
    'Reverse Induct':'RTS','Reverse Pick Stage':'RTS',
    'Reverse Indirect':'RTS','Reverse Problem Solve':'RTS',
  };

  // ─── STORAGE ──────────────────────────────────────────────────────────────────
  const K = {
    SITE:'fclm_wash_site',SLACK:'fclm_wash_slack',
    PLAN:'fclm_wash_qs_plan',SENT:'fclm_wash_unplanned_sent',
    SUMMARY_SENT:'fclm_wash_summary_sent'
  };
  const ls = {
    get: k=>{try{return JSON.parse(localStorage.getItem(k));}catch{return null;}},
    set: (k,v)=>localStorage.setItem(k,JSON.stringify(v)),
    del: k=>localStorage.removeItem(k)
  };
  function getPlan(site,date)       {return(ls.get(K.PLAN)||{})[`${site}::${date}`]||null;}
  function savePlan(site,date,data) {const s=ls.get(K.PLAN)||{};s[`${site}::${date}`]=data;ls.set(K.PLAN,s);}
  function delPlan(site,date)       {const s=ls.get(K.PLAN)||{};delete s[`${site}::${date}`];ls.set(K.PLAN,s);}
  function getLoadedDates(site)     {return Object.keys(ls.get(K.PLAN)||{}).filter(k=>k.startsWith(site+'::')).map(k=>k.replace(site+'::','')).sort();}
  function getSlack() {
    const saved=ls.get(K.SLACK)||{};
    if(!saved.webhookUrl){
      saved.webhookUrl='https://hooks.slack.com/services/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';
      saved.alertUnplanned=true;saved.summary14h=true;saved.summary23h=true;
    }
    return saved;
  }
  function saveSlack(cfg){ls.set(K.SLACK,cfg);}

  // ─── NUMBER PARSER v1.3.5 ─────────────────────────────────────────────────────
  function parseNum(s) {
    let v=(s||'').toString().trim().replace(/[%]/g,'');
    if(!v)return 0;
    v=v.replace(/\u00a0/g,'').replace(/\s/g,'');
    if(!v)return 0;
    // Nombre simple avec point décimal ex: "7.685" → 7.685
    if(/^\d+\.\d+$/.test(v))return parseFloat(v);
    // Format FR : "203,16" → virgule = décimal
    if(/^\d+,\d+$/.test(v))return parseFloat(v.replace(',','.'));
    // Format EN milliers : "5,537.0" → virgule = milliers, point = décimal
    if(/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(v))return parseFloat(v.replace(/,/g,''));
    // Format EU milliers : "5.537,0" → point = milliers, virgule = décimal
    if(/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(v))return parseFloat(v.replace(/\./g,'').replace(',','.'));
    return parseFloat(v)||0;
  }

  // ─── CSV HELPERS ──────────────────────────────────────────────────────────────
  function detectSeparator(text) {
    const l=text.split('\n')[0];
    const c=(l.match(/,/g)||[]).length,s=(l.match(/;/g)||[]).length,t=(l.match(/\t/g)||[]).length;
    if(t>c&&t>s)return '\t';
    return s>c?';':',';
  }
  function parseCSVLine(line,sep=',') {
    const r=[];let cur='',inQ=false;
    for(const ch of line){
      if(ch==='"')inQ=!inQ;
      else if(ch===sep&&!inQ){r.push(cur);cur='';}
      else cur+=ch;
    }
    r.push(cur);return r;
  }
  function parseQsCsv(text,site) {
    const sep=detectSeparator(text);
    const lines=text.split('\n');
    if(lines.length<2)return{};
    const headers=lines[0].replace(/^\uFEFF/,'').split(sep).map(h=>h.replace(/"/g,'').trim());
    const idx={};headers.forEach((h,i)=>idx[h]=i);
    console.log('[Wash] QS CSV sep:',JSON.stringify(sep),'| Headers:',headers);
    const byDate={};
    for(let i=1;i<lines.length;i++){
      const line=lines[i].trim();if(!line)continue;
      const cols=parseCSVLine(line,sep);
      if(cols.length<headers.length)continue;
      const loc=(cols[idx['location_allocated']]||'').replace(/"/g,'').trim();
      if(site&&loc!==site)continue;
      const rawDate=(cols[idx['reporting_date']]||'').replace(/"/g,'').trim();
      const date=parseFrDate(rawDate);
      if(!date)continue;
      const proc=(cols[idx['Labor Process']]||'').replace(/"/g,'').trim();
      const hrs=parseNum((cols[idx['planned_hours']]||'').replace(/"/g,''));
      const units=parseNum((cols[idx['planned_units']]||'').replace(/"/g,''));
      if(!proc)continue;
      const cat=CSV_TO_CAT[proc]||null;
      if(!byDate[date])byDate[date]={};
      byDate[date][proc]={plannedHours:hrs,plannedUnits:units,category:cat};
    }
    console.log('[Wash] QS Parsed dates:',Object.keys(byDate));
    return byDate;
  }

  // ─── AUTO-LOAD PLAN ───────────────────────────────────────────────────────────
  function fetchUrl(url) {
    return new Promise(resolve=>{
      GM_xmlhttpRequest({
        method:'GET',url,
        onload: r=>resolve({ok:r.status===200,status:r.status,text:r.responseText}),
        onerror:()=>resolve({ok:false,status:0,text:''})
      });
    });
  }
  async function autoLoadPlan(site,dateStr) {
    if(getPlan(site,dateStr)){console.log('[Wash] Plan cached for',dateStr);return{success:true,cached:true,source:'cache'};}
    const spUrl=`${SHAREPOINT_BASE}/FCLM_Plan_${dateStr}.csv`;
    const wikiUrl=`${WIKI_BASE}/FCLM_Plan_${dateStr}.csv`;
    console.log('[Wash] Trying SharePoint:',spUrl);
    const spRes=await fetchUrl(spUrl);
    if(spRes.ok&&spRes.text){
      try{const byDate=parseQsCsv(spRes.text,site);const dates=Object.keys(byDate);if(dates.length){dates.forEach(d=>savePlan(site,d,byDate[d]));return{success:true,source:'sharepoint',dates};}}
      catch(e){console.warn('[Wash] SharePoint parse error:',e.message);}
    }
    console.warn('[Wash] SharePoint failed, trying Wiki...');
    const wikiRes=await fetchUrl(wikiUrl);
    if(wikiRes.ok&&wikiRes.text){
      try{const byDate=parseQsCsv(wikiRes.text,site);const dates=Object.keys(byDate);if(dates.length){dates.forEach(d=>savePlan(site,d,byDate[d]));return{success:true,source:'wiki',dates};}}
      catch(e){console.warn('[Wash] Wiki parse error:',e.message);}
    }
    return{success:false,source:'none',error:spRes.status===404||wikiRes.status===404?`Fichier FCLM_Plan_${dateStr}.csv introuvable.`:`SharePoint (${spRes.status}) et Wiki (${wikiRes.status}) inaccessibles.`};
  }

  // ─── PLAN HELPERS ─────────────────────────────────────────────────────────────
  function getProcPlanHrs(planData,cat,ppaName){
    if(!planData?.[cat]?.processes)return null;
    if(planData[cat].processes[ppaName])return planData[cat].processes[ppaName].plannedHours;
    const csv=PPA_TO_CSV[ppaName];
    if(csv&&planData[cat].processes[csv])return planData[cat].processes[csv].plannedHours;
    // ✅ Yard Marshall fix : Inbound + Outbound Traffic Controller partagent le plan Yard Marshall
    // Ex: plan = 35h → chacun est considéré comme couvert (pas UNPLANNED), delta = actuel - 35h total
    const yardMarshallProcs=['Inbound Traffic Controller','Outbound Traffic Controller'];
    if(cat==='UTR'&&yardMarshallProcs.includes(ppaName)){
      const ymPlan=planData[cat].processes['Yard Marshall (100410/100941)'];
      if(ymPlan)return ymPlan.plannedHours; // on retourne le total — la détection unplanned sera désactivée
    }
    const rtsSubs=['RTS Associate','RTS'];
    if(cat==='RTS'&&rtsSubs.some(r=>ppaName.toLowerCase().includes(r.toLowerCase())||r.toLowerCase().includes(ppaName.toLowerCase()))){
      for(const k of['RTS (100412)','RTS Associate (100412)','RTS']){if(planData[cat].processes[k])return planData[cat].processes[k].plannedHours;}
      const total=Object.values(planData[cat].processes).reduce((s,p)=>s+(p.plannedHours||0),0);
      if(total>0)return total;
    }
    return null;
  }
  function buildPlanByCat(rawPlan){
    const r={};CATS.forEach(c=>r[c]={hours:0,units:0,processes:{}});
    for(const[csvName,data]of Object.entries(rawPlan)){
      const cat=data.category;if(!cat||!r[cat])continue;
      r[cat].hours+=data.plannedHours;r[cat].units+=data.plannedUnits;
      r[cat].processes[csvName]={plannedHours:data.plannedHours,plannedUnits:data.plannedUnits};
    }
    return r;
  }

  // ─── SITE / DATE ──────────────────────────────────────────────────────────────
  function getSite(){const p=new URLSearchParams(window.location.search);const s=p.get('warehouseId')||p.get('site');if(s){ls.set(K.SITE,s);return s;}return ls.get(K.SITE)||null;}
  function setSite(s){ls.set(K.SITE,s);}
  function getDateStr(){const p=new URLSearchParams(window.location.search);const d=p.get('startDateDay')||p.get('startDateIntraday')||p.get('date');if(d)return d.replace(/\//g,'-');const n=new Date();return`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}-${String(n.getDate()).padStart(2,'0')}`;}
  function getTimeRange(){const p=new URLSearchParams(window.location.search);const sh=p.get('startHourIntraday'),eh=p.get('endHourIntraday');if(sh!=null&&eh!=null)return{startH:parseInt(sh),startM:parseInt(p.get('startMinuteIntraday')||'0'),endH:parseInt(eh),endM:parseInt(p.get('endMinuteIntraday')||'0')};return null;}
  function buildIntradayParams(dateStr,tr){
    if(!tr)return'';
    const d=dateStr.replace(/-/g,'/');
    const nd=new Date(dateStr);nd.setDate(nd.getDate()+1);
    const nx=`${nd.getFullYear()}/${String(nd.getMonth()+1).padStart(2,'0')}/${String(nd.getDate()).padStart(2,'0')}`;
    return`&spanType=Intraday&maxIntradayDays=1&startDateDay=${encodeURIComponent(nx)}&startDateWeek=${encodeURIComponent(nx)}&startDateMonth=${encodeURIComponent(d.substring(0,7)+'/01')}&startDateIntraday=${encodeURIComponent(d)}&startHourIntraday=${tr.startH}&startMinuteIntraday=${tr.startM}&endDateIntraday=${encodeURIComponent(d)}&endHourIntraday=${tr.endH}&endMinuteIntraday=${tr.endM}`;
  }

  // ─── FETCH FCLM ───────────────────────────────────────────────────────────────
  function gmGet(url){return new Promise((res,rej)=>GM_xmlhttpRequest({method:'GET',url,onload:r=>res(r),onerror:e=>rej(e)}));}
  async function fetchPpaHtml(site,date,tr){return(await gmGet(`${BASE}/ppa/inspect/node?warehouseId=${site}${buildIntradayParams(date,tr)}`)).responseText;}
  async function fetchPpaCsv(site,date,tr){return(await gmGet(`${BASE}/ppa/inspect/node/csv?warehouseId=${site}&nodeType=DS${buildIntradayParams(date,tr)}`)).responseText;}

  // ─── PPA CSV PARSER v1.3.5 ────────────────────────────────────────────────────
  function parsePpaCsv(csv) {
    const result={};CATS.forEach(c=>result[c]=[]);
    const lines=csv.split('\n');
    if(lines.length<2)return result;
    // ✅ Détecter le séparateur réel du CSV (tab, point-virgule ou virgule)
    const sep=detectSeparator(csv);
    // ✅ Lire les headers pour trouver les bonnes colonnes dynamiquement
    const rawHeaders=lines[0].replace(/^\uFEFF/,'');
    const headers=parseCSVLine(rawHeaders,sep).map(h=>h.replace(/"/g,'').trim());
    // Indices des colonnes clés — noms exacts du CSV FCLM
    const iMain =headers.findIndex(h=>h==='Main Processes'||h==='Main Process'||/main.*process/i.test(h));
    const iProc =headers.findIndex(h=>h==='Labor Process Name'||/labor.*process.*name/i.test(h));
    const iAttr =headers.findIndex(h=>h==='Attributes'||/^attribute/i.test(h));
    const iUnits=headers.findIndex(h=>h==='Actual Units');
    const iHours=headers.findIndex(h=>h==='Actual Hours');
    // Fallback sur les indices fixes FCLM connus si headers non trouvés
    // Structure CSV FCLM : 0=MainProcess, 1=CoreProcess, 2=LaborProcess, 3=Code, 4=Attribute, 5=ActualUnits, 6=ActualQty, 7=ActualHours
    const colMain =iMain >=0?iMain :0;
    const colProc =iProc >=0?iProc :2;
    const colAttr =iAttr >=0?iAttr :4;
    const colUnits=iUnits>=0?iUnits:5;  // ✅ col 5 = Actual Units
    const colHours=iHours>=0?iHours:7;  // ✅ col 7 = Actual Hours
    console.log('[Wash] CSV sep:',JSON.stringify(sep),'cols: main='+colMain+' proc='+colProc+' attr='+colAttr+' units='+colUnits+' hours='+colHours);
    // ✅ DEBUG : afficher les 5 premières lignes parsées pour vérifier les colonnes
    console.log('[Wash] HEADERS:',headers);
    for(let di=1;di<=Math.min(5,lines.length-1);di++){
      const dl=lines[di].trim();if(!dl)continue;
      const dc=parseCSVLine(dl,sep).map(c=>c.replace(/"/g,'').trim());
      console.log(`[Wash] Row${di} (${dc.length} cols):`,dc.slice(0,12));
    }
    const seen={};
    for(let i=1;i<lines.length;i++){
      const line=lines[i].trim();if(!line)continue;
      const cols=parseCSVLine(line,sep).map(c=>c.replace(/"/g,'').trim());
      if(cols.length<Math.max(colMain,colProc,colHours)+1)continue;
      const mainProc=cols[colMain]||'';
      const procName=cols[colProc]||'';
      const hours=parseNum(cols[colHours]||'');
      const units=parseNum(cols[colUnits]||'');
      if(!mainProc||!procName)continue;
      let cat=CATS.find(c=>mainProc.toLowerCase()===c.toLowerCase());
      if(!cat)cat=CATS.find(c=>mainProc.toLowerCase().startsWith(c.toLowerCase()));
      if(!cat)cat=CSV_MAIN_TO_CAT[mainProc.toLowerCase()];
      if(!cat)continue;
      if(IGNORED_PROCESSES.has(procName))continue;
      const key=`${cat}::${procName}`;
      // ✅ Sommer toutes les lignes du même process (attributs multiples)
      if(!seen[key]){
        seen[key]={cat,name:procName,units:0,hours:0};
      }
      seen[key].units+=units;
      seen[key].hours+=hours;
    }
    for(const entry of Object.values(seen)){
      if(!result[entry.cat])continue;
      result[entry.cat].push({name:entry.name,units:entry.units,hours:parseFloat(entry.hours.toFixed(3))});
    }
    console.log('[Wash] PPA processes:',Object.entries(result).map(([k,v])=>`${k}:${v.length}`).join(', '));
    return result;
  }

  // ─── PPA HTML PARSER v1.3.5 ───────────────────────────────────────────────────
  // HTML FCLM utilise format FR : espaces = milliers, virgule = décimal
  function parsePpaHtml(html) {
    const doc=new DOMParser().parseFromString(html,'text/html');
    const result={};
    const tbl=doc.querySelector('#summary-table')||doc.querySelector('table.result-table');
    if(!tbl)return result;
    for(const row of tbl.rows){
      const cells=row.cells;if(cells.length<6)continue;
      const name=cells[0].textContent.trim();
      if(!name.toLowerCase().endsWith('total'))continue;
      const base=name.replace(/\s*Total$/i,'').trim();
      let cat=CATS.find(c=>base.toLowerCase()===c.toLowerCase());
      if(!cat)cat=CATS.find(c=>base.toLowerCase().startsWith(c.toLowerCase()));
      if(!cat)cat=CSV_MAIN_TO_CAT[base.toLowerCase()];
      if(!cat)continue;
      result[cat]={
        units:parseNum(cells[1]?.textContent),
        hours:parseNum(cells[3]?.textContent)
      };
    }
    return result;
  }

  // ─── HELPERS ──────────────────────────────────────────────────────────────────
  const fmt  =(v,d=0)=>v==null||isNaN(v)?'-':Number(v).toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d});
  const fmtD =(v,d=1)=>v==null||isNaN(v)?'-':(v>0?'+':'')+v.toFixed(d);

  // ─── UNPLANNED DETECTION ──────────────────────────────────────────────────────
  function detectUnplannedLogs(ppaData,planData,site,dateStr){
    const unplanned=[];
    if(!ppaData||!planData)return unplanned;
    for(const cat of CATS){
      // ✅ ppaData[cat] peut être {totals, processes} (depuis render) ou un tableau direct
      const procs=Array.isArray(ppaData[cat])
        ? ppaData[cat]
        : Array.isArray(ppaData[cat]?.processes)
          ? ppaData[cat].processes
          : [];
      for(const proc of procs){
        if(!proc||proc.hours<=0)continue;
        if(IGNORED_PROCESSES.has(proc.name))continue; // ✅ ne pas signaler les process ignorés
        const planHrs=getProcPlanHrs(planData,cat,proc.name);
        if(planHrs===null||planHrs===0)unplanned.push({cat,name:proc.name,hours:proc.hours,units:proc.units});
      }
    }
    return unplanned;
  }

  // ─── SENT TRACKERS ────────────────────────────────────────────────────────────
  function wasUnplannedAlertSent(site,dateStr,procName){return((ls.get(K.SENT)||{})[`${site}::${dateStr}`]||[]).includes(procName);}
  function markUnplannedAlertSent(site,dateStr,procName){const s=ls.get(K.SENT)||{},k=`${site}::${dateStr}`;if(!s[k])s[k]=[];if(!s[k].includes(procName))s[k].push(procName);ls.set(K.SENT,s);}
  function wasSummarySentToday(site,dateStr,hour){return(ls.get(K.SUMMARY_SENT)||{})[`${site}::${dateStr}::${hour}`]===true;}
  function markSummarySent(site,dateStr,hour){const s=ls.get(K.SUMMARY_SENT)||{};s[`${site}::${dateStr}::${hour}`]=true;ls.set(K.SUMMARY_SENT,s);}

  // ─── SLACK ────────────────────────────────────────────────────────────────────
  function sendSlack(payload){
    const cfg=getSlack();if(!cfg.webhookUrl)return Promise.reject('No webhook URL');
    return new Promise((res,rej)=>GM_xmlhttpRequest({
      method:'POST',url:cfg.webhookUrl,headers:{'Content-Type':'application/json'},
      data:JSON.stringify(payload),
      onload:r=>r.status===200?res():rej(`${r.status}: ${r.responseText}`),
      onerror:e=>rej(e)
    }));
  }
  async function sendUnplannedAlert(site,dateStr,unplanned){
    const newOnes=unplanned.filter(u=>!wasUnplannedAlertSent(site,dateStr,u.name));
    if(!newOnes.length)return;
    const lines=[`🚨 Unplanned Logs Detected — ${site} | ${dateStr}`,`_These processes have actual hours but are NOT in the QS Plan:_`,''];
    newOnes.forEach(u=>lines.push(`• ${u.name} (${u.cat}) — ${u.hours.toFixed(1)} hrs | ${fmt(u.units)} units`));
    lines.push('',`🔍 <https://fclm-portal.amazon.com/ppa/inspect/process?spanType=Day&startDateDay=${dateStr.replace(/-/g,'%2F')}&warehouseId=${site}|View in FCLM PPA>`);
    try{await sendSlack({text:lines.join('\n')});newOnes.forEach(u=>markUnplannedAlertSent(site,dateStr,u.name));}
    catch(e){console.error('[Wash] Unplanned alert failed:',e);}
  }
  async function sendSummary(site,dateStr,catData,hasPlan,totalActHrs,totalPlanHrs,totalDelta,unplanned,summaryType){
    const now=new Date().toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',hour12:false});
    const is14h=summaryType==='14h',is23h=summaryType==='23h';
    const lines=[`📊 Wash Sheet ${is14h?'14:00':is23h?'23:00':''} Summary — ${site} | ${dateStr} | ${now}`,hasPlan?'':'_⚠️ No QS Plan loaded_',''];
    if(!hasPlan){lines.push('No plan loaded.');}
    else if(is14h){
      for(const cat of['Sort','Pick Stage','UTR']){
        const d=catData.find(x=>x.cat===cat);if(!d)continue;
        lines.push(`${dEmoji(d.delta)} ${cat} — Actual: ${d.actHrs.toFixed(1)}h | Plan: ${d.planHrs!=null?d.planHrs.toFixed(1):'-'}h | ${d.delta!=null?(d.delta>0?'+':'')+d.delta.toFixed(1)+' hrs':''}`);
        (d._processes||[]).filter(p=>p.planHrs>0).forEach(p=>{const pd=p.actHrs-p.planHrs;lines.push(`  ${dEmoji(pd)} ${p.name}: ${p.actHrs.toFixed(1)}h / Plan ${p.planHrs.toFixed(1)}h (${pd>0?'+':''}${pd.toFixed(1)}h)`);});
        unplanned.filter(u=>u.cat===cat).forEach(u=>lines.push(`  🚨 ${u.name} (UNPLANNED): ${u.hours.toFixed(1)}h`));
        lines.push('');
      }
      const other=unplanned.filter(u=>!['Sort','Pick Stage','UTR'].includes(u.cat));
      if(other.length){lines.push('🚨 _Other Unplanned:_');other.forEach(u=>lines.push(`• ${u.name} (${u.cat}): ${u.hours.toFixed(1)}h`));lines.push('');}
    }else if(is23h){
      lines.push('*All Planned Processes — End of Day:*','');
      for(const d of catData){
        if(!(d._processes||[]).filter(p=>p.planHrs>0).length)continue;
        lines.push(`${dEmoji(d.delta)} ${d.cat} — ${d.actHrs.toFixed(1)}h / Plan ${d.planHrs!=null?d.planHrs.toFixed(1):'-'}h`);
        d._processes.filter(p=>p.planHrs>0).forEach(p=>{const pd=p.actHrs-p.planHrs;lines.push(`  ${dEmoji(pd)} ${p.name}: ${p.actHrs.toFixed(1)}h / Plan ${p.planHrs.toFixed(1)}h (${pd>0?'+':''}${pd.toFixed(1)}h)`);});
        unplanned.filter(u=>u.cat===d.cat).forEach(u=>lines.push(`  🚨 ${u.name} (UNPLANNED): ${u.hours.toFixed(1)}h`));
        lines.push('');
      }
    }
    lines.push(`🔍 <https://fclm-portal.amazon.com/ppa/inspect/process?spanType=Day&startDateDay=${dateStr.replace(/-/g,'%2F')}&warehouseId=${site}|View in FCLM PPA>`);
    try{await sendSlack({text:lines.filter(l=>l!==undefined).join('\n')});console.log('[Wash] Summary sent:',summaryType);}
    catch(e){console.error('[Wash] Summary failed:',e);}
  }
  function checkScheduledSummaries(site,dateStr,catData,hasPlan,totalActHrs,totalPlanHrs,totalDelta,unplanned){
    const now=new Date(),h=now.getHours(),m=now.getMinutes();
    if((h===14||h===23)&&m<=5&&!wasSummarySentToday(site,dateStr,h)){
      markSummarySent(site,dateStr,h);
      sendSummary(site,dateStr,catData,hasPlan,totalActHrs,totalPlanHrs,totalDelta,unplanned,h===14?'14h':'23h');
    }
  }

  // ─── STYLES ───────────────────────────────────────────────────────────────────
  function injectStyles(){
    GM_addStyle(`
      *{box-sizing:border-box}
      body{margin:0;padding:0;background:#f0f2f5;font-family:'Amazon Ember',Arial,sans-serif}
      #wc{max-width:1400px;margin:0 auto;padding:16px}
      .w-header{background:linear-gradient(135deg,#1565c0,#0d47a1);color:#fff;padding:12px 20px;border-radius:10px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
      .w-header h1{margin:0;font-size:1.2rem;font-weight:700}
      .w-hdr-left,.w-hdr-right{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
      .w-site-input{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.3);padding:5px 10px;border-radius:5px;font-size:14px;width:80px;text-align:center;text-transform:uppercase}
      .w-site-input:focus{outline:none;background:rgba(255,255,255,.25)}
      .w-date-badge{background:rgba(255,255,255,.15);padding:5px 10px;border-radius:5px;font-size:13px}
      .w-intraday{background:#ff8f00;padding:3px 8px;border-radius:4px;font-size:11px;font-weight:700}
      .w-btn{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.25);padding:5px 14px;border-radius:5px;cursor:pointer;font-size:13px;white-space:nowrap;transition:background .2s}
      .w-btn:hover{background:rgba(255,255,255,.28)}
      .w-btn-sm{padding:4px 10px;font-size:12px}
      .w-btn-plan-ok{background:rgba(46,125,50,.6)!important;border-color:#4caf50!important}
      .w-btn-plan-miss{background:rgba(198,40,40,.6)!important;border-color:#ef5350!important}
      .w-btn-slack{background:rgba(74,21,75,.7)!important;border-color:#9c27b0!important}
      .w-divider{width:1px;height:22px;background:rgba(255,255,255,.25);flex-shrink:0}
      .w-time-sel{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.25);padding:4px 6px;border-radius:4px;font-size:12px}
      .w-time-sel option{color:#333;background:#fff}
      .w-banner{background:#fff3e0;border:1px solid #ff8f00;border-radius:7px;padding:10px 16px;margin-bottom:12px;font-size:13px;color:#e65100;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
      .w-banner-unplanned{background:#ffebee;border-color:#ef5350;color:#c62828}
      .w-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:16px}
      .w-card{background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 2px 6px rgba(0,0,0,.08);border-left:4px solid #1565c0}
      .w-card-title{font-size:11px;font-weight:700;color:#546e7a;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
      .w-card-main{font-size:26px;font-weight:700;color:#263238;font-variant-numeric:tabular-nums;margin-bottom:4px}
      .w-card-sub{font-size:12px;color:#888;display:flex;justify-content:space-between}
      .w-card-delta{font-size:14px;font-weight:700;margin-top:4px}
      .w-table-wrap{background:#fff;border-radius:10px;box-shadow:0 2px 6px rgba(0,0,0,.08);overflow:hidden;margin-bottom:16px}
      .w-table-header{background:#263238;color:#fff;padding:12px 16px;font-size:14px;font-weight:600;display:flex;justify-content:space-between;align-items:center}
      table.w-table{width:100%;border-collapse:collapse;font-size:13px}
      table.w-table thead th{background:#37474f;color:#fff;padding:9px 12px;text-align:center;font-weight:500;white-space:nowrap}
      table.w-table thead th:first-child{text-align:left}
      table.w-table thead th.th-plan{background:#1b5e20}
      table.w-table thead th.th-wash{background:#4a148c}
      table.w-table tbody td{padding:7px 12px;border-bottom:1px solid #eceff1;text-align:center}
      table.w-table tbody td:first-child{text-align:left}
      table.w-table tbody tr.w-cat-row{cursor:pointer;background:#f5f5f5}
      table.w-table tbody tr.w-cat-row:hover{background:#e3f2fd}
      table.w-table tbody tr.w-cat-row td:first-child{font-weight:700;color:#263238}
      table.w-table tbody tr.w-proc-row td:first-child{padding-left:28px;color:#546e7a;font-size:12px}
      table.w-table tbody tr.w-proc-row{background:#fafafa}
      table.w-table tbody tr.w-proc-row:hover{background:#f5f5f5}
      table.w-table tbody tr.w-total-row{background:#eceff1;font-weight:700}
      table.w-table tbody tr.w-total-row td{border-top:2px solid #90a4ae}
      table.w-table tbody tr.w-unplanned-row{background:#fff8e1}
      table.w-table tbody tr.w-unplanned-row td:first-child{color:#e65100}
      .w-collapsed{display:none!important}
      .w-num{font-variant-numeric:tabular-nums}
      .w-green{color:#2e7d32;font-weight:600}.w-yellow{color:#f57f17;font-weight:600}.w-red{color:#c62828;font-weight:600}
      .w-bg-green{background:#e8f5e9!important;color:#2e7d32;font-weight:600}
      .w-bg-yellow{background:#fff8e1!important;color:#f57f17;font-weight:600}
      .w-bg-red{background:#ffebee!important;color:#c62828;font-weight:600}
      .w-unplanned-badge{background:#ff5722;color:#fff;font-size:10px;padding:2px 6px;border-radius:10px;margin-left:6px;font-weight:700}
      .w-loading{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:300px;color:#1565c0;font-size:16px;gap:10px}
      .w-spinner{width:36px;height:36px;border:4px solid #e3f2fd;border-top-color:#1565c0;border-radius:50%;animation:w-spin 1s linear infinite}
      @keyframes w-spin{to{transform:rotate(360deg)}}
      .w-error{color:#c62828;background:#ffebee;padding:16px;border-radius:8px;margin:16px}
      .w-footer{text-align:center;padding:12px 0 6px;font-size:11px;color:#90a4ae}
      .w-footer a{color:#1565c0;text-decoration:none}
      .w-source-badge{font-size:11px;padding:2px 8px;border-radius:4px;font-weight:600;margin-left:6px}
      .w-source-sp{background:#e3f2fd;color:#0d47a1}
      .w-source-wiki{background:#f3e5f5;color:#6a1b9a}
      .w-source-cache{background:#e8f5e9;color:#2e7d32}
      .w-source-none{background:#ffebee;color:#c62828}
    `);
  }

  // ─── PLAN MODAL ───────────────────────────────────────────────────────────────
  function showPlanModal(site,dateStr){
    document.getElementById('w-plan-modal')?.remove();
    const loaded=getLoadedDates(site);
    const modal=document.createElement('div');modal.id='w-plan-modal';
    modal.setAttribute('style','position:fixed!important;top:0!important;left:0!important;right:0!important;bottom:0!important;background:rgba(0,0,0,0.6)!important;z-index:2147483647!important;display:flex!important;align-items:center!important;justify-content:center!important;visibility:visible!important;opacity:1!important');
    const box=document.createElement('div');
    Object.assign(box.style,{background:'#fff',borderRadius:'10px',padding:'22px',width:'600px',maxWidth:'95vw',maxHeight:'85vh',display:'flex',flexDirection:'column',boxShadow:'0 8px 30px rgba(0,0,0,0.25)',fontFamily:'-apple-system,sans-serif',overflowY:'auto'});
    const title=document.createElement('h3');title.textContent=`📅 QS Plan — ${site}`;Object.assign(title.style,{margin:'0 0 14px',fontSize:'16px',color:'#263238'});box.appendChild(title);
    const infoDiv=document.createElement('div');Object.assign(infoDiv.style,{background:'#e3f2fd',borderRadius:'8px',padding:'12px 16px',marginBottom:'12px',fontSize:'13px'});
    infoDiv.innerHTML=`
      <div style="font-weight:600;color:#0d47a1;margin-bottom:6px">☁️ Chargement automatique</div>
      <div style="color:#546e7a;font-size:12px;margin-bottom:4px">Le script cherche <strong>FCLM_Plan_${dateStr}.csv</strong> dans cet ordre :</div>
      <div style="font-size:12px;color:#546e7a">
        1️⃣ <strong>SharePoint</strong> — amazonfra.sharepoint.com/sites/DWP2_Station/...<br>
        2️⃣ <strong>Wiki Amazon</strong> — w.amazon.com/bin/download/Users/haoulati/fclm-plan/
      </div>
      <div style="margin-top:10px;display:flex;gap:8px;align-items:center">
        <button id="w-sp-reload" style="padding:6px 14px;background:#1565c0;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:600">🔄 Recharger</button>
        <span id="w-sp-status" style="font-size:11px;color:#888"></span>
      </div>`;
    box.appendChild(infoDiv);
    if(loaded.length){const ld=document.createElement('div');Object.assign(ld.style,{marginBottom:'12px',fontSize:'12px',color:'#2e7d32',background:'#e8f5e9',padding:'8px 12px',borderRadius:'6px'});ld.innerHTML=`✓ Plans en cache : <strong>${loaded.join(', ')}</strong>`;box.appendChild(ld);}
    const ml=document.createElement('p');ml.innerHTML='📤 <strong>Upload manuel</strong> <span style="color:#888;font-size:11px">(fallback si les deux sources sont inaccessibles)</span>';Object.assign(ml.style,{fontSize:'13px',color:'#546e7a',margin:'0 0 6px'});box.appendChild(ml);
    const fileInput=document.createElement('input');fileInput.type='file';fileInput.accept='.csv';Object.assign(fileInput.style,{display:'block',width:'100%',padding:'10px',border:'2px dashed #90a4ae',borderRadius:'8px',background:'#fafafa',cursor:'pointer',fontSize:'13px',color:'#546e7a',marginBottom:'10px'});box.appendChild(fileInput);
    const statusEl=document.createElement('div');Object.assign(statusEl.style,{fontSize:'12px',marginBottom:'10px',display:'none'});box.appendChild(statusEl);
    const closeBtn=document.createElement('button');closeBtn.textContent='Fermer';Object.assign(closeBtn.style,{padding:'7px 18px',borderRadius:'5px',border:'none',cursor:'pointer',fontSize:'13px',fontWeight:'600',background:'#eee',color:'#333',alignSelf:'flex-end'});closeBtn.onclick=()=>modal.remove();box.appendChild(closeBtn);
    modal.appendChild(box);document.body.appendChild(modal);
    modal.addEventListener('click',e=>{if(e.target===modal)modal.remove();});
    box.querySelector('#w-sp-reload').onclick=async function(){
      const st=box.querySelector('#w-sp-status');st.textContent='⏳ Chargement...';st.style.color='#e65100';
      delPlan(site,dateStr);
      const result=await autoLoadPlan(site,dateStr);
      if(result.success){st.textContent=`✅ Chargé depuis ${result.source==='sharepoint'?'SharePoint':result.source==='wiki'?'Wiki':'cache'} !`;st.style.color='#2e7d32';setTimeout(()=>{modal.remove();showPlanModal(site,dateStr);init();},1000);}
      else{st.textContent=`❌ ${result.error}`;st.style.color='#c62828';}
    };
    fileInput.addEventListener('change',e=>{
      const file=e.target.files[0];if(!file)return;
      const reader=new FileReader();
      reader.onload=ev=>{
        try{
          const byDate=parseQsCsv(ev.target.result,site);const dates=Object.keys(byDate);
          if(!dates.length){showStatus('error',`Aucune donnée pour "${site}".`);return;}
          dates.forEach(d=>savePlan(site,d,byDate[d]));
          showStatus('success',`✓ Plan chargé pour : ${dates.join(', ')}`);
          setTimeout(()=>{modal.remove();showPlanModal(site,dateStr);init();},1500);
        }catch(err){showStatus('error','Erreur CSV : '+err.message);}
      };
      reader.readAsText(file);
    });
    function showStatus(type,msg){statusEl.style.display='block';statusEl.style.padding='8px 12px';statusEl.style.borderRadius='4px';statusEl.style.color=type==='success'?'#2e7d32':'#c62828';statusEl.style.background=type==='success'?'#e8f5e9':'#ffebee';statusEl.textContent=msg;}
  }

  // ─── SLACK MODAL ──────────────────────────────────────────────────────────────
  function showSlackModal(){
    document.getElementById('w-slack-modal')?.remove();
    const cfg=getSlack();
    const modal=document.createElement('div');modal.id='w-slack-modal';
    modal.setAttribute('style','position:fixed!important;top:0!important;left:0!important;right:0!important;bottom:0!important;background:rgba(0,0,0,0.6)!important;z-index:2147483647!important;display:flex!important;align-items:center!important;justify-content:center!important;visibility:visible!important;opacity:1!important');
    const box=document.createElement('div');box.setAttribute('style','background:#fff;border-radius:10px;padding:22px;width:480px;max-width:95vw;display:flex;flex-direction:column;box-shadow:0 8px 30px rgba(0,0,0,0.25);font-family:-apple-system,sans-serif;overflow-y:auto');
    box.innerHTML='<h3 style="margin:0 0 14px;font-size:16px;color:#263238">⚙️ Slack Settings</h3>'
      +'<label style="font-size:12px;color:#666;display:block;margin-bottom:4px">Slack Webhook URL</label>'
      +'<input type="url" id="w-slack-url" value="'+(cfg.webhookUrl||'')+'" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;font-size:13px;margin-bottom:12px"/>'
      +'<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;margin-bottom:8px"><input type="checkbox" id="w-slack-unplanned" '+(cfg.alertUnplanned!==false?'checked':'')+'/> Alerte logs non prévus</label>'
      +'<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;margin-bottom:8px"><input type="checkbox" id="w-slack-14h" '+(cfg.summary14h!==false?'checked':'')+'/> Summary à 14:00</label>'
      +'<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;margin-bottom:14px"><input type="checkbox" id="w-slack-23h" '+(cfg.summary23h!==false?'checked':'')+'/> Summary à 23:00</label>'
      +'<div style="display:flex;gap:8px;justify-content:flex-end">'
      +'<button id="w-slack-test" style="padding:7px 18px;border-radius:5px;border:none;cursor:pointer;font-size:13px;font-weight:600;background:#7b1fa2;color:#fff">Test</button>'
      +'<button id="w-slack-save" style="padding:7px 18px;border-radius:5px;border:none;cursor:pointer;font-size:13px;font-weight:600;background:#1565c0;color:#fff">Save</button>'
      +'<button id="w-slack-close" style="padding:7px 18px;border-radius:5px;border:none;cursor:pointer;font-size:13px;font-weight:600;background:#eee;color:#333">Cancel</button>'
      +'</div>';
    modal.appendChild(box);document.body.appendChild(modal);
    box.querySelector('#w-slack-save').onclick=()=>{saveSlack({webhookUrl:box.querySelector('#w-slack-url').value.trim(),alertUnplanned:box.querySelector('#w-slack-unplanned').checked,summary14h:box.querySelector('#w-slack-14h').checked,summary23h:box.querySelector('#w-slack-23h').checked});modal.remove();};
    box.querySelector('#w-slack-test').onclick=async()=>{const url=box.querySelector('#w-slack-url').value.trim();if(!url){alert('Enter URL first');return;}saveSlack(Object.assign({},getSlack(),{webhookUrl:url}));try{await sendSlack({text:'✅ FCLM Wash — Slack test OK!'});alert('✅ Sent!');}catch(e){alert('❌ Failed: '+e);}};
    box.querySelector('#w-slack-close').onclick=()=>modal.remove();
    modal.addEventListener('click',e=>{if(e.target===modal)modal.remove();});
  }

  // ─── RENDER ───────────────────────────────────────────────────────────────────
  function render(ppaData,planData,site,dateStr,timeRange,hasPlan,planSource){
    document.body.innerHTML='';document.documentElement.classList.add('wash-ready');
    const wc=document.createElement('div');wc.id='wc';
    const catData=[];let totalActHrs=0,totalPlanHrs=0,totalActUnits=0,totalPlanUnits=0;
    for(const cat of CATS){
      const ppa=ppaData[cat]?.totals||{units:0,hours:0};
      const plan=planData?.[cat]||{hours:0,units:0};
      const planHrs=hasPlan?(plan.hours||0):null,planUnits=hasPlan?(plan.units||0):null;
      const delta=hasPlan?(ppa.hours||0)-planHrs:null;
      totalActHrs+=ppa.hours||0;totalActUnits+=ppa.units||0;
      if(hasPlan){totalPlanHrs+=planHrs;totalPlanUnits+=planUnits;}
      const procs=(ppaData[cat]?.processes||[]).map(proc=>({name:proc.name,actHrs:proc.hours||0,actUnits:proc.units||0,planHrs:hasPlan?getProcPlanHrs(planData,cat,proc.name)??null:null}));
      if(hasPlan&&planData[cat]?.processes){Object.entries(planData[cat].processes).forEach(([csvName,planProc])=>{if(planProc.plannedHours>0&&!procs.find(p=>p.name===csvName||PPA_TO_CSV[p.name]===csvName))procs.push({name:csvName,actHrs:0,actUnits:0,planHrs:planProc.plannedHours});});}
      catData.push({cat,actHrs:ppa.hours||0,actUnits:ppa.units||0,planHrs,planUnits,delta,_processes:procs});
    }
    const totalDelta=hasPlan?totalActHrs-totalPlanHrs:null;
    const unplanned=hasPlan?detectUnplannedLogs(ppaData,planData,site,dateStr):[];
    const slackCfg=getSlack();
    if(slackCfg.webhookUrl){
      if(hasPlan&&slackCfg.alertUnplanned!==false&&unplanned.length>0)sendUnplannedAlert(site,dateStr,unplanned);
      if(slackCfg.summary14h!==false||slackCfg.summary23h!==false)checkScheduledSummaries(site,dateStr,catData,hasPlan,totalActHrs,totalPlanHrs,totalDelta,unplanned);
    }
    const sourceBadge=planSource==='sharepoint'?'<span class="w-source-badge w-source-sp">☁️ SharePoint</span>'
                     :planSource==='wiki'?'<span class="w-source-badge w-source-wiki">📖 Wiki</span>'
                     :hasPlan?'<span class="w-source-badge w-source-cache">💾 Cache</span>'
                     :'<span class="w-source-badge w-source-none">⚠️ No Plan</span>';
    const loaded=getLoadedDates(site);
    const trLabel=timeRange?`<span class="w-intraday">Intraday ${String(timeRange.startH).padStart(2,'0')}:00-${String(timeRange.endH).padStart(2,'0')}:00</span>`:'';
    const hourOpts=sel=>Array.from({length:24},(_,i)=>`<option value="${i}"${i===sel?' selected':''}>${String(i).padStart(2,'0')}:00</option>`).join('');
    const header=document.createElement('div');header.className='w-header';
    header.innerHTML=`
      <div class="w-hdr-left">
        <h1>📋 Wash Sheet</h1>
        <input class="w-site-input" id="w-site-inp" value="${site}" maxlength="6" spellcheck="false"/>
        <span class="w-date-badge">${dateStr}</span>
        ${trLabel} ${sourceBadge}
        <button id="w-prev-day" class="w-btn w-btn-sm">◀ Prev</button>
        <button id="w-next-day" class="w-btn w-btn-sm" ${dateStr>=new Date().toISOString().split('T')[0]?'style="opacity:.4;cursor:not-allowed"':''}>Next ▶</button>
      </div>
      <div class="w-hdr-right">
        <div style="display:flex;align-items:center;gap:6px">
          <select id="w-sh" class="w-time-sel"><option value="">Start</option>${hourOpts(timeRange?.startH??'')}</select>
          <span style="color:rgba(255,255,255,.5)">-</span>
          <select id="w-eh" class="w-time-sel"><option value="">End</option>${hourOpts(timeRange?.endH??'')}</select>
          <button id="w-time-go" class="w-btn w-btn-sm">Go</button>
          ${timeRange?'<button id="w-time-clr" class="w-btn w-btn-sm">Clear</button>':''}
        </div>
        <div class="w-divider"></div>
        <button id="w-refresh" class="w-btn">↻ Refresh</button>
        <div class="w-divider"></div>
        <button id="w-plan-btn" class="w-btn ${hasPlan?'w-btn-plan-ok':'w-btn-plan-miss'}">
          📅 Plan ${hasPlan?'✓':'!'} (${loaded.length} days)
        </button>
        <div class="w-divider"></div>
        <button id="w-slack-now" class="w-btn w-btn-slack">📤 Send to Slack</button>
        <button id="w-slack-cfg" class="w-btn w-btn-sm">⚙️</button>
      </div>`;
    wc.appendChild(header);
    if(!hasPlan){const b=document.createElement('div');b.className='w-banner';b.innerHTML=`⚠️ Aucun plan pour <strong>${dateStr}</strong>. Upload <strong>FCLM_Plan_${dateStr}.csv</strong> sur SharePoint ou Wiki. <button id="w-banner-plan" class="w-btn w-btn-sm" style="background:#ff8f00;border-color:#ff8f00;margin-left:8px">📤 Gérer</button>`;wc.appendChild(b);}
    if(unplanned.length>0){const b=document.createElement('div');b.className='w-banner w-banner-unplanned';b.innerHTML=`🚨 <strong>${unplanned.length} log${unplanned.length>1?'s':''} non prévu${unplanned.length>1?'s':''}:</strong> ${unplanned.map(u=>`<strong>${u.name}</strong> (${u.cat}, ${u.hours.toFixed(1)}h)`).join(' · ')} &nbsp;<a href="https://fclm-portal.amazon.com/ppa/inspect/process?spanType=Day&startDateDay=${dateStr.replace(/-/g,'%2F')}&warehouseId=${site}" target="_blank" style="color:#c62828;font-weight:700">Voir →</a>`;wc.appendChild(b);}
    const cards=document.createElement('div');cards.className='w-cards';
    const oc=document.createElement('div');oc.className='w-card';oc.style.borderLeftColor=dClr(totalDelta);
    oc.innerHTML=`<div class="w-card-title">Overall</div><div class="w-card-main">${fmt(totalActHrs,1)} hrs</div><div class="w-card-sub"><span>Plan</span><span>${hasPlan?fmt(totalPlanHrs,1)+' hrs':'—'}</span></div>${hasPlan?`<div class="w-card-delta ${dCls(totalDelta)}">${fmtD(totalDelta)} hrs vs plan</div>`:''}`;
    cards.appendChild(oc);
    catData.forEach(d=>{
      if(d.actHrs===0&&(!hasPlan||d.planHrs===0))return;
      const c=document.createElement('div');c.className='w-card';c.style.borderLeftColor=dClr(d.delta);
      c.innerHTML=`<div class="w-card-title">${d.cat}</div><div class="w-card-main">${fmt(d.actHrs,1)} hrs</div><div class="w-card-sub"><span>Units</span><span>${fmt(d.actUnits)}</span></div>${hasPlan?`<div class="w-card-delta ${dCls(d.delta)}">${fmtD(d.delta)} hrs vs plan</div>`:''}`;
      cards.appendChild(c);
    });
    wc.appendChild(cards);
    const tw=document.createElement('div');tw.className='w-table-wrap';
    tw.innerHTML=`<div class="w-table-header"><span>📊 Detail by Category & Process</span><span style="font-size:12px;color:#90a4ae">Click a category to expand/collapse</span></div>`;
    const table=document.createElement('table');table.className='w-table';
    table.innerHTML=`<thead><tr><th rowspan="2">Category / Process</th><th colspan="2">PPA Actuals</th><th colspan="2" class="th-plan">QS Plan</th><th colspan="2" class="th-wash">Wash</th></tr><tr><th>Units</th><th>Hours</th><th class="th-plan">Plan Hrs</th><th class="th-plan">Plan Units</th><th class="th-wash">Hrs +/- Plan</th><th class="th-wash">% to Plan</th></tr></thead>`;
    const tbody=document.createElement('tbody');
    for(const d of catData){
      const pct=(hasPlan&&d.planHrs)?(d.actHrs/d.planHrs*100):null;
      const ctr=document.createElement('tr');ctr.className='w-cat-row';ctr.dataset.cat=d.cat;
      ctr.innerHTML=`<td>${d.cat} <span style="font-size:11px;color:#90a4ae">▼</span></td><td class="w-num">${fmt(d.actUnits)}</td><td class="w-num">${fmt(d.actHrs,1)}</td><td class="w-num">${hasPlan?fmt(d.planHrs,1):'-'}</td><td class="w-num">${hasPlan?fmt(d.planUnits):'-'}</td><td class="w-num ${dBg(d.delta)}">${hasPlan?fmtD(d.delta):'-'}</td><td class="w-num">${pct!=null?fmt(pct,1)+'%':'-'}</td>`;
      tbody.appendChild(ctr);
      for(const proc of(ppaData[d.cat]?.processes||[])){
        const ph=hasPlan?getProcPlanHrs(planData,d.cat,proc.name):null;
        const pd=(hasPlan&&ph!=null)?proc.hours-ph:null;
        const isU=hasPlan&&ph===null&&proc.hours>0; // ✅ null = pas dans le plan = UNPLANNED
        const pp=(hasPlan&&ph)?proc.hours/ph*100:null;
        // ✅ Yard Marshall : Inbound + Outbound partagent le plan — afficher delta combiné
        const isYard=['Inbound Traffic Controller','Outbound Traffic Controller'].includes(proc.name);
        let displayPh=ph,displayPd=pd,displayPp=pp;
        if(isYard&&hasPlan&&ph!=null){
          const ymProcs=ppaData[d.cat]?.processes?.filter(p=>['Inbound Traffic Controller','Outbound Traffic Controller'].includes(p.name))||[];
          const ymTotal=ymProcs.reduce((s,p)=>s+p.hours,0);
          displayPd=ymTotal-ph; // delta = total des deux vs plan Yard Marshall
          displayPp=ph?ymTotal/ph*100:null;
        }
        const ptr=document.createElement('tr');ptr.className=`w-proc-row${isU?' w-unplanned-row':''}`;ptr.dataset.cat=d.cat;
        ptr.innerHTML=`<td>${proc.name}${isU?'<span class="w-unplanned-badge">UNPLANNED</span>':''}</td><td class="w-num">${fmt(proc.units)}</td><td class="w-num">${fmt(proc.hours,1)}</td><td class="w-num">${displayPh!=null?fmt(displayPh,1):'-'}</td><td class="w-num">-</td><td class="w-num ${dBg(displayPd)}">${displayPd!=null?fmtD(displayPd):'-'}</td><td class="w-num">${displayPp!=null?fmt(displayPp,1)+'%':'-'}</td>`;
        tbody.appendChild(ptr);
      }
    }
    const tpct=(hasPlan&&totalPlanHrs)?(totalActHrs/totalPlanHrs*100):null;
    const ttr=document.createElement('tr');ttr.className='w-total-row';
    ttr.innerHTML=`<td>TOTAL</td><td class="w-num">${fmt(totalActUnits)}</td><td class="w-num">${fmt(totalActHrs,1)}</td><td class="w-num">${hasPlan?fmt(totalPlanHrs,1):'-'}</td><td class="w-num">${hasPlan?fmt(totalPlanUnits):'-'}</td><td class="w-num ${dBg(totalDelta)}">${hasPlan?fmtD(totalDelta):'-'}</td><td class="w-num">${tpct!=null?fmt(tpct,1)+'%':'-'}</td>`;
    tbody.appendChild(ttr);table.appendChild(tbody);tw.appendChild(table);wc.appendChild(tw);
    const footer=document.createElement('div');footer.className='w-footer';
    footer.innerHTML=`FCLM Wash Sheet v1.3.5 — Created by <a href="https://phonetool.amazon.com/users/haoulati" target="_blank">haoulati</a>`;
    wc.appendChild(footer);document.body.appendChild(wc);
    tbody.querySelectorAll('.w-cat-row').forEach(row=>{row.addEventListener('click',()=>tbody.querySelectorAll(`.w-proc-row[data-cat="${row.dataset.cat}"]`).forEach(r=>r.classList.toggle('w-collapsed')));});
    document.getElementById('w-plan-btn')?.addEventListener('click',()=>showPlanModal(site,dateStr));
    document.getElementById('w-banner-plan')?.addEventListener('click',()=>showPlanModal(site,dateStr));
    document.getElementById('w-refresh')?.addEventListener('click',()=>init());
    document.getElementById('w-slack-cfg')?.addEventListener('click',()=>showSlackModal());
    document.getElementById('w-slack-now')?.addEventListener('click',async function(){
      const btn=this;btn.textContent='...';
      try{await sendSummary(site,dateStr,catData,hasPlan,totalActHrs,totalPlanHrs,totalDelta,unplanned,'23h');btn.textContent='✓ Sent';}
      catch(e){btn.textContent='❌ Failed';}
      setTimeout(()=>{btn.textContent='📤 Send to Slack';},2000);
    });
    document.getElementById('w-prev-day')?.addEventListener('click',()=>{const d=new Date(dateStr);d.setDate(d.getDate()-1);const p=new URLSearchParams(window.location.search);p.set('startDateDay',d.toISOString().split('T')[0].replace(/-/g,'/'));['startDateIntraday','endDateIntraday','startHourIntraday','endHourIntraday','startMinuteIntraday','endMinuteIntraday','spanType','maxIntradayDays'].forEach(k=>p.delete(k));history.replaceState(null,'','?'+p.toString());init();});
    document.getElementById('w-next-day')?.addEventListener('click',()=>{const d=new Date(dateStr);d.setDate(d.getDate()+1);if(d.toISOString().split('T')[0]>new Date().toISOString().split('T')[0])return;const p=new URLSearchParams(window.location.search);p.set('startDateDay',d.toISOString().split('T')[0].replace(/-/g,'/'));['startDateIntraday','endDateIntraday','startHourIntraday','endHourIntraday','startMinuteIntraday','endMinuteIntraday','spanType','maxIntradayDays'].forEach(k=>p.delete(k));history.replaceState(null,'','?'+p.toString());init();});
    document.getElementById('w-time-go')?.addEventListener('click',()=>{const sh=document.getElementById('w-sh').value,eh=document.getElementById('w-eh').value;if(!sh||!eh)return;const p=new URLSearchParams(window.location.search);p.set('startHourIntraday',sh);p.set('startMinuteIntraday','0');p.set('endHourIntraday',eh);p.set('endMinuteIntraday','0');const d=dateStr.replace(/-/g,'/');p.set('startDateIntraday',d);p.set('endDateIntraday',d);history.replaceState(null,'','?'+p.toString());init();});
    document.getElementById('w-time-clr')?.addEventListener('click',()=>{const p=new URLSearchParams(window.location.search);['startHourIntraday','startMinuteIntraday','endHourIntraday','endMinuteIntraday','startDateIntraday','endDateIntraday','spanType','maxIntradayDays'].forEach(k=>p.delete(k));history.replaceState(null,'','?'+p.toString());init();});
    const si=document.getElementById('w-site-inp');
    if(si){const apply=()=>{const v=si.value.trim().toUpperCase();if(v&&v!==site){setSite(v);init();}};si.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();apply();}});si.addEventListener('blur',apply);}
  }

  // ─── INIT ─────────────────────────────────────────────────────────────────────
  async function init(){
    injectStyles();
    let site=getSite();
    if(!site){
      document.body.innerHTML='';document.documentElement.classList.add('wash-ready');
      const div=document.createElement('div');div.className='w-loading';
      div.innerHTML=`<div>Entre ton code site pour continuer</div><input id="w-site-prompt" style="padding:8px 14px;font-size:16px;text-transform:uppercase;border:2px solid #1565c0;border-radius:5px;width:140px;text-align:center" placeholder="ex: DWP2" maxlength="6"/><button onclick="(()=>{const v=document.getElementById('w-site-prompt').value.trim().toUpperCase();if(v){localStorage.setItem('fclm_wash_site',v);location.reload();}})()" style="padding:8px 20px;background:#1565c0;color:#fff;border:none;border-radius:5px;cursor:pointer;font-size:14px">Go</button>`;
      document.body.appendChild(div);
      document.getElementById('w-site-prompt')?.addEventListener('keydown',e=>{if(e.key==='Enter'){const v=e.target.value.trim().toUpperCase();if(v){setSite(v);init();}}});
      return;
    }
    const dateStr=getDateStr(),timeRange=getTimeRange();
    document.body.innerHTML='';document.documentElement.classList.add('wash-ready');
    const loading=document.createElement('div');loading.className='w-loading';
    loading.innerHTML=`<div class="w-spinner"></div><div>Chargement — ${site} | ${dateStr}</div><div id="w-load-status" style="font-size:12px;color:#888">Récupération FCLM...</div>`;
    document.body.appendChild(loading);
    try{
      const[ppaHtml,ppaCsv]=await Promise.all([fetchPpaHtml(site,dateStr,timeRange),fetchPpaCsv(site,dateStr,timeRange)]);
      const ppaTotals=parsePpaHtml(ppaHtml),ppaProcesses=parsePpaCsv(ppaCsv);
      const ppaData={};for(const cat of CATS)ppaData[cat]={totals:ppaTotals[cat]||{units:0,hours:0},processes:ppaProcesses[cat]||[]};
      const ls2=document.getElementById('w-load-status');if(ls2)ls2.textContent='☁️ Chargement plan...';
      const planResult=await autoLoadPlan(site,dateStr);
      const rawPlan=getPlan(site,dateStr);
      const planData=rawPlan?buildPlanByCat(rawPlan):null;
      render(ppaData,planData,site,dateStr,timeRange,!!planData,planResult.source);
      setTimeout(init,5*60*1000);
    }catch(err){
      document.body.innerHTML=`<div class="w-error"><h2>Erreur</h2><p>${err.message||err}</p><button onclick="location.reload()" style="margin-top:12px;padding:8px 16px;cursor:pointer;border-radius:5px;border:none;background:#1565c0;color:#fff">Retry</button></div>`;
      document.documentElement.classList.add('wash-ready');
      console.error('[Wash] Init error:',err);
    }
  }

  // ─── BACKGROUND CHECKER ───────────────────────────────────────────────────────
  async function backgroundCheck(){
    const site=getSite(),dateStr=getDateStr();if(!site)return;
    const slackCfg=getSlack();if(!slackCfg.webhookUrl)return;
    await autoLoadPlan(site,dateStr);
    const rawPlan=getPlan(site,dateStr);if(!rawPlan)return;
    try{
      const timeRange=getTimeRange();
      const ppaCsvText=(await gmGet(`${BASE}/ppa/inspect/node/csv?warehouseId=${site}&nodeType=DS${buildIntradayParams(dateStr,timeRange)}`)).responseText;
      const ppaProcesses=parsePpaCsv(ppaCsvText);
      const planData=buildPlanByCat(rawPlan);
      const ppaData={};for(const cat of CATS)ppaData[cat]={processes:ppaProcesses[cat]||[]};
      if(slackCfg.alertUnplanned!==false){const unplanned=detectUnplannedLogs(ppaData,planData,site,dateStr);if(unplanned.length>0)await sendUnplannedAlert(site,dateStr,unplanned);}
      const now=new Date(),h=now.getHours(),m=now.getMinutes();
      if((h===14||h===23)&&m<=5&&!wasSummarySentToday(site,dateStr,h)){
        const ppaTotalsText=(await gmGet(`${BASE}/ppa/inspect/node?warehouseId=${site}${buildIntradayParams(dateStr,timeRange)}`)).responseText;
        const ppaTotals=parsePpaHtml(ppaTotalsText);
        const catData=[];let totalActHrs=0,totalPlanHrs=0;
        for(const cat of CATS){
          const ppa=ppaTotals[cat]||{units:0,hours:0},plan=planData[cat]||{hours:0,units:0};
          totalActHrs+=ppa.hours||0;totalPlanHrs+=plan.hours||0;
          const procs=(ppaData[cat]?.processes||[]).map(proc=>({name:proc.name,actHrs:proc.hours||0,actUnits:proc.units||0,planHrs:getProcPlanHrs(planData,cat,proc.name)||0}));
          if(planData[cat]?.processes){Object.entries(planData[cat].processes).forEach(([n,p])=>{if(p.plannedHours>0&&!procs.find(x=>x.name===n||PPA_TO_CSV[x.name]===n))procs.push({name:n,actHrs:0,actUnits:0,planHrs:p.plannedHours});});}
          catData.push({cat,actHrs:ppa.hours||0,planHrs:plan.hours||0,delta:(ppa.hours||0)-(plan.hours||0),_processes:procs});
        }
        const unplanned=detectUnplannedLogs(ppaData,planData,site,dateStr);
        markSummarySent(site,dateStr,h);
        await sendSummary(site,dateStr,catData,true,totalActHrs,totalPlanHrs,totalActHrs-totalPlanHrs,unplanned,h===14?'14h':'23h');
      }
    }catch(e){console.log('[Wash BG] Error:',e.message);}
  }

  // ─── ROUTE ────────────────────────────────────────────────────────────────────
  function route(){
    setTimeout(backgroundCheck,3000);
    setInterval(backgroundCheck,5*60*1000);
    if(window.location.pathname.startsWith('/ppa/wash'))init();
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',route);
  else route();

})();
