/**
 * fetch-planner.js — SDC Operations Dashboard
 * Fetches tasks, detects stuck/stale items, builds team accountability view.
 *
 * TO ADD TEAMS NOTIFICATIONS LATER:
 * 1. Add TEAMS_WEBHOOK_URL to GitHub Secrets
 * 2. Uncomment the postToTeams() call at the bottom of main()
 */

const fs = require('fs');

const TENANT_ID    = process.env.TENANT_ID;
const CLIENT_ID    = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
// const TEAMS_WEBHOOK_URL = process.env.TEAMS_WEBHOOK_URL;

const PLANNERS = [
  { key: 'trades',    label: 'Trades & Distributions', planId: process.env.PLAN_ID_TRADES,     color: '#2563eb' },
  { key: 'paperwork', label: 'Paperwork',               planId: process.env.PLAN_ID_PAPERWORK,  color: '#d97706' },
  { key: 'advisor',   label: 'Advisor Flow',            planId: process.env.PLAN_ID_ADVISOR,    color: '#16a34a' },
  { key: 'locations', label: 'Locations',               planId: process.env.PLAN_ID_LOCATIONS,  color: '#7c3aed' },
  { key: 'conts',     label: 'CONTs & Checks',          planId: process.env.PLAN_ID_CONTS,      color: '#dc2626' },
];

const CLIENT_DEPENDENT_BUCKETS = [
  'waiting on signatures','waiting on client','pending client signature',
  'sent to client','in wip','reassign to acs','reassign to ejd','paperwork submitted',
];

const STUCK_THRESHOLDS = { queue: 1, inprogress: 7, waitingclient: 14, default: 7 };
const COMM_GAP_DAYS = 5;

function isClientDependent(b) {
  if (!b) return false;
  return CLIENT_DEPENDENT_BUCKETS.some(x => b.toLowerCase().includes(x));
}

function getBucketType(b) {
  const bl = (b || '').toLowerCase();
  if (bl.includes('queue') || bl.includes('task recognized')) return 'queue';
  if (bl.includes('in progress') || bl.includes('being pulled')) return 'inprogress';
  if (isClientDependent(b)) return 'waitingclient';
  return 'default';
}

function daysSince(iso) {
  if (!iso) return null;
  return Math.floor((Date.now() - new Date(iso).getTime()) / 86400000);
}

function businessDaysSince(iso) {
  if (!iso) return null;
  let count = 0;
  const cur = new Date(iso);
  const now = new Date();
  while (cur < now) {
    const d = cur.getDay();
    if (d !== 0 && d !== 6) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

/**
 * Extract a date from Advisor Flow task titles.
 * Handles formats like: "05/07", "04/06", "03/26 10:30pm", "JUNE/JULY"
 * Returns a Date object or null.
 */
function extractTitleDate(title) {
  if (!title) return null;
  // Match MM/DD at start of title (e.g. "05/07", "04/06 1:30pm")
  const match = title.match(/^(\d{1,2})\/(\d{1,2})/);
  if (match) {
    const month = parseInt(match[1], 10) - 1;
    const day   = parseInt(match[2], 10);
    const year  = new Date().getFullYear();
    const d     = new Date(year, month, day);
    // If the date is more than 6 months in the future, it's probably last year
    if (d.getTime() - Date.now() > 180 * 86400000) d.setFullYear(year - 1);
    return d;
  }
  return null;
}

function detectStuck(task, notesLastMod) {
  if (task.status === 'complete' || task.status === 'waiting-on-client') return null;

  // Advisor Flow: if title has a date, don't flag as stuck until 7 days after that date
  if (task.plannerKey === 'advisor') {
    const titleDate = extractTitleDate(task.title);
    if (titleDate) {
      const daysSinceEvent = daysSince(titleDate.toISOString());
      if (daysSinceEvent === null || daysSinceEvent < 7) return null; // event hasn't passed + 7 days yet
      // After 7 days past the event date, check for communication gap
      const commGap = notesLastMod
        ? daysSince(notesLastMod) > COMM_GAP_DAYS
        : daysSinceEvent > COMM_GAP_DAYS;
      if (!commGap && task.status !== 'late') return null;
      return {
        type:  task.status === 'late' ? 'overdue-silent' : 'comm-gap',
        age: daysSinceEvent, threshold: 7, commGap,
        label: task.status === 'late'
          ? `Follow-up overdue — ${daysSinceEvent} days since event`
          : `No follow-up notes in ${daysSince(notesLastMod) ?? daysSinceEvent} days`,
      };
    }
  }

  const bt        = getBucketType(task.bucketName);
  const threshold = STUCK_THRESHOLDS[bt];
  const age       = bt === 'queue' ? businessDaysSince(task.lastModifiedDateTime) : daysSince(task.lastModifiedDateTime);
  if (age === null) return null;
  const commGap   = notesLastMod ? daysSince(notesLastMod) > COMM_GAP_DAYS : age > COMM_GAP_DAYS;
  const isStuck   = age >= threshold;
  const silentOD  = task.status === 'late' && commGap;
  if (!isStuck && !commGap && !silentOD) return null;
  return {
    type:  isStuck ? 'stuck' : silentOD ? 'overdue-silent' : 'comm-gap',
    age, threshold, commGap,
    label: isStuck
      ? `No movement for ${age} day${age !== 1 ? 's' : ''}`
      : silentOD ? `Overdue — no client communication logged`
      : `No notes update in ${daysSince(notesLastMod)} days`,
  };
}

function normalizeName(raw) {
  if (!raw) return null;
  let n = raw.trim()
    .replace(/\s+(TOD|NQ|IRA|ROTH|401K|403B|529|UTMA|BL|NW|FID|AMF|NFS|BES|FJD|ZWB|CMB|KSO|ACS|JEK|EJD|BAW|MRB)\b.*/i,'')
    .replace(/\s+[A-Z]{2,4}$/,'').trim();
  if (n.includes(',')) {
    const [last,...rest] = n.split(',').map(p=>p.trim());
    const first = rest.join(',').trim();
    return first ? `${toTC(last)}, ${toTC(first)}` : toTC(last);
  }
  const w = n.split(/\s+/).filter(Boolean);
  if (w.length === 1) return toTC(w[0]);
  if (w.length === 2) return `${toTC(w[1])}, ${toTC(w[0])}`;
  if (w.some(x => x==='&'||x.toLowerCase()==='and')) return w.map(toTC).join(' ');
  return `${toTC(w[w.length-1])}, ${w.slice(0,-1).map(toTC).join(' ')}`;
}
function toTC(s) { return (s||'').toLowerCase().replace(/(?:^|[-\s])(\w)/g,c=>c.toUpperCase()); }

function nameSim(a,b) {
  if (!a||!b) return 0;
  const la=a.split(',')[0]?.toLowerCase().trim()||'';
  const lb=b.split(',')[0]?.toLowerCase().trim()||'';
  if (la!==lb) return 0;
  const fa=(a.split(',')[1]||'').toLowerCase().trim().split(/\s+/)[0]||'';
  const fb=(b.split(',')[1]||'').toLowerCase().trim().split(/\s+/)[0]||'';
  if (!fa||!fb) return 0;
  if (fa.includes(fb)||fb.includes(fa)) return 0.85;
  const m=fa.length,nn=fb.length;
  const dp=Array.from({length:m+1},(_,i)=>Array.from({length:nn+1},(_,j)=>i===0?j:j===0?i:0));
  for(let i=1;i<=m;i++) for(let j=1;j<=nn;j++) dp[i][j]=fa[i-1]===fb[j-1]?dp[i-1][j-1]:1+Math.min(dp[i-1][j],dp[i][j-1],dp[i-1][j-1]);
  const sim=1-dp[m][nn]/Math.max(m,nn);
  return sim>0.5?sim:0;
}

function extractClient(title, key) {
  if (!title) return null;
  const em = title.match(/^(.+?)\s*[—–]\s*.+/);
  if (em) return em[1].trim();
  if (['paperwork','locations','advisor'].includes(key)) {
    const h = title.match(/^(.+?)\s+-\s+.{2,}/);
    if (h) return h[1].trim();
    const p = title.match(/^(.+?)\s*\|\|\s*.+/);
    if (p) return p[1].trim();
  }
  if (['trades','conts'].includes(key)) {
    const w = title.trim().split(/\s+/);
    return w.length>=2 ? `${w[0]} ${w[1]}` : w[0]||null;
  }
  return null;
}

function extractAdvisor(names) {
  const KW = ['brent','frank','zach','cheyenne','katie','melissa','elizabeth','jaiden'];
  return names.find(n => KW.some(k => n.toLowerCase().includes(k))) || null;
}

function fmtDate(iso) {
  if (!iso) return null;
  return new Date(iso).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
}

async function getAccessToken() {
  const url  = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({ grant_type:'client_credentials', client_id:CLIENT_ID, client_secret:CLIENT_SECRET, scope:'https://graph.microsoft.com/.default' });
  const res  = await fetch(url,{method:'POST',body});
  const data = await res.json();
  if (!data.access_token) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

function sleep(ms){return new Promise(r=>setTimeout(r,ms));}

async function graphGet(token, path, retries=4) {
  for (let i=0;i<=retries;i++) {
    const res=await fetch(`https://graph.microsoft.com/v1.0${path}`,{headers:{Authorization:`Bearer ${token}`}});
    if (res.ok) return res.json();
    if (res.status===429) { await sleep((parseInt(res.headers.get('Retry-After')||'0')*1000)||Math.pow(2,i)*2000); continue; }
    throw new Error(`Graph ${path} → ${res.status}`);
  }
  throw new Error(`Graph ${path} → exceeded retries`);
}

const userCache={};
async function resolveUser(token,uid) {
  if (userCache[uid]) return userCache[uid];
  try {
    const d=await graphGet(token,`/users/${uid}?$select=displayName,userPrincipalName,mail`);
    const n=d.displayName||d.mail?.split('@')[0]||d.userPrincipalName?.split('@')[0]||'Unknown';
    userCache[uid]=n; return n;
  } catch { userCache[uid]='Unknown'; return 'Unknown'; }
}

function deriveStatus(t, bucket) {
  if ((t.percentComplete??0)===100) return 'complete';
  if (isClientDependent(bucket)) return 'waiting-on-client';
  if (t.dueDateTime && new Date(t.dueDateTime)<new Date()) return 'late';
  if ((t.percentComplete??0)>0) return 'in-progress';
  return 'not-started';
}

async function aiGroupClients(wipArr, allTasks) {
  if (!GITHUB_TOKEN) return null;
  const sums={};
  for (const t of allTasks.filter(t=>t.status!=='complete'&&t.clientName)) {
    if (!sums[t.clientName]) sums[t.clientName]={planners:new Set(),buckets:[],statuses:[],stuck:0};
    sums[t.clientName].planners.add(t.plannerKey);
    sums[t.clientName].buckets.push(t.bucketName);
    sums[t.clientName].statuses.push(t.status);
    if (t.stuckInfo) sums[t.clientName].stuck++;
  }
  const compact=Object.entries(sums).map(([name,d])=>({
    name, planners:[...d.planners], buckets:[...new Set(d.buckets)].slice(0,3),
    overdue:d.statuses.includes('late'), stuck:d.stuck>0, count:d.buckets.length,
  }));
  const priority=compact.filter(c=>c.overdue||c.stuck||c.planners.length>1).sort((a,b)=>(b.overdue?1:0)-(a.overdue?1:0)).slice(0,60);
  if (!priority.length) return null;
  const prompt=`Financial advisory ops review. Priority clients:\n${JSON.stringify(priority,null,1)}\nAll names: ${compact.map(c=>c.name).join(', ')}\n\nMERGE same person/household (nicknames, couples, format diffs). Give brief stage for multi-planner/overdue clients.\n\nReturn JSON array max 50:\n[{"canonicalName":"Smith, John","merges":["John Smith"],"currentStage":"brief","nextAction":"single action","blockers":null,"urgency":"high|normal|waiting"}]`;
  try {
    const res=await fetch('https://models.inference.ai.azure.com/chat/completions',{
      method:'POST', headers:{'Content-Type':'application/json','Authorization':`Bearer ${GITHUB_TOKEN}`},
      body:JSON.stringify({model:'gpt-4o-mini',messages:[{role:'system',content:'Financial ops analyst. Be concise.'},{role:'user',content:prompt}],max_tokens:2000,temperature:0.1,response_format:{type:'json_object'}}),
    });
    if (!res.ok) { console.log(`AI failed: ${res.status}`); return null; }
    const d=await res.json();
    let text=(d.choices?.[0]?.message?.content||'[]').replace(/```json|```/g,'').trim();
    const parsed=JSON.parse(text);
    return Array.isArray(parsed)?parsed:(parsed.clients||parsed.groups||Object.values(parsed)[0]||[]);
  } catch(e) { console.log(`AI error: ${e.message}`); return null; }
}

// ── TEAMS PLACEHOLDER ─────────────────────────────────────────────────────────
// async function postToTeams(stuck) {
//   if (!process.env.TEAMS_WEBHOOK_URL || !stuck.length) return;
//   await fetch(process.env.TEAMS_WEBHOOK_URL, {
//     method:'POST', headers:{'Content-Type':'application/json'},
//     body: JSON.stringify({ text: `⚠️ ${stuck.length} items need attention:\n` +
//       stuck.slice(0,10).map(t=>`• **${t.clientName||'?'}** — ${t.title} (${t.stuckInfo?.label})`).join('\n') }),
//   });
// }

async function main() {
  console.log('Authenticating...');
  const token = await getAccessToken();

  const data = {
    fetchedAt: new Date().toISOString(),
    planners:[], wip:{}, teamView:{}, overdueTasks:[], stuckTasks:[], possibleDuplicates:[], aiMerges:[],
    stats:{ totalOpen:0, completedToday:0, overdue:0, waitingOnClient:0, wipClients:0, stuck:0, commGap:0 },
  };

  for (const planner of PLANNERS) {
    console.log(`Fetching: ${planner.label}`);
    await sleep(300);
    const [tasksRes, bucketsRes] = await Promise.all([
      graphGet(token,`/planner/plans/${planner.planId}/tasks?$select=id,title,percentComplete,dueDateTime,completedDateTime,assignments,bucketId,lastModifiedDateTime`),
      graphGet(token,`/planner/plans/${planner.planId}/buckets`),
    ]);
    const bmap={};
    for (const b of (bucketsRes.value??[])) bmap[b.id]=b.name;

    const ids=new Set();
    for (const t of (tasksRes.value??[])) for (const uid of Object.keys(t.assignments??{})) ids.add(uid);
    await Promise.all([...ids].map(uid=>resolveUser(token,uid)));

    const today=new Date().toDateString();

    const tasks=(tasksRes.value??[]).map(t=>{
      const bucket=bmap[t.bucketId]??'Unknown';
      const status=deriveStatus(t,bucket);
      const raw=extractClient(t.title,planner.key);
      const client=raw?normalizeName(raw):null;
      const aIds=Object.keys(t.assignments??{});
      const aNames=aIds.map(uid=>userCache[uid]||'Unknown');
      const advisor=extractAdvisor(aNames);
      const doneToday=status==='complete'&&t.completedDateTime&&new Date(t.completedDateTime).toDateString()===today;
      if (status!=='complete') data.stats.totalOpen++;
      if (doneToday)           data.stats.completedToday++;
      if (status==='late')     data.stats.overdue++;
      if (status==='waiting-on-client') data.stats.waitingOnClient++;
      return { id:t.id, title:t.title, status, clientName:client, bucketName:bucket,
        assigneeNames:aNames, advisor, isUnassigned:aIds.length===0,
        dueDateTime:t.dueDateTime??null, dueDateFormatted:fmtDate(t.dueDateTime),
        lastModifiedDateTime:t.lastModifiedDateTime??null,
        percentComplete:t.percentComplete??0, plannerKey:planner.key,
        notes:null, notesLastModified:null, stuckInfo:null };
    });

    // Fetch notes for active tasks
    const active=tasks.filter(t=>t.status!=='complete');
    console.log(`  Fetching notes for ${active.length} active tasks...`);
    for (const task of active) {
      try {
        await sleep(400);
        const r=await fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${task.id}/details`,{headers:{Authorization:`Bearer ${token}`}});
        if (r.ok) { const d=await r.json(); task.notes=d.description||null; task.notesLastModified=d.lastModifiedDateTime||null; }
        else if (r.status===429) {
          const wait=parseInt(r.headers.get('Retry-After')||'10')*1000;
          await sleep(wait);
          const r2=await fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${task.id}/details`,{headers:{Authorization:`Bearer ${token}`}});
          if (r2.ok) { const d=await r2.json(); task.notes=d.description||null; task.notesLastModified=d.lastModifiedDateTime||null; }
        }
      } catch {}

      task.stuckInfo=detectStuck(task,task.notesLastModified);
      if (task.stuckInfo) console.log(`  STUCK: ${task.title.slice(0,40)} — ${task.stuckInfo.label}`);

      if (task.stuckInfo) {
        if (task.stuckInfo.type==='stuck')    data.stats.stuck++;
        if (task.stuckInfo.type==='comm-gap') data.stats.commGap++;
        data.stuckTasks.push({ plannerKey:planner.key, plannerLabel:planner.label, plannerColor:planner.color,
          title:task.title, clientName:task.clientName, bucketName:task.bucketName,
          assigneeNames:task.assigneeNames, isUnassigned:task.isUnassigned,
          dueDateFormatted:task.dueDateFormatted, stuckInfo:task.stuckInfo, status:task.status });
      }

      if (task.status==='late') {
        data.overdueTasks.push({ plannerKey:planner.key, plannerLabel:planner.label, plannerColor:planner.color,
          title:task.title, clientName:task.clientName, bucketName:task.bucketName,
          assigneeNames:task.assigneeNames, advisor:task.advisor, isUnassigned:task.isUnassigned,
          dueDateFormatted:task.dueDateFormatted, stuckInfo:task.stuckInfo });
      }

      // Team view
      const members=task.assigneeNames.length?task.assigneeNames:['Unassigned'];
      for (const m of members) {
        if (!data.teamView[m]) data.teamView[m]={ name:m, tasks:[], counts:{total:0,overdue:0,stuck:0,commGap:0,waitingClient:0} };
        const tv=data.teamView[m];
        tv.tasks.push({ id:task.id, title:task.title, clientName:task.clientName,
          plannerKey:planner.key, plannerLabel:planner.label, plannerColor:planner.color,
          bucketName:task.bucketName, status:task.status, dueDateFormatted:task.dueDateFormatted,
          lastModifiedDateTime:task.lastModifiedDateTime, stuckInfo:task.stuckInfo, notes:task.notes });
        tv.counts.total++;
        if (task.status==='late')              tv.counts.overdue++;
        if (task.status==='waiting-on-client') tv.counts.waitingClient++;
        if (task.stuckInfo?.type==='stuck')    tv.counts.stuck++;
        if (task.stuckInfo?.type==='comm-gap') tv.counts.commGap++;
      }

      // WIP
      if (task.clientName) {
        if (!data.wip[task.clientName]) data.wip[task.clientName]={
          client:task.clientName, advisor:task.advisor||null, planners:[], byPlanner:{},
          hasOverdue:false, hasWaitingOnClient:false, hasStuck:false, hasCommGap:false,
          totalItems:0, doneItems:0,
        };
        const wip=data.wip[task.clientName];
        if (!wip.advisor&&task.advisor) wip.advisor=task.advisor;
        if (!wip.byPlanner[planner.key]) wip.byPlanner[planner.key]={key:planner.key,label:planner.label,color:planner.color,items:[]};
        wip.byPlanner[planner.key].items.push({ plannerKey:planner.key, title:task.title, status:task.status,
          bucketName:task.bucketName, assigneeNames:task.assigneeNames, isUnassigned:task.isUnassigned,
          dueDateFormatted:task.dueDateFormatted, lastModifiedDateTime:task.lastModifiedDateTime,
          stuckInfo:task.stuckInfo, notes:task.notes });
        if (!wip.planners.includes(planner.key)) wip.planners.push(planner.key);
        if (task.status==='late')              wip.hasOverdue=true;
        if (task.status==='waiting-on-client') wip.hasWaitingOnClient=true;
        if (task.stuckInfo?.type==='stuck')    wip.hasStuck=true;
        if (task.stuckInfo?.type==='comm-gap') wip.hasCommGap=true;
        wip.totalItems++;
        if (task.status==='complete') wip.doneItems++;
      }
    }

    data.planners.push({ key:planner.key, label:planner.label, color:planner.color,
      lastUpdated:new Date().toISOString(), tasks,
      openCount:tasks.filter(t=>t.status!=='complete').length,
      overdueCount:tasks.filter(t=>t.status==='late').length,
      waitingCount:tasks.filter(t=>t.status==='waiting-on-client').length,
      stuckCount:tasks.filter(t=>t.stuckInfo?.type==='stuck').length,
    });
  }

  // Enrich WIP
  const wipArr=Object.values(data.wip);
  data.stats.wipClients=wipArr.filter(w=>w.planners.length>1).length;
  for (const w of wipArr) {
    w.progress=w.totalItems>0?Math.round((w.doneItems/w.totalItems)*100):0;
    w.plannerGroups=Object.values(w.byPlanner);
    w.currentStage=null; w.nextAction=null; w.blockers=null; w.aiUrgency=null;
  }

  // Sort team tasks
  for (const m of Object.values(data.teamView)) {
    m.tasks.sort((a,b)=>{
      const p=t=>t.status==='late'?0:t.stuckInfo?.type==='stuck'?1:t.stuckInfo?.type==='comm-gap'?2:t.status==='waiting-on-client'?3:4;
      return p(a)-p(b);
    });
  }

  // AI grouping
  const allActive=data.planners.flatMap(p=>p.tasks.filter(t=>t.status!=='complete'&&t.clientName));
  if (GITHUB_TOKEN&&allActive.length>0) {
    console.log(`Running AI grouping for ${wipArr.length} clients...`);
    const groups=await aiGroupClients(wipArr,allActive);
    if (groups?.length>0) {
      console.log(`AI returned ${groups.length} groupings`);
      for (const g of groups) {
        const canon=g.canonicalName;
        let ce=wipArr.find(w=>w.client===canon);
        if (g.merges?.length>0) {
          for (const alt of g.merges) {
            const ae=wipArr.find(w=>w.client===alt);
            if (!ae||ae.client===canon) continue;
            console.log(`  Merging "${alt}" → "${canon}"`);
            data.aiMerges.push({from:alt,to:canon});
            if (!ce) { ae.client=canon; ce=ae; }
            else {
              for (const pg of (ae.plannerGroups||[])) {
                const ex=ce.plannerGroups.find(x=>x.key===pg.key);
                if (ex) ex.items.push(...pg.items); else ce.plannerGroups.push(pg);
              }
              for (const p of ae.planners) if (!ce.planners.includes(p)) ce.planners.push(p);
              if (ae.hasOverdue) ce.hasOverdue=true;
              if (ae.hasWaitingOnClient) ce.hasWaitingOnClient=true;
              if (ae.hasStuck) ce.hasStuck=true;
              if (ae.hasCommGap) ce.hasCommGap=true;
              ce.totalItems+=ae.totalItems; ce.doneItems+=ae.doneItems;
              ce.progress=ce.totalItems>0?Math.round((ce.doneItems/ce.totalItems)*100):0;
              ae._merged=true;
            }
          }
        }
        if (ce) {
          ce.currentStage=g.currentStage||null; ce.nextAction=g.nextAction||null;
          ce.blockers=g.blockers||null; ce.aiUrgency=g.urgency||'normal';
          console.log(`  ✓ ${canon}`);
        }
      }
    }
  }

  // Duplicate detection
  const names=wipArr.map(w=>w.client);
  const seen=new Set();
  for (let i=0;i<names.length;i++) for (let j=i+1;j<names.length;j++) {
    const s=nameSim(names[i],names[j]);
    if (s>=0.75&&s<1) { const k=[names[i],names[j]].sort().join('|||'); if (!seen.has(k)) { seen.add(k); data.possibleDuplicates.push({nameA:names[i],nameB:names[j],score:Math.round(s*100)}); } }
  }

  data.wip=wipArr
    .filter(w=>!w._merged&&w.plannerGroups.some(g=>g.items.some(i=>i.status!=='complete')))
    .sort((a,b)=>((b.hasOverdue?4:0)+(b.hasStuck?2:0)+(b.hasCommGap?1:0))-((a.hasOverdue?4:0)+(a.hasStuck?2:0)+(a.hasCommGap?1:0)));

  data.teamView=Object.values(data.teamView)
    .filter(m=>m.counts.total>0)
    .sort((a,b)=>(b.counts.overdue+b.counts.stuck)-(a.counts.overdue+a.counts.stuck));

  // await postToTeams(data.stuckTasks); // Uncomment when Teams webhook is ready

  fs.writeFileSync('planner-data.json',JSON.stringify(data,null,2));
  console.log('Done. Stats:',data.stats);
  console.log(`Stuck: ${data.stats.stuck} | Comm gaps: ${data.stats.commGap} | Stuck tasks flagged: ${data.stuckTasks.length}`);
}

main().catch(err=>{ console.error('Error:',err); process.exit(1); });
