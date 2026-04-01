const fs = require('fs');

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const PLANNERS = [
  { key: 'trades',    label: 'Trades & Distributions', planId: process.env.PLAN_ID_TRADES,     color: '#4f9cf9' },
  { key: 'paperwork', label: 'Paperwork',               planId: process.env.PLAN_ID_PAPERWORK,  color: '#f9a84f' },
  { key: 'advisor',   label: 'Advisor Flow',            planId: process.env.PLAN_ID_ADVISOR,    color: '#7fd8a0' },
  { key: 'locations', label: 'Locations',               planId: process.env.PLAN_ID_LOCATIONS,  color: '#c084fc' },
];

async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });
  const res  = await fetch(url, { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function graphGet(token, path) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph ${path} → ${res.status}`);
  return res.json();
}

function deriveStatus(task) {
  const pct = task.percentComplete ?? 0;
  if (pct === 100) return 'complete';
  if (task.dueDateTime && new Date(task.dueDateTime) < new Date()) return 'late';
  if (pct > 0) return 'in-progress';
  return 'not-started';
}

function extractClient(title) {
  const match = title.match(/^([^—–-]+?)\s*[—–-]/);
  return match ? match[1].trim() : null;
}

async function main() {
  console.log('Authenticating with Microsoft Graph...');
  const token = await getAccessToken();

  const plannerData = {
    fetchedAt: new Date().toISOString(),
    planners:  [],
    wip:       {},
    stats: { totalOpen: 0, completedToday: 0, overdue: 0, wipClients: 0 },
  };

  for (const planner of PLANNERS) {
    console.log(`Fetching: ${planner.label}`);
    const [tasksRes, bucketsRes] = await Promise.all([
      graphGet(token, `/planner/plans/${planner.planId}/tasks`),
      graphGet(token, `/planner/plans/${planner.planId}/buckets`),
    ]);

    const bucketMap = {};
    for (const b of (bucketsRes.value ?? [])) bucketMap[b.id] = b.name;

    const today = new Date().toDateString();
    const tasks = (tasksRes.value ?? []).map(t => {
      const status = deriveStatus(t);
      const clientName = extractClient(t.title);
      const completedToday = status === 'complete' && t.completedDateTime &&
        new Date(t.completedDateTime).toDateString() === today;

      if (status !== 'complete') plannerData.stats.totalOpen++;
      if (completedToday)        plannerData.stats.completedToday++;
      if (status === 'late')     plannerData.stats.overdue++;

      if (clientName) {
        if (!plannerData.wip[clientName]) {
          plannerData.wip[clientName] = { client: clientName, items: [], planners: [] };
        }
        const wip = plannerData.wip[clientName];
        wip.items.push({ plannerKey: planner.key, title: t.title, status, bucketName: bucketMap[t.bucketId] ?? 'Unknown', dueDateTime: t.dueDateTime ?? null });
        if (!wip.planners.includes(planner.key)) wip.planners.push(planner.key);
      }

      return {
        id: t.id, title: t.title, status, clientName,
        bucketName: bucketMap[t.bucketId] ?? 'Unknown',
        assignees: Object.keys(t.assignments ?? {}),
        dueDateTime: t.dueDateTime ?? null,
        completedDateTime: t.completedDateTime ?? null,
      };
    });

    plannerData.planners.push({
      key: planner.key, label: planner.label, color: planner.color, tasks,
      openCount: tasks.filter(t => t.status !== 'complete').length,
    });
  }

  const wipArray = Object.values(plannerData.wip);
  plannerData.stats.wipClients = wipArray.filter(w => w.planners.length > 1).length;

  for (const entry of wipArray) {
    const total = entry.items.length;
    const done  = entry.items.filter(i => i.status === 'complete').length;
    entry.progress  = total > 0 ? Math.round((done / total) * 100) : 0;
    entry.hasOverdue = entry.items.some(i => i.status === 'late');
  }

  plannerData.wip = wipArray
    .filter(w => w.items.some(i => i.status !== 'complete'))
    .sort((a, b) => {
      if (a.hasOverdue !== b.hasOverdue) return a.hasOverdue ? -1 : 1;
      return b.items.length - a.items.length;
    });

  fs.writeFileSync('planner-data.json', JSON.stringify(plannerData, null, 2));
  console.log('planner-data.json written. Stats:', plannerData.stats);
}

main().catch(err => { console.error('Error:', err); process.exit(1); });
