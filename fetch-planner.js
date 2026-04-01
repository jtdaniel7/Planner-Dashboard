/**
 * fetch-planner.js
 * Pulls tasks from all 4 Planners, resolves assignee names,
 * handles client-dependent buckets, groups WIP by client.
 */

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

// Buckets that are client-dependent — never count as overdue
const CLIENT_DEPENDENT_BUCKETS = [
  'waiting on signatures',
  'waiting on client',
  'pending client signature',
  'sent to client',
];

function isClientDependent(bucketName) {
  if (!bucketName) return false;
  return CLIENT_DEPENDENT_BUCKETS.some(b => bucketName.toLowerCase().includes(b));
}

async function getAccessToken() {
  const url  = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
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

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function graphGet(token, path, retries = 4) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.ok) return res.json();

    if (res.status === 429) {
      // Respect Retry-After header if present, otherwise back off exponentially
      const retryAfter = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = retryAfter > 0 ? retryAfter * 1000 : Math.pow(2, attempt) * 2000;
      console.log(`Rate limited on ${path} — waiting ${wait}ms (attempt ${attempt + 1})`);
      await sleep(wait);
      continue;
    }

    throw new Error(`Graph ${path} → ${res.status}`);
  }
  throw new Error(`Graph ${path} → exceeded retries`);
}

const userCache = {};

async function resolveUserName(token, userId) {
  if (userCache[userId]) return userCache[userId];
  try {
    const data = await graphGet(token, `/users/${userId}`);
    const name = data.displayName || data.userPrincipalName || 'Unknown';
    userCache[userId] = name;
    return name;
  } catch {
    userCache[userId] = 'Unknown';
    return 'Unknown';
  }
}

function deriveStatus(task, bucketName) {
  const pct = task.percentComplete ?? 0;
  if (pct === 100) return 'complete';

  // Client-dependent bucket — never overdue regardless of due date
  if (isClientDependent(bucketName)) return 'waiting-on-client';

  if (task.dueDateTime && new Date(task.dueDateTime) < new Date()) return 'late';
  if (pct > 0) return 'in-progress';
  return 'not-started';
}

function extractClient(title) {
  const match = title.match(/^([^—–\-]+?)\s*[—–\-]/);
  return match ? match[1].trim() : null;
}

function formatDueDate(iso) {
  if (!iso) return null;
  return new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

async function main() {
  console.log('Authenticating with Microsoft Graph...');
  const token = await getAccessToken();

  const plannerData = {
    fetchedAt: new Date().toISOString(),
    planners:  [],
    wip:       {},
    overdueTasks: [],  // flat list of genuinely overdue tasks across all planners
    stats: { totalOpen: 0, completedToday: 0, overdue: 0, waitingOnClient: 0, wipClients: 0 },
  };

  for (const planner of PLANNERS) {
    console.log(`Fetching: ${planner.label}`);
    await sleep(500); // brief pause between planners

    const [tasksRes, bucketsRes] = await Promise.all([
      graphGet(token, `/planner/plans/${planner.planId}/tasks`),
      graphGet(token, `/planner/plans/${planner.planId}/buckets`),
    ]);

    const bucketMap = {};
    for (const b of (bucketsRes.value ?? [])) bucketMap[b.id] = b.name;

    // Pre-resolve all assignee IDs
    const allAssigneeIds = new Set();
    for (const t of (tasksRes.value ?? [])) {
      for (const uid of Object.keys(t.assignments ?? {})) allAssigneeIds.add(uid);
    }
    await Promise.all([...allAssigneeIds].map(uid => resolveUserName(token, uid)));

    const today = new Date().toDateString();
    const plannerLastUpdated = new Date().toISOString();

    const tasks = await Promise.all((tasksRes.value ?? []).map(async t => {
      const bucketName  = bucketMap[t.bucketId] ?? 'Unknown';
      const status      = deriveStatus(t, bucketName);
      const clientName  = extractClient(t.title);
      const assigneeIds = Object.keys(t.assignments ?? {});
      const assigneeNames = assigneeIds.map(uid => userCache[uid] || 'Unknown');
      const completedToday = status === 'complete' && t.completedDateTime &&
        new Date(t.completedDateTime).toDateString() === today;

      if (status !== 'complete')            plannerData.stats.totalOpen++;
      if (completedToday)                   plannerData.stats.completedToday++;
      if (status === 'late')                plannerData.stats.overdue++;
      if (status === 'waiting-on-client')   plannerData.stats.waitingOnClient++;

      // Add to overdue flat list
      if (status === 'late') {
        plannerData.overdueTasks.push({
          plannerKey:   planner.key,
          plannerLabel: planner.label,
          plannerColor: planner.color,
          title:        t.title,
          clientName,
          bucketName,
          assigneeNames,
          isUnassigned: assigneeIds.length === 0,
          dueDateFormatted: formatDueDate(t.dueDateTime),
        });
      }

      // Fetch notes — small delay between calls to avoid rate limiting
      let notes = null;
      try {
        await sleep(200);
        const detail = await graphGet(token, `/planner/tasks/${t.id}/details`);
        notes = detail.description || null;
      } catch { /* optional */ }

      const taskObj = {
        id:               t.id,
        title:            t.title,
        status,
        clientName,
        bucketName,
        assigneeNames,
        isUnassigned:     assigneeIds.length === 0,
        dueDateTime:      t.dueDateTime ?? null,
        dueDateFormatted: formatDueDate(t.dueDateTime),
        completedDateTime: t.completedDateTime ?? null,
        notes,
        percentComplete:  t.percentComplete ?? 0,
      };

      // WIP grouping — keyed by client name
      if (clientName) {
        if (!plannerData.wip[clientName]) {
          plannerData.wip[clientName] = {
            client:   clientName,
            planners: [],
            byPlanner: {},  // items grouped by planner key
            hasOverdue: false,
            hasWaitingOnClient: false,
            totalItems: 0,
            doneItems: 0,
          };
        }
        const wip = plannerData.wip[clientName];
        if (!wip.byPlanner[planner.key]) {
          wip.byPlanner[planner.key] = {
            key:   planner.key,
            label: planner.label,
            color: planner.color,
            items: [],
          };
        }
        wip.byPlanner[planner.key].items.push({
          title:        t.title,
          status,
          bucketName,
          assigneeNames,
          isUnassigned: assigneeIds.length === 0,
          dueDateFormatted: formatDueDate(t.dueDateTime),
          notes,
        });
        if (!wip.planners.includes(planner.key)) wip.planners.push(planner.key);
        if (status === 'late')              wip.hasOverdue = true;
        if (status === 'waiting-on-client') wip.hasWaitingOnClient = true;
        wip.totalItems++;
        if (status === 'complete') wip.doneItems++;
      }

      return taskObj;
    }));

    plannerData.planners.push({
      key:          planner.key,
      label:        planner.label,
      color:        planner.color,
      lastUpdated:  plannerLastUpdated,
      tasks,
      openCount:    tasks.filter(t => t.status !== 'complete').length,
      overdueCount: tasks.filter(t => t.status === 'late').length,
      waitingCount: tasks.filter(t => t.status === 'waiting-on-client').length,
    });
  }

  // Enrich WIP
  const wipArray = Object.values(plannerData.wip);
  plannerData.stats.wipClients = wipArray.filter(w => w.planners.length > 1).length;

  for (const entry of wipArray) {
    entry.progress = entry.totalItems > 0
      ? Math.round((entry.doneItems / entry.totalItems) * 100) : 0;
    // Convert byPlanner object to array for easier rendering
    entry.plannerGroups = Object.values(entry.byPlanner);
  }

  plannerData.wip = wipArray
    .filter(w => w.plannerGroups.some(g => g.items.some(i => i.status !== 'complete')))
    .sort((a, b) => {
      if (a.hasOverdue !== b.hasOverdue) return a.hasOverdue ? -1 : 1;
      return b.planners.length - a.planners.length;
    });

  fs.writeFileSync('planner-data.json', JSON.stringify(plannerData, null, 2));
  console.log('planner-data.json written. Stats:', plannerData.stats);
}

main().catch(err => { console.error('Error:', err); process.exit(1); });
