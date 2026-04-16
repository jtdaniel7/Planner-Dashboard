/**
 * fetch-planner.js
 * Pulls tasks from all 5 Planners, resolves assignee names,
 * normalizes client names, groups WIP by client,
 * and uses Claude AI to assess each client's current stage,
 * next action, and any blockers.
 */

const fs = require('fs');

const TENANT_ID      = process.env.TENANT_ID;
const CLIENT_ID      = process.env.CLIENT_ID;
const CLIENT_SECRET  = process.env.CLIENT_SECRET;
const GITHUB_TOKEN   = process.env.GITHUB_TOKEN;

const PLANNERS = [
  { key: 'trades',    label: 'Trades & Distributions', planId: process.env.PLAN_ID_TRADES,     color: '#2563eb' },
  { key: 'paperwork', label: 'Paperwork',               planId: process.env.PLAN_ID_PAPERWORK,  color: '#d97706' },
  { key: 'advisor',   label: 'Advisor Flow',            planId: process.env.PLAN_ID_ADVISOR,    color: '#16a34a' },
  { key: 'locations', label: 'Locations',               planId: process.env.PLAN_ID_LOCATIONS,  color: '#7c3aed' },
  { key: 'conts',     label: 'CONTs & Checks',          planId: process.env.PLAN_ID_CONTS,      color: '#dc2626' },
];

// Buckets that are client-dependent — never count as overdue
const CLIENT_DEPENDENT_BUCKETS = [
  'waiting on signatures',
  'waiting on client',
  'pending client signature',
  'sent to client',
  'in wip',
  'reassign to acs',
  'reassign to ejd',
  'paperwork submitted',
];

function isClientDependent(bucketName) {
  if (!bucketName) return false;
  return CLIENT_DEPENDENT_BUCKETS.some(b => bucketName.toLowerCase().includes(b));
}

// ── WIP Stage definitions (used in AI prompt) ────────────────────────────────
const WIP_STAGE_CONTEXT = `
You are an operations analyst for SD Capital Advisors, a financial advisory firm.
You understand their exact workflow for client accounts. Here are the sequential stages:

PAPERWORK STAGES (in order):
1. Queue / Task Recognized — paperwork has been requested, not yet started
2. In Progress - Being Pulled — paperwork being prepared at Cambridge
3. Waiting on Signatures — paperwork sent to client via DocuSign or in-office
4. Paperwork Submitted — signed paperwork submitted to custodian (Fidelity, AMF, NFS, etc.)
5. Account Established — custodian has opened the account

TRANSFER/CONTRIBUTION STAGES (after account established):
6. Transfer Submitted — TOA or rollover submitted to delivering firm
7. Waiting on Funds — waiting for transfer to complete (can take 10-30 business days)
8. Contribution Processed — ACH/ICP contribution or check deposited

INVESTMENT STAGES (after funds arrive):
9. Trades / Link to Model — advisor needs to place trades or link to investment model
10. Fee Billing Setup — ensure fee schedule is configured correctly

CLEANUP STAGES:
11. RT/CSTAT Update — Redtail and CSTAT records updated
12. Online Access / Household — client portal access and household setup
13. New Client Onboarding — if new client, onboarding team assigned
14. Complete — all steps done, WIP entry can be deleted

SPECIAL CONDITIONS:
- "Waiting on Signatures" and "Paperwork Submitted" are CLIENT-DEPENDENT — not overdue even if past due date
- "In WIP / Reassign to ACS/EJD" = client-dependent, not a team failure
- Fidelity BL accounts route through CIR for RIA signature before going to Fidelity
- Inherited/Bene IRA accounts need: RMD tracking, tag group, decedent info in Redtail
- GBU/Insurance accounts need: FedEx tracking, policy approval, DocuFast delivery, maturity date reminder
- Transfers from accounts starting with "Y" (Fidelity Bene accounts) require opening a new account first

Your job: Given a client's active tasks across multiple planners, determine:
1. currentStage — the most accurate single-sentence description of where this client is right now
2. nextAction — the single most important next step the ops team needs to take
3. blockers — any issues preventing forward progress (missing info, waiting on client, pending advisor decision)
4. urgency — "high" (overdue or blocking), "normal" (on track), or "waiting" (client-dependent, nothing ops can do)
`;

// ── AI Client Grouping ───────────────────────────────────────────────────────
/**
 * Single AI call that takes all active tasks across all planners
 * and returns an intelligent parent/child grouping by client.
 * Handles name variations (Becca/Rebecca, Murphy/Murphy Brandy etc.)
 * and produces a clean canonical client name for each group.
 */
async function aiGroupClients(wipArray, allTasks) {
  if (!GITHUB_TOKEN) return null;

  // Build compact client list — only clients worth AI analysis
  // Hard cap at 60 to stay well under 8000 token limit
  const clientSummaries = {};
  for (const t of allTasks.filter(t => t.status !== 'complete' && t.clientName)) {
    const name = t.clientName;
    if (!clientSummaries[name]) clientSummaries[name] = { planners: new Set(), buckets: [], statuses: [] };
    clientSummaries[name].planners.add(t.plannerKey);
    clientSummaries[name].buckets.push(t.bucketName);
    clientSummaries[name].statuses.push(t.status);
  }

  const allCompact = Object.entries(clientSummaries).map(([name, data]) => ({
    name,
    planners:  [...data.planners],
    buckets:   [...new Set(data.buckets)].slice(0, 3), // max 3 buckets
    overdue:   data.statuses.includes('late'),
    taskCount: data.buckets.length,
  }));

  if (!allCompact.length) return null;

  // Only send clients that genuinely need AI attention:
  // 1. Overdue clients
  // 2. Multi-planner clients (likely household grouping candidates)
  // 3. Clients with unusual name formats that may be duplicates
  const priority = allCompact
    .filter(c => c.overdue || c.planners.length > 1)
    .sort((a, b) => (b.overdue ? 1 : 0) - (a.overdue ? 1 : 0))
    .slice(0, 60); // hard cap at 60

  // Also include all client names for duplicate detection only (just names, no details)
  const allNames = allCompact.map(c => c.name);

  const toProcess = priority;

  const prompt = `You are reviewing active client work at SD Capital Advisors, a financial advisory firm.

Priority clients (overdue or multi-planner) — analyze for grouping and stage:
${JSON.stringify(toProcess, null, 1)}

All client names (for duplicate/nickname detection only):
${allNames.join(', ')}

Your job — TWO things only:
1. MERGE: Identify names that refer to the same person/household. Common patterns:
   - Same last name, different first name format: "Becca Ferguson" = "Ferguson, Rebecca"
   - Couple/household tasks: "Kuch, Earl" + "Reineke-Kuch, Donna" + "Kuch, Earl & Reineke-Kuch, donna" = same household
   - Nickname vs legal name: "Bob" = "Robert", "Becca" = "Rebecca", "Mike" = "Michael"
   - Format differences: "SMITH JOHN" = "Smith, John"

2. ASSESS: For merged groups and multi-planner clients, provide a brief stage summary.

Respond with ONLY a JSON array of clients needing merges or with meaningful stage info.
Skip clients that are clearly standalone and on track.
Max 50 entries in response.

[
  {
    "canonicalName": "Kuch, Earl & Donna",
    "merges": ["Kuch, Earl", "Reineke-Kuch, Donna", "Kuch, Earl & Reineke-Kuch, Donna 03/23"],
    "currentStage": "Multiple items across Paperwork and Advisor Flow",
    "nextAction": "Review overdue paperwork items",
    "blockers": null,
    "urgency": "high"
  }
]`;

  try {
    const res = await fetch('https://models.inference.ai.azure.com/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${GITHUB_TOKEN}`,
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          { role: 'system', content: WIP_STAGE_CONTEXT },
          { role: 'user',   content: prompt },
        ],
        max_tokens: 2000,
        temperature: 0.1,
        response_format: { type: 'json_object' },
      }),
    });

    if (!res.ok) {
      const err = await res.text();
      console.log(`AI grouping failed: ${res.status} ${err}`);
      return null;
    }

    const data = await res.json();
    let text = data.choices?.[0]?.message?.content || '[]';
    text = text.replace(/```json|```/g, '').trim();

    // GPT-4o-mini with json_object returns an object, extract array if wrapped
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : (parsed.clients || parsed.groups || Object.values(parsed)[0] || []);
  } catch (e) {
    console.log(`AI grouping parse error: ${e.message}`);
    return null;
  }
}

// ── Name normalization ────────────────────────────────────────────────────────
function normalizeName(raw) {
  if (!raw) return null;
  let name = raw.trim();

  name = name
    .replace(/\s+(TOD|NQ|IRA|ROTH|401K|403B|529|UTMA|BL|NW|FID|AMF|NFS|BES|FJD|ZWB|CMB|KSO|ACS|JEK|EJD|BAW|MRB)\b.*/i, '')
    .replace(/\s+[A-Z]{2,4}$/, '')
    .trim();

  if (name.includes(',')) {
    const parts = name.split(',').map(p => p.trim());
    const last  = toTitleCase(parts[0]);
    const first = toTitleCase(parts.slice(1).join(',').trim());
    return first ? `${last}, ${first}` : last;
  }

  const words = name.split(/\s+/).filter(Boolean);
  if (words.length === 1) return toTitleCase(words[0]);
  if (words.length === 2) return `${toTitleCase(words[1])}, ${toTitleCase(words[0])}`;

  const ampIdx = words.findIndex(w => w === '&' || w === 'and');
  if (ampIdx !== -1) return words.map(toTitleCase).join(' ');

  const last  = toTitleCase(words[words.length - 1]);
  const first = words.slice(0, -1).map(toTitleCase).join(' ');
  return `${last}, ${first}`;
}

function toTitleCase(str) {
  if (!str) return '';
  return str.toLowerCase().replace(/(?:^|[-\s])(\w)/g, c => c.toUpperCase());
}

function nameSimilarity(a, b) {
  if (!a || !b) return 0;
  const na = a.toLowerCase().replace(/[^a-z]/g, '');
  const nb = b.toLowerCase().replace(/[^a-z]/g, '');
  if (na === nb) return 1;

  const lastA = a.split(',')[0]?.toLowerCase().trim() || '';
  const lastB = b.split(',')[0]?.toLowerCase().trim() || '';
  if (lastA !== lastB) return 0;

  const firstA = (a.split(',')[1] || '').toLowerCase().trim().split(/\s+/)[0] || '';
  const firstB = (b.split(',')[1] || '').toLowerCase().trim().split(/\s+/)[0] || '';
  if (!firstA || !firstB) return 0;

  if (firstA.includes(firstB) || firstB.includes(firstA)) return 0.85;

  const dist = levenshtein(firstA, firstB);
  const maxLen = Math.max(firstA.length, firstB.length);
  const similarity = 1 - dist / maxLen;
  return similarity > 0.5 ? similarity : 0;
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({length: m+1}, (_, i) =>
    Array.from({length: n+1}, (_, j) => i === 0 ? j : j === 0 ? i : 0));
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i-1] === b[j-1]
        ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[m][n];
}

// ── Auth ──────────────────────────────────────────────────────────────────────
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
      const retryAfter = parseInt(res.headers.get('Retry-After') || '0', 10);
      const wait = retryAfter > 0 ? retryAfter * 1000 : Math.pow(2, attempt) * 2000;
      console.log(`Rate limited on ${path} — waiting ${wait}ms`);
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
    const data = await graphGet(token, `/users/${userId}?$select=displayName,userPrincipalName,mail`);
    let name = data.displayName;
    if (!name || name === 'Unknown') {
      const email = data.mail || data.userPrincipalName || '';
      name = email.split('@')[0].replace(/[._]/g, ' ') || 'Unknown';
    }
    userCache[userId] = name;
    return name;
  } catch (e) {
    console.log(`Could not resolve user ${userId}: ${e.message}`);
    userCache[userId] = 'Unknown';
    return 'Unknown';
  }
}

function deriveStatus(task, bucketName) {
  const pct = task.percentComplete ?? 0;
  if (pct === 100) return 'complete';
  if (isClientDependent(bucketName)) return 'waiting-on-client';
  if (task.dueDateTime && new Date(task.dueDateTime) < new Date()) return 'late';
  if (pct > 0) return 'in-progress';
  return 'not-started';
}

function extractClient(title, plannerKey) {
  if (!title) return null;

  const emDash = title.match(/^(.+?)\s*[—–]\s*.+/);
  if (emDash) return emDash[1].trim();

  if (['paperwork', 'locations', 'advisor'].includes(plannerKey)) {
    const hyphen = title.match(/^(.+?)\s+-\s+.{2,}/);
    if (hyphen) return hyphen[1].trim();
    const pipe = title.match(/^(.+?)\s*\|\|\s*.+/);
    if (pipe) return pipe[1].trim();
  }

  if (['trades', 'conts'].includes(plannerKey)) {
    const words = title.trim().split(/\s+/);
    if (words.length >= 2) return `${words[0]} ${words[1]}`;
    return words[0] || null;
  }

  return null;
}

function extractAdvisor(assigneeNames) {
  const ADVISOR_KEYWORDS = ['brent', 'frank', 'zach', 'cheyenne', 'katie', 'melissa', 'elizabeth', 'jaiden'];
  for (const name of assigneeNames) {
    const lower = name.toLowerCase();
    if (ADVISOR_KEYWORDS.some(k => lower.includes(k))) return name;
  }
  return null;
}

function formatDueDate(iso) {
  if (!iso) return null;
  return new Date(iso).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  console.log('Authenticating with Microsoft Graph...');
  const token = await getAccessToken();

  const plannerData = {
    fetchedAt:          new Date().toISOString(),
    planners:           [],
    wip:                {},
    overdueTasks:       [],
    possibleDuplicates: [],
    aiMerges:           [], // { from: "alt name", to: "canonical name" }
    stats: { totalOpen: 0, completedToday: 0, overdue: 0, waitingOnClient: 0, wipClients: 0 },
  };

  for (const planner of PLANNERS) {
    console.log(`Fetching: ${planner.label}`);
    await sleep(500);

    const [tasksRes, bucketsRes] = await Promise.all([
      graphGet(token, `/planner/plans/${planner.planId}/tasks`),
      graphGet(token, `/planner/plans/${planner.planId}/buckets`),
    ]);

    const bucketMap = {};
    for (const b of (bucketsRes.value ?? [])) bucketMap[b.id] = b.name;

    const allAssigneeIds = new Set();
    for (const t of (tasksRes.value ?? [])) {
      for (const uid of Object.keys(t.assignments ?? {})) allAssigneeIds.add(uid);
    }
    await Promise.all([...allAssigneeIds].map(uid => resolveUserName(token, uid)));

    const today = new Date().toDateString();
    const plannerLastUpdated = new Date().toISOString();

    const tasks = await Promise.all((tasksRes.value ?? []).map(async t => {
      const bucketName    = bucketMap[t.bucketId] ?? 'Unknown';
      const status        = deriveStatus(t, bucketName);
      const rawClient     = extractClient(t.title, planner.key);
      const clientName    = rawClient ? normalizeName(rawClient) : null;
      const assigneeIds   = Object.keys(t.assignments ?? {});
      const assigneeNames = assigneeIds.map(uid => userCache[uid] || 'Unknown');
      const advisor       = extractAdvisor(assigneeNames);
      const completedToday = status === 'complete' && t.completedDateTime &&
        new Date(t.completedDateTime).toDateString() === today;

      if (status !== 'complete')          plannerData.stats.totalOpen++;
      if (completedToday)                 plannerData.stats.completedToday++;
      if (status === 'late')              plannerData.stats.overdue++;
      if (status === 'waiting-on-client') plannerData.stats.waitingOnClient++;

      if (status === 'late') {
        plannerData.overdueTasks.push({
          plannerKey:       planner.key,
          plannerLabel:     planner.label,
          plannerColor:     planner.color,
          title:            t.title,
          clientName,
          bucketName,
          assigneeNames,
          advisor,
          isUnassigned:     assigneeIds.length === 0,
          dueDateFormatted: formatDueDate(t.dueDateTime),
        });
      }

      // Notes fetched separately after all tasks are processed (see below)
      let notes = null;

      const taskObj = {
        id:                t.id,
        title:             t.title,
        status,
        clientName,
        rawClientName:     rawClient,
        bucketName,
        assigneeNames,
        advisor,
        isUnassigned:      assigneeIds.length === 0,
        dueDateTime:       t.dueDateTime ?? null,
        dueDateFormatted:  formatDueDate(t.dueDateTime),
        completedDateTime: t.completedDateTime ?? null,
        notes,
        percentComplete:   t.percentComplete ?? 0,
      };

      if (clientName) {
        if (!plannerData.wip[clientName]) {
          plannerData.wip[clientName] = {
            client:             clientName,
            advisor:            advisor || null,
            planners:           [],
            byPlanner:          {},
            hasOverdue:         false,
            hasWaitingOnClient: false,
            totalItems:         0,
            doneItems:          0,
          };
        }
        const wip = plannerData.wip[clientName];
        if (!wip.advisor && advisor) wip.advisor = advisor;

        if (!wip.byPlanner[planner.key]) {
          wip.byPlanner[planner.key] = {
            key:   planner.key,
            label: planner.label,
            color: planner.color,
            items: [],
          };
        }
        wip.byPlanner[planner.key].items.push({
          plannerKey:       planner.key,
          title:            t.title,
          status,
          bucketName,
          assigneeNames,
          advisor,
          isUnassigned:     assigneeIds.length === 0,
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

    // Fetch notes only for active tasks — with respectful rate limiting
    const activeTasks = tasks.filter(t => t.status !== 'complete');
    console.log(`  Fetching notes for ${activeTasks.length} active tasks in ${planner.label}...`);
    for (const task of activeTasks) {
      try {
        await sleep(400); // respectful delay — slower but avoids 429s
        const notesRes = await fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${task.id}/details`, {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (notesRes.ok) {
          const detail = await notesRes.json();
          task.notes = detail.description || null;

          // Also update the notes in WIP byPlanner items
          if (task.clientName && plannerData.wip[task.clientName]?.byPlanner[planner.key]) {
            const wipItem = plannerData.wip[task.clientName].byPlanner[planner.key].items
              .find(i => i.title === task.title);
            if (wipItem) wipItem.notes = task.notes;
          }
        } else if (notesRes.status === 429) {
          const retryAfter = parseInt(notesRes.headers.get('Retry-After') || '10', 10);
          console.log(`  Rate limited on notes — waiting ${retryAfter}s before continuing...`);
          await sleep(retryAfter * 1000);
          // Retry once after waiting
          const retry = await fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${task.id}/details`, {
            headers: { Authorization: `Bearer ${token}` },
          });
          if (retry.ok) {
            const detail = await retry.json();
            task.notes = detail.description || null;
            if (task.clientName && plannerData.wip[task.clientName]?.byPlanner[planner.key]) {
              const wipItem = plannerData.wip[task.clientName].byPlanner[planner.key].items
                .find(i => i.title === task.title);
              if (wipItem) wipItem.notes = task.notes;
            }
          }
        }
      } catch (e) {
        console.log(`  Notes fetch error for ${task.id.slice(0,8)}: ${e.message}`);
      }
    }

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

  // ── Enrich WIP ──────────────────────────────────────────────────────────────
  const wipArray = Object.values(plannerData.wip);
  plannerData.stats.wipClients = wipArray.filter(w => w.planners.length > 1).length;

  for (const entry of wipArray) {
    entry.progress      = entry.totalItems > 0 ? Math.round((entry.doneItems / entry.totalItems) * 100) : 0;
    entry.plannerGroups = Object.values(entry.byPlanner);
  }

  // ── AI Client Grouping & Stage Assessment (single call) ────────────────────
  // Collect all active tasks across all planners for the AI to process at once
  const allActiveTasks = plannerData.planners.flatMap(p =>
    p.tasks
      .filter(t => t.status !== 'complete' && t.clientName)
      .map(t => ({ ...t, plannerKey: p.key }))
  );

  // Initialize all WIP entries with null AI fields
  wipArray.forEach(entry => {
    entry.currentStage = null;
    entry.nextAction   = null;
    entry.blockers     = null;
    entry.aiUrgency    = null;
  });

  if (GITHUB_TOKEN && allActiveTasks.length > 0) {
    const priorityCount = allActiveTasks.filter(t => t.status !== 'complete' && t.clientName).length;
    console.log(`Running AI grouping — ${wipArray.length} total clients, sending ${Math.min(60, wipArray.filter(w => w.hasOverdue || w.planners.length > 1).length)} priority clients...`);
    const groupings = await aiGroupClients(wipArray, allActiveTasks);

    if (groupings && groupings.length > 0) {
      console.log(`AI returned ${groupings.length} client groupings`);

      for (const group of groupings) {
        const canonical = group.canonicalName;

        // Apply merges — remap any merged names to the canonical name
        if (group.merges && group.merges.length > 0) {
          let canonEntry = wipArray.find(w => w.client === canonical);

          for (const altName of group.merges) {
            const altEntry = wipArray.find(w => w.client === altName);
            if (!altEntry) continue;
            if (altEntry.client === canonical) continue;

            console.log(`  Merging "${altName}" → "${canonical}"`);
            plannerData.aiMerges.push({ from: altName, to: canonical });

            if (!canonEntry) {
              // Canonical name is new — rename first alt entry to canonical
              altEntry.client = canonical;
              canonEntry = altEntry;
            } else {
              // Merge alt into canonical
              for (const pg of (altEntry.plannerGroups || [])) {
                const existing = canonEntry.plannerGroups.find(g => g.key === pg.key);
                if (existing) existing.items.push(...pg.items);
                else canonEntry.plannerGroups.push(pg);
              }
              for (const p of altEntry.planners) {
                if (!canonEntry.planners.includes(p)) canonEntry.planners.push(p);
              }
              if (altEntry.hasOverdue)         canonEntry.hasOverdue = true;
              if (altEntry.hasWaitingOnClient) canonEntry.hasWaitingOnClient = true;
              canonEntry.totalItems += altEntry.totalItems;
              canonEntry.doneItems  += altEntry.doneItems;
              // Recalculate progress
              canonEntry.progress = canonEntry.totalItems > 0
                ? Math.round((canonEntry.doneItems / canonEntry.totalItems) * 100) : 0;
              altEntry._merged = true;
            }
          }

          // Apply AI assessment to canonical entry
          if (canonEntry) {
            canonEntry.currentStage = group.currentStage || canonEntry.currentStage;
            canonEntry.nextAction   = group.nextAction   || canonEntry.nextAction;
            canonEntry.blockers     = group.blockers      || canonEntry.blockers;
            canonEntry.aiUrgency    = group.urgency       || canonEntry.aiUrgency;
          }
        }

        // Apply AI stage assessment to the canonical entry
        const entry = wipArray.find(w => w.client === canonical);
        if (entry) {
          entry.currentStage = group.currentStage || null;
          entry.nextAction   = group.nextAction   || null;
          entry.blockers     = group.blockers      || null;
          entry.aiUrgency    = group.urgency       || 'normal';
          console.log(`  ✓ ${canonical}: ${group.currentStage || 'grouped'}`);
        }
      }
    }
  } else {
    console.log('GITHUB_TOKEN not available — skipping AI grouping');
  }

  // ── Duplicate detection ─────────────────────────────────────────────────────
  const clientNames = wipArray.map(w => w.client);
  const seenPairs   = new Set();
  for (let i = 0; i < clientNames.length; i++) {
    for (let j = i + 1; j < clientNames.length; j++) {
      const score = nameSimilarity(clientNames[i], clientNames[j]);
      if (score >= 0.75 && score < 1) {
        const key = [clientNames[i], clientNames[j]].sort().join('|||');
        if (!seenPairs.has(key)) {
          seenPairs.add(key);
          plannerData.possibleDuplicates.push({
            nameA: clientNames[i],
            nameB: clientNames[j],
            score: Math.round(score * 100),
          });
        }
      }
    }
  }

  // ── Sort & write ────────────────────────────────────────────────────────────
  plannerData.wip = wipArray
    .filter(w => !w._merged && w.plannerGroups.some(g => g.items.some(i => i.status !== 'complete')))
    .sort((a, b) => {
      // AI urgency sort: high → normal → waiting
      const urgencyOrder = { high: 0, normal: 1, waiting: 2 };
      const ua = urgencyOrder[a.aiUrgency] ?? 1;
      const ub = urgencyOrder[b.aiUrgency] ?? 1;
      if (ua !== ub) return ua - ub;
      // Then overdue
      if (a.hasOverdue !== b.hasOverdue) return a.hasOverdue ? -1 : 1;
      // Then multi-planner
      return b.planners.length - a.planners.length;
    });

  fs.writeFileSync('planner-data.json', JSON.stringify(plannerData, null, 2));
  console.log('planner-data.json written. Stats:', plannerData.stats);
  if (plannerData.possibleDuplicates.length) {
    console.log(`Possible duplicates: ${plannerData.possibleDuplicates.length}`);
  }
}

main().catch(err => { console.error('Error:', err); process.exit(1); });
