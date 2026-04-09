/**
 * parse-paperwork-forms.js
 * Pulls completed tasks from the Paperwork planner, parses task notes
 * for document names, and outputs a JSON lookup table mapping
 * account type + custodian → required forms.
 *
 * Uses the same auth/fetch pattern as fetch-planner.js.
 *
 * Usage:
 *   node parse-paperwork-forms.js
 *
 * Output:
 *   form-lookup.json  — machine-readable lookup for Phase 2 auto-fill engine
 *   form-lookup.md    — human-readable summary for review
 */

const fs = require('fs');

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const PLAN_ID       = process.env.PLAN_ID_PAPERWORK;

// ── Helpers (same as fetch-planner.js) ───────────────────────────────────────

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
      console.log(`Rate limited — waiting ${wait}ms (attempt ${attempt + 1})`);
      await sleep(wait);
      continue;
    }
    throw new Error(`Graph ${path} → ${res.status}`);
  }
  throw new Error(`Graph ${path} → exceeded retries`);
}

// ── Parsers ───────────────────────────────────────────────────────────────────

function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractAccountType(title, notes) {
  const combined = (title + ' ' + notes).toUpperCase();

  const patterns = [
    [/ROTH\s+IRA/,                    'Roth IRA'],
    [/TRADITIONAL\s+IRA/,             'Traditional IRA'],
    [/ROLLOVER\s+IRA/,                'Rollover IRA'],
    [/SEP\s+IRA/,                     'SEP IRA'],
    [/INHERITED\s+IRA|BENEFICIARY\s+IRA/, 'Inherited/Beneficiary IRA'],
    [/SIMPLE\s+IRA/,                  'SIMPLE IRA'],
    [/401\(K\)|401K/,                 '401(k)'],
    [/403\(B\)|403B/,                 '403(b)'],
    [/INDIVIDUAL\s+TOD|IND\s+TOD/,    'Individual TOD'],
    [/JOINT.*TOD/,                    'Joint TOD'],
    [/JOINT.*SURVIVORSHIP|JTWROS/,    'Joint WROS'],
    [/JOINT.*COMMON/,                 'Joint Tenants in Common'],
    [/UTMA|UGMA/,                     'UTMA/UGMA'],
    [/529/,                           '529'],
    [/TRUST/,                         'Trust'],
    [/INDIVIDUAL/,                    'Individual'],
  ];

  for (const [pattern, label] of patterns) {
    if (pattern.test(combined)) return label;
  }
  return 'Unknown';
}

function extractCustodian(title, notes) {
  const combined = (title + ' ' + notes).toUpperCase();
  if (/FIDELITY|FID[^E]|FIWS|NFS/.test(combined)) return 'Fidelity';
  if (/SCHWAB/.test(combined))                      return 'Schwab';
  if (/WEALTHPORT|CAMBRIDGE/.test(combined))        return 'WealthPort/Cambridge';
  if (/PERSHING/.test(combined))                    return 'Pershing';
  if (/AMF|GUARDIAN/.test(combined))                return 'AMF/Guardian';
  if (/DELAWARE/.test(combined))                    return 'Delaware Life';
  return 'Unknown';
}

function extractDocuments(notes) {
  if (!notes) return [];

  const docs = new Set();

  // Split on common list separators
  const lines = notes.split(/[\n\r•\-\*\/|]+/);

  for (let line of lines) {
    line = line.trim();
    if (line.length < 4) continue;

    // Skip obvious non-document lines
    if (/^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}/.test(line)) continue;  // dates
    if (/^\$[\d,]+/.test(line)) continue;                           // dollar amounts
    if (/^(please|note:|send|sign|complete|submitted|received|client|advisor|per|see|attach)/i.test(line)) continue;

    // Keep lines that look like form or document names
    const docKeywords = [
      'FORM', 'APPLICATION', 'AGREEMENT', 'EXHIBIT', 'ADV',
      'CRS', 'TOD', 'ACH', 'RECEIPT', 'DISCLOSURE',
      'TRANSFER', 'BENEFICIAR', 'SUITABILITY', 'W-9', 'W-8',
      'WEALTHPORT', 'FIDELITY', 'CAMBRIDGE', 'NEW ACCOUNT',
      'IRA', 'WRAP', 'CUSTODIAN', 'DESIGNATION', 'AUTHORIZATION',
      'CHANGE', 'REQUEST', 'UPDATE', 'OPENING', 'BROKERAGE',
    ];

    if (docKeywords.some(kw => line.toUpperCase().includes(kw))) {
      // Clean up and normalize
      const clean = line
        .replace(/^\d+[\.\)]\s*/, '')  // remove leading "1." or "1)"
        .replace(/\s+/g, ' ')
        .trim();
      if (clean.length > 3) docs.add(clean);
    }
  }

  return [...docs];
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  console.log('Authenticating...');
  const token = await getAccessToken();
  console.log('✓ Token acquired');

  console.log('Pulling Paperwork planner tasks...');
  const tasksRes = await graphGet(token, `/planner/plans/${PLAN_ID}/tasks`);
  const allTasks = tasksRes.value ?? [];
  console.log(`  Total tasks: ${allTasks.length}`);

  const completed = allTasks.filter(t => t.percentComplete === 100);
  console.log(`  Completed tasks: ${completed.length}`);

  // Pull notes for each completed task
  console.log('Fetching task notes...');
  const parsed = [];

  for (let i = 0; i < completed.length; i++) {
    const task = completed[i];
    await sleep(150); // same rate-limit courtesy as fetch-planner.js

    let notes = '';
    try {
      const detail = await graphGet(token, `/planner/tasks/${task.id}/details`);
      notes = stripHtml(detail.description || '');
    } catch (e) {
      console.log(`  Could not fetch notes for task ${task.id}: ${e.message}`);
    }

    const accountType = extractAccountType(task.title, notes);
    const custodian   = extractCustodian(task.title, notes);
    const documents   = extractDocuments(notes);

    parsed.push({
      title:       task.title,
      accountType,
      custodian,
      documents,
      notes:       notes.slice(0, 300), // truncated for output
    });

    if ((i + 1) % 20 === 0) console.log(`  Processed ${i + 1}/${completed.length}...`);
  }

  console.log(`✓ Done processing ${completed.length} tasks`);

  // ── Build lookup table ──────────────────────────────────────────────────────
  // Group by accountType + custodian, union all docs seen across tasks
  const lookup = {};

  for (const task of parsed) {
    const key = `${task.accountType}__${task.custodian}`;
    if (!lookup[key]) {
      lookup[key] = {
        accountType: task.accountType,
        custodian:   task.custodian,
        taskCount:   0,
        documents:   new Set(),
        examples:    [],
      };
    }
    lookup[key].taskCount++;
    task.documents.forEach(d => lookup[key].documents.add(d));
    if (lookup[key].examples.length < 3) lookup[key].examples.push(task.title);
  }

  // Convert Sets to sorted arrays
  const lookupArray = Object.values(lookup)
    .map(entry => ({
      accountType: entry.accountType,
      custodian:   entry.custodian,
      taskCount:   entry.taskCount,
      documents:   [...entry.documents].sort(),
      examples:    entry.examples,
    }))
    .sort((a, b) => b.taskCount - a.taskCount);

  // ── Write JSON ──────────────────────────────────────────────────────────────
  fs.writeFileSync('form-lookup.json', JSON.stringify({
    generatedAt: new Date().toISOString(),
    totalTasksAnalyzed: completed.length,
    lookup: lookupArray,
  }, null, 2));
  console.log('✓ form-lookup.json written');

  // ── Write Markdown summary ──────────────────────────────────────────────────
  let md = `# SDC Form Lookup by Account Type\n`;
  md += `_Generated: ${new Date().toLocaleString()} — ${completed.length} completed tasks analyzed_\n\n`;
  md += `> **Review required:** Document names are parsed from task notes. `;
  md += `Verify each row before using in production.\n\n`;

  for (const entry of lookupArray) {
    md += `## ${entry.accountType} — ${entry.custodian}\n`;
    md += `**Tasks seen:** ${entry.taskCount}\n\n`;
    if (entry.documents.length > 0) {
      md += `**Documents found in notes:**\n`;
      entry.documents.forEach(d => { md += `- ${d}\n`; });
    } else {
      md += `_No document names parsed from notes — manual review needed_\n`;
    }
    md += `\n**Example tasks:**\n`;
    entry.examples.forEach(e => { md += `- ${e}\n`; });
    md += '\n---\n\n';
  }

  fs.writeFileSync('form-lookup.md', md);
  console.log('✓ form-lookup.md written');

  // ── Also write raw dump for debugging ──────────────────────────────────────
  fs.writeFileSync('form-lookup-raw.json', JSON.stringify(parsed, null, 2));
  console.log('✓ form-lookup-raw.json written (full task dump for review)');

  console.log(`\nDone. ${lookupArray.length} account type/custodian combinations found.`);
  console.log('Bring form-lookup.json and form-lookup.md back to Claude for cleanup.');
}

main().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
