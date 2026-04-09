/**
 * SDC Capital Advisors — Paperwork Automation Engine
 * Phase 2: fill-forms.js
 *
 * Flow:
 *   1. Poll Planner for new Paperwork tasks
 *   2. Parse task note → extract client/account/suitability fields
 *   3. Look up required forms from Form Lookup table
 *   4. Pull client data from Redtail API
 *   5. Fill blank PDFs using field mapping
 *   6. Save completed packet + checklist to SharePoint
 *   7. Comment on Planner task with summary
 */

import { execSync } from "child_process";
import { readFileSync, writeFileSync, existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));

// ─── CONFIG ─────────────────────────────────────────────────────────────────
const TENANT_ID          = process.env.AZURE_TENANT_ID;
const CLIENT_ID          = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET      = process.env.AZURE_CLIENT_SECRET;
const REDTAIL_API_KEY    = process.env.REDTAIL_API_KEY;
const PLAN_ID            = process.env.PLANNER_PLAN_ID || "cZbzO4Musk2xh4XhQBvfK2UACK_v";
const SHAREPOINT_SITE    = process.env.SHAREPOINT_SITE_URL || "https://sdcapitale3.sharepoint.com/sites/SDC%20%202";
const SHAREPOINT_FOLDER  = process.env.SHAREPOINT_FOLDER_PATH || "/OPERATIONS/Doc Ops/JTs Sandbox";
const FORMS_DIR          = join(__dirname, "../forms/blank");
const PROCESSED_LOG      = join(__dirname, "processed-tasks.json");

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function log(msg) { console.log(`[${new Date().toISOString()}] ${msg}`); }

async function getToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function graphGet(token, path) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph GET ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPost(token, path, body) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Graph POST ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPut(token, path, buffer, contentType = "application/pdf") {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": contentType },
    body: buffer,
  });
  if (!res.ok) throw new Error(`Graph PUT ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

// ─── REDTAIL ─────────────────────────────────────────────────────────────────
async function redtailSearch(name) {
  const auth = Buffer.from(`key:${REDTAIL_API_KEY}`).toString("base64");
  const res = await fetch(
    `https://api.redtailtechnology.com/crm/v1/contacts?keyword=${encodeURIComponent(name)}`,
    { headers: { Authorization: `Basic ${auth}`, Accept: "application/json" } }
  );
  if (!res.ok) throw new Error(`Redtail search failed: ${res.status}`);
  const data = await res.json();
  const contacts = data.contacts || data.data || [];
  if (!contacts.length) throw new Error(`No Redtail contact found for: ${name}`);
  return contacts[0];
}

async function redtailGetContact(contactId) {
  const auth = Buffer.from(`key:${REDTAIL_API_KEY}`).toString("base64");
  const res = await fetch(
    `https://api.redtailtechnology.com/crm/v1/contacts/${contactId}`,
    { headers: { Authorization: `Basic ${auth}`, Accept: "application/json" } }
  );
  if (!res.ok) throw new Error(`Redtail contact fetch failed: ${res.status}`);
  return res.json();
}

function extractRedtailFields(contact) {
  // Normalize Redtail contact fields to our standard mapping keys
  const addr = contact.addresses?.[0] || {};
  const phone = contact.phones?.find(p => p.phone_type === "Mobile")
             || contact.phones?.find(p => p.phone_type === "Home")
             || contact.phones?.[0] || {};
  const email = contact.emails?.[0] || {};

  return {
    firstName:   contact.first_name || "",
    middleName:  contact.middle_name || "",
    lastName:    contact.last_name || "",
    fullName:    [contact.first_name, contact.middle_name, contact.last_name].filter(Boolean).join(" "),
    SSN:         contact.tax_id || "",
    DOB:         contact.birth_date || "",
    address:     addr.street_address || "",
    address2:    addr.secondary_address || "",
    city:        addr.city || "",
    state:       addr.state || "",
    zip:         addr.zip || "",
    country:     addr.country || "USA",
    phone:       phone.number || "",
    email:       email.address || "",
    citizenship: "USA",
  };
}

// ─── TASK NOTE PARSER ────────────────────────────────────────────────────────
function parseTaskNote(note) {
  const get = (label) => {
    const regex = new RegExp(`${label}[:\\s]+([^\\n]+)`, "i");
    const match = note.match(regex);
    return match ? match[1].trim() : "";
  };

  const custodian = get("Sponsor Company");
  const accountType = get("Type of Account");
  const paperworkType = get("Type of Paperwork Being Pulled");
  const clientName = get("Name of Client");
  const fundingMethod = get("How the Account being funded");
  const bankLinking = get("Are we linking the bank").toLowerCase().includes("yes");
  const hasVoidedCheck = get("Do we have a voided check").toLowerCase().includes("yes");
  const fullSuitability = get("Full Suitability Needed").toLowerCase().includes("yes");
  const signatureMethod = get("Signature Method");
  const isNewClient = get("Is this a new client").toLowerCase().includes("yes");
  const isMinor = get("Is this a Minor Custodial Account").toLowerCase().includes("yes");
  const beneficiaries = get("Beneficiaries");
  const maintenanceType = get("Type of Maintenance");

  // Financial/suitability data
  const annualIncome = get("Household Annual Income") || get("Annual Income");
  const netWorth = get("Net Worth");
  const taxBracket = get("Tax Bracket");
  const netInvestableAssets = get("Net Investable Assets");
  const riskTolerance = get("Risk Tolerance");
  const primaryObjective = get("Primary Objective");
  const timeHorizon = get("Account Time Horizon");
  const annualExpenses = get("Annual Expenses");
  const estimatedSpecialExpenses = get("Estimated Special Expenses");
  const specialExpenseTimeline = get("Special Expense Timeline");

  // Account numbers from note (maintenance forms)
  const accountNumbers = [...note.matchAll(/\b\d{9,12}\b/g)].map(m => m[0]);

  return {
    clientName,
    custodian,
    accountType,
    paperworkType,
    fundingMethod,
    bankLinking,
    hasVoidedCheck,
    fullSuitability,
    signatureMethod,
    isNewClient,
    isMinor,
    beneficiaries,
    maintenanceType,
    accountNumbers,
    financial: {
      annualIncome,
      netWorth,
      taxBracket,
      netInvestableAssets,
      riskTolerance,
      primaryObjective,
      timeHorizon,
      annualExpenses,
      estimatedSpecialExpenses,
      specialExpenseTimeline,
    },
    rawNote: note,
  };
}

// ─── FORM LOOKUP ─────────────────────────────────────────────────────────────
function getRequiredForms(parsed) {
  const { custodian, accountType, fundingMethod, bankLinking, isNewClient, paperworkType, maintenanceType } = parsed;
  const forms = [];
  const flags = [];

  const isFidelity = custodian?.toUpperCase().includes("FID");
  const isIRA = /IRA|ROTH|ROLLOVER|SEP|SIMPLE/i.test(accountType + " " + custodian);
  const isNonretirement = /INDIVIDUAL|JOINT|TOD|JTOD|UTMA|UGMA/i.test(accountType + " " + custodian);
  const isNewAccount = /new account/i.test(paperworkType);
  const isMaintenance = /maintenance/i.test(paperworkType);
  const isTransfer = /transfer/i.test(fundingMethod);
  const isRothConversion = /roth conversion/i.test(maintenanceType + " " + custodian);
  const isBIA = /BIA/i.test(maintenanceType);

  if (isFidelity) {
    if (isNewAccount) {
      if (isIRA) {
        forms.push({ file: "Premier Select IRA App.pdf", label: "Fidelity Premiere Select IRA Application" });
        if (bankLinking) {
          forms.push({ file: "premiere_select_standing_payment_instruction_iraspi.pdf", label: "Fidelity SPI IRA" });
          flags.push("⚠️  Bank routing/account number — NEEDS MANUAL INPUT (from voided check)");
        }
      }

      if (isNonretirement) {
        forms.push({ file: "Brokerage App.pdf", label: "Fidelity Brokerage Account Application" });
        forms.push({ file: "TOD Registration.pdf", label: "Fidelity TOD Bene Designation" });
        if (bankLinking) {
          forms.push({ file: "standing_payment_instructions_nonretirement_stand.pdf", label: "Fidelity SPI Nonretirement" });
          flags.push("⚠️  Bank routing/account number — NEEDS MANUAL INPUT (from voided check)");
        }
      }

      // New account disclosures
      if (isTransfer) {
        forms.push({ file: "transfer_of_assets_nfs_acat.pdf", label: "Fidelity Transfer of Assets" });
        flags.push("⚠️  Delivering firm account number — NEEDS MANUAL INPUT (from old statement)");
        flags.push("⚠️  Delivering firm address/DTC — NEEDS MANUAL INPUT");
        flags.push("⚠️  Attach all pages of most recent account statement");
      }

      // WealthPort forms for new accounts
      forms.push({ file: "wealthport_custodian_wrap_client_agreement_and_application.pdf", label: "WealthPort Custodian Wrap Client Agreement & Application" });
      forms.push({ file: "wealthport_custodian_wrap_exhibit.pdf", label: "WealthPort Custodian Exhibit" });

      // Disclosures for new clients
      if (isNewClient) {
        forms.push({ file: "CIRFormCRS.pdf", label: "CIR Form CRS" });
        forms.push({ file: "CIRAFormCRS.pdf", label: "CIRA Form CRS" });
        forms.push({ file: "form_adv_delivery_receipt_discl.pdf", label: "Form ADV Delivery Receipt" });
      }
    }

    if (isMaintenance) {
      if (isRothConversion) {
        forms.push({ file: "RothConversion_Feb2026.pdf", label: "Fidelity Premiere Select Roth IRA Conversion" });
        forms.push({ file: "CIRFormCRS.pdf", label: "CIR Form CRS" });
        forms.push({ file: "CIRAFormCRS.pdf", label: "CIRA Form CRS" });
        forms.push({ file: "form_adv_delivery_receipt_discl.pdf", label: "Form ADV Delivery Receipt" });
        flags.push("⚠️  Both IRA account numbers required — confirm in task note");
      }

      if (isBIA) {
        forms.push({ file: "IRA Bene.pdf", label: "Fidelity IRA/HSA Beneficiary Designation" });
        forms.push({ file: "CIRFormCRS.pdf", label: "CIR Form CRS" });
        forms.push({ file: "CIRAFormCRS.pdf", label: "CIRA Form CRS" });
        forms.push({ file: "form_adv_delivery_receipt_discl.pdf", label: "Form ADV Delivery Receipt" });
      }
    }
  }

  // Beneficiary flags
  if (parsed.beneficiaries && parsed.beneficiaries.toLowerCase().includes("unknown")) {
    flags.push("⚠️  Minor SSN unknown — NEEDS MANUAL INPUT before submission");
  }
  if (!parsed.beneficiaries && isIRA) {
    flags.push("⚠️  No beneficiary info in task note — confirm bene details before submitting");
  }

  // Bank linking without voided check
  if (bankLinking && !parsed.hasVoidedCheck) {
    flags.push("⚠️  Bank linking requested but no voided check on file — obtain before submitting SPI");
  }

  return { forms, flags };
}

// ─── PDF FILLER ──────────────────────────────────────────────────────────────
async function fillPDF(formFile, fieldValues, outputPath) {
  // Uses python3 + pymupdf (fitz) to fill PDF fields
  // Build a python script inline and execute it
  const mapping = JSON.stringify(fieldValues);
  const script = `
import fitz, json, sys

doc = fitz.open(${JSON.stringify(join(FORMS_DIR, formFile))})
mapping = json.loads(${JSON.stringify(mapping)})
unfilled = []

for page in doc:
    for widget in page.widgets():
        field = widget.field_name
        if field in mapping and mapping[field]:
            if widget.field_type_string == "CheckBox":
                widget.field_value = mapping[field] == "true" or mapping[field] == True
            else:
                widget.field_value = str(mapping[field])
            widget.update()
        elif field and field not in ["Item_Code","Form_ID","Form_ID_02","Form_ID_03",
            "Form_ID_04","Form_ID_05","Form_ID_06","Form_ID_07","Form_ID_08",
            "Print","Reset","Clear Form","TransactionID"]:
            unfilled.append(field)

doc.save(${JSON.stringify(outputPath)})
print(json.dumps({"unfilled": unfilled}))
`;

  const result = execSync(`python3 -c '${script.replace(/'/g, `'"'"'`)}'`, {
    encoding: "utf8",
    maxBuffer: 10 * 1024 * 1024,
  });

  const { unfilled } = JSON.parse(result.trim());
  return unfilled;
}

// ─── FIELD VALUE BUILDER ─────────────────────────────────────────────────────
function buildFieldValues(formFile, rtFields, parsed) {
  const { financial } = parsed;

  // Common fields shared across most forms
  const common = {
    // Redtail identity
    "1own.FName":    rtFields.firstName,
    "1own.MName":    rtFields.middleName,
    "1own.LName":    rtFields.lastName,
    "1own.FullName": rtFields.fullName,
    "1own.SSN":      rtFields.SSN,
    "1own.DOB":      rtFields.DOB,
    "1own.H.Addr1":  rtFields.address,
    "1own.H.Addr2":  rtFields.address2,
    "1own.H.City":   rtFields.city,
    "1own.H.State":  rtFields.state,
    "1own.H.Zip":    rtFields.zip,
    "1own.H.Country": rtFields.country,
    "1own.H.Email":  rtFields.email,
    "1own.H.Mobile": rtFields.phone,
    "1own.H.Phone":  rtFields.phone,
    // Firm defaults
    "1rep.BrCompany":  "Cambridge Investment Research Advisors",
    "FirmName":        "Cambridge Investment Research Advisors",
    // Signature pages (print name)
    "SD_PrintAccountOwnerName": rtFields.fullName,
    "SD_PrintIRAOwnerName":     rtFields.fullName,
    // SPI forms
    "1own.1own_FName": rtFields.firstName,
    "1own.1own_MName": rtFields.middleName,
    "1own.1own_LName": rtFields.lastName,
    "1own.1own_FullName": rtFields.fullName,
    "1acc.1acc_AcctNum": parsed.accountNumbers[0] || "",
    // Transfer of assets
    "1own.FullName": rtFields.fullName,
    "1own.SSN":      rtFields.SSN,
    // Roth conversion
    "AI_First":  rtFields.firstName,
    "AI_MI":     rtFields.middleName,
    "AI_Last":   rtFields.lastName,
    "AI_PrimaryPhone": rtFields.phone,
    // Investment Exchange
    "Primary investor first MI lastauthorized signertrustentityminor print name": rtFields.fullName,
    "SSNTIN": rtFields.SSN,
  };

  // Suitability fields (Premiere Select IRA App page 12)
  const suitability = {
    "1own.Income.Range":       mapIncomeRange(financial.annualIncome),
    "1own.NetWorth.Range":     mapNetWorthRange(financial.netWorth),
    "1own.LiquidAssets.Range": mapNetWorthRange(financial.netInvestableAssets),
    "1own.FedTaxRange":        mapTaxBracket(financial.taxBracket),
    "1own.RiskTolerance":      mapRiskTolerance(financial.riskTolerance),
    "1own.PortTimeHoriz":      mapTimeHorizon(financial.timeHorizon),
    "1own.ExpAnnual.Range":    mapExpenseRange(financial.annualExpenses),
  };

  return { ...common, ...suitability };
}

// ─── RANGE MAPPERS ───────────────────────────────────────────────────────────
function mapIncomeRange(income) {
  if (!income) return "";
  const n = parseInt(income.replace(/[^0-9]/g, "")) || 0;
  if (n < 25000)  return "Under $25,000";
  if (n < 50000)  return "$25,000-$49,999";
  if (n < 100000) return "$50,000-$99,999";
  if (n < 200000) return "$100,000-$199,999";
  return "$200,000 or more";
}

function mapNetWorthRange(nw) {
  if (!nw) return "";
  const n = parseInt(nw.replace(/[^0-9]/g, "")) || 0;
  if (n < 50000)   return "Under $50,000";
  if (n < 100000)  return "$50,000-$99,999";
  if (n < 250000)  return "$100,000-$249,999";
  if (n < 500000)  return "$250,000-$499,999";
  if (n < 1000000) return "$500,000-$999,999";
  return "$1,000,000 or more";
}

function mapTaxBracket(bracket) {
  if (!bracket) return "";
  const n = parseInt(bracket.replace(/[^0-9]/g, "")) || 0;
  if (n <= 12) return "10% or 12%";
  if (n <= 22) return "22%";
  if (n <= 24) return "24%";
  if (n <= 32) return "32%";
  return "35% or 37%";
}

function mapRiskTolerance(risk) {
  if (!risk) return "";
  const r = risk.toLowerCase();
  if (r.includes("conservative")) return "Conservative";
  if (r.includes("mod") && r.includes("con")) return "Moderately Conservative";
  if (r.includes("moderate")) return "Moderate";
  if (r.includes("mod") && r.includes("agg")) return "Moderately Aggressive";
  if (r.includes("aggressive") || r.includes("high")) return "Aggressive";
  return risk;
}

function mapTimeHorizon(horizon) {
  if (!horizon) return "";
  const h = horizon.toLowerCase();
  if (h.includes("short") || h.includes("1") || h.includes("2")) return "Short (1-3 years)";
  if (h.includes("medium") || h.includes("3") || h.includes("5")) return "Medium (3-5 years)";
  if (h.includes("long") || h.includes("10")) return "Long (more than 10 years)";
  return horizon;
}

function mapExpenseRange(expenses) {
  if (!expenses) return "";
  const n = parseInt(expenses.replace(/[^0-9]/g, "")) || 0;
  if (n < 25000)  return "Under $25,000";
  if (n < 50000)  return "$25,000-$49,999";
  if (n < 75000)  return "$50,000-$74,999";
  if (n < 100000) return "$75,000-$99,999";
  return "$100,000 or more";
}

// ─── SHAREPOINT UPLOAD ───────────────────────────────────────────────────────
async function uploadToSharePoint(token, localPath, fileName, clientName) {
  // Get SharePoint site ID
  const hostname = "sdcapitale3.sharepoint.com";
  const sitePath = "/sites/SDC  2";
  const siteRes = await graphGet(token, `/sites/${hostname}:${sitePath}`);
  const siteId = siteRes.id;

  // Get drive ID
  const drivesRes = await graphGet(token, `/sites/${siteId}/drives`);
  const drive = drivesRes.value.find(d => d.name === "Documents") || drivesRes.value[0];
  const driveId = drive.id;

  // Build folder path with client name subfolder
  const folderPath = `${SHAREPOINT_FOLDER}/Completed Packets/${clientName} - ${new Date().toISOString().slice(0,10)}`;
  const uploadPath = `/drives/${driveId}/root:${folderPath}/${fileName}:/content`;

  const fileBuffer = readFileSync(localPath);
  const result = await graphPut(token, uploadPath, fileBuffer);
  return result.webUrl;
}

// ─── PLANNER COMMENT ─────────────────────────────────────────────────────────
async function commentOnTask(token, taskId, message) {
  // Get task details etag first
  const details = await graphGet(token, `/planner/tasks/${taskId}/details`);
  const etag = details["@odata.etag"];

  await fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      "If-Match": etag,
    },
    body: JSON.stringify({
      description: `${details.description || ""}\n\n--- AUTOMATION LOG ---\n${message}`,
    }),
  });
}

// ─── MAIN ────────────────────────────────────────────────────────────────────
async function main() {
  log("SDC Paperwork Automation — starting");

  // Load processed task log to avoid reprocessing
  const processed = existsSync(PROCESSED_LOG)
    ? JSON.parse(readFileSync(PROCESSED_LOG, "utf8"))
    : {};

  const token = await getToken();
  log("Microsoft Graph token acquired");

  // Pull all tasks from Paperwork planner
  const tasksRes = await graphGet(token, `/planner/plans/${PLAN_ID}/tasks`);
  const tasks = tasksRes.value || [];
  log(`Found ${tasks.length} tasks in Paperwork planner`);

  // Filter to unprocessed, incomplete new account tasks
  const toProcess = tasks.filter(t =>
    !processed[t.id] &&
    t.percentComplete < 100 &&
    t.title?.toLowerCase().includes("new account") ||
    t.title?.toLowerCase().includes("maintenance")
  );

  log(`${toProcess.length} tasks to process`);

  for (const task of toProcess) {
    log(`\nProcessing: ${task.title}`);

    try {
      // Get task note
      const details = await graphGet(token, `/planner/tasks/${task.id}/details`);
      const note = (details.description || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();

      if (!note) {
        log("  ⚠️  No task note — skipping");
        continue;
      }

      // Parse task note
      const parsed = parseTaskNote(note);
      log(`  Client: ${parsed.clientName}`);
      log(`  Account: ${parsed.accountType} — ${parsed.custodian}`);

      // Get required forms
      const { forms, flags } = getRequiredForms(parsed);
      if (!forms.length) {
        log("  ⚠️  No forms matched — may be AMF or unmapped account type, skipping");
        continue;
      }
      log(`  Forms: ${forms.map(f => f.label).join(", ")}`);

      // Pull client from Redtail
      let rtFields;
      try {
        const contact = await redtailSearch(parsed.clientName);
        const fullContact = await redtailGetContact(contact.id);
        rtFields = extractRedtailFields(fullContact);
        log(`  Redtail: found ${rtFields.fullName}`);
      } catch (err) {
        log(`  ⚠️  Redtail lookup failed: ${err.message}`);
        flags.push(`⚠️  Redtail lookup failed for "${parsed.clientName}" — all client fields NEED MANUAL INPUT`);
        rtFields = {
          firstName: "", middleName: "", lastName: "", fullName: parsed.clientName,
          SSN: "", DOB: "", address: "", address2: "", city: "", state: "",
          zip: "", country: "USA", phone: "", email: "", citizenship: "USA",
        };
      }

      // Fill each form
      const filledPaths = [];
      const allUnfilled = [];

      for (const form of forms) {
        const formPath = join(FORMS_DIR, form.file);
        if (!existsSync(formPath)) {
          log(`  ⚠️  Blank form not found: ${form.file}`);
          flags.push(`⚠️  Blank form missing: ${form.file} — PULL MANUALLY`);
          continue;
        }

        const outputPath = `/tmp/${task.id}_${form.file}`;
        const fieldValues = buildFieldValues(form.file, rtFields, parsed);

        try {
          const unfilled = await fillPDF(form.file, fieldValues, outputPath);
          filledPaths.push({ path: outputPath, label: form.label, file: form.file });
          if (unfilled.length) {
            allUnfilled.push(...unfilled.map(f => `${form.label}: ${f}`));
          }
          log(`  ✅ Filled: ${form.label} (${unfilled.length} unfilled fields)`);
        } catch (err) {
          log(`  ❌ Fill failed for ${form.label}: ${err.message}`);
          flags.push(`⚠️  Could not auto-fill ${form.label} — PULL MANUALLY`);
        }
      }

      // Build checklist
      const checklist = buildChecklist(parsed, forms, flags, allUnfilled, rtFields);
      const checklistPath = `/tmp/${task.id}_CHECKLIST.txt`;
      writeFileSync(checklistPath, checklist);
      log(`  Checklist written`);

      // Upload to SharePoint
      const uploadedUrls = [];
      for (const filled of filledPaths) {
        try {
          const url = await uploadToSharePoint(token, filled.path, filled.file, parsed.clientName);
          uploadedUrls.push({ label: filled.label, url });
          log(`  ✅ Uploaded: ${filled.label}`);
        } catch (err) {
          log(`  ❌ SharePoint upload failed for ${filled.label}: ${err.message}`);
        }
      }

      // Upload checklist
      try {
        const checklistUrl = await uploadToSharePoint(
          token, checklistPath, `CHECKLIST_${parsed.clientName}.txt`, parsed.clientName
        );
        uploadedUrls.push({ label: "Action Checklist", url: checklistUrl });
      } catch (err) {
        log(`  ❌ Checklist upload failed: ${err.message}`);
      }

      // Comment on Planner task
      const summary = buildSummary(parsed, forms, flags, uploadedUrls);
      await commentOnTask(token, task.id, summary);
      log(`  ✅ Planner task updated`);

      // Mark as processed
      processed[task.id] = {
        processedAt: new Date().toISOString(),
        clientName: parsed.clientName,
        formsCount: forms.length,
      };

    } catch (err) {
      log(`  ❌ Error processing task ${task.id}: ${err.message}`);
    }
  }

  // Save processed log
  writeFileSync(PROCESSED_LOG, JSON.stringify(processed, null, 2));
  log("\nDone.");
}

// ─── CHECKLIST BUILDER ───────────────────────────────────────────────────────
function buildChecklist(parsed, forms, flags, unfilled, rtFields) {
  const lines = [];
  const date = new Date().toLocaleDateString("en-US");

  lines.push(`SDC CAPITAL ADVISORS — PAPERWORK CHECKLIST`);
  lines.push(`Generated: ${date}`);
  lines.push(`Client: ${parsed.clientName}`);
  lines.push(`Account: ${parsed.accountType} — ${parsed.custodian}`);
  lines.push(`Signature Method: ${parsed.signatureMethod}`);
  lines.push("");
  lines.push("─".repeat(50));
  lines.push("FORMS IN THIS PACKET:");
  lines.push("─".repeat(50));
  forms.forEach((f, i) => lines.push(`  ${i + 1}. ${f.label}`));
  lines.push("");

  if (flags.length) {
    lines.push("─".repeat(50));
    lines.push("ACTION REQUIRED BEFORE SENDING:");
    lines.push("─".repeat(50));
    flags.forEach(f => lines.push(`  ${f}`));
    lines.push("");
  }

  if (unfilled.length) {
    lines.push("─".repeat(50));
    lines.push("UNFILLED PDF FIELDS (review before sending):");
    lines.push("─".repeat(50));
    unfilled.slice(0, 30).forEach(f => lines.push(`  • ${f}`));
    if (unfilled.length > 30) lines.push(`  ... and ${unfilled.length - 30} more`);
    lines.push("");
  }

  lines.push("─".repeat(50));
  lines.push("AUTO-FILLED FROM REDTAIL:");
  lines.push("─".repeat(50));
  const filled = Object.entries(rtFields).filter(([, v]) => v);
  filled.forEach(([k, v]) => lines.push(`  ✅ ${k}: ${k === "SSN" ? "***-**-" + String(v).slice(-4) : v}`));
  lines.push("");
  lines.push("─".repeat(50));
  lines.push("REVIEW NOTES:");
  lines.push("─".repeat(50));
  lines.push("  □ Review all auto-filled fields for accuracy");
  lines.push("  □ Complete all ACTION REQUIRED items above");
  lines.push("  □ Send via " + (parsed.signatureMethod || "DocuSign"));
  if (parsed.signatureMethod?.toLowerCase().includes("mail")) {
    lines.push("  □ Print, tag, and mail to client");
  }

  return lines.join("\n");
}

function buildSummary(parsed, forms, flags, uploadedUrls) {
  const lines = [];
  lines.push(`🤖 Automation complete — ${new Date().toLocaleDateString("en-US")}`);
  lines.push(`Forms filled: ${forms.length}`);
  if (flags.length) lines.push(`Action items: ${flags.length} — see checklist`);
  lines.push("");
  uploadedUrls.forEach(u => lines.push(`📄 ${u.label}: ${u.url}`));
  return lines.join("\n");
}

main().catch(err => {
  console.error("Fatal error:", err);
  process.exit(1);
});
