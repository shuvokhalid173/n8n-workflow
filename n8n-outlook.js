/**
 * n8n-outlook.js  —  Outlook Invoice Downloader (Phase 1)
 *
 * Key fixes in this version:
 *  - Webhook uses responseMode "onReceived" so it returns instantly (no timeout).
 *  - Execution poller searches ALL recent executions and matches by workflowId,
 *    with a longer initial wait so the record has time to appear.
 *  - Cleanup (deactivate + delete) happens AFTER polling finishes, not before.
 *  - Folder IDs read from .env if present, skipping the resolver round-trip.
 */

require('dotenv').config();
const axios = require('axios');

// ─── Config ────────────────────────────────────────────────────────────────────

const N8N_HOST = process.env.N8N_HOST || 'http://localhost:5678';
const N8N_API_KEY = process.env.N8N_API_KEY;
const OUTLOOK_CREDENTIAL_ID = process.env.OUTLOOK_CREDENTIAL_ID;
const LOCAL_DOWNLOAD_PATH = process.env.LOCAL_DOWNLOAD_PATH;
const N8N_CONTAINER_DOWNLOAD_PATH = process.env.N8N_CONTAINER_DOWNLOAD_PATH || '/home/node/downloads/';

const FOLDER_NAMES = ['FredFeuer', 'Waschbeckenmanufaktur'];
const EXECUTION_TIMEOUT_MS = 15 * 60 * 1000;  // 15 min max
const POLL_INTERVAL_MS = 4000;
const INITIAL_WAIT_MS = 8000;             // wait 8 s before first poll

// ─── Validation ────────────────────────────────────────────────────────────────

if (!N8N_API_KEY || !OUTLOOK_CREDENTIAL_ID || !LOCAL_DOWNLOAD_PATH) {
  console.error('❌  Missing required environment variables.');
  console.error('    Required: N8N_API_KEY, OUTLOOK_CREDENTIAL_ID, LOCAL_DOWNLOAD_PATH');
  process.exit(1);
}

const withSlash = (p) => (p.endsWith('/') ? p : `${p}/`);
const containerBase = withSlash(N8N_CONTAINER_DOWNLOAD_PATH);

// ─── n8n API client ────────────────────────────────────────────────────────────

const api = axios.create({
  baseURL: `${N8N_HOST}/api/v1`,
  headers: { 'X-N8N-API-KEY': N8N_API_KEY, Accept: 'application/json', 'Content-Type': 'application/json' },
});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ─── Folder ID resolver ────────────────────────────────────────────────────────

async function resolveFolderIds() {
  if (process.env.FREDFEUER_FOLDER_ID && process.env.WASCH_FOLDER_ID) {
    console.log('📁  Using folder IDs from .env (skipping auto-resolve).');
    return {
      FredFeuer: process.env.FREDFEUER_FOLDER_ID,
      Waschbeckenmanufaktur: process.env.WASCH_FOLDER_ID,
    };
  }

  console.log('🔍  Resolving Outlook folder IDs by display name…');
  const resolverPath = `resolve-folders-${Date.now()}`;
  const resolverSchema = {
    name: '__FolderResolver (auto-delete)',
    nodes: [
      {
        parameters: { httpMethod: 'POST', path: resolverPath, responseMode: 'lastNode', options: {} },
        name: 'Trigger', type: 'n8n-nodes-base.webhook', typeVersion: 1, position: [250, 300],
      },
      {
        parameters: {
          method: 'GET',
          url: 'https://graph.microsoft.com/v1.0/me/mailFolders?$top=50',
          authentication: 'predefinedCredentialType',
          nodeCredentialType: 'microsoftOutlookOAuth2Api',
          options: {},
        },
        name: 'Get Mail Folders', type: 'n8n-nodes-base.httpRequest', typeVersion: 4, position: [500, 300],
        credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
      },
    ],
    connections: { Trigger: { main: [[{ node: 'Get Mail Folders', type: 'main', index: 0 }]] } },
    settings: {},
  };

  let resolverId = null;
  try {
    const cr = await api.post('/workflows', resolverSchema);
    resolverId = cr.data.id;
    await api.post(`/workflows/${resolverId}/activate`, {});
    await sleep(2000);

    const tr = await axios.post(`${N8N_HOST}/webhook/${resolverPath}`, {}, { timeout: 15000 });
    const folders = tr.data?.value || [];
    const map = {};
    for (const name of FOLDER_NAMES) {
      const found = folders.find((f) => f.displayName.toLowerCase() === name.toLowerCase());
      if (!found) throw new Error(`Folder "${name}" not found. Available: ${folders.map((f) => f.displayName).join(', ')}`);
      map[name] = found.id;
      console.log(`   ✅  "${name}" → ${found.id}`);
    }
    return map;
  } finally {
    if (resolverId) await safeCleanup(resolverId, false);
  }
}

// ─── Workflow schema ───────────────────────────────────────────────────────────

function getWorkflowSchema(webhookPath, folderIds) {
  const makeSaveCode = (company) => `
const fs   = require('fs');
const path = require('path');
const BASE    = ${JSON.stringify(containerBase)};
const COMPANY = ${JSON.stringify(company)};
const outputItems = [];

for (const item of $input.all()) {
  const email = item.json;
  const rawDate = email.receivedDateTime || email.lastModifiedDateTime || new Date().toISOString();
  const d     = new Date(rawDate);
  const year  = d.getFullYear().toString();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day   = String(d.getDate()).padStart(2, '0');
  const dirPath = path.join(BASE, 'Invoices', COMPANY, year, month, day);
  fs.mkdirSync(dirPath, { recursive: true });

  const binaryKeys = Object.keys(item.binary || {});
  const pdfKeys = binaryKeys.filter((k) => {
    const b = item.binary[k];
    return b.mimeType === 'application/pdf' || (b.fileName && b.fileName.toLowerCase().endsWith('.pdf'));
  });

  if (pdfKeys.length === 0) {
    outputItems.push({ json: { skipped: true, subject: email.subject, company: COMPANY, reason: 'No PDF attachment' } });
    continue;
  }
  for (const key of pdfKeys) {
    const b        = item.binary[key];
    const fileName = b.fileName || ('invoice_' + Date.now() + '.pdf');
    const filePath = path.join(dirPath, fileName);
    outputItems.push({
      json:   { ...email, _savePath: filePath, _binaryKey: key, company: COMPANY },
      binary: { attachment_0: b },
    });
  }
}
return outputItems;
`;

  return {
    name: 'Automated Outlook Invoice Downloader',
    nodes: [
      {
        parameters: { httpMethod: 'POST', path: webhookPath, responseMode: 'onReceived', options: {} },
        name: 'Webhook Trigger', type: 'n8n-nodes-base.webhook', typeVersion: 1, position: [250, 300],
      },
      {
        parameters: {
          operation: 'getAll', folderId: folderIds['FredFeuer'],
          returnAll: true, downloadAttachments: true,
          filters: { hasAttachments: true }, options: {},
        },
        name: 'Outlook (FredFeuer)', type: 'n8n-nodes-base.microsoftOutlook', typeVersion: 2, position: [500, 160],
        credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
      },
      {
        parameters: {
          operation: 'getAll', folderId: folderIds['Waschbeckenmanufaktur'],
          returnAll: true, downloadAttachments: true,
          filters: { hasAttachments: true }, options: {},
        },
        name: 'Outlook (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.microsoftOutlook', typeVersion: 2, position: [500, 460],
        credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
      },
      {
        parameters: { mode: 'runOnceForAllItems', jsCode: makeSaveCode('FredFeuer'), options: {} },
        name: 'Prepare Path (FredFeuer)', type: 'n8n-nodes-base.code', typeVersion: 2, position: [750, 160],
      },
      {
        parameters: { mode: 'runOnceForAllItems', jsCode: makeSaveCode('Waschbeckenmanufaktur'), options: {} },
        name: 'Prepare Path (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.code', typeVersion: 2, position: [750, 460],
      },
      {
        parameters: { fileName: '={{ $json._savePath }}', dataPropertyName: 'attachment_0', options: {} },
        name: 'Save Invoice (FredFeuer)', type: 'n8n-nodes-base.writeBinaryFile', typeVersion: 1, position: [1000, 160],
      },
      {
        parameters: { fileName: '={{ $json._savePath }}', dataPropertyName: 'attachment_0', options: {} },
        name: 'Save Invoice (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.writeBinaryFile', typeVersion: 1, position: [1000, 460],
      },
    ],
    connections: {
      'Webhook Trigger': {
        main: [[
          { node: 'Outlook (FredFeuer)', type: 'main', index: 0 },
          { node: 'Outlook (Waschbeckenmanufaktur)', type: 'main', index: 0 },
        ]]
      },
      'Outlook (FredFeuer)': { main: [[{ node: 'Prepare Path (FredFeuer)', type: 'main', index: 0 }]] },
      'Outlook (Waschbeckenmanufaktur)': { main: [[{ node: 'Prepare Path (Waschbeckenmanufaktur)', type: 'main', index: 0 }]] },
      'Prepare Path (FredFeuer)': { main: [[{ node: 'Save Invoice (FredFeuer)', type: 'main', index: 0 }]] },
      'Prepare Path (Waschbeckenmanufaktur)': { main: [[{ node: 'Save Invoice (Waschbeckenmanufaktur)', type: 'main', index: 0 }]] },
    },
    settings: {},
  };
}

// ─── Execution poller ──────────────────────────────────────────────────────────

async function waitForExecution(workflowId) {
  const deadline = Date.now() + EXECUTION_TIMEOUT_MS;

  // Give n8n time to create the execution record before we start polling
  process.stdout.write(`⏳  Waiting ${INITIAL_WAIT_MS / 1000}s for execution to register`);
  await sleep(INITIAL_WAIT_MS);
  process.stdout.write('\n');

  let executionId = null;
  let attempts = 0;

  while (Date.now() < deadline) {
    attempts++;
    try {
      // Fetch last 20 executions and find the one matching our workflowId
      const res = await api.get('/executions?limit=20');
      const execs = res.data?.data || [];
      const match = execs.find((e) => e.workflowId === workflowId);

      if (match) {
        if (!executionId) {
          executionId = match.id;
          process.stdout.write(`🔄  Execution ${executionId} found — polling for completion`);
        }

        const status = match.status;
        const finished = match.finished;

        if (status === 'success' || (finished === true && status !== 'error')) {
          process.stdout.write('\n');
          // Fetch full execution data for the summary
          const full = await api.get(`/executions/${executionId}`);
          return { ok: true, exec: full.data };
        }
        if (status === 'error' || status === 'crashed') {
          process.stdout.write('\n');
          const full = await api.get(`/executions/${executionId}`);
          return { ok: false, exec: full.data };
        }

        process.stdout.write('.');
      } else {
        // Not found yet — print a dot every 3 attempts so the terminal doesn't look frozen
        if (attempts % 3 === 0) process.stdout.write('·');
      }
    } catch (_) {
      // Transient API error — keep polling
    }

    await sleep(POLL_INTERVAL_MS);
  }

  process.stdout.write('\n');
  throw new Error(`Execution did not finish within ${EXECUTION_TIMEOUT_MS / 60000} minutes.`);
}

// ─── Error details printer ────────────────────────────────────────────────────

function printErrorDetails(exec) {
  const runData = exec.data?.resultData?.runData || {};
  const topErr = exec.data?.resultData?.error;

  console.log('\n❌  Execution failed. Node-level details:');
  console.log('─────────────────────────────────────────');

  let foundError = false;
  for (const [nodeName, runs] of Object.entries(runData)) {
    for (const run of runs || []) {
      const itemCount = run?.data?.main?.[0]?.length ?? 0;
      if (run?.error) {
        foundError = true;
        console.log(`\n  ❌  Node: "${nodeName}"`);
        console.log(`      Message:  ${run.error.message}`);
        if (run.error.description) console.log(`      Detail:   ${run.error.description}`);
        if (run.error.context) console.log(`      Context:  ${JSON.stringify(run.error.context)}`);
        if (run.error.httpCode) console.log(`      HTTP code: ${run.error.httpCode}`);
      } else {
        console.log(`  ✅  Node: "${nodeName}" — ${itemCount} item(s) OK`);
      }
    }
  }

  if (topErr && !foundError) {
    console.log(`\n  ❌  Top-level error in node: "${topErr.node?.name || 'unknown'}"`);
    console.log(`      Message: ${topErr.message}`);
    if (topErr.description) console.log(`      Detail:  ${topErr.description}`);
  }

  console.log('\n─────────────────────────────────────────');
}

// ─── Summary printer ──────────────────────────────────────────────────────────

function printExecutionSummary(exec) {
  const nodeData = exec.data?.resultData?.runData || {};
  let saved = 0, skipped = 0, errored = 0;

  for (const [nodeName, runs] of Object.entries(nodeData)) {
    for (const run of runs || []) {
      if (run?.error) { errored++; continue; }
      if (!nodeName.startsWith('Prepare Path') && !nodeName.startsWith('Save Invoice')) continue;
      for (const item of run?.data?.main?.[0] || []) {
        const j = item?.json || {};
        if (j.skipped) skipped++;
        else if (j._savePath) saved++;
      }
    }
  }

  console.log('\n─────────────────────────────────────────');
  console.log('📊  Summary:');
  console.log(`    ✅  Saved:   ${saved} PDF(s)`);
  if (skipped) console.log(`    ⚠️   Skipped: ${skipped} (no PDF attachment)`);
  if (errored) console.log(`    ❌  Errors:  ${errored} node error(s) — check n8n UI → Executions`);
  console.log(`\n📂  Files are in: ${withSlash(LOCAL_DOWNLOAD_PATH)}Invoices/`);
  console.log('─────────────────────────────────────────\n');
}

// ─── Cleanup ──────────────────────────────────────────────────────────────────

async function safeCleanup(workflowId, log = true) {
  try { await api.post(`/workflows/${workflowId}/deactivate`, {}); } catch (_) { }
  try {
    await api.delete(`/workflows/${workflowId}`);
    if (log) console.log(`🗑   Cleaned up workflow ${workflowId}`);
  } catch (_) { }
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  let workflowId = null;

  try {
    // 1. Resolve folder IDs
    const folderIds = await resolveFolderIds();

    // 2. Register workflow
    const webhookPath = `outlook-invoices-${Date.now()}`;
    console.log('\n📝  Registering workflow in n8n…');
    const cr = await api.post('/workflows', getWorkflowSchema(webhookPath, folderIds));
    workflowId = cr.data.id;
    console.log(`✅  Workflow registered — ID: ${workflowId}`);

    // 3. Activate — give n8n a moment to fully register the webhook route
    await api.post(`/workflows/${workflowId}/activate`, {});
    console.log('⚡  Workflow activated.');
    await sleep(3000);   // 3 s is enough for webhook route registration

    // 4. Fire webhook — returns instantly (onReceived mode)
    const webhookUrl = `${N8N_HOST}/webhook/${webhookPath}`;
    console.log(`\n🚀  Triggering: ${webhookUrl}`);
    await axios.post(webhookUrl, {}, { timeout: 10_000 });
    console.log('✅  Webhook accepted. n8n is downloading your emails…');

    // 5. Poll until done  ← cleanup happens AFTER this
    const { ok, exec } = await waitForExecution(workflowId);

    if (ok) {
      printExecutionSummary(exec);
      // Only delete workflow on success — on error we keep it so the UI shows the trace
      if (workflowId) await safeCleanup(workflowId);
      workflowId = null; // prevent double-cleanup in finally
    } else {
      printErrorDetails(exec);
      process.exitCode = 1;
      // Deactivate but DO NOT delete — leave it in n8n UI for inspection
      try { await api.post(`/workflows/${workflowId}/deactivate`, {}); } catch (_) { }
      console.log(`⚠️   Workflow ${workflowId} left in n8n for inspection (not deleted).`);
      console.log(`    Open http://localhost:5678 → Executions to see the full trace.`);
      workflowId = null;
    }

  } catch (err) {
    console.error('\n❌  Error:');
    if (err.response) {
      console.error(`    HTTP ${err.response.status}`);
      console.error('    Body:', JSON.stringify(err.response.data, null, 2));
    } else {
      console.error('   ', err.message);
    }
    process.exitCode = 1;
  } finally {
    // Only runs if workflowId wasn't already cleaned up above
    if (workflowId) await safeCleanup(workflowId);
  }
}

main();