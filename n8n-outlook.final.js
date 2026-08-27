/**
 * n8n-outlook.js  —  Outlook Invoice Downloader (Folder-Move Architecture)
 */

require('dotenv').config();
const axios = require('axios');
const fs = require('fs');
const path = require('path');

// ─── Config ────────────────────────────────────────────────────────────────────

const N8N_HOST = process.env.N8N_HOST || 'http://localhost:5678';
const N8N_API_KEY = process.env.N8N_API_KEY;
const OUTLOOK_CREDENTIAL_ID = process.env.OUTLOOK_CREDENTIAL_ID;
const LOCAL_DOWNLOAD_PATH = process.env.LOCAL_DOWNLOAD_PATH;

const FOLDER_NAMES = ['FredFeuer', 'Waschbeckenmanufaktur'];
const SUBFOLDER_NAME = 'Processed';
const EXECUTION_TIMEOUT_MS = 15 * 60 * 1000;
const POLL_INTERVAL_MS = 4000;
const INITIAL_WAIT_MS = 8000;

// ─── Validation ────────────────────────────────────────────────────────────────

if (!N8N_API_KEY || !OUTLOOK_CREDENTIAL_ID || !LOCAL_DOWNLOAD_PATH) {
    console.error('❌  Missing required environment variables.');
    console.error('    Required: N8N_API_KEY, OUTLOOK_CREDENTIAL_ID, LOCAL_DOWNLOAD_PATH');
    process.exit(1);
}

// ─── n8n API client ────────────────────────────────────────────────────────────

const api = axios.create({
    baseURL: `${N8N_HOST}/api/v1`,
    headers: { 'X-N8N-API-KEY': N8N_API_KEY, Accept: 'application/json', 'Content-Type': 'application/json' },
});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ─── Folder ID resolver ────────────────────────────────────────────────────────

async function resolveFolderIds() {
    console.log('🔍  Resolving Outlook folder IDs & "Processed" subfolders…');
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
                    url: 'https://graph.microsoft.com/v1.0/me/mailFolders?$top=50&$expand=childFolders',
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

            const childFolders = found.childFolders || [];
            const processedChild = childFolders.find((c) => c.displayName.toLowerCase() === SUBFOLDER_NAME.toLowerCase());

            if (!processedChild) {
                throw new Error(`Subfolder "${SUBFOLDER_NAME}" not found under "${name}". Please create "${name}/${SUBFOLDER_NAME}" in Outlook.`);
            }

            map[name] = {
                id: found.id,
                processedId: processedChild.id,
            };
            console.log(`   ✅  "${name}" → ${found.id}`);
            console.log(`       ↳ Subfolder "${SUBFOLDER_NAME}" → ${processedChild.id}`);
        }
        return map;
    } finally {
        if (resolverId) await safeCleanup(resolverId, false);
    }
}

// ─── Workflow schema ───────────────────────────────────────────────────────────

function getWorkflowSchema(webhookPath, folderIds) {
    const makeSplitCode = () => `
const items = $input.all();
const out = [];
for (const item of items) {
  const messages = item.json?.value || [];
  for (const msg of messages) {
    if (msg && msg.id && msg.hasAttachments) {
      out.push({ json: msg });
    }
  }
}
return out;
`;

    const makeSaveCode = (company, moveNodeName) => `
const COMPANY = ${JSON.stringify(company)};
const outputItems = [];
const items = $input.all();

for (let i = 0; i < items.length; i++) {
  const item = items[i];
  const attachments = item.json?.value || [];
  
  let rawDate = new Date().toISOString();
  let msgId = Date.now().toString().slice(-6);
  try {
    const movedItems = $("${moveNodeName}").all();
    if (movedItems[i] && movedItems[i].json) {
      if (movedItems[i].json.receivedDateTime) rawDate = movedItems[i].json.receivedDateTime;
      if (movedItems[i].json.id) msgId = movedItems[i].json.id.slice(-6);
    }
  } catch (_) {}

  const d     = new Date(rawDate);
  const year  = d.getFullYear().toString();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day   = String(d.getDate()).padStart(2, '0');
  
  const relDir = 'Invoices/' + COMPANY + '/' + year + '/' + month + '/' + day;

  for (const att of attachments) {
    const name = att.name || '';
    const mime = att.contentType || '';
    const isPdf = mime === 'application/pdf' || name.toLowerCase().endsWith('.pdf');
    
    if (isPdf && att.contentBytes) {
      const fileName = name ? (msgId + '_' + name) : ('invoice_' + msgId + '_' + Date.now() + '.pdf');
      outputItems.push({
        json: {
          relDir: relDir,
          fileName: fileName,
          company: COMPANY,
          contentBytes: att.contentBytes,
        }
      });
    }
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
            // ─── FredFeuer Pipeline ───
            {
                parameters: {
                    method: 'GET',
                    url: `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderIds['FredFeuer'].id)}/messages?$top=50&$filter=hasAttachments eq true&$select=id,receivedDateTime,subject,hasAttachments`,
                    authentication: 'predefinedCredentialType',
                    nodeCredentialType: 'microsoftOutlookOAuth2Api',
                    options: {},
                },
                name: 'Get Messages (FredFeuer)', type: 'n8n-nodes-base.httpRequest', typeVersion: 4, position: [450, 160],
                credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
            },
            {
                parameters: { mode: 'runOnceForAllItems', jsCode: makeSplitCode(), options: {} },
                name: 'Split Messages (FredFeuer)', type: 'n8n-nodes-base.code', typeVersion: 2, position: [650, 160],
            },
            {
                parameters: {
                    method: 'POST',
                    url: '=https://graph.microsoft.com/v1.0/me/messages/{{ $json.id }}/move',
                    authentication: 'predefinedCredentialType',
                    nodeCredentialType: 'microsoftOutlookOAuth2Api',
                    sendBody: true,
                    specifyBody: 'json',
                    jsonBody: `{"destinationId": "${folderIds['FredFeuer'].processedId}"}`,
                    options: {},
                },
                name: 'Move Message (FredFeuer)', type: 'n8n-nodes-base.httpRequest', typeVersion: 4, position: [850, 160],
                credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
            },
            {
                parameters: {
                    method: 'GET',
                    url: '=https://graph.microsoft.com/v1.0/me/messages/{{ $json.id }}/attachments',
                    authentication: 'predefinedCredentialType',
                    nodeCredentialType: 'microsoftOutlookOAuth2Api',
                    options: {},
                },
                name: 'Get Attachments (FredFeuer)', type: 'n8n-nodes-base.httpRequest', typeVersion: 4, position: [1050, 160],
                credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
            },
            {
                parameters: { mode: 'runOnceForAllItems', jsCode: makeSaveCode('FredFeuer', 'Move Message (FredFeuer)'), options: {} },
                name: 'Prepare Path (FredFeuer)', type: 'n8n-nodes-base.code', typeVersion: 2, position: [1250, 160],
            },

            // ─── Waschbeckenmanufaktur Pipeline ───
            {
                parameters: {
                    method: 'GET',
                    url: `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderIds['Waschbeckenmanufaktur'].id)}/messages?$top=50&$filter=hasAttachments eq true&$select=id,receivedDateTime,subject,hasAttachments`,
                    authentication: 'predefinedCredentialType',
                    nodeCredentialType: 'microsoftOutlookOAuth2Api',
                    options: {},
                },
                name: 'Get Messages (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.httpRequest', typeVersion: 4, position: [450, 460],
                credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
            },
            {
                parameters: { mode: 'runOnceForAllItems', jsCode: makeSplitCode(), options: {} },
                name: 'Split Messages (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.code', typeVersion: 2, position: [650, 460],
            },
            {
                parameters: {
                    method: 'POST',
                    url: '=https://graph.microsoft.com/v1.0/me/messages/{{ $json.id }}/move',
                    authentication: 'predefinedCredentialType',
                    nodeCredentialType: 'microsoftOutlookOAuth2Api',
                    sendBody: true,
                    specifyBody: 'json',
                    jsonBody: `{"destinationId": "${folderIds['Waschbeckenmanufaktur'].processedId}"}`,
                    options: {},
                },
                name: 'Move Message (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.httpRequest', typeVersion: 4, position: [850, 460],
                credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
            },
            {
                parameters: {
                    method: 'GET',
                    url: '=https://graph.microsoft.com/v1.0/me/messages/{{ $json.id }}/attachments',
                    authentication: 'predefinedCredentialType',
                    nodeCredentialType: 'microsoftOutlookOAuth2Api',
                    options: {},
                },
                name: 'Get Attachments (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.httpRequest', typeVersion: 4, position: [1050, 460],
                credentials: { microsoftOutlookOAuth2Api: { id: OUTLOOK_CREDENTIAL_ID } },
            },
            {
                parameters: { mode: 'runOnceForAllItems', jsCode: makeSaveCode('Waschbeckenmanufaktur', 'Move Message (Waschbeckenmanufaktur)'), options: {} },
                name: 'Prepare Path (Waschbeckenmanufaktur)', type: 'n8n-nodes-base.code', typeVersion: 2, position: [1250, 460],
            },
        ],
        connections: {
            'Webhook Trigger': {
                main: [[
                    { node: 'Get Messages (FredFeuer)', type: 'main', index: 0 },
                    { node: 'Get Messages (Waschbeckenmanufaktur)', type: 'main', index: 0 },
                ]]
            },
            'Get Messages (FredFeuer)': { main: [[{ node: 'Split Messages (FredFeuer)', type: 'main', index: 0 }]] },
            'Split Messages (FredFeuer)': { main: [[{ node: 'Move Message (FredFeuer)', type: 'main', index: 0 }]] },
            'Move Message (FredFeuer)': { main: [[{ node: 'Get Attachments (FredFeuer)', type: 'main', index: 0 }]] },
            'Get Attachments (FredFeuer)': { main: [[{ node: 'Prepare Path (FredFeuer)', type: 'main', index: 0 }]] },

            'Get Messages (Waschbeckenmanufaktur)': { main: [[{ node: 'Split Messages (Waschbeckenmanufaktur)', type: 'main', index: 0 }]] },
            'Split Messages (Waschbeckenmanufaktur)': { main: [[{ node: 'Move Message (Waschbeckenmanufaktur)', type: 'main', index: 0 }]] },
            'Move Message (Waschbeckenmanufaktur)': { main: [[{ node: 'Get Attachments (Waschbeckenmanufaktur)', type: 'main', index: 0 }]] },
            'Get Attachments (Waschbeckenmanufaktur)': { main: [[{ node: 'Prepare Path (Waschbeckenmanufaktur)', type: 'main', index: 0 }]] },
        },
        settings: {},
    };
}

// ─── Execution poller ──────────────────────────────────────────────────────────

async function waitForExecution(workflowId) {
    const deadline = Date.now() + EXECUTION_TIMEOUT_MS;

    process.stdout.write(`⏳  Waiting ${INITIAL_WAIT_MS / 1000}s for execution to register`);
    await sleep(INITIAL_WAIT_MS);
    process.stdout.write('\n');

    let executionId = null;
    let attempts = 0;

    while (Date.now() < deadline) {
        attempts++;
        try {
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
                    const full = await api.get(`/executions/${executionId}?includeData=true`);
                    return { ok: true, exec: full.data };
                }
                if (status === 'error' || status === 'crashed') {
                    process.stdout.write('\n');
                    const full = await api.get(`/executions/${executionId}?includeData=true`);
                    return { ok: false, exec: full.data };
                }

                process.stdout.write('.');
            } else {
                if (attempts % 3 === 0) process.stdout.write('·');
            }
        } catch (_) { }
        await sleep(POLL_INTERVAL_MS);
    }

    process.stdout.write('\n');
    throw new Error(`Execution did not finish within ${EXECUTION_TIMEOUT_MS / 60000} minutes.`);
}

// ─── Local File Saver ──────────────────────────────────────────────────────────

function saveInvoicesFromExecution(exec) {
    const runData = exec.resultData?.runData || exec.data?.resultData?.runData || {};
    let totalFetched = 0;
    let saved = 0;

    for (const [nodeName, runs] of Object.entries(runData)) {
        if (nodeName.startsWith('Split Messages')) {
            for (const run of runs || []) {
                totalFetched += run?.data?.main?.[0]?.length || 0;
            }
        }

        if (nodeName.startsWith('Prepare Path')) {
            for (const run of runs || []) {
                const items = run?.data?.main?.[0] || [];
                for (const item of items) {
                    const data = item.json;
                    if (data && data.contentBytes && data.relDir && data.fileName) {
                        const targetDir = path.join(LOCAL_DOWNLOAD_PATH, data.relDir);
                        fs.mkdirSync(targetDir, { recursive: true });

                        const targetFilePath = path.join(targetDir, data.fileName);
                        fs.writeFileSync(targetFilePath, Buffer.from(data.contentBytes, 'base64'));
                        saved++;
                    }
                }
            }
        }
    }

    const skipped = totalFetched - saved;

    console.log('\n─────────────────────────────────────────');
    console.log('📊  Summary:');
    console.log(`    📥  Fetched & Moved: ${totalFetched} email(s)`);
    console.log(`    ✅  Saved:           ${saved} PDF(s)`);
    if (skipped > 0) console.log(`    ⚠️   Skipped:         ${skipped} (no PDF attachment)`);
    console.log(`\n📂  Files saved locally to: ${path.resolve(LOCAL_DOWNLOAD_PATH, 'Invoices')}`);
    console.log('─────────────────────────────────────────\n');
}

// ─── Error details printer ────────────────────────────────────────────────────

function printErrorDetails(exec) {
    const resData = exec.resultData || exec.data?.resultData || {};
    const runData = resData.runData || {};
    const topErr = resData.error;

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
        const folderIds = await resolveFolderIds();
        const webhookPath = `outlook-invoices-${Date.now()}`;

        console.log('\n📝  Registering workflow in n8n…');
        const cr = await api.post('/workflows', getWorkflowSchema(webhookPath, folderIds));
        workflowId = cr.data.id;
        console.log(`✅  Workflow registered — ID: ${workflowId}`);

        await api.post(`/workflows/${workflowId}/activate`, {});
        console.log('⚡  Workflow activated.');
        await sleep(3000);

        const webhookUrl = `${N8N_HOST}/webhook/${webhookPath}`;
        console.log(`\n🚀  Triggering: ${webhookUrl}`);
        await axios.post(webhookUrl, {}, { timeout: 10_000 });
        console.log('✅  Webhook accepted. n8n is processing your emails…');

        const { ok, exec } = await waitForExecution(workflowId);

        if (ok) {
            saveInvoicesFromExecution(exec);
            if (workflowId) await safeCleanup(workflowId);
            workflowId = null;
        } else {
            printErrorDetails(exec);
            process.exitCode = 1;
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
        if (workflowId) await safeCleanup(workflowId);
    }
}

main();