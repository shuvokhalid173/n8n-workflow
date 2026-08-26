require('dotenv').config();
const axios = require('axios');

// Configuration from environment variables
const N8N_HOST = process.env.N8N_HOST || 'http://localhost:5678';
const N8N_API_KEY = process.env.N8N_API_KEY;
const OUTLOOK_CREDENTIAL_ID = process.env.OUTLOOK_CREDENTIAL_ID;
const LOCAL_DOWNLOAD_PATH = process.env.LOCAL_DOWNLOAD_PATH;
const N8N_CONTAINER_DOWNLOAD_PATH = process.env.N8N_CONTAINER_DOWNLOAD_PATH || '/home/node/downloads/';

if (!N8N_API_KEY || !OUTLOOK_CREDENTIAL_ID || !LOCAL_DOWNLOAD_PATH) {
  console.error('❌ Error: Missing required environment variables. Please check your .env file.');
  console.error('   Required: N8N_API_KEY, OUTLOOK_CREDENTIAL_ID, LOCAL_DOWNLOAD_PATH');
  process.exit(1);
}

const withTrailingSlash = (path) => (path.endsWith('/') ? path : `${path}/`);
const localDownloadPath = withTrailingSlash(LOCAL_DOWNLOAD_PATH);
const containerDownloadPath = withTrailingSlash(N8N_CONTAINER_DOWNLOAD_PATH);

// Axios instance configured for n8n API
const n8nApi = axios.create({
  baseURL: `${N8N_HOST}/api/v1`,
  headers: {
    'X-N8N-API-KEY': N8N_API_KEY,
    'Accept': 'application/json',
    'Content-Type': 'application/json',
  },
});

const summarizeWorkflowResult = (data) => {
  if (!data || typeof data !== 'object') {
    return data;
  }

  // Handle array responses from parallel execution branches
  if (Array.isArray(data)) {
    return data.map((item) => ({
      company: item.folderName || item.from?.emailAddress?.name || 'Unknown',
      subject: item.subject,
      receivedDate: item.receivedDateTime || item.date,
      hasAttachments: item.hasAttachments,
      savedFile: item.fileName || item.binary?.attachment_0?.fileName || null,
      status: item.fileName ? '✅ Downloaded' : '⚠️ Check execution log',
    }));
  }

  // Handle single object response
  return {
    company: data.folderName || data.from?.emailAddress?.name || 'Unknown',
    subject: data.subject,
    receivedDate: data.receivedDateTime || data.date,
    hasAttachments: data.hasAttachments,
    savedFile: data.fileName || data.binary?.attachment_0?.fileName || null,
    status: data.fileName ? '✅ Downloaded' : '⚠️ Check execution log',
  };
};

// Construct the declarative JSON schema for the n8n workflow
const getWorkflowSchema = (webhookPath) => {
  return {
    name: 'Automated Outlook Invoice Downloader',
    nodes: [
      // ── Webhook Trigger ──────────────────────────────────────────────
      {
        parameters: {
          httpMethod: 'POST',
          path: webhookPath,
          responseMode: 'lastNode',
          options: {}
        },
        name: 'Webhook Trigger',
        type: 'n8n-nodes-base.webhook',
        typeVersion: 1,
        position: [250, 300]
      },

      // ── Branch 1: Outlook (FredFeuer) ────────────────────────────────
      {
        parameters: {
          operation: 'getAll',
          folderId: 'FredFeuer',
          returnAll: true,
          filters: {
            hasAttachments: true
          },
          options: {
            downloadAttachments: true
          }
        },
        name: 'Outlook (FredFeuer)',
        type: 'n8n-nodes-base.microsoftOutlook',
        typeVersion: 1,
        position: [500, 180],
        credentials: {
          microsoftOutlookOAuth2Api: {
            id: OUTLOOK_CREDENTIAL_ID
          }
        }
      },

      // ── Branch 2: Outlook (Waschbeckenmanufaktur) ────────────────────
      {
        parameters: {
          operation: 'getAll',
          folderId: 'Waschbeckenmanufaktur',
          returnAll: true,
          filters: {
            hasAttachments: true
          },
          options: {
            downloadAttachments: true
          }
        },
        name: 'Outlook (Waschbeckenmanufaktur)',
        type: 'n8n-nodes-base.microsoftOutlook',
        typeVersion: 1,
        position: [500, 460],
        credentials: {
          microsoftOutlookOAuth2Api: {
            id: OUTLOOK_CREDENTIAL_ID
          }
        }
      },

      // ── Write Binary File – FredFeuer ────────────────────────────────
      {
        parameters: {
          fileName: `=${containerDownloadPath}Invoices/FredFeuer/{{ $now.format('YYYY/MM/DD') }}/{{ $binary.attachment_0.fileName }}`,
          dataPropertyName: 'attachment_0',
          options: {}
        },
        name: 'Save Invoice (FredFeuer)',
        type: 'n8n-nodes-base.writeBinaryFile',
        typeVersion: 1,
        position: [750, 180]
      },

      // ── Write Binary File – Waschbeckenmanufaktur ────────────────────
      {
        parameters: {
          fileName: `=${containerDownloadPath}Invoices/Waschbeckenmanufaktur/{{ $now.format('YYYY/MM/DD') }}/{{ $binary.attachment_0.fileName }}`,
          dataPropertyName: 'attachment_0',
          options: {}
        },
        name: 'Save Invoice (Waschbeckenmanufaktur)',
        type: 'n8n-nodes-base.writeBinaryFile',
        typeVersion: 1,
        position: [750, 460]
      }
    ],
    connections: {
      // Webhook fans out to both Outlook branches in parallel
      'Webhook Trigger': {
        main: [
          [
            {
              node: 'Outlook (FredFeuer)',
              type: 'main',
              index: 0
            },
            {
              node: 'Outlook (Waschbeckenmanufaktur)',
              type: 'main',
              index: 0
            }
          ]
        ]
      },
      // Each Outlook branch feeds into its dedicated Write Binary File node
      'Outlook (FredFeuer)': {
        main: [
          [
            {
              node: 'Save Invoice (FredFeuer)',
              type: 'main',
              index: 0
            }
          ]
        ]
      },
      'Outlook (Waschbeckenmanufaktur)': {
        main: [
          [
            {
              node: 'Save Invoice (Waschbeckenmanufaktur)',
              type: 'main',
              index: 0
            }
          ]
        ]
      }
    },
    settings: {}
  };
};

async function main() {
  let workflowId = null;

  try {
    console.log('🚀 Starting n8n Outlook invoice automation...');

    // 1. Create/Register the workflow in n8n
    console.log('📝 Registering workflow in local n8n instance...');
    const uniqueWebhookPath = `trigger-outlook-invoices-${Date.now()}`;
    const workflowSchema = getWorkflowSchema(uniqueWebhookPath);
    const createResponse = await n8nApi.post('/workflows', workflowSchema);
    workflowId = createResponse.data.id;
    console.log(`✅ Workflow registered successfully with ID: ${workflowId}`);

    console.log('⚡ Activating workflow...');
    await n8nApi.post(`/workflows/${workflowId}/activate`, {});
    console.log('✅ Workflow activated successfully.');

    // Wait a brief moment to ensure the active webhook is fully registered internally by n8n
    await new Promise(resolve => setTimeout(resolve, 1000));

    // 2. Trigger the workflow instantly via the webhook
    const webhookUrl = `${N8N_HOST}/webhook/${uniqueWebhookPath}`;
    console.log(`⚡ Triggering workflow via webhook at ${webhookUrl}...`);

    const triggerResponse = await axios.post(webhookUrl, {});
    console.log('✅ Workflow triggered successfully!');

    const result = summarizeWorkflowResult(triggerResponse.data);
    if (Array.isArray(result)) {
      console.log(`\n📊 Invoice Download Summary (${result.length} items):`);
      result.forEach((entry, idx) => {
        console.log(`  [${idx + 1}] ${entry.company} — ${entry.savedFile || 'no file'} ${entry.status}`);
      });
    } else {
      console.log('Result:', result);
    }

    console.log(`\n🎉 Process complete. Check your local directory: ${localDownloadPath} for downloaded invoices.`);
    console.log(`n8n wrote inside the container at: ${containerDownloadPath}`);

  } catch (error) {
    console.error('\n❌ An error occurred during execution:');

    if (error.response) {
      // The request was made and the server responded with a status code outside of the 2xx range
      console.error(`Status Code: ${error.response.status}`);
      console.error('Response Data:', JSON.stringify(error.response.data, null, 2));
    } else if (error.request) {
      // The request was made but no response was received
      console.error('No response received from the server. Is n8n running on port 5678?');
      console.error(error.message);
    } else {
      // Something happened in setting up the request
      console.error('Error Details:', error.message);
    }
  }
}

main();
