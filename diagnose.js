/**
 * diagnose.js — run this in a SECOND terminal while n8n-outlook.js is running.
 * It shows you: all recent executions, their status, and any error messages.
 * Usage: node diagnose.js
 */
require('dotenv').config();
const axios = require('axios');

const N8N_HOST = process.env.N8N_HOST || 'http://localhost:5678';
const N8N_API_KEY = process.env.N8N_API_KEY;

if (!N8N_API_KEY) {
    console.error('❌  N8N_API_KEY not set in .env');
    process.exit(1);
}

const api = axios.create({
    baseURL: `${N8N_HOST}/api/v1`,
    headers: { 'X-N8N-API-KEY': N8N_API_KEY, Accept: 'application/json' },
});

async function main() {
    console.log(`\n🔍  Querying n8n at ${N8N_HOST}\n`);

    // 1. Show all executions (last 10)
    try {
        const res = await api.get('/executions?limit=10');
        const execs = res.data?.data || [];

        if (execs.length === 0) {
            console.log('⚠️  No executions found at all.\n');
            console.log('This means the webhook trigger fired but n8n never started the workflow.');
            console.log('Possible causes:');
            console.log('  • The workflow was cleaned up before the execution could be recorded.');
            console.log('  • The 1500ms sleep before triggering was not enough — n8n needs more time.');
            console.log('  • The webhook path did not match (race condition on activation).\n');
        } else {
            console.log(`Found ${execs.length} execution(s):\n`);
            for (const e of execs) {
                console.log(`  ID:         ${e.id}`);
                console.log(`  Workflow:   ${e.workflowId} — ${e.workflowData?.name || '(name not returned)'}`);
                console.log(`  Status:     ${e.status || e.finished}`);
                console.log(`  Started:    ${e.startedAt}`);
                console.log(`  Finished:   ${e.stoppedAt || 'still running'}`);
                if (e.data?.resultData?.error) {
                    console.log(`  ❌ Error:   ${e.data.resultData.error.message}`);
                }
                console.log('');
            }
        }
    } catch (err) {
        console.error('❌  Failed to list executions:', err.response?.data || err.message);
    }

    // 2. Show all active workflows
    try {
        const res = await api.get('/workflows?active=true');
        const wfs = res.data?.data || [];
        console.log(`\n📋  Active workflows right now: ${wfs.length}`);
        for (const w of wfs) {
            console.log(`  • [${w.id}] ${w.name}`);
        }
    } catch (err) {
        console.error('❌  Failed to list workflows:', err.response?.data || err.message);
    }

    // 3. Check n8n health
    try {
        const h = await axios.get(`${N8N_HOST}/healthz`);
        console.log(`\n💚  n8n health: ${h.data?.status || 'ok'}`);
    } catch {
        console.log('\n❌  n8n health check failed — is it still running?');
    }
}

main();