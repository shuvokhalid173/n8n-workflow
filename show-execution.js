/**
 * show-execution.js
 * Fetches the full details of a specific execution and prints the error.
 * Usage: node show-execution.js <execution-id>
 * Example: node show-execution.js 18
 */
require('dotenv').config();
const axios = require('axios');

const N8N_HOST = process.env.N8N_HOST || 'http://localhost:5678';
const N8N_API_KEY = process.env.N8N_API_KEY;

const api = axios.create({
    baseURL: `${N8N_HOST}/api/v1`,
    headers: { 'X-N8N-API-KEY': N8N_API_KEY, Accept: 'application/json' },
});

async function main() {
    const execId = process.argv[2];
    if (!execId) {
        console.error('Usage: node show-execution.js <execution-id>');
        process.exit(1);
    }

    const res = await api.get(`/executions/${execId}`);
    const exec = res.data;

    console.log(`\n📋  Execution ${execId}`);
    console.log(`    Status:   ${exec.status}`);
    console.log(`    Started:  ${exec.startedAt}`);
    console.log(`    Finished: ${exec.stoppedAt}`);

    const runData = exec.data?.resultData?.runData || {};

    console.log('\n─── Node results ───────────────────────────────────────────────');
    for (const [nodeName, runs] of Object.entries(runData)) {
        for (const run of runs || []) {
            const itemCount = run?.data?.main?.[0]?.length ?? 0;
            if (run?.error) {
                console.log(`\n  ❌  ${nodeName}`);
                console.log(`      Error:   ${run.error.message}`);
                console.log(`      Details: ${JSON.stringify(run.error.description || run.error.context || '')}`);
                if (run.error.cause) {
                    console.log(`      Cause:   ${JSON.stringify(run.error.cause)}`);
                }
            } else {
                console.log(`  ✅  ${nodeName} — ${itemCount} item(s)`);
            }
        }
    }

    // Top-level error
    const topErr = exec.data?.resultData?.error;
    if (topErr) {
        console.log('\n─── Top-level error ────────────────────────────────────────────');
        console.log(`  Message: ${topErr.message}`);
        console.log(`  Node:    ${topErr.node?.name || 'unknown'}`);
        console.log(`  Details: ${JSON.stringify(topErr.description || topErr.context || '')}`);
    }

    console.log('\n────────────────────────────────────────────────────────────────\n');
}

main().catch(e => console.error('❌', e.response?.data || e.message));