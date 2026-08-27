/**
 * cleanup-orphans.js
 * Run this ONCE to delete all the leftover "Automated Outlook Invoice Downloader"
 * workflows that piled up. Safe to run any time.
 * Usage: node cleanup-orphans.js
 */
require('dotenv').config();
const axios = require('axios');

const N8N_HOST = process.env.N8N_HOST || 'http://localhost:5678';
const N8N_API_KEY = process.env.N8N_API_KEY;

const api = axios.create({
    baseURL: `${N8N_HOST}/api/v1`,
    headers: { 'X-N8N-API-KEY': N8N_API_KEY, Accept: 'application/json', 'Content-Type': 'application/json' },
});

async function main() {
    console.log('🧹  Fetching all workflows…');
    const res = await api.get('/workflows?limit=100');
    const all = res.data?.data || [];
    const targets = all.filter(w =>
        w.name === 'Automated Outlook Invoice Downloader' ||
        w.name === '__FolderResolver (auto-delete)'
    );

    if (targets.length === 0) {
        console.log('✅  Nothing to clean up.');
        return;
    }

    console.log(`Found ${targets.length} orphan workflow(s) to delete.\n`);
    for (const w of targets) {
        try {
            await api.post(`/workflows/${w.id}/deactivate`, {});
        } catch (_) { }
        try {
            await api.delete(`/workflows/${w.id}`);
            console.log(`  🗑  Deleted [${w.id}] ${w.name}`);
        } catch (e) {
            console.log(`  ⚠️  Could not delete [${w.id}]: ${e.response?.data?.message || e.message}`);
        }
    }
    console.log('\n✅  Done. n8n is clean.');
}

main().catch(e => console.error('❌', e.message));