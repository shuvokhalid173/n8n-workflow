const axios = require('axios');
require('dotenv').config();

async function test() {
  const n8nApi = axios.create({
    baseURL: 'http://localhost:5678/api/v1',
    headers: { 'X-N8N-API-KEY': process.env.N8N_API_KEY, 'Accept': 'application/json', 'Content-Type': 'application/json' }
  });
  try {
    const res = await n8nApi.post('/workflows', { name: 'Test', nodes: [], connections: {}, settings: {} });
    console.log('Created:', res.data.id);
    const id = res.data.id;
    // try activate
    const res2 = await n8nApi.post(`/workflows/${id}/activate`, {});
    console.log('Activated?', res2.data);
  } catch(e) {
    console.error(e.response ? e.response.data : e.message);
  }
}
test();
