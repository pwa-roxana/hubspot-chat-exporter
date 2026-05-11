const express = require('express');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const HS_TOKEN = process.env.HUBSPOT_TOKEN;
const HS_BASE = 'https://api.hubapi.com';

function hsHeaders() {
  return {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + HS_TOKEN
  };
}

// Search contacts
app.post('/api/contacts/search', async (req, res) => {
  try {
    const r = await fetch(`${HS_BASE}/crm/v3/objects/contacts/search`, {
      method: 'POST',
      headers: hsHeaders(),
      body: JSON.stringify(req.body)
    });
    const data = await r.json();
    res.status(r.status).json(data);
  } catch (e) {
    res.status(500).json({ message: e.message });
  }
});

// Create contact
app.post('/api/contacts', async (req, res) => {
  try {
    const r = await fetch(`${HS_BASE}/crm/v3/objects/contacts`, {
      method: 'POST',
      headers: hsHeaders(),
      body: JSON.stringify(req.body)
    });
    const data = await r.json();
    res.status(r.status).json(data);
  } catch (e) {
    res.status(500).json({ message: e.message });
  }
});

// Update contact
app.patch('/api/contacts/:id', async (req, res) => {
  try {
    const r = await fetch(`${HS_BASE}/crm/v3/objects/contacts/${req.params.id}`, {
      method: 'PATCH',
      headers: hsHeaders(),
      body: JSON.stringify(req.body)
    });
    const data = await r.json();
    res.status(r.status).json(data);
  } catch (e) {
    res.status(500).json({ message: e.message });
  }
});

// Log note
app.post('/api/notes', async (req, res) => {
  try {
    const r = await fetch(`${HS_BASE}/crm/v3/objects/notes`, {
      method: 'POST',
      headers: hsHeaders(),
      body: JSON.stringify(req.body)
    });
    const data = await r.json();
    res.status(r.status).json(data);
  } catch (e) {
    res.status(500).json({ message: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PWA Call Logger running on port ${PORT}`));
