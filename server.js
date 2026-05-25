const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const { Pool } = require('pg');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── DATABASE ──────────────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

async function initDB() {
  const client = await pool.connect();
  try {
    await client.query(`
      CREATE TABLE IF NOT EXISTS subs (
        id SERIAL PRIMARY KEY,
        codice VARCHAR(50) NOT NULL,
        ex_sub VARCHAR(50),
        location VARCHAR(50),
        piano VARCHAR(100),
        inquilino_id INTEGER,
        note TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS fornitori (
        id SERIAL PRIMARY KEY,
        codice_zuc VARCHAR(50),
        ragione_sociale VARCHAR(200) NOT NULL,
        piva VARCHAR(20), cf VARCHAR(20),
        indirizzo TEXT, cap VARCHAR(10),
        citta VARCHAR(100), provincia VARCHAR(5),
        tel VARCHAR(50), email VARCHAR(100), spec VARCHAR(200),
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS inquilini (
        id SERIAL PRIMARY KEY,
        codice_zuc VARCHAR(50),
        ragione_sociale VARCHAR(200) NOT NULL,
        piva VARCHAR(20), cf VARCHAR(20),
        indirizzo TEXT, cap VARCHAR(10),
        citta VARCHAR(100), provincia VARCHAR(5),
        tel VARCHAR(50), email VARCHAR(100),
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS interventi (
        id SERIAL PRIMARY KEY,
        location VARCHAR(50),
        sub_id INTEGER, inquilino_id INTEGER, fornitore_id INTEGER,
        protocollo VARCHAR(100), num_fattura VARCHAR(100),
        data_intervento DATE, data_fattura DATE,
        prezzo DECIMAL(10,2), descrizione TEXT, note TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS settings (
        key VARCHAR(50) PRIMARY KEY,
        value TEXT
      );
      INSERT INTO settings (key, value) VALUES ('password', 'immobili2024')
        ON CONFLICT (key) DO NOTHING;
    `);
    console.log('✅ Database inizializzato');
  } finally { client.release(); }
}

// ── AUTH ──────────────────────────────────────────────────────
app.post('/api/auth/login', async (req, res) => {
  const { password } = req.body;
  const r = await pool.query("SELECT value FROM settings WHERE key='password'");
  if (r.rows[0]?.value === password) res.json({ ok: true });
  else res.status(401).json({ ok: false });
});

app.post('/api/auth/change-password', async (req, res) => {
  const { oldPassword, newPassword } = req.body;
  const r = await pool.query("SELECT value FROM settings WHERE key='password'");
  if (r.rows[0]?.value !== oldPassword) return res.status(401).json({ error: 'Password errata' });
  await pool.query("UPDATE settings SET value=$1 WHERE key='password'", [newPassword]);
  res.json({ ok: true });
});

// ── SUBS ──────────────────────────────────────────────────────
app.get('/api/subs', async (req, res) => {
  const r = await pool.query('SELECT * FROM subs ORDER BY id');
  res.json(r.rows);
});
app.post('/api/subs', async (req, res) => {
  const { codice, ex_sub, location, piano, inquilino_id, note } = req.body;
  const r = await pool.query(
    'INSERT INTO subs (codice,ex_sub,location,piano,inquilino_id,note) VALUES ($1,$2,$3,$4,$5,$6) RETURNING *',
    [codice, ex_sub||null, location, piano||null, inquilino_id||null, note||null]
  );
  res.json(r.rows[0]);
});
app.put('/api/subs/:id', async (req, res) => {
  const { codice, ex_sub, location, piano, inquilino_id, note } = req.body;
  const r = await pool.query(
    'UPDATE subs SET codice=$1,ex_sub=$2,location=$3,piano=$4,inquilino_id=$5,note=$6 WHERE id=$7 RETURNING *',
    [codice, ex_sub||null, location, piano||null, inquilino_id||null, note||null, req.params.id]
  );
  res.json(r.rows[0]);
});
app.delete('/api/subs/:id', async (req, res) => {
  await pool.query('DELETE FROM subs WHERE id=$1', [req.params.id]);
  res.json({ ok: true });
});

// ── FORNITORI ─────────────────────────────────────────────────
app.get('/api/fornitori', async (req, res) => {
  const r = await pool.query('SELECT * FROM fornitori ORDER BY ragione_sociale');
  res.json(r.rows);
});
app.post('/api/fornitori', async (req, res) => {
  const f = req.body;
  const r = await pool.query(
    `INSERT INTO fornitori (codice_zuc,ragione_sociale,piva,cf,indirizzo,cap,citta,provincia,tel,email,spec)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING *`,
    [f.codice_zuc||null,f.ragione_sociale,f.piva||null,f.cf||null,f.indirizzo||null,f.cap||null,f.citta||null,f.provincia||null,f.tel||null,f.email||null,f.spec||null]
  );
  res.json(r.rows[0]);
});
app.put('/api/fornitori/:id', async (req, res) => {
  const f = req.body;
  const r = await pool.query(
    `UPDATE fornitori SET codice_zuc=$1,ragione_sociale=$2,piva=$3,cf=$4,indirizzo=$5,cap=$6,citta=$7,provincia=$8,tel=$9,email=$10,spec=$11 WHERE id=$12 RETURNING *`,
    [f.codice_zuc||null,f.ragione_sociale,f.piva||null,f.cf||null,f.indirizzo||null,f.cap||null,f.citta||null,f.provincia||null,f.tel||null,f.email||null,f.spec||null,req.params.id]
  );
  res.json(r.rows[0]);
});
app.delete('/api/fornitori/:id', async (req, res) => {
  await pool.query('DELETE FROM fornitori WHERE id=$1', [req.params.id]);
  res.json({ ok: true });
});

// Import massivo fornitori (Zucchetti)
app.post('/api/fornitori/import-bulk', async (req, res) => {
  const { items } = req.body;
  const client = await pool.connect();
  let added = 0, skipped = 0;
  try {
    await client.query('BEGIN');
    for (const f of items) {
      const exists = await client.query(
        'SELECT id FROM fornitori WHERE LOWER(TRIM(ragione_sociale))=LOWER(TRIM($1))', [f.ragione_sociale]
      );
      if (exists.rows.length) { skipped++; continue; }
      await client.query(
        `INSERT INTO fornitori (codice_zuc,ragione_sociale,piva,cf,indirizzo,cap,citta,provincia,tel,email)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)`,
        [f.codice_zuc||null,f.ragione_sociale,f.piva||null,f.cf||null,f.indirizzo||null,f.cap||null,f.citta||null,f.provincia||null,f.tel||null,f.email||null]
      );
      added++;
    }
    await client.query('COMMIT');
    res.json({ added, skipped });
  } catch(e) { await client.query('ROLLBACK'); res.status(500).json({ error: e.message }); }
  finally { client.release(); }
});

// ── INQUILINI ─────────────────────────────────────────────────
app.get('/api/inquilini', async (req, res) => {
  const r = await pool.query('SELECT * FROM inquilini ORDER BY ragione_sociale');
  res.json(r.rows);
});
app.post('/api/inquilini', async (req, res) => {
  const i = req.body;
  const r = await pool.query(
    `INSERT INTO inquilini (codice_zuc,ragione_sociale,piva,cf,indirizzo,cap,citta,provincia,tel,email)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) RETURNING *`,
    [i.codice_zuc||null,i.ragione_sociale,i.piva||null,i.cf||null,i.indirizzo||null,i.cap||null,i.citta||null,i.provincia||null,i.tel||null,i.email||null]
  );
  res.json(r.rows[0]);
});
app.put('/api/inquilini/:id', async (req, res) => {
  const i = req.body;
  const r = await pool.query(
    `UPDATE inquilini SET codice_zuc=$1,ragione_sociale=$2,piva=$3,cf=$4,indirizzo=$5,cap=$6,citta=$7,provincia=$8,tel=$9,email=$10 WHERE id=$11 RETURNING *`,
    [i.codice_zuc||null,i.ragione_sociale,i.piva||null,i.cf||null,i.indirizzo||null,i.cap||null,i.citta||null,i.provincia||null,i.tel||null,i.email||null,req.params.id]
  );
  res.json(r.rows[0]);
});
app.delete('/api/inquilini/:id', async (req, res) => {
  await pool.query('DELETE FROM inquilini WHERE id=$1', [req.params.id]);
  res.json({ ok: true });
});

app.post('/api/inquilini/import-bulk', async (req, res) => {
  const { items } = req.body;
  const client = await pool.connect();
  let added = 0, skipped = 0;
  try {
    await client.query('BEGIN');
    for (const i of items) {
      const exists = await client.query(
        'SELECT id FROM inquilini WHERE LOWER(TRIM(ragione_sociale))=LOWER(TRIM($1))', [i.ragione_sociale]
      );
      if (exists.rows.length) { skipped++; continue; }
      await client.query(
        `INSERT INTO inquilini (codice_zuc,ragione_sociale,piva,cf,indirizzo,cap,citta,provincia,tel,email)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)`,
        [i.codice_zuc||null,i.ragione_sociale,i.piva||null,i.cf||null,i.indirizzo||null,i.cap||null,i.citta||null,i.provincia||null,i.tel||null,i.email||null]
      );
      added++;
    }
    await client.query('COMMIT');
    res.json({ added, skipped });
  } catch(e) { await client.query('ROLLBACK'); res.status(500).json({ error: e.message }); }
  finally { client.release(); }
});

// ── INTERVENTI ────────────────────────────────────────────────
app.get('/api/interventi', async (req, res) => {
  const r = await pool.query('SELECT * FROM interventi ORDER BY id');
  res.json(r.rows);
});

app.post('/api/interventi/check-duplicate', async (req, res) => {
  const { sub_id, fornitore_id, descrizione } = req.body;
  const r = await pool.query(
    'SELECT * FROM interventi WHERE sub_id=$1 AND fornitore_id=$2', [sub_id, fornitore_id]
  );
  const words = (descrizione||'').toLowerCase().split(/\s+/).filter(w=>w.length>3);
  const similar = r.rows.filter(x => {
    const d = (x.descrizione||'').toLowerCase();
    const matches = words.filter(w=>d.includes(w)).length;
    return words.length && matches/words.length > 0.4;
  });
  res.json({ duplicates: similar });
});

app.post('/api/interventi', async (req, res) => {
  const v = req.body;
  const r = await pool.query(
    `INSERT INTO interventi (location,sub_id,inquilino_id,fornitore_id,protocollo,num_fattura,data_intervento,data_fattura,prezzo,descrizione,note)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING *`,
    [v.location,v.sub_id||null,v.inquilino_id||null,v.fornitore_id||null,v.protocollo||null,v.num_fattura||null,v.data_intervento||null,v.data_fattura||null,v.prezzo||null,v.descrizione,v.note||null]
  );
  res.json(r.rows[0]);
});

app.put('/api/interventi/:id', async (req, res) => {
  const v = req.body;
  const r = await pool.query(
    `UPDATE interventi SET location=$1,sub_id=$2,inquilino_id=$3,fornitore_id=$4,protocollo=$5,num_fattura=$6,data_intervento=$7,data_fattura=$8,prezzo=$9,descrizione=$10,note=$11 WHERE id=$12 RETURNING *`,
    [v.location,v.sub_id||null,v.inquilino_id||null,v.fornitore_id||null,v.protocollo||null,v.num_fattura||null,v.data_intervento||null,v.data_fattura||null,v.prezzo||null,v.descrizione,v.note||null,req.params.id]
  );
  res.json(r.rows[0]);
});

app.delete('/api/interventi/:id', async (req, res) => {
  await pool.query('DELETE FROM interventi WHERE id=$1', [req.params.id]);
  res.json({ ok: true });
});

// Import storico massivo con smart matching
app.post('/api/interventi/import-storico', async (req, res) => {
  const { rows } = req.body;
  const client = await pool.connect();
  let added = 0, errors = [];
  try {
    await client.query('BEGIN');
    const fornitori = (await client.query('SELECT * FROM fornitori')).rows;
    const inquilini = (await client.query('SELECT * FROM inquilini')).rows;
    const subs = (await client.query('SELECT * FROM subs')).rows;

    const normalize = s => (s||'').toLowerCase().trim().replace(/\s+/g,' ');
    const findOrCreate = async (table, arr, nome) => {
      if (!nome) return null;
      const n = normalize(nome);
      let found = arr.find(x => normalize(x.ragione_sociale) === n);
      if (!found) {
        found = arr.find(x => normalize(x.ragione_sociale).includes(n) || n.includes(normalize(x.ragione_sociale)));
      }
      if (found) return found.id;
      const col = table === 'fornitori' ? 'ragione_sociale' : 'ragione_sociale';
      const r = await client.query(`INSERT INTO ${table} (ragione_sociale) VALUES ($1) RETURNING *`, [nome.trim()]);
      const newItem = r.rows[0];
      arr.push(newItem);
      return newItem.id;
    };

    for (const row of rows) {
      try {
        const subNorm = normalize(row.sub_codice);
        const sub = subs.find(s =>
          normalize(s.codice) === subNorm || normalize(s.ex_sub||'') === subNorm
        );
        const fornitore_id = await findOrCreate('fornitori', fornitori, row.fornitore_nome);
        const inquilino_id = row.inquilino_nome ? await findOrCreate('inquilini', inquilini, row.inquilino_nome) : null;

        // Parse date
        const parseDate = d => {
          if (!d) return null;
          const s = String(d).trim();
          if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
          if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s.split('/').reverse().join('-');
          if (/^\d{2}\/\d{2}\/\d{2}$/.test(s)) {
            const [dd,mm,yy] = s.split('/');
            return `20${yy}-${mm}-${dd}`;
          }
          return null;
        };

        await client.query(
          `INSERT INTO interventi (location,sub_id,fornitore_id,inquilino_id,protocollo,num_fattura,data_intervento,data_fattura,prezzo,descrizione,note)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)`,
          [row.location||null, sub?.id||null, fornitore_id, inquilino_id, row.protocollo||null, row.num_fattura||null,
           parseDate(row.data_intervento), parseDate(row.data_fattura),
           parseFloat(row.prezzo)||null, row.descrizione||null, row.note||null]
        );
        added++;
      } catch(e) { errors.push({ row, error: e.message }); }
    }
    await client.query('COMMIT');
    res.json({ added, errors });
  } catch(e) { await client.query('ROLLBACK'); res.status(500).json({ error: e.message }); }
  finally { client.release(); }
});

// ── RIEPILOGO ─────────────────────────────────────────────────
app.get('/api/riepilogo', async (req, res) => {
  const r = await pool.query(`
    SELECT i.sub_id, s.codice as sub, s.ex_sub, i.location,
           iq.ragione_sociale as inquilino,
           COUNT(i.id) as count,
           COALESCE(SUM(i.prezzo),0) as totale,
           json_agg(json_build_object('fornitore', f.ragione_sociale, 'prezzo', COALESCE(i.prezzo,0))) as dettagli
    FROM interventi i
    LEFT JOIN subs s ON i.sub_id = s.id
    LEFT JOIN inquilini iq ON s.inquilino_id = iq.id
    LEFT JOIN fornitori f ON i.fornitore_id = f.id
    GROUP BY i.sub_id, s.codice, s.ex_sub, i.location, iq.ragione_sociale
    ORDER BY totale DESC
  `);
  const result = r.rows.map(row => {
    const fornitori = {};
    (row.dettagli||[]).forEach(d => {
      if (d.fornitore) fornitori[d.fornitore] = (fornitori[d.fornitore]||0) + parseFloat(d.prezzo);
    });
    return { ...row, totale: parseFloat(row.totale), count: parseInt(row.count), fornitori };
  });
  res.json(result);
});

// ── EXPORT EXCEL ──────────────────────────────────────────────
app.get('/api/export', async (req, res) => {
  const r = await pool.query(`
    SELECT s.codice as sub, s.ex_sub, i.location,
           iq.ragione_sociale as inquilino, f.ragione_sociale as fornitore,
           i.protocollo, i.data_intervento, i.data_fattura, i.num_fattura,
           i.prezzo, i.descrizione, i.note
    FROM interventi i
    LEFT JOIN subs s ON i.sub_id=s.id
    LEFT JOIN inquilini iq ON i.inquilino_id=iq.id
    LEFT JOIN fornitori f ON i.fornitore_id=f.id
    ORDER BY s.codice, i.id
  `);
  const rows = r.rows.map(row => ({
    'SUB': row.sub||'', 'Ex SUB': row.ex_sub||'', 'Sede': row.location||'',
    'Inquilino': row.inquilino||'', 'Fornitore': row.fornitore||'',
    'N° Protocollo': row.protocollo||'', 'Data Intervento': row.data_intervento||'',
    'Data Fattura': row.data_fattura||'', 'N° Fattura': row.num_fattura||'',
    'Prezzo (€)': row.prezzo||'', 'Descrizione': row.descrizione||'', 'Note': row.note||''
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Storico Interventi');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="storico_interventi.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── START ─────────────────────────────────────────────────────
initDB().then(() => {
  app.listen(PORT, '0.0.0.0', () => console.log(`\n✅ Gestionale avviato su http://localhost:${PORT}\n`));
}).catch(err => {
  console.error('Errore connessione DB:', err.message);
  process.exit(1);
});
