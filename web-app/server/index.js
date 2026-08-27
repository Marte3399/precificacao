import cors from 'cors';
import express from 'express';
import { getEntriesBySystem, initializeDatabase, replaceEntriesBySystem } from './db.js';

const app = express();
const port = Number(process.env.PORT || 3001);

app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.get('/health', (_req, res) => {
  res.json({ ok: true });
});

app.get('/api/daily-grid/:systemKey', async (req, res) => {
  try {
    const { systemKey } = req.params;
    const rows = await getEntriesBySystem(systemKey);
    res.json({ rows });
  } catch (error) {
    console.error('Erro ao consultar registros daily:', error);
    res.status(500).json({ message: 'Erro ao consultar registros daily.' });
  }
});

app.put('/api/daily-grid/:systemKey', async (req, res) => {
  try {
    const { systemKey } = req.params;
    const { rows } = req.body || {};

    if (!Array.isArray(rows)) {
      res.status(400).json({ message: 'Payload inválido. Envie { rows: [] }.' });
      return;
    }

    await replaceEntriesBySystem(systemKey, rows);
    res.json({ ok: true });
  } catch (error) {
    console.error('Erro ao salvar registros daily:', error);
    res.status(500).json({ message: 'Erro ao salvar registros daily.' });
  }
});

initializeDatabase()
  .then(() => {
    app.listen(port, () => {
      console.log(`API daily rodando em http://localhost:${port}`);
    });
  })
  .catch((error) => {
    console.error('Falha ao iniciar banco/API:', error);
    process.exit(1);
  });
