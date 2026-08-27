import fs from 'node:fs';
import path from 'node:path';
import sqlite3 from 'sqlite3';

const databaseDirectory = path.resolve('server', 'data');
const databasePath = path.join(databaseDirectory, 'daily.db');

if (!fs.existsSync(databaseDirectory)) {
  fs.mkdirSync(databaseDirectory, { recursive: true });
}

const db = new sqlite3.Database(databasePath);

function run(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function onRun(error) {
      if (error) {
        reject(error);
        return;
      }
      resolve(this);
    });
  });
}

function all(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (error, rows) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(rows);
    });
  });
}

export async function initializeDatabase() {
  await run(`
    CREATE TABLE IF NOT EXISTS daily_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      system_key TEXT NOT NULL,
      row_order INTEGER NOT NULL,
      prioridade TEXT,
      ticket TEXT,
      descricao TEXT,
      status TEXT,
      responsavel TEXT,
      entrada TEXT,
      prazo TEXT,
      entrega TEXT,
      observacoes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `);

  await run('CREATE INDEX IF NOT EXISTS idx_daily_entries_system ON daily_entries(system_key)');
  await run('CREATE INDEX IF NOT EXISTS idx_daily_entries_order ON daily_entries(system_key, row_order)');
}

export async function getEntriesBySystem(systemKey) {
  const rows = await all(
    `
      SELECT
        row_order,
        prioridade,
        ticket,
        descricao,
        status,
        responsavel,
        entrada,
        prazo,
        entrega,
        observacoes
      FROM daily_entries
      WHERE system_key = ?
      ORDER BY row_order ASC
    `,
    [systemKey]
  );

  return rows.map((row) => ({
    prioridade: row.prioridade || '',
    ticket: row.ticket || '',
    descricao: row.descricao || '',
    status: row.status || '',
    responsavel: row.responsavel || '',
    entrada: row.entrada || '',
    prazo: row.prazo || '',
    entrega: row.entrega || '',
    observacoes: row.observacoes || ''
  }));
}

export async function replaceEntriesBySystem(systemKey, rows) {
  await run('BEGIN TRANSACTION');
  try {
    await run('DELETE FROM daily_entries WHERE system_key = ?', [systemKey]);

    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index] || {};
      await run(
        `
          INSERT INTO daily_entries (
            system_key,
            row_order,
            prioridade,
            ticket,
            descricao,
            status,
            responsavel,
            entrada,
            prazo,
            entrega,
            observacoes,
            updated_at
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
        `,
        [
          systemKey,
          index,
          String(row.prioridade || ''),
          String(row.ticket || ''),
          String(row.descricao || ''),
          String(row.status || ''),
          String(row.responsavel || ''),
          String(row.entrada || ''),
          String(row.prazo || ''),
          String(row.entrega || ''),
          String(row.observacoes || '')
        ]
      );
    }

    await run('COMMIT');
  } catch (error) {
    await run('ROLLBACK');
    throw error;
  }
}
