import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import multer from 'multer';
import { readSheet, detectMapping, analyze, buildWorkbook } from './analyze.js';

const app = express();
const PORT = process.env.PORT || 5000;
const origins = (process.env.CORS_ORIGIN || 'http://localhost:3000').split(',');

app.use(helmet());
app.use(cors({ origin: origins, exposedHeaders: ['Content-Disposition'] }));
app.use(rateLimit({ windowMs: 60_000, limit: 120 }));

const upload = multer({
  storage: multer.memoryStorage(), // nothing is saved on disk / DB
  limits: { fileSize: 25 * 1024 * 1024 },
  fileFilter: (_req, f, cb) => (/\.(xlsx|xls|csv)$/i.test(f.originalname) ? cb(null, true) : cb(new Error('Only .xlsx, .xls and .csv files are allowed'))),
});

const parse = (s, d) => { try { return s ? JSON.parse(s) : d; } catch { return d; } };

function run(req) {
  if (!req.file) throw Object.assign(new Error('No file was uploaded'), { status: 400 });
  const { rows, columns } = readSheet(req.file.buffer);
  if (!rows.length) throw Object.assign(new Error('The sheet is empty'), { status: 422 });
  const picked = Object.fromEntries(Object.entries(parse(req.body.mapping, {})).filter(([, v]) => v));
  const mapping = { ...detectMapping(columns, rows), ...picked };
  if (!mapping.produced)
    throw Object.assign(new Error('Could not detect the production column. Please select it in the column mapping.'), { status: 422, columns });
  const result = analyze(rows, mapping, parse(req.body.filters, {}));
  result.meta.columns = columns;
  return result;
}

app.get('/api/health', (_req, res) => res.json({ ok: true }));

app.post('/api/analyze', upload.single('file'), (req, res, next) => {
  try { res.json(run(req)); } catch (e) { next(e); }
});

app.post('/api/export', upload.single('file'), (req, res, next) => {
  try {
    const buf = buildWorkbook(run(req));
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="production-report.xlsx"');
    res.send(buf);
  } catch (e) { next(e); }
});

app.use((err, _req, res, _next) => {
  const status = err.status || (err.code === 'LIMIT_FILE_SIZE' ? 413 : 400);
  if (status >= 500) console.error(err);
  res.status(status).json({ error: err.message || 'Server error', columns: err.columns });
});

app.listen(PORT, () => console.log(`API running on :${PORT}`));
