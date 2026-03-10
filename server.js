require('dotenv').config();
const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const https   = require('https');
const path    = require('path');
const { parse } = require('csv-parse/sync');

const app = express();

// ─── Safe Multer Wrapper — always returns JSON on error ───────────────────────
const _multer = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024 }, // 30 MB
});

function safeUpload(field) {
  return (req, res, next) => {
    _multer.single(field)(req, res, (err) => {
      if (!err) return next();
      return res.status(400).json({
        error: err instanceof multer.MulterError
          ? `ไฟล์ใหญ่เกินไป (${err.message})` 
          : err.message
      });
    });
  };
}

app.use(cors());
app.use(express.json({ limit: '150mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const RESEND_API_KEY = process.env.RESEND_API_KEY;
const FROM_EMAIL     = process.env.FROM_EMAIL || 'noreply@shd-technology.co.th';
const FROM_NAME      = process.env.FROM_NAME  || 'SHD Technology';

// Resend does not accept video/audio attachments
const BLOCKED_MIME_PREFIXES = ['video/', 'audio/'];

// ─── Resend API call ──────────────────────────────────────────────────────────
function callResend(payload) {
  return new Promise((resolve) => {
    const body = JSON.stringify(payload);
    const options = {
      hostname: 'api.resend.com',
      path: '/emails',
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${RESEND_API_KEY}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body),
      },
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (res.statusCode === 200 || res.statusCode === 201) {
            resolve({ success: true, id: json.id });
          } else {
            resolve({ success: false, error: json.message || json.name || `HTTP ${res.statusCode}` });
          }
        } catch (e) {
          resolve({ success: false, error: `Parse error: ${data.slice(0, 120)}` });
        }
      });
    });
    req.on('error', e => resolve({ success: false, error: e.message }));
    req.write(body);
    req.end();
  });
}

// ─── SSE helper ───────────────────────────────────────────────────────────────
function sseStream(res) {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  return (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);
}

// ═══════════════════════════════════════════════════════════════════════════════
// MODE 1 — Access Code Email
// ═══════════════════════════════════════════════════════════════════════════════
function buildAccessCodeEmail(email, code, senderNote = '') {
  const now  = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
  const year = new Date().getFullYear();
  return `<!DOCTYPE html>
<html lang="th"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>รหัสการเข้าใช้งาน</title></head>
<body style="margin:0;padding:0;background:#F1F5F9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#F1F5F9;"><tr><td align="center" style="padding:32px 16px 48px;">
<table width="580" cellpadding="0" cellspacing="0" style="max-width:580px;width:100%;">
<tr><td align="center" style="padding-bottom:20px;"><img src="https://shd-technology.co.th/images/logo.png" alt="SHD Technology" width="110" style="display:block;height:auto;"></td></tr>
<tr><td style="background:#1D4ED8;border-radius:12px 12px 0 0;padding:32px 40px 28px;">
  <p style="color:rgba(255,255,255,0.5);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:2.5px;margin:0 0 8px;">System Access Credentials</p>
  <h1 style="color:#fff;font-size:22px;font-weight:700;margin:0 0 6px;letter-spacing:-0.3px;">รหัสการเข้าใช้งานระบบ</h1>
  <p style="color:rgba(255,255,255,0.45);font-size:13px;margin:0;">${now}</p>
</td></tr>
<tr><td style="background:#fff;padding:36px 40px 32px;border:1px solid #E2E8F0;border-top:none;">
  <p style="color:#374151;font-size:15px;line-height:1.75;margin:0 0 24px;">เรียน คุณ,<br><br>บัญชีผู้ใช้งานของคุณใน <strong style="color:#1D4ED8;">SHD Technology</strong> พร้อมใช้งานแล้ว กรุณาใช้ข้อมูลด้านล่างเพื่อเข้าสู่ระบบครั้งแรก</p>
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;"><tr><td style="background:#F8FAFF;border:1px solid #DBEAFE;border-radius:10px;padding:28px 32px;">
    <p style="color:#6B7280;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:0 0 16px;text-align:center;">ข้อมูลการเข้าสู่ระบบ</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:12px;"><tr><td style="background:#fff;border:1px solid #E5E7EB;border-radius:8px;padding:14px 18px;">
      <p style="color:#9CA3AF;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin:0 0 4px;">Email Address</p>
      <p style="color:#111827;font-size:14px;font-weight:600;margin:0;font-family:'Courier New',monospace;">${email}</p>
    </td></tr></table>
    <table width="100%" cellpadding="0" cellspacing="0"><tr><td style="background:#1D4ED8;border-radius:8px;padding:18px 24px;text-align:center;">
      <p style="color:rgba(255,255,255,0.55);font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:2px;margin:0 0 8px;">Access Code</p>
      <p style="color:#fff;font-size:28px;font-weight:800;letter-spacing:8px;margin:0;font-family:'Courier New',monospace;">${code}</p>
    </td></tr></table>
  </td></tr></table>
  ${senderNote ? `<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr><td style="background:#F0FDF4;border-left:3px solid #16A34A;border-radius:0 8px 8px 0;padding:14px 18px;"><p style="color:#166534;font-size:13px;line-height:1.6;margin:0;"><strong>หมายเหตุ:</strong> ${senderNote}</p></td></tr></table>` : ''}
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;"><tr><td style="background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;padding:12px 16px;"><p style="color:#92400E;font-size:12px;line-height:1.6;margin:0;"><strong>ข้อควรระวัง:</strong> อย่าเปิดเผย Access Code กับผู้อื่น หากไม่ได้ร้องขอกรุณาติดต่อ IT ทันที</p></td></tr></table>
</td></tr>
<tr><td style="background:#1E293B;border-radius:0 0 12px 12px;padding:20px 40px;">
  <table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td><p style="color:#475569;font-size:11px;margin:0;line-height:1.6;">&copy; ${year} SHD Technology Co., Ltd.<br><a href="https://shd-technology.co.th" style="color:#60A5FA;text-decoration:none;">shd-technology.co.th</a></p></td>
    <td align="right"><p style="color:#334155;font-size:11px;margin:0;line-height:1.6;">อีเมลนี้ส่งโดยอัตโนมัติ<br>กรุณาอย่าตอบกลับ</p></td>
  </tr></table>
</td></tr>
</table></td></tr></table>
</body></html>`;
}

// POST /api/send-access-codes
app.post('/api/send-access-codes', async (req, res) => {
  const { employees, senderNote } = req.body;
  if (!employees || !Array.isArray(employees) || !employees.length) {
    return res.status(400).json({ error: 'ไม่มีข้อมูลพนักงาน' });
  }
  const emit = sseStream(res);
  emit({ type: 'start', total: employees.length });
  let success = 0, failed = 0;
  const CONCURRENCY = 3;
  for (let i = 0; i < employees.length; i += CONCURRENCY) {
    const batch = employees.slice(i, i + CONCURRENCY);
    const results = await Promise.all(batch.map(emp => callResend({
      from: `${FROM_NAME} <${FROM_EMAIL}>`,
      to: [emp.email],
      subject: '[SHD Technology] รหัสการเข้าใช้งานระบบของคุณ',
      html: buildAccessCodeEmail(emp.email, emp.code, senderNote || ''),
    })));
    results.forEach((r, idx) => {
      const emp = batch[idx];
      if (r.success) { success++; emit({ type: 'progress', email: emp.email, status: 'ok', success, failed, total: employees.length }); }
      else           { failed++;  emit({ type: 'progress', email: emp.email, status: 'err', reason: r.error, success, failed, total: employees.length }); }
    });
    if (i + CONCURRENCY < employees.length) await new Promise(r => setTimeout(r, 300));
  }
  emit({ type: 'done', success, failed, total: employees.length });
  res.end();
});

// POST /api/preview-access-code
app.post('/api/preview-access-code', (req, res) => {
  const { email = 'employee@shd-technology.co.th', code = 'EMP2024-XYZ', note = '' } = req.body;
  res.send(buildAccessCodeEmail(email, code, note));
});

// ═══════════════════════════════════════════════════════════════════════════════
// MODE 2 — Custom Email Composer
// ═══════════════════════════════════════════════════════════════════════════════

// POST /api/send-custom
app.post('/api/send-custom', async (req, res) => {
  const { recipients, subject, htmlBody, fromName, replyTo, attachments = [] } = req.body;

  if (!recipients || !Array.isArray(recipients) || !recipients.length) {
    return res.status(400).json({ error: 'ไม่มีผู้รับ' });
  }
  if (!subject || !subject.trim()) {
    return res.status(400).json({ error: 'ไม่มีหัวข้ออีเมล' });
  }
  if (!htmlBody || !htmlBody.trim()) {
    return res.status(400).json({ error: 'ไม่มีเนื้อหาอีเมล' });
  }

  const emit = sseStream(res);
  emit({ type: 'start', total: recipients.length });

  const senderName = (fromName && fromName.trim()) ? fromName.trim() : FROM_NAME;

  // Build Resend attachment array (only supported MIME)
  const resendAttachments = attachments
    .filter(a => !BLOCKED_MIME_PREFIXES.some(p => (a.contentType || '').startsWith(p)))
    .map(a => ({ filename: a.filename, content: a.content }));

  let success = 0, failed = 0;
  const CONCURRENCY = 3;

  for (let i = 0; i < recipients.length; i += CONCURRENCY) {
    const batch = recipients.slice(i, i + CONCURRENCY);
    const results = await Promise.all(batch.map(email => {
      const payload = {
        from: `${senderName} <${FROM_EMAIL}>`,
        to: [email.trim()],
        subject: subject.trim(),
        html: htmlBody,
        ...(replyTo && replyTo.trim() && { reply_to: replyTo.trim() }),
        ...(resendAttachments.length > 0 && { attachments: resendAttachments }),
      };
      return callResend(payload);
    }));

    results.forEach((r, idx) => {
      const email = batch[idx];
      if (r.success) { success++; emit({ type: 'progress', email, status: 'ok', success, failed, total: recipients.length }); }
      else           { failed++;  emit({ type: 'progress', email, status: 'err', reason: r.error, success, failed, total: recipients.length }); }
    });
    if (i + CONCURRENCY < recipients.length) await new Promise(r => setTimeout(r, 300));
  }

  emit({ type: 'done', success, failed, total: recipients.length });
  res.end();
});

// ═══════════════════════════════════════════════════════════════════════════════
// Shared Routes
// ═══════════════════════════════════════════════════════════════════════════════

// POST /api/parse-file
app.post('/api/parse-file', safeUpload('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'ไม่พบไฟล์' });
  const filename = req.file.originalname.toLowerCase();
  let rows = [];
  try {
    if (filename.endsWith('.csv')) {
      rows = parse(req.file.buffer.toString('utf-8'), { skip_empty_lines: true, trim: true });
    } else if (filename.endsWith('.xlsx') || filename.endsWith('.xls')) {
      const XLSX = require('xlsx');
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }).map(r => r.map(c => String(c).trim()));
    } else {
      return res.status(400).json({ error: 'รองรับเฉพาะ .csv, .xlsx, .xls' });
    }
    let emailCol = -1, codeCol = -1;
    (rows[0] || []).forEach((h, i) => {
      const l = String(h).toLowerCase();
      if (l.includes('email') || l.includes('อีเมล') || l.includes('mail')) emailCol = i;
      if (l.includes('code') || l.includes('รหัส') || l.includes('password') || l.includes('pass')) codeCol = i;
    });
    res.json({ rows, emailCol, codeCol, headers: rows[0] || [] });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/upload-attachment — receives file, returns base64 to frontend
app.post('/api/upload-attachment', safeUpload('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'ไม่พบไฟล์' });

  const mimeType = req.file.mimetype || 'application/octet-stream';

  // Reject video / audio — Resend does not support them
  if (BLOCKED_MIME_PREFIXES.some(prefix => mimeType.startsWith(prefix))) {
    return res.status(400).json({
      error: `ไฟล์ประเภท Video/Audio ไม่รองรับ — Resend API ไม่สามารถส่งไฟล์ .mp4, .avi, .mp3 ฯลฯ ได้ กรุณาใช้ PDF, Word, Excel, รูปภาพ หรือ ZIP แทน`
    });
  }

  res.json({
    filename: req.file.originalname,
    contentType: mimeType,
    size: req.file.size,
    content: req.file.buffer.toString('base64'),
  });
});

// ─── Global error handler — always JSON ───────────────────────────────────────
app.use((err, req, res, next) => {
  console.error('Server error:', err);
  res.status(500).json({ error: err.message || 'Internal server error' });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`\n  SHD Mailer  →  http://localhost:${PORT}\n`);
});