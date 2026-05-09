const DEFAULT_REPORT_RECIPIENT = 'ale21rock@gmail.com';

function doPost(e) {
  try {
    const payload = parseRequestPayload_(e);
    const action = String(payload.action || '').toLowerCase();

    if (action !== 'sendreport') {
      return jsonResponse_({ ok: false, error: 'Unsupported action.' }, 400);
    }

    const result = sendCobrosReport_(payload);
    return jsonResponse_(result);
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) }, 500);
  }
}

function parseRequestPayload_(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  return {};
}

function sendCobrosReport_(payload) {
  const report = payload.report || {};
  const rows = Array.isArray(report.rows) ? report.rows : [];
  const recipient = String(payload.recipient || DEFAULT_REPORT_RECIPIENT).trim();
  const subject = String(payload.subject || `Cobros ${report.monthName || ''} ${report.year || ''}`).trim();

  if (!recipient) throw new Error('Missing recipient email.');

  const html = buildCobrosReportHtml_(report, rows, payload.generatedAt);
  const pdfBlob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .getAs(MimeType.PDF)
    .setName(`${subject.replace(/[^a-z0-9]+/gi, '_').replace(/_+/g, '_').replace(/^_|_$/g, '') || 'cobros_report'}.pdf`);

  GmailApp.sendEmail(
    recipient,
    subject,
    buildCobrosPlainText_(report, rows),
    {
      attachments: [pdfBlob],
      name: 'Control de Gastos'
    }
  );

  return { ok: true, recipient, subject };
}

function buildCobrosPlainText_(report, rows) {
  const lines = [
    `Cobros ${report.monthName || ''} ${report.year || ''}`.trim(),
    `Total a cobrar: ${formatCurrency_(report.totalToCollect)}`,
    `Total pagado: ${formatCurrency_(report.totalPaid)}`,
    `Total descontado: ${formatCurrency_(report.totalDiscounted)}`,
    ''
  ];
  rows.forEach(row => {
    lines.push(`${row.nombre}: total ${formatCurrency_(row.total)} | pagado ${formatCurrency_(row.pagado)} | desc. ${formatCurrency_(row.descuento)} | restante ${formatCurrency_(row.restante)}`);
  });
  return lines.join('\n');
}

function buildCobrosReportHtml_(report, rows, generatedAt) {
  const totalToCollect = formatCurrency_(report.totalToCollect);
  const totalPaid = formatCurrency_(report.totalPaid);
  const totalDiscounted = formatCurrency_(report.totalDiscounted);
  const tableRows = rows.map(row => `
    <tr>
      <td>${escapeHtml_(row.nombre)}</td>
      <td>${formatCurrency_(row.total)}</td>
      <td>${formatCurrency_(row.pagado)}</td>
      <td>${formatCurrency_(row.descuento)}</td>
      <td>${formatCurrency_(row.restante)}</td>
    </tr>
  `).join('');

  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <style>
    body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:28px;color:#111827;background:#fff}
    h1{margin:0 0 6px;font-size:28px}
    .sub{color:#4b5563;margin:0 0 16px;font-size:14px}
    .meta{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:18px 0 22px}
    .card{border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;background:#f9fafb}
    .card span{display:block;color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.06em;font-weight:700;margin-bottom:6px}
    .card strong{font-size:22px}
    table{width:100%;border-collapse:collapse;margin-top:8px}
    th,td{padding:10px 8px;border-bottom:1px solid #e5e7eb;text-align:left;font-size:13px}
    th{background:#f3f4f6;color:#374151;text-transform:uppercase;letter-spacing:.05em;font-size:11px}
    td:last-child,th:last-child{text-align:right}
    .note{margin-top:14px;color:#6b7280;font-size:12px}
    @media print{body{padding:16px}.note{position:fixed;bottom:12px;left:16px;right:16px}}
  </style>
</head>
<body>
  <h1>Cobros a cobrar - ${escapeHtml_(String(report.monthName || ''))} ${escapeHtml_(String(report.year || ''))}</h1>
  <p class="sub">Solo muestra lo que paga cada persona, lo pagado, los descuentos y lo pendiente.</p>
  <div class="meta">
    <div class="card"><span>Total a cobrar</span><strong>${totalToCollect}</strong></div>
    <div class="card"><span>Total pagado</span><strong>${totalPaid}</strong></div>
    <div class="card"><span>Total descontado</span><strong>${totalDiscounted}</strong></div>
  </div>
  <table>
    <thead>
      <tr><th>Persona</th><th>Total</th><th>Pagado</th><th>Descuento</th><th>Restante</th></tr>
    </thead>
    <tbody>${tableRows}</tbody>
  </table>
  <p class="note">Generado ${escapeHtml_(String(generatedAt || new Date().toISOString()))} desde Control de Gastos.</p>
</body>
</html>`;
}

function escapeHtml_(value) {
  return String(value).replace(/[&<>"]/g, char => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;'
  }[char]));
}

function formatCurrency_(value) {
  const number = Number(value || 0);
  return `$${number.toLocaleString('es-UY', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function jsonResponse_(payload, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(payload));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}