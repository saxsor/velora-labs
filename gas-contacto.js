/**
 * Velora Labs — Google Apps Script
 * Recibe el formulario de contacto, guarda en Sheets y manda email de notificación.
 *
 * SETUP (una sola vez):
 *  1. Abre script.google.com → Nuevo proyecto
 *  2. Pega este código
 *  3. Cambia NOTIFICATION_EMAIL a tu correo
 *  4. Cambia SHEET_NAME si quieres otro nombre de hoja
 *  5. Guarda (Ctrl+S) → Implementar → Nueva implementación
 *     - Tipo: Aplicación web
 *     - Ejecutar como: Yo (tu cuenta)
 *     - Quién tiene acceso: Cualquier persona
 *  6. Copia la URL que te da → pégala en contacto.html donde dice GAS_URL
 *  7. Listo. Cada envío llega a Sheets y a tu correo.
 */

const NOTIFICATION_EMAIL = 'jalilbonilla@outlook.com';
const SHEET_NAME          = 'Leads Velora';

/* ─── HEADERS de hoja ──────────────────────────────── */
const COLUMNS = [
  'Fecha',
  'Nombre',
  'Empresa',
  'Email',
  'Teléfono',
  'Sitio web',
  'Necesidad',
  'Mensaje',
];

/* ─── doPost ───────────────────────────────────────── */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    appendToSheet(data);
    sendNotificationEmail(data);
  } catch (err) {
    Logger.log('Error: ' + err.message);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ─── Guardar en Sheets ────────────────────────────── */
function appendToSheet(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, COLUMNS.length)
         .setValues([COLUMNS])
         .setFontWeight('bold')
         .setBackground('#5B5FE8')
         .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    data.fecha     || new Date().toLocaleString('es-MX'),
    data.nombre    || '',
    data.empresa   || '',
    data.email     || '',
    data.telefono  || '',
    data.sitio     || '—',
    data.necesidad || '',
    data.mensaje   || '',
  ]);
}

/* ─── Email de notificación ────────────────────────── */
function sendNotificationEmail(data) {
  const subject = `[Velora Labs] Nuevo lead — ${data.nombre} · ${data.empresa}`;

  const html = `
<div style="font-family:'Helvetica Neue',Arial,sans-serif;max-width:560px;margin:0 auto;background:#0E0D12;color:#E8EAF4;border-radius:12px;overflow:hidden;border:1px solid #1E1C28;">

  <!-- Header -->
  <div style="background:#5B5FE8;padding:28px 32px;">
    <p style="margin:0;font-size:11px;letter-spacing:2px;text-transform:uppercase;opacity:.7;font-family:monospace;">Velora Labs</p>
    <h1 style="margin:6px 0 0;font-size:22px;font-weight:700;color:#fff;">Nuevo diagnóstico recibido</h1>
  </div>

  <!-- Body -->
  <div style="padding:32px;">

    <table style="width:100%;border-collapse:collapse;">
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;width:36%;">
          <span style="font-size:11px;font-family:monospace;letter-spacing:1px;text-transform:uppercase;color:#9698B8;">Nombre</span>
        </td>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <strong style="font-size:15px;">${escapeHtml(data.nombre)}</strong>
        </td>
      </tr>
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <span style="font-size:11px;font-family:monospace;letter-spacing:1px;text-transform:uppercase;color:#9698B8;">Empresa</span>
        </td>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <span style="font-size:14px;">${escapeHtml(data.empresa)}</span>
        </td>
      </tr>
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <span style="font-size:11px;font-family:monospace;letter-spacing:1px;text-transform:uppercase;color:#9698B8;">Email</span>
        </td>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <a href="mailto:${escapeHtml(data.email)}" style="color:#818CF8;font-size:14px;">${escapeHtml(data.email)}</a>
        </td>
      </tr>
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <span style="font-size:11px;font-family:monospace;letter-spacing:1px;text-transform:uppercase;color:#9698B8;">WhatsApp</span>
        </td>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <a href="https://wa.me/${sanitizePhone(data.telefono)}" style="color:#22D3EE;font-size:14px;">${escapeHtml(data.telefono)}</a>
        </td>
      </tr>
      <tr>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <span style="font-size:11px;font-family:monospace;letter-spacing:1px;text-transform:uppercase;color:#9698B8;">Sitio web</span>
        </td>
        <td style="padding:10px 0;border-bottom:1px solid #1E1C28;">
          <span style="font-size:14px;">${data.sitio && data.sitio !== '—' ? `<a href="${escapeHtml(data.sitio)}" style="color:#818CF8;">${escapeHtml(data.sitio)}</a>` : '<span style="color:#4E5075;">—</span>'}</span>
        </td>
      </tr>
      <tr>
        <td style="padding:10px 0;">
          <span style="font-size:11px;font-family:monospace;letter-spacing:1px;text-transform:uppercase;color:#9698B8;">Necesita</span>
        </td>
        <td style="padding:10px 0;">
          <span style="background:#262432;border-radius:6px;padding:4px 10px;font-size:13px;">${escapeHtml(data.necesidad)}</span>
        </td>
      </tr>
    </table>

    <!-- Mensaje -->
    <div style="margin-top:24px;background:#141318;border:1px solid #1E1C28;border-radius:10px;padding:20px;">
      <p style="margin:0 0 8px;font-size:11px;font-family:monospace;letter-spacing:1px;text-transform:uppercase;color:#9698B8;">Mensaje</p>
      <p style="margin:0;font-size:14px;line-height:1.7;color:#E8EAF4;">${escapeHtml(data.mensaje).replace(/\n/g, '<br>')}</p>
    </div>

    <!-- Actions -->
    <div style="margin-top:24px;display:flex;gap:12px;">
      <a href="mailto:${escapeHtml(data.email)}?subject=Re: Diagnóstico Velora Labs"
         style="display:inline-block;background:#5B5FE8;color:#fff;padding:11px 22px;border-radius:8px;font-size:13px;font-weight:600;text-decoration:none;">
        Responder por email
      </a>
      <a href="https://wa.me/${sanitizePhone(data.telefono)}"
         style="display:inline-block;background:#1E1C28;color:#E8EAF4;padding:11px 22px;border-radius:8px;font-size:13px;font-weight:600;text-decoration:none;border:1px solid #262432;">
        Abrir WhatsApp
      </a>
    </div>

  </div>

  <!-- Footer -->
  <div style="padding:20px 32px;border-top:1px solid #1E1C28;">
    <p style="margin:0;font-size:12px;color:#4E5075;">
      Enviado desde el formulario de diagnóstico de veloralabs.studio · ${data.fecha || ''}
    </p>
  </div>

</div>`;

  GmailApp.sendEmail(NOTIFICATION_EMAIL, subject, '', { htmlBody: html });
}

/* ─── Helpers ──────────────────────────────────────── */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function sanitizePhone(phone) {
  if (!phone) return '';
  return String(phone).replace(/[^0-9+]/g, '');
}
