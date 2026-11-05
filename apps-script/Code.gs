/**
 * Manito – registro de proformas y envío por correo
 * Pega este código en: Hoja de cálculo de Google → Extensiones → Apps Script.
 * Luego: Implementar → Nueva implementación → Aplicación web
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier persona
 * Copia la URL de la aplicación web y pégala en Manito (SCRIPT_URL).
 */

// Debe ser igual al TOKEN que está en index.html de Manito.
const TOKEN = 'manito-2026';
const HOJA = 'Proformas';
const ENCABEZADOS = ['Fecha registro', 'N.º', 'Negocio', 'Servicio', 'Correo usuario', 'WhatsApp usuario',
  'Cliente', 'Correo cliente', 'WhatsApp cliente', 'Lugar', 'Detalle', 'Subtotal', 'IVA', 'Total',
  'Anticipo %', 'Validez (días)', 'Reenvío', 'Correo enviado a'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return json_({ ok: false, error: 'token' });

    const hoja = hoja_();
    const destinos = [];
    let estadoCorreo = '';

    const okMail = s => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim());
    const cliente = okMail(d.cemail) ? d.cemail.trim() : '';
    const usuario = okMail(d.uemail) ? d.uemail.trim() : '';
    const dueno = Session.getEffectiveUser().getEmail();

    if (d.pdf && (cliente || usuario)) {
      const pdf = Utilities.newBlob(Utilities.base64Decode(d.pdf), 'application/pdf', d.filename || 'Proforma.pdf');
      const negocio = d.negocio || 'Tu proveedor';
      const total = '$' + Number(d.total || 0).toFixed(2);
      const para = cliente || usuario;
      const cc = cliente && usuario ? usuario : '';
      const html =
        '<div style="font-family:Arial,sans-serif;font-size:15px;color:#14202C;line-height:1.5">' +
        '<p>Hola ' + esc_(d.cliente || '') + ',</p>' +
        '<p><b>' + esc_(negocio) + '</b> te envía la proforma <b>N.º ' + esc_(d.no) + '</b> por <b>' + total + '</b>.</p>' +
        '<p>La encontrarás adjunta en PDF. Para cualquier consulta, responde a este correo' +
        (d.utel ? ' o escribe al WhatsApp ' + esc_(d.utel) : '') + '.</p>' +
        '<p style="color:#7A8794;font-size:12px;margin-top:24px">Proforma hecha con Manito</p></div>';
      const opciones = { name: negocio + ' vía Manito', htmlBody: html, attachments: [pdf] };
      if (cc) opciones.cc = cc;
      if (usuario) opciones.replyTo = usuario;
      if (dueno && dueno !== para && dueno !== cc) opciones.bcc = dueno;
      MailApp.sendEmail(para, 'Proforma N.º ' + d.no + ' – ' + negocio, 'Adjunto la proforma N.º ' + d.no + ' por ' + total + '.', opciones);
      destinos.push(para); if (cc) destinos.push(cc);
      estadoCorreo = destinos.join(', ');
    }

    const detalle = (d.items || []).map(i => i.q + ' x ' + i.d + ' ($' + Number(i.p).toFixed(2) + ')').join(' | ');
    hoja.appendRow([new Date(), d.no, d.negocio, d.oficio, usuario, d.utel || '', d.cliente, cliente,
      d.ctel || '', d.lugar || '', detalle, d.sub, d.iva, d.total, d.anticipo, d.validez, d.reenvio ? 'Sí' : '', estadoCorreo]);

    return json_({ ok: true, enviados: destinos, restantes: MailApp.getRemainingDailyQuota() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() { return json_({ ok: true, app: 'Manito' }); }

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJA);
  if (!h) { h = ss.insertSheet(HOJA); }
  if (h.getLastRow() === 0) {
    h.appendRow(ENCABEZADOS);
    h.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold').setBackground('#16324F').setFontColor('#FFFFFF');
    h.setFrozenRows(1);
  }
  return h;
}

function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Ejecuta esta función una vez desde el editor para dar los permisos de Gmail y Hojas.
function autorizar() { hoja_(); MailApp.getRemainingDailyQuota(); }
