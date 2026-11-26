// Kniestep Verhuur – reserveringen en bezetting
// Zet in kolom M "geannuleerd" om een reservering vrij te geven.
const CAP = 2;
const MAIL = 'kniestepverhuur@outlook.com';
const TZ = 'Europe/Amsterdam';
const KOP = ['Ontvangen','Naam','Telefoon','E-mail','Start','Terug','Weken','Tijdvak','Adres','Huur','Bezorging','Totaal','Status'];

function blad() {
  const s = SpreadsheetApp.getActive().getSheets()[0];
  if (s.getLastRow() === 0) { s.appendRow(KOP); s.setFrozenRows(1); s.getRange('E:F').setNumberFormat('dd-mm-yyyy'); }
  return s;
}
function dag(x) { return Utilities.formatDate(new Date(x), TZ, 'yyyy-MM-dd'); }
function plus(k, n) { const a = new Date(k + 'T12:00:00Z'); a.setUTCDate(a.getUTCDate() + n); return a.toISOString().slice(0, 10); }

function volleDagen() {
  const s = blad(), n = s.getLastRow(), tel = {};
  if (n < 2) return [];
  s.getRange(2, 1, n - 1, 13).getValues().forEach(r => {
    if (/annul/i.test(String(r[12])) || !r[4] || !r[5]) return;
    for (let k = dag(r[4]), eind = dag(r[5]); k <= eind; k = plus(k, 1)) tel[k] = (tel[k] || 0) + 1;
  });
  const vandaag = dag(new Date());
  return Object.keys(tel).filter(k => tel[k] >= CAP && k >= vandaag).sort();
}

function uit(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() { return uit({ vol: volleDagen() }); }

function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (x) { return uit({ ok: false, fout: 'ongeldig' }); }
  if (p.web) return uit({ ok: true });                       // spam-val
  const w = Math.max(2, Math.min(26, parseInt(p.weken, 10) || 0));
  if (!p.naam || !p.email || !/^\d{4}-\d{2}-\d{2}$/.test(p.start || '') || !p.adres) return uit({ ok: false, fout: 'onvolledig' });
  const eind = plus(p.start, w * 7);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const vol = volleDagen();
    if (vol.some(k => k >= p.start && k <= eind)) return uit({ ok: false, fout: 'vol', vol: vol });
    const s = blad();
    s.appendRow([new Date(), p.naam, "'" + (p.tel || ''), p.email, new Date(p.start + 'T12:00:00'), new Date(eind + 'T12:00:00'), w, p.slot || '', p.adres, p.huur || '', p.bezorg || '', p.totaal || '', '']);
    SpreadsheetApp.flush();
    const body = 'Nieuwe reservering via kniestepverhuur.nl\n\n' +
      'Naam: ' + p.naam + '\nE-mail: ' + p.email + '\nTelefoon: ' + (p.tel || '-') +
      '\nStart: ' + p.start + '\nTerug: ' + eind + ' (' + w + ' weken)\nTijdvak: ' + (p.slot || '') +
      '\nAdres: ' + p.adres + '\nHuur: ' + (p.huur || '') + '\nBezorging: ' + (p.bezorg || '') + '\nTotaal: ' + (p.totaal || '') +
      '\n\nAnnuleert de klant? Zet in de sheet bij Status "geannuleerd".\n' + SpreadsheetApp.getActive().getUrl();
    MailApp.sendEmail({ to: MAIL, replyTo: p.email, subject: 'Reservering kniestep – ' + p.naam + ' – ' + p.start, body: body, name: 'Kniestep Verhuur' });
    return uit({ ok: true, eind: eind });
  } finally { lock.releaseLock(); }
}
