// Kniestep Verhuur – reserveringen, bezetting, mails en agenda
// Kolom M (Status): zet "geannuleerd" om een reservering vrij te geven.
// Eenmalig: voer setup() uit. Optioneel bij Projectinstellingen > Scripteigenschappen:
//   START_ADRES = vertrekadres voor de reistijd, REVIEW_URL = link naar je Google-reviews.
const SHEET_ID = '1RqGoUG2nLT2gCK7lYnsjt0hIzw_XKORaQcFoeZrAraQ';
const CAP = 2;
const MAIL = 'kniestepverhuur@outlook.com';
const TZ = 'Europe/Amsterdam';
const BUNQ = 'https://bunq.me/DRB/';
const SITE = 'https://kniestepverhuur.nl';
const MAX_PER_DAG = 12; // rem tegen misbruik: daarna geen automatische klantmails meer
const KOP = ['Ontvangen', 'Naam', 'Telefoon', 'E-mail', 'Start', 'Terug', 'Weken', 'Tijdvak', 'Adres', 'Huur', 'Bezorging', 'Totaal', 'Status', 'Taal', 'Ophalen', 'Herinnering', 'Review'];
const DH = [52.0808, 4.3248];
const SLOTS = [[/ochtend|morning/i, 9, 12], [/middag|afternoon/i, 12, 17], [/avond|evening/i, 17, 20]];

function prop(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function blad() {
  const s = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (s.getLastRow() === 0) { s.appendRow(KOP); s.setFrozenRows(1); s.getRange('E:F').setNumberFormat('dd-mm-yyyy'); }
  return s;
}
function wachtblad() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName('Wachtlijst');
  if (!s) { s = ss.insertSheet('Wachtlijst'); s.appendRow(['Ontvangen', 'E-mail', 'Gewenste start', 'Weken']); s.setFrozenRows(1); }
  return s;
}
function dag(x) { return Utilities.formatDate(new Date(x), TZ, 'yyyy-MM-dd'); }
function plus(k, n) { const a = new Date(k + 'T12:00:00Z'); a.setUTCDate(a.getUTCDate() + n); return a.toISOString().slice(0, 10); }
function netjes(x, n) { return String(x || '').replace(/[\r\n<>]/g, ' ').replace(/https?:\/\/\S+/gi, '').trim().slice(0, n); }
function mooi(k, en) { return Utilities.formatDate(new Date(k + 'T12:00:00'), TZ, en ? 'd MMMM yyyy' : 'd-M-yyyy'); }

function volleDagen() {
  const s = blad(), n = s.getLastRow(), tel = {};
  if (n < 2) return [];
  s.getRange(2, 1, n - 1, 13).getValues().forEach(r => {
    if (/annul|verlopen/i.test(String(r[12])) || !r[4] || !r[5]) return;
    for (let k = dag(r[4]), eind = dag(r[5]); k <= eind; k = plus(k, 1)) tel[k] = (tel[k] || 0) + 1;
  });
  const vandaag = dag(new Date());
  return Object.keys(tel).filter(k => tel[k] >= CAP && k >= vandaag).sort();
}

// Prijs altijd op de server berekenen; het bedrag uit de browser is niet te vertrouwen.
function huurVan(w) { let h = 0; for (let i = 1; i <= w; i++) h += 25; return h; }
function euro(v) { return '€' + (Number.isInteger(v) ? v : v.toFixed(2).replace('.', ',')); }
function geoZoek(adres) {
  try {
    const u = 'https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?rows=1&fq=type:adres&fl=weergavenaam,centroide_ll,gemeentenaam,provincienaam&q=' + encodeURIComponent(adres);
    const d = JSON.parse(UrlFetchApp.fetch(u, { muteHttpExceptions: true }).getContentText()).response.docs[0];
    if (!d) return null;
    const m = d.centroide_ll.match(/POINT\(([\d.]+) ([\d.]+)\)/);
    return { naam: d.weergavenaam, gem: d.gemeentenaam, prov: d.provincienaam, pt: [parseFloat(m[2]), parseFloat(m[1])] };
  } catch (x) { return null; }
}
function km(a, b) { const r = x => x * Math.PI / 180, R = 6371, dl = r(b[0] - a[0]), dn = r(b[1] - a[1]); const q = Math.sin(dl / 2) ** 2 + Math.cos(r(a[0])) * Math.cos(r(b[0])) * Math.sin(dn / 2) ** 2; return 2 * R * Math.asin(Math.sqrt(q)); }
// Geeft {bz, op} (null = onbekend) of {buiten:true}
function bezorgkosten(adres, w, ophalen) {
  const g = geoZoek(adres);
  if (!g) return { bz: null, op: null };
  if (g.prov !== 'Zuid-Holland') return { buiten: true };
  const dh = g.gem === "'s-Gravenhage";
  const bz = dh ? 15 : Math.round((15 + Math.round(km(DH, g.pt) * 1.3) * 0.35) * 2) / 2;
  const op = !ophalen ? 0 : (dh && w >= 6 ? 0 : bz);
  return { bz, op };
}

function uit(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return uit({ vol: volleDagen() }); }

function vandaagAantal(s) {
  const n = s.getLastRow(); if (n < 2) return 0;
  const v = dag(new Date());
  return s.getRange(2, 1, n - 1, 1).getValues().filter(r => r[0] && dag(r[0]) === v).length;
}

function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (x) { return uit({ ok: false, fout: 'ongeldig' }); }
  if (p.web) return uit({ ok: true }); // spam-val
  const email = netjes(p.email, 120);
  if (!/^\S+@\S+\.\S+$/.test(email)) return uit({ ok: false, fout: 'onvolledig' });
  const startOk = /^\d{4}-\d{2}-\d{2}$/.test(p.start || '');
  const wIn = parseInt(p.weken, 10);
  if (p.type !== 'wachtlijst' && !(wIn >= 2 && wIn <= 26)) return uit({ ok: false, fout: 'weken' });
  const w = Math.max(2, Math.min(26, wIn || 0));

  if (p.type === 'wachtlijst') {
    if (!startOk) return uit({ ok: false, fout: 'onvolledig' });
    const ws = wachtblad();
    if (vandaagAantal(ws) < MAX_PER_DAG) ws.appendRow([new Date(), email, p.start, w]);
    return uit({ ok: true });
  }

  const naam = netjes(p.naam, 60), adres = netjes(p.adres, 160), tel = netjes(p.tel, 30);
  if (!naam || !startOk || !adres) return uit({ ok: false, fout: 'onvolledig' });
  if (p.start < dag(new Date())) return uit({ ok: false, fout: 'onvolledig' });
  const en = p.taal === 'en', ophalen = !!p.ophalen;
  const slot = netjes(p.slot, 40);
  const hu = huurVan(w), k = bezorgkosten(adres, w, ophalen);
  if (k.buiten) return uit({ ok: false, fout: 'buiten' });
  const bekend = k.bz !== null;
  const bedrag = bekend ? hu + k.bz + k.op : 0;
  const huur = euro(hu);
  const bezorg = bekend ? euro(k.bz) + (ophalen ? ' + ophalen ' + (k.op ? euro(k.op) : 'gratis') : ' / zelf terugbrengen') : 'nog te bepalen';
  const totaal = (bekend ? euro(bedrag) : euro(hu) + ' + bezorging') + ' + €40 borg';
  if (bekend && Number(p.bedrag) && Math.abs(Number(p.bedrag) - bedrag) > 0.01) console.warn('Bedrag browser wijkt af', p.bedrag, bedrag, adres);
  const eind = plus(p.start, w * 7);
  const link = bedrag ? BUNQ + bedrag.toFixed(2) + '/' + encodeURIComponent('Kniestep ' + naam + ' ' + p.start) : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let druk;
  try {
    const vol = volleDagen();
    if (vol.some(k => k >= p.start && k <= eind)) return uit({ ok: false, fout: 'vol', vol: vol });
    const s = blad();
    druk = vandaagAantal(s) >= MAX_PER_DAG;
    s.appendRow([new Date(), naam, "'" + tel, email, new Date(p.start + 'T12:00:00'), new Date(eind + 'T12:00:00'), w, slot, adres, huur, bezorg, totaal, '', en ? 'en' : 'nl', ophalen ? 'ja' : 'nee', '', '']);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  const R = { naam, email, tel, start: p.start, eind, w, slot, adres, huur, bezorg, totaal, bedrag, link, en, ophalen };
  try { mailEigenaar(R, druk); } catch (x) { console.error(x); }
  if (!druk) {
    try { mailKlant(R); } catch (x) { console.error(x); }
    try { agenda(R); } catch (x) { console.error(x); }
  }
  return uit({ ok: true, eind: eind, bedrag: bedrag, link: link });
}

function mailEigenaar(R, druk) {
  MailApp.sendEmail({
    to: MAIL, replyTo: R.email, name: 'Kniestep Verhuur',
    subject: 'Reservering kniestep – ' + R.naam + ' – ' + R.start,
    body: 'Nieuwe reservering via kniestepverhuur.nl\n\nNaam: ' + R.naam + '\nE-mail: ' + R.email + '\nTelefoon: ' + (R.tel || '-') +
      '\nStart: ' + R.start + '\nTerug: ' + R.eind + ' (' + R.w + ' weken)\nTijdvak: ' + R.slot + '\nAdres: ' + R.adres +
      '\nHuur: ' + R.huur + '\nBezorging: ' + R.bezorg + '\nTotaal: ' + R.totaal + '\nNa afloop: ' + (R.ophalen ? 'ophalen' : 'klant brengt terug') +
      (R.link ? '\n\nBetaallink die de klant kreeg: ' + R.link + '\nControleer in bunq of de betaling binnen is.' : '\n\nBezorgkosten onbekend: stuur de klant zelf een betaallink.') +
      (druk ? '\n\nLET OP: veel reserveringen vandaag. Deze klant kreeg GEEN automatische bevestiging en er is geen agenda-item gemaakt.' : '') +
      '\n\nAnnuleert de klant? Zet in de sheet bij Status "geannuleerd".\n' + SpreadsheetApp.openById(SHEET_ID).getUrl()
  });
}

function mailKlant(R) {
  const nl = !R.en;
  const regels = nl ? [
    'Hoi ' + R.naam + ',', '',
    'Bedankt voor je reservering. Dit heb ik genoteerd:', '',
    'Periode: ' + mooi(R.start) + ' t/m ' + mooi(R.eind) + ' (' + R.w + ' weken)',
    'Bezorging: ' + R.slot + ', ' + R.adres,
    'Na afloop: ' + (R.ophalen ? 'ik haal de kniestep bij je op' : 'je brengt de kniestep zelf terug'),
    'Huur: ' + R.huur, 'Bezorging/ophalen: ' + R.bezorg, 'Totaal: ' + R.totaal, '',
    R.link ? 'Nog niet betaald? Dat kan hier met iDEAL: ' + R.link : 'Ik reken de bezorgkosten uit en stuur je een betaallink.',
    'Je reservering is definitief zodra de betaling binnen is. De borg van €40 betaal je bij de overdracht.', '',
    'Annuleren kan kosteloos tot 24 uur voor de bezorging. Start je huur binnen 24 uur, dan kun je niet meer annuleren.',
    'Huurvoorwaarden: ' + SITE + '/#voorwaarden', '',
    'Vragen? Beantwoord gewoon deze mail.', '', 'Groet,', 'Kniestep Verhuur'
  ] : [
    'Hi ' + R.naam + ',', '',
    'Thanks for your booking. These are the details:', '',
    'Period: ' + mooi(R.start, 1) + ' to ' + mooi(R.eind, 1) + ' (' + R.w + ' weeks)',
    'Delivery: ' + R.slot + ', ' + R.adres,
    'Afterwards: ' + (R.ophalen ? 'I will collect the knee scooter from you' : 'you return the knee scooter yourself'),
    'Rent: ' + R.huur, 'Delivery/collection: ' + R.bezorg, 'Total: ' + R.totaal, '',
    R.link ? 'Not paid yet? You can pay here with iDEAL: ' + R.link : 'I will calculate the delivery fee and send you a payment link.',
    'Your booking is final once the payment has arrived. The €40 deposit is paid at handover.', '',
    'You can cancel free of charge up to 24 hours before delivery. If your rental starts within 24 hours, cancelling is no longer possible.',
    'Rental terms (in Dutch): ' + SITE + '/#voorwaarden', '',
    'Questions? Just reply to this email.', '', 'Kind regards,', 'Kniestep Verhuur'
  ];
  MailApp.sendEmail({ to: R.email, replyTo: MAIL, name: 'Kniestep Verhuur', subject: nl ? 'Je reservering bij Kniestep Verhuur' : 'Your booking with Kniestep Verhuur', body: regels.join('\n') });
}

function reisMinuten(adres) {
  const van = prop('START_ADRES');
  if (!van) return 30;
  try {
    const d = Maps.newDirectionFinder().setOrigin(van).setDestination(adres + ', Nederland').setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
    return Math.max(10, Math.ceil(d.routes[0].legs[0].duration.value / 60));
  } catch (x) { return 30; }
}
function tijd(k, uur) { return new Date(Utilities.formatDate(new Date(k + 'T12:00:00Z'), TZ, "yyyy-MM-dd'T'") + ('0' + uur).slice(-2) + ':00:00' + Utilities.formatDate(new Date(k + 'T12:00:00Z'), TZ, 'XXX')); }

function agenda(R) {
  const cal = CalendarApp.getDefaultCalendar();
  const sl = SLOTS.find(x => x[0].test(R.slot)) || SLOTS[0];
  const info = 'Klant: ' + R.naam + '\nTel: ' + R.tel + '\nMail: ' + R.email + '\nPeriode: ' + R.start + ' t/m ' + R.eind + '\nTotaal: ' + R.totaal + '\nBorg €40 innen bij overdracht.';
  const reis = reisMinuten(R.adres) + 30; // reistijd + 30 minuten speling
  const b0 = tijd(R.start, sl[1]), b1 = tijd(R.start, sl[2]);
  cal.createEvent('[Zakelijk] Kniestep bezorgen – ' + R.naam, b0, b1, { location: R.adres, description: info });
  cal.createEvent('[Zakelijk] Reistijd kniestep bezorgen', new Date(b0.getTime() - reis * 60000), b0, { description: 'Reistijd incl. 30 min speling naar ' + R.adres });
  const e0 = tijd(R.eind, sl[1]), e1 = tijd(R.eind, sl[2]);
  if (R.ophalen) {
    cal.createEvent('[Zakelijk] Kniestep ophalen – ' + R.naam, e0, e1, { location: R.adres, description: info + '\nBorg teruggeven.' });
    cal.createEvent('[Zakelijk] Reistijd kniestep ophalen', new Date(e0.getTime() - reis * 60000), e0, { description: 'Reistijd incl. 30 min speling naar ' + R.adres });
  } else {
    cal.createEvent('[Zakelijk] Kniestep terug – ' + R.naam + ' brengt terug', e0, e1, { description: info + '\nBorg teruggeven.' });
  }
}

// Draait dagelijks (zie setup): herinnering 3 dagen voor het einde, reviewverzoek de dag na het einde.
function dagelijks() {
  const s = blad(), n = s.getLastRow(); if (n < 2) return;
  const v = dag(new Date()), over3 = plus(v, 3), gisteren = plus(v, -1), review = prop('REVIEW_URL');
  const verl = 'mailto:' + MAIL + '?subject=Kniestep%20verlengen';
  s.getRange(2, 1, n - 1, 17).getValues().forEach((r, i) => {
    if (/annul|verlopen/i.test(String(r[12])) || !r[3] || !r[5]) return;
    const eind = dag(r[5]), en = r[13] === 'en', naam = r[1], rij = i + 2;
    if (eind === over3 && !r[15]) {
      MailApp.sendEmail({
        to: r[3], replyTo: MAIL, name: 'Kniestep Verhuur',
        subject: en ? 'Your knee scooter rental ends in 3 days' : 'Je kniestep-huur loopt over 3 dagen af',
        body: en
          ? 'Hi ' + naam + ',\n\nYour rental ends on ' + mooi(eind, 1) + '. Need it longer? Reply to this email with the number of extra weeks and I will check availability and send a payment link.\n\n' + (r[14] === 'ja' ? 'Otherwise I will collect the knee scooter on that day in the agreed time slot.' : 'Otherwise please return the knee scooter on that day; reply to arrange a time.') + '\n\nKind regards,\nKniestep Verhuur'
          : 'Hoi ' + naam + ',\n\nJe huur loopt af op ' + mooi(eind) + '. Langer nodig? Beantwoord deze mail met het aantal extra weken, dan kijk ik of het kan en stuur ik een betaallink.\n\n' + (r[14] === 'ja' ? 'Anders haal ik de kniestep die dag op in het afgesproken tijdvak.' : 'Anders breng je de kniestep die dag terug; beantwoord deze mail om een tijd af te spreken.') + '\n\nGroet,\nKniestep Verhuur'
      });
      s.getRange(rij, 16).setValue(new Date());
    }
    if (review && eind === gisteren && !r[16]) {
      MailApp.sendEmail({
        to: r[3], replyTo: MAIL, name: 'Kniestep Verhuur',
        subject: en ? 'How was the knee scooter?' : 'Hoe beviel de kniestep?',
        body: en
          ? 'Hi ' + naam + ',\n\nThanks for renting. I hope the knee scooter helped and that your recovery goes well. A short Google review would help me a lot:\n' + review + '\n\nKind regards,\nKniestep Verhuur'
          : 'Hoi ' + naam + ',\n\nBedankt voor het huren. Ik hoop dat de kniestep heeft geholpen en wens je een goed herstel. Met een korte Google-review help je mij enorm:\n' + review + '\n\nGroet,\nKniestep Verhuur'
      });
      s.getRange(rij, 17).setValue(new Date());
    }
  });
}

function setup() {
  blad(); wachtblad();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dagelijks').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dagelijks').timeBased().everyDays(1).atHour(9).inTimezone(TZ).create();
}

// OPTIONEEL (staat uit): niet-betaalde reserveringen automatisch vrijgeven.
// Werkwijze: zet in kolom M (Status) "betaald" zodra de bunq-betaling binnen is. Stel scripteigenschap
// ONBETAALD_VERVAL_UUR in (bijv. 12) en voer setupVrijgeven() eenmalig uit. Reserveringen zonder status
// die ouder zijn dan dat aantal uur krijgen status "verlopen (niet betaald)" en tellen niet meer mee voor de bezetting.
function vrijgevenOnbetaald() {
  const uur = Number(prop('ONBETAALD_VERVAL_UUR')); if (!uur) return;
  const s = blad(), n = s.getLastRow(); if (n < 2) return;
  const grens = Date.now() - uur * 3600e3;
  s.getRange(2, 1, n - 1, 13).getValues().forEach((r, i) => {
    if (!String(r[12]).trim() && r[0] && new Date(r[0]).getTime() < grens) {
      s.getRange(i + 2, 13).setValue('verlopen (niet betaald)');
      MailApp.sendEmail({ to: MAIL, name: 'Kniestep Verhuur', subject: 'Reservering vrijgegeven (niet betaald) – ' + r[1], body: 'De reservering van ' + r[1] + ' (start ' + dag(r[4]) + ') is na ' + uur + ' uur zonder status "betaald" vrijgegeven. Klopt dat niet? Zet de status terug op leeg of "betaald".' });
    }
  });
}
function setupVrijgeven() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'vrijgevenOnbetaald').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('vrijgevenOnbetaald').timeBased().everyHours(1).create();
}
