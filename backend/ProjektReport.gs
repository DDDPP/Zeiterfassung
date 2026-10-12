// ═══════════════════════════════════════════════════════════════════════════════
// GARTEN HOCH 3 – Projektreport (Backend-Modul für den Apps-Script-Webhook)
// ═══════════════════════════════════════════════════════════════════════════════
// Liefert für ein Projekt und einen Zeitraum alle Leistungsrunden (Leistungserfassung)
// und alle Baustellenberichte als JSON an die Seite projektreport.html.
//
// EINRICHTUNG (einmalig)
//  1. Diese Datei als neue Skriptdatei „ProjektReport.gs" im Apps-Script-Projekt anlegen.
//  2. In der bestehenden doGet(e) ganz oben diese Zeile einfügen:
//        if (e.parameter.action === 'projekt_report') return projektReportAntwort_(e);
//  3. Projekteinstellungen (Zahnrad) → Skripteigenschaften → Eigenschaft hinzufügen:
//        REPORT_TOKEN = <langer, zufälliger Code>
//     Dieser Code ist der Zugang zum Report (?token=… in der Report-URL). Er ist bewusst
//     getrennt von den Mitarbeiter-Tokens, weil der Report die Stunden aller zeigt.
//  4. Der Tab mit den Leistungsrunden heißt „Leistungserfassung" (PR_SHEET_LEISTUNG). Bei Umbenennung hier anpassen.
//  5. Bereitstellen → Bereitstellungen verwalten → Stift → Neue Version (NICHT „Neue Bereitstellung").
//
// Die Spalten der Leistungsrunden werden über ihre Überschriften gefunden (Reihenfolge egal).
// Findet das Skript sie nicht, antwortet es mit einer Fehlermeldung, die die gefundenen
// Überschriften nennt – dann nur PR_LE_SPALTEN unten anpassen.
// ═══════════════════════════════════════════════════════════════════════════════

const PR_SHEET_BERICHTE = 'Baustellenberichte';
const PR_SHEET_LEISTUNG = 'Leistungserfassung';

// Uhrzeiten im Report immer in deutscher Zeit ausgeben – unabhängig von der Zeitzone der Tabelle
const PR_ZEITZONE = 'Europe/Berlin';

// Mögliche Überschriften je Feld (klein geschrieben, ohne Sonderzeichen vergleichen)
const PR_LE_SPALTEN = {
  mitarbeiter: ['mitarbeiter', 'name'],
  kategorie:   ['kategorie', 'taetigkeit', 'tätigkeit'],
  projekt:     ['projekt'],
  start:       ['start', 'startzeit', 'beginn'],
  ende:        ['ende', 'endzeit']
};

function projektReportAntwort_(e) {
  const p = (e && e.parameter) || {};
  try {
    const erwartet = PropertiesService.getScriptProperties().getProperty('REPORT_TOKEN');
    if (!erwartet) return prJson_({ status: 'error', message: 'REPORT_TOKEN ist in den Skripteigenschaften nicht gesetzt.' });
    if (!p.token || p.token !== erwartet) return prJson_({ status: 'error', message: 'Zugang ungültig.' });

    const projekt = prNorm_(p.projekt);
    if (!projekt) return prJson_({ status: 'error', message: 'Kein Projekt angegeben.' });
    const von = prIsoDatum_(p.von);   // leer = ohne Untergrenze
    const bis = prIsoDatum_(p.bis);   // leer = ohne Obergrenze

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hinweise = [];
    const leistungen = prLeistungen_(ss, projekt, von, bis, hinweise);
    const berichte = prBerichte_(ss, projekt, von, bis, hinweise);

    return prJson_({ status: 'ok', projekt: p.projekt, von: von, bis: bis, leistungen: leistungen, berichte: berichte, hinweise: hinweise });
  } catch (err) {
    return prJson_({ status: 'error', message: String((err && err.message) || err) });
  }
}

// ── LEISTUNGSRUNDEN ───────────────────────────────────────────────────────────
function prLeistungen_(ss, projekt, von, bis, hinweise) {
  const sh = ss.getSheetByName(PR_SHEET_LEISTUNG);
  if (!sh) throw new Error('Tab „' + PR_SHEET_LEISTUNG + '" nicht gefunden (PR_SHEET_LEISTUNG anpassen).');
  const tz = PR_ZEITZONE;
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  const range = sh.getRange(1, 1, lastRow, lastCol);
  const werte = range.getValues();
  const anzeige = range.getDisplayValues();

  // Kopfzeile in den ersten 5 Zeilen suchen
  let kopf = -1, idx = {};
  for (let r = 0; r < Math.min(5, werte.length) && kopf < 0; r++) {
    const namen = anzeige[r].map(prKopfName_);
    const treffer = {};
    let alle = true;
    Object.keys(PR_LE_SPALTEN).forEach(function (feld) {
      const i = namen.findIndex(function (n) { return PR_LE_SPALTEN[feld].indexOf(n) >= 0; });
      if (i < 0) alle = false; else treffer[feld] = i;
    });
    if (alle) { kopf = r; idx = treffer; }
  }
  if (kopf < 0) {
    throw new Error('Spalten im Tab „' + PR_SHEET_LEISTUNG + '" nicht erkannt. Gefundene Überschriften in Zeile 1: ' +
      anzeige[0].filter(String).join(' | ') + ' – bitte PR_LE_SPALTEN anpassen.');
  }

  const out = [];
  let ohneZeit = 0;
  for (let r = kopf + 1; r < werte.length; r++) {
    if (prNorm_(anzeige[r][idx.projekt]).toLowerCase() !== projekt.toLowerCase()) continue;
    const start = prZeit_(werte[r][idx.start], anzeige[r][idx.start], tz);
    const ende = prZeit_(werte[r][idx.ende], anzeige[r][idx.ende], tz);
    if (!start) { ohneZeit++; continue; }
    const tag = start.slice(0, 10);
    if ((von && tag < von) || (bis && tag > bis)) continue;
    const offen = !ende;
    out.push({
      mitarbeiter: String(anzeige[r][idx.mitarbeiter]).trim(),
      kategorie: String(anzeige[r][idx.kategorie]).trim(),
      start: start,
      ende: ende,
      dauerMin: offen ? 0 : Math.max(0, Math.round((prMs_(ende) - prMs_(start)) / 60000)),
      offen: offen
    });
  }
  if (ohneZeit) hinweise.push(ohneZeit + ' Leistungsrunde(n) ohne lesbare Startzeit wurden übersprungen.');
  out.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
  return out;
}

// ── BAUSTELLENBERICHTE ────────────────────────────────────────────────────────
// Spalten: A Datum, B Projekt, C Erfasst von, D Zeitstempel, E–N Gruppen 1–5 (Std/Was),
//          O Fahrten, P Belege, Q Foto-Links, R Bemerkungen
function prBerichte_(ss, projekt, von, bis, hinweise) {
  const sh = ss.getSheetByName(PR_SHEET_BERICHTE);
  if (!sh) throw new Error('Tab „' + PR_SHEET_BERICHTE + '" nicht gefunden.');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const range = sh.getRange(2, 1, lastRow - 1, 18);
  const anzeige = range.getDisplayValues();
  const formeln = range.getFormulas();

  const out = [];
  for (let r = 0; r < anzeige.length; r++) {
    const z = anzeige[r];
    if (prNorm_(z[1]).toLowerCase() !== projekt.toLowerCase()) continue;
    const datum = prIsoDatum_(z[0]);
    if (!datum) continue;
    if ((von && datum < von) || (bis && datum > bis)) continue;

    const gruppen = {};
    for (let g = 0; g < 5; g++) {
      const std = parseFloat(String(z[4 + 2 * g]).replace(',', '.')) || 0;
      const was = String(z[5 + 2 * g]).trim();
      if (std || was) gruppen['g' + (g + 1)] = { stunden: std, was: was };
    }
    const fotos = (String(z[16]) + ' ' + String(formeln[r][16] || '')).match(/https?:\/\/[^\s,;"')]+/g) || [];
    out.push({
      datum: datum,
      erfasstVon: String(z[2]).trim(),
      zeitstempel: String(z[3]).trim(),
      gruppen: gruppen,
      fahrten: String(z[14]).trim(),
      belege: String(z[15]).trim(),
      fotos: fotos.filter(function (u, i) { return fotos.indexOf(u) === i; }),
      notizen: String(z[17]).trim()
    });
  }
  out.sort(function (a, b) { return a.datum < b.datum ? -1 : a.datum > b.datum ? 1 : (a.zeitstempel < b.zeitstempel ? -1 : 1); });
  return out;
}

// ── HILFSFUNKTIONEN (alle mit Präfix pr, damit nichts mit bestehendem Code kollidiert) ──
function prJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function prNorm_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function prKopfName_(s) { return prNorm_(s).toLowerCase().replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/[^a-z0-9]/g, ''); }
function prPad_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

// 'yyyy-mm-dd' oder 'dd.mm.yyyy' → 'yyyy-mm-dd' (sonst '')
function prIsoDatum_(s) {
  s = prNorm_(s);
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) return m[3] + '-' + prPad_(m[2]) + '-' + prPad_(m[1]);
  return '';
}

// Zelle (Date-Objekt oder Text) → 'yyyy-MM-dd HH:mm' in der Zeitzone der Tabelle (sonst '')
function prZeit_(wert, anzeige, tz) {
  if (Object.prototype.toString.call(wert) === '[object Date]' && !isNaN(wert.getTime())) return Utilities.formatDate(wert, tz, 'yyyy-MM-dd HH:mm');
  const s = prNorm_(wert) || prNorm_(anzeige);
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}T.*(Z|[+-]\d{2}:?\d{2})$/.test(s)) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm');
  }
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3] + ' ' + prPad_(m[4]) + ':' + m[5];
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})[ ,]+(\d{1,2}):(\d{2})/);
  if (m) return m[3] + '-' + prPad_(m[2]) + '-' + prPad_(m[1]) + ' ' + prPad_(m[4]) + ':' + m[5];
  return '';
}
// 'yyyy-MM-dd HH:mm' → Millisekunden (nur für Differenzen)
function prMs_(z) {
  const m = z.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/);
  return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
}
