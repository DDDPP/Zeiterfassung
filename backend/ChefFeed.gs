// ═══════════════════════════════════════════════════════════════════════════════
// GARTEN HOCH 3 – Chef-Überblick (Backend-Modul für den Apps-Script-Webhook)
// ═══════════════════════════════════════════════════════════════════════════════
// Liefert für die Seite chef.html einen Feed aus Ereignissen der letzten Tage:
//   Warnungen:  Pause zu kurz · Schicht ohne Ausstempeln · Leistung ohne Ende · sehr langer Arbeitstag
//   Hinweise:   gemeldeter Mehraufwand
// Alles wird beim Abruf direkt aus den Tabs berechnet (kein eigener Ereignis-Tab, keine
// Änderung an bestehenden Skripten). Nutzt Hilfsfunktionen aus ProjektReport.gs (gleiches
// Apps-Script-Projekt): prJson_, prNorm_, prIsoDatum_, prKopfName_, prAusgeschlossen_,
// prLeistungen_, prZeit_, PR_ZEITZONE.
//
// EINRICHTUNG (einmalig)
//  1. Diese Datei als neue Skriptdatei „ChefFeed.gs" im Apps-Script-Projekt anlegen.
//  2. In der bestehenden doGet(e) ganz oben diese Zeile einfügen (neben der projekt_report-Zeile):
//        if (e.parameter.action === 'chef_feed') return chefFeedAntwort_(e);
//  3. Projekteinstellungen (Zahnrad) → Skripteigenschaften → Eigenschaft hinzufügen:
//        CHEF_TOKEN = <langer, zufälliger Code>
//     Das ist der Zugang zur Chef-Seite (?token=… in der URL). Eigener Code, getrennt von
//     Mitarbeiter- und Report-Token.
//  4. Bereitstellen → Bereitstellungen verwalten → Stift → Neue Version (NICHT „Neue Bereitstellung").
//
// ANNAHMEN (nicht an echten Daten geprüft – bitte mit &diagnose=1 an der Chef-URL kontrollieren):
//  A1. Jeder Mitarbeiter-Tab hat in Zeile 55 die Überschriften „Datum | WT | Typ | Beginn | Ende | Pause …"
//      und ab Zeile 56 die Tagesdaten (laut technischer Dokumentation). Tabs ohne diese Kopfzeile
//      werden übersprungen und bei diagnose=1 aufgelistet.
//  A2. Der Tab-Name ist der Mitarbeitername.
//  A3. Uhrzeiten stehen als „HH:MM" in der Anzeige der Zellen; die Pause als Minuten-Zahl oder „H:MM".
//  A4. Die Pausenregel entspricht der Stempeluhr-App: je Schicht über 6 Std. mindestens 30 Min.,
//      über 9 Std. mindestens 45 Min. (Anwesenheitszeit Ende − Beginn).
//  A5. Schwelle „sehr langer Arbeitstag": mehr als 10 Std. netto am Tag (CF_LANGER_TAG_MIN) – Richtwert, keine Messung.
// ═══════════════════════════════════════════════════════════════════════════════

const CF_TAB_MEHRAUFWAND = 'Mehraufwand';
const CF_KOPFZEILE = 55;          // Überschriften der Mitarbeiter-Tabs
const CF_ERSTE_DATENZEILE = 56;   // erste Tageszeile (01.01.2026)
const CF_MAX_TAGE = 60;
const CF_LANGER_TAG_MIN = 600;    // Annahme A5

function chefFeedAntwort_(e) {
  const p = (e && e.parameter) || {};
  try {
    const erwartet = PropertiesService.getScriptProperties().getProperty('CHEF_TOKEN');
    if (!erwartet) return prJson_({ status: 'error', message: 'CHEF_TOKEN ist in den Skripteigenschaften nicht gesetzt.' });
    if (!p.token || p.token !== erwartet) return prJson_({ status: 'error', message: 'Zugang ungültig.' });

    let tage = parseInt(p.tage, 10);
    if (!(tage >= 1)) tage = 7;
    tage = Math.min(tage, CF_MAX_TAGE);
    const heute = Utilities.formatDate(new Date(), PR_ZEITZONE, 'yyyy-MM-dd');
    const von = cfTagePlus_(heute, -(tage - 1));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const diagnose = String(p.diagnose || '') === '1';

    const ereignisse = [];
    const hinweise = [];
    const diag = { mitarbeiterTabs: [], uebersprungeneTabs: [], beispiele: [] };

    cfAbschnitt_('Stempeluhr', hinweise, function () { cfStempeluhr_(ss, von, heute, ereignisse, hinweise, diag); });
    cfAbschnitt_('Leistungserfassung', hinweise, function () { cfOffeneLeistungen_(ss, von, heute, ereignisse); });
    cfAbschnitt_('Mehraufwand', hinweise, function () { cfMehraufwand_(ss, von, heute, ereignisse); });

    ereignisse.sort(function (a, b) {
      if (a.schwere !== b.schwere) return a.schwere === 'warnung' ? -1 : 1;
      if (a.datum !== b.datum) return a.datum < b.datum ? 1 : -1;
      return (a.zeit || '') < (b.zeit || '') ? 1 : (a.zeit || '') > (b.zeit || '') ? -1 : 0;
    });

    const antwort = { status: 'ok', heute: heute, von: von, tage: tage, ereignisse: ereignisse, hinweise: hinweise };
    if (diagnose) antwort.diagnose = diag;
    return prJson_(antwort);
  } catch (err) {
    return prJson_({ status: 'error', message: String((err && err.message) || err) });
  }
}

// Ein Fehler in einem Abschnitt soll nicht den ganzen Feed verhindern
function cfAbschnitt_(name, hinweise, fn) {
  try { fn(); } catch (err) { hinweise.push(name + ' konnte nicht gelesen werden: ' + String((err && err.message) || err)); }
}

// ── STEMPELUHR (Mitarbeiter-Tabs) ───────────────────────────────────────────────
function cfStempeluhr_(ss, von, heute, out, hinweise, diag) {
  const schichten = [[3, 4, 5], [9, 10, 11], [13, 14, 15]];   // Spalten D/E/F, J/K/L, N/O/P (0-basiert)
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (prAusgeschlossen_(name)) return;
    const letzteZeile = sh.getLastRow(), letzteSpalte = sh.getLastColumn();
    let kopf = [];
    if (letzteZeile >= CF_KOPFZEILE && letzteSpalte >= 6) {
      kopf = sh.getRange(CF_KOPFZEILE, 1, 1, Math.min(20, letzteSpalte)).getDisplayValues()[0].map(prKopfName_);
    }
    if (kopf[0] !== 'datum' || kopf[3] !== 'beginn' || kopf[4] !== 'ende' || kopf[5] !== 'pause') {
      diag.uebersprungeneTabs.push(name);
      return;
    }
    diag.mitarbeiterTabs.push(name);

    // Nur die Zeilen im gewünschten Zeitraum lesen (Datum steht in Spalte A als TT.MM.JJJJ)
    const anzahl = letzteZeile - CF_ERSTE_DATENZEILE + 1;
    if (anzahl < 1) return;
    const spalteA = sh.getRange(CF_ERSTE_DATENZEILE, 1, anzahl, 1).getDisplayValues();
    let erste = -1, letzte = -1;
    for (let i = 0; i < spalteA.length; i++) {
      const d = prIsoDatum_(spalteA[i][0]);
      if (d && d >= von && d <= heute) { if (erste < 0) erste = i; letzte = i; }
    }
    if (erste < 0) return;
    const zeilen = sh.getRange(CF_ERSTE_DATENZEILE + erste, 1, letzte - erste + 1, Math.min(20, letzteSpalte)).getDisplayValues();

    zeilen.forEach(function (z) {
      const datum = prIsoDatum_(z[0]);
      if (!datum || datum < von || datum > heute) return;
      if (prKopfName_(z[2]) !== 'arbeitstag') return;

      let nettoTag = 0, anwesendTag = 0, pauseTag = 0, hatSchicht = false;
      schichten.forEach(function (ix, n) {
        if (ix[2] >= z.length) return;
        const b = cfMin_(z[ix[0]]);
        if (b == null) return;
        hatSchicht = true;
        const e = cfMin_(z[ix[1]]);
        const schichtNr = n + 1;
        if (diag.beispiele.length < 6) diag.beispiele.push({ tab: name, datum: z[0], schicht: schichtNr, beginn: z[ix[0]], ende: z[ix[1]], pause: z[ix[2]] });

        if (e == null) {
          // Offene Schicht eines früheren Tages = Ausstempeln vergessen
          if (datum < heute) {
            out.push(cfEreignis_('schicht_offen', 'warnung', datum, '', name, '',
              'Schicht ohne Ausstempeln',
              name + ' hat am ' + cfDatumKurz_(datum) + ' um ' + z[ix[0]] + ' Uhr eingestempelt, aber kein Ende eingetragen.', String(schichtNr)));
          }
          return;
        }
        const brutto = e - b;
        if (brutto <= 0) return;   // über Mitternacht o. Ä.: wie in der App nicht bewertet
        const pauseRoh = cfPause_(z[ix[2]]);
        const pause = pauseRoh == null ? 0 : pauseRoh;
        const pflicht = brutto > 540 ? 45 : (brutto > 360 ? 30 : 0);
        if (pause < pflicht) {
          out.push(cfEreignis_('pause_kurz', 'warnung', datum, '', name, '',
            'Pause zu kurz',
            name + ' war am ' + cfDatumKurz_(datum) + ' ' + cfStdMin_(brutto) + ' Std. da (' + z[ix[0]] + '–' + z[ix[1]] + ') und hat ' +
              (pauseRoh == null ? 'keine Pause' : pause + ' Min. Pause') + ' eingetragen – vorgeschrieben sind ' + pflicht + ' Min.', String(schichtNr)));
        }
        nettoTag += brutto - pause;
        anwesendTag += brutto;
        pauseTag += pause;
      });

      if (hatSchicht && nettoTag > CF_LANGER_TAG_MIN) {
        out.push(cfEreignis_('lange_zeit', 'warnung', datum, '', name, '',
          'Sehr langer Arbeitstag',
          name + ' war am ' + cfDatumKurz_(datum) + ' ' + cfStdMin_(anwesendTag) + ' Std. anwesend, davon ' + pauseTag + ' Min. Pause – ' +
            'Arbeitszeit nach Abzug der Pause: ' + cfStdMin_(nettoTag) + ' Std.', ''));
      }
    });
  });
  if (!diag.mitarbeiterTabs.length) {
    hinweise.push('In der Tabelle wurde kein Mitarbeiter-Tab erkannt (Zeile 55 mit „Datum | … | Beginn | Ende | Pause" erwartet). Mit &diagnose=1 an der URL sieht man die übersprungenen Tabs.');
  }
}

// ── OFFENE LEISTUNGEN (Leistungserfassung) ──────────────────────────────────────
function cfOffeneLeistungen_(ss, von, heute, out) {
  const liste = prLeistungen_(ss, '', von, heute, []);
  liste.forEach(function (l) {
    if (!l.offen) return;
    const tag = l.start.slice(0, 10);
    if (tag >= heute) return;   // heutige laufende Runden gehören später in „Jetzt aktiv"
    out.push(cfEreignis_('leistung_offen', 'warnung', tag, l.start.slice(11), l.mitarbeiter, l.projekt,
      'Leistung ohne Ende',
      l.mitarbeiter + ' hat am ' + cfDatumKurz_(tag) + ' um ' + l.start.slice(11) + ' Uhr „' + (l.kategorie || 'Leistung') + '" gestartet' +
        (l.projekt ? ' (' + l.projekt + ')' : ' (ohne Projekt)') + ' und nicht beendet.', l.start));
  });
}

// ── MEHRAUFWAND (Tab „Mehraufwand": ID | Datum | Mitarbeiter | Projekt | Unterkategorie | Minuten | Begründung | Erfasst am | Tätigkeit) ──
function cfMehraufwand_(ss, von, heute, out) {
  const sh = ss.getSheetByName(CF_TAB_MEHRAUFWAND);
  if (!sh || sh.getLastRow() < 2) return;
  const z = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(9, sh.getMaxColumns())).getDisplayValues();
  z.forEach(function (r) {
    if (!r[0] || prAusgeschlossen_(r[2])) return;
    const datum = prIsoDatum_(r[1]);
    if (!datum || datum < von || datum > heute) return;
    const min = parseFloat(String(r[5]).replace(',', '.')) || 0;
    if (min <= 0) return;
    const erfasst = prZeit_(r[7], r[7], PR_ZEITZONE);   // 'yyyy-MM-dd HH:mm' oder ''
    const projekt = prNorm_(r[3]) + (prNorm_(r[4]) ? ' · ' + prNorm_(r[4]) : '');
    const taet = r.length > 8 ? prNorm_(r[8]) : '';
    out.push(cfEreignis_('mehraufwand', 'hinweis', datum, erfasst ? erfasst.slice(11) : '', prNorm_(r[2]), prNorm_(r[3]),
      'Mehraufwand: ' + cfStdMin_(min) + ' Std.',
      prNorm_(r[2]) + ' meldet auf ' + projekt + (taet ? ' (' + taet + ')' : '') + ': ' + prNorm_(r[6]), String(r[0]).trim()));
  });
}

// ── HILFSFUNKTIONEN (Präfix cf) ─────────────────────────────────────────────────
// Woher ein Ereignis stammt (wird in der Chef-Seite als Herkunft angezeigt)
const CF_QUELLE = { pause_kurz: 'Zeiterfassung', schicht_offen: 'Zeiterfassung', lange_zeit: 'Zeiterfassung',
                    leistung_offen: 'Leistungserfassung', mehraufwand: 'Mehraufwand' };
function cfEreignis_(typ, schwere, datum, zeit, mitarbeiter, projekt, titel, text, schluessel) {
  return { id: [typ, mitarbeiter, datum, schluessel].join('|'), typ: typ, schwere: schwere, quelle: CF_QUELLE[typ] || '', datum: datum, zeit: zeit || '',
           mitarbeiter: mitarbeiter, projekt: projekt || '', titel: titel, text: text };
}
function cfTagePlus_(iso, n) {
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3] + n)).toISOString().slice(0, 10);
}
function cfDatumKurz_(iso) { return iso.slice(8, 10) + '.' + iso.slice(5, 7) + '.'; }
function cfStdMin_(min) { min = Math.round(min); return Math.floor(min / 60) + ':' + prPad_(min % 60); }
// „7:30" → Minuten seit Mitternacht, sonst null
function cfMin_(s) {
  const m = prNorm_(s).match(/^(\d{1,2}):(\d{2})/);
  return m ? (+m[1]) * 60 + (+m[2]) : null;
}
// Pause: „30" → 30 · „0:30" → 30 · 0,0208 (Tagesbruchteil) → 30 · leer → null (Annahme A3)
function cfPause_(s) {
  s = prNorm_(s);
  if (!s) return null;
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return (+m[1]) * 60 + (+m[2]);
  const x = parseFloat(s.replace(',', '.'));
  if (isNaN(x)) return null;
  return Math.round(x > 0 && x < 1 ? x * 1440 : x);
}
