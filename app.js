/* =====================================================================
   app.js — Archivio dei verbali del Consiglio d'Istituto
   Indicizzazione locale (DOCX / PDF / TXT) e ricerca per parole o frase.
   ===================================================================== */
(function () {
  'use strict';

  var P = window.VerbaliParser;
  var VERSIONE_INDICE = 6;

  var stato = {
    config: null,
    files: [],           // elenco dal server
    indice: null,        // { versione, generato, docs: {nome: docIndicizzato} }
    normCache: {},       // nome -> array di paragrafi normalizzati
    cestino: [],
    cartellaCestino: null,
    risultati: [],
    ultimaRicerca: null
  };

  /* ------------------------------------------------------------- utilita' */

  function el(id) { return document.getElementById(id); }
  function esc(s) { return P.escHtml(s == null ? '' : s); }

  function chiediJson(url, opzioni) {
    return fetch(url, opzioni || {}).then(function (r) {
      return r.text().then(function (t) {
        var dati = null;
        try { dati = t ? JSON.parse(t) : null; } catch (e) { dati = null; }
        if (!r.ok) {
          var msg = (dati && dati.errore) ? dati.errore : ('HTTP ' + r.status);
          var err = new Error(msg);
          err.dati = dati; err.status = r.status;
          throw err;
        }
        return dati;
      });
    });
  }

  function velo(mostra, titolo, testo, percentuale) {
    var v = el('velo');
    if (!mostra) { v.hidden = true; return; }
    v.hidden = false;
    if (titolo != null) el('velo-titolo').textContent = titolo;
    if (testo != null) el('velo-testo').textContent = testo;
    if (percentuale != null) el('velo-barra').style.width = Math.round(percentuale) + '%';
  }

  function attendi(ms) { return new Promise(function (r) { setTimeout(r, ms || 0); }); }

  function dimensioneLeggibile(byte) {
    var b = byte || 0;
    if (b >= 1048576) return (b / 1048576).toFixed(1).replace('.', ',') + ' MB';
    return Math.max(1, Math.round(b / 1024)) + ' KB';
  }

  function etichettaVerbale(d) {
    var s = 'Verbale';
    if (d.numero != null) s += ' n. ' + d.numero;
    if (d.annoScolastico) s += '/' + d.annoScolastico;
    return s;
  }

  function ordinaDocs(docs, criterio) {
    var arr = docs.slice();
    arr.sort(function (a, b) {
      if (criterio === 'pertinenza') return (b._occ || 0) - (a._occ || 0);
      var ka = a.dataISO || '', kb = b.dataISO || '';
      if (ka !== kb) return criterio === 'cronologico' ? (ka < kb ? -1 : 1) : (ka < kb ? 1 : -1);
      var na = a.numero == null ? -1 : a.numero, nb = b.numero == null ? -1 : b.numero;
      return criterio === 'cronologico' ? na - nb : nb - na;
    });
    return arr;
  }

  /* ------------------------------------------------ estrazione del testo */

  function estraiDocx(buffer) {
    return JSZip.loadAsync(buffer).then(function (zip) {
      var f = zip.file('word/document.xml');
      if (!f) throw new Error('il file non contiene word/document.xml');
      return f.async('string');
    }).then(function (xml) {
      return {
        paragrafi: P.docxToParagraphs(xml),
        presenze: P.estraiPresenze(P.docxToTabelle(xml))
      };
    });
  }

  function estraiPdf(buffer) {
    return pdfjsLib.getDocument({ data: buffer, isEvalSupported: false }).promise.then(function (pdf) {
      var pagine = [];
      var seq = Promise.resolve();
      for (var i = 1; i <= pdf.numPages; i++) {
        seq = seq.then(function (n) {
          return function () {
            return pdf.getPage(n).then(function (page) {
              return page.getTextContent().then(function (tc) {
                pagine.push(tc.items.map(function (it) {
                  return { str: it.str, x: it.transform[4], y: it.transform[5], w: it.width || 0, h: it.height || 0 };
                }));
              });
            });
          };
        }(i));
      }
      return seq.then(function () {
        return {
          paragrafi: P.pdfItemsToParagraphs(pagine),
          presenze: P.pdfItemsToPresenze(pagine)
        };
      });
    });
  }

  function estraiTxt(buffer) {
    var t = new TextDecoder('utf-8').decode(new Uint8Array(buffer));
    return Promise.resolve({
      paragrafi: t.split(/\r?\n/).map(function (r) { return { text: r.trim(), ilvl: -1, numId: -1 }; }),
      presenze: { rilevata: false, righe: [] }
    });
  }

  function estraiParagrafi(nome, buffer) {
    var e = (nome.split('.').pop() || '').toLowerCase();
    if (e === 'docx') return estraiDocx(buffer);
    if (e === 'pdf') return estraiPdf(buffer);
    if (e === 'txt') return estraiTxt(buffer);
    return Promise.reject(new Error('formato non indicizzabile (' + e + ')'));
  }

  /* ------------------------------------------------------ indicizzazione */

  function indicizzaFile(meta) {
    return fetch('/api/file?nome=' + encodeURIComponent(meta.nome))
      .then(function (r) { if (!r.ok) throw new Error('lettura non riuscita'); return r.arrayBuffer(); })
      .then(function (buf) { return estraiParagrafi(meta.nome, buf); })
      .then(function (estratto) {
        var paras = estratto.paragrafi || estratto;
        var d = P.parseVerbale(paras, meta.nome, estratto.presenze);
        d.nome = meta.nome;
        d.dimensione = meta.dimensione;
        d.modificato = meta.modificato;
        d.errore = null;
        if (d.caratteri < 400) {
          d.avviso = 'testo quasi assente: se si tratta di una scansione va prima applicato un riconoscimento OCR';
        }
        return d;
      })
      .catch(function (e) {
        return {
          nome: meta.nome, dimensione: meta.dimensione, modificato: meta.modificato,
          errore: e.message || String(e), numero: null, annoScolastico: null,
          dataISO: null, dataLabel: null, odg: [], sezioni: [], puntiNonLocalizzati: [],
          presenze: { rilevata: false, righe: [] },
          delibere: [], paragrafi: [], caratteri: 0
        };
      });
  }

  function daRiindicizzare(meta, forza) {
    if (forza) return true;
    if (!stato.indice || !stato.indice.docs) return true;
    var vecchio = stato.indice.docs[meta.nome];
    if (!vecchio) return true;
    if (vecchio.dimensione !== meta.dimensione) return true;
    if (vecchio.modificato !== meta.modificato) return true;
    return false;
  }

  function costruisciIndice(forza) {
    return chiediJson('/api/list').then(function (r) {
      stato.files = (r && r.files) ? r.files : [];
      if (!stato.indice || stato.indice.versione !== VERSIONE_INDICE || forza) {
        stato.indice = { versione: VERSIONE_INDICE, generato: null, docs: {} };
      }

      var daFare = stato.files.filter(function (f) { return daRiindicizzare(f, forza); });
      var presenti = {};
      stato.files.forEach(function (f) { presenti[f.nome] = true; });
      // rimuove dall'indice i file non piu' presenti nella cartella
      Object.keys(stato.indice.docs).forEach(function (n) {
        if (!presenti[n]) { delete stato.indice.docs[n]; delete stato.normCache[n]; }
      });

      if (!daFare.length) { aggiornaStato(); return Promise.resolve(0); }

      velo(true, 'Indicizzazione dei verbali', '', 0);
      var fatti = 0;
      var seq = Promise.resolve();
      daFare.forEach(function (meta) {
        seq = seq.then(function () {
          velo(true, 'Indicizzazione dei verbali', meta.nome, (fatti / daFare.length) * 100);
          return attendi(10).then(function () { return indicizzaFile(meta); }).then(function (d) {
            stato.indice.docs[meta.nome] = d;
            delete stato.normCache[meta.nome];
            fatti++;
            velo(true, 'Indicizzazione dei verbali', meta.nome, (fatti / daFare.length) * 100);
          });
        });
      });

      return seq.then(function () {
        stato.indice.generato = new Date().toISOString();
        return salvaIndice();
      }).then(function () {
        velo(false);
        return daFare.length;
      });
    });
  }

  function salvaIndice() {
    return chiediJson('/api/indice', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(stato.indice)
    }).catch(function (e) {
      console.warn('indice non salvato:', e.message);
    });
  }

  function docs() {
    if (!stato.indice || !stato.indice.docs) return [];
    return Object.keys(stato.indice.docs).map(function (k) { return stato.indice.docs[k]; });
  }

  function paragrafiNormalizzati(d) {
    if (!stato.normCache[d.nome]) {
      stato.normCache[d.nome] = d.paragrafi.map(function (p) { return P.norm(p.t); });
    }
    return stato.normCache[d.nome];
  }

  /* -------------------------------------------------------------- ricerca */

  /* ==================================================================
     PRESENZE DEI COMPONENTI E ALLARMI PER ASSENZE CONSECUTIVE
     ================================================================== */

  var VERSIONE_PRESENZE = 1;

  function chiaveNome(n) { return P.norm(n).replace(/[^a-z0-9]+/g, ' ').trim(); }

  /* Nominativi esclusi dal monitoraggio delle presenze fin dalla prima
     installazione: componenti decaduti dalla carica in epoca anteriore
     all'archivio, dei quali non si vuole traccia fra gli avvisi. Il
     riconoscimento avviene per cognome, ignorando accenti e maiuscole, poiché
     la grafia con cui compaiono nei verbali non è nota a priori. L'elenco
     resta visibile e modificabile nella scheda Presenze. */
  /* Nella versione pubblicata su GitHub Pages il codice è pubblico: i
     nominativi esclusi risiedono soltanto in dati/presenze.json, nel
     repository privato, dove erano già stati registrati. */
  var ESCLUSI_PREDEFINITI = [];

  function statoPresenze() {
    if (!stato.presenzeDati) {
      stato.presenzeDati = { versione: VERSIONE_PRESENZE };
    }
    var s = stato.presenzeDati;
    if (!s.correzioni) s.correzioni = {};
    if (!s.decadenze) s.decadenze = {};
    if (!s.silenziati) s.silenziati = {};
    if (!s.cessati) s.cessati = {};
    if (!s.esclusi) s.esclusi = [];
    if (!s.esclusiInizializzati) {
      ESCLUSI_PREDEFINITI.forEach(function (e) {
        var gia = s.esclusi.some(function (x) { return x.pattern === e.pattern; });
        if (!gia) s.esclusi.push({ pattern: e.pattern, etichetta: e.etichetta, predefinito: true });
      });
      s.esclusiInizializzati = true;
    }
    return s;
  }

  /* Vero se il nominativo corrisponde a uno degli esclusi. Il confronto è per
     parole intere consecutive, così un cognome composto non intercetta un cognome più lungo che lo contiene. */
  function nominativoEscluso(nome) {
    var parole = chiaveNome(nome).split(' ').filter(Boolean);
    if (!parole.length) return false;
    return statoPresenze().esclusi.some(function (e) {
      var p = String(e.pattern || '').split(' ').filter(Boolean);
      if (!p.length) return false;
      for (var i = 0; i + p.length <= parole.length; i++) {
        var ok = true;
        for (var j = 0; j < p.length; j++) {
          if (parole[i + j] !== p[j]) { ok = false; break; }
        }
        if (ok) return true;
      }
      return false;
    });
  }

  function salvaPresenze() {
    return chiediJson('/api/presenze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(statoPresenze())
    }).catch(function (e) { console.warn('presenze non salvate:', e.message); });
  }

  /* Sedute in ordine cronologico, dalla più antica alla più recente. Le
     assenze consecutive hanno senso soltanto su una successione ordinata. */
  function seduteOrdinate() {
    return docs().filter(function (d) { return !d.errore; })
      .sort(function (a, b) {
        var ka = a.dataISO || '', kb = b.dataISO || '';
        if (ka !== kb) return ka < kb ? -1 : 1;
        return (a.numero == null ? 0 : a.numero) - (b.numero == null ? 0 : b.numero);
      });
  }

  /* Stato di un componente in una seduta, tenuto conto delle correzioni
     manuali: 'presente', 'assente', oppure null se non risulta in carica. */
  function statoIn(doc, chiave) {
    /* La cessazione dal mandato ha una decorrenza: le sedute successive non
       riguardano più il componente, quelle anteriori conservano il dato che i
       verbali documentano. */
    var cess = statoPresenze().cessati[chiave];
    if (cess && cess.data && doc.dataISO && doc.dataISO >= cess.data) return null;

    var corr = statoPresenze().correzioni[doc.nome + '||' + chiave];
    if (corr) return corr === 'assente' ? 'assente' : (corr === 'presente' ? 'presente' : null);
    var righe = (doc.presenze && doc.presenze.righe) ? doc.presenze.righe : [];
    for (var i = 0; i < righe.length; i++) {
      if (chiaveNome(righe[i].nome) === chiave) {
        return righe[i].stato === 'assente' ? 'assente' : (righe[i].stato === 'presente' ? 'presente' : null);
      }
    }
    return null;
  }

  /* Prospetto completo: per ciascun componente la successione degli stati,
     la serie di assenze consecutive in corso e il livello di allarme. */
  function calcolaPresenze() {
    var sedute = seduteOrdinate();
    var conTabella = sedute.filter(function (d) { return d.presenze && d.presenze.rilevata; });

    // elenco dei componenti: unione dei nominativi rilevati, più quelli
    // introdotti dalle correzioni manuali
    var membri = {};
    function registra(nome) {
      var k = chiaveNome(nome);
      if (!k) return null;
      // gli esclusi non vengono nemmeno registrati: non compaiono in alcun
      // punto del monitoraggio, né fra gli avvisi né nel prospetto
      if (nominativoEscluso(nome)) return null;
      if (!membri[k]) membri[k] = { chiave: k, nome: nome, ruolo: '', stati: {}, note: {} };
      return membri[k];
    }
    sedute.forEach(function (d) {
      ((d.presenze && d.presenze.righe) || []).forEach(function (r) {
        var m = registra(r.nome);
        if (m && r.ruolo && !m.ruolo) m.ruolo = r.ruolo;
        if (m && r.nota) m.note[d.nome] = r.nota;
      });
    });
    Object.keys(statoPresenze().correzioni).forEach(function (k) {
      var p = k.split('||');
      if (p.length === 2 && nominativoEscluso(p[1])) return;
      if (p.length === 2 && membri[p[1]] === undefined) {
        // una correzione può riguardare un componente non rilevato in alcuna
        // tabella: se ne conserva il nominativo così come memorizzato
        membri[p[1]] = { chiave: p[1], nome: p[1], ruolo: '', stati: {}, note: {} };
      }
    });

    var ultima = conTabella.length ? conTabella[conTabella.length - 1] : null;

    var elenco = Object.keys(membri).map(function (k) { return membri[k]; });
    elenco.forEach(function (m) {
      var serie = 0, serieMax = 0, ultimaAssenza = null, totAssenze = 0, totPresenze = 0;
      var primaDellaSerie = null, sedutaRosso = null, serieRossa = [];
      var correnti = [];
      sedute.forEach(function (d) {
        var st = statoIn(d, m.chiave);
        m.stati[d.nome] = st;
        if (st === 'assente') {
          if (serie === 0) { primaDellaSerie = d; correnti = []; }
          serie++;
          correnti.push(d);
          totAssenze++;
          ultimaAssenza = d;
          if (serie > serieMax) serieMax = serie;
          if (serie === 3 && !sedutaRosso) { sedutaRosso = d; serieRossa = correnti.slice(); }
        } else if (st === 'presente') {
          serie = 0;
          primaDellaSerie = null;
          correnti = [];
          totPresenze++;
        }
      });
      m.serie = serie;
      m.serieMax = serieMax;
      m.totAssenze = totAssenze;
      m.totPresenze = totPresenze;
      m.ultimaAssenza = ultimaAssenza;
      m.primaDellaSerie = primaDellaSerie;
      m.sedutaRosso = sedutaRosso;
      m.serieRossa = serieRossa;
      m.cessato = statoPresenze().cessati[m.chiave] || null;
      m.inCarica = m.cessato ? false : (ultima ? statoIn(ultima, m.chiave) !== null : false);
      m.decadenza = statoPresenze().decadenze[m.chiave] || null;
      m.silenziato = !!statoPresenze().silenziati[m.chiave];

      /* L'avviso arancione cessa da sé non appena il componente risulta
         presente in una seduta successiva, poiché la serie riparte da zero.
         L'avviso rosso, una volta raggiunte tre assenze consecutive, permane
         anche se il componente torna a presentarsi: si chiude soltanto con la
         registrazione della delibera di decadenza, oppure si tace a mano. */
      if (m.cessato) m.livello = 'cessato';
      else if (m.decadenza) m.livello = 'chiuso';
      else if (serieMax > 2) m.livello = 'rosso';
      else if (serie === 2) m.livello = 'arancione';
      else m.livello = 'nessuno';

      // lampeggia in apertura soltanto se il componente risulta ancora in
      // carica e l'avviso non è stato silenziato a mano
      m.lampeggia = (m.livello === 'rosso' || m.livello === 'arancione') &&
                    m.inCarica && !m.silenziato;
    });

    elenco.sort(function (a, b) {
      var ord = { rosso: 0, arancione: 1, nessuno: 2, chiuso: 3, cessato: 4 };
      if (ord[a.livello] !== ord[b.livello]) return ord[a.livello] - ord[b.livello];
      if (b.serie !== a.serie) return b.serie - a.serie;
      return a.nome.localeCompare(b.nome, 'it');
    });

    return {
      sedute: sedute,
      conTabella: conTabella,
      senzaTabella: sedute.filter(function (d) { return !(d.presenze && d.presenze.rilevata); }),
      membri: elenco,
      attuali: elenco.filter(function (m) { return !m.cessato; }),
      cessati: elenco.filter(function (m) { return !!m.cessato; }),
      ultima: ultima,
      allarmi: elenco.filter(function (m) { return m.lampeggia; }),
      sospesi: elenco.filter(function (m) {
        return (m.livello === 'rosso' || m.livello === 'arancione') && !m.lampeggia;
      })
    };
  }

  /* ------------------------------------------------ ricerca per numero */

  /* Riconosce nel testo digitato i riferimenti a un verbale o a una delibera.
     Sono ammesse le forme "10", "verbale 10", "v10", "delibera 96",
     "d 96/2025-26", anche piu' d'una nella medesima interrogazione. Un numero
     privo di qualificazione viene cercato come verbale e come delibera. */
  function analizzaNumero(qRaw) {
    var n = P.norm(qRaw).replace(/\bn[.°*]?\s*(?=\d)/g, '').replace(/\s+/g, ' ').trim();
    if (!n) return [];

    var VERB = /^(verbale|verbali|verb|vrb|v)$/;
    var DELIB = /^(delibera|delibere|delib|del|d)$/;
    var NUM = /^(\d{1,4})(?:\/(\d{4}[-\/]\d{2,4}))?$/;
    var NUM_ATTACCATO = /^(verbale|verb|v|delibera|delib|del|d)\.?(\d{1,4})(?:\/(\d{4}[-\/]\d{2,4}))?$/;

    var tok = n.split(' ');
    var richieste = [], i = 0;

    function aggiungi(tipo, numero, anno) {
      if (!numero) return;
      richieste.push({ tipo: tipo, numero: parseInt(numero, 10), anno: anno || null });
    }

    while (i < tok.length) {
      var t = tok[i];
      var m;

      if (VERB.test(t) || DELIB.test(t)) {
        var tipo = VERB.test(t) ? 'verbale' : 'delibera';
        m = NUM.exec(tok[i + 1] || '');
        if (m) { aggiungi(tipo, m[1], m[2]); i += 2; continue; }
        i++; continue;
      }

      m = NUM_ATTACCATO.exec(t);
      if (m) {
        aggiungi(/^(v|vrb|verb|verbale)$/.test(m[1]) ? 'verbale' : 'delibera', m[2], m[3]);
        i++; continue;
      }

      m = NUM.exec(t);
      if (m) { aggiungi('entrambi', m[1], m[2]); i++; continue; }

      i++;
    }
    return richieste;
  }

  function stessoAnno(a, b) {
    if (!a || !b) return true;
    return P.norm(a).replace(/[^0-9]/g, '') === P.norm(b).replace(/[^0-9]/g, '');
  }

  function cercaPerNumero(q) {
    var richieste = analizzaNumero(q);
    var esiti = { verbali: [], delibere: [], richieste: richieste };
    if (!richieste.length) return esiti;

    var visti = {};
    richieste.forEach(function (r) {
      if (r.tipo === 'verbale' || r.tipo === 'entrambi') {
        docs().forEach(function (d) {
          if (d.errore || d.numero !== r.numero) return;
          if (!stessoAnno(r.anno, d.annoScolastico)) return;
          if (visti['v' + d.nome]) return;
          visti['v' + d.nome] = true;
          esiti.verbali.push(d);
        });
      }
      if (r.tipo === 'delibera' || r.tipo === 'entrambi') {
        docs().forEach(function (d) {
          if (d.errore) return;
          d.sezioni.forEach(function (s, si) {
            (s.dettaglioDelibere || []).forEach(function (del) {
              if (del.numero !== r.numero) return;
              if (!stessoAnno(r.anno, del.anno)) return;
              var chiave = 'd' + d.nome + '#' + si + '#' + del.id;
              if (visti[chiave]) return;
              visti[chiave] = true;
              esiti.delibere.push({ doc: d, sezione: s, sezioneIdx: si, delibera: del });
            });
          });
        });
      }
    });

    esiti.verbali = ordinaDocs(esiti.verbali, 'recenti');
    esiti.delibere.sort(function (a, b) {
      var ka = a.doc.dataISO || '', kb = b.doc.dataISO || '';
      if (ka !== kb) return ka < kb ? 1 : -1;
      return b.delibera.numero - a.delibera.numero;
    });
    return esiti;
  }

  /* -------------------------------------------------- ricerca per data */

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function giorniDelMese(anno, mese) { return new Date(anno, mese, 0).getDate(); }

  /* Ogni espressione di data viene convertita nell'intervallo che essa
     designa: un giorno singolo, un mese intero, un anno solare o un anno
     scolastico, che si assume compreso fra il 1 settembre e il 31 agosto. */
  function intervalloDa(espr) {
    var a = espr.anno, m = espr.mese, g = espr.giorno;
    if (espr.tipo === 'as') {
      return { da: a + '-09-01', a: (a + 1) + '-08-31' };
    }
    if (g) return { da: a + '-' + pad2(m) + '-' + pad2(g), a: a + '-' + pad2(m) + '-' + pad2(g) };
    if (m) return { da: a + '-' + pad2(m) + '-01', a: a + '-' + pad2(m) + '-' + pad2(giorniDelMese(a, m)) };
    return { da: a + '-01-01', a: a + '-12-31' };
  }

  function analizzaData(qRaw) {
    var n = P.norm(qRaw)
      .replace(/\ba\.?\s*s\.?\s*/g, 'as ')
      .replace(/\s+/g, ' ').trim();
    if (!n) return null;

    var MESI = P.mesi;
    var espressioni = [];

    // anno scolastico: "as 2025/26", "2025/26", "2025-26"
    var reAS = /(?:as\s*)?(\d{4})\s*[\/-]\s*(\d{2,4})(?![\d\/-])/g, m;
    var consumato = n;
    while ((m = reAS.exec(n)) !== null) {
      var a1 = parseInt(m[1], 10);
      var secondo = m[2].length === 2 ? (Math.floor(a1 / 100) * 100 + parseInt(m[2], 10)) : parseInt(m[2], 10);
      if (secondo === a1 + 1 && a1 >= 1990 && a1 <= 2100) {
        espressioni.push({ tipo: 'as', anno: a1, testo: m[0] });
        consumato = consumato.replace(m[0], ' ');
      }
    }

    // giorno numerico: 23/04/2026, 23-4-2026, 23.04.2026
    var reGG = /(\d{1,2})\s*[\/.-]\s*(\d{1,2})\s*[\/.-]\s*(\d{4})/g;
    while ((m = reGG.exec(consumato)) !== null) {
      var gg = +m[1], mm = +m[2], aa = +m[3];
      if (gg >= 1 && gg <= 31 && mm >= 1 && mm <= 12) {
        espressioni.push({ tipo: 'giorno', giorno: gg, mese: mm, anno: aa, testo: m[0] });
      }
    }
    consumato = consumato.replace(reGG, ' ');

    // giorno con mese in lettere: "23 aprile 2026"
    var reGGM = new RegExp('(\\d{1,2})\\s+(' + MESI.join('|') + ')\\s+(\\d{4})', 'g');
    while ((m = reGGM.exec(consumato)) !== null) {
      espressioni.push({ tipo: 'giorno', giorno: +m[1], mese: MESI.indexOf(m[2]) + 1, anno: +m[3], testo: m[0] });
    }
    consumato = consumato.replace(reGGM, ' ');

    // mese in lettere con anno: "aprile 2026"
    var reM = new RegExp('(' + MESI.join('|') + ')\\s+(\\d{4})', 'g');
    while ((m = reM.exec(consumato)) !== null) {
      espressioni.push({ tipo: 'mese', mese: MESI.indexOf(m[1]) + 1, anno: +m[2], testo: m[0] });
    }
    consumato = consumato.replace(reM, ' ');

    // mese numerico con anno: "04/2026"
    var reMN = /(\d{1,2})\s*[\/.-]\s*(\d{4})/g;
    while ((m = reMN.exec(consumato)) !== null) {
      if (+m[1] >= 1 && +m[1] <= 12) espressioni.push({ tipo: 'mese', mese: +m[1], anno: +m[2], testo: m[0] });
    }
    consumato = consumato.replace(reMN, ' ');

    // anno solare isolato
    var reA = /(?:^|\s)(\d{4})(?:\s|$)/g;
    while ((m = reA.exec(consumato)) !== null) {
      var y = +m[1];
      if (y >= 1990 && y <= 2100) espressioni.push({ tipo: 'anno', anno: y, testo: m[1] });
    }

    if (!espressioni.length) return null;

    var intervalli = espressioni.map(intervalloDa);
    // "dal ... al ...": due estremi si fondono in un unico intervallo
    var haCongiunzione = /\b(dal|da|al|a|fra|tra)\b/.test(n) || /\d\s*-\s*\d/.test(n);
    if (intervalli.length === 2 && haCongiunzione) {
      var da = intervalli[0].da < intervalli[1].da ? intervalli[0].da : intervalli[1].da;
      var aFin = intervalli[0].a > intervalli[1].a ? intervalli[0].a : intervalli[1].a;
      intervalli = [{ da: da, a: aFin }];
      espressioni = [{ tipo: 'intervallo' }];
    }
    return { espressioni: espressioni, intervalli: intervalli, descrizione: descriviIntervalli(intervalli) };
  }

  function itData(iso) {
    if (!iso) return '—';
    var p = iso.split('-');
    return p[2] + '/' + p[1] + '/' + p[0];
  }

  function descriviIntervalli(intervalli) {
    return intervalli.map(function (iv) {
      if (iv.da === iv.a) return 'del ' + itData(iv.da);
      return 'dal ' + itData(iv.da) + ' al ' + itData(iv.a);
    }).join(' e ');
  }

  function cercaPerData(q) {
    var an = analizzaData(q);
    if (!an) return { verbali: [], analisi: null, senzaData: 0 };
    var trovati = [], visti = {}, senzaData = 0;
    docs().forEach(function (d) {
      if (d.errore) return;
      if (!d.dataISO) { senzaData++; return; }
      var dentro = an.intervalli.some(function (iv) { return d.dataISO >= iv.da && d.dataISO <= iv.a; });
      if (dentro && !visti[d.nome]) { visti[d.nome] = true; trovati.push(d); }
    });
    return { verbali: ordinaDocs(trovati, 'recenti'), analisi: an, senzaData: senzaData };
  }

  function disegnaEsitiData(ricerca) {
    var v = el('vista-risultati');
    var e = ricerca.esiti;

    if (!e.analisi) {
      v.innerHTML = '<div class="vuoto"><span class="grande">Nessuna data riconosciuta</span>' +
        'Vanno digitate forme quali «23/04/2026», «23 aprile 2026», «aprile 2026», «2026», ' +
        '«a.s. 2025/26», «dal 01/01/2026 al 30/06/2026».</div>';
      return;
    }

    if (!e.verbali.length) {
      v.innerHTML = '<div class="vuoto"><span class="grande">Nessuna seduta</span>' +
        'Nell\'archivio non risultano verbali ' + esc(e.analisi.descrizione) + '.' +
        (e.senzaData ? '<br>' + e.senzaData + (e.senzaData === 1 ? ' verbale non ha' : ' verbali non hanno') +
          ' una data rilevata e resta escluso da questa ricerca: si veda la scheda Diagnostica.' : '') +
        '</div>';
      return;
    }

    var h = '<div class="riepilogo"><strong>' + e.verbali.length + '</strong> ' +
      (e.verbali.length === 1 ? 'seduta' : 'sedute') + ' ' + esc(e.analisi.descrizione) + '.' +
      (e.senzaData ? ' <span style="color:var(--allerta)">' + e.senzaData +
        (e.senzaData === 1 ? ' verbale è escluso' : ' verbali sono esclusi') +
        ' perché privi di data rilevata.</span>' : '') + '</div>';
    e.verbali.forEach(function (d) { h += schedaVerbaleCompleto(d); });
    v.innerHTML = h;
  }

  /* ------------------------------------------------------------- ricerca */

  function cerca() {
    var q = el('q').value.trim();
    var modo = document.querySelector('input[name=modo]:checked').value;
    var prefisso = el('opt-prefisso').checked;
    var ambito = el('opt-ambito').value;
    var ordine = el('opt-ordine').value;
    var soloDelibere = el('opt-solo-delibere').checked;

    el('btn-pulisci').hidden = !q;
    el('avviso-modo').textContent = '';

    if (!q) { stato.risultati = []; stato.ultimaRicerca = null; disegnaRisultati(null); return; }

    if (modo === 'numero') {
      var esitiNum = cercaPerNumero(q);
      stato.risultati = [];
      stato.ultimaRicerca = { q: q, modo: 'numero', esiti: esitiNum };
      if (!esitiNum.richieste.length) {
        el('avviso-modo').textContent = 'Non è stato riconosciuto alcun numero. Vanno digitate forme quali «10», «verbale 10», «delibera 96».';
      }
      disegnaEsitiNumero(stato.ultimaRicerca);
      return;
    }

    if (modo === 'data') {
      var esitiData = cercaPerData(q);
      stato.risultati = [];
      stato.ultimaRicerca = { q: q, modo: 'data', esiti: esitiData };
      if (!esitiData.analisi) {
        el('avviso-modo').textContent = 'Non è stata riconosciuta alcuna data. Vanno digitate forme quali «23/04/2026», «aprile 2026», «2026».';
      }
      disegnaEsitiData(stato.ultimaRicerca);
      return;
    }

    var termini = P.norm(q).split(/\s+/).filter(Boolean);
    if (modo === 'parole' && termini.length > 3) {
      el('avviso-modo').textContent = 'Sono stati indicati ' + termini.length +
        ' termini: la ricerca per parole è pensata per un massimo di tre. Vengono usati i primi cinque.';
    }

    var matchers = P.buildMatchers(q, modo, prefisso);
    if (!matchers || !matchers.length) { stato.risultati = []; disegnaRisultati(null); return; }
    var singoli = modo === 'frase' ? matchers : termini.map(function (t) { return P.buildMatchers(t, 'frase', prefisso)[0]; });

    var gruppi = [];
    docs().forEach(function (d) {
      if (d.errore || !d.paragrafi.length) return;
      var nrm = paragrafiNormalizzati(d);
      var perSezione = {};

      function aggiungi(si, idxPar, occ) {
        var k = String(si);
        if (!perSezione[k]) perSezione[k] = { sezione: si, occorrenze: 0, paragrafi: [] };
        perSezione[k].occorrenze += occ;
        perSezione[k].paragrafi.push(idxPar);
      }

      if (ambito === 'paragrafo') {
        for (var i = 0; i < nrm.length; i++) {
          var occ = P.testParagrafo(nrm[i], matchers);
          if (occ) aggiungi(d.paragrafi[i].s, i, occ);
        }
      } else {
        // raggruppa gli indici dei paragrafi per ambito
        var chiavi = {};
        for (var j = 0; j < d.paragrafi.length; j++) {
          var ch = ambito === 'verbale' ? '*' : String(d.paragrafi[j].s);
          if (!chiavi[ch]) chiavi[ch] = [];
          chiavi[ch].push(j);
        }
        Object.keys(chiavi).forEach(function (ch) {
          var lista = chiavi[ch];
          var trovatiPerTermine = singoli.map(function () { return []; });
          lista.forEach(function (idx) {
            singoli.forEach(function (m, k) {
              if (P.testParagrafo(nrm[idx], [m])) trovatiPerTermine[k].push(idx);
            });
          });
          var completo = trovatiPerTermine.every(function (a) { return a.length > 0; });
          if (!completo) return;
          var unione = {}, tot = 0;
          trovatiPerTermine.forEach(function (a) {
            a.forEach(function (idx) { if (!unione[idx]) { unione[idx] = true; tot++; } });
          });
          Object.keys(unione).map(Number).sort(function (a, b) { return a - b; }).forEach(function (idx) {
            var si = ambito === 'verbale' ? d.paragrafi[idx].s : parseInt(ch, 10);
            aggiungi(si, idx, 1);
          });
        });
      }

      var chiaviSez = Object.keys(perSezione);
      if (!chiaviSez.length) return;

      var voci = [];
      chiaviSez.forEach(function (k) {
        var g = perSezione[k];
        var sez = (g.sezione >= 0 && d.sezioni[g.sezione]) ? d.sezioni[g.sezione] : null;
        if (soloDelibere && (!sez || !sez.delibere || !sez.delibere.length)) return;
        voci.push({ doc: d, sezione: sez, sezioneIdx: g.sezione, occorrenze: g.occorrenze, paragrafi: g.paragrafi });
      });
      if (!voci.length) return;

      var totDoc = voci.reduce(function (a, v) { return a + v.occorrenze; }, 0);
      d._occ = totDoc;
      // i punti all'ordine del giorno precedono le parti accessorie
      voci.sort(function (a, b) {
        var pa = (a.sezione && a.sezione.punto) ? 0 : 1;
        var pb = (b.sezione && b.sezione.punto) ? 0 : 1;
        if (pa !== pb) return pa - pb;
        return a.paragrafi[0] - b.paragrafi[0];
      });
      gruppi.push({ doc: d, occorrenze: totDoc, voci: voci });
    });

    var ordinati = ordinaDocs(gruppi.map(function (g) { return g.doc; }), ordine);
    var mappa = {};
    gruppi.forEach(function (g) { mappa[g.doc.nome] = g; });
    stato.risultati = ordinati.map(function (d) { return mappa[d.nome]; });
    stato.ultimaRicerca = {
      q: q, modo: modo, prefisso: prefisso, ambito: ambito,
      matchers: matchers, singoli: singoli
    };
    disegnaRisultati(stato.ultimaRicerca);
  }

  /* ------------------------------------------------------ resa risultati */

  function tagSezione(sez) {
    var h = '';
    if (sez && sez.punto) h += '<span class="tag punto">punto ' + esc(sez.punto) + ' o.d.g.</span>';
    else h += '<span class="tag">parte accessoria</span>';
    if (sez && sez.delibere) {
      sez.delibere.forEach(function (d) {
        h += '<span class="tag delibera">delibera n. ' + esc(d) + '</span>';
      });
    }
    if (sez && sez.esiti) {
      sez.esiti.forEach(function (e) {
        if (e === 'deliberato') return;
        h += '<span class="tag esito">' + esc(e) + '</span>';
      });
    }
    return h;
  }

  /* ------------------------------------- resa degli esiti per numero */

  var RE_FORMULA = /^(delibera|il consiglio d[i'’]istituto|il consiglio di istituto|delibera:)\s*$/i;
  var RE_PREMESSA_RIGA = /^(vist[aoie]|visionat|considerat|preso atto|presa d|premesso|richiamat|sentit|udit|accertat|tenuto conto|ritenut|rilevat|acquisit|esaminat|dato atto|verificat)/i;

  function formattaTestoDelibera(testo) {
    if (!testo) return '<p><em>Testo non disponibile.</em></p>';
    return testo.split('\n').map(function (riga) {
      var t = riga.trim();
      if (!t) return '';
      if (RE_FORMULA.test(t)) return '<p class="formula">' + esc(t) + '</p>';
      if (RE_PREMESSA_RIGA.test(t)) return '<p class="premessa">' + esc(t) + '</p>';
      return '<p>' + esc(t) + '</p>';
    }).join('');
  }

  /* Comandi comuni alle schede di un atto. Il primo argomento e' il nome
     del file, il secondo l'eventuale comando che aggiunge l'atto alla
     raccolta da condividere. */
  function azioniFile(nome, raccolta) {
    return '<div class="azioni-punto">' + (raccolta || '') +
      '<button class="bottone minuto" data-apri="' + esc(nome) + '">Apri il verbale</button>' +
      '<button class="bottone minuto" data-cartella="' + esc(nome) + '">Mostra su GitHub</button></div>';
  }

  function schedaVerbaleCompleto(d) {
    var h = '<div class="scheda-numero">';
    h += '<div class="titolo-atto">' + esc(etichettaVerbale(d)) + '</div>';
    h += '<div class="sottotitolo-atto">' +
      (d.dataLabel ? 'Seduta del ' + esc(d.dataLabel) : 'Data non rilevata') +
      ' · ' + d.sezioni.filter(function (s) { return s.punto; }).length + ' punti all\'ordine del giorno' +
      ' · ' + (d.delibere || []).length + (d.delibere.length === 1 ? ' delibera adottata' : ' delibere adottate') +
      '</div>';

    if (!d.odg.length) {
      h += '<p><em>L\'ordine del giorno non è stato riconosciuto in questo verbale. Il testo resta comunque ricercabile.</em></p>';
    } else {
      // ai punti dell'o.d.g. si associano le delibere rilevate nella trattazione
      var perPunto = {};
      d.sezioni.forEach(function (s) { if (s.punto) perPunto[s.punto] = s; });

      h += '<ul class="odg-elenco">';
      d.odg.forEach(function (it) {
        var s = perPunto[it.num];
        h += '<li' + (it.level > 0 ? ' class="sotto"' : '') + '>';
        h += '<span class="num-punto">' + esc(it.num) + '.</span>';
        h += '<span class="testo-punto">' + esc(it.titolo);
        if (s) {
          (s.delibere || []).forEach(function (num) {
            h += ' <span class="tag delibera">delibera n. ' + esc(num) + '</span>';
          });
          (s.esiti || []).forEach(function (e) {
            if (e === 'deliberato') return;
            h += ' <span class="tag esito">' + esc(e) + '</span>';
          });
        } else if (it.level === 0) {
          h += ' <span class="tag">trattazione non individuata</span>';
        }
        h += '</span></li>';
      });
      h += '</ul>';
    }
    h += azioniFile(d.nome,
      '<button class="bottone minuto rilievo" data-raccogli-odg="' + esc(d.nome) +
      '">Aggiungi l\'ordine del giorno alla raccolta</button>');
    return h + '</div>';
  }

  function schedaDelibera(voce) {
    var d = voce.doc, s = voce.sezione, del = voce.delibera;
    var h = '<div class="scheda-numero">';
    h += '<div class="titolo-atto">Delibera n. ' + esc(del.id) + '</div>';
    h += '<div class="sottotitolo-atto">' + esc(etichettaVerbale(d)) +
      (d.dataLabel ? ' · seduta del ' + esc(d.dataLabel) : '') +
      (s.punto ? ' · punto ' + esc(s.punto) + ' all\'ordine del giorno' : '') + '</div>';
    if (s.titolo) {
      h += '<div class="punto-riga1"><span class="tag punto">oggetto</span>' +
        '<span class="punto-titolo" style="font-size:14.5px">' + esc(s.titolo) + '</span></div>';
    }
    h += '<div class="testo-delibera">' + formattaTestoDelibera(del.testo) + '</div>';
    h += azioniFile(d.nome,
      '<button class="bottone minuto rilievo" data-raccogli-delibera="' + esc(d.nome) +
      '" data-sez="' + voce.sezioneIdx + '" data-delibera="' + esc(del.id) +
      '">Aggiungi la delibera alla raccolta</button>');
    return h + '</div>';
  }

  function disegnaEsitiNumero(ricerca) {
    var v = el('vista-risultati');
    var e = ricerca.esiti;

    if (!e.richieste.length) {
      v.innerHTML = '<div class="vuoto"><span class="grande">Nessun numero riconosciuto</span>' +
        'In questa modalità va digitato un numero. Sono ammesse le forme «10», «verbale 10», ' +
        '«delibera 96», «d 96/2025-26».</div>';
      return;
    }

    if (!e.verbali.length && !e.delibere.length) {
      var elenco = e.richieste.map(function (r) {
        var q = r.tipo === 'verbale' ? 'il verbale n. ' : (r.tipo === 'delibera' ? 'la delibera n. ' : 'il numero ');
        return q + r.numero + (r.anno ? '/' + r.anno : '');
      }).join(', ');
      v.innerHTML = '<div class="vuoto"><span class="grande">Nessun riscontro</span>' +
        'Nell\'archivio non risulta ' + esc(elenco) + '.<br>' +
        'L\'elenco completo è consultabile nelle schede Verbali e Delibere.</div>';
      return;
    }

    var h = '<div class="riepilogo">';
    var parti = [];
    if (e.verbali.length) parti.push('<strong>' + e.verbali.length + '</strong> ' + (e.verbali.length === 1 ? 'verbale' : 'verbali'));
    if (e.delibere.length) parti.push('<strong>' + e.delibere.length + '</strong> ' + (e.delibere.length === 1 ? 'delibera' : 'delibere'));
    h += parti.join(' e ') + '.</div>';

    e.verbali.forEach(function (d) { h += schedaVerbaleCompleto(d); });
    e.delibere.forEach(function (voce) { h += schedaDelibera(voce); });

    v.innerHTML = h;
  }

  /* -------------------------------- trattazione integrale di un punto */

  var contatoreEspansioni = 0;

  function classeCapoverso(t) {
    var s = t.trim();
    if (RE_FORMULA.test(s)) return 'formula';
    if (RE_PREMESSA_RIGA.test(s)) return 'premessa';
    return '';
  }

  /* Ricompone il testo di una sezione a partire dai paragrafi che vi
     appartengono, evidenziando ogni occorrenza dei termini cercati. La
     finestra di contesto e' infinita: il testo non viene troncato. */
  function testoIntegraleSezione(d, sezioneIdx, ricerca) {
    var matchers = ricerca ? (ricerca.singoli && ricerca.singoli.length ? ricerca.singoli : ricerca.matchers) : null;
    var parti = [], occorrenze = 0;
    d.paragrafi.forEach(function (p) {
      if (p.s !== sezioneIdx) return;
      var cl = classeCapoverso(p.t);
      var html;
      if (matchers) {
        var ev = P.evidenzia(p.t, matchers, Infinity, true);
        occorrenze += ev.occorrenze || 0;
        html = ev.html;
      } else {
        html = esc(p.t);
      }
      parti.push('<p' + (cl ? ' class="' + cl + '"' : '') + '>' + html + '</p>');
    });
    if (!parti.length) return { html: '<p><em>Testo non disponibile per questo punto.</em></p>', occorrenze: 0 };
    return { html: parti.join(''), occorrenze: occorrenze, capoversi: parti.length };
  }

  function commutaEspansione(bottone) {
    var cont = el(bottone.dataset.espandi);
    if (!cont) return;

    if (!cont.hidden) {
      cont.hidden = true;
      bottone.textContent = bottone.dataset.etichetta || 'Mostra il punto per intero';
      bottone.classList.add('rilievo');
      return;
    }

    if (!cont.dataset.pronto) {
      var d = stato.indice && stato.indice.docs ? stato.indice.docs[bottone.dataset.doc] : null;
      var si = parseInt(bottone.dataset.sez, 10);
      if (!d) { cont.innerHTML = '<p><em>Verbale non più disponibile nell\'archivio.</em></p>'; }
      else {
        var sez = (si >= 0 && d.sezioni[si]) ? d.sezioni[si] : null;
        var r = testoIntegraleSezione(d, si, stato.ultimaRicerca);
        var intestazione = '<div class="intestazione-integrale">' +
          (sez && sez.punto ? 'Punto ' + esc(sez.punto) + ' all\'ordine del giorno · ' : '') +
          esc(etichettaVerbale(d)) + (d.dataLabel ? ' del ' + esc(d.dataLabel) : '') +
          ' · ' + r.capoversi + (r.capoversi === 1 ? ' capoverso' : ' capoversi') +
          (r.occorrenze ? ' · ' + r.occorrenze + (r.occorrenze === 1 ? ' occorrenza evidenziata' : ' occorrenze evidenziate') : '') +
          '</div>';
        cont.innerHTML = intestazione + r.html;
      }
      cont.dataset.pronto = '1';
    }
    cont.hidden = false;
    if (!bottone.dataset.etichetta) bottone.dataset.etichetta = bottone.textContent;
    bottone.textContent = 'Nascondi il testo integrale';
    bottone.classList.remove('rilievo');
  }

  function disegnaRisultati(ricerca) {
    var v = el('vista-risultati');

    if (!ricerca) {
      var n = docs().filter(function (d) { return !d.errore; }).length;
      var nd = 0;
      docs().forEach(function (d) { nd += (d.delibere || []).length; });
      v.innerHTML = '<div class="vuoto"><span class="grande">Archivio pronto alla consultazione</span>' +
        esc(n) + ' verbali indicizzati, ' + esc(nd) + ' delibere censite.<br>' +
        'Digitare una, due o tre parole nel campo qui sopra.</div>';
      return;
    }

    if (!stato.risultati.length) {
      v.innerHTML = '<div class="vuoto"><span class="grande">Nessun riscontro</span>' +
        'La ricerca ' + (ricerca.modo === 'frase' ? 'della frase' : 'delle parole') + ' &laquo;' + esc(ricerca.q) +
        '&raquo; non ha prodotto risultati.<br>Si può provare con la ricerca per parole, ' +
        'con le forme derivate, oppure ampliando l\'ambito al punto all\'o.d.g. o al verbale.</div>';
      return;
    }

    var totPunti = 0, totOcc = 0;
    stato.risultati.forEach(function (g) { totPunti += g.voci.length; totOcc += g.occorrenze; });

    function pl(n, sing, plur) { return '<strong>' + n + '</strong> ' + (n === 1 ? sing : plur); }

    var h = '<div class="riepilogo">' + pl(totOcc, 'occorrenza', 'occorrenze') + ' in ' +
      pl(totPunti, 'punto', 'punti') + ' di ' + pl(stato.risultati.length, 'verbale', 'verbali') +
      (ricerca.ambito !== 'paragrafo' ? ' (ambito: ' + (ricerca.ambito === 'punto' ? 'stesso punto all\'o.d.g.' : 'stesso verbale') + ')' : '') +
      '. <button class="bottone minuto" data-raccogli-ricerca="1">Aggiungi tutti gli stralci alla raccolta</button></div>';

    stato.risultati.forEach(function (g) {
      var d = g.doc;
      h += '<div class="gruppo-verbale">';
      h += '<div class="intestazione-verbale">' +
        '<span class="numero">' + esc(etichettaVerbale(d)) + '</span>' +
        '<span class="data">' + (d.dataLabel ? 'seduta del ' + esc(d.dataLabel) : 'data non rilevata') + '</span>' +
        '<span class="nomefile">' + esc(d.nome) + '</span></div>';

      g.voci.forEach(function (voce) {
        var sez = voce.sezione;
        h += '<div class="scheda-punto">';
        h += '<div class="punto-riga1">' + tagSezione(sez) +
          '<span class="tag occorrenze">' + voce.occorrenze + (voce.occorrenze === 1 ? ' occorrenza' : ' occorrenze') + '</span></div>';
        h += '<div class="punto-titolo">' + esc(sez ? (sez.titolo || sez.intestazione || '—') : 'Testo non riferibile a un punto specifico') + '</div>';
        if (sez && sez.sottopunti && sez.sottopunti.length) {
          h += '<div class="punto-sottopunti">Sottopunti: ' + esc(sez.sottopunti.join(' · ')) + '</div>';
        }

        h += '<div class="frammenti">';
        var mostrati = 0;
        // Con ambito esteso i termini possono stare in paragrafi diversi: si
        // evidenzia quello che effettivamente compare in ciascun passaggio.
        var perEvidenza = ricerca.ambito === 'paragrafo' ? ricerca.matchers : ricerca.singoli;
        var qualsiasi = ricerca.ambito !== 'paragrafo';
        voce.paragrafi.forEach(function (idx) {
          if (mostrati >= 3) return;
          var testo = d.paragrafi[idx].t;
          var ev = P.evidenzia(testo, perEvidenza, 150, qualsiasi);
          h += '<div class="frammento">' + ev.html + '</div>';
          mostrati++;
        });
        if (voce.paragrafi.length > mostrati) {
          h += '<div class="altri-frammenti">e altri ' + (voce.paragrafi.length - mostrati) + ' passaggi nello stesso punto.</div>';
        }
        h += '</div>';

        // contenitore della trattazione integrale, riempito al primo clic
        var idEsp = 'esp-' + (++contatoreEspansioni);
        h += '<div class="testo-integrale" id="' + idEsp + '" hidden></div>';

        h += '<div class="azioni-punto">' +
          '<button class="bottone minuto rilievo" data-espandi="' + idEsp + '" ' +
          'data-doc="' + esc(d.nome) + '" data-sez="' + voce.sezioneIdx + '">' +
          (sez && sez.punto ? 'Mostra il punto per intero' : 'Mostra il passo per intero') + '</button>' +
          '<button class="bottone minuto" data-raccogli-punto="' + esc(d.nome) + '" ' +
          'data-sez="' + voce.sezioneIdx + '" data-paragrafi="' + esc(JSON.stringify(voce.paragrafi)) + '">' +
          'Aggiungi alla raccolta</button>' +
          '<button class="bottone minuto" data-apri="' + esc(d.nome) + '">Apri il verbale</button>' +
          '<button class="bottone minuto" data-cartella="' + esc(d.nome) + '">Mostra su GitHub</button>' +
          '</div>';
        h += '</div>';
      });
      h += '</div>';
    });

    v.innerHTML = h;
  }

  /* ------------------------------------------------------- vista verbali */

  function disegnaVerbali() {
    var arr = ordinaDocs(docs(), 'recenti');
    if (!arr.length) {
      el('vista-verbali').innerHTML = '<div class="vuoto"><span class="grande">Archivio vuoto</span>' +
        'Nessun verbale presente nell\'archivio.</div>';
      return;
    }
    var h = '<div class="riepilogo">' + arr.length + (arr.length === 1 ? ' file' : ' file') + ' nell\'archivio. Un clic sulla riga mostra i punti all\'ordine del giorno.</div>';
    h += '<div class="tabella-contenitore"><table><thead><tr>' +
      '<th class="num">Verbale</th><th>Seduta</th><th class="num">Punti</th><th class="num">Delibere</th>' +
      '<th>Delibere adottate</th><th>File</th></tr></thead><tbody>';
    arr.forEach(function (d, i) {
      var punti = d.sezioni.filter(function (s) { return s.punto; }).length;
      h += '<tr class="riga-espandi" data-riga="' + i + '">' +
        '<td class="num">' + (d.numero != null ? esc(d.numero) : '—') + (d.annoScolastico ? '<span style="color:var(--testo-debole)">/' + esc(d.annoScolastico) + '</span>' : '') + '</td>' +
        '<td>' + esc(d.dataLabel || '—') + '</td>' +
        '<td class="num">' + punti + '</td>' +
        '<td class="num">' + (d.delibere || []).length + '</td>' +
        '<td class="mono">' + esc((d.delibere || []).join(', ') || '—') + '</td>' +
        '<td class="mono">' + esc(d.nome) + (d.errore ? ' <span style="color:var(--allerta)">· non indicizzato</span>' : '') + '</td>' +
        '</tr>';
      h += '<tr class="dettaglio-verbale" data-dettaglio="' + i + '" hidden><td colspan="6">';
      if (d.errore) {
        h += '<em style="color:var(--allerta)">Indicizzazione non riuscita: ' + esc(d.errore) + '</em>';
      } else if (!d.odg.length) {
        h += '<em>Ordine del giorno non riconosciuto. Il testo resta comunque interamente ricercabile.</em>';
      } else {
        h += '<ol style="list-style:none">';
        d.sezioni.forEach(function (s) {
          if (!s.punto) return;
          h += '<li><strong>' + esc(s.punto) + '.</strong> ' + esc(s.titolo) +
            (s.delibere.length ? ' <span class="tag delibera">delibera n. ' + esc(s.delibere.join(', n. ')) + '</span>' : '') +
            (s.esiti && s.esiti.length ? ' <span class="tag esito">' + esc(s.esiti.join(' · ')) + '</span>' : '') +
            '</li>';
        });
        h += '</ol>';
      }
      h += '<div class="azioni-punto" style="margin-top:10px">' +
        (d.errore ? '' :
          '<button class="bottone minuto" data-raccogli-odg="' + esc(d.nome) + '">Aggiungi l\'o.d.g. alla raccolta</button>' +
          '<button class="bottone minuto" data-apri="' + esc(d.nome) + '">Apri il verbale</button>' +
          '<button class="bottone minuto" data-cartella="' + esc(d.nome) + '">Mostra su GitHub</button>') +
        '<button class="bottone minuto pericolo" data-rimuovi="' + esc(d.nome) + '">Rimuovi dall\'archivio</button>' +
        '</div>';
      h += '</td></tr>';
    });
    h += '</tbody></table></div>';
    el('vista-verbali').innerHTML = h;
    stato._verbaliOrdinati = arr;
  }

  /* ------------------------------------------------------ vista delibere */

  function elencoDelibere() {
    var out = [];
    docs().forEach(function (d) {
      d.sezioni.forEach(function (s, si) {
        (s.delibere || []).forEach(function (num) {
          /* dell'atto si tiene anche la collocazione, perche' la scheda
             consenta di raccoglierne il testo per la condivisione */
          var det = null;
          (s.dettaglioDelibere || []).forEach(function (x) {
            if (!det && String(x.numero) === String(parseInt(num, 10))) det = x;
          });
          out.push({
            numero: num, ordn: parseInt(num, 10) || 0,
            verbale: d.numero, dataISO: d.dataISO, data: d.dataLabel,
            punto: s.punto, oggetto: s.titolo, nome: d.nome,
            sez: si, id: det ? det.id : null
          });
        });
      });
    });
    out.sort(function (a, b) {
      if ((b.dataISO || '') !== (a.dataISO || '')) return (b.dataISO || '') < (a.dataISO || '') ? -1 : 1;
      return b.ordn - a.ordn;
    });
    return out;
  }

  function disegnaDelibere(filtro) {
    var lista = elencoDelibere();
    var f = P.norm(filtro || '');
    if (f) {
      lista = lista.filter(function (r) {
        return P.norm(r.oggetto + ' ' + r.numero + ' ' + (r.data || '') + ' verbale ' + r.verbale).indexOf(f) >= 0;
      });
    }
    var h = '<div class="filtro-tabella"><input id="filtro-delibere" type="search" placeholder="Filtra per oggetto o numero…" value="' + esc(filtro || '') + '">' +
      '<span>' + lista.length + (lista.length === 1 ? ' delibera' : ' delibere') + '</span></div>';
    if (!lista.length) {
      h += '<div class="vuoto"><span class="grande">Nessuna delibera</span>Non risultano delibere corrispondenti.</div>';
      el('vista-delibere').innerHTML = h;
      agganciaFiltroDelibere();
      return;
    }
    h += '<div class="tabella-contenitore"><table><thead><tr>' +
      '<th class="num">Delibera</th><th class="num">Verbale</th><th>Seduta</th><th class="num">Punto</th><th>Oggetto</th></tr></thead><tbody>';
    lista.forEach(function (r) {
      h += '<tr>' +
        '<td class="num"><strong>' + esc(r.numero) + '</strong></td>' +
        '<td class="num">' + (r.verbale != null ? 'n. ' + esc(r.verbale) : '—') + '</td>' +
        '<td>' + esc(r.data || '—') + '</td>' +
        '<td class="num">' + esc(r.punto || '—') + '</td>' +
        '<td>' + esc(r.oggetto) +
        (r.id ? ' <button class="bottone minuto" data-raccogli-delibera="' + esc(r.nome) +
          '" data-sez="' + r.sez + '" data-delibera="' + esc(r.id) + '">raccogli</button>' : '') +
        ' <button class="bottone minuto" data-apri="' + esc(r.nome) + '">apri</button></td>' +
        '</tr>';
    });
    h += '</tbody></table></div>';
    el('vista-delibere').innerHTML = h;
    agganciaFiltroDelibere();
  }

  function agganciaFiltroDelibere() {
    var inp = el('filtro-delibere');
    if (!inp) return;
    inp.addEventListener('input', function () {
      var pos = this.selectionStart, val = this.value;
      disegnaDelibere(val);
      var nuovo = el('filtro-delibere');
      if (nuovo) { nuovo.focus(); try { nuovo.setSelectionRange(pos, pos); } catch (e) { } }
    });
  }

  function csvDelibere() {
    var righe = [['Delibera', 'Verbale', 'Seduta', 'Punto o.d.g.', 'Oggetto', 'File']];
    elencoDelibere().forEach(function (r) {
      righe.push([r.numero, r.verbale == null ? '' : r.verbale, r.data || '', r.punto || '', r.oggetto, r.nome]);
    });
    return '\uFEFF' + righe.map(function (r) {
      return r.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(';');
    }).join('\r\n');
  }

  /* -------------------------------------------- rimozione e ripristino */

  /* La rimozione sposta il file nel Cestino e cancella immediatamente dal
     memorizzato i dati del verbale, senza attendere la riconciliazione con la
     cartella: l'eliminazione dei dati dall'archivio deve essere certa. */
  function rimuoviVerbale(nome) {
    var d = stato.indice && stato.indice.docs ? stato.indice.docs[nome] : null;
    var descr = nome;
    if (d && d.numero != null) {
      descr = etichettaVerbale(d) + (d.dataLabel ? ' del ' + d.dataLabel : '') + '\n' + nome;
    }
    if (!confirm('Rimuovere dall\'archivio il seguente verbale?\n\n' + descr +
                 '\n\nIl file viene spostato nel Cestino, da cui potrà essere ripristinato. ' +
                 'I dati corrispondenti vengono eliminati dall\'indice.')) return;

    velo(true, 'Rimozione in corso', nome, 50);
    chiediJson('/api/rimuovi?nome=' + encodeURIComponent(nome), { method: 'POST' })
      .then(function () {
        if (stato.indice && stato.indice.docs) delete stato.indice.docs[nome];
        delete stato.normCache[nome];
        stato.indice.generato = new Date().toISOString();
        return salvaIndice();
      })
      .then(function () { return costruisciIndice(false); })
      .then(function () { return aggiornaCestino(); })
      .then(function () {
        velo(false);
        ridisegnaTutto();
        alert('Verbale rimosso dall\'archivio e cancellato dall\'indice.\n\n' +
              'Il file si trova ora nel Cestino, consultabile dall\'omonima scheda.');
      })
      .catch(function (e) {
        velo(false);
        alert('La rimozione non è andata a buon fine: ' + e.message);
      });
  }

  function ripristinaVerbale(stored, originale, sovrascrivi) {
    var url = '/api/ripristina?nome=' + encodeURIComponent(stored) + (sovrascrivi ? '&sovrascrivi=1' : '');
    velo(true, 'Ripristino in corso', originale, 50);
    return chiediJson(url, { method: 'POST' })
      .then(function () { return costruisciIndice(false); })
      .then(function () { return aggiornaCestino(); })
      .then(function () {
        velo(false);
        ridisegnaTutto();
        alert('Verbale ripristinato nell\'archivio e nuovamente indicizzato.');
      })
      .catch(function (e) {
        velo(false);
        if (e.status === 409) {
          if (confirm('Nell\'archivio esiste già un file di nome:\n\n' + originale +
                      '\n\nSostituirlo con quello che si sta ripristinando?')) {
            return ripristinaVerbale(stored, originale, true);
          }
          return null;
        }
        alert('Il ripristino non è andato a buon fine: ' + e.message);
      });
  }

  function eliminaDalCestino(stored, originale) {
    if (!confirm('Cancellare definitivamente il seguente file?\n\n' + originale +
                 '\n\nL\'operazione non è reversibile: il file viene eliminato dal repository e ' +
                 'nessuna funzione dell\'applicazione potrà recuperarlo.')) return;
    chiediJson('/api/elimina?nome=' + encodeURIComponent(stored), { method: 'POST' })
      .then(function () { return aggiornaCestino(); })
      .then(function () { disegnaCestino(); })
      .catch(function (e) { alert('La cancellazione non è andata a buon fine: ' + e.message); });
  }

  function aggiornaCestino() {
    return chiediJson('/api/cestino').then(function (r) {
      stato.cestino = (r && r.files) ? r.files : [];
      stato.cartellaCestino = r ? r.cartella : null;
      var c = el('conta-cestino');
      c.textContent = stato.cestino.length;
      c.hidden = stato.cestino.length === 0;
      return stato.cestino;
    }).catch(function () {
      stato.cestino = [];
      el('conta-cestino').hidden = true;
      return [];
    });
  }

  function disegnaCestino() {
    var lista = stato.cestino || [];
    if (!lista.length) {
      el('vista-cestino').innerHTML = '<div class="vuoto"><span class="grande">Cestino vuoto</span>' +
        'I verbali rimossi dall\'archivio finiscono qui e restano recuperabili.<br>' +
        'La rimozione si comanda dalla scheda Verbali, espandendo la riga del verbale.</div>';
      return;
    }
    var h = '<div class="riquadro"><h3>Verbali rimossi dall\'archivio</h3>' +
      '<p>Questi file non compaiono più nell\'archivio e i loro dati sono stati cancellati ' +
      'dall\'indice, ma restano nel repository dei dati, nella cartella <span class="mono">' +
      esc(stato.cartellaCestino || 'Cestino') + '</span>. Il ripristino li riporta ' +
      'nell\'archivio e li reindicizza. La cancellazione definitiva non è reversibile.</p></div>';

    h += '<div class="tabella-contenitore"><table><thead><tr>' +
      '<th>File rimosso</th><th>Rimosso il</th><th class="num">Dimensione</th><th>Operazioni</th>' +
      '</tr></thead><tbody>';
    lista.forEach(function (f) {
      var quando = '—';
      if (f.rimossoIl) {
        var dt = new Date(f.rimossoIl);
        if (!isNaN(dt.getTime())) quando = dt.toLocaleString('it-IT');
      }
      h += '<tr>' +
        '<td class="mono">' + esc(f.originale) + '</td>' +
        '<td>' + esc(quando) + '</td>' +
        '<td class="num">' + dimensioneLeggibile(f.dimensione) + '</td>' +
        '<td><div class="azioni-punto" style="margin:0">' +
        '<button class="bottone minuto" data-ripristina="' + esc(f.nome) + '" data-orig="' + esc(f.originale) + '">Ripristina</button>' +
        '<button class="bottone minuto" data-apri-cestino="' + esc(f.nome) + '">Apri</button>' +
        '<button class="bottone minuto pericolo" data-elimina="' + esc(f.nome) + '" data-orig="' + esc(f.originale) + '">Cancella definitivamente</button>' +
        '</div></td></tr>';
    });
    h += '</tbody></table></div>';
    el('vista-cestino').innerHTML = h;
  }

  /* ----------------------------------------------------- vista presenze */

  function etichettaSeduta(d) {
    return (d.numero != null ? 'n. ' + d.numero : '—') + '  ' + (d.dataLabel || '');
  }

  function disegnaAvvisoPresenze(p) {
    var box = el('avviso-presenze');
    var conta = el('conta-presenze');

    var rossi = p.allarmi.filter(function (m) { return m.livello === 'rosso'; });
    var arancioni = p.allarmi.filter(function (m) { return m.livello === 'arancione'; });

    conta.textContent = p.allarmi.length;
    conta.hidden = p.allarmi.length === 0;
    conta.className = 'contatore' + (rossi.length ? ' allarme' : '');

    if (!p.allarmi.length) { box.hidden = true; return; }

    box.hidden = false;
    box.className = 'avviso-presenze' + (rossi.length ? '' : ' solo-arancione');

    var h = '<div class="corpo">';
    if (rossi.length) {
      h += '<div class="lampeggio"><span class="pallino"></span><strong>' +
        (rossi.length === 1 ? 'Un componente ha' : rossi.length + ' componenti hanno') +
        ' più di due assenze consecutive.</strong></div>' +
        '<div class="nomi">' + rossi.map(function (m) {
          return esc(m.nome) + ' (' + m.serieMax + ' assenze consecutive)';
        }).join('; ') + '</div>';
    }
    if (arancioni.length) {
      h += '<div class="lampeggio"' + (rossi.length ? ' style="margin-top:7px"' : '') + '>' +
        '<span class="pallino"></span><strong>' +
        (arancioni.length === 1 ? 'Un componente ha' : arancioni.length + ' componenti hanno') +
        ' due assenze consecutive.</strong></div>' +
        '<div class="nomi">' + arancioni.map(function (m) { return esc(m.nome); }).join('; ') + '</div>';
    }
    h += '</div><button class="bottone minuto" data-scheda-vai="presenze">Apri la scheda Presenze</button>';
    box.innerHTML = h;
  }

  function schedaAllarme(m) {
    var h = '<div class="scheda-allarme ' + m.livello + '">';
    h += '<div class="punto-riga1">';
    if (m.livello === 'rosso') h += '<span class="tag rosso">oltre due assenze consecutive</span>';
    else if (m.livello === 'arancione') h += '<span class="tag arancione">due assenze consecutive</span>';
    else if (m.livello === 'chiuso') h += '<span class="tag chiuso">decadenza registrata</span>';
    if (!m.inCarica) h += '<span class="tag">non più negli elenchi</span>';
    if (m.silenziato) h += '<span class="tag">avviso silenziato</span>';
    h += '</div>';
    h += '<div class="nome-componente">' + esc(m.nome) + '</div>';
    h += '<div class="dettaglio">' + esc(m.ruolo || 'ruolo non rilevato') +
      ' · ' + m.totPresenze + ' presenze, ' + m.totAssenze + ' assenze su ' + p_conta(m) + ' sedute';
    if (m.serieRossa && m.serieRossa.length) {
      h += '<br>Serie che ha determinato l\'allarme: ' +
        esc(m.serieRossa.map(etichettaSeduta).join(' · '));
    } else if (m.livello === 'arancione' && m.primaDellaSerie) {
      h += '<br>Assenze consecutive in corso dalla seduta ' + esc(etichettaSeduta(m.primaDellaSerie));
    }
    if (m.decadenza) {
      h += '<br>Decadenza dichiarata con delibera n. ' + esc(m.decadenza.delibera) +
        (m.decadenza.data ? ' del ' + esc(m.decadenza.data) : '') +
        (m.decadenza.annotato ? ' · registrata il ' + esc(new Date(m.decadenza.annotato).toLocaleDateString('it-IT')) : '');
    }
    if (!m.inCarica && m.livello === 'rosso' && !m.decadenza) {
      h += '<br><em>Il componente non compare nella tabella presenze del verbale più recente. ' +
        'L\'avviso non lampeggia più in apertura, ma resta aperto finché non viene registrata ' +
        'la delibera di decadenza o non se ne dà atto in altro modo.</em>';
    }
    h += '</div>';

    h += '<div class="azioni-punto">';
    h += '<button class="bottone minuto" data-silenzia="' + esc(m.chiave) + '">' +
      (m.silenziato ? 'Riattiva l\'avviso' : 'Silenzia l\'avviso') + '</button>';
    if (m.decadenza) {
      h += '<button class="bottone minuto pericolo" data-revoca-decadenza="' + esc(m.chiave) + '">Revoca la registrazione</button>';
    } else if (m.livello === 'rosso') {
      h += '<button class="bottone minuto rilievo" data-apri-decadenza="' + esc(m.chiave) + '">Registra la delibera di decadenza</button>';
    }
    h += '</div>';

    if (m.livello === 'rosso' && !m.decadenza) {
      h += '<div class="modulo-decadenza" id="dec-' + esc(m.chiave).replace(/[^a-z0-9]/g, '-') + '" hidden>' +
        '<div><label>Delibera n.</label><input class="stretto" data-campo="delibera" placeholder="107/2025-26"></div>' +
        '<div><label>del</label><input class="stretto" data-campo="data" placeholder="15/05/2026"></div>' +
        '<button class="bottone minuto primario" data-salva-decadenza="' + esc(m.chiave) + '">Registra</button>' +
        '</div>';
    }
    return h + '</div>';
  }

  function p_conta(m) {
    return m.totPresenze + m.totAssenze;
  }

  /* Composizione del Consiglio: dichiarazione di cessazione dal mandato,
     revoca, e comando che riporta in vista i componenti cessati. */
  function riquadroComposizione(p) {
    var h = '<div class="riquadro" style="margin-top:18px"><h3>Composizione del Consiglio</h3>';
    h += '<p>L\'elenco dei componenti è ricavato dalle tabelle presenze dei verbali. ' +
      'Chi ha cessato il mandato va dichiarato con la relativa decorrenza: le sedute anteriori ' +
      'conservano il dato che i verbali documentano, quelle successive non lo riguardano più, ' +
      'e gli avvisi a suo carico si chiudono.</p>';

    h += '<div class="azioni-punto" style="margin:0 0 4px">' +
      '<button class="bottone minuto rilievo" data-apri-cessazione="1">' +
      (stato.pannelloCessazione ? 'Chiudi l\'elenco' : 'Dichiara componenti cessati dal mandato') + '</button>' +
      '<button class="bottone minuto" data-apri-esclusione="1">' +
      (stato.pannelloEsclusione ? 'Chiudi l\'elenco' : 'Escludi dal monitoraggio') + '</button>';
    if (p.cessati.length) {
      h += '<label class="opzione-inline" style="margin-left:10px">' +
        '<input type="checkbox" id="opt-mostra-cessati"' + (stato.mostraCessati ? ' checked' : '') + '> ' +
        'mostra anche i ' + p.cessati.length + ' cessati nel prospetto</label>';
    }
    h += '</div>';

    if (stato.pannelloCessazione) {
      h += '<div class="modulo-cessazione">';
      h += '<div style="width:100%"><label>Cessazione a decorrere dal</label>' +
        '<input id="data-cessazione" placeholder="01/01/2026 oppure gennaio 2026" value="' +
        esc(stato.bozzaDataCessazione || '') + '"></div>';
      h += '<div class="elenco-cessazione">';
      var disponibili = p.attuali.slice().sort(function (a, b) { return a.nome.localeCompare(b.nome, 'it'); });
      if (!disponibili.length) h += '<em>Nessun componente da dichiarare.</em>';
      disponibili.forEach(function (m) {
        h += '<label class="voce-cessazione"><input type="checkbox" data-cessa="' + esc(m.chiave) + '"> ' +
          esc(m.nome) + '<small>' + esc(m.ruolo || '—') + '</small></label>';
      });
      h += '</div>';
      h += '<button class="bottone minuto primario" data-salva-cessazione="1">Registra la cessazione</button>';
      h += '</div>';
    }

    if (stato.pannelloEsclusione) {
      h += '<div class="modulo-cessazione">';
      h += '<p style="font-size:13px;color:var(--testo-tenue);margin:0 0 8px">' +
        'L\'esclusione toglie il componente dal monitoraggio delle presenze in modo integrale: ' +
        'non compare fra gli avvisi, non compare nel prospetto, non viene conteggiato. ' +
        'Il testo dei verbali resta immutato e i suoi nominativi restano ricercabili.</p>';
      h += '<div class="elenco-cessazione">';
      var candidati = p.membri.slice().sort(function (a, b) { return a.nome.localeCompare(b.nome, 'it'); });
      if (!candidati.length) h += '<em>Nessun componente da escludere.</em>';
      candidati.forEach(function (m) {
        h += '<label class="voce-cessazione"><input type="checkbox" data-escludi="' + esc(m.chiave) + '"> ' +
          esc(m.nome) + '<small>' + esc(m.ruolo || '—') + '</small></label>';
      });
      h += '</div>';
      h += '<button class="bottone minuto primario" data-salva-esclusione="1">Escludi i selezionati</button>';
      h += '</div>';
    }

    var esclusi = statoPresenze().esclusi || [];
    if (esclusi.length) {
      h += '<div class="riquadro-esclusi"><strong>Esclusi dal monitoraggio delle presenze:</strong> ';
      h += esclusi.map(function (e, i) {
        return '<span class="voce-esclusa">' + esc(e.etichetta || e.pattern) +
          '<button class="bottone minuto" data-reintegra="' + i + '" title="Riporta il nominativo nel monitoraggio">reintegra</button></span>';
      }).join(' ');
      h += '<div style="margin-top:6px;font-size:12px;color:var(--testo-debole)">' +
        'Il riconoscimento avviene per cognome, senza distinzione di accenti e maiuscole. ' +
        'Un eventuale omonimo verrebbe escluso anch\'esso: in tal caso va reintegrato e ' +
        'sostituito con l\'esclusione del solo componente interessato.</div>';
      h += '</div>';
    }

    if (p.cessati.length) {
      h += '<div class="tabella-contenitore" style="margin-top:12px"><table><thead><tr>' +
        '<th>Componente cessato</th><th>Ruolo</th><th>Dal</th><th class="num">Sedute</th><th></th>' +
        '</tr></thead><tbody>';
      p.cessati.forEach(function (m) {
        h += '<tr><td>' + esc(m.nome) + '</td><td>' + esc(m.ruolo || '—') + '</td><td>' +
          esc(m.cessato.dataLabel || m.cessato.data || '—') + '</td><td class="num">' +
          (m.totPresenze + m.totAssenze) + '</td><td>' +
          '<button class="bottone minuto" data-revoca-cessazione="' + esc(m.chiave) + '">Revoca</button>' +
          '</td></tr>';
      });
      h += '</tbody></table></div>';
    }

    return h + '</div>';
  }

  function disegnaPresenze() {
    var p = calcolaPresenze();
    stato.presenzeCalcolo = p;
    disegnaAvvisoPresenze(p);

    var v = el('vista-presenze');

    if (!p.conTabella.length) {
      v.innerHTML = '<div class="vuoto"><span class="grande">Nessuna tabella delle presenze</span>' +
        'In nessuno dei verbali indicizzati è stata riconosciuta la tabella dei componenti.<br>' +
        'Il riconoscimento è affidabile sui file .docx, dove la tabella ha una struttura propria; ' +
        'sui .pdf dipende dall\'allineamento dei segni di spunta.</div>';
      return;
    }

    var h = '';

    // avvisi in corso e provvedimenti
    var daMostrare = p.membri.filter(function (m) {
      return m.livello === 'rosso' || m.livello === 'arancione' || m.livello === 'chiuso';
    });
    if (daMostrare.length) {
      h += '<div class="riepilogo">' + daMostrare.length +
        (daMostrare.length === 1 ? ' componente segnalato' : ' componenti segnalati') + '.</div>';
      daMostrare.forEach(function (m) { h += schedaAllarme(m); });
    } else {
      h += '<div class="riquadro"><h3>Nessuna segnalazione</h3>' +
        '<p>Nessun componente presenta due o più assenze consecutive nelle sedute indicizzate.</p></div>';
    }

    // prospetto completo
    h += '<div class="riquadro" style="margin-top:18px"><h3>Prospetto delle presenze</h3>' +
      '<div class="azioni-punto" style="margin:0 0 8px">' +
      '<button class="bottone minuto rilievo" data-stampa-presenze="1">Esporta il prospetto in PDF</button>' +
      '<button class="bottone minuto" data-raccogli-presenze="1">Aggiungi il prospetto alla raccolta</button>' +
      '</div>' +
      '<p>Una colonna per seduta, in ordine cronologico. <strong>P</strong> indica presenza, ' +
      '<strong>A</strong> assenza, il trattino che il componente non risulta negli elenchi di quella seduta. ' +
      'Un clic sulla casella ne corregge il contenuto, quando l\'estrazione automatica abbia sbagliato; ' +
      'le caselle corrette portano un puntino.</p>';
    if (p.senzaTabella.length) {
      h += '<p style="color:var(--allerta)">' + p.senzaTabella.length +
        (p.senzaTabella.length === 1 ? ' verbale non ha' : ' verbali non hanno') +
        ' una tabella presenze riconosciuta e non compaiono nel prospetto: ' +
        esc(p.senzaTabella.map(function (d) { return etichettaSeduta(d); }).join(' · ')) + '.</p>';
    }
    h += '</div>';

    var visibili = stato.mostraCessati ? p.membri : p.attuali;
    if (p.cessati.length && !stato.mostraCessati) {
      h += '<p style="font-size:12.5px;color:var(--testo-debole);margin:-4px 0 10px">' +
        p.cessati.length + (p.cessati.length === 1 ? ' componente cessato dal mandato non compare' : ' componenti cessati dal mandato non compaiono') +
        ' nel prospetto. Il comando che li riporta in vista si trova nel riquadro «Composizione del Consiglio», più sotto.</p>';
    }

    h += '<div class="tabella-contenitore"><table class="griglia-presenze"><thead><tr>' +
      '<th>Componente</th>';
    p.conTabella.forEach(function (d) {
      h += '<th class="seduta">' + esc(etichettaSeduta(d)) + '</th>';
    });
    h += '<th class="num">Consec.</th><th class="num">Ass.</th><th>Stato</th></tr></thead><tbody>';

    visibili.forEach(function (m) {
      var cls = m.livello === 'rosso' ? ' class="riga-rosso"' :
        (m.livello === 'arancione' ? ' class="riga-arancione"' :
          (m.cessato ? ' class="riga-cessato"' : ''));
      h += '<tr' + cls + '><td class="nome">' + esc(m.nome) +
        '<small>' + esc(m.ruolo || '—') +
        (m.cessato ? ' · cessato dal ' + esc(m.cessato.dataLabel || m.cessato.data) : '') + '</small></td>';
      p.conTabella.forEach(function (d) {
        var st = m.stati[d.nome];
        var corretta = statoPresenze().correzioni[d.nome + '||' + m.chiave] ? ' corretta' : '';
        var lab = st === 'presente' ? 'P' : (st === 'assente' ? 'A' : '–');
        var c = st === 'presente' ? 'presente' : (st === 'assente' ? 'assente' : 'vuoto');
        h += '<td class="cella ' + c + corretta + '" data-cella="' + esc(d.nome) + '" data-chiave="' + esc(m.chiave) +
          '" title="' + esc(m.nome + ' · ' + etichettaSeduta(d)) + '">' + lab + '</td>';
      });
      h += '<td class="num">' + (m.serie || '') + '</td><td class="num">' + m.totAssenze + '</td><td>';
      if (m.livello === 'rosso') h += '<span class="tag rosso">rosso</span>';
      else if (m.livello === 'arancione') h += '<span class="tag arancione">arancione</span>';
      else if (m.livello === 'chiuso') h += '<span class="tag chiuso">decaduto</span>';
      else if (m.livello === 'cessato') h += '<span class="tag chiuso">cessato</span>';
      else h += '<span style="color:var(--testo-debole)">—</span>';
      if (!m.inCarica && !m.cessato) h += ' <span class="tag">fuori</span>';
      h += '</td></tr>';
    });
    h += '</tbody></table></div>';

    h += riquadroComposizione(p);

    el('vista-presenze').innerHTML = h;
  }

  /* --------------------------------------- comandi della scheda presenze */

  /* Stato ricavato dal documento, senza tenere conto delle correzioni. */
  function statoEstratto(doc, chiave) {
    var righe = (doc && doc.presenze && doc.presenze.righe) ? doc.presenze.righe : [];
    for (var i = 0; i < righe.length; i++) {
      if (chiaveNome(righe[i].nome) === chiave) {
        return righe[i].stato === 'assente' ? 'assente' : (righe[i].stato === 'presente' ? 'presente' : null);
      }
    }
    return null;
  }

  /* Il clic fa ruotare la casella fra presente, assente e non in carica.
     Quando il valore raggiunto coincide con quello estratto dal documento la
     correzione viene rimossa, così che la casella torni al dato originario. */
  function commutaCella(nomeFile, chiave) {
    var s = statoPresenze();
    var k = nomeFile + '||' + chiave;
    var d = stato.indice.docs[nomeFile];
    if (!d) return;
    var attuale = statoIn(d, chiave);
    var prossimo = attuale === 'presente' ? 'assente' : (attuale === 'assente' ? null : 'presente');
    if (prossimo === statoEstratto(d, chiave)) delete s.correzioni[k];
    else s.correzioni[k] = prossimo === null ? 'fuori' : prossimo;
    salvaPresenze().then(function () { disegnaPresenze(); });
  }

  function commutaSilenzio(chiave) {
    var s = statoPresenze();
    if (s.silenziati[chiave]) delete s.silenziati[chiave];
    else s.silenziati[chiave] = new Date().toISOString();
    salvaPresenze().then(function () { disegnaPresenze(); });
  }

  function registraDecadenza(chiave, delibera, data) {
    if (!delibera) { alert('Va indicato il numero della delibera con cui il Consiglio ha dichiarato la decadenza.'); return; }
    var s = statoPresenze();
    s.decadenze[chiave] = { delibera: delibera, data: data || '', annotato: new Date().toISOString() };
    salvaPresenze().then(function () { disegnaPresenze(); });
  }

  function revocaDecadenza(chiave) {
    if (!confirm('Revocare la registrazione della decadenza? L\'avviso torna attivo.')) return;
    delete statoPresenze().decadenze[chiave];
    salvaPresenze().then(function () { disegnaPresenze(); });
  }

  /* ------------------------------------- cessazione dal mandato */

  /* La decorrenza si ricava con il medesimo riconoscimento usato dalla ricerca
     per data: sono ammesse tanto "01/01/2026" quanto "gennaio 2026". */
  function decorrenzaCessazione(testo) {
    var an = analizzaData(testo || '');
    if (!an || !an.intervalli.length) return null;
    return an.intervalli[0].da;
  }

  function registraCessazione() {
    var campo = el('data-cessazione');
    var testo = campo ? campo.value.trim() : '';
    var data = decorrenzaCessazione(testo);
    if (!data) {
      alert('Non è stata riconosciuta la data di decorrenza.\n\n' +
        'Sono ammesse forme quali 01/01/2026, 1 gennaio 2026, gennaio 2026.');
      return;
    }

    var scelti = Array.prototype.slice.call(document.querySelectorAll('[data-cessa]:checked'))
      .map(function (i) { return i.getAttribute('data-cessa'); });
    if (!scelti.length) { alert('Non è stato selezionato alcun componente.'); return; }

    var p = stato.presenzeCalcolo || calcolaPresenze();
    var nomi = scelti.map(function (k) {
      var m = p.membri.filter(function (x) { return x.chiave === k; })[0];
      return m ? m.nome : k;
    });

    if (!confirm('Dichiarare cessati dal mandato, a decorrere dal ' + itData(data) + ', i seguenti componenti?\n\n· ' +
      nomi.join('\n· ') + '\n\nLe sedute anteriori a quella data conservano le loro presenze; ' +
      'quelle successive non li riguarderanno e gli avvisi a loro carico si chiuderanno.')) return;

    var s = statoPresenze();
    scelti.forEach(function (k, i) {
      s.cessati[k] = { data: data, dataLabel: itData(data), nome: nomi[i], annotato: new Date().toISOString() };
    });
    stato.pannelloCessazione = false;
    stato.bozzaDataCessazione = '';
    salvaPresenze().then(function () {
      disegnaPresenze();
      alert(nomi.length + (nomi.length === 1 ? ' componente dichiarato cessato' : ' componenti dichiarati cessati') +
        ' dal mandato a decorrere dal ' + itData(data) + '.');
    });
  }

  function registraEsclusione() {
    var scelti = Array.prototype.slice.call(document.querySelectorAll('[data-escludi]:checked'))
      .map(function (i) { return i.getAttribute('data-escludi'); });
    if (!scelti.length) { alert('Non è stato selezionato alcun componente.'); return; }

    var p = stato.presenzeCalcolo || calcolaPresenze();
    var s = statoPresenze();
    var nomi = [];
    scelti.forEach(function (k) {
      var m = p.membri.filter(function (x) { return x.chiave === k; })[0];
      var nome = m ? m.nome : k;
      nomi.push(nome);
      // si esclude il nominativo per intero, non il solo cognome: la grafia è
      // qui nota, essendo stata ricavata dai verbali
      if (!s.esclusi.some(function (e) { return e.pattern === chiaveNome(nome); })) {
        s.esclusi.push({ pattern: chiaveNome(nome), etichetta: nome });
      }
    });

    if (!confirm('Escludere dal monitoraggio delle presenze i seguenti componenti?\n\n· ' +
      nomi.join('\n· ') + '\n\nNon compariranno più fra gli avvisi né nel prospetto. ' +
      'Il testo dei verbali non viene modificato.')) {
      nomi.forEach(function (n) {
        s.esclusi = s.esclusi.filter(function (e) { return e.pattern !== chiaveNome(n); });
      });
      return;
    }

    stato.pannelloEsclusione = false;
    salvaPresenze().then(function () { disegnaPresenze(); });
  }

  function reintegraEscluso(indice) {
    var s = statoPresenze();
    var e = s.esclusi[indice];
    if (!e) return;
    if (!confirm('Riportare «' + (e.etichetta || e.pattern) + '» nel monitoraggio delle presenze?')) return;
    s.esclusi.splice(indice, 1);
    salvaPresenze().then(function () { disegnaPresenze(); });
  }

  function revocaCessazione(chiave) {
    var s = statoPresenze();
    var nome = (s.cessati[chiave] && s.cessati[chiave].nome) || chiave;
    if (!confirm('Revocare la cessazione dal mandato di ' + nome + '?\n\n' +
      'Il componente torna a figurare fra quelli in carica.')) return;
    delete s.cessati[chiave];
    salvaPresenze().then(function () { disegnaPresenze(); });
  }

  /* --------------------------------------------------- vista diagnostica */

  function disegnaDiagnostica() {
    var arr = ordinaDocs(docs(), 'recenti');
    var h = '';
    var errori = arr.filter(function (d) { return d.errore; });
    var scansioni = arr.filter(function (d) { return !d.errore && d.avviso; });
    var senzaOdg = arr.filter(function (d) { return !d.errore && !d.odg.length; });
    var nonLocalizzati = arr.filter(function (d) { return !d.errore && d.puntiNonLocalizzati && d.puntiNonLocalizzati.length; });
    var senzaData = arr.filter(function (d) { return !d.errore && !d.dataLabel; });
    var senzaNumero = arr.filter(function (d) { return !d.errore && d.numero == null; });

    if (!arr.length) {
      el('vista-diagnostica').innerHTML = '<div class="vuoto"><span class="grande">Nulla da verificare</span>L\'archivio è vuoto.</div>';
      return;
    }

    var tot = arr.length, ok = arr.filter(function (d) { return !d.errore; }).length;
    var car = arr.reduce(function (a, d) { return a + (d.caratteri || 0); }, 0);
    var par = arr.reduce(function (a, d) { return a + (d.paragrafi || []).length; }, 0);

    h += '<div class="riquadro"><h3>Stato dell\'indice</h3>' +
      '<p>' + ok + ' verbali su ' + tot + ' risultano indicizzati, per complessivi ' +
      par.toLocaleString('it-IT') + ' paragrafi e ' + car.toLocaleString('it-IT') + ' caratteri ricercabili.' +
      (stato.indice && stato.indice.generato ? ' Ultimo aggiornamento: ' + new Date(stato.indice.generato).toLocaleString('it-IT') + '.' : '') +
      '</p></div>';

    function blocco(titolo, testo, lista, grave) {
      if (!lista.length) return '';
      var s = '<div class="riquadro' + (grave ? ' attenzione' : '') + '"><h3>' + esc(titolo) + '</h3><p>' + testo + '</p><ul>';
      lista.forEach(function (x) { s += '<li>' + x + '</li>'; });
      return s + '</ul></div>';
    }

    h += blocco('File non indicizzati', 'Questi file non è stato possibile leggerli. I formati .doc di vecchia generazione vanno riaperti con Word e salvati in .docx.',
      errori.map(function (d) { return esc(d.nome) + ' — ' + esc(d.errore); }), true);

    h += blocco('Verbali privi di testo estraibile', 'Si tratta con ogni probabilità di scansioni prive di riconoscimento ottico dei caratteri: non sono ricercabili finché non si applica un OCR.',
      scansioni.map(function (d) { return esc(d.nome) + ' — ' + esc(d.caratteri) + ' caratteri estratti'; }), true);

    h += blocco('Ordine del giorno non riconosciuto', 'Il testo resta interamente ricercabile, ma i riscontri non potranno essere attribuiti a un punto specifico.',
      senzaOdg.map(function (d) { return esc(d.nome); }), false);

    h += blocco('Punti non localizzati nel corpo del verbale', 'Questi punti compaiono nell\'elenco iniziale ma la relativa trattazione non è stata individuata nel corpo del documento, di norma perché l\'intestazione usata nel testo differisce sensibilmente dalla dizione dell\'ordine del giorno.',
      nonLocalizzati.map(function (d) { return '<strong>' + esc(d.nome) + '</strong>: ' + esc(d.puntiNonLocalizzati.join(' | ')); }), false);

    h += blocco('Data della seduta non rilevata', 'Conviene indicare la data nel nome del file, per esempio «VERBALE n.10 ... 23 aprile 2026.docx».',
      senzaData.map(function (d) { return esc(d.nome); }), false);

    h += blocco('Numero del verbale non rilevato', 'Conviene indicare il numero nel nome del file, per esempio «VERBALE n.10 ...».',
      senzaNumero.map(function (d) { return esc(d.nome); }), false);

    if (!errori.length && !scansioni.length && !senzaOdg.length && !nonLocalizzati.length && !senzaData.length && !senzaNumero.length) {
      h += '<div class="riquadro"><h3>Nessuna anomalia</h3><p>Tutti i verbali sono stati riconosciuti nella loro struttura: numero, data, punti all\'ordine del giorno e delibere.</p></div>';
    }

    h += '<div class="riquadro"><h3>Attendibilità del riconoscimento dei punti</h3>' +
      '<p>Per ciascun punto viene calcolato un grado di corrispondenza fra la dizione dell\'ordine del giorno e l\'intestazione trovata nel corpo del verbale. Un valore inferiore a 0,60 merita una verifica manuale.</p>' +
      '<div class="tabella-contenitore"><table><thead><tr><th class="num">Verbale</th><th class="num">Punto</th><th>Dizione nell\'o.d.g.</th><th>Intestazione nel testo</th><th class="num">Corrisp.</th></tr></thead><tbody>';
    var dubbi = 0;
    arr.forEach(function (d) {
      d.sezioni.forEach(function (s) {
        if (!s.punto || s.affidabilita == null || s.affidabilita >= 0.6) return;
        dubbi++;
        h += '<tr><td class="num">' + esc(d.numero) + '</td><td class="num">' + esc(s.punto) + '</td><td>' +
          esc(s.titolo) + '</td><td>' + esc(s.intestazione || '— non individuata —') + '</td><td class="num">' +
          esc(String(s.affidabilita).replace('.', ',')) + '</td></tr>';
      });
    });
    if (!dubbi) h += '<tr><td colspan="5"><em>Nessun punto con corrispondenza dubbia.</em></td></tr>';
    h += '</tbody></table></div></div>';

    el('vista-diagnostica').innerHTML = h;
  }

  /* ------------------------------------------------------------ caricamento */

  function caricaFiles(lista) {
    var files = Array.prototype.slice.call(lista).filter(function (f) {
      return /\.(docx|pdf|txt|doc)$/i.test(f.name);
    });
    if (!files.length) {
      alert('Nessun file utilizzabile. Sono ammessi i formati .docx, .pdf e .txt.');
      return Promise.resolve();
    }

    var fatti = 0, caricati = [], saltati = [];
    velo(true, 'Caricamento nell\'archivio', '', 0);

    var seq = Promise.resolve();
    files.forEach(function (f) {
      seq = seq.then(function () {
        velo(true, 'Caricamento nell\'archivio', f.name, (fatti / files.length) * 100);
        return f.arrayBuffer().then(function (buf) {
          return fetch('/api/upload?nome=' + encodeURIComponent(f.name), { method: 'POST', body: buf })
            .then(function (r) {
              if (r.status === 409) {
                velo(false);
                var sovrascrivi = confirm('Nell\'archivio esiste già un file di nome:\n\n' + f.name +
                  '\n\nSostituirlo con quello che si sta caricando?');
                velo(true, 'Caricamento nell\'archivio', f.name, (fatti / files.length) * 100);
                if (!sovrascrivi) { saltati.push(f.name); return null; }
                return fetch('/api/upload?nome=' + encodeURIComponent(f.name) + '&sovrascrivi=1', { method: 'POST', body: buf });
              }
              return r;
            })
            .then(function (r) {
              if (r && !r.ok) throw new Error('HTTP ' + r.status);
              if (r) caricati.push(f.name);
            });
        }).catch(function (e) {
          saltati.push(f.name + ' (' + (e.message || e) + ')');
        }).then(function () { fatti++; });
      });
    });

    return seq.then(function () {
      velo(false);
      return costruisciIndice(false);
    }).then(function () {
      ridisegnaTutto();
      var msg = caricati.length + (caricati.length === 1 ? ' verbale aggiunto' : ' verbali aggiunti') + ' all\'archivio e indicizzato.';
      if (saltati.length) msg += '\n\nNon acquisiti:\n· ' + saltati.join('\n· ');
      alert(msg);
    });
  }

  /* ---------------------------------------------------------------- stato */

  function aggiornaStato() {
    var c = stato.config || {};
    var n = docs().length;
    var ok = docs().filter(function (d) { return !d.errore && !d.avviso; }).length;
    var nd = 0;
    docs().forEach(function (d) { nd += (d.delibere || []).length; });
    var problemi = n - ok;

    var h = '';
    h += '<span class="pill' + (c.esiste ? '' : ' err') + '">' + (c.esiste ? 'archivio collegato' : 'archivio non trovato') + '</span>';
    h += '<span class="pill">' + n + ' verbali</span>';
    h += '<span class="pill">' + nd + ' delibere</span>';
    if (problemi > 0) h += '<span class="pill err">' + problemi + ' da verificare</span>';
    h += '<span class="percorso">' + esc(c.archivePath || '') + '</span>';
    el('stato-archivio').innerHTML = h;
  }

  function ridisegnaTutto() {
    aggiornaStato();
    if (el('q').value.trim()) cerca(); else disegnaRisultati(null);
    disegnaVerbali();
    disegnaDelibere(el('filtro-delibere') ? el('filtro-delibere').value : '');
    disegnaCestino();
    disegnaPresenze();
    disegnaCondivisione();
    disegnaDiagnostica();
  }

  var SEGNAPOSTO = {
    parole: 'Una, due o tre parole: palestra, stage linguistico, conto consuntivo…',
    frase: 'La frase esatta da cercare: impianti sportivi, conto consuntivo…',
    numero: 'Il numero: 10, verbale 10, delibera 96, d 96/2025-26…',
    data: 'La data: 23/04/2026, aprile 2026, 2026, a.s. 2025/26…'
  };

  function aggiornaModo() {
    var modo = document.querySelector('input[name=modo]:checked').value;
    document.body.classList.toggle('modo-numero', modo === 'numero');
    document.body.classList.toggle('modo-data', modo === 'data');
    el('q').placeholder = SEGNAPOSTO[modo] || SEGNAPOSTO.parole;
  }

  /* ==================================================================
     RACCOLTA DA CONDIVIDERE E STAMPA IN PDF

     Quanto puo' essere utile a terzi - il prospetto delle presenze, il
     testo di una delibera con lo sviluppo del punto all'ordine del
     giorno, gli stralci restituiti da una ricerca, l'ordine del giorno
     di una seduta - viene raccolto in un elenco e composto in un unico
     foglio, che si stampa oppure si salva in PDF dalla finestra di
     stampa del browser (voce «Salva come PDF» fra le destinazioni).
     ================================================================== */

  var CHIAVE_RACCOLTA = 'archivio-verbali-raccolta-1';
  var MARGINE_MM = 10;          /* margine della pagina, in millimetri */
  var SCALA_MINIMA = 0.34;      /* oltre questa riduzione il testo non si legge */
  var contatoreVoci = 0;

  stato.raccolta = [];
  stato.intestazione = { titolo: '', nota: '' };
  stato.stampa = null;          /* { voci, opzioni } in corso di anteprima */

  var MESI_IT = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
    'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'];

  function dataEstesa(d) {
    var x = d || new Date();
    return x.getDate() + ' ' + MESI_IT[x.getMonth()] + ' ' + x.getFullYear();
  }

  /* Data della seduta in forma compatta per le intestazioni di colonna:
     23/04/2026 diventa 23/04/26, che in otto punti tipografici sta. */
  function dataBreve(d) {
    var t = d.dataLabel || '';
    var m = t.match(/^(\d{2})\/(\d{2})\/(\d{2})(\d{2})$/);
    return m ? m[1] + '/' + m[2] + '/' + m[4] : t;
  }

  function messaggio(testo) {
    var box = el('messaggio');
    box.textContent = testo;
    box.hidden = false;
    clearTimeout(box._timer);
    box._timer = setTimeout(function () { box.hidden = true; }, 2600);
  }

  /* ------------------------------------------------ tenuta della raccolta */

  /* La raccolta sopravvive alla chiusura del browser: viene conservata
     nella memoria locale della pagina. Se questa non e' disponibile, la
     raccolta resta valida per la sola sessione in corso. */
  function caricaRaccolta() {
    try {
      var t = window.localStorage.getItem(CHIAVE_RACCOLTA);
      if (!t) return;
      var d = JSON.parse(t);
      if (d && d.voci) {
        stato.raccolta = d.voci;
        stato.intestazione = d.intestazione || { titolo: '', nota: '' };
        stato.raccolta.forEach(function (v) {
          var n = parseInt(String(v.id).replace(/\D/g, ''), 10);
          if (n > contatoreVoci) contatoreVoci = n;
        });
      }
    } catch (e) { /* memoria locale non disponibile: nessuna conseguenza */ }
  }

  function salvaRaccolta() {
    try {
      window.localStorage.setItem(CHIAVE_RACCOLTA, JSON.stringify({
        voci: stato.raccolta, intestazione: stato.intestazione
      }));
    } catch (e) { /* come sopra */ }
  }

  function aggiornaContatoreRaccolta() {
    var c = el('conta-raccolta');
    if (!c) return;
    c.textContent = stato.raccolta.length;
    c.hidden = stato.raccolta.length === 0;
  }

  /* Impronta di una voce, per non raccogliere due volte la stessa cosa. */
  function improntaVoce(v) {
    return [v.tipo, v.nome || '', v.sez == null ? '' : v.sez,
      v.idDelibera || '', v.firma || ''].join('|');
  }

  function aggiungiVoce(v, descrizione) {
    var imp = improntaVoce(v);
    var gia = stato.raccolta.some(function (x) { return improntaVoce(x) === imp; });
    if (gia && v.tipo !== 'nota' && v.tipo !== 'ricerca') {
      messaggio('Questo elemento è già nella raccolta.');
      return;
    }
    v.id = 'r' + (++contatoreVoci);
    stato.raccolta.push(v);
    salvaRaccolta();
    aggiornaContatoreRaccolta();
    disegnaCondivisione();
    messaggio((descrizione || 'Elemento aggiunto') + ' · ' + stato.raccolta.length +
      (stato.raccolta.length === 1 ? ' elemento nella raccolta' : ' elementi nella raccolta'));
  }

  function vocePerId(id) {
    for (var i = 0; i < stato.raccolta.length; i++) {
      if (stato.raccolta[i].id === id) return stato.raccolta[i];
    }
    return null;
  }

  function rimuoviVoce(id) {
    stato.raccolta = stato.raccolta.filter(function (v) { return v.id !== id; });
    salvaRaccolta();
    aggiornaContatoreRaccolta();
    disegnaCondivisione();
  }

  function spostaVoce(id, verso) {
    var i = -1;
    stato.raccolta.forEach(function (v, k) { if (v.id === id) i = k; });
    var j = i + verso;
    if (i < 0 || j < 0 || j >= stato.raccolta.length) return;
    var t = stato.raccolta[i];
    stato.raccolta[i] = stato.raccolta[j];
    stato.raccolta[j] = t;
    salvaRaccolta();
    disegnaCondivisione();
  }

  /* --------------------------------------------- comandi di raccolta */

  function docPerNome(nome) {
    return (stato.indice && stato.indice.docs) ? stato.indice.docs[nome] : null;
  }

  function contestoRicerca() {
    var r = stato.ultimaRicerca;
    if (!r || (r.modo !== 'parole' && r.modo !== 'frase')) return null;
    return { q: r.q, modo: r.modo, prefisso: !!r.prefisso, ambito: r.ambito };
  }

  function matchersDa(ric) {
    if (!ric || !ric.q) return null;
    var m = P.buildMatchers(ric.q, ric.modo, ric.prefisso);
    return (m && m.length) ? m : null;
  }

  function raccogliPunto(nome, sez, paragrafi) {
    var d = docPerNome(nome);
    if (!d) return;
    aggiungiVoce({
      tipo: 'punto', nome: nome, sez: sez,
      paragrafi: paragrafi || [],
      ricerca: contestoRicerca(),
      opz: { integrale: true, evidenzia: true }
    }, 'Punto aggiunto alla raccolta');
  }

  function raccogliDelibera(nome, sez, idDelibera) {
    aggiungiVoce({
      tipo: 'delibera', nome: nome, sez: sez, idDelibera: idDelibera,
      opz: { sviluppo: true }
    }, 'Delibera aggiunta alla raccolta');
  }

  function raccogliOdg(nome) {
    aggiungiVoce({ tipo: 'odg', nome: nome, opz: { delibere: true } },
      'Ordine del giorno aggiunto alla raccolta');
  }

  function raccogliPresenze() {
    aggiungiVoce({
      tipo: 'presenze',
      opz: { cessati: false, segnalazioni: true, totali: true, legenda: true }
    }, 'Prospetto delle presenze aggiunto alla raccolta');
  }

  /* L'esito di una ricerca viene fissato al momento della raccolta: i
     verbali possono essere in seguito aggiunti o rimossi, ma cio' che si
     e' deciso di condividere non deve cambiare da sé. */
  function raccogliRicerca() {
    if (!stato.risultati.length) { messaggio('Non vi sono risultati da raccogliere.'); return; }
    var ric = contestoRicerca();
    var voci = [];
    stato.risultati.forEach(function (g) {
      g.voci.forEach(function (voce) {
        voci.push({ nome: g.doc.nome, sez: voce.sezioneIdx, paragrafi: voce.paragrafi.slice() });
      });
    });
    aggiungiVoce({
      tipo: 'ricerca', ricerca: ric, elementi: voci,
      firma: (ric ? ric.q : '') + '#' + voci.length + '#' + Date.now(),
      opz: { integrale: false, evidenzia: true }
    }, 'Stralci della ricerca aggiunti alla raccolta');
  }

  /* ------------------------------------- composizione del foglio stampato */

  /* Il titolo del blocco si omette quando ripeterebbe quello che
     l'intestazione del foglio gia' riporta. */
  function titoloBlocco(testo, opzGen) {
    return (opzGen && opzGen.unico) ? '' : '<h2>' + testo + '</h2>';
  }

  function sezioneDi(d, si) {
    return (d && si != null && si >= 0 && d.sezioni[si]) ? d.sezioni[si] : null;
  }

  function riferimentoVerbale(d, sez) {
    var parti = [etichettaVerbale(d)];
    if (d.dataLabel) parti.push('seduta del ' + d.dataLabel);
    if (sez && sez.punto) parti.push('punto ' + sez.punto + ' all\'ordine del giorno');
    return parti.join(' · ');
  }

  /* Testo di una sezione, per intero oppure nei soli passaggi indicati,
     con i termini cercati evidenziati quando richiesto. */
  function testoPerStampa(d, si, paragrafi, matchers, integrale) {
    var out = [];
    if (integrale) {
      d.paragrafi.forEach(function (p) {
        if (p.s !== si) return;
        var cl = classeCapoverso(p.t);
        var html = matchers ? P.evidenzia(p.t, matchers, Infinity, true).html : esc(p.t);
        out.push('<p' + (cl ? ' class="' + cl + '"' : '') + '>' + html + '</p>');
      });
    } else {
      (paragrafi || []).forEach(function (idx) {
        var p = d.paragrafi[idx];
        if (!p) return;
        var html = matchers ? P.evidenzia(p.t, matchers, 320, true).html : esc(p.t);
        out.push('<p class="passaggio">' + html + '</p>');
      });
    }
    if (!out.length) out.push('<p><em>Testo non disponibile.</em></p>');
    return out.join('');
  }

  function bloccoPunto(v, opzGen) {
    var d = docPerNome(v.nome);
    if (!d) return '<div class="blocco"><p><em>Il verbale ' + esc(v.nome) +
      ' non è più presente nell\'archivio.</em></p></div>';
    var sez = sezioneDi(d, v.sez);
    var integrale = v.opz.integrale !== false;
    var matchers = (opzGen.evidenzia && v.opz.evidenzia !== false) ? matchersDa(v.ricerca) : null;

    var h = '<div class="blocco">';
    h += titoloBlocco((sez && sez.punto ? 'Punto ' + esc(sez.punto) + ' all\'ordine del giorno' : 'Stralcio del verbale'), opzGen);
    h += '<div class="riferimento">' + esc(riferimentoVerbale(d, null)) + '</div>';
    if (sez && (sez.titolo || sez.intestazione)) {
      h += '<p class="oggetto">' + esc(sez.titolo || sez.intestazione) + '</p>';
    }
    if (sez && sez.delibere && sez.delibere.length) {
      h += '<p class="riferimento">Delibera n. ' + esc(sez.delibere.join(', n. ')) + '</p>';
    }
    h += testoPerStampa(d, v.sez, v.paragrafi, matchers, integrale);
    return h + '</div>';
  }

  function bloccoDelibera(v, opzGen) {
    var d = docPerNome(v.nome);
    if (!d) return '<div class="blocco"><p><em>Il verbale ' + esc(v.nome) +
      ' non è più presente nell\'archivio.</em></p></div>';
    var sez = sezioneDi(d, v.sez);
    var del = null;
    ((sez && sez.dettaglioDelibere) || []).forEach(function (x) {
      if (x.id === v.idDelibera) del = x;
    });

    var h = '<div class="blocco">';
    h += titoloBlocco('Delibera n. ' + esc(del ? del.id : v.idDelibera), opzGen);
    h += '<div class="riferimento">' + esc(riferimentoVerbale(d, sez)) + '</div>';
    if (sez && sez.titolo) h += '<p class="oggetto">Oggetto: ' + esc(sez.titolo) + '</p>';
    h += formattaTestoDelibera(del ? del.testo : '');

    if (v.opz.sviluppo !== false && sez) {
      h += '<div class="sotto-titolo-blocco">Sviluppo del punto all\'ordine del giorno</div>';
      h += testoPerStampa(d, v.sez, null, null, true);
    }
    return h + '</div>';
  }

  function bloccoOdg(v, opzGen) {
    var d = docPerNome(v.nome);
    if (!d) return '<div class="blocco"><p><em>Il verbale ' + esc(v.nome) +
      ' non è più presente nell\'archivio.</em></p></div>';
    var perPunto = {};
    d.sezioni.forEach(function (s) { if (s.punto) perPunto[s.punto] = s; });

    var h = '<div class="blocco">';
    h += titoloBlocco('Ordine del giorno · ' + esc(etichettaVerbale(d)), opzGen);
    h += '<div class="riferimento">' +
      (d.dataLabel ? 'Seduta del ' + esc(d.dataLabel) : 'Data non rilevata') +
      ' · ' + (d.delibere || []).length +
      ((d.delibere || []).length === 1 ? ' delibera adottata' : ' delibere adottate') + '</div>';
    if (!d.odg.length) {
      h += '<p><em>L\'ordine del giorno non è stato riconosciuto in questo verbale.</em></p>';
      return h + '</div>';
    }
    h += '<ul class="odg-stampa">';
    d.odg.forEach(function (it) {
      var s = perPunto[it.num];
      h += '<li' + (it.level > 0 ? ' class="sotto"' : '') + '><strong>' + esc(it.num) + '.</strong> ' + esc(it.titolo);
      if (s && v.opz.delibere !== false) {
        (s.delibere || []).forEach(function (num) {
          h += ' <span class="segno">[delibera n. ' + esc(num) + ']</span>';
        });
        (s.esiti || []).forEach(function (e) {
          if (e === 'deliberato') return;
          h += ' <span class="segno">[' + esc(e) + ']</span>';
        });
      }
      h += '</li>';
    });
    h += '</ul>';
    return h + '</div>';
  }

  function bloccoRicerca(v, opzGen) {
    var ric = v.ricerca;
    var matchers = (opzGen.evidenzia && v.opz.evidenzia !== false) ? matchersDa(ric) : null;
    var integrale = v.opz.integrale === true;

    var h = '<div class="blocco">';
    h += titoloBlocco('Stralci dei verbali', opzGen);
    h += '<div class="riferimento">' +
      (ric ? 'Ricerca ' + (ric.modo === 'frase' ? 'della frase' : 'delle parole') +
        ' «' + esc(ric.q) + '»' : 'Selezione di passaggi') +
      ' · ' + v.elementi.length + (v.elementi.length === 1 ? ' punto' : ' punti') + '</div>';

    var perVerbale = [];
    var mappa = {};
    v.elementi.forEach(function (e) {
      if (!mappa[e.nome]) { mappa[e.nome] = { nome: e.nome, voci: [] }; perVerbale.push(mappa[e.nome]); }
      mappa[e.nome].voci.push(e);
    });

    perVerbale.forEach(function (g) {
      var d = docPerNome(g.nome);
      if (!d) {
        h += '<p><em>Il verbale ' + esc(g.nome) + ' non è più presente nell\'archivio.</em></p>';
        return;
      }
      h += '<div class="sotto-titolo-blocco">' + esc(etichettaVerbale(d)) +
        (d.dataLabel ? ' · seduta del ' + esc(d.dataLabel) : '') + '</div>';
      g.voci.forEach(function (e) {
        var sez = sezioneDi(d, e.sez);
        h += '<p class="oggetto">' +
          (sez && sez.punto ? 'Punto ' + esc(sez.punto) + ' o.d.g. — ' : '') +
          esc(sez ? (sez.titolo || sez.intestazione || 'parte accessoria') : 'parte accessoria') +
          (sez && sez.delibere && sez.delibere.length ?
            ' <span class="segno">[delibera n. ' + esc(sez.delibere.join(', n. ')) + ']</span>' : '') +
          '</p>';
        h += testoPerStampa(d, e.sez, e.paragrafi, matchers, integrale);
      });
    });
    return h + '</div>';
  }

  function bloccoNota(v) {
    var t = String(v.testo || '').trim();
    if (!t) return '';
    var h = '<div class="blocco">';
    t.split(/\n{2,}/).forEach(function (par, i) {
      var righe = par.split('\n').filter(function (r) { return r.trim(); });
      if (!righe.length) return;
      if (i === 0 && righe.length > 1 && righe[0].length < 90 && !/[.;:]$/.test(righe[0].trim())) {
        h += '<h2>' + esc(righe.shift().trim()) + '</h2>';
      }
      if (righe.length) h += '<p>' + righe.map(esc).join('<br>') + '</p>';
    });
    return h + '</div>';
  }

  /* ------------------------------------ prospetto delle presenze stampato */

  function bloccoPresenze(v, opzGen) {
    var p = calcolaPresenze();
    if (!p.conTabella.length) {
      return '<div class="blocco"><h2>Prospetto delle presenze</h2>' +
        '<p><em>In nessuno dei verbali indicizzati è stata riconosciuta la tabella dei componenti.</em></p></div>';
    }
    var opz = v.opz || {};
    var visibili = opz.cessati ? p.membri : p.attuali;
    var prima = p.conTabella[0], ultima = p.conTabella[p.conTabella.length - 1];

    var h = '<div class="blocco">';
    h += titoloBlocco('Prospetto delle presenze', opzGen);
    h += '<div class="riferimento">' + p.conTabella.length +
      (p.conTabella.length === 1 ? ' seduta' : ' sedute') +
      (prima.dataLabel && ultima.dataLabel ? ' · dal ' + esc(prima.dataLabel) + ' al ' + esc(ultima.dataLabel) : '') +
      ' · ' + visibili.length + (visibili.length === 1 ? ' componente' : ' componenti') + '</div>';

    h += '<div class="riquadro-tabella"><table class="prospetto"><thead><tr><th class="nome">Componente</th>';
    p.conTabella.forEach(function (d) {
      h += '<th class="seduta">' + (d.numero != null ? 'n. ' + esc(d.numero) : '—') +
        '<small>' + esc(dataBreve(d)) + '</small></th>';
    });
    h += '<th>Pres.</th><th>Ass.</th><th>Cons.</th>';
    if (opz.stato !== false) h += '<th>Stato</th>';
    h += '</tr></thead><tbody>';

    visibili.forEach(function (m) {
      h += '<tr><td class="nome">' + esc(m.nome);
      var sotto = [];
      if (m.ruolo) sotto.push(m.ruolo);
      if (m.cessato) sotto.push('cessato dal ' + (m.cessato.dataLabel || m.cessato.data));
      if (sotto.length) h += '<small>' + esc(sotto.join(' · ')) + '</small>';
      h += '</td>';
      p.conTabella.forEach(function (d) {
        var st = m.stati[d.nome];
        var corretta = statoPresenze().correzioni[d.nome + '||' + m.chiave] ? ' corretta' : '';
        var lab = st === 'presente' ? 'P' : (st === 'assente' ? 'A' : '–');
        var cl = st === 'presente' ? 'p' : (st === 'assente' ? 'a' : 'v');
        h += '<td class="' + cl + corretta + '">' + lab + '</td>';
      });
      h += '<td>' + m.totPresenze + '</td><td>' + m.totAssenze + '</td><td>' + (m.serie || '') + '</td>';
      if (opz.stato !== false) {
        var st = '—';
        if (m.livello === 'rosso') st = 'oltre 2 ass. consec.';
        else if (m.livello === 'arancione') st = '2 ass. consecutive';
        else if (m.livello === 'chiuso') st = 'decaduto';
        else if (m.livello === 'cessato') st = 'cessato';
        h += '<td class="stato-cella">' + esc(st) + '</td>';
      }
      h += '</tr>';
    });

    if (opz.totali !== false) {
      h += '<tr class="totali"><td class="nome">Presenti per seduta</td>';
      p.conTabella.forEach(function (d) {
        var n = 0;
        visibili.forEach(function (m) { if (m.stati[d.nome] === 'presente') n++; });
        h += '<td>' + n + '</td>';
      });
      h += '<td colspan="' + (opz.stato !== false ? 4 : 3) + '"></td></tr>';
    }
    h += '</tbody></table></div>';

    if (opz.legenda !== false) {
      h += '<div class="legenda"><strong>P</strong> presente · <strong>A</strong> assente · ' +
        '<strong>–</strong> non risulta negli elenchi di quella seduta · ' +
        '<strong>•</strong> dato corretto a mano rispetto all\'estrazione automatica. ' +
        'La colonna «Cons.» riporta le assenze consecutive in corso alla data dell\'ultima seduta.';
      if (p.senzaTabella.length) {
        h += '<br>' + p.senzaTabella.length +
          (p.senzaTabella.length === 1 ? ' verbale non ha' : ' verbali non hanno') +
          ' una tabella delle presenze riconosciuta e non compare nel prospetto: ' +
          esc(p.senzaTabella.map(etichettaSeduta).join(' · ')) + '.';
      }
      if (!opz.cessati && p.cessati.length) {
        h += '<br>' + p.cessati.length +
          (p.cessati.length === 1 ? ' componente cessato dal mandato non compare' : ' componenti cessati dal mandato non compaiono') +
          ' nel prospetto.';
      }
      h += '</div>';
    }

    if (opz.segnalazioni !== false) {
      var segnalati = p.membri.filter(function (m) {
        return (m.livello === 'rosso' || m.livello === 'arancione' || m.livello === 'chiuso') &&
          (opz.cessati || !m.cessato);
      });
      if (segnalati.length) {
        h += '<div class="segnalazioni"><strong>Assenze consecutive da segnalare</strong><ul>';
        segnalati.forEach(function (m) {
          var r = esc(m.nome) + ' — ';
          if (m.livello === 'rosso') r += m.serieMax + ' assenze consecutive';
          else if (m.livello === 'arancione') r += 'due assenze consecutive in corso';
          else r += 'decadenza registrata';
          if (m.serieRossa && m.serieRossa.length) {
            r += ' (sedute ' + esc(m.serieRossa.map(etichettaSeduta).join(' · ')) + ')';
          } else if (m.livello === 'arancione' && m.primaDellaSerie) {
            r += ' (dalla seduta ' + esc(etichettaSeduta(m.primaDellaSerie)) + ')';
          }
          if (m.decadenza) {
            r += ' · decadenza dichiarata con delibera n. ' + esc(m.decadenza.delibera) +
              (m.decadenza.data ? ' del ' + esc(m.decadenza.data) : '');
          }
          h += '<li>' + r + '.</li>';
        });
        h += '</ul></div>';
      }
    }
    return h + '</div>';
  }

  /* ------------------------------------------------- foglio completo */

  function bloccoVoce(v, opzGen) {
    if (v.tipo === 'presenze') return bloccoPresenze(v, opzGen);
    if (v.tipo === 'delibera') return bloccoDelibera(v, opzGen);
    if (v.tipo === 'punto') return bloccoPunto(v, opzGen);
    if (v.tipo === 'odg') return bloccoOdg(v, opzGen);
    if (v.tipo === 'ricerca') return bloccoRicerca(v, opzGen);
    if (v.tipo === 'nota') return bloccoNota(v);
    return '';
  }

  function componiFoglio(voci, opzGen) {
    var proprio = (stato.intestazione.titolo || '').trim();
    var titolo = proprio || titoloPredefinito(voci);
    var nota = (stato.intestazione.nota || '').trim();
    opzGen.unico = (voci.length === 1 && !proprio);

    var h = '<div class="doc-stampa"><div class="intestazione-foglio">';
    h += '<div class="ente">Liceo Scientifico Statale «Nicolò Palmeri» · Termini Imerese</div>';
    h += '<div class="organo">Consiglio d\'Istituto</div>';
    h += '<h1>' + esc(titolo) + '</h1>';
    h += '<div class="sottotitolo-foglio">Estratto dall\'archivio dei verbali · ' + esc(dataEstesa()) + '</div>';
    h += '</div>';
    if (nota) h += '<div class="nota-intestazione">' + esc(nota).replace(/\n/g, '<br>') + '</div>';

    voci.forEach(function (v) { h += bloccoVoce(v, opzGen); });

    h += '<div class="piede-foglio"><span>Archivio dei verbali del Consiglio d\'Istituto · ' +
      'documento generato il ' + esc(dataEstesa()) + '</span>' +
      '<span>Il documento riporta stralci dei verbali approvati; fa fede il testo originale.</span></div>';
    return h + '</div>';
  }

  function titoloPredefinito(voci) {
    if (voci.length === 1) {
      var v = voci[0];
      if (v.tipo === 'presenze') return 'Prospetto delle presenze';
      if (v.tipo === 'delibera') return 'Delibera n. ' + (v.idDelibera || '');
      if (v.tipo === 'odg') {
        var d = docPerNome(v.nome);
        return 'Ordine del giorno · ' + (d ? etichettaVerbale(d) : '');
      }
      if (v.tipo === 'ricerca') return 'Stralci dei verbali';
      if (v.tipo === 'punto') return 'Stralcio del verbale';
    }
    return 'Estratto dai verbali del Consiglio d\'Istituto';
  }

  /* --------------------------------------------- anteprima e stampa */

  /* Quanti pixel misura un millimetro sullo schermo: serve a confrontare
     l'altezza del foglio composto con quella della pagina stampabile. */
  function pixelPerMm() {
    var probe = document.createElement('div');
    probe.style.cssText = 'position:absolute;visibility:hidden;height:100mm;width:1mm';
    document.body.appendChild(probe);
    var r = probe.offsetHeight / 100;
    document.body.removeChild(probe);
    return r || 3.7795;
  }

  function regolaPagina(orientamento) {
    var s = document.getElementById('regola-pagina');
    if (!s) {
      s = document.createElement('style');
      s.id = 'regola-pagina';
      document.head.appendChild(s);
    }
    s.textContent = '@page { size: A4 ' +
      (orientamento === 'verticale' ? 'portrait' : 'landscape') +
      '; margin: ' + MARGINE_MM + 'mm; }';
  }

  /* Una tabella con molte sedute puo' eccedere la larghezza della pagina.
     Se ne riduce dapprima il corpo tipografico; se ancora non basta, si
     rimpicciolisce la sola tabella, lasciando intatto il resto del foglio:
     una riduzione dell'intero documento ne comprometterebbe l'impaginazione
     su piu' pagine. Restituisce la riduzione minima applicata. */
  function adattaTabelle(interno, disponibile) {
    var minimo = 1;
    Array.prototype.forEach.call(interno.querySelectorAll('table.prospetto'), function (t) {
      t.style.fontSize = '';
      t.style.transform = '';
      var corpo = 8;
      for (var i = 0; i < 14 && t.offsetWidth > disponibile + 1 && corpo > 4.4; i++) {
        corpo -= 0.3;
        t.style.fontSize = corpo.toFixed(1) + 'pt';
      }
      if (t.offsetWidth > disponibile + 1) {
        var k = disponibile / t.offsetWidth;
        t.style.transform = 'scale(' + k.toFixed(4) + ')';
        var cont = t.parentNode;
        if (cont && cont.className === 'riquadro-tabella') {
          cont.style.height = Math.ceil(t.offsetHeight * k) + 'px';
          cont.style.overflow = 'hidden';
        }
        if (k < minimo) minimo = k;
      }
    });
    return minimo;
  }

  function apriAnteprima(voci, opzioni) {
    if (!voci.length) { messaggio('Non vi è nulla da stampare.'); return; }
    stato.stampa = { voci: voci, opz: opzioni };
    el('st-orientamento').value = opzioni.orientamento || 'orizzontale';
    el('st-una-pagina').checked = opzioni.unaPagina !== false;
    el('st-evidenzia').checked = opzioni.evidenzia !== false;
    el('velo-stampa').hidden = false;
    document.body.classList.add('stampa-attiva');
    componiAnteprima();
  }

  function chiudiAnteprima() {
    el('velo-stampa').hidden = true;
    document.body.classList.remove('stampa-attiva');
    stato.stampa = null;
  }

  function componiAnteprima() {
    if (!stato.stampa) return;
    var o = stato.stampa.opz;
    o.orientamento = el('st-orientamento').value;
    o.unaPagina = el('st-una-pagina').checked;
    o.evidenzia = el('st-evidenzia').checked;
    regolaPagina(o.orientamento);

    var larghMm = (o.orientamento === 'verticale' ? 210 : 297) - 2 * MARGINE_MM;
    var altMm = (o.orientamento === 'verticale' ? 297 : 210) - 2 * MARGINE_MM;
    var mm = pixelPerMm();

    var foglio = el('foglio'), interno = el('foglio-interno');
    foglio.style.width = larghMm + 'mm';
    foglio.style.height = '';
    foglio.style.minHeight = altMm + 'mm';
    interno.style.width = larghMm + 'mm';
    interno.style.transform = 'none';
    interno.innerHTML = componiFoglio(stato.stampa.voci, o);

    var avviso = el('st-avviso');
    avviso.textContent = '';

    /* Un paio di pixel di margine: il foglio non deve traboccare sulla
       pagina seguente per un arrotondamento. */
    var spazioAlto = altMm * mm - 3;
    var spazioLargo = larghMm * mm;
    var kTabella = adattaTabelle(interno, spazioLargo);

    var alt = interno.scrollHeight;
    var k = 1;
    if (o.unaPagina && alt > spazioAlto) k = spazioAlto / alt;

    if (k < SCALA_MINIMA) {
      k = SCALA_MINIMA;
      avviso.textContent = 'Il contenuto eccede la pagina anche alla riduzione massima: ' +
        'converrà l\'orientamento orizzontale, oppure meno elementi nella raccolta.';
    }

    /* La riduzione si applica al solo caso della pagina unica: in un
       documento di piu' pagine il contenuto rimpicciolito non si
       ripartirebbe correttamente fra le pagine. */
    foglio.classList.toggle('ritagliato', k < 1);
    if (k < 1) {
      interno.style.transform = 'scale(' + k.toFixed(4) + ')';
      foglio.style.height = Math.ceil(alt * k) + 'px';
      foglio.style.minHeight = '0';
      if (!avviso.textContent && k < 0.8) {
        avviso.textContent = 'Il contenuto è stato ridotto al ' + Math.round(k * 100) +
          '% per rientrare in una pagina.';
      }
    } else if (o.unaPagina) {
      foglio.style.height = Math.ceil(spazioAlto) + 'px';
    } else {
      foglio.style.height = '';
      foglio.style.minHeight = altMm + 'mm';
      /* Stima per difetto: un blocco che non sta nello spazio residuo passa
         intero alla pagina seguente, e le pagine possono risultare di piu'. */
      var pagine = Math.max(1, Math.ceil(alt / spazioAlto - 0.02));
      if (pagine > 1) avviso.textContent = 'Il documento occuperà all\'incirca ' + pagine + ' pagine.';
    }

    if (kTabella < 1 && !avviso.textContent) {
      avviso.textContent = 'Il prospetto è stato ridotto al ' + Math.round(kTabella * 100) +
        '% per rientrare nella larghezza della pagina: l\'orientamento orizzontale lo rende più leggibile.';
    }
  }

  /* Il titolo della pagina diventa il nome proposto dal browser per il
     file PDF: conviene che sia quello del documento composto. */
  function stampaFoglio() {
    if (!stato.stampa) return;
    var titoloPrecedente = document.title;
    var t = (stato.intestazione.titolo || '').trim() || titoloPredefinito(stato.stampa.voci);
    document.title = t + ' - Consiglio d\'Istituto - Liceo Palmeri';
    function ripristina() {
      document.title = titoloPrecedente;
      window.removeEventListener('afterprint', ripristina);
    }
    window.addEventListener('afterprint', ripristina);
    setTimeout(function () { window.print(); }, 60);
    setTimeout(ripristina, 8000);
  }

  /* Stampa immediata del solo prospetto delle presenze, senza passare
     dalla raccolta: e' l'uso piu' frequente. */
  function stampaProspettoPresenze() {
    var p = calcolaPresenze();
    if (!p.conTabella.length) {
      alert('Non risulta alcuna tabella delle presenze nei verbali indicizzati: non vi è prospetto da stampare.');
      return;
    }
    apriAnteprima([{
      id: 'tmp', tipo: 'presenze',
      opz: { cessati: !!stato.mostraCessati, segnalazioni: true, totali: true, legenda: true }
    }], { orientamento: 'orizzontale', unaPagina: true, evidenzia: false });
  }

  /* ------------------------------------------- scheda «Da condividere» */

  function descriviVoce(v) {
    var d = v.nome ? docPerNome(v.nome) : null;
    var sez = d ? sezioneDi(d, v.sez) : null;
    if (v.tipo === 'presenze') {
      return { titolo: 'Prospetto delle presenze',
        sommario: 'Tabella dei componenti per seduta, con le assenze consecutive da segnalare.' };
    }
    if (v.tipo === 'delibera') {
      return { titolo: 'Delibera n. ' + (v.idDelibera || ''),
        sommario: (d ? riferimentoVerbale(d, sez) : v.nome) + (sez && sez.titolo ? ' — ' + sez.titolo : '') };
    }
    if (v.tipo === 'punto') {
      return { titolo: (sez && sez.punto ? 'Punto ' + sez.punto + ' o.d.g. — ' : 'Stralcio — ') +
          (sez ? (sez.titolo || sez.intestazione || 'parte accessoria') : ''),
        sommario: d ? riferimentoVerbale(d, null) : v.nome };
    }
    if (v.tipo === 'odg') {
      return { titolo: 'Ordine del giorno · ' + (d ? etichettaVerbale(d) : v.nome),
        sommario: d && d.dataLabel ? 'Seduta del ' + d.dataLabel : v.nome };
    }
    if (v.tipo === 'ricerca') {
      return { titolo: 'Stralci della ricerca' + (v.ricerca ? ' «' + v.ricerca.q + '»' : ''),
        sommario: v.elementi.length + (v.elementi.length === 1 ? ' punto' : ' punti') +
          ' di ' + contaVerbaliRicerca(v) + ' verbali, come restituiti dalla ricerca.' };
    }
    if (v.tipo === 'nota') {
      return { titolo: 'Nota libera', sommario: 'Testo redatto a mano, riportato nel documento.' };
    }
    return { titolo: '—', sommario: '' };
  }

  function contaVerbaliRicerca(v) {
    var s = {}, n = 0;
    v.elementi.forEach(function (e) { if (!s[e.nome]) { s[e.nome] = 1; n++; } });
    return n;
  }

  function disegnaCondivisione() {
    var vista = el('vista-condivisione');
    if (!vista) return;
    aggiornaContatoreRaccolta();

    var h = '';
    h += '<div class="riquadro"><h3>Documento da condividere</h3>' +
      '<p>Gli elementi raccolti nelle altre schede vengono composti in un unico documento, ' +
      'che si stampa oppure si salva in PDF scegliendo «Salva come PDF» fra le destinazioni ' +
      'della finestra di stampa.</p>' +
      '<div class="intestazione-raccolta">' +
      '<div class="campo-largo"><label for="int-titolo">Titolo del documento</label>' +
      '<input id="int-titolo" type="text" placeholder="' + esc(titoloPredefinito(stato.raccolta)) + '" value="' +
      esc(stato.intestazione.titolo || '') + '"></div>' +
      '<div class="campo-largo"><label for="int-nota">Nota in apertura (facoltativa)</label>' +
      '<input id="int-nota" type="text" placeholder="Si trasmette, per quanto di competenza, …" value="' +
      esc(stato.intestazione.nota || '') + '"></div>' +
      '</div></div>';

    if (!stato.raccolta.length) {
      h += '<div class="vuoto"><span class="grande">La raccolta è vuota</span>' +
        'Il comando <strong>Aggiungi alla raccolta</strong> compare accanto a ciascun risultato di ricerca, ' +
        'a ciascuna delibera, all\'ordine del giorno di ogni verbale e al prospetto delle presenze.<br>' +
        'Gli elementi raccolti si ritrovano qui, in questa scheda, e vi restano anche dopo la chiusura del programma.</div>';
      h += '<div class="azioni-punto"><button class="bottone minuto" data-aggiungi-nota="1">Aggiungi una nota libera</button>' +
        '<button class="bottone minuto" data-raccogli-presenze="1">Aggiungi il prospetto delle presenze</button></div>';
      vista.innerHTML = h;
      agganciaCampiRaccolta();
      return;
    }

    h += '<div class="riepilogo">' + stato.raccolta.length +
      (stato.raccolta.length === 1 ? ' elemento nella raccolta' : ' elementi nella raccolta') +
      '. L\'ordine è quello del documento.</div>';

    stato.raccolta.forEach(function (v, i) {
      var des = descriviVoce(v);
      h += '<div class="voce-raccolta">';
      h += '<div class="riga-ordine"><span class="indice-voce">' + (i + 1) + '.</span>' +
        '<button class="bottone minuto" data-su="' + esc(v.id) + '"' + (i === 0 ? ' disabled' : '') + ' title="Sposta in su">↑</button>' +
        '<button class="bottone minuto" data-giu="' + esc(v.id) + '"' + (i === stato.raccolta.length - 1 ? ' disabled' : '') + ' title="Sposta in giù">↓</button>' +
        '<button class="bottone minuto pericolo" data-togli="' + esc(v.id) + '">Togli</button>';
      if (v.nome) {
        h += '<button class="bottone minuto" data-apri="' + esc(v.nome) + '">Apri il verbale</button>';
      }
      h += '</div>';
      h += '<div class="titolo-voce">' + esc(des.titolo) + '</div>';
      h += '<div class="sommario-voce">' + esc(des.sommario) + '</div>';

      if (v.tipo === 'nota') {
        h += '<textarea data-testo-nota="' + esc(v.id) + '" placeholder="Prima riga: titolo del paragrafo. Righe successive: il testo.">' +
          esc(v.testo || '') + '</textarea>';
      }
      h += '<div class="opzioni-voce">';
      if (v.tipo === 'punto' || v.tipo === 'ricerca') {
        h += '<label><input type="checkbox" data-opz="integrale" data-voce="' + esc(v.id) + '"' +
          (v.opz.integrale ? ' checked' : '') + '> testo integrale ' +
          (v.tipo === 'ricerca' ? 'dei punti' : 'del punto') + ', non i soli passaggi trovati</label>';
        if (v.ricerca) {
          h += '<label><input type="checkbox" data-opz="evidenzia" data-voce="' + esc(v.id) + '"' +
            (v.opz.evidenzia !== false ? ' checked' : '') + '> evidenzia i termini cercati</label>';
        }
      }
      if (v.tipo === 'delibera') {
        h += '<label><input type="checkbox" data-opz="sviluppo" data-voce="' + esc(v.id) + '"' +
          (v.opz.sviluppo !== false ? ' checked' : '') + '> con lo sviluppo del punto all\'ordine del giorno</label>';
      }
      if (v.tipo === 'odg') {
        h += '<label><input type="checkbox" data-opz="delibere" data-voce="' + esc(v.id) + '"' +
          (v.opz.delibere !== false ? ' checked' : '') + '> indica le delibere adottate su ciascun punto</label>';
      }
      if (v.tipo === 'presenze') {
        h += '<label><input type="checkbox" data-opz="cessati" data-voce="' + esc(v.id) + '"' +
          (v.opz.cessati ? ' checked' : '') + '> comprendi i componenti cessati dal mandato</label>';
        h += '<label><input type="checkbox" data-opz="segnalazioni" data-voce="' + esc(v.id) + '"' +
          (v.opz.segnalazioni !== false ? ' checked' : '') + '> elenca le assenze consecutive da segnalare</label>';
        h += '<label><input type="checkbox" data-opz="totali" data-voce="' + esc(v.id) + '"' +
          (v.opz.totali !== false ? ' checked' : '') + '> riga dei presenti per seduta</label>';
        h += '<label><input type="checkbox" data-opz="stato" data-voce="' + esc(v.id) + '"' +
          (v.opz.stato !== false ? ' checked' : '') + '> colonna dello stato</label>';
      }
      h += '</div></div>';
    });

    h += '<div class="azioni-punto" style="margin-top:14px">' +
      '<button class="bottone primario" data-anteprima="1">Anteprima e stampa (PDF)</button>' +
      '<button class="bottone minuto" data-aggiungi-nota="1">Aggiungi una nota libera</button>' +
      '<button class="bottone minuto" data-raccogli-presenze="1">Aggiungi il prospetto delle presenze</button>' +
      '<button class="bottone minuto pericolo" data-svuota-raccolta="1">Svuota la raccolta</button>' +
      '</div>';

    vista.innerHTML = h;
    agganciaCampiRaccolta();
  }

  /* I campi di testo si aggiornano senza ridisegnare la scheda, perché il
     ridisegno interromperebbe la digitazione. */
  function agganciaCampiRaccolta() {
    var t = el('int-titolo'), n = el('int-nota');
    if (t) t.addEventListener('input', function () { stato.intestazione.titolo = this.value; salvaRaccolta(); });
    if (n) n.addEventListener('input', function () { stato.intestazione.nota = this.value; salvaRaccolta(); });
    Array.prototype.forEach.call(document.querySelectorAll('[data-testo-nota]'), function (ta) {
      ta.addEventListener('input', function () {
        var v = vocePerId(this.dataset.testoNota);
        if (v) { v.testo = this.value; salvaRaccolta(); }
      });
    });
  }

  function anteprimaRaccolta() {
    if (!stato.raccolta.length) { messaggio('La raccolta è vuota.'); return; }
    var soloPresenze = stato.raccolta.length === 1 && stato.raccolta[0].tipo === 'presenze';
    /* Il prospetto delle presenze e' largo: quando fa parte della raccolta
       si propone l'orientamento orizzontale, che lo rende leggibile. */
    var conProspetto = stato.raccolta.some(function (v) { return v.tipo === 'presenze'; });
    apriAnteprima(stato.raccolta, {
      orientamento: conProspetto ? 'orizzontale' : 'verticale',
      unaPagina: soloPresenze,
      evidenzia: true
    });
  }

  /* Eventi della raccolta e dell'anteprima: si agganciano una sola volta. */
  function agganciaEventiStampa() {
    el('st-orientamento').addEventListener('change', componiAnteprima);
    el('st-una-pagina').addEventListener('change', componiAnteprima);
    el('st-evidenzia').addEventListener('change', componiAnteprima);
    el('st-stampa').addEventListener('click', stampaFoglio);
    el('st-chiudi').addEventListener('click', chiudiAnteprima);
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && !el('velo-stampa').hidden) chiudiAnteprima();
    });

    document.addEventListener('click', function (e) {
      var sel = '[data-raccogli-punto],[data-raccogli-delibera],[data-raccogli-odg],' +
        '[data-raccogli-presenze],[data-raccogli-ricerca],[data-stampa-presenze],' +
        '[data-su],[data-giu],[data-togli],[data-anteprima],[data-svuota-raccolta],[data-aggiungi-nota]';
      var b = e.target.closest ? e.target.closest(sel) : null;
      if (!b) return;
      e.stopPropagation();

      if (b.dataset.raccogliPunto) {
        var par = [];
        try { par = JSON.parse(b.dataset.paragrafi || '[]'); } catch (x) { par = []; }
        raccogliPunto(b.dataset.raccogliPunto, parseInt(b.dataset.sez, 10), par);
      } else if (b.dataset.raccogliDelibera) {
        raccogliDelibera(b.dataset.raccogliDelibera, parseInt(b.dataset.sez, 10), b.dataset.delibera);
      } else if (b.dataset.raccogliOdg) {
        raccogliOdg(b.dataset.raccogliOdg);
      } else if (b.dataset.raccogliPresenze) {
        raccogliPresenze();
      } else if (b.dataset.raccogliRicerca) {
        raccogliRicerca();
      } else if (b.dataset.stampaPresenze) {
        stampaProspettoPresenze();
      } else if (b.dataset.su) {
        spostaVoce(b.dataset.su, -1);
      } else if (b.dataset.giu) {
        spostaVoce(b.dataset.giu, 1);
      } else if (b.dataset.togli) {
        rimuoviVoce(b.dataset.togli);
      } else if (b.dataset.anteprima) {
        anteprimaRaccolta();
      } else if (b.dataset.aggiungiNota) {
        aggiungiVoce({ tipo: 'nota', testo: '', firma: 'n' + Date.now(), opz: {} }, 'Nota aggiunta');
      } else if (b.dataset.svuotaRaccolta) {
        if (!confirm('Togliere dalla raccolta tutti gli elementi?')) return;
        stato.raccolta = [];
        salvaRaccolta();
        aggiornaContatoreRaccolta();
        disegnaCondivisione();
      }
    });

    document.addEventListener('change', function (e) {
      var t = e.target;
      if (!t || !t.dataset || !t.dataset.opz) return;
      var v = vocePerId(t.dataset.voce);
      if (!v) return;
      v.opz[t.dataset.opz] = t.checked;
      salvaRaccolta();
      disegnaCondivisione();
    });
  }

  /* Comando di raccolta da apporre alle schede dei risultati. */
  function bottoneRaccolta(attributi, etichetta) {
    return '<button class="bottone minuto" ' + attributi + '>' + etichetta + '</button>';
  }


  /* ------------------------------------------------------------- eventi */

  function agganciaEventi() {
    var timer = null;
    el('q').addEventListener('input', function () {
      el('btn-pulisci').hidden = !this.value;
      clearTimeout(timer);
      timer = setTimeout(cerca, 220);
    });
    el('q').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { clearTimeout(timer); cerca(); }
    });
    el('btn-cerca').addEventListener('click', cerca);
    el('btn-pulisci').addEventListener('click', function () {
      el('q').value = ''; this.hidden = true; stato.ultimaRicerca = null; disegnaRisultati(null); el('q').focus();
    });

    ['opt-prefisso', 'opt-ambito', 'opt-ordine', 'opt-solo-delibere'].forEach(function (id) {
      el(id).addEventListener('change', cerca);
    });
    Array.prototype.forEach.call(document.querySelectorAll('input[name=modo]'), function (r) {
      r.addEventListener('change', function () { aggiornaModo(); cerca(); });
    });

    // schede
    Array.prototype.forEach.call(document.querySelectorAll('.scheda'), function (b) {
      b.addEventListener('click', function () {
        Array.prototype.forEach.call(document.querySelectorAll('.scheda'), function (x) { x.classList.remove('attiva'); });
        b.classList.add('attiva');
        ['risultati', 'verbali', 'delibere', 'presenze', 'condivisione', 'cestino', 'diagnostica'].forEach(function (v) {
          el('vista-' + v).hidden = (v !== b.dataset.scheda);
        });
      });
    });

    // menu
    el('btn-altro').addEventListener('click', function (e) {
      e.stopPropagation();
      el('menu-altro').hidden = !el('menu-altro').hidden;
    });
    document.addEventListener('click', function () { el('menu-altro').hidden = true; });
    el('menu-altro').addEventListener('click', function (e) {
      var a = e.target.dataset.azione;
      if (!a) return;
      el('menu-altro').hidden = true;
      if (a === 'reindicizza') {
        if (!confirm('Ricostruire l\'indice leggendo da capo tutti i verbali? L\'operazione può richiedere qualche minuto.')) return;
        costruisciIndice(true).then(ridisegnaTutto);
      } else if (a === 'cartella') {
        fetch('/api/apri', { method: 'POST' });
      } else if (a === 'esporta') {
        var blob = new Blob([csvDelibere()], { type: 'text/csv;charset=utf-8' });
        var url = URL.createObjectURL(blob);
        var link = document.createElement('a');
        link.href = url; link.download = 'Indice delibere Consiglio d\'Istituto.csv';
        document.body.appendChild(link); link.click(); document.body.removeChild(link);
        setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
      } else if (a === 'stampa-presenze') {
        stampaProspettoPresenze();
      } else if (a === 'raccolta') {
        var sc = document.querySelector('.scheda[data-scheda="condivisione"]');
        if (sc) { sc.click(); sc.scrollIntoView({ block: 'nearest' }); }
      } else if (a === 'impostazioni') {
        if (window.ArchivioGitHub) window.ArchivioGitHub.impostazioni();
      }
    });

    el('btn-aggiorna').addEventListener('click', function () {
      costruisciIndice(false).then(function (n) {
        ridisegnaTutto();
        if (!n) alert('Nessun verbale nuovo o modificato: l\'indice è già aggiornato.');
      });
    });

    // caricamento
    el('btn-carica').addEventListener('click', function () { el('file-input').click(); });
    el('file-input').addEventListener('change', function () {
      if (this.files && this.files.length) caricaFiles(this.files);
      this.value = '';
    });

    var contatoreDrag = 0;
    window.addEventListener('dragenter', function (e) {
      if (!e.dataTransfer || Array.prototype.indexOf.call(e.dataTransfer.types || [], 'Files') < 0) return;
      contatoreDrag++; el('zona-trascinamento').hidden = false;
    });
    window.addEventListener('dragleave', function () {
      contatoreDrag--; if (contatoreDrag <= 0) { contatoreDrag = 0; el('zona-trascinamento').hidden = true; }
    });
    window.addEventListener('dragover', function (e) { e.preventDefault(); });
    window.addEventListener('drop', function (e) {
      e.preventDefault(); contatoreDrag = 0; el('zona-trascinamento').hidden = true;
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) caricaFiles(e.dataTransfer.files);
    });

    // azioni delegate
    document.addEventListener('click', function (e) {
      var sel = '[data-apri],[data-cartella],[data-riga],[data-rimuovi],[data-ripristina],[data-elimina],' +
        '[data-apri-cestino],[data-espandi],[data-cella],[data-silenzia],[data-apri-decadenza],' +
        '[data-salva-decadenza],[data-revoca-decadenza],[data-scheda-vai],' +
        '[data-apri-cessazione],[data-salva-cessazione],[data-revoca-cessazione],' +
        '[data-apri-esclusione],[data-salva-esclusione],[data-reintegra]';
      var b = e.target.closest ? e.target.closest(sel) : null;
      if (!b) return;

      if (b.dataset.schedaVai) {
        var t = document.querySelector('.scheda[data-scheda="' + b.dataset.schedaVai + '"]');
        if (t) { t.click(); t.scrollIntoView({ block: 'nearest' }); }
        return;
      }
      if (b.dataset.apriCessazione) {
        var campoData = el('data-cessazione');
        stato.bozzaDataCessazione = campoData ? campoData.value : '';
        stato.pannelloCessazione = !stato.pannelloCessazione;
        disegnaPresenze();
        var nuovo = el('data-cessazione');
        if (nuovo) nuovo.focus();
        return;
      }
      if (b.dataset.salvaCessazione) { registraCessazione(); return; }
      if (b.dataset.revocaCessazione) { revocaCessazione(b.dataset.revocaCessazione); return; }
      if (b.dataset.apriEsclusione) {
        stato.pannelloEsclusione = !stato.pannelloEsclusione;
        disegnaPresenze();
        return;
      }
      if (b.dataset.salvaEsclusione) { registraEsclusione(); return; }
      if (b.dataset.reintegra) { reintegraEscluso(parseInt(b.dataset.reintegra, 10)); return; }
      if (b.dataset.cella) { commutaCella(b.dataset.cella, b.dataset.chiave); return; }
      if (b.dataset.silenzia) { commutaSilenzio(b.dataset.silenzia); return; }
      if (b.dataset.revocaDecadenza) { revocaDecadenza(b.dataset.revocaDecadenza); return; }
      if (b.dataset.apriDecadenza) {
        var box = document.getElementById('dec-' + b.dataset.apriDecadenza.replace(/[^a-z0-9]/g, '-'));
        if (box) { box.hidden = !box.hidden; if (!box.hidden) { var i = box.querySelector('input'); if (i) i.focus(); } }
        return;
      }
      if (b.dataset.salvaDecadenza) {
        var cont = b.closest('.modulo-decadenza');
        var del = cont ? cont.querySelector('[data-campo=delibera]').value.trim() : '';
        var dat = cont ? cont.querySelector('[data-campo=data]').value.trim() : '';
        registraDecadenza(b.dataset.salvaDecadenza, del, dat);
        return;
      }

      if (b.dataset.espandi) {
        e.stopPropagation();
        commutaEspansione(b);
      } else if (b.dataset.apri) {
        fetch('/api/apri?nome=' + encodeURIComponent(b.dataset.apri), { method: 'POST' });
      } else if (b.dataset.cartella) {
        fetch('/api/apri?nome=' + encodeURIComponent(b.dataset.cartella) + '&cartella=1', { method: 'POST' });
      } else if (b.dataset.apriCestino) {
        fetch('/api/apri?da=cestino&nome=' + encodeURIComponent(b.dataset.apriCestino), { method: 'POST' });
      } else if (b.dataset.rimuovi) {
        e.stopPropagation();
        rimuoviVerbale(b.dataset.rimuovi);
      } else if (b.dataset.ripristina) {
        ripristinaVerbale(b.dataset.ripristina, b.dataset.orig || b.dataset.ripristina, false);
      } else if (b.dataset.elimina) {
        eliminaDalCestino(b.dataset.elimina, b.dataset.orig || b.dataset.elimina);
      } else if (b.dataset.riga != null) {
        var det = document.querySelector('[data-dettaglio="' + b.dataset.riga + '"]');
        if (det) det.hidden = !det.hidden;
      }
    });

    // casella che riporta in vista i componenti cessati
    document.addEventListener('change', function (e) {
      if (e.target && e.target.id === 'opt-mostra-cessati') {
        stato.mostraCessati = e.target.checked;
        disegnaPresenze();
      }
    });

    document.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') { e.preventDefault(); el('q').focus(); el('q').select(); }
    });
  }

  /* --------------------------------------------------------------- avvio */

  function avvia() {
    if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
    agganciaEventi();
    agganciaEventiStampa();
    caricaRaccolta();
    aggiornaContatoreRaccolta();
    aggiornaModo();

    chiediJson('/api/config').then(function (c) {
      stato.config = c || {};
      if (!stato.config.esiste) {
        el('stato-archivio').innerHTML = '<span class="pill err">cartella non trovata</span>' +
          '<span class="percorso">' + esc(stato.config.archivePath || '') + '</span>';
        el('vista-risultati').innerHTML = '<div class="riquadro attenzione"><h3>La cartella dei verbali non è stata trovata</h3>' +
          '<p>Il percorso configurato è:</p><p class="mono"><strong>' + esc(stato.config.archivePath) + '</strong></p>' +
          '<p>Occorre aprire il file <strong>config.json</strong> nella cartella dell\'applicazione, ' +
          'correggere la voce <em>ArchivePath</em> con il percorso effettivo e riavviare l\'archivio. ' +
          'Nel percorso le barre rovesciate vanno raddoppiate, per esempio: ' +
          '<span class="mono">C:\\\\Users\\\\giann\\\\Desktop\\\\...</span></p></div>';
        return null;
      }
      return chiediJson('/api/indice');
    }).then(function (idx) {
      if (idx === null && (!stato.config || !stato.config.esiste)) return;
      if (idx && idx.versione === VERSIONE_INDICE) stato.indice = idx;
      return costruisciIndice(false)
        .then(function () { return aggiornaCestino(); })
        .then(function () {
          return chiediJson('/api/presenze').then(function (pr) {
            if (pr && pr.versione === VERSIONE_PRESENZE) stato.presenzeDati = pr;
          }).catch(function () { });
        })
        .then(ridisegnaTutto);
    }).catch(function (e) {
      velo(false);
      el('stato-archivio').innerHTML = '<span class="pill err">errore di avvio</span><span class="percorso">' + esc(e.message) + '</span>';
      console.error(e);
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', avvia);
  else avvia();
})();
