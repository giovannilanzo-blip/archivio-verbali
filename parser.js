/* =======================================================================
   parser.js — Estrazione e strutturazione dei verbali del Consiglio
   d'Istituto. Funziona sia nel browser sia in Node (per i test).
   ======================================================================= */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.VerbaliParser = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------------------------------------------------------------- utils */

  function unescapeXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
            .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(+d); })
            .replace(/&amp;/g, '&');
  }

  /* Normalizzazione per la ricerca (minuscole, senza diacritici, apostrofi e
     spaziatura uniformi) con mappa posizionale verso il testo originale: serve
     per evidenziare le occorrenze nel testo così come è scritto nel verbale. */
  var APOSTROFI = '‘’ʼ´`';
  var VIRGOLETTE = '“”«»';
  var LINEETTE = '–—−';

  function normWithMap(s) {
    var out = '', map = [], prevSpace = true;
    if (!s) return { n: '', map: map };
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (APOSTROFI.indexOf(ch) >= 0) ch = "'";
      else if (VIRGOLETTE.indexOf(ch) >= 0) ch = '"';
      else if (LINEETTE.indexOf(ch) >= 0) ch = '-';
      else if (ch === ' ' || ch === ' ' || ch === ' ') ch = ' ';

      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === '\f' || ch === '\v') {
        if (prevSpace) continue;
        prevSpace = true; out += ' '; map.push(i); continue;
      }
      prevSpace = false;
      var d = ch.normalize ? ch.normalize('NFD').replace(/[̀-ͯ]/g, '') : ch;
      d = d.toLowerCase();
      for (var k = 0; k < d.length; k++) { out += d.charAt(k); map.push(i); }
    }
    // rimuove lo spazio finale eventualmente accumulato
    while (out.length && out.charAt(out.length - 1) === ' ') { out = out.slice(0, -1); map.pop(); }
    return { n: out, map: map };
  }

  function norm(s) { return normWithMap(s).n; }

  var STOPWORDS = ('di a da in con su per tra fra del dello della dei degli delle dal dalla al alla allo ai agli alle ' +
    'nel nella nei negli nelle sul sulla sui sugli sulle il lo la i gli le un uno una che non come anche ' +
    'della dell alla all sui nell cui suo sua loro alcuni altre altri e o ed ad se si sono stato stata ' +
    'presso circa quale quali ogni presente presenti seguenti seguente').split(/\s+/);
  var STOPSET = {};
  STOPWORDS.forEach(function (w) { STOPSET[w] = true; });

  function tokens(s) {
    var out = [], seen = {};
    norm(s).split(/[^a-z0-9]+/).forEach(function (w) {
      if (w.length < 4 && !/^\d{4}$/.test(w)) return;
      if (STOPSET[w]) return;
      if (!seen[w]) { seen[w] = true; out.push(w); }
    });
    return out;
  }

  function similarity(a, b) {
    var ta = tokens(a), tb = tokens(b);
    if (!ta.length || !tb.length) return { score: 0, shared: 0 };
    var set = {}, shared = 0;
    tb.forEach(function (w) { set[w] = true; });
    ta.forEach(function (w) { if (set[w]) shared++; });
    var contA = shared / ta.length;
    var contB = shared / tb.length;
    var score = Math.max(contA, tb.length >= 3 ? contB : 0);
    return { score: score, shared: shared };
  }

  /* ------------------------------------------------- estrazione da DOCX */

  var RE_P = /<w:p(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/w:p>)/g;
  var RE_T = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;

  function docxToParagraphs(xml) {
    var body = xml;
    var mb = /<w:body[^>]*>([\s\S]*)<\/w:body>/.exec(xml);
    if (mb) body = mb[1];
    // rimuove il testo cancellato con revisioni tracciate
    body = body.replace(/<w:del\b[\s\S]*?<\/w:del>/g, '');

    var paras = [], m;
    RE_P.lastIndex = 0;
    while ((m = RE_P.exec(body)) !== null) {
      var inner = m[1] || '';
      var pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(inner);
      var ilvl = -1, numId = -1;
      if (pPr) {
        var np = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(pPr[1]);
        if (np) {
          var mi = /<w:ilvl\s+w:val="(\d+)"/.exec(np[1]);
          var mn = /<w:numId\s+w:val="(\d+)"/.exec(np[1]);
          ilvl = mi ? parseInt(mi[1], 10) : 0;
          numId = mn ? parseInt(mn[1], 10) : 0;
        }
      }
      var t = '', mt;
      RE_T.lastIndex = 0;
      while ((mt = RE_T.exec(inner)) !== null) t += mt[1];
      t = unescapeXml(t).replace(/ /g, ' ').replace(/[ \t]+/g, ' ').trim();
      paras.push({ text: t, ilvl: ilvl, numId: numId });
    }
    return paras;
  }

  /* ------------------------------------------- tabelle nei documenti .docx */

  var RE_TBL = /<w:tbl>([\s\S]*?)<\/w:tbl>/g;
  var RE_TR = /<w:tr(?:\s[^>]*)?>([\s\S]*?)<\/w:tr>/g;
  var RE_TC = /<w:tc>([\s\S]*?)<\/w:tc>/g;

  /* Le porzioni di testo interne al medesimo capoverso vanno unite senza
     interposizione di spazi, poiché Word suddivide una parola in più
     porzioni; sono invece i capoversi a doversi separare con uno spazio. */
  function testoCella(xml) {
    var capoversi = [], mp, m;
    var re = /<w:p(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/w:p>)/g;
    while ((mp = re.exec(xml)) !== null) {
      var t = '';
      RE_T.lastIndex = 0;
      while ((m = RE_T.exec(mp[1] || '')) !== null) t += m[1];
      t = unescapeXml(t).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
      if (t) capoversi.push(t);
    }
    if (!capoversi.length) {
      var u = '';
      RE_T.lastIndex = 0;
      while ((m = RE_T.exec(xml)) !== null) u += m[1];
      return unescapeXml(u).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
    }
    return capoversi.join(' ');
  }

  function docxToTabelle(xml) {
    var body = xml;
    var mb = /<w:body[^>]*>([\s\S]*)<\/w:body>/.exec(xml);
    if (mb) body = mb[1];
    body = body.replace(/<w:del\b[\s\S]*?<\/w:del>/g, '');

    var tabelle = [], mt;
    RE_TBL.lastIndex = 0;
    while ((mt = RE_TBL.exec(body)) !== null) {
      var righe = [], mr;
      RE_TR.lastIndex = 0;
      while ((mr = RE_TR.exec(mt[1])) !== null) {
        var celle = [], mc;
        RE_TC.lastIndex = 0;
        while ((mc = RE_TC.exec(mr[1])) !== null) celle.push(testoCella(mc[1]));
        if (celle.length) righe.push(celle);
      }
      if (righe.length) tabelle.push(righe);
    }
    return tabelle;
  }

  /* --------------------------------------- presenze dei componenti */

  var RE_SEGNO_PRESENZA = /^(x|×|si|s[iì]|p|presente|v|✓|✔)\b/i;
  var RE_INTESTA_COMPONENTE = /(componente|component|cognome|nominativo|nome)/;
  var RE_INTESTA_PRESENTE = /^present/;
  var RE_INTESTA_ASSENTE = /^assent/;
  var RE_INTESTA_RUOLO = /^(ruolo|qualifica|componente di|categoria)/;

  // Un nome di persona: almeno due parole, iniziali maiuscole, senza cifre.
  var RE_NOME = /^[A-ZÀ-Þ][\wÀ-ÿ'’.-]+(?:\s+[A-ZÀ-Þ][\wÀ-ÿ'’.-]+){1,4}$/;

  function sembraNome(t) {
    if (!t || t.length < 5 || t.length > 60) return false;
    if (/\d/.test(t)) return false;
    return RE_NOME.test(t.trim());
  }

  /* Nei PDF il nominativo va a capo, e la prima riga può recare il solo
     cognome: una parola sola, iniziale maiuscola, è quindi un inizio di
     nominativo accettabile quando la riga porta il segno di spunta. */
  var RE_INIZIO_NOME = /^[A-ZÀ-Þ][a-zà-öø-ÿ'’.-]{2,}$/;

  function sembraInizioNome(t) {
    if (!t) return false;
    t = t.trim();
    if (t.length < 3 || t.length > 30 || /\d/.test(t)) return false;
    return RE_INIZIO_NOME.test(t);
  }

  /* Individua la tabella delle presenze fra quelle del documento e ne ricava,
     per ciascun componente, il ruolo e lo stato. Il ruolo si eredita dalle
     celle unite lasciate vuote. Lo stato si determina dalla colonna
     "Presente": una cella non vuota vale presenza, anche quando reca
     un'annotazione oraria; la presenza si considera assente soltanto se la
     colonna "Presente" è vuota e quella "Assente" non lo è. */
  /* Individua nella riga indicata le colonne della tabella delle presenze. */
  function intestazionePresenze(riga) {
    var intest = riga.map(function (c) { return norm(c); });
    var iComp = -1, iPres = -1, iAss = -1, iRuolo = -1;
    intest.forEach(function (c, k) {
      if (iComp < 0 && RE_INTESTA_COMPONENTE.test(c)) iComp = k;
      if (iPres < 0 && RE_INTESTA_PRESENTE.test(c)) iPres = k;
      if (iAss < 0 && RE_INTESTA_ASSENTE.test(c)) iAss = k;
      if (iRuolo < 0 && RE_INTESTA_RUOLO.test(c)) iRuolo = k;
    });
    if (iComp < 0 || (iPres < 0 && iAss < 0)) return null;
    // l'intestazione "Presente Assente" può finire in una sola colonna
    if (iAss === iPres) iAss = -1;
    return { comp: iComp, pres: iPres, ass: iAss, ruolo: iRuolo };
  }

  /* In mancanza di intestazioni riconoscibili, si cerca la colonna che
     contenga nomi di persona e quelle che contengano segni di spunta. */
  function colonnePerContenuto(tb) {
    var larghezza = 0;
    tb.forEach(function (rr) { if (rr.length > larghezza) larghezza = rr.length; });
    if (larghezza < 2) return null;
    var colNomi = -1;
    for (var c1 = 0; c1 < larghezza && colNomi < 0; c1++) {
      var nomi = 0;
      tb.forEach(function (rr) { if (sembraNome(rr[c1] || '')) nomi++; });
      if (nomi >= Math.max(3, Math.floor(tb.length * 0.4))) colNomi = c1;
    }
    if (colNomi < 0) return null;
    var colSegni = [];
    for (var c2 = 0; c2 < larghezza; c2++) {
      if (c2 === colNomi) continue;
      var segni = 0;
      tb.forEach(function (rr) { if (/^[x×v✓✔]$/i.test((rr[c2] || '').trim())) segni++; });
      if (segni >= 2) colSegni.push(c2);
    }
    if (!colSegni.length) return null;
    return {
      comp: colNomi,
      pres: colSegni[0],
      ass: colSegni.length > 1 ? colSegni[1] : -1,
      ruolo: colNomi > 0 ? colNomi - 1 : -1
    };
  }

  /* Ricava i componenti da tutte le tabelle riconoscibili. La tabella delle
     presenze può proseguire su più pagine, ripetendo l'intestazione, e nei
     PDF i nominativi lunghi vanno a capo: una riga priva di segni di spunta
     che rechi soltanto testo nella colonna del componente viene perciò
     considerata continuazione del nominativo precedente. */
  function estraiPresenze(tabelle, continuazioni) {
    if (!tabelle || !tabelle.length) return { rilevata: false, righe: [] };

    var righe = [], visti = {}, ruoloCorrente = '', qualcosa = false;

    tabelle.forEach(function (t) {
      if (!t || t.length < 3) return;

      var colonne = null, inizio = 0;
      for (var r = 0; r < t.length && !colonne; r++) {
        var c = intestazionePresenze(t[r]);
        if (c) { colonne = c; inizio = r + 1; }
      }
      if (!colonne && !qualcosa) {
        colonne = colonnePerContenuto(t);
        inizio = 0;
      }
      if (!colonne) return;
      qualcosa = true;

      var digiuno = 0;
      for (var k = inizio; k < t.length; k++) {
        var riga = t[k];
        var nome = (riga[colonne.comp] || '').trim();
        var cellaPres = colonne.pres >= 0 ? (riga[colonne.pres] || '').trim() : '';
        var cellaAss = colonne.ass >= 0 ? (riga[colonne.ass] || '').trim() : '';
        var haSegni = !!(cellaPres || cellaAss);

        if (colonne.ruolo >= 0) {
          var ru = (riga[colonne.ruolo] || '').trim();
          if (ru) {
            // nei PDF anche il ruolo va a capo: le righe prive di segni ne
            // costituiscono la continuazione
            if (continuazioni && !haSegni && ruoloCorrente) ruoloCorrente = (ruoloCorrente + ' ' + ru).trim();
            else ruoloCorrente = ru;
          }
        }

        /* Continuazione del nominativo andato a capo: la riga non reca segni
           di spunta e riporta soltanto altro testo nella colonna del
           componente. Vale per le sole tabelle ricostruite dai PDF; in un
           .docx ciascun componente occupa una riga sola. */
        if (continuazioni && nome && !haSegni && righe.length) {
          var ultimo = righe[righe.length - 1];
          var esteso = (ultimo.nome + ' ' + nome).replace(/\s+/g, ' ').trim();
          if (esteso.length <= 60 && !/\d/.test(nome) && nome.length <= 30) {
            delete visti[chiavePresenza(ultimo.nome)];
            ultimo.nome = esteso;
            visti[chiavePresenza(esteso)] = true;
            digiuno = 0;
            continue;
          }
        }

        var accettabile = sembraNome(nome) ||
          (continuazioni && haSegni && sembraInizioNome(nome));
        if (!nome || !accettabile) {
          if (righe.length) { digiuno++; if (digiuno >= 8) break; }
          continue;
        }
        digiuno = 0;

        var stato, nota = '';
        if (cellaPres) {
          stato = 'presente';
          if (!/^[x×v✓✔]$/i.test(cellaPres)) nota = cellaPres.replace(/^[x×v✓✔]\s*/i, '').trim();
          if (cellaAss && !RE_SEGNO_PRESENZA.test(cellaAss)) nota = (nota ? nota + ' · ' : '') + cellaAss;
        } else if (cellaAss) {
          stato = 'assente';
          if (!/^[x×]$/i.test(cellaAss)) nota = cellaAss;
        } else {
          stato = 'ignoto';
        }

        var ch = chiavePresenza(nome);
        if (visti[ch]) continue;
        visti[ch] = true;
        righe.push({ ruolo: ruoloCorrente, nome: nome, stato: stato, nota: nota });
      }
    });

    return { rilevata: righe.length > 0, righe: righe };
  }

  function chiavePresenza(n) {
    return norm(n).replace(/[^a-z0-9]+/g, ' ').trim();
  }

  /* -------------------------------------------------- estrazione da PDF */
  /* Ricostruisce righe e paragrafi dalle coordinate degli elementi di testo. */

  /* Raggruppa gli elementi di testo di una pagina in righe, per coordinata y. */
  function pdfRighe(items) {
    var lines = [];
    items.forEach(function (it) {
      if (!it.str) return;
      var y = Math.round(it.y);
      var line = null;
      for (var i = lines.length - 1; i >= 0 && i >= lines.length - 4; i--) {
        if (Math.abs(lines[i].y - y) <= 2.5) { line = lines[i]; break; }
      }
      if (!line) { line = { y: y, parts: [], x0: it.x, x1: it.x + it.w }; lines.push(line); }
      line.parts.push(it);
      if (it.x < line.x0) line.x0 = it.x;
      if (it.x + it.w > line.x1) line.x1 = it.x + it.w;
    });
    lines.sort(function (a, b) { return b.y - a.y; });
    lines.forEach(function (l) {
      l.parts.sort(function (a, b) { return a.x - b.x; });
      l.text = l.parts.map(function (p) { return p.str; }).join('').replace(/\s+/g, ' ').trim();
      var hh = l.parts.map(function (p) { return p.h; }).filter(function (v) { return v > 0; }).sort(function (a, b) { return a - b; });
      l.altezza = hh.length ? hh[Math.floor(hh.length / 2)] : 0;
    });
    return lines.filter(function (l) { return l.text.length > 0; });
  }

  function pdfItemsToParagraphs(pages) {
    var paras = [];
    pages.forEach(function (items) {
      var lines = pdfRighe(items);
      if (!lines.length) return;

      var widths = lines.map(function (l) { return l.x1 - l.x0; }).sort(function (a, b) { return a - b; });
      var maxW = widths[widths.length - 1] || 1;

      // interlinea di riferimento: si ricava dall'altezza del carattere, non
      // dalla mediana degli stacchi, che in un testo di paragrafi brevi
      // coincide con lo stacco fra paragrafi e impedirebbe ogni separazione.
      var alt = [];
      lines.forEach(function (l) {
        l.parts.forEach(function (p) { if (p.h > 0) alt.push(p.h); });
      });
      alt.sort(function (a, b) { return a - b; });
      var altMediana = alt.length ? alt[Math.floor(alt.length / 2)] : 11;

      var gaps = [];
      for (var i = 1; i < lines.length; i++) gaps.push(Math.abs(lines[i - 1].y - lines[i].y));
      var gapsOrd = gaps.slice().sort(function (a, b) { return a - b; });
      var q25 = gapsOrd.length ? gapsOrd[Math.floor(gapsOrd.length * 0.25)] : altMediana * 1.2;
      var soglia = Math.max(altMediana * 1.6, q25 * 1.32);

      var cur = null;
      for (var j = 0; j < lines.length; j++) {
        var L = lines[j];
        var startNew = true;
        if (cur) {
          var prev = lines[j - 1];
          var gap = Math.abs(prev.y - L.y);
          var prevShort = (prev.x1 - prev.x0) < 0.62 * maxW;
          var prevEnds = /[.:;!?]$/.test(prev.text);
          var rientro = (L.x0 - prev.x0) > altMediana * 0.9;
          // un cambio di corpo tipografico segnala un'intestazione: le
          // intestazioni dei punti vanno separate dal testo che le circonda
          var cambioCorpo = (L.altezza > 0 && prev.altezza > 0) &&
                            (Math.abs(L.altezza - prev.altezza) > altMediana * 0.12);
          startNew = (gap > soglia) || (prevShort && prevEnds) || rientro || cambioCorpo ||
                     /^\s*(\d{1,2}[.)]|[a-z][.)])\s/.test(L.text);
        }
        if (startNew || !cur) {
          cur = { text: L.text, ilvl: -1, numId: -1 };
          paras.push(cur);
        } else {
          cur.text += ' ' + L.text;
        }
      }
      paras.push({ text: '', ilvl: -1, numId: -1 });
    });
    return unisciCapoversiSpezzati(paras);
  }

  /* Nei PDF la ricostruzione per coordinate interrompe talvolta un capoverso a
     meta' frase. Dove un capoverso non termina con segno di punteggiatura
     conclusivo e quello successivo si apre in minuscolo, i due vengono
     ricongiunti: la frattura non corrisponde a un a capo del documento. */
  var RE_FINE_CONCLUSA = /[.:;!?»"'\)\]]\s*$/;
  var RE_INIZIO_MINUSCOLO = /^[a-zàèéìòùáíóúâêîôûäëïöü(«"']/;

  function unisciCapoversiSpezzati(paras) {
    var out = [], ultimo = -1;
    for (var i = 0; i < paras.length; i++) {
      var p = paras[i], t = p.text.trim();
      if (!t) { out.push(p); continue; }
      if (ultimo >= 0) {
        var prec = out[ultimo].text.trim();
        if (prec.length > 25 && !RE_FINE_CONCLUSA.test(prec) && RE_INIZIO_MINUSCOLO.test(t)) {
          out[ultimo].text = prec + ' ' + t;
          continue;
        }
      }
      out.push(p);
      ultimo = out.length - 1;
    }
    return out;
  }

  /* ------------------------------ tabella delle presenze nei documenti PDF */

  /* Nei PDF non esiste struttura di tabella: le colonne vanno ricostruite
     dalle ascisse dei segni di spunta. Si raggruppano le ascisse di tutti i
     segni isolati, si assumono come colonne i due raggruppamenti più
     frequenti, e si attribuisce a ciascuna riga il nome che precede la prima
     colonna. Il procedimento è congetturale: un PDF che non collochi i segni
     in colonna non viene riconosciuto, e in tal caso le presenze si
     inseriscono a mano. */
  /* Ricostruisce una matrice righe/colonne dalle ascisse degli elementi di
     testo: le colonne si ricavano raggruppando le ascisse di inizio, e
     ciascun elemento viene attribuito alla colonna il cui inizio lo precede
     immediatamente. Il risultato ha la stessa forma di una tabella .docx e
     viene analizzato dalla medesima procedura. */
  function pdfItemsToTabella(pages) {
    var tabelle = [];

    pages.forEach(function (items) {
      var matrice = [];
      var lines = pdfRighe(items);
      if (lines.length < 3) return;

      var altezze = [];
      lines.forEach(function (l) { if (l.altezza > 0) altezze.push(l.altezza); });
      altezze.sort(function (a, b) { return a - b; });
      var corpo = altezze.length ? altezze[Math.floor(altezze.length / 2)] : 11;
      var tolleranza = Math.max(8, corpo * 0.85);

      var ascisse = [];
      lines.forEach(function (l) {
        l.parts.forEach(function (p) { if ((p.str || '').trim()) ascisse.push(p.x); });
      });
      if (!ascisse.length) return;
      ascisse.sort(function (a, b) { return a - b; });

      var gruppi = [];
      ascisse.forEach(function (x) {
        var g = gruppi.length ? gruppi[gruppi.length - 1] : null;
        if (g && x - g.ultimo <= tolleranza) { g.somma += x; g.n++; g.ultimo = x; }
        else gruppi.push({ somma: x, n: 1, ultimo: x, inizio: x });
      });
      // si trattengono soltanto i raggruppamenti ricorrenti: una colonna vera
      // si ripete su più righe, un rientro occasionale no
      var colonne = gruppi.filter(function (g) { return g.n >= 3; })
        .map(function (g) { return g.inizio; })
        .sort(function (a, b) { return a - b; });
      if (colonne.length < 2) return;

      function colonnaDi(x) {
        var k = 0;
        for (var i = 0; i < colonne.length; i++) {
          if (x >= colonne[i] - tolleranza) k = i; else break;
        }
        return k;
      }

      lines.forEach(function (l) {
        var celle = [];
        for (var i = 0; i < colonne.length; i++) celle.push('');
        l.parts.forEach(function (p) {
          if (!(p.str || '').trim() && !p.str) return;
          celle[colonnaDi(p.x)] += p.str;
        });
        for (var j = 0; j < celle.length; j++) celle[j] = celle[j].replace(/\s+/g, ' ').trim();
        if (celle.some(function (c) { return c; })) matrice.push(celle);
      });

      // una matrice per pagina: la tabella può proseguire nella pagina
      // seguente con una diversa disposizione delle colonne
      if (matrice.length) { tabelle.push(matrice); matrice = []; }
    });

    return tabelle;
  }

  function pdfItemsToPresenze(pages) {
    var r = estraiPresenze(pdfItemsToTabella(pages), true);
    if (r.rilevata) r.congetturale = true;
    return r;
  }

  /* --------------------------------------------------------- date e numeri */

  var MESI = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
              'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'];

  var UNITA = {
    zero: 0, uno: 1, due: 2, tre: 3, quattro: 4, cinque: 5, sei: 6, sette: 7, otto: 8, nove: 9,
    dieci: 10, undici: 11, dodici: 12, tredici: 13, quattordici: 14, quindici: 15, sedici: 16,
    diciassette: 17, diciotto: 18, diciannove: 19, venti: 20, ventuno: 21, ventidue: 22,
    ventitre: 23, ventitré: 23, ventiquattro: 24, venticinque: 25, ventisei: 26, ventisette: 27,
    ventotto: 28, ventinove: 29, trenta: 30, trentuno: 31, trentadue: 32, trentatre: 33,
    trentaquattro: 34, trentacinque: 35, trentasei: 36, trentasette: 37, trentotto: 38,
    trentanove: 39, quaranta: 40, quarantuno: 41, quarantadue: 42, quarantatre: 43,
    quarantaquattro: 44, quarantacinque: 45, quarantasei: 46, quarantasette: 47,
    quarantotto: 48, quarantanove: 49, cinquanta: 50
  };

  function annoDaParole(w) {
    var s = norm(w).replace(/\s+/g, '');
    var m = /^(due)?mila(.*)$/.exec(s);
    if (!m) return null;
    var rest = m[2];
    if (!rest) return 2000;
    if (UNITA[rest] !== undefined) return 2000 + UNITA[rest];
    return null;
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function estraiData(paras, filename) {
    var head = paras.slice(0, 60).map(function (p) { return p.text; }).join('\n');
    var h = norm(head);
    var m;

    // "Il giorno 23 del mese di aprile dell'anno duemilaventisei"
    m = /il giorno\s+(\d{1,2})\s*(?:del mese di|del)?\s*([a-z]+)\s*(?:dell'anno|del)?\s*([a-z]+|\d{4})/.exec(h);
    if (m) {
      var mi = MESI.indexOf(m[2]);
      if (mi >= 0) {
        var y = /^\d{4}$/.test(m[3]) ? parseInt(m[3], 10) : annoDaParole(m[3]);
        if (y) return { iso: y + '-' + pad2(mi + 1) + '-' + pad2(+m[1]), label: pad2(+m[1]) + '/' + pad2(mi + 1) + '/' + y };
      }
    }
    // "23 aprile 2026"
    m = new RegExp('(\\d{1,2})\\s+(' + MESI.join('|') + ')\\s+(\\d{4})').exec(h);
    if (m) {
      var mi2 = MESI.indexOf(m[2]);
      return { iso: m[3] + '-' + pad2(mi2 + 1) + '-' + pad2(+m[1]), label: pad2(+m[1]) + '/' + pad2(mi2 + 1) + '/' + m[3] };
    }
    // dal nome del file
    var f = norm(filename || '');
    m = new RegExp('(\\d{1,2})[ _-]+(' + MESI.join('|') + ')[ _-]+(\\d{4})').exec(f);
    if (m) {
      var mi3 = MESI.indexOf(m[2]);
      return { iso: m[3] + '-' + pad2(mi3 + 1) + '-' + pad2(+m[1]), label: pad2(+m[1]) + '/' + pad2(mi3 + 1) + '/' + m[3] };
    }
    m = /(\d{1,2})[._\-\/](\d{1,2})[._\-\/](\d{4})/.exec(f);
    if (m) return { iso: m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]), label: pad2(+m[1]) + '/' + pad2(+m[2]) + '/' + m[3] };

    return { iso: null, label: null };
  }

  function estraiNumero(paras, filename) {
    var head = norm(paras.slice(0, 40).map(function (p) { return p.text; }).join('\n'));
    var m = /verbale[^\n]{0,60}?n[.°*]?\s*(\d{1,3})\s*(?:\/\s*(\d{4}\s*[-\/]\s*\d{2,4}))?/.exec(head);
    if (m) return { numero: parseInt(m[1], 10), anno: m[2] ? m[2].replace(/\s+/g, '') : null };
    m = /n[.°*]?\s*(\d{1,3})\s*\/\s*(\d{4}\s*[-\/]\s*\d{2,4})/.exec(head);
    if (m) return { numero: parseInt(m[1], 10), anno: m[2].replace(/\s+/g, '') };
    var f = norm(filename || '');
    m = /verbale[ _\-]*n[.°_ -]*(\d{1,3})/.exec(f);
    if (m) return { numero: parseInt(m[1], 10), anno: null };
    return { numero: null, anno: null };
  }

  /* ------------------------------------------------------ ordine del giorno */

  var RE_ODG_ANCHOR = /(seguenti\s+punti|punti\s+all|ordine\s+del\s+giorno|o\.?\s?d\.?\s?g)/;
  var RE_ODG_STOP = /^(verifica del numero legale|risultano present|sono present|presiede|constatata|il consiglio d|partecipa alla riunione|si procede alla verifica)/;

  function estraiOdg(paras) {
    var anchor = -1;
    for (var i = 0; i < Math.min(paras.length, 60); i++) {
      var n = norm(paras[i].text);
      if (!n) continue;
      if (RE_ODG_ANCHOR.test(n) && /:\s*$/.test(paras[i].text) && n.length > 40) { anchor = i; break; }
    }
    if (anchor < 0) {
      for (var k = 0; k < Math.min(paras.length, 60); k++) {
        var n2 = norm(paras[k].text);
        if (n2 && RE_ODG_ANCHOR.test(n2) && n2.length > 40) { anchor = k; break; }
      }
    }
    if (anchor < 0) return { items: [], end: 0, anchor: -1 };

    var items = [], vuoti = 0, end = anchor;
    for (var j = anchor + 1; j < paras.length && items.length < 40; j++) {
      var p = paras[j], t = p.text.trim();
      if (!t) { vuoti++; if (items.length >= 2 && vuoti >= 2) break; continue; }
      var nt = norm(t);
      if (RE_ODG_STOP.test(nt)) break;
      if (nt.length > 400) break;
      vuoti = 0;

      var lvl = p.ilvl >= 0 ? p.ilvl : 0;
      var clean = t;
      var mm = /^\s*(\d{1,2})\s*[.)]\s+(.*)$/.exec(t);
      if (mm && p.ilvl < 0) { clean = mm[2]; lvl = 0; }
      else {
        var ms = /^\s*([a-z])\s*[.)]\s+(.*)$/.exec(t);
        if (ms && p.ilvl < 0) { clean = ms[2]; lvl = 1; }
      }
      clean = clean.replace(/[;,.\s]+$/, '').trim();
      if (!clean) continue;
      items.push({ level: Math.min(lvl, 2), titolo: clean });
      end = j;
    }

    // numerazione gerarchica
    var counters = [0, 0, 0];
    items.forEach(function (it) {
      counters[it.level]++;
      for (var z = it.level + 1; z < 3; z++) counters[z] = 0;
      var parts = [];
      for (var q = 0; q <= it.level; q++) parts.push(counters[q]);
      it.num = parts.join('.');
    });
    return { items: items, end: end, anchor: anchor };
  }

  /* ------------------------------------------- sezioni del corpo del verbale */

  var RE_BODY_START = /(dichiara aperta la seduta|constatata la sussistenza del numero legale|aperta la seduta|partecipa alla riunione)/;
  var RE_TAIL = /(esaurita la trattazione|alle ore [\d.:]+\s*(la seduta|si scioglie|e' tolta|viene tolta)|letto, approvato e sottoscritto|la seduta (e'|viene) (tolta|sciolta)|non essendovi altro)/;

  /* Nei PDF l'intestazione di un punto resta talvolta inglobata nel paragrafo
     che la precede. Dove la dizione dell'ordine del giorno si ritrova
     all'interno di un paragrafo, questo viene diviso in quel punto, così che
     l'intestazione torni a essere un paragrafo autonomo. */
  function separaIntestazioni(paras, odg, dopo) {
    var frasi = [];
    odg.items.filter(function (it) { return it.level === 0; }).forEach(function (it) {
      var parole = norm(it.titolo).split(/\s+/).filter(Boolean).slice(0, 5).join(' ');
      if (parole.length >= 12) frasi.push(parole);
    });
    if (!frasi.length) return paras;

    var out = [];
    for (var i = 0; i < paras.length; i++) {
      var p = paras[i];
      if (i <= dopo || p.text.length < 150) { out.push(p); continue; }

      var nm = normWithMap(p.text);
      var taglio = -1, off = -1;
      for (var k = 0; k < frasi.length; k++) {
        var pos = nm.n.indexOf(frasi[k]);
        if (pos <= 30) continue;
        var o = nm.map[pos];
        if (o === undefined || o <= 0 || o >= p.text.length - 5) continue;
        /* Il taglio si opera soltanto dove la dizione apre un periodo: la
           lettera iniziale è maiuscola e il testo che precede si chiude con un
           segno conclusivo. Senza questa verifica la medesima dizione ricorrente
           nella prosa ("illustra la proposta di regolamento per l'uso dei
           laboratori") verrebbe scambiata per un'intestazione. */
        if (!/[A-ZÀÈÉÌÒÙ«"']/.test(p.text.charAt(o))) continue;
        var precedente = p.text.substring(0, o).replace(/\s+$/, '');
        if (!/[.;:!?»)\]]$/.test(precedente)) continue;
        if (taglio < 0 || pos < taglio) { taglio = pos; off = o; }
      }
      if (taglio < 0) { out.push(p); continue; }
      var a = p.text.substring(0, off).trim();
      var b = p.text.substring(off).trim();
      if (!a || !b) { out.push(p); continue; }
      out.push({ text: a, ilvl: p.ilvl, numId: p.numId });
      out.push({ text: b, ilvl: -1, numId: -1 });
    }
    return out;
  }

  function estraiSezioni(paras, odg, odgEnd) {
    var start = odgEnd + 1;
    for (var i = odgEnd; i < Math.min(paras.length, odgEnd + 200); i++) {
      if (RE_BODY_START.test(norm(paras[i].text))) { start = i + 1; break; }
    }

    var level0 = odg.items.filter(function (it) { return it.level === 0; });
    var sezioni = [];
    var cursor = start;

    level0.forEach(function (it, idx) {
      var best = -1, bestScore = 0;
      var limit = paras.length;
      for (var j = cursor; j < limit; j++) {
        var t = paras[j].text.trim();
        if (t.length < 5) continue;
        var sim = similarity(it.titolo, t.length > 400 ? t.substring(0, 400) : t);
        // se l'intestazione è rimasta incollata al testo che segue, il
        // confronto va tentato anche sul solo esordio del paragrafo
        if (t.length > 130) {
          var esordio = t.substring(0, 130);
          var p = esordio.lastIndexOf('. ');
          if (p > 20) esordio = esordio.substring(0, p);
          var sim2 = similarity(it.titolo, esordio);
          if (sim2.score > sim.score) sim = sim2;
        }
        if (sim.shared >= 2 && sim.score > bestScore) { bestScore = sim.score; best = j; }
        if (bestScore >= 0.9) break;
      }
      if (best >= 0 && bestScore >= 0.45) {
        sezioni.push({ punto: it.num, titolo: it.titolo, intestazione: paras[best].text.trim(), from: best, score: Math.round(bestScore * 100) / 100 });
        cursor = best + 1;
      } else {
        sezioni.push({ punto: it.num, titolo: it.titolo, intestazione: null, from: -1, score: Math.round(bestScore * 100) / 100 });
      }
    });

    // sottopunti agganciati al punto padre
    odg.items.filter(function (it) { return it.level > 0; }).forEach(function (it) {
      var parent = it.num.split('.')[0];
      sezioni.forEach(function (s) {
        if (s.punto === parent) { if (!s.sottopunti) s.sottopunti = []; s.sottopunti.push(it.num + ' ' + it.titolo); }
      });
    });

    // coda del verbale
    var tailStart = -1;
    for (var k = paras.length - 1; k >= start; k--) {
      if (RE_TAIL.test(norm(paras[k].text))) tailStart = k;
    }

    var located = sezioni.filter(function (s) { return s.from >= 0; }).sort(function (a, b) { return a.from - b.from; });
    for (var z = 0; z < located.length; z++) {
      var nxt = (z + 1 < located.length) ? located[z + 1].from - 1 : (tailStart > located[z].from ? tailStart - 1 : paras.length - 1);
      located[z].to = nxt;
    }

    var out = [];
    if (located.length && located[0].from > start - 1) {
      // la parte iniziale viene distinta in due: intestazione e presenze da un
      // lato, elenco dei punti all'ordine del giorno dall'altro, perché un
      // riscontro nell'elenco non equivale alla trattazione dell'argomento
      var fineIntro = located[0].from - 1;
      var a = odg.anchor;
      if (a >= 0 && odgEnd > a && odgEnd <= fineIntro) {
        if (a > 0) out.push({ punto: null, titolo: 'Intestazione, convocazione e verifica delle presenze', from: 0, to: a - 1, accessoria: true });
        out.push({ punto: null, titolo: 'Elenco dei punti all\'ordine del giorno', from: a, to: odgEnd, accessoria: true });
        if (odgEnd < fineIntro) out.push({ punto: null, titolo: 'Verifica del numero legale e apertura della seduta', from: odgEnd + 1, to: fineIntro, accessoria: true });
      } else {
        out.push({ punto: null, titolo: 'Parte preliminare del verbale', from: 0, to: fineIntro, accessoria: true });
      }
    } else if (!located.length) {
      out.push({ punto: null, titolo: 'Verbale non segmentato per punti', from: 0, to: paras.length - 1 });
    }
    located.forEach(function (s) { out.push(s); });
    if (tailStart >= 0) {
      out.push({ punto: null, titolo: 'Varie ed eventuali e chiusura della seduta', from: tailStart, to: paras.length - 1, accessoria: true });
    }
    var nonLocated = sezioni.filter(function (s) { return s.from < 0; });
    return { sezioni: out, nonLocalizzati: nonLocated };
  }

  /* ------------------------------------------------------------- delibere */

  var RE_DEL_FULL = /delibera\s*n[.°*]?\s*(\d{1,4})\s*\/\s*(\d{4}\s*[-\/]\s*\d{2,4})/g;
  var RE_DEL_BARE = /delibera\s*n[.°*]?\s*(\d{1,4})(?!\d)/g;
  var RE_ALTRO_ORGANO = /(collegio|docenti|giunta|consiglio di classe|revisori|comune|citta metropolitana|regione|assemblea|comitato)/;

  /* Il testo di una sezione viene analizzato per intero, non paragrafo per
     paragrafo: nei PDF il riferimento "(Delibera n. 77/2025-26)" si trova con
     una certa frequenza spezzato fra due righe, e quindi fra due paragrafi
     ricostruiti. La concatenazione rende l'estrazione indipendente da come il
     testo è stato segmentato. */
  function costruisciTesto(paras, from, to) {
    var parti = [], indici = [], pos = 0;
    for (var i = Math.max(0, from); i <= Math.min(to, paras.length - 1); i++) {
      var t = paras[i].text.trim();
      if (!t) continue;
      if (parti.length) pos += 1; // il separatore occupa un carattere
      indici.push({ para: i, start: pos, end: pos + t.length });
      parti.push(t);
      pos += t.length;
    }
    // separatore che conta come spazio: così un riferimento spezzato fra due
    // paragrafi, per esempio "(Delibera" + "n. 77/2025-26)", resta leggibile.
    return { testo: parti.join('\n'), indici: indici };
  }

  function testoSezione(paras, from, to) {
    return costruisciTesto(paras, from, to).testo;
  }

  // Individua il paragrafo che contiene una data posizione del testo concatenato.
  function paragrafoDaPosizione(indici, pos) {
    for (var i = 0; i < indici.length; i++) {
      if (pos >= indici[i].start && pos <= indici[i].end) return indici[i].para;
      if (pos < indici[i].start) return indici[i].para;
    }
    return indici.length ? indici[indici.length - 1].para : -1;
  }

  var RE_MARCATORE_DISPOSITIVO = /(^|\s)DELIBERA(\s|$)/;
  var RE_PREMESSA = /(vist[ao]|richiamat|preso atto|considerat|premesso|precedent|parere)/;

  var RE_ESORDIO_ORGANO = /^il consiglio d[i'’]/;
  var RE_SOLO_DELIBERA = /^delibera\s*:?\s*$/;

  function estraiDelibere(paras, from, to) {
    var tc = costruisciTesto(paras, from, to);
    var testoOrig = tc.testo;
    if (!testoOrig) return [];
    var nm = normWithMap(testoOrig);
    var n = nm.n;

    // Posizione del dispositivo: le delibere adottate compaiono dopo la formula
    // "DELIBERA" in maiuscolo, mentre i richiami ad altre delibere stanno nelle
    // premesse (VISTA, VISTO, PRESO ATTO) che la precedono.
    var mm = RE_MARCATORE_DISPOSITIVO.exec(testoOrig);
    var inizioDispositivo = mm ? mm.index : -1;
    // la posizione va riportata sul testo normalizzato
    if (inizioDispositivo >= 0) {
      var k = 0;
      for (k = 0; k < nm.map.length; k++) { if (nm.map[k] >= inizioDispositivo) break; }
      inizioDispositivo = k;
    }

    var out = [], seenNum = {}, m;

    function accetta(numTxt, anno, pos) {
      var num = parseInt(numTxt, 10);
      if (!num || seenNum[num]) return;
      seenNum[num] = true;
      out.push({ id: anno ? numTxt + '/' + anno : numTxt, numero: num, anno: anno || null, pos: pos });
    }

    /* Il riferimento a una delibera di altro organo ha forma attributiva:
       "delibera n. 98 del Collegio dei Docenti". La finestra da esaminare è
       perciò quella immediatamente adiacente al riferimento, non un intorno
       ampio, che comprenderebbe le premesse ("preso atto del parere dei
       Revisori dei Conti") facendo scartare delibere proprie. */
    function altroOrgano(testo, inizio, lunghezza) {
      var dopo = testo.substring(inizio + lunghezza, inizio + lunghezza + 45);
      var prima = testo.substring(Math.max(0, inizio - 28), inizio);
      return RE_ALTRO_ORGANO.test(dopo) || RE_ALTRO_ORGANO.test(prima);
    }

    // 1) forma completa "Delibera n. 96/2025-26"
    RE_DEL_FULL.lastIndex = 0;
    while ((m = RE_DEL_FULL.exec(n)) !== null) {
      if (altroOrgano(n, m.index, m[0].length)) continue;
      var nelDispositivo = inizioDispositivo >= 0 && m.index >= inizioDispositivo;
      var prima = n.substring(Math.max(0, m.index - 70), m.index);
      // fuori dal dispositivo si accetta solo se non è un richiamo a un atto precedente
      if (!nelDispositivo && RE_PREMESSA.test(prima)) continue;
      accetta(m[1], m[2].replace(/\s+/g, ''), m.index);
    }

    // 2) forma breve "Delibera n. 96", ammessa solo dentro il dispositivo
    if (inizioDispositivo >= 0) {
      var residuo = n.replace(RE_DEL_FULL, function (s) { return new Array(s.length + 1).join(' '); });
      RE_DEL_BARE.lastIndex = inizioDispositivo;
      while ((m = RE_DEL_BARE.exec(residuo)) !== null) {
        if (altroOrgano(residuo, m.index, m[0].length)) continue;
        accetta(m[1], null, m.index);
      }
    }

    out.sort(function (a, b) { return a.pos - b.pos; });

    /* Testo integrale di ciascuna delibera. Il blocco deliberativo si estende
       dall'enunciazione dell'organo ("IL CONSIGLIO DI ISTITUTO"), o in
       mancanza dalla formula "DELIBERA", fino al paragrafo che reca il numero
       della delibera. I blocchi non si sovrappongono: la ricerca all'indietro
       si arresta alla delibera precedente. */
    var limiteInferiore = Math.max(0, from);
    out.forEach(function (d) {
      var posOrig = nm.map[Math.min(d.pos, nm.map.length - 1)];
      if (posOrig === undefined) posOrig = 0;
      var paraFine = paragrafoDaPosizione(tc.indici, posOrig);
      if (paraFine < 0) { d.testo = ''; d.paraDa = -1; d.paraA = -1; return; }

      var paraInizio = paraFine, trovatoOrgano = -1, trovataFormula = -1;
      for (var i = paraFine; i >= limiteInferiore; i--) {
        var t = norm(paras[i].text);
        if (!t) continue;
        if (trovatoOrgano < 0 && RE_ESORDIO_ORGANO.test(t)) { trovatoOrgano = i; break; }
        if (trovataFormula < 0 && RE_SOLO_DELIBERA.test(t)) trovataFormula = i;
      }
      if (trovatoOrgano >= 0) paraInizio = trovatoOrgano;
      else if (trovataFormula >= 0) paraInizio = trovataFormula;

      d.paraDa = paraInizio;
      d.paraA = paraFine;
      d.testo = testoSezione(paras, paraInizio, paraFine);
      limiteInferiore = paraFine + 1;
    });

    return out;
  }

  var RE_ESITO_DELIBERA = /^delibera\s*:?\s*$|\bdelibera\b.*\bunanimita\b|^(approva|delibera)\b/;
  var RE_ESITO_ATTO = /(prende atto|si prende atto|preso atto dei|ne prende atto)/;
  var RE_ESITO_VIS = /(visiona|visionat|presa visione|esamina|si esprime favorevolmente|esprime parere)/;

  function classificaEsito(paras, from, to, delibere) {
    var esiti = [];
    if (delibere && delibere.length) esiti.push('deliberato');
    var n = norm(testoSezione(paras, from, to));
    if (RE_ESITO_ATTO.test(n)) esiti.push('presa d\'atto');
    if (RE_ESITO_VIS.test(n)) esiti.push('visionato');
    if (!esiti.length) esiti.push('trattato');
    return esiti;
  }

  /* ------------------------------------------------------ parsing completo */

  function parseVerbale(paras0, filename, presenze) {
    var num = estraiNumero(paras0, filename);
    var data = estraiData(paras0, filename);
    var odg = estraiOdg(paras0);
    var paras = separaIntestazioni(paras0, odg, odg.end);
    var sez = estraiSezioni(paras, odg, odg.end);

    sez.sezioni.forEach(function (s) {
      var dett = estraiDelibere(paras, s.from, s.to);
      s.delibere = dett.map(function (d) { return d.id; });
      s.dettaglioDelibere = dett.map(function (d) {
        return { id: d.id, numero: d.numero, anno: d.anno, testo: d.testo || '' };
      });
      s.esiti = classificaEsito(paras, s.from, s.to, s.delibere);
    });

    var tutteDelibere = [];
    sez.sezioni.forEach(function (s) {
      s.delibere.forEach(function (d) { if (tutteDelibere.indexOf(d) < 0) tutteDelibere.push(d); });
    });

    // paragrafi indicizzati con riferimento alla sezione
    var mappa = new Array(paras.length);
    sez.sezioni.forEach(function (s, si) {
      for (var i = Math.max(0, s.from); i <= Math.min(s.to, paras.length - 1); i++) {
        if (mappa[i] === undefined) mappa[i] = si;
      }
    });

    var parag = [];
    for (var i = 0; i < paras.length; i++) {
      var t = paras[i].text.trim();
      if (t.length < 3) continue;
      parag.push({ s: mappa[i] === undefined ? -1 : mappa[i], t: t });
    }

    return {
      numero: num.numero,
      annoScolastico: num.anno,
      dataISO: data.iso,
      dataLabel: data.label,
      odg: odg.items,
      sezioni: sez.sezioni.map(function (s) {
        return {
          punto: s.punto, titolo: s.titolo, intestazione: s.intestazione || null,
          sottopunti: s.sottopunti || null, delibere: s.delibere,
          dettaglioDelibere: s.dettaglioDelibere || [], esiti: s.esiti || [],
          accessoria: !!s.accessoria,
          affidabilita: s.score === undefined ? null : s.score
        };
      }),
      puntiNonLocalizzati: sez.nonLocalizzati.map(function (s) { return s.punto + ' ' + s.titolo; }),
      presenze: presenze || { rilevata: false, righe: [] },
      delibere: tutteDelibere,
      paragrafi: parag,
      caratteri: parag.reduce(function (a, p) { return a + p.t.length; }, 0)
    };
  }

  /* -------------------------------------------------------------- ricerca */

  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  function buildMatchers(query, modo, prefisso) {
    var q = norm(query);
    if (!q) return null;
    if (modo === 'frase') {
      var body = q.split(/\s+/).map(escRe).join('[^a-z0-9]+');
      var tail = prefisso ? '' : '(?![a-z0-9])';
      return [new RegExp('(?<![a-z0-9])' + body + tail, 'g')];
    }
    return q.split(/\s+/).filter(Boolean).slice(0, 5).map(function (w) {
      var tail = prefisso ? '' : '(?![a-z0-9])';
      return new RegExp('(?<![a-z0-9])' + escRe(w) + tail, 'g');
    });
  }

  // fallback per motori senza lookbehind
  function buildMatchersSafe(query, modo, prefisso) {
    try { return buildMatchers(query, modo, prefisso); }
    catch (e) {
      var q = norm(query);
      if (!q) return null;
      if (modo === 'frase') {
        var body = q.split(/\s+/).map(escRe).join('[^a-z0-9]+');
        return [new RegExp('\\b' + body + (prefisso ? '' : '\\b'), 'g')];
      }
      return q.split(/\s+/).filter(Boolean).slice(0, 5).map(function (w) {
        return new RegExp('\\b' + escRe(w) + (prefisso ? '' : '\\b'), 'g');
      });
    }
  }

  /* Restituisce il numero di occorrenze (0 se non tutti i termini compaiono)
     e gli intervalli trovati sul testo normalizzato. */
  /* qualsiasi = true: basta che compaia almeno uno dei termini (serve per
     evidenziare i passaggi quando l'ambito di ricerca è il punto o il verbale,
     dove i termini possono essere distribuiti su paragrafi diversi). */
  function trovaOccorrenze(normText, matchers, qualsiasi) {
    var ranges = [], totale = 0;
    for (var i = 0; i < matchers.length; i++) {
      var re = matchers[i];
      re.lastIndex = 0;
      var c = 0, m;
      while ((m = re.exec(normText)) !== null) {
        c++;
        ranges.push([m.index, m.index + m[0].length]);
        if (m.index === re.lastIndex) re.lastIndex++;
        if (c > 400) break;
      }
      if (c === 0 && !qualsiasi) return { occorrenze: 0, ranges: [] };
      totale += c;
    }
    ranges.sort(function (a, b) { return a[0] - b[0]; });
    return { occorrenze: totale, ranges: ranges };
  }

  function testParagrafo(normText, matchers) {
    return trovaOccorrenze(normText, matchers).occorrenze;
  }

  /* Evidenzia le occorrenze nel testo originale e restituisce HTML sicuro. */
  function evidenzia(originale, matchers, contesto, qualsiasi) {
    var nm = normWithMap(originale);
    var r = trovaOccorrenze(nm.n, matchers, qualsiasi);
    if (!r.ranges.length) return { html: escHtml(taglia(originale, 0, contesto)), occorrenze: 0 };

    // unisce intervalli sovrapposti e li riporta sul testo originale
    var origRanges = [];
    r.ranges.forEach(function (rg) {
      var a = nm.map[rg[0]];
      var b = nm.map[Math.max(rg[0], rg[1] - 1)] + 1;
      if (a === undefined || b === undefined) return;
      if (origRanges.length && a <= origRanges[origRanges.length - 1][1]) {
        if (b > origRanges[origRanges.length - 1][1]) origRanges[origRanges.length - 1][1] = b;
      } else origRanges.push([a, b]);
    });
    if (!origRanges.length) return { html: escHtml(taglia(originale, 0, contesto)), occorrenze: 0 };

    // finestra di contesto centrata sulla prima occorrenza
    var win = finestra(originale, origRanges[0][0], origRanges[origRanges.length - 1][1], contesto);
    var html = '', pos = win.from;
    origRanges.forEach(function (rg) {
      if (rg[1] <= win.from || rg[0] >= win.to) return;
      var a = Math.max(rg[0], win.from), b = Math.min(rg[1], win.to);
      html += escHtml(originale.substring(pos, a)) + '<mark>' + escHtml(originale.substring(a, b)) + '</mark>';
      pos = b;
    });
    html += escHtml(originale.substring(pos, win.to));
    if (win.from > 0) html = '… ' + html;
    if (win.to < originale.length) html = html + ' …';
    return { html: html, occorrenze: r.occorrenze };
  }

  function finestra(s, a, b, contesto) {
    contesto = contesto || 160;
    if (s.length <= contesto * 2) return { from: 0, to: s.length };
    var from = Math.max(0, a - contesto);
    var to = Math.min(s.length, Math.max(b + contesto, from + contesto * 2));
    // allinea ai confini di parola
    while (from > 0 && /\S/.test(s.charAt(from - 1))) from--;
    while (to < s.length && /\S/.test(s.charAt(to))) to++;
    return { from: from, to: to };
  }

  function taglia(s, from, contesto) {
    var lim = (contesto || 160) * 2;
    return s.length > lim ? s.substring(from, from + lim) + ' …' : s;
  }

  function escHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  return {
    mesi: MESI,
    norm: norm,
    normWithMap: normWithMap,
    tokens: tokens,
    similarity: similarity,
    docxToParagraphs: docxToParagraphs,
    docxToTabelle: docxToTabelle,
    estraiPresenze: estraiPresenze,
    pdfItemsToPresenze: pdfItemsToPresenze,
    pdfItemsToTabella: pdfItemsToTabella,
    sembraNome: sembraNome,
    pdfItemsToParagraphs: pdfItemsToParagraphs,
    parseVerbale: parseVerbale,
    buildMatchers: buildMatchersSafe,
    testParagrafo: testParagrafo,
    trovaOccorrenze: trovaOccorrenze,
    evidenzia: evidenzia,
    escHtml: escHtml,
    escRe: escRe
  };
});
