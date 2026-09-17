/* =====================================================================
   github.js — Archivio dei verbali del Consiglio d'Istituto
   Sostituisce il server locale: le richieste "/api/..." dell'applicazione
   vengono intercettate e tradotte in operazioni sul repository privato
   GitHub che custodisce verbali, indice, presenze e cestino.

   Struttura del repository dei dati:
     verbali/            i verbali in archivio
     cestino/            i verbali rimossi, "AAAAMMGG-hhmmss__nome"
     dati/indice.json    l'indice generato dall'applicazione
     dati/presenze.json  correzioni, decadenze, cessazioni, esclusioni

   La chiave di accesso resta nella memoria locale del browser di ciascun
   dispositivo e viene inviata soltanto ad api.github.com.
   ===================================================================== */
(function () {
  'use strict';

  var CHIAVE_IMPOSTAZIONI = 'archivio-verbali-github';
  var API = 'https://api.github.com';
  var ESTENSIONI = /\.(pdf|docx|doc|txt)$/i;
  var fetchOriginale = window.fetch.bind(window);

  var imp = leggiImpostazioni();
  var ramo = null;          // ramo predefinito del repository
  var shaNoti = {};         // percorso -> sha del blob, per le scritture

  /* ------------------------------------------------------- impostazioni */

  function leggiImpostazioni() {
    try {
      var t = window.localStorage.getItem(CHIAVE_IMPOSTAZIONI);
      var d = t ? JSON.parse(t) : null;
      if (d && d.proprietario && d.repository && d.chiave) return d;
    } catch (e) { /* memoria locale non disponibile */ }
    return null;
  }

  function scriviImpostazioni(d) {
    try { window.localStorage.setItem(CHIAVE_IMPOSTAZIONI, JSON.stringify(d)); } catch (e) { }
  }

  function cancellaImpostazioni() {
    try { window.localStorage.removeItem(CHIAVE_IMPOSTAZIONI); } catch (e) { }
  }

  /* ------------------------------------------------------ chiamate API */

  function codificaPercorso(p) {
    return p.split('/').map(encodeURIComponent).join('/');
  }

  function gh(metodo, percorso, corpo, accetta) {
    var opz = {
      method: metodo,
      cache: 'no-store',
      headers: {
        'Authorization': 'Bearer ' + imp.chiave,
        'Accept': accetta || 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
      }
    };
    if (corpo !== undefined) {
      opz.headers['Content-Type'] = 'application/json';
      opz.body = JSON.stringify(corpo);
    }
    return fetchOriginale(API + percorso, opz);
  }

  function ghJson(metodo, percorso, corpo) {
    return gh(metodo, percorso, corpo).then(function (r) {
      return r.text().then(function (t) {
        var d = null;
        try { d = t ? JSON.parse(t) : null; } catch (e) { d = null; }
        if (!r.ok) {
          var err = new Error(messaggioErrore(r.status, d));
          err.status = r.status;
          throw err;
        }
        return d;
      });
    });
  }

  function messaggioErrore(status, d) {
    if (status === 401) return 'chiave di accesso non valida o scaduta';
    if (status === 403) return 'la chiave non ha i permessi di scrittura sul repository';
    if (status === 404) return 'repository o file non trovato';
    return (d && d.message) ? d.message : ('GitHub HTTP ' + status);
  }

  function repo() {
    return '/repos/' + encodeURIComponent(imp.proprietario) + '/' + encodeURIComponent(imp.repository);
  }

  function assicuraRamo() {
    if (ramo) return Promise.resolve(ramo);
    return ghJson('GET', repo()).then(function (r) {
      ramo = r.default_branch || 'main';
      return ramo;
    });
  }

  /* Elenco di una cartella: [] se la cartella non esiste ancora. */
  function elenco(cartella) {
    return assicuraRamo().then(function () {
      return gh('GET', repo() + '/contents/' + codificaPercorso(cartella) + '?ref=' + encodeURIComponent(ramo));
    }).then(function (r) {
      if (r.status === 404) return [];
      if (!r.ok) throw Object.assign(new Error(messaggioErrore(r.status)), { status: r.status });
      return r.json();
    }).then(function (voci) {
      voci = Array.isArray(voci) ? voci : [];
      voci.forEach(function (v) { if (v.type === 'file') shaNoti[v.path] = v.sha; });
      return voci.filter(function (v) { return v.type === 'file'; });
    });
  }

  /* Contenuto grezzo di un file: null se non esiste. */
  function leggiGrezzo(percorso) {
    return assicuraRamo().then(function () {
      return gh('GET', repo() + '/contents/' + codificaPercorso(percorso) + '?ref=' + encodeURIComponent(ramo),
        undefined, 'application/vnd.github.raw+json');
    }).then(function (r) {
      if (r.status === 404) return null;
      if (!r.ok) throw Object.assign(new Error(messaggioErrore(r.status)), { status: r.status });
      return r.arrayBuffer();
    });
  }

  function base64(buffer) {
    var u = new Uint8Array(buffer), s = '', passo = 0x8000;
    for (var i = 0; i < u.length; i += passo) {
      s += String.fromCharCode.apply(null, u.subarray(i, i + passo));
    }
    return btoa(s);
  }

  function shaDi(percorso) {
    if (shaNoti[percorso]) return Promise.resolve(shaNoti[percorso]);
    var i = percorso.lastIndexOf('/');
    return elenco(percorso.slice(0, i)).then(function () { return shaNoti[percorso] || null; });
  }

  /* Scrittura di un file. Se nel frattempo il file è stato modificato da un
     altro dispositivo, lo sha viene riletto e la scrittura ripetuta una volta. */
  function scrivi(percorso, buffer, messaggio, tentativo) {
    return shaDi(percorso).then(function (sha) {
      var corpo = { message: messaggio, content: base64(buffer), branch: ramo };
      if (sha) corpo.sha = sha;
      return gh('PUT', repo() + '/contents/' + codificaPercorso(percorso), corpo);
    }).then(function (r) {
      if ((r.status === 409 || r.status === 422) && !tentativo) {
        delete shaNoti[percorso];
        var i = percorso.lastIndexOf('/');
        return elenco(percorso.slice(0, i)).then(function () {
          return scrivi(percorso, buffer, messaggio, true);
        });
      }
      return r.json().then(function (d) {
        if (!r.ok) throw Object.assign(new Error(messaggioErrore(r.status, d)), { status: r.status });
        if (d && d.content) shaNoti[percorso] = d.content.sha;
        return d;
      });
    });
  }

  /* Spostamenti e cancellazioni con un solo commit, senza ricaricare il
     contenuto dei file: si ricompone l'albero del repository. */
  function commitAlbero(modifiche, messaggio) {
    var rif, commitBase;
    return assicuraRamo().then(function () {
      return ghJson('GET', repo() + '/git/ref/heads/' + encodeURIComponent(ramo));
    }).then(function (r) {
      rif = r.object.sha;
      return ghJson('GET', repo() + '/git/commits/' + rif);
    }).then(function (c) {
      commitBase = c;
      return ghJson('POST', repo() + '/git/trees', {
        base_tree: c.tree.sha,
        tree: modifiche.map(function (m) {
          return { path: m.path, mode: '100644', type: 'blob', sha: m.sha };
        })
      });
    }).then(function (albero) {
      return ghJson('POST', repo() + '/git/commits', {
        message: messaggio, tree: albero.sha, parents: [rif]
      });
    }).then(function (nuovo) {
      return ghJson('PATCH', repo() + '/git/refs/heads/' + encodeURIComponent(ramo), { sha: nuovo.sha });
    }).then(function () {
      modifiche.forEach(function (m) {
        if (m.sha) shaNoti[m.path] = m.sha; else delete shaNoti[m.path];
      });
    });
  }

  /* ------------------------------------------------- utilità del cestino */

  function marcaTemporale() {
    var d = new Date();
    function z(n) { return (n < 10 ? '0' : '') + n; }
    return d.getFullYear() + z(d.getMonth() + 1) + z(d.getDate()) + '-' +
      z(d.getHours()) + z(d.getMinutes()) + z(d.getSeconds());
  }

  function voceCestino(stored) {
    var m = /^(\d{8})-(\d{6})__(.+)$/.exec(stored);
    if (!m) return { originale: stored, quando: null };
    var q = m[1].slice(0, 4) + '-' + m[1].slice(4, 6) + '-' + m[1].slice(6, 8) + 'T' +
      m[2].slice(0, 2) + ':' + m[2].slice(2, 4) + ':' + m[2].slice(4, 6);
    return { originale: m[3], quando: q };
  }

  function nomeValido(n) {
    return !!n && n.trim() && n.indexOf('/') < 0 && n.indexOf('\\') < 0 && n.indexOf('..') < 0;
  }

  /* ------------------------------------------------- risposte simulate */

  function risposta(dati, status) {
    return new Response(typeof dati === 'string' ? dati : JSON.stringify(dati), {
      status: status || 200,
      headers: { 'Content-Type': 'application/json; charset=utf-8' }
    });
  }

  function errore(msg, status, extra) {
    var d = { errore: msg };
    if (extra) Object.keys(extra).forEach(function (k) { d[k] = extra[k]; });
    return risposta(d, status || 500);
  }

  function tipoMime(nome) {
    var e = (nome.split('.').pop() || '').toLowerCase();
    return {
      pdf: 'application/pdf',
      docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      doc: 'application/msword',
      txt: 'text/plain;charset=utf-8'
    }[e] || 'application/octet-stream';
  }

  function paginaGitHub(percorso) {
    return 'https://github.com/' + encodeURIComponent(imp.proprietario) + '/' +
      encodeURIComponent(imp.repository) + '/' + (percorso ? 'blob/' + (ramo || 'main') + '/' + codificaPercorso(percorso)
        : 'tree/' + (ramo || 'main') + '/verbali');
  }

  /* Apertura di un verbale. I PDF e i testi si aprono in una nuova scheda,
     che va creata subito, dentro il clic, perché i browser dei telefoni
     bloccano le schede aperte dopo un'attesa; i .docx vengono scaricati e
     il dispositivo li apre con l'applicazione predefinita. */
  function apriFile(percorso, nome) {
    var visibile = /\.(pdf|txt)$/i.test(nome);
    var scheda = visibile ? window.open('', '_blank') : null;
    if (scheda) {
      try { scheda.document.write('<p style="font-family:sans-serif;padding:2em">Apertura del verbale in corso…</p>'); } catch (e) { }
    }
    return leggiGrezzo(percorso).then(function (buf) {
      if (!buf) throw new Error('file inesistente');
      var url = URL.createObjectURL(new Blob([buf], { type: tipoMime(nome) }));
      if (scheda) {
        scheda.location.href = url;
      } else {
        var a = document.createElement('a');
        a.href = url; a.download = nome;
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
      }
      setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
      return risposta({ ok: true });
    }).catch(function (e) {
      if (scheda) scheda.close();
      alert('Non è stato possibile aprire il verbale: ' + e.message);
      return errore(e.message, 500);
    });
  }

  /* ------------------------------------------------ gestore delle API */

  function gestisci(url, opzioni) {
    var u = new URL(url, location.href);
    var percorso = u.pathname.replace(/^.*\/api\//, '/api/');
    var q = function (k) { return u.searchParams.get(k) || ''; };
    var metodo = ((opzioni && opzioni.method) || 'GET').toUpperCase();
    var corpo = opzioni ? opzioni.body : undefined;

    if (percorso === '/api/config') {
      return pronto().then(function () {
        return risposta({
          archivePath: 'github.com/' + imp.proprietario + '/' + imp.repository,
          esiste: true, remoto: true
        });
      });
    }

    return pronto().then(function () {

      if (percorso === '/api/list') {
        return elenco('verbali').then(function (voci) {
          var files = voci.filter(function (v) { return ESTENSIONI.test(v.name); })
            .sort(function (a, b) { return a.name < b.name ? -1 : 1; })
            .map(function (v) { return { nome: v.name, dimensione: v.size, modificato: v.sha }; });
          return risposta({ files: files });
        });
      }

      if (percorso === '/api/file') {
        var nf = q('nome');
        if (!nomeValido(nf)) return errore('nome non valido', 400);
        return leggiGrezzo('verbali/' + nf).then(function (buf) {
          if (!buf) return errore('file inesistente', 404);
          return new Response(buf, { status: 200, headers: { 'Content-Type': tipoMime(nf) } });
        });
      }

      if (percorso === '/api/indice' || percorso === '/api/presenze') {
        var file = 'dati/' + percorso.slice(5) + '.json';
        if (metodo === 'GET') {
          return leggiGrezzo(file).then(function (buf) {
            return risposta(buf ? new TextDecoder('utf-8').decode(buf) : 'null');
          });
        }
        var testo = typeof corpo === 'string' ? corpo : '';
        return scrivi(file, new TextEncoder().encode(testo).buffer,
          percorso === '/api/indice' ? 'Aggiornamento dell\'indice' : 'Aggiornamento delle presenze')
          .then(function () { return risposta({ ok: true }); });
      }

      if (percorso === '/api/cestino') {
        return elenco('cestino').then(function (voci) {
          var files = voci.sort(function (a, b) { return a.name < b.name ? 1 : -1; }).map(function (v) {
            var p = voceCestino(v.name);
            return { nome: v.name, originale: p.originale, rimossoIl: p.quando, dimensione: v.size };
          });
          return risposta({ files: files, cartella: 'cestino/ del repository ' + imp.repository });
        });
      }

      if (percorso === '/api/upload') {
        var nu = q('nome');
        if (!nomeValido(nu) || !ESTENSIONI.test(nu)) return errore('nome non valido', 400);
        return elenco('verbali').then(function (voci) {
          var esiste = voci.some(function (v) { return v.name === nu; });
          if (esiste && q('sovrascrivi') !== '1') return errore('esiste', 409, { nome: nu });
          var buf = corpo instanceof ArrayBuffer ? corpo : (corpo && corpo.buffer ? corpo.buffer : corpo);
          return scrivi('verbali/' + nu, buf, 'Acquisizione: ' + nu)
            .then(function () { return risposta({ ok: true, nome: nu }); });
        });
      }

      if (percorso === '/api/rimuovi') {
        var nr = q('nome');
        if (!nomeValido(nr)) return errore('nome non valido', 400);
        return Promise.all([elenco('verbali'), elenco('cestino')]).then(function (ris) {
          var v = ris[0].filter(function (x) { return x.name === nr; })[0];
          if (!v) return errore('file inesistente', 404);
          var occupati = {};
          ris[1].forEach(function (x) { occupati[x.name] = true; });
          var dest = marcaTemporale() + '__' + nr, i = 2;
          while (occupati[dest]) { dest = dest.replace(/( \(\d+\))?(\.[^.]+)$/, ' (' + (i++) + ')$2'); }
          return commitAlbero([
            { path: 'cestino/' + dest, sha: v.sha },
            { path: 'verbali/' + nr, sha: null }
          ], 'Rimozione nel cestino: ' + nr).then(function () {
            return risposta({ ok: true, nome: nr, cestino: dest });
          });
        });
      }

      if (percorso === '/api/ripristina') {
        var ns = q('nome');
        if (!nomeValido(ns)) return errore('nome non valido', 400);
        return Promise.all([elenco('cestino'), elenco('verbali')]).then(function (ris) {
          var c = ris[0].filter(function (x) { return x.name === ns; })[0];
          if (!c) return errore('file inesistente', 404);
          var orig = voceCestino(ns).originale;
          var esiste = ris[1].some(function (x) { return x.name === orig; });
          if (esiste && q('sovrascrivi') !== '1') return errore('esiste', 409, { nome: orig });
          return commitAlbero([
            { path: 'verbali/' + orig, sha: c.sha },
            { path: 'cestino/' + ns, sha: null }
          ], 'Ripristino dal cestino: ' + orig).then(function () {
            return risposta({ ok: true, nome: orig });
          });
        });
      }

      if (percorso === '/api/elimina') {
        var ne = q('nome');
        if (!nomeValido(ne)) return errore('nome non valido', 400);
        return elenco('cestino').then(function (voci) {
          if (!voci.some(function (x) { return x.name === ne; })) return errore('file inesistente', 404);
          return commitAlbero([{ path: 'cestino/' + ne, sha: null }], 'Cancellazione definitiva: ' + ne)
            .then(function () { return risposta({ ok: true, nome: ne }); });
        });
      }

      if (percorso === '/api/apri') {
        var na = q('nome');
        if (!na) { window.open(paginaGitHub(null), '_blank'); return risposta({ ok: true }); }
        if (!nomeValido(na)) return errore('nome non valido', 400);
        var base = q('da') === 'cestino' ? 'cestino/' : 'verbali/';
        if (q('cartella') === '1') { window.open(paginaGitHub(base + na), '_blank'); return risposta({ ok: true }); }
        return apriFile(base + na, q('da') === 'cestino' ? voceCestino(na).originale : na);
      }

      if (percorso === '/api/chiudi') return risposta({ ok: true });

      return errore('richiesta non riconosciuta', 404);
    }).catch(function (e) {
      if (e && e.status === 401) segnalaChiave();
      return errore(e.message || String(e), e.status || 500);
    });
  }

  /* Le aperture di file richiedono che la scheda nasca dentro il clic:
     per questo la gestione di /api/apri parte in modo sincrono quando le
     impostazioni sono già presenti. */
  window.fetch = function (risorsa, opzioni) {
    var url = typeof risorsa === 'string' ? risorsa : (risorsa && risorsa.url) || '';
    if (/(^|\/)api\//.test(url) && !/^https?:\/\/api\.github\.com/.test(url)) {
      if (imp && /\/api\/apri/.test(url)) {
        try { return gestisciSincrono(url, opzioni); } catch (e) { }
      }
      return gestisci(url, opzioni);
    }
    return fetchOriginale(risorsa, opzioni);
  };

  function gestisciSincrono(url, opzioni) {
    var u = new URL(url, location.href);
    var na = u.searchParams.get('nome') || '';
    if (na && u.searchParams.get('cartella') !== '1' && nomeValido(na)) {
      var base = u.searchParams.get('da') === 'cestino' ? 'cestino/' : 'verbali/';
      var nomeVis = base === 'cestino/' ? voceCestino(na).originale : na;
      return apriFile(base + na, nomeVis);
    }
    return gestisci(url, opzioni);
  }

  /* ------------------------------------------- finestra di collegamento */

  var attesa = null;

  function pronto() {
    if (imp) return Promise.resolve();
    if (!attesa) attesa = new Promise(function (risolvi) { mostraFinestra(risolvi); });
    return attesa;
  }

  function segnalaChiave() {
    var b = document.getElementById('gh-avviso-chiave');
    if (b) return;
    b = document.createElement('div');
    b.id = 'gh-avviso-chiave';
    b.className = 'gh-avviso';
    b.innerHTML = 'GitHub non accetta la chiave di accesso di questo dispositivo, forse perché è scaduta. ' +
      '<button type="button" class="bottone minuto">Aggiorna il collegamento</button>';
    b.querySelector('button').addEventListener('click', function () { apriImpostazioni(); });
    document.body.insertBefore(b, document.body.firstChild);
  }

  function mostraFinestra(risolvi, precedenti) {
    var p = precedenti || imp || { proprietario: 'giovannilanzo-blip', repository: 'archivio-verbali-dati', chiave: '' };
    var velo = document.createElement('div');
    velo.className = 'gh-velo';
    velo.innerHTML =
      '<form class="gh-dialogo" autocomplete="off">' +
      '<h2>Collegamento all\'archivio</h2>' +
      '<p>I verbali sono custoditi in un repository privato su GitHub. Su questo dispositivo occorre ' +
      'indicarlo una sola volta, insieme alla chiave di accesso.</p>' +
      '<label>Proprietario<input name="proprietario" required value="' + attr(p.proprietario) + '"></label>' +
      '<label>Repository dei dati<input name="repository" required value="' + attr(p.repository) + '"></label>' +
      '<label>Chiave di accesso (fine-grained token)<input name="chiave" type="password" required ' +
      'placeholder="github_pat_…" value="' + attr(p.chiave) + '"></label>' +
      '<p class="gh-errore" hidden></p>' +
      '<div class="gh-azioni">' +
      (imp ? '<button type="button" class="bottone" data-gh="scollega">Scollega questo dispositivo</button>' +
        '<button type="button" class="bottone" data-gh="annulla">Annulla</button>' : '') +
      '<button type="submit" class="bottone primario">Collega</button></div>' +
      '</form>';
    document.body.appendChild(velo);

    var form = velo.querySelector('form');
    var msg = velo.querySelector('.gh-errore');

    velo.addEventListener('click', function (e) {
      var a = e.target.dataset ? e.target.dataset.gh : null;
      if (a === 'annulla') velo.remove();
      if (a === 'scollega') {
        if (!confirm('Cancellare da questo dispositivo la chiave di accesso? I verbali restano su GitHub.')) return;
        cancellaImpostazioni();
        location.reload();
      }
    });

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var nuovo = {
        proprietario: form.proprietario.value.trim(),
        repository: form.repository.value.trim(),
        chiave: form.chiave.value.trim()
      };
      msg.hidden = true;
      var bottone = form.querySelector('[type=submit]');
      bottone.disabled = true; bottone.textContent = 'Verifica in corso…';
      var vecchio = imp;
      imp = nuovo; ramo = null; shaNoti = {};
      ghJson('GET', repo()).then(function (r) {
        if (!r.permissions || !r.permissions.push) {
          throw new Error('la chiave consente la lettura ma non la scrittura: in GitHub va concesso "Contents: Read and write"');
        }
        if (!r.private) {
          throw new Error('il repository risulta pubblico: i verbali contengono dati personali e il repository deve essere privato');
        }
        ramo = r.default_branch || 'main';
        scriviImpostazioni(nuovo);
        velo.remove();
        if (risolvi) risolvi(); else location.reload();
      }).catch(function (err) {
        imp = vecchio;
        msg.textContent = 'Collegamento non riuscito: ' + err.message + '.';
        msg.hidden = false;
        bottone.disabled = false; bottone.textContent = 'Collega';
      });
    });
  }

  function attr(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  }

  function apriImpostazioni() { mostraFinestra(null, imp); }

  window.ArchivioGitHub = {
    impostazioni: apriImpostazioni,
    paginaArchivio: function () { return imp ? paginaGitHub(null) : null; }
  };
})();
