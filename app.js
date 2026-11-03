// ============================================================
// RAPPORTINO DI CANTIERE DIGITALE (Mod. 22.03)
// - Caposquadra (telefono): compila la pagina 1 e la invia. Dopo l'invio non la vede più.
// - Responsabile (PC): riceve i rapportini, scrive le ore (pagina 2), convalida e scrive su Excel.
// I dati passano da Firebase (Firestore); le regole di sicurezza stanno in firebase/firestore.rules.
// Gli elenchi di operai e commesse arrivano dal file Excel Master_Operai (mai scritti nel codice).
// ============================================================

import { firebaseConfig, DOMINIO_ACCESSI } from './firebase-config.js';
import { initializeApp, deleteApp } from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js';
import {
  getAuth, initializeAuth, inMemoryPersistence, onAuthStateChanged, signInWithEmailAndPassword,
  createUserWithEmailAndPassword, signOut, connectAuthEmulator
} from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js';
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager, doc, collection, getDoc,
  getDocFromCache, getDocs, setDoc, updateDoc, deleteDoc, onSnapshot, query, where, serverTimestamp, writeBatch,
  waitForPendingWrites, connectFirestoreEmulator
} from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore.js';

// ---------------------------------------------------------------- utilità
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const oggiISO = () => new Date().toLocaleDateString('sv-SE'); // AAAA-MM-GG
const dataIT = iso => { const [a, m, g] = String(iso || '').split('-'); return g ? `${g}/${m}/${a}` : ''; };
const dataOraIT = ts => {
  const d = ts?.toDate ? ts.toDate() : (ts ? new Date(ts) : null);
  return d ? d.toLocaleString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
};
const num = v => { const n = parseFloat(String(v ?? '').replace(',', '.')); return isNaN(n) ? 0 : n; };
const fmtOre = n => (Math.round(n * 100) / 100).toString().replace('.', ',');
const MESI = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
const VOCI_ORE = [['or', 'Ord. D'], ['sd', 'Straord. D'], ['on', 'Ord. N'], ['sn', 'Straord. N'], ['v', 'Viaggio'], ['rol', 'ROL']];
const ls = {
  get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (_) { return null; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (_) {} },
  del(k) { try { localStorage.removeItem(k); } catch (_) {} }
};

function avviso(testo, tipo = '') {
  const el = document.createElement('div');
  el.className = `avviso-pop ${tipo}`;
  el.textContent = testo;
  $('avvisi').appendChild(el);
  setTimeout(() => el.remove(), tipo === 'errore' ? 7000 : 4000);
}

// Finestra di dialogo: restituisce il valore del bottone premuto (o null se chiusa)
function finestra(titolo, html, bottoni = [{ testo: 'OK', valore: true, primario: true }]) {
  $('finestraTitolo').textContent = titolo;
  $('finestraCorpo').innerHTML = html;
  const box = $('finestraBottoni');
  box.innerHTML = '';
  $('finestra').hidden = false;
  return new Promise(resolve => {
    bottoni.forEach(b => {
      const btn = document.createElement('button');
      btn.className = `btn ${b.primario ? 'primario' : ''} ${b.pericolo ? 'pericolo' : ''}`;
      btn.textContent = b.testo;
      btn.onclick = () => { $('finestra').hidden = true; resolve(b.valore); };
      box.appendChild(btn);
    });
  });
}
const conferma = (titolo, html, si = 'Conferma') =>
  finestra(titolo, html, [{ testo: 'Annulla', valore: false }, { testo: si, valore: true, primario: true }]);

function mostraSchermo(id) {
  ['scrAvvio', 'scrConfig', 'scrLogin', 'scrNonAbilitato', 'scrApp'].forEach(s => { $(s).hidden = s !== id; });
}

const codiceValido = c => /^[a-z0-9._-]{2,40}$/.test(c);
const normCodice = c => String(c || '').trim().toLowerCase();
const emailDi = codice => `${normCodice(codice)}@${DOMINIO_ACCESSI}`;

function messaggioErrore(err) {
  const c = err?.code || '';
  if (/invalid-credential|wrong-password|user-not-found|invalid-email/.test(c)) return 'Codice o password errati';
  if (/network-request-failed|unavailable/.test(c)) return 'Nessuna connessione a internet: riprova quando c\'è campo';
  if (/too-many-requests/.test(c)) return 'Troppi tentativi: aspetta qualche minuto e riprova';
  if (/email-already-in-use/.test(c)) return 'Questo codice utente esiste già: scegline un altro';
  if (/weak-password/.test(c)) return 'Password troppo corta: almeno 6 caratteri';
  if (/permission-denied/.test(c)) return 'Operazione non permessa per questo utente';
  if (/operation-not-allowed/.test(c)) return 'Accesso con password non attivato in Firebase (vedi la guida)';
  return err?.message || String(err);
}

// ---------------------------------------------------------------- avvio
// Intestazione come sul modulo cartaceo (pagina 1 completa, pagina 2 solo la riga di identificazione).
// I testi e il logo si cambiano in "Elenchi e utenti > Intestazione del modulo"; questi sono i valori di partenza.
const INTESTAZIONE_BASE = {
  sedeLegale: 'Sede legale: Viale Michelangelo n. 33 – 80127 Napoli (NA)',
  sedeOperativa: 'Sede operativa: Strada Consortile SNC - 81032 Carinaro (CE)',
  telefono: 'Tel. e Fax 081/502.97.30',
  sito: 'www.vicla.it',
  titolo: 'MODELLO DEL SISTEMA DI GESTIONE INTEGRATO QUALITA’, AMBIENTE E SICUREZZA',
  sottotitolo: 'RAPPORTO DI CANTIERE',
  identificazione: 'Mod. 22.03',
  revisione: '03',
  dataRevisione: '08/01/2016',
  norma1: 'UNI EN ISO 9001:2015',
  norma2: 'UNI EN ISO 14001:2015',
  norma3: 'BS OHSAS 18001:2007'
};
const LOGO_BASE = 'icons/logo-vicla.png';
let intestazione = { ...INTESTAZIONE_BASE, logo: '' };

function testataModulo(pagina, t = intestazione) {
  const griglia = `<div class="t-griglia">
      <div>Identificazione: ${esc(t.identificazione)}</div><div>Revisione: ${esc(t.revisione)}</div><div>Data: ${esc(t.dataRevisione)}</div><div>Pagina ${pagina} di 2</div>
      <div>Rif. Norme</div><div>${esc(t.norma1)}</div><div>${esc(t.norma2)}</div><div>${esc(t.norma3)}</div>
    </div>`;
  if (pagina === 2) return `<div class="testata">${griglia}</div>`;
  // il logo è un'immagine caricata dall'ufficio (data:image/...) oppure quello originale dell'app
  const logo = /^data:image\/(png|jpeg|webp);base64,/.test(t.logo || '') ? t.logo : LOGO_BASE;
  return `<div class="testata">
    <div class="t-alto">
      <div class="t-logo"><img src="${logo}" alt="Logo"></div>
      <div class="t-sede">
        ${[t.sedeLegale, t.sedeOperativa, t.telefono, t.sito].filter(Boolean).map(r => `<div>${esc(r)}</div>`).join('')}
      </div>
    </div>
    <div class="t-titolo">${esc(t.titolo)}</div>
    <div class="t-sotto">${esc(t.sottotitolo)}</div>
    ${griglia}
  </div>`;
}
// Le intestazioni del modulo in compilazione e della pagina 2 seguono sempre l'intestazione attuale
function aggiornaTestate() {
  document.querySelectorAll('.testata-pos').forEach(el => { el.innerHTML = testataModulo(Number(el.dataset.pagina)); });
}
aggiornaTestate();
$('versioneApp').textContent = `Versione ${self.APP_VERSION} del ${self.APP_VERSION_DATE} · Novità`;
$('versioneApp').onclick = () => finestra('Novità delle versioni', (self.APP_STORICO || []).map(v => `
  <h3>Versione ${esc(v.versione)} <span class="aiuto">del ${esc(v.data)}</span></h3>
  <ul>${v.novita.map(n => `<li>${esc(n)}</li>`).join('')}</ul>`).join(''));
if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

// Prove sul PC: con l'indirizzo http://localhost:...?emulatore l'app usa il simulatore di Firebase
// (dati finti, nessun account). Online questa modalità non si attiva mai.
const EMULATORE = new URLSearchParams(location.search).has('emulatore') && ['localhost', '127.0.0.1'].includes(location.hostname);
const config = EMULATORE
  ? { apiKey: 'demo-key', authDomain: 'demo-rapportino.firebaseapp.com', projectId: 'demo-rapportino', appId: 'demo' }
  : firebaseConfig;
const configurato = !Object.values(config).some(v => String(v).includes('DA-COMPILARE'));
let app, auth, db;
if (!configurato) {
  mostraSchermo('scrConfig');
} else {
  app = initializeApp(config);
  auth = getAuth(app);
  // Cache sul telefono: gli elenchi si vedono anche senza campo e i rapportini inviati senza campo
  // restano in coda e partono da soli appena torna la connessione
  db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });
  if (EMULATORE) {
    connectAuthEmulator(auth, 'http://127.0.0.1:9099', { disableWarnings: true });
    connectFirestoreEmulator(db, '127.0.0.1', 8080);
    document.title += ' (PROVA)';
  }
  onAuthStateChanged(auth, utenteCambiato);
}

// ---------------------------------------------------------------- stato
const stato = {
  utente: null,          // { uid, nome, ruolo, codice }
  elenchi: { operai: [], commesse: [] },
  presenti: new Map(),   // nome -> { partenza, dalle, alle, rientro }
  automezzi: [],         // elenco dei mezzi (dal database)
  mezziScelti: new Set(), // mezzi spuntati nel rapportino in compilazione
  arrivati: [],
  aperto: null,          // rapportino aperto dal responsabile: { id, dati, blocchi }
  cartellaPronta: false,
  annulla: []            // funzioni per chiudere gli ascolti di Firestore all'uscita
};

// ---------------------------------------------------------------- accesso
async function utenteCambiato(user) {
  stato.annulla.forEach(f => { try { f(); } catch (_) {} });
  stato.annulla = [];
  if (!user) {
    stato.utente = null;
    mostraSchermo('scrLogin');
    return;
  }
  let profilo;
  try {
    const ref = doc(db, 'utenti', user.uid);
    const snap = await getDoc(ref).catch(() => getDocFromCache(ref));
    profilo = snap.exists() ? snap.data() : null;
  } catch (err) {
    $('nonAbilitatoTesto').textContent = 'Non riesco a leggere il tuo profilo. Al primo accesso serve internet: riprova quando c\'è campo.';
    mostraSchermo('scrNonAbilitato');
    return;
  }
  if (!profilo || profilo.attivo !== true) {
    $('nonAbilitatoTesto').textContent = 'Questo utente non è abilitato a usare il rapportino. Rivolgiti all\'ufficio.';
    mostraSchermo('scrNonAbilitato');
    return;
  }
  stato.utente = { uid: user.uid, nome: profilo.nome, ruolo: profilo.ruolo, codice: profilo.codice };
  $('barraUtente').textContent = `${profilo.nome} · ${profilo.ruolo === 'responsabile' ? 'Responsabile' : 'Caposquadra'}`;
  mostraSchermo('scrApp');
  avviaCompilazione();
  if (profilo.ruolo === 'responsabile') {
    $('menuResponsabile').hidden = false;
    avviaResponsabile();
    mostraVista('arrivati');
  } else {
    $('menuResponsabile').hidden = true;
    mostraVista('compila');
  }
}

$('formLogin').addEventListener('submit', async e => {
  e.preventDefault();
  const codice = normCodice($('loginCodice').value);
  $('loginErrore').hidden = true;
  $('btnLogin').disabled = true;
  try {
    await signInWithEmailAndPassword(auth, emailDi(codice), $('loginPassword').value);
    $('loginPassword').value = '';
  } catch (err) {
    $('loginErrore').textContent = messaggioErrore(err);
    $('loginErrore').hidden = false;
  } finally {
    $('btnLogin').disabled = false;
  }
});

$('btnApriSetup').onclick = () => { $('formLogin').hidden = true; $('formSetup').hidden = false; };
$('btnChiudiSetup').onclick = () => { $('formSetup').hidden = true; $('formLogin').hidden = false; };

// Prima configurazione: crea l'account e, nello stesso momento, il profilo responsabile e il segno "setup fatto".
// Le regole permettono questa operazione una volta sola.
$('formSetup').addEventListener('submit', async e => {
  e.preventDefault();
  const nome = $('setupNome').value.trim();
  const codice = normCodice($('setupCodice').value);
  $('setupErrore').hidden = true;
  if (!codiceValido(codice)) {
    $('setupErrore').textContent = 'Codice utente: solo lettere minuscole, numeri, punto, trattino (2-40 caratteri), senza spazi';
    $('setupErrore').hidden = false;
    return;
  }
  $('btnSetup').disabled = true;
  try {
    const cred = await createUserWithEmailAndPassword(auth, emailDi(codice), $('setupPassword').value);
    try {
      const batch = writeBatch(db);
      batch.set(doc(db, 'utenti', cred.user.uid), { nome, codice, ruolo: 'responsabile', attivo: true, creatoIl: serverTimestamp() });
      batch.set(doc(db, 'config', 'setup'), { fattoIl: serverTimestamp(), da: cred.user.uid });
      await batch.commit();
      utenteCambiato(cred.user); // il profilo ora esiste: rilegge e apre l'app
    } catch (err) {
      await cred.user.delete().catch(() => {});
      throw /permission-denied/.test(err?.code) ? new Error('La prima configurazione è già stata fatta: entra con il tuo codice') : err;
    }
    $('formSetup').hidden = true;
    $('formLogin').hidden = false;
  } catch (err) {
    $('setupErrore').textContent = messaggioErrore(err);
    $('setupErrore').hidden = false;
  } finally {
    $('btnSetup').disabled = false;
  }
});

async function esci() {
  const coda = ls.get(chiaveCoda()) || 0;
  if (coda > 0) {
    const ok = await conferma('Rapportini non ancora partiti',
      `<p>Hai <b>${coda}</b> rapportino/i non ancora arrivato/i al responsabile (manca la connessione).</p>
       <p>Se esci adesso <b>potrebbero andare persi</b>. Aspetta di avere campo, poi esci.</p>`, 'Esci lo stesso');
    if (!ok) return;
    ls.del(chiaveCoda());
  }
  await signOut(auth);
}
$('btnEsci').onclick = esci;
$('btnEsciNonAbilitato').onclick = () => signOut(auth);

// ---------------------------------------------------------------- viste
const VISTE = { compila: 'vistaCompila', inviato: 'vistaInviato', arrivati: 'vistaArrivati', convalidati: 'vistaConvalidati', dettaglio: 'vistaDettaglio', gestione: 'vistaGestione' };
function mostraVista(nome) {
  Object.entries(VISTE).forEach(([k, id]) => { $(id).hidden = k !== nome; });
  document.querySelectorAll('.scheda-btn').forEach(b => b.classList.toggle('attiva', b.dataset.vista === nome || (nome === 'dettaglio' && b.dataset.vista === 'arrivati') || (nome === 'inviato' && b.dataset.vista === 'compila')));
  window.scrollTo(0, 0);
  // rapportino aperto: tutta la larghezza dello schermo per le due pagine affiancate, ognuna dall'inizio
  document.querySelector('.contenuto').classList.toggle('largo', nome === 'dettaglio');
  if (nome === 'dettaglio') document.querySelectorAll('.affiancati > .modulo').forEach(m => { m.scrollTop = 0; });
  if (nome === 'convalidati') caricaConvalidati();
}
document.querySelectorAll('.scheda-btn').forEach(b => { b.onclick = () => mostraVista(b.dataset.vista); });

// ================================================================
// PAGINA 1 · COMPILAZIONE (caposquadra, e anche responsabile)
// ================================================================
const chiaveBozza = () => `rapp_bozza_${stato.utente?.uid}`;
const chiaveCoda = () => `rapp_coda_${stato.utente?.uid}`;
const CAMPI_TESTO = { localita: 'rLocalita', responsabile: 'rResponsabile', mensaDalle: 'rMensaDalle', mensaAlle: 'rMensaAlle', lavorazioni: 'rLavorazioni', notizie: 'rNotizie', controlli: 'rControlli' };

function avviaCompilazione() {
  $('rCaposquadra').value = stato.utente.nome;
  $('firmaCapo').textContent = `${stato.utente.nome} (registrata all'invio)`;
  caricaBozza();
  const off = onSnapshot(doc(db, 'elenchi', 'correnti'), snap => {
    const d = snap.exists() ? snap.data() : {};
    stato.elenchi = { operai: d.operai || [], commesse: d.commesse || [], nascoste: d.nascoste || [], file: d.file || '', aggiornatoIl: d.aggiornatoIl };
    $('erroreElenchi').hidden = stato.elenchi.operai.length > 0;
    $('erroreElenchi').textContent = 'Gli elenchi di operai e commesse non sono ancora stati caricati dall\'ufficio (Elenchi e utenti > Aggiorna elenchi).';
    riempiCommesse();
    disegnaOperai();
    mostraInfoElenchi();
  }, err => avviso(`Elenchi non disponibili: ${messaggioErrore(err)}`, 'errore'));
  stato.annulla.push(off);
  // elenco automezzi: lo gestisce il responsabile in "Elenchi e utenti"
  const offM = onSnapshot(doc(db, 'elenchi', 'automezzi'), snap => {
    stato.automezzi = snap.exists() ? (snap.data().voci || []) : [];
    if (snap.exists() || !snap.metadata.fromCache) {
      for (const m of [...stato.mezziScelti]) if (!stato.automezzi.includes(m)) stato.mezziScelti.delete(m);
    }
    disegnaMezzi();
    disegnaGestioneMezzi();
  }, () => {});
  stato.annulla.push(offM);
  // intestazione del modulo (testi e logo): la decide l'ufficio
  const offI = onSnapshot(doc(db, 'elenchi', 'intestazione'), snap => {
    intestazione = { ...INTESTAZIONE_BASE, logo: '', ...(snap.exists() ? snap.data() : {}) };
    aggiornaTestate();
    disegnaGestioneIntestazione();
  }, () => {});
  stato.annulla.push(offI);
  controllaCoda();
}

// Automezzi: elenco a comparsa; ogni mezzo scelto compare sotto con la ✕ per toglierlo (se ne possono scegliere più)
function disegnaMezzi() {
  const box = $('rAutomezzi');
  if (!stato.automezzi.length) {
    box.innerHTML = '<span class="tenue">Nessun automezzo in elenco (lo inserisce l\'ufficio)</span>';
    return;
  }
  const scelti = stato.automezzi.filter(m => stato.mezziScelti.has(m));
  const liberi = stato.automezzi.filter(m => !stato.mezziScelti.has(m));
  box.innerHTML = `
    ${liberi.length ? `<select id="sceltaMezzo" aria-label="Aggiungi automezzo">
      <option value="">${scelti.length ? '+ Aggiungi un altro automezzo' : '— Scegli l\'automezzo —'}</option>
      ${liberi.map(m => `<option>${esc(m)}</option>`).join('')}
    </select>` : ''}
    <div class="mezzi-scelti">${scelti.map(m => `<span class="chip si">${esc(m)}
      <button type="button" class="togli-mezzo" data-togli="${esc(m)}" aria-label="Togli ${esc(m)}">✕</button></span>`).join('')}</div>`;
}
$('rAutomezzi').addEventListener('change', e => {
  if (e.target.id !== 'sceltaMezzo' || !e.target.value) return;
  stato.mezziScelti.add(e.target.value);
  disegnaMezzi();
  salvaBozza();
});
$('rAutomezzi').addEventListener('click', e => {
  const m = e.target.dataset.togli;
  if (m === undefined) return;
  stato.mezziScelti.delete(m);
  disegnaMezzi();
  salvaBozza();
});

function riempiCommesse() {
  const sel = $('rCommessa');
  const attuale = sel.value || ls.get(chiaveBozza())?.commessa || '';
  // sul telefono solo le commesse vere (le assenze le gestisce il responsabile)
  const voci = stato.elenchi.commesse.filter(c => !stato.elenchi.nascoste.includes(c));
  if (attuale && !voci.includes(attuale)) voci.unshift(attuale);
  sel.innerHTML = '<option value="">— Scegli la commessa —</option>' + voci.map(c => `<option ${c === attuale ? 'selected' : ''}>${esc(c)}</option>`).join('');
}

function disegnaOperai() {
  const box = $('listaOperai');
  if (stato.elenchi.operai.length === 0) {
    // elenco non ancora arrivato: le X della bozza restano da parte
    box.innerHTML = '<div class="nessuno">Elenco operai in arrivo…</div>';
    aggiornaContaPresenti();
    return;
  }
  // chi era segnato ma non è più nell'elenco viene tolto
  for (const nome of [...stato.presenti.keys()]) if (!stato.elenchi.operai.includes(nome)) stato.presenti.delete(nome);
  box.innerHTML = stato.elenchi.operai.map((nome, i) => rigaOperaio(nome, i)).join('');
  filtraOperai();
  aggiornaContaPresenti();
}

// Una riga della tabella come sul modulo: nome, casella P, partenza, Dalle/Alle, rientro
function rigaOperaio(nome, i) {
  const o = stato.presenti.get(nome);
  const v = k => esc(o?.[k] || '');
  return `<div class="operaio ${o ? 'presente' : ''}" data-nome="${esc(nome)}">
    <div class="o-nome" data-tocca><span class="o-num">${i + 1}.</span>${esc(nome)}</div>
    <button type="button" class="o-p casella-p" data-tocca aria-pressed="${!!o}" aria-label="Presente: ${esc(nome)}"></button>
    <div class="o-ora o-partenza"><span class="o-et">Partenza deposito</span><input type="time" data-orario="partenza" value="${v('partenza')}"></div>
    <div class="o-ora o-cantiere"><span class="o-et">Dalle</span><input type="time" data-orario="dalle" value="${v('dalle')}"><span class="o-et">Alle</span><input type="time" data-orario="alle" value="${v('alle')}"></div>
    <div class="o-ora o-rientro"><span class="o-et">Rientro deposito</span><input type="time" data-orario="rientro" value="${v('rientro')}"></div>
  </div>`;
}

function segnaPresente(div) {
  const nome = div.dataset.nome;
  if (stato.presenti.has(nome)) stato.presenti.delete(nome);
  else stato.presenti.set(nome, { partenza: '', dalle: '', alle: '', rientro: '' });
  const si = stato.presenti.has(nome);
  div.classList.toggle('presente', si);
  div.querySelector('.casella-p').setAttribute('aria-pressed', si);
  if (!si) div.querySelectorAll('[data-orario]').forEach(inp => { inp.value = ''; });
  aggiornaContaPresenti();
  salvaBozza();
}

$('listaOperai').addEventListener('click', e => {
  if (e.target.closest('[data-tocca]')) segnaPresente(e.target.closest('.operaio'));
});$('listaOperai').addEventListener('input', e => {
  const k = e.target.dataset.orario;
  if (!k) return;
  const nome = e.target.closest('.operaio').dataset.nome;
  const o = stato.presenti.get(nome);
  if (o) { o[k] = e.target.value; salvaBozza(); }
});

function filtraOperai() {
  const f = $('cercaOperaio').value.trim().toLowerCase();
  document.querySelectorAll('#listaOperai .operaio').forEach(div => {
    div.hidden = !!f && !div.dataset.nome.toLowerCase().includes(f);
  });
}
$('cercaOperaio').addEventListener('input', filtraOperai);

function aggiornaContaPresenti() {
  $('contaPresenti').textContent = stato.presenti.size;
  $('orariTutti').hidden = stato.presenti.size < 2;
}

$('btnApplicaTutti').onclick = () => {
  const valori = {};
  document.querySelectorAll('[data-tutti]').forEach(inp => { if (inp.value) valori[inp.dataset.tutti] = inp.value; });
  if (Object.keys(valori).length === 0) { avviso('Scrivi prima almeno un orario qui sopra', 'errore'); return; }
  for (const o of stato.presenti.values()) Object.assign(o, valori);
  document.querySelectorAll('#listaOperai .operaio.presente').forEach(div => {
    div.querySelectorAll('[data-orario]').forEach(inp => { if (valori[inp.dataset.orario]) inp.value = valori[inp.dataset.orario]; });
  });
  salvaBozza();
  avviso(`Orari applicati a ${stato.presenti.size} operai`, 'ok');
};

// Bozza sul telefono finché il rapportino non viene inviato (se si chiude l'app non si perde il lavoro)
function salvaBozza() {
  if (!stato.utente) return;
  const b = { commessa: $('rCommessa').value, data: $('rData').value, presenti: Object.fromEntries(stato.presenti), automezzi: [...stato.mezziScelti] };
  for (const [k, id] of Object.entries(CAMPI_TESTO)) b[k] = $(id).value;
  ls.set(chiaveBozza(), b);
}
function caricaBozza() {
  const b = ls.get(chiaveBozza()) || {};
  $('rData').value = b.data || oggiISO();
  for (const [k, id] of Object.entries(CAMPI_TESTO)) $(id).value = b[k] || '';
  stato.presenti = new Map(Object.entries(b.presenti || {}));
  stato.mezziScelti = new Set(Array.isArray(b.automezzi) ? b.automezzi : []);
  riempiCommesse();
  disegnaOperai();
  disegnaMezzi();
}
function svuotaModulo() {
  ls.del(chiaveBozza());
  $('rCommessa').value = '';
  document.querySelectorAll('[data-tutti]').forEach(inp => { inp.value = ''; });
  $('cercaOperaio').value = '';
  caricaBozza();
}
['rCommessa', 'rData', ...Object.values(CAMPI_TESTO)].forEach(id => $(id).addEventListener('input', salvaBozza));
$('rCommessa').addEventListener('change', salvaBozza);

$('btnSvuota').onclick = async () => {
  if (await conferma('Svuotare il rapportino?', '<p>Verranno tolti tutti i dati scritti finora.</p>', 'Svuota')) svuotaModulo();
};

$('formRapportino').addEventListener('submit', async e => {
  e.preventDefault();
  const commessa = $('rCommessa').value;
  const data = $('rData').value;
  const errori = [];
  if (!commessa) errori.push('scegli la <b>commessa</b>');
  if (!data) errori.push('scrivi la <b>data</b>');
  if (stato.presenti.size === 0) errori.push('segna almeno un <b>operaio presente</b> (casella P)');
  if (data && data > oggiISO()) errori.push('la data è nel futuro: controlla');
  if (errori.length) {
    await finestra('Manca qualcosa', `<ul>${errori.map(x => `<li>${x}</li>`).join('')}</ul>`);
    return;
  }

  // ordine dell'elenco, come sul modulo
  const operai = stato.elenchi.operai.filter(n => stato.presenti.has(n)).map(nome => ({ nome, ...stato.presenti.get(nome) }));
  const ok = await conferma('Inviare il rapportino?', `
    <p><b>${dataIT(data)}</b> · commessa <b>${esc(commessa)}</b></p>
    <p>${operai.length} operai presenti:</p>
    <ul>${operai.map(o => `<li>${esc(o.nome)}${o.dalle || o.alle ? ` <span class="aiuto">(${esc(o.dalle)}–${esc(o.alle)})</span>` : ''}</li>`).join('')}</ul>
    <p class="errore">Dopo l'invio <b>non potrai più vederlo né modificarlo</b>.</p>`, 'Invia');
  if (!ok) return;

  const dati = {
    stato: 'arrivato',
    data,
    mese: data.slice(0, 7),
    commessa,
    localita: $('rLocalita').value.trim(),
    responsabileCommessa: $('rResponsabile').value.trim(),
    caposquadraUid: stato.utente.uid,
    caposquadraNome: stato.utente.nome,
    operai,
    automezzi: stato.automezzi.filter(m => stato.mezziScelti.has(m)), // nell'ordine dell'elenco
    // testi dell'intestazione in vigore al momento dell'invio (es. la revisione del modulo)
    intestazione: Object.fromEntries(Object.keys(INTESTAZIONE_BASE).map(k => [k, intestazione[k] ?? INTESTAZIONE_BASE[k]])),
    mensaDalle: $('rMensaDalle').value,
    mensaAlle: $('rMensaAlle').value,
    lavorazioni: $('rLavorazioni').value.trim(),
    altreNotizie: $('rNotizie').value.trim(),
    controlli: $('rControlli').value.trim(),
    compilatoIl: new Date().toISOString(),
    inviatoIl: serverTimestamp(),
    versioneApp: self.APP_VERSION
  };
  inviaRapportino(dati);
});

// L'invio parte subito; se manca il campo resta in coda sul telefono e parte da solo quando torna la rete.
// Il modulo si svuota in ogni caso: il caposquadra non rivede il rapportino inviato.
function inviaRapportino(dati) {
  const bozza = ls.get(chiaveBozza());
  const ref = doc(collection(db, 'rapportini'));
  const chiave = chiaveCoda();
  ls.set(chiave, (ls.get(chiave) || 0) + 1);
  let consegnato = false;

  const scrittura = setDoc(ref, dati).then(() => {
    consegnato = true;
    ls.set(chiave, Math.max(0, (ls.get(chiave) || 1) - 1));
    controllaCoda();
  });

  svuotaModulo();
  $('inviatoTitolo').textContent = 'Invio in corso…';
  $('inviatoTesto').textContent = '';
  mostraVista('inviato');

  const attesa = setTimeout(() => {
    if (consegnato) return;
    $('inviatoTitolo').textContent = 'Rapportino in partenza';
    $('inviatoTesto').textContent = 'Ora non c\'è connessione: il rapportino è conservato sul telefono e partirà da solo appena torna il campo. Non uscire dall\'app finché non parte.';
    controllaCoda();
  }, 6000);

  scrittura.then(() => {
    clearTimeout(attesa);
    $('inviatoTitolo').textContent = 'Rapportino inviato';
    $('inviatoTesto').textContent = `Il rapportino del ${dataIT(dati.data)} è arrivato al responsabile.`;
  }).catch(async err => {
    clearTimeout(attesa);
    ls.set(chiave, Math.max(0, (ls.get(chiave) || 1) - 1));
    controllaCoda();
    // invio rifiutato: il rapportino torna nel modulo per non perderlo
    if (bozza) { ls.set(chiaveBozza(), bozza); caricaBozza(); }
    mostraVista('compila');
    await finestra('Invio non riuscito', `<p>${esc(messaggioErrore(err))}</p><p>Il rapportino è stato rimesso nel modulo: puoi riprovare.</p>`);
  });
}

function controllaCoda() {
  const n = ls.get(chiaveCoda()) || 0;
  $('avvisoCoda').hidden = n === 0;
  $('avvisoCoda').textContent = `⏳ ${n} rapportino/i in attesa di connessione: partiranno da soli appena c'è campo.`;
  if (n > 0) {
    waitForPendingWrites(db).then(() => { ls.del(chiaveCoda()); $('avvisoCoda').hidden = true; }).catch(() => {});
  }
}

$('btnNuovo').onclick = () => mostraVista('compila');

// ================================================================
// RESPONSABILE
// ================================================================
let responsabileAvviato = false;
function avviaResponsabile() {
  // elenco dei rapportini arrivati, aggiornato in tempo reale
  const off = onSnapshot(query(collection(db, 'rapportini'), where('stato', '==', 'arrivato')), snap => {
    stato.arrivati = snap.docs.map(d => ({ id: d.id, ...d.data() }))
      .sort((a, b) => (b.data || '').localeCompare(a.data || '') || ((b.inviatoIl?.seconds || 0) - (a.inviatoIl?.seconds || 0)));
    disegnaArrivati();
  }, err => avviso(`Rapportini non disponibili: ${messaggioErrore(err)}`, 'errore'));
  stato.annulla.push(off);

  const offU = onSnapshot(collection(db, 'utenti'), snap => disegnaUtenti(snap.docs.map(d => ({ uid: d.id, ...d.data() }))));
  stato.annulla.push(offU);

  if (!$('meseConvalidati').value) $('meseConvalidati').value = oggiISO().slice(0, 7);
  if (!responsabileAvviato) {
    responsabileAvviato = true;
    CartellaExcel.restore().then(s => (s === 'pronta' ? cartellaPronta() : mostraAvvisoCartella(s)));
  }
}

function voceRapportino(r) {
  const extra = r.stato === 'convalidato'
    ? `<span class="etichetta ok">Convalidato ${esc(dataOraIT(r.convalidatoIl))}</span>`
    : `<span class="etichetta">Inviato ${esc(dataOraIT(r.inviatoIl))}</span>`;
  return `<button class="voce" data-id="${esc(r.id)}">
    <div class="voce-princ">
      <div class="voce-titolo">${esc(dataIT(r.data))} · ${esc(r.commessa)}${r.localita ? ` · ${esc(r.localita)}` : ''}</div>
      <div class="voce-sotto">Caposquadra ${esc(r.caposquadraNome)} · ${r.operai?.length || 0} operai</div>
    </div>${extra}</button>`;
}

function disegnaArrivati() {
  const n = stato.arrivati.length;
  $('contaArrivati').textContent = n;
  $('contaArrivati').classList.toggle('zero', n === 0);
  $('listaArrivati').innerHTML = n ? stato.arrivati.map(voceRapportino).join('') : '<div class="nessuno">Nessun rapportino da convalidare</div>';
}
$('listaArrivati').addEventListener('click', e => {
  const v = e.target.closest('.voce');
  if (v) apriRapportino(stato.arrivati.find(r => r.id === v.dataset.id));
});

let offConvalidati = null;
let convalidati = [];
function caricaConvalidati() {
  if (offConvalidati) offConvalidati();
  const mese = $('meseConvalidati').value;
  if (!mese) return;
  offConvalidati = onSnapshot(query(collection(db, 'rapportini'), where('stato', '==', 'convalidato'), where('mese', '==', mese)), snap => {
    convalidati = snap.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => (b.data || '').localeCompare(a.data || ''));
    $('listaConvalidati').innerHTML = convalidati.length ? convalidati.map(voceRapportino).join('') : '<div class="nessuno">Nessun rapportino convalidato in questo mese</div>';
  }, err => avviso(messaggioErrore(err), 'errore'));
  stato.annulla.push(() => offConvalidati && offConvalidati());
}
$('meseConvalidati').addEventListener('change', caricaConvalidati);
$('listaConvalidati').addEventListener('click', e => {
  const v = e.target.closest('.voce');
  if (v) apriRapportino(convalidati.find(r => r.id === v.dataset.id));
});

// ---------------------------------------------------------------- dettaglio
function orarioInOre(hhmm) {
  const m = String(hhmm || '').match(/^(\d{1,2}):(\d{2})/);
  return m ? parseInt(m[1]) + parseInt(m[2]) / 60 : null;
}
function differenzaOre(da, a) {
  const x = orarioInOre(da), y = orarioInOre(a);
  if (x === null || y === null) return 0;
  return y >= x ? y - x : y + 24 - x; // turno a cavallo della mezzanotte
}
const oreVuote = () => Object.fromEntries(VOCI_ORE.map(([k]) => [k, '']));

function bloccoIniziale(r) {
  const tutti = (r.operai || []).map(o => o.nome);
  const comune = k => { const s = new Set((r.operai || []).map(o => o[k]).filter(Boolean)); return s.size === 1 ? [...s][0] : ''; };
  const pausa = differenzaOre(r.mensaDalle, r.mensaAlle);
  return {
    righe: [
      { commessa: r.commessa || '', entrata: comune('dalle'), uscita: comune('alle'), pausa: pausa ? fmtOre(pausa) : '' },
      { commessa: '', entrata: '', uscita: '', pausa: '' }
    ],
    ore: { t1: oreVuote(), t2: oreVuote() },
    operai: tutti
  };
}

function apriRapportino(r) {
  if (!r) return;
  stato.aperto = { id: r.id, dati: r, blocchi: (r.blocchi?.length ? structuredClone(r.blocchi) : [bloccoIniziale(r)]) };
  disegnaFronte(r);
  disegnaBlocchi();
  const conv = r.stato === 'convalidato';
  $('btnConvalida').textContent = conv ? 'Riscrivi su Excel' : 'Convalida e scrivi su Excel';
  $('btnSalvaOre').hidden = conv;
  $('infoFileExcel').textContent = conv
    ? `Convalidato da ${r.convalidatoDaNome || ''} il ${dataOraIT(r.convalidatoIl)} · scritto in ${r.fileExcel || ''}`
    : '';
  mostraVista('dettaglio');
}
$('btnTornaElenco').onclick = () => mostraVista(stato.aperto?.dati?.stato === 'convalidato' ? 'convalidati' : 'arrivati');

// Pagina 1 in sola lettura, con la stessa grafica del modulo
function disegnaFronte(r) {
  const val = v => esc(v) || '&nbsp;';
  const testo = v => `<div class="r-valore testo">${val(v)}</div>`;
  $('dettaglioFronte').innerHTML = `
    ${testataModulo(1, { ...intestazione, ...(r.intestazione || {}) })}
    <div class="riquadro">
      <div class="r-riga">
        <div class="r-cella largo"><span class="r-et">COMMESSA:</span><div class="r-valore">${val(r.commessa)}</div></div>
        <div class="r-cella"><span class="r-et">DATA:</span><div class="r-valore">${val(dataIT(r.data))}</div></div>
      </div>
      <div class="r-riga"><div class="r-cella"><span class="r-et grande">LOCALIZZAZIONE DEL SITO:</span><div class="r-valore">${val(r.localita)}</div></div></div>
      <div class="r-riga">
        <div class="r-cella largo"><span class="r-et">Caposquadra:</span><div class="r-valore">${val(r.caposquadraNome)}</div></div>
        <div class="r-cella"><span class="r-et">Responsabile di Commessa:</span><div class="r-valore">${val(r.responsabileCommessa)}</div></div>
      </div>
    </div>
    <div class="tab-operai sola-lettura">
      <div class="operaio testa">
        <div class="o-nome">COGNOME E NOME</div><div class="o-p">P</div>
        <div class="o-ora o-partenza">ORARIO PARTENZA DA DEPOSITO</div>
        <div class="o-ora o-cantiere">ORARIO DI LAVORO IN CANTIERE</div>
        <div class="o-ora o-rientro">ORARIO RIENTRO IN DEPOSITO</div>
      </div>
      ${(r.operai || []).map((o, i) => `<div class="operaio presente">
        <div class="o-nome"><span class="o-num">${i + 1}.</span>${esc(o.nome)}</div>
        <div class="o-p casella-p"></div>
        <div class="o-ora o-partenza"><span class="o-et">Partenza deposito</span><b>${esc(o.partenza) || '…'}</b></div>
        <div class="o-ora o-cantiere"><span class="o-et">Dalle</span><b>${esc(o.dalle) || '…'}</b><span class="o-et">Alle</span><b>${esc(o.alle) || '…'}</b></div>
        <div class="o-ora o-rientro"><span class="o-et">Rientro deposito</span><b>${esc(o.rientro) || '…'}</b></div>
      </div>`).join('')}
    </div>
    <div class="riquadro">
      <div class="r-riga">
        <div class="r-cella largo"><span class="r-et">AUTOMEZZI UTILIZZATI:</span>${testo(Array.isArray(r.automezzi) ? r.automezzi.join('\n') : r.automezzi)}</div>
        <div class="r-cella"><span class="r-et normale">ORARIO DI PAUSA MENSA:</span>
          <div class="dalle-alle"><span>Dalle</span><b>${esc(r.mensaDalle) || '…'}</b><span>Alle</span><b>${esc(r.mensaAlle) || '…'}</b></div></div>
      </div>
    </div>
    <div class="riquadro">
      <div class="r-riga">
        <div class="r-cella largo"><span class="r-et fascia">LAVORAZIONI ESEGUITE:</span>${testo(r.lavorazioni)}</div>
        <div class="r-cella"><span class="r-et fascia">ALTRE NOTIZIE</span>${testo(r.altreNotizie)}</div>
      </div>
      <div class="r-riga"><div class="r-cella"><span class="r-et fascia sinistra">CONTROLLI EFFETTUATI:</span>${testo(r.controlli)}</div></div>
    </div>
    <div class="riquadro">
      <div class="r-riga">
        <div class="r-cella largo"><span class="r-et">Firma Caposquadra:</span><span class="firma">${esc(r.caposquadraNome)} · inviato il ${esc(dataOraIT(r.inviatoIl))}</span></div>
        <div class="r-cella"><span class="r-et">Firma RC:</span><span class="firma">${r.convalidatoDaNome ? `${esc(r.convalidatoDaNome)} · ${esc(dataOraIT(r.convalidatoIl))}` : '<span class="tenue">alla convalida</span>'}</span></div>
      </div>
    </div>`;
}

function opzioniCommessa(scelta) {
  const voci = [...stato.elenchi.commesse];
  if (scelta && !voci.includes(scelta)) voci.unshift(scelta);
  return '<option value=""></option>' + voci.map(c => `<option ${c === scelta ? 'selected' : ''}>${esc(c)}</option>`).join('');
}

function disegnaBlocchi() {
  const { dati, blocchi } = stato.aperto;
  const presenti = (dati.operai || []).map(o => o.nome);
  $('blocchi').innerHTML = blocchi.map((b, i) => `
    <div class="blocco" data-b="${i}">
      <div class="blocco-testa">
        <span>TURNO DI LAVORO ${blocchi.length > 1 ? i + 1 : ''}</span>
        ${blocchi.length > 1 ? `<button class="btn piccolo pericolo" data-azione="togli">Togli blocco</button>` : ''}
      </div>
      <div class="blocco-corpo">
        <div class="tabella-scorri"><table class="tab">
          <thead><tr><th></th><th>Commessa</th><th>Entrata</th><th>Uscita</th><th>Sub-ore</th><th>A detrarre pausa</th><th>Parziali</th></tr></thead>
          <tbody>${b.righe.map((r, t) => `<tr>
            <th>${t + 1}° turno</th>
            <td><select data-riga="${t}" data-k="commessa">${opzioniCommessa(r.commessa)}</select></td>
            <td><input type="time" data-riga="${t}" data-k="entrata" value="${esc(r.entrata)}"></td>
            <td><input type="time" data-riga="${t}" data-k="uscita" value="${esc(r.uscita)}"></td>
            <td class="calc" data-calc="sub${t}"></td>
            <td><input class="ore" inputmode="decimal" data-riga="${t}" data-k="pausa" value="${esc(r.pausa)}"></td>
            <td class="calc" data-calc="parz${t}"></td>
          </tr>`).join('')}
          <tr><th colspan="6" style="text-align:right">Totale</th><td class="calc" data-calc="totTurni"></td></tr>
          </tbody>
        </table></div>
        <div class="tabella-scorri"><table class="tab">
          <thead><tr><th></th><th>1° turno</th><th>2° turno</th></tr></thead>
          <tbody>${VOCI_ORE.map(([k, et]) => `<tr><th>${et}</th>
            <td><input class="ore" inputmode="decimal" data-t="t1" data-k="${k}" value="${esc(b.ore.t1[k])}"></td>
            <td><input class="ore" inputmode="decimal" data-t="t2" data-k="${k}" value="${esc(b.ore.t2[k])}"></td></tr>`).join('')}
            <tr><th>Totale ore</th><td class="calc" data-calc="ore1"></td><td class="calc" data-calc="ore2"></td></tr>
          </tbody>
        </table></div>
      </div>
      <div class="aiuto" style="margin-top:8px">Vale per:
        <button class="link" data-azione="tutti">tutti</button> ·
        <button class="link" data-azione="nessuno">nessuno</button></div>
      <div class="blocco-operai">${presenti.map(n => `<label class="chip ${b.operai.includes(n) ? 'si' : ''}">
        <input type="checkbox" data-operaio="${esc(n)}" ${b.operai.includes(n) ? 'checked' : ''}>${esc(n)}</label>`).join('')}</div>
    </div>`).join('');
  blocchi.forEach((_, i) => calcolaBlocco(i));
  aggiornaRiepilogo();
}

function calcolaBlocco(i) {
  const b = stato.aperto.blocchi[i];
  const box = document.querySelector(`.blocco[data-b="${i}"]`);
  if (!box) return;
  const set = (k, v) => { box.querySelector(`[data-calc="${k}"]`).textContent = v ? fmtOre(v) : ''; };
  let tot = 0;
  b.righe.forEach((r, t) => {
    const sub = differenzaOre(r.entrata, r.uscita);
    const parz = sub ? Math.max(0, sub - num(r.pausa)) : 0;
    set(`sub${t}`, sub); set(`parz${t}`, parz);
    tot += parz;
  });
  set('totTurni', tot);
  ['t1', 't2'].forEach((t, x) => set(`ore${x + 1}`, VOCI_ORE.reduce((s, [k]) => s + num(b.ore[t][k]), 0)));
}

$('blocchi').addEventListener('input', e => {
  const box = e.target.closest('.blocco');
  if (!box) return;
  const i = parseInt(box.dataset.b);
  const b = stato.aperto.blocchi[i];
  const el = e.target;
  if (el.dataset.riga !== undefined) b.righe[parseInt(el.dataset.riga)][el.dataset.k] = el.value;
  else if (el.dataset.t) b.ore[el.dataset.t][el.dataset.k] = el.value;
  else if (el.dataset.operaio !== undefined) {
    const n = el.dataset.operaio;
    b.operai = el.checked ? [...new Set([...b.operai, n])] : b.operai.filter(x => x !== n);
    el.parentElement.classList.toggle('si', el.checked);
  }
  calcolaBlocco(i);
  aggiornaRiepilogo();
});
$('blocchi').addEventListener('change', e => { if (e.target.tagName === 'SELECT') e.target.dispatchEvent(new Event('input', { bubbles: true })); });
$('blocchi').addEventListener('click', e => {
  const az = e.target.dataset.azione;
  if (!az) return;
  const i = parseInt(e.target.closest('.blocco').dataset.b);
  const bl = stato.aperto.blocchi;
  if (az === 'togli') bl.splice(i, 1);
  if (az === 'tutti') bl[i].operai = (stato.aperto.dati.operai || []).map(o => o.nome);
  if (az === 'nessuno') bl[i].operai = [];
  disegnaBlocchi();
});
$('btnAggiungiBlocco').onclick = () => {
  const r = stato.aperto.dati;
  stato.aperto.blocchi.push({
    righe: [{ commessa: r.commessa || '', entrata: '', uscita: '', pausa: '' }, { commessa: '', entrata: '', uscita: '', pausa: '' }],
    ore: { t1: oreVuote(), t2: oreVuote() },
    operai: []
  });
  disegnaBlocchi();
};

// Dai blocchi alle righe da scrivere in Excel (una per operaio e turno) + controlli
function calcolaRecord() {
  const { dati, blocchi } = stato.aperto;
  const records = [], errori = [], avvisi = [];
  const visti = new Map();
  blocchi.forEach((b, i) => {
    const nomeB = blocchi.length > 1 ? `Blocco ${i + 1}` : 'Il blocco';
    let qualcosa = false;
    ['t1', 't2'].forEach((t, x) => {
      const ore = Object.fromEntries(VOCI_ORE.map(([k]) => [k, num(b.ore[t][k])]));
      const somma = Object.values(ore).reduce((s, v) => s + v, 0);
      if (somma <= 0) return;
      qualcosa = true;
      if (Object.values(ore).some(v => v < 0 || v > 24)) errori.push(`${nomeB}, ${x + 1}° turno: ore non valide`);
      const commessa = b.righe[x].commessa || (x === 1 ? b.righe[0].commessa : '') || dati.commessa;
      if (stato.elenchi.commesse.length && !stato.elenchi.commesse.includes(commessa)) {
        avvisi.push(`${nomeB}, ${x + 1}° turno: la commessa "${commessa}" non è nell'elenco del file Excel`);
      }
      b.operai.forEach(nome => {
        const chiave = `${nome}|${x + 1}`;
        if (visti.has(chiave)) errori.push(`${nome}: ${x + 1}° turno presente in due blocchi`);
        visti.set(chiave, true);
        records.push({ workerName: nome, turno: x + 1, cantiere: commessa, commessa, ...ore, malattia: 0 });
      });
    });
    if (qualcosa && b.operai.length === 0) errori.push(`${nomeB}: spunta gli operai a cui vale`);
    if (!qualcosa && b.operai.length > 0) errori.push(`${nomeB}: scrivi le ore nel riquadro 1°/2° turno`);
  });
  const conOre = new Set(records.map(r => r.workerName));
  const senza = (dati.operai || []).map(o => o.nome).filter(n => !conOre.has(n));
  if (senza.length) avvisi.push(`Senza ore: ${senza.join(', ')}`);
  if (records.length === 0 && errori.length === 0) errori.push('Nessuna ora da scrivere');
  return { records, errori, avvisi };
}

function aggiornaRiepilogo() {
  const { records, errori, avvisi } = calcolaRecord();
  const totale = records.reduce((s, r) => s + VOCI_ORE.reduce((x, [k]) => x + r[k], 0), 0);
  $('riepilogoOre').innerHTML =
    errori.map(e => `<div class="errore">${esc(e)}</div>`).join('') +
    avvisi.map(a => `<div class="errore" style="background:var(--attenzione-sfondo);color:var(--testo)">${esc(a)}</div>`).join('') +
    (records.length && !errori.length ? `<div class="ok-box">${new Set(records.map(r => r.workerName)).size} operai · ${records.length} turni · ${fmtOre(totale)} ore in tutto</div>` : '');
}

$('btnSalvaOre').onclick = async () => {
  try {
    await updateDoc(doc(db, 'rapportini', stato.aperto.id), { blocchi: stato.aperto.blocchi });
    avviso('Ore salvate (il rapportino resta tra gli arrivati)', 'ok');
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
};

$('btnElimina').onclick = async () => {
  const r = stato.aperto.dati;
  const ok = await conferma('Eliminare il rapportino?',
    `<p>${esc(dataIT(r.data))} · ${esc(r.commessa)} · caposquadra ${esc(r.caposquadraNome)}</p>
     <p>Il rapportino verrà cancellato definitivamente. Il file Excel non viene toccato.</p>`, 'Elimina');
  if (!ok) return;
  try {
    await deleteDoc(doc(db, 'rapportini', r.id));
    avviso('Rapportino eliminato', 'ok');
    mostraVista(r.stato === 'convalidato' ? 'convalidati' : 'arrivati');
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
};

// ---------------------------------------------------------------- scrittura su Excel
function mostraAvvisoCartella(s) {
  stato.cartellaPronta = false;
  $('avvisoCartella').hidden = false;
  if (!CartellaExcel.supported()) {
    $('avvisoCartellaTesto').textContent = 'Per scrivere su Excel apri l\'app sul PC dell\'ufficio con Microsoft Edge o Google Chrome.';
    $('btnCartella').hidden = true;
  } else if (s === 'da-confermare') {
    $('avvisoCartellaTesto').textContent = `Conferma l'accesso alla cartella "${CartellaExcel.name()}" dei file Excel.`;
    $('btnCartella').textContent = 'Consenti accesso';
    $('btnCartella').dataset.azione = 'conferma';
  } else {
    $('avvisoCartellaTesto').textContent = 'Collega la cartella dei file Excel (serve una volta sola).';
    $('btnCartella').textContent = 'Collega cartella Excel';
    $('btnCartella').dataset.azione = 'scegli';
  }
  $('infoCartella').textContent = 'Cartella non collegata.';
}
function cartellaPronta() {
  stato.cartellaPronta = true;
  $('avvisoCartella').hidden = true;
  $('infoCartella').textContent = `Cartella collegata: ${CartellaExcel.name()}`;
}
$('btnCartella').onclick = async () => {
  try {
    if ($('btnCartella').dataset.azione === 'conferma' && await CartellaExcel.confirmAccess()) return cartellaPronta();
    await CartellaExcel.choose();
    cartellaPronta();
  } catch (err) { if (err.name !== 'AbortError') avviso(`Cartella non collegata: ${err.message}`, 'errore'); }
};
$('btnCambiaCartella').onclick = async () => {
  try { await CartellaExcel.choose(); cartellaPronta(); avviso(`Cartella collegata: ${CartellaExcel.name()}`, 'ok'); }
  catch (err) { if (err.name !== 'AbortError') avviso(err.message, 'errore'); }
};
async function richiediCartella() {
  if (stato.cartellaPronta) return true;
  if (!CartellaExcel.supported()) {
    await finestra('Excel non raggiungibile', '<p>Per scrivere su Excel usa il PC dell\'ufficio con Microsoft Edge o Google Chrome.</p>');
    return false;
  }
  avviso('Collega prima la cartella dei file Excel (barra gialla in alto)', 'errore');
  return false;
}

// File del mese del rapportino (es. Presenze_11-Novembre_2026.xlsm). Se non c'è, propone di crearlo dal Master.
async function fileDelMese(anno, mese) {
  const files = await CartellaExcel.listExcelFiles();
  const master = CartellaExcel.masterName(files);
  const trovato = files.find(f => f !== master && /presenze/i.test(f) && f.includes(String(anno)) &&
    (() => { const p = ExcelXlsm.parseMonthAndYear(f); return p.month === mese && p.year === anno; })());
  if (trovato) return trovato;
  const nuovo = `Presenze_${String(mese).padStart(2, '0')}-${MESI[mese - 1]}_${anno}.xlsm`;
  if (!master) throw new Error('nella cartella non trovo né il file del mese né il Master_Operai da copiare');
  const ok = await conferma('Manca il file del mese',
    `<p>Nella cartella non c'è il file presenze di <b>${MESI[mese - 1]} ${anno}</b>.</p>
     <p>Lo creo adesso copiando <b>${esc(master)}</b> con il nome <b>${esc(nuovo)}</b>?</p>`, 'Crea il file');
  if (!ok) return null;
  const bytes = await ExcelXlsm.prepareNewMonth(await CartellaExcel.readBytes(master), mese, anno);
  await CartellaExcel.writeBytes(nuovo, bytes, true);
  avviso(`Creato ${nuovo}`, 'ok');
  return nuovo;
}

const etichettaTurno = v => [v.commessa || '(senza commessa)', ...VOCI_ORE.filter(([k]) => v[k]).map(([k, et]) => `${et} ${fmtOre(v[k])}`)].join(' · ');

$('btnConvalida').onclick = async () => {
  const { dati } = stato.aperto;
  const { records, errori, avvisi } = calcolaRecord();
  if (errori.length) { await finestra('Da sistemare prima di scrivere', `<ul>${errori.map(e => `<li>${esc(e)}</li>`).join('')}</ul>`); return; }
  if (avvisi.length && !await conferma('Controlla', `<ul>${avvisi.map(a => `<li>${esc(a)}</li>`).join('')}</ul><p>Vuoi scrivere lo stesso?</p>`, 'Scrivi lo stesso')) return;
  if (!await richiediCartella()) return;

  const btn = $('btnConvalida');
  btn.disabled = true;
  const testoBtn = btn.textContent;
  btn.textContent = 'Scrittura su Excel in corso…';
  try {
    const [anno, mese, giorno] = dati.data.split('-').map(Number);
    const fileName = await fileDelMese(anno, mese);
    if (!fileName) return;
    if (await CartellaExcel.isOpenInExcel(fileName)) throw new Error(`il file ${fileName} è aperto in Excel: chiudilo e riprova`);

    let bytes = await CartellaExcel.readBytes(fileName);
    // posizione di ogni operaio nel file del mese (si cerca per nome: l'ordine nel file è quello che conta)
    const operaiFile = await ExcelXlsm.readWorkers(bytes);
    const indice = new Map(operaiFile.map(w => [w.name.trim().toUpperCase(), w.index]));
    const mancanti = [...new Set(records.map(r => r.workerName).filter(n => !indice.has(n.trim().toUpperCase())))];
    if (mancanti.length) throw new Error(`questi operai non sono nel file ${fileName}: ${mancanti.join(', ')}`);
    let daScrivere = records.map(r => ({ ...r, workerIndex: indice.get(r.workerName.trim().toUpperCase()) }));

    const conflitti = await ExcelXlsm.findConflicts(bytes, giorno, daScrivere);
    let saltati = 0;
    if (conflitti.length) {
      const scelta = await finestra('Dati già presenti in Excel', `
        <p>In <b>${esc(fileName)}</b>, giorno ${giorno}, questi turni contengono già dati diversi (forse da un altro rapportino):</p>
        <ul>${conflitti.map(c => `<li><b>${esc(c.workerName)} – ${c.turno}° turno</b><br>In Excel: ${esc(etichettaTurno(c.existing))}<br>Nuovo: ${esc(etichettaTurno(c.nuovo))}</li>`).join('')}</ul>`,
        [{ testo: 'Annulla', valore: 'annulla' }, { testo: 'Lascia quelli già scritti', valore: 'salta' }, { testo: 'Sovrascrivi', valore: 'sovrascrivi', primario: true }]);
      if (scelta === 'annulla' || !scelta) { avviso('Scrittura annullata: Excel non è stato modificato'); return; }
      if (scelta === 'salta') {
        const inConflitto = new Set(conflitti.map(c => `${c.workerIndex}-${c.turno}`));
        const prima = daScrivere.length;
        daScrivere = daScrivere.filter(r => !inConflitto.has(`${r.workerIndex}-${r.turno}`));
        saltati = prima - daScrivere.length;
      }
    }

    if (daScrivere.length) {
      // copia di sicurezza alla prima scrittura del giorno per ogni file
      const chiaveCopia = `rapp_copia:${CartellaExcel.name()}/${fileName}`;
      if (ls.get(chiaveCopia) !== oggiISO()) {
        try { await CartellaExcel.datedBackup(fileName, bytes); ls.set(chiaveCopia, oggiISO()); }
        catch (e) { throw new Error(`copia di sicurezza non riuscita (${e.message}): Excel non è stato modificato`); }
      }
      const risultato = await ExcelXlsm.writePresenze(bytes, giorno, daScrivere);
      await CartellaExcel.writeBytes(fileName, risultato.bytes);
    }

    await updateDoc(doc(db, 'rapportini', dati.id), {
      stato: 'convalidato',
      blocchi: stato.aperto.blocchi,
      convalidatoIl: serverTimestamp(),
      convalidatoDaUid: stato.utente.uid,
      convalidatoDaNome: stato.utente.nome,
      fileExcel: fileName
    });
    await finestra('Convalidato', `<p>Scritti <b>${daScrivere.length}</b> turni in <b>${esc(fileName)}</b>, giorno ${giorno}.${saltati ? ` ${saltati} lasciati com'erano.` : ''}</p>
      <p>Il rapportino è passato tra i convalidati.</p>`);
    mostraVista('arrivati');
  } catch (err) {
    const msg = /NoModificationAllowed|InvalidState|being used|in uso/i.test(`${err.name} ${err.message}`)
      ? 'il file è in uso (forse aperto in Excel): chiudilo e riprova' : messaggioErrore(err);
    await finestra('Scrittura non riuscita', `<p>${esc(msg)}</p>`);
  } finally {
    btn.disabled = false;
    btn.textContent = testoBtn;
  }
};

// ---------------------------------------------------------------- elenchi dal file Excel
function mostraInfoElenchi() {
  const e = stato.elenchi;
  $('infoElenchi').textContent = e.operai.length
    ? `Ora sui telefoni: ${e.operai.length} operai e ${e.commesse.length} commesse${e.file ? ` (da ${e.file})` : ''}${e.aggiornatoIl ? `, aggiornati il ${dataOraIT(e.aggiornatoIl)}` : ''}.`
    : 'Elenchi non ancora caricati.';
  $('listaVisibili').innerHTML = e.commesse.map(c => {
    const si = !e.nascoste.includes(c);
    return `<label class="chip ${si ? 'si' : ''}"><input type="checkbox" data-voce="${esc(c)}" ${si ? 'checked' : ''}>${esc(c)}</label>`;
  }).join('');
}

// Voci che non sono commesse di lavoro: nascoste in automatico sui telefoni
const VOCE_NON_COMMESSA = /^(ASSENTE|FERIE|MALATTIA|SMONTO|CORSO|VISITA|PERMESSO|ROL\b|INFORTUNIO)/i;

$('listaVisibili').addEventListener('change', async e => {
  const voce = e.target.dataset.voce;
  if (voce === undefined) return;
  const nascoste = e.target.checked ? stato.elenchi.nascoste.filter(c => c !== voce) : [...new Set([...stato.elenchi.nascoste, voce])];
  try { await updateDoc(doc(db, 'elenchi', 'correnti'), { nascoste }); }
  catch (err) { avviso(messaggioErrore(err), 'errore'); mostraInfoElenchi(); }
});

$('btnAggiornaElenchi').onclick = async () => {
  if (!await richiediCartella()) return;
  try {
    const files = await CartellaExcel.listExcelFiles();
    const master = CartellaExcel.masterName(files);
    if (!master) throw new Error('nella cartella non trovo il file Master_Operai');
    const bytes = await CartellaExcel.readBytes(master);
    const operai = (await ExcelXlsm.readWorkers(bytes)).map(w => w.name);
    const commesse = await ExcelXlsm.readCommesse(bytes);
    const ok = await conferma('Aggiornare gli elenchi?', `
      <p>Dal file <b>${esc(master)}</b>: <b>${operai.length}</b> operai e <b>${commesse.length}</b> commesse.</p>
      <p>I telefoni dei capisquadra li riceveranno alla prossima connessione.</p>`, 'Aggiorna');
    if (!ok) return;
    // le scelte già fatte dal responsabile restano; le voci nuove seguono la regola automatica
    const giaViste = new Set(stato.elenchi.commesse);
    const nascoste = commesse.filter(c => giaViste.has(c) ? stato.elenchi.nascoste.includes(c) : VOCE_NON_COMMESSA.test(c));
    await setDoc(doc(db, 'elenchi', 'correnti'), { operai, commesse, nascoste, file: master, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
    avviso('Elenchi aggiornati', 'ok');
  } catch (err) { await finestra('Elenchi non aggiornati', `<p>${esc(messaggioErrore(err))}</p>`); }
};

// ---------------------------------------------------------------- intestazione del modulo
const ETICHETTE_INTESTAZIONE = {
  sedeLegale: 'Sede legale', sedeOperativa: 'Sede operativa', telefono: 'Telefono e fax', sito: 'Sito internet',
  titolo: 'Titolo', sottotitolo: 'Sottotitolo', identificazione: 'Identificazione (es. Mod. 22.03)', revisione: 'Revisione',
  dataRevisione: 'Data della revisione', norma1: 'Norma 1', norma2: 'Norma 2', norma3: 'Norma 3'
};
let logoInModifica = null;      // logo scelto ma non ancora salvato ('' = logo originale)
let intestazioneToccata = false; // il responsabile sta scrivendo: non sovrascrivere i suoi campi

function valoriFormIntestazione() {
  const v = {};
  document.querySelectorAll('#campiIntestazione [data-campo]').forEach(inp => { v[inp.dataset.campo] = inp.value.trim(); });
  v.logo = logoInModifica !== null ? logoInModifica : intestazione.logo;
  return v;
}
function anteprimaIntestazione() {
  $('anteprimaIntestazione').innerHTML = testataModulo(1, valoriFormIntestazione());
}
function disegnaGestioneIntestazione() {
  if (stato.utente?.ruolo !== 'responsabile' || intestazioneToccata) return;
  $('campiIntestazione').innerHTML = Object.entries(ETICHETTE_INTESTAZIONE).map(([k, et]) => `
    <label class="campo"><span>${et}</span><input data-campo="${k}" maxlength="120" value="${esc(intestazione[k])}"></label>`).join('');
  logoInModifica = null;
  anteprimaIntestazione();
}
$('campiIntestazione').addEventListener('input', () => { intestazioneToccata = true; anteprimaIntestazione(); });

// Logo caricato dal PC: rimpicciolito (max 800 pixel di larghezza) perché stia nel database
$('fileLogo').addEventListener('change', async e => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const img = await createImageBitmap(file);
    const scala = Math.min(1, 800 / img.width);
    const c = document.createElement('canvas');
    c.width = Math.round(img.width * scala); c.height = Math.round(img.height * scala);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    let dato = c.toDataURL('image/png');
    if (dato.length > 700000) { // troppo pesante: JPEG su fondo bianco
      const ctx = c.getContext('2d');
      ctx.globalCompositeOperation = 'destination-over'; ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      dato = c.toDataURL('image/jpeg', 0.85);
    }
    if (dato.length > 900000) throw new Error('immagine troppo grande anche dopo averla rimpicciolita');
    logoInModifica = dato;
    intestazioneToccata = true;
    anteprimaIntestazione();
    avviso('Logo pronto: premi "Salva intestazione" per confermarlo');
  } catch (err) { avviso(`Logo non caricato: ${err.message}`, 'errore'); }
});
$('btnLogoOriginale').onclick = () => { logoInModifica = ''; intestazioneToccata = true; anteprimaIntestazione(); };
$('btnIntestazioneBase').onclick = () => {
  document.querySelectorAll('#campiIntestazione [data-campo]').forEach(inp => { inp.value = INTESTAZIONE_BASE[inp.dataset.campo]; });
  intestazioneToccata = true;
  anteprimaIntestazione();
};
$('formIntestazione').addEventListener('submit', async e => {
  e.preventDefault();
  const v = valoriFormIntestazione();
  if (!await conferma('Salvare l\'intestazione?', `<p>La nuova intestazione comparirà sui rapportini compilati da adesso in poi.</p>
    <div class="modulo">${testataModulo(1, v)}</div>`, 'Salva')) return;
  try {
    await setDoc(doc(db, 'elenchi', 'intestazione'), { ...v, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
    intestazioneToccata = false;
    avviso('Intestazione salvata', 'ok');
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// ---------------------------------------------------------------- automezzi
function disegnaGestioneMezzi() {
  if (stato.utente?.ruolo !== 'responsabile') return;
  $('listaMezzi').innerHTML = stato.automezzi.length
    ? stato.automezzi.map((m, i) => `<div class="voce">
        <div class="voce-princ"><div class="voce-titolo">${esc(m)}</div></div>
        <button class="btn piccolo pericolo" data-togli-mezzo="${i}">Togli</button></div>`).join('')
    : '<div class="nessuno">Nessun automezzo: aggiungili qui sotto</div>';
}
const salvaMezzi = voci => setDoc(doc(db, 'elenchi', 'automezzi'), { voci, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });

$('formMezzo').addEventListener('submit', async e => {
  e.preventDefault();
  const nuovo = $('nuovoMezzo').value.trim().replace(/\s+/g, ' ');
  if (!nuovo) return;
  if (stato.automezzi.some(m => m.toUpperCase() === nuovo.toUpperCase())) { avviso('Questo automezzo è già in elenco', 'errore'); return; }
  try {
    await salvaMezzi([...stato.automezzi, nuovo]);
    $('nuovoMezzo').value = '';
    avviso(`Aggiunto: ${nuovo}`, 'ok');
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
});
$('listaMezzi').addEventListener('click', async e => {
  const i = e.target.dataset.togliMezzo;
  if (i === undefined) return;
  const m = stato.automezzi[Number(i)];
  if (!await conferma('Togliere l\'automezzo?', `<p><b>${esc(m)}</b> non comparirà più sui telefoni. I rapportini già inviati non cambiano.</p>`, 'Togli')) return;
  try { await salvaMezzi(stato.automezzi.filter(x => x !== m)); }
  catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// ---------------------------------------------------------------- utenti
function disegnaUtenti(utenti) {
  utenti.sort((a, b) => (a.ruolo || '').localeCompare(b.ruolo || '') || (a.nome || '').localeCompare(b.nome || ''));
  $('listaUtenti').innerHTML = utenti.map(u => `
    <div class="voce">
      <div class="voce-princ">
        <div class="voce-titolo">${esc(u.nome)}</div>
        <div class="voce-sotto">Codice: ${esc(u.codice)} · ${u.ruolo === 'responsabile' ? 'Responsabile' : 'Caposquadra'}</div>
      </div>
      <span class="etichetta ${u.attivo ? 'ok' : 'no'}">${u.attivo ? 'Attivo' : 'Disattivato'}</span>
      ${u.uid === stato.utente.uid ? '' : `<button class="btn piccolo" data-uid="${esc(u.uid)}" data-attivo="${u.attivo ? 1 : 0}">${u.attivo ? 'Disattiva' : 'Riattiva'}</button>`}
    </div>`).join('');
}
$('listaUtenti').addEventListener('click', async e => {
  const b = e.target.closest('[data-uid]');
  if (!b) return;
  const attiva = b.dataset.attivo !== '1';
  try { await updateDoc(doc(db, 'utenti', b.dataset.uid), { attivo: attiva }); }
  catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// Il nuovo account si crea con una seconda connessione "usa e getta", così il responsabile resta collegato
$('formUtente').addEventListener('submit', async e => {
  e.preventDefault();
  const nome = $('nuNome').value.trim();
  const codice = normCodice($('nuCodice').value);
  const password = $('nuPassword').value;
  const ruolo = $('nuRuolo').value;
  if (!codiceValido(codice)) { avviso('Codice utente: solo lettere minuscole, numeri, punto o trattino, senza spazi', 'errore'); return; }
  $('btnCreaUtente').disabled = true;
  let app2;
  try {
    app2 = initializeApp(config, `nuovo-utente-${Date.now()}`);
    const auth2 = initializeAuth(app2, { persistence: inMemoryPersistence });
    if (EMULATORE) connectAuthEmulator(auth2, 'http://127.0.0.1:9099', { disableWarnings: true });
    const cred = await createUserWithEmailAndPassword(auth2, emailDi(codice), password);
    await setDoc(doc(db, 'utenti', cred.user.uid), { nome, codice, ruolo, attivo: true, creatoIl: serverTimestamp(), creatoDa: stato.utente.nome });
    await signOut(auth2);
    $('formUtente').reset();
    await finestra('Utente creato', `<p><b>${esc(nome)}</b> può entrare con:</p><p>Codice utente: <b>${esc(codice)}</b><br>Password: <b>${esc(password)}</b></p><p>Comunicaglieli di persona.</p>`);
  } catch (err) {
    avviso(messaggioErrore(err), 'errore');
  } finally {
    if (app2) deleteApp(app2).catch(() => {});
    $('btnCreaUtente').disabled = false;
  }
});

// Solo per le prove automatiche con il simulatore: cartella finta e lettura diretta (per verificare le regole)
if (EMULATORE) {
  window.__prova = {
    usaCartella: async handle => { await CartellaExcel.useHandle(handle); cartellaPronta(); },
    leggiRapportini: async () => (await getDocs(collection(db, 'rapportini'))).size,
    modificaRapportino: async (id, campi) => updateDoc(doc(db, 'rapportini', id), campi)
  };
}
