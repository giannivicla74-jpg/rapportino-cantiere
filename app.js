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
  createUserWithEmailAndPassword, signOut, connectAuthEmulator, updatePassword, reauthenticateWithCredential,
  EmailAuthProvider
} from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js';
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager, doc, collection, getDoc,
  getDocFromCache, getDocFromServer, getDocs, setDoc, updateDoc, deleteDoc, onSnapshot, query, where, serverTimestamp, writeBatch,
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
// Il codice di accesso nasce dal nome: "Mario Rossi" → "mario.rossi" (senza accenti, maiuscole e spazi).
// I codici vecchi (es. "001") restano uguali.
const normCodice = c => String(c || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim()
  .replace(/[^a-z0-9._-]+/g, '.').replace(/^\.+|\.+$/g, '');
const emailDi = codice => `${normCodice(codice)}@${DOMINIO_ACCESSI}`;

function messaggioErrore(err) {
  const c = err?.code || '';
  if (/invalid-credential|wrong-password|user-not-found|invalid-email/.test(c)) return 'Nome o PIN errati';
  if (/network-request-failed|unavailable/.test(c)) return 'Nessuna connessione a internet: riprova quando c\'è campo';
  if (/too-many-requests/.test(c)) return 'Troppi tentativi: aspetta qualche minuto e riprova';
  if (/email-already-in-use/.test(c)) return 'Esiste già un utente con questo nome (anche eliminato): aggiungi per esempio la seconda iniziale o un numero';
  if (/weak-password/.test(c)) return 'Password troppo corta: almeno 6 caratteri';
  if (/permission-denied/.test(c)) return 'Operazione non permessa per questo utente';
  if (/operation-not-allowed/.test(c)) return 'Accesso con password non attivato in Firebase (vedi la guida)';
  return err?.message || String(err);
}

// ---------------------------------------------------------------- avvio
// Intestazione come sul modulo cartaceo (pagina 1 completa, pagina 2 solo la riga di identificazione).
// I testi e il logo si cambiano in "Elenchi e utenti > Intestazione del modulo"; questi sono i valori di partenza.
// Grafica 2025 (Edizione 02): ragione sociale, due sedi, dati fiscali e contatti; 5 caselle nella riga di identificazione.
const INTESTAZIONE_BASE = {
  ragioneSociale: 'VI.CLA. FUTURE S.r.l.',
  sedeLegale: 'Viale Michelangelo n. 33 – 80129 Napoli (Na) - Italia',
  sedeAmministrativa: 'Strada Consortile snc - 81032 Carinaro (CE) - Italia',
  datiFiscali: 'P.IVA/C.F./n° di Iscrizione C.C.I.A.A. di Napoli: 05485601214 - R.E.A. NA-756780',
  telefono: 'Tel +39 081 3041343 - Fax +39 081 5029730',
  contatti: 'e-mail: info@vicla.it - web: www.vicla.it - pec: info@pec.vicla.it',
  titolo: 'MODELLO del SISTEMA di GESTIONE INTEGRATO QUALITÀ, AMBIENTE, SALUTE e SICUREZZA',
  sottotitolo: 'RAPPORTO DI LAVORO',
  identificazione: 'Mod. 22.03',
  edizione: '02',
  revisione: '00',
  dataRevisione: '07/04/2025',
  norma1: 'UNI EN ISO 9001:2015',
  norma2: 'UNI EN ISO 14001:2015',
  norma3: 'UNI ISO 45001:2018',
  // titoli delle colonne degli orari: seguono le stesse regole dell'intestazione (un rapportino inviato conserva i suoi)
  colPartenza: 'ORARIO PARTENZA DA DEPOSITO',
  colCantiere: 'ORARIO DI LAVORO IN CANTIERE',
  colRientro: 'ORARIO RIENTRO IN DEPOSITO'
};
const titoloColonna = (t, k) => t[k] || INTESTAZIONE_BASE[k];
const LOGO_BASE = 'icons/logo-vicla.png';
let intestazione = { ...INTESTAZIONE_BASE, logo: '' };
// Intestazione di un rapportino: quella salvata al momento dell'invio (anche nella grafica vecchia), altrimenti l'attuale
const intestazioneDi = r => (r.intestazione ? { logo: intestazione.logo, ...r.intestazione } : intestazione);

function testataModulo(pagina, t = intestazione) {
  return t.ragioneSociale ? testataNuova(pagina, t) : testataVecchia(pagina, t);
}
function logoDi(t) {
  // il logo è un'immagine caricata dall'ufficio (data:image/...) oppure quello originale dell'app
  return /^data:image\/(png|jpeg|webp);base64,/.test(t.logo || '') ? t.logo : LOGO_BASE;
}
function testataNuova(pagina, t) {
  const righe = `<div class="t2-id">
      <div>Identificazione: ${esc(t.identificazione)}</div><div>Edizione: ${esc(t.edizione)}</div><div>Revisione.: ${esc(t.revisione)}</div>
      <div>Data: ${esc(t.dataRevisione)}</div><div>Pagina ${pagina} di 2</div>
    </div>
    <div class="t2-norme"><div>Rif. Norme</div><div>${esc(t.norma1)}</div><div>${esc(t.norma2)}</div><div>${esc(t.norma3)}</div></div>`;
  if (pagina === 2) return `<div class="testata t2">${righe}</div>`;
  return `<div class="testata t2">
    <div class="t2-alto">
      <div class="t-logo"><img src="${logoDi(t)}" alt="Logo"></div>
      <div class="t2-dati">
        <div class="t2-rs">${esc(t.ragioneSociale)}</div>
        <div class="t2-sedi">
          <div>Sede legale:<br>${esc(t.sedeLegale)}</div>
          <div>Sede Amministrativa:<br>${esc(t.sedeAmministrativa)}</div>
        </div>
        <div class="t2-fisc">${[t.datiFiscali, t.telefono, t.contatti].filter(Boolean).map(r => `<div>${esc(r)}</div>`).join('')}</div>
      </div>
    </div>
    <div class="t-titolo">${esc(t.titolo)}</div>
    <div class="t-sotto">${esc(t.sottotitolo)}</div>
    ${righe}
  </div>`;
}
// Grafica precedente (Revisione 03 del 2016): resta per i rapportini inviati prima dell'aggiornamento
function testataVecchia(pagina, t) {
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
  // titoli delle colonne nel modulo in compilazione
  $('titoloPartenza').textContent = titoloColonna(intestazione, 'colPartenza');
  $('titoloCantiere').textContent = titoloColonna(intestazione, 'colCantiere');
  $('titoloRientro').textContent = titoloColonna(intestazione, 'colRientro');
}
aggiornaTestate();
$('versioneApp').textContent = `Versione ${self.APP_VERSION} del ${self.APP_VERSION_DATE}`;
$('versioneAccesso').textContent = `Versione ${self.APP_VERSION} del ${self.APP_VERSION_DATE}`;
// Aggiornamenti: il service worker controlla se c'è una versione nuova ogni volta che l'app torna in primo piano
// (l'app installata sul telefono spesso non si chiude davvero) e ogni 30 minuti. Quando la versione nuova è pronta
// compare una barra con "Aggiorna": niente ricarica automatica, per non perdere ore che si stanno scrivendo.
if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  const avevaControllo = !!navigator.serviceWorker.controller; // false alla primissima installazione
  navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then(reg => {
    const controlla = () => reg.update().catch(() => {});
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') controlla(); });
    setInterval(controlla, 30 * 60 * 1000);
  }).catch(() => {});
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (avevaControllo) mostraAggiornamento(); });
}
// Tema: Automatico (segue il telefono/PC) → Chiaro → Scuro → Automatico… La scelta resta su questo dispositivo.
const TEMI = { auto: { icona: '◐', nome: 'Automatico' }, chiaro: { icona: '☀', nome: 'Chiaro' }, scuro: { icona: '🌙', nome: 'Scuro' } };
function applicaTema(t) {
  if (t === 'chiaro' || t === 'scuro') document.documentElement.dataset.tema = t;
  else { delete document.documentElement.dataset.tema; t = 'auto'; }
  document.querySelectorAll('[data-tema-btn]').forEach(b => {
    b.textContent = TEMI[t].icona;
    b.title = `Tema: ${TEMI[t].nome} (tocca per cambiare)`;
    b.setAttribute('aria-label', b.title);
  });
  const scuro = t === 'scuro' || (t === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
  document.querySelector('meta[name="theme-color"]')?.setAttribute('content', scuro ? '#10151c' : '#1f3a5f');
  return t;
}
const temaSalvato = (() => { try { return localStorage.getItem('rapp_tema'); } catch (_) { return null; } })();
let temaAttuale = applicaTema(temaSalvato || 'auto');
document.querySelectorAll('[data-tema-btn]').forEach(b => {
  b.onclick = () => {
    temaAttuale = applicaTema({ auto: 'chiaro', chiaro: 'scuro', scuro: 'auto' }[temaAttuale]);
    try { localStorage.setItem('rapp_tema', temaAttuale); } catch (_) {} // letto anche dallo script in <head>
    avviso(`Tema: ${TEMI[temaAttuale].nome}`);
  };
});

// Telefono: l'app installata resta in orizzontale (dove il sistema lo permette); nel browser l'avviso si può chiudere

function mostraAggiornamento() {
  if (document.getElementById('barraAggiorna')) return;
  const barra = document.createElement('div');
  barra.id = 'barraAggiorna';
  barra.className = 'barra-aggiorna';
  barra.innerHTML = '<span>È disponibile la versione nuova dell\'app.</span><button class="btn primario piccolo" type="button">Aggiorna</button>';
  barra.querySelector('button').onclick = () => location.reload();
  document.body.appendChild(barra);
}

// Prove sul PC: con l'indirizzo http://localhost:...?emulatore l'app usa il simulatore di Firebase
// (dati finti, nessun account). Online questa modalità non si attiva mai.
const EMULATORE = new URLSearchParams(location.search).has('emulatore') && ['localhost', '127.0.0.1'].includes(location.hostname);
const config = EMULATORE
  ? { apiKey: 'demo-key', authDomain: 'demo-rapportino.firebaseapp.com', projectId: 'demo-rapportino', appId: 'demo' }
  : firebaseConfig;
const configurato = !Object.values(config).some(v => String(v).includes('DA-COMPILARE'));
let app, auth, db;
// All'apertura Firebase verifica online l'accesso già fatto e, con il campo debole, aspetterebbe 30 secondi.
// Solo durante l'avvio quella verifica ha 2 secondi: poi si entra con l'accesso salvato nel telefono.
// Le regole del database continuano a controllare ogni lettura e scrittura.
let avvioAccesso = true;
const fetchVero = self.fetch.bind(self);
self.fetch = (url, opzioni) => avvioAccesso && self.AbortSignal?.timeout && /accounts:lookup|securetoken/.test(String(url))
  ? fetchVero(url, { ...opzioni, signal: AbortSignal.timeout(2000) })
  : fetchVero(url, opzioni);
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
  responsabili: { voci: [], predefinito: '' }, // responsabili di commessa (dal database)
  orarioStandard: {},    // orari del pulsante "Orario standard" (dal database)
  orarioDeposito: {},    // orario speciale del pulsante con la commessa COMM.0 (dal database)
  presenti: new Map(),   // nome -> { partenza, dalle, alle, rientro }
  automezzi: [],         // elenco dei mezzi (dal database)
  mezziScelti: new Set(), // mezzi spuntati nel rapportino in compilazione
  autisti: new Map(),    // mezzo -> autista nel rapportino in compilazione
  arrivati: [],
  aperto: null,          // rapportino aperto dal responsabile: { id, dati, blocchi }
  cartellaPronta: false,
  annulla: []            // funzioni per chiudere gli ascolti di Firestore all'uscita
};

// ---------------------------------------------------------------- accesso
async function utenteCambiato(user) {
  avvioAccesso = false;
  stato.annulla.forEach(f => { try { f(); } catch (_) {} });
  stato.annulla = [];
  stato.areaAperta = false;
  $('avvisoCartella').hidden = true; // la barra della cartella riguarda solo il responsabile sul PC
  $('avvisoArchivio').hidden = true;
  if (!user) {
    stato.utente = null;
    mostraSchermo('scrLogin');
    return;
  }
  let profilo;
  try {
    const ref = doc(db, 'utenti', user.uid);
    // Con il campo debole non si resta ad aspettare: dopo 1,5 secondi si usa il profilo già nel telefono.
    // Al primo accesso la copia non c'è ancora, e allora si attende la risposta del server.
    const dalServer = getDoc(ref);
    let snap;
    try {
      snap = await Promise.race([dalServer, new Promise((_, no) => setTimeout(no, 1500))]);
    } catch (_) {
      snap = await getDocFromCache(ref).catch(() => dalServer);
    }
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
  // sviluppatore: l'unico che vede "Elenchi e utenti" (lo impongono anche le regole del database)
  const sviluppatore = profilo.ruolo === 'responsabile' && profilo.sviluppatore === true;
  stato.utente = { uid: user.uid, nome: profilo.nome, ruolo: profilo.ruolo, codice: profilo.codice, sviluppatore };
  $('barraUtente').textContent = `${profilo.nome} · ${profilo.ruolo === 'responsabile' ? 'Responsabile' : 'Caposquadra'}`;
  mostraSchermo('scrApp');
  avviaCompilazione();
  if (profilo.ruolo === 'responsabile') {
    $('menuResponsabile').hidden = false;
    // "Elenchi e utenti" nascosto per tutti: lo sviluppatore lo apre con 5 tocchi sul simbolo GC + PIN
    stato.areaAperta = false;
    document.querySelector('.scheda-btn[data-vista="gestione"]').hidden = true;
    avviaResponsabile();
    mostraVista('arrivati');
    // 5 tocchi sul simbolo nella schermata di accesso: dopo l'ingresso chiede subito il PIN
    if (ingressoSviluppatore && sviluppatore) apriAreaSviluppatore();
  } else {
    $('menuResponsabile').hidden = true;
    mostraVista('compila');
  }
  ingressoSviluppatore = false;
  $('sottotitoloAccesso').textContent = 'Mod. 22.03 digitale';
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

// Il collegamento "Prima configurazione" è nascosto: compare solo aprendo l'indirizzo con ?configura
// (oltre alla regola del database che rifiuta una seconda configurazione).
$('btnApriSetup').hidden = !(EMULATORE || new URLSearchParams(location.search).has('configura'));
$('btnApriSetup').onclick =() => { $('formLogin').hidden = true; $('formSetup').hidden = false; };
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
      batch.set(doc(db, 'utenti', cred.user.uid), { nome, codice, ruolo: 'responsabile', sviluppatore: true, attivo: true, creatoIl: serverTimestamp() });
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
  if (nome === 'gestione' && !stato.areaAperta) return;
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

// Il caposquadra è chi è entrato; se compila l'ufficio (responsabile, es. copiando un cartaceo) lo sceglie dall'elenco operai
const compilaUfficio = () => stato.utente?.ruolo === 'responsabile';

function riempiCapisquadra() {
  const sel = $('rCaposquadraScelta');
  const attuale = sel.value || ls.get(chiaveBozza())?.caposquadra || '';
  const voci = [...stato.elenchi.operai];
  if (attuale && !voci.includes(attuale)) voci.unshift(attuale);
  sel.innerHTML = '<option value="">— Scegli il caposquadra —</option>' + voci.map(n => `<option ${n === attuale ? 'selected' : ''}>${esc(n)}</option>`).join('');
  aggiornaFirmaCapo();
}
function aggiornaFirmaCapo() {
  if (!compilaUfficio()) { $('firmaCapo').textContent = stato.utente.nome; return; }
  const capo = $('rCaposquadraScelta').value;
  $('firmaCapo').textContent = `${capo || '…'} · compilato dall'ufficio (${stato.utente.nome})`;
}

function avviaCompilazione() {
  $('rCaposquadra').value = stato.utente.nome;
  $('rCaposquadra').hidden = compilaUfficio();
  $('rCaposquadraScelta').hidden = !compilaUfficio();
  aggiornaFirmaCapo();
  caricaBozza();
  const off = onSnapshot(doc(db, 'elenchi', 'correnti'), snap => {
    const d = snap.exists() ? snap.data() : {};
    stato.elenchi = { operai: d.operai || [], commesse: d.commesse || [], nascoste: d.nascoste || [], viaggi: d.viaggi || {}, file: d.file || '', aggiornatoIl: d.aggiornatoIl };
    $('erroreElenchi').hidden = stato.elenchi.operai.length > 0;
    $('erroreElenchi').textContent = 'Gli elenchi di operai e commesse non sono ancora stati caricati dall\'ufficio (Elenchi e utenti > Aggiorna elenchi).';
    riempiCommesse();
    riempiCapisquadra();
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
  // responsabili di commessa: li gestisce lo sviluppatore
  const offR = onSnapshot(doc(db, 'elenchi', 'responsabili'), snap => {
    const d = snap.exists() ? snap.data() : {};
    stato.responsabili = { voci: d.voci || [], predefinito: d.predefinito || '' };
    riempiResponsabili();
    disegnaGestioneResponsabili();
  }, () => {});
  stato.annulla.push(offR);
  // orario standard: lo imposta lo sviluppatore; finché non lo salva valgono gli orari di partenza
  const offO = onSnapshot(doc(db, 'elenchi', 'orarioStandard'), snap => {
    stato.orarioStandard = snap.exists() ? snap.data() : { ...ORARIO_STANDARD_BASE };
    aggiornaPulsanteStandard();
    disegnaGestioneOrarioStandard();
  }, () => {});
  stato.annulla.push(offO);
  const offD = onSnapshot(doc(db, 'elenchi', 'orarioDeposito'), snap => {
    stato.orarioDeposito = snap.exists() ? snap.data() : { ...ORARIO_DEPOSITO_BASE };
    aggiornaPulsanteStandard();
    disegnaGestioneOrarioDeposito();
  }, () => {});
  stato.annulla.push(offD);
  // intestazione del modulo (testi e logo): la decide l'ufficio
  const offI = onSnapshot(doc(db, 'elenchi', 'intestazione'), snap => {
    // un'intestazione salvata con la grafica vecchia (senza ragione sociale) non vale più: si usa la nuova, tenendo il logo
    const salvata = snap.exists() ? snap.data() : {};
    intestazione = salvata.ragioneSociale ? { ...INTESTAZIONE_BASE, logo: '', ...salvata } : { ...INTESTAZIONE_BASE, logo: salvata.logo || '' };
    aggiornaTestate();
    disegnaGestioneIntestazione();
  }, () => {});
  stato.annulla.push(offI);
  controllaCoda();
  recuperaCopie(stato.utente.uid).catch(() => {});
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
    <div class="mezzi-scelti">${scelti.map(m => {
      // l'autista si sceglie tra gli operai presenti in questo rapportino
      const autista = stato.autisti.get(m) || '';
      const presenti = stato.elenchi.operai.filter(n => stato.presenti.has(n));
      return `<div class="riga-mezzo"><span class="chip si">${esc(m)}
        <button type="button" class="togli-mezzo" data-togli="${esc(m)}" aria-label="Togli ${esc(m)}">✕</button></span>
        <label class="scelta-autista"><span>Autista:</span><select data-autista-mezzo="${esc(m)}" aria-label="Autista di ${esc(m)}">
          <option value="">${presenti.length ? '— Scegli —' : 'Segna prima i presenti'}</option>
          ${presenti.map(n => `<option ${n === autista ? 'selected' : ''}>${esc(n)}</option>`).join('')}
        </select></label></div>`;
    }).join('')}</div>`;
}
$('rAutomezzi').addEventListener('change', e => {
  const mezzoAutista = e.target.dataset.autistaMezzo;
  if (mezzoAutista !== undefined) {
    if (e.target.value) stato.autisti.set(mezzoAutista, e.target.value);
    else stato.autisti.delete(mezzoAutista);
    salvaBozza();
    return;
  }
  if (e.target.id !== 'sceltaMezzo' || !e.target.value) return;
  stato.mezziScelti.add(e.target.value);
  disegnaMezzi();
  salvaBozza();
});
$('rAutomezzi').addEventListener('click', e => {
  const m = e.target.dataset.togli;
  if (m === undefined) return;
  stato.mezziScelti.delete(m);
  stato.autisti.delete(m);
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
  commessaDiPrima = sel.value;
  aggiornaPulsanteStandard();
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
    <div class="o-nome" data-tocca title="${esc(nome)}"><span class="o-num">${i + 1}.</span><span class="o-testo">${esc(nome)}</span></div>
    <button type="button" class="o-p casella-p" data-tocca aria-pressed="${!!o}" aria-label="Presente: ${esc(nome)}"></button>
    <div class="o-ora o-partenza"><span class="o-et">Partenza deposito</span><input type="time" data-orario="partenza" value="${v('partenza')}"></div>
    <div class="o-ora o-cantiere"><span class="o-et">Dalle</span><input type="time" data-orario="dalle" value="${v('dalle')}"><span class="o-et">Alle</span><input type="time" data-orario="alle" value="${v('alle')}"></div>
    <div class="o-ora o-rientro"><span class="o-et">Rientro deposito</span><input type="time" data-orario="rientro" value="${v('rientro')}"></div>
  </div>`;
}

// Orari scritti nella riga (si leggono dallo schermo: niente si perde se un evento non è arrivato)
const CHIAVI_ORARI = ['partenza', 'dalle', 'alle', 'rientro'];
function orariRiga(div) {
  const o = {};
  CHIAVI_ORARI.forEach(k => { o[k] = div.querySelector(`[data-orario="${k}"]`)?.value || ''; });
  return o;
}
// Orari del riquadro "Stessi orari per tutti i presenti" (solo quelli scritti)
function orariComuni() {
  const v = {};
  document.querySelectorAll('[data-tutti]').forEach(inp => { if (inp.value) v[inp.dataset.tutti] = inp.value; });
  return v;
}
function segnaPresente(div) {
  const nome = div.dataset.nome;
  if (stato.presenti.has(nome)) stato.presenti.delete(nome);
  else {
    // chi viene segnato tiene gli orari già scritti nella riga; quelli vuoti si prendono dagli orari comuni
    const o = orariRiga(div);
    const comuni = orariComuni();
    CHIAVI_ORARI.forEach(k => { if (!o[k] && comuni[k]) o[k] = comuni[k]; });
    div.querySelectorAll('[data-orario]').forEach(inp => { inp.value = o[inp.dataset.orario]; });
    stato.presenti.set(nome, o);
  }
  const si = stato.presenti.has(nome);
  div.classList.toggle('presente', si);
  div.querySelector('.casella-p').setAttribute('aria-pressed', si);
  if (!si) div.querySelectorAll('[data-orario]').forEach(inp => { inp.value = ''; });
  aggiornaContaPresenti();
  salvaBozza();
}

$('listaOperai').addEventListener('click', e => {
  if (e.target.closest('[data-tocca]')) segnaPresente(e.target.closest('.operaio'));
});
// Orario scritto in una riga: se l'operaio non ha la X, la mette (chi ha un orario era presente)
function orarioCambiato(e) {
  const k = e.target.dataset.orario;
  if (!k) return;
  const div = e.target.closest('.operaio');
  const o = stato.presenti.get(div.dataset.nome);
  if (o) { o[k] = e.target.value; salvaBozza(); }
  else if (e.target.value) segnaPresente(div);
}
$('listaOperai').addEventListener('input', orarioCambiato);
$('listaOperai').addEventListener('change', orarioCambiato);
// "Stessi orari per tutti": appena si scrive un orario vale per tutti i presenti (il tasto resta)
document.querySelectorAll('[data-tutti]').forEach(inp => inp.addEventListener('change', () => {
  if (!inp.value) return;
  for (const o of stato.presenti.values()) o[inp.dataset.tutti] = inp.value;
  document.querySelectorAll(`#listaOperai .operaio.presente [data-orario="${inp.dataset.tutti}"]`).forEach(x => { x.value = inp.value; });
  salvaBozza();
}));

function filtraOperai() {
  const f = $('cercaOperaio').value.trim().toLowerCase();
  document.querySelectorAll('#listaOperai .operaio').forEach(div => {
    div.hidden = !!f && !div.dataset.nome.toLowerCase().includes(f);
  });
}
$('cercaOperaio').addEventListener('input', filtraOperai);

function aggiornaContaPresenti() {
  $('contaPresenti').textContent = stato.presenti.size;
  $('orariTutti').hidden = stato.presenti.size < 1;
  // l'autista deve essere tra i presenti: se un operaio viene tolto, il suo mezzo resta senza autista
  for (const [m, a] of [...stato.autisti]) if (!stato.presenti.has(a)) stato.autisti.delete(m);
  disegnaMezzi();
}

function applicaOrariATutti() {
  const valori = {};
  document.querySelectorAll('[data-tutti]').forEach(inp => { if (inp.value) valori[inp.dataset.tutti] = inp.value; });
  if (Object.keys(valori).length === 0) { avviso('Scrivi prima almeno un orario qui sopra', 'errore'); return; }
  for (const o of stato.presenti.values()) Object.assign(o, valori);
  document.querySelectorAll('#listaOperai .operaio.presente').forEach(div => {
    div.querySelectorAll('[data-orario]').forEach(inp => { if (valori[inp.dataset.orario]) inp.value = valori[inp.dataset.orario]; });
  });
  salvaBozza();
  const n = stato.presenti.size;
  avviso(`Orari applicati a ${n} ${n === 1 ? 'operaio' : 'operai'}`, 'ok');
}
$('btnApplicaTutti').onclick = applicaOrariATutti;

// Orario standard: si imposta nell'area sviluppatore (elenchi/orarioStandard). Il pulsante riempie il riquadro
// e applica subito gli orari ai presenti; le righe degli operai restano sempre modificabili.
const ORARIO_STANDARD_BASE = { nome: 'Orario standard 8 & 1', partenza: '07:00', dalle: '07:30', alle: '16:30', rientro: '17:00' };
// Solo per il deposito COMM.0 c'è un orario speciale (elenchi/orarioDeposito), che comprende anche la pausa.
const ORARIO_DEPOSITO_BASE = { nome: 'Orario magazzino 8 & 1', partenza: '06:30', dalle: '07:00', alle: '17:00', rientro: '17:30', mensaDalle: '12:00', mensaAlle: '14:00' };
const eDeposito = c => /^COMM\.0$/i.test(String(c || '').trim());
const haOrari = s => CHIAVI_ORARI.some(k => s[k]);
// Con COMM.0 vale l'orario speciale; se lì non ci sono orari, resta quello standard
const orarioDelPulsante = () => (eDeposito($('rCommessa').value) && haOrari(stato.orarioDeposito) ? stato.orarioDeposito : stato.orarioStandard);
function aggiornaPulsanteStandard() {
  const s = orarioDelPulsante();
  const speciale = s === stato.orarioDeposito;
  $('btnOrarioStandard').textContent = s.nome || (speciale ? ORARIO_DEPOSITO_BASE : ORARIO_STANDARD_BASE).nome;
  $('btnOrarioStandard').hidden = !haOrari(s);
}
$('btnOrarioStandard').onclick = () => {
  const s = orarioDelPulsante();
  document.querySelectorAll('[data-tutti]').forEach(inp => { inp.value = s[inp.dataset.tutti] || ''; });
  if (s === stato.orarioDeposito) {
    $('rMensaDalle').value = s.mensaDalle || '';
    $('rMensaAlle').value = s.mensaAlle || '';
    aggiornaTastoPausa();
  }
  applicaOrariATutti();
};
// Lasciando COMM.0 per un'altra commessa, la pausa del magazzino non deve restare per sbaglio
let commessaDiPrima = '';
$('rCommessa').addEventListener('change', () => {
  const era = commessaDiPrima;
  commessaDiPrima = $('rCommessa').value;
  const d = stato.orarioDeposito;
  if (eDeposito(era) && !eDeposito(commessaDiPrima) && d.mensaAlle && $('rMensaDalle').value === d.mensaDalle && $('rMensaAlle').value === d.mensaAlle) {
    $('rMensaDalle').value = PAUSA_DALLE;
    $('rMensaAlle').value = PAUSA_ALLE;
    aggiornaTastoPausa();
    salvaBozza();
  }
  aggiornaPulsanteStandard();
});
let orarioDepositoToccato = false;
function disegnaGestioneOrarioDeposito() {
  if (stato.utente?.ruolo !== 'responsabile') return;
  $('avvisoViaggioDeposito').hidden = viaggioMassimo('COMM.0') !== 0;
  if (orarioDepositoToccato) return;
  document.querySelectorAll('[data-deposito]').forEach(inp => { inp.value = stato.orarioDeposito[inp.dataset.deposito] || ''; });
}
$('formOrarioDeposito').addEventListener('input', () => { orarioDepositoToccato = true; });
$('formOrarioDeposito').addEventListener('submit', async e => {
  e.preventDefault();
  const v = {};
  document.querySelectorAll('[data-deposito]').forEach(inp => { v[inp.dataset.deposito] = inp.value.trim(); });
  v.nome = v.nome.replace(/\s+/g, ' ') || ORARIO_DEPOSITO_BASE.nome;
  try {
    await setDoc(doc(db, 'elenchi', 'orarioDeposito'), { ...v, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
    orarioDepositoToccato = false;
    disegnaGestioneOrarioDeposito();
    avviso('Orario di COMM.0 salvato', 'ok');
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
});
let orarioStandardToccato = false; // lo sviluppatore sta scrivendo: non sovrascrivere i suoi campi
function disegnaGestioneOrarioStandard() {
  if (stato.utente?.ruolo !== 'responsabile' || orarioStandardToccato) return;
  document.querySelectorAll('[data-standard]').forEach(inp => { inp.value = stato.orarioStandard[inp.dataset.standard] || ''; });
}
$('formOrarioStandard').addEventListener('input', () => { orarioStandardToccato = true; });
$('formOrarioStandard').addEventListener('submit', async e => {
  e.preventDefault();
  const v = {};
  document.querySelectorAll('[data-standard]').forEach(inp => { v[inp.dataset.standard] = inp.value.trim(); });
  v.nome = v.nome.replace(/\s+/g, ' ') || ORARIO_STANDARD_BASE.nome;
  try {
    await setDoc(doc(db, 'elenchi', 'orarioStandard'), { ...v, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
    orarioStandardToccato = false;
    disegnaGestioneOrarioStandard();
    avviso('Orario standard salvato', 'ok');
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// Bozza sul telefono finché il rapportino non viene inviato (se si chiude l'app non si perde il lavoro)
// Responsabile di commessa: resta la scelta della bozza, altrimenti il predefinito.
// Un nome tolto dall'elenco, o scritto a mano con una versione vecchia, resta scelto finché non lo si cambia.
function riempiResponsabili() {
  const sel = $('rResponsabile');
  const { voci, predefinito } = stato.responsabili;
  const attuale = ls.get(chiaveBozza())?.responsabile || predefinito;
  const tutte = attuale && !voci.includes(attuale) ? [attuale, ...voci] : voci;
  sel.innerHTML = '<option value="">— Scegli —</option>' + tutte.map(n => `<option ${n === attuale ? 'selected' : ''}>${esc(n)}</option>`).join('');
}

const PAUSA_DALLE = '12:00', PAUSA_ALLE = '13:00';
function aggiornaTastoPausa() {
  $('btnPausa').textContent = $('rMensaDalle').value || $('rMensaAlle').value ? 'Nessuna pausa' : `Rimetti ${PAUSA_DALLE}–${PAUSA_ALLE}`;
}
$('btnPausa').onclick = () => {
  const vuota = !$('rMensaDalle').value && !$('rMensaAlle').value;
  $('rMensaDalle').value = vuota ? PAUSA_DALLE : '';
  $('rMensaAlle').value = vuota ? PAUSA_ALLE : '';
  aggiornaTastoPausa();
  salvaBozza();
};
['rMensaDalle', 'rMensaAlle'].forEach(id => $(id).addEventListener('input', aggiornaTastoPausa));

function salvaBozza() {
  if (!stato.utente) return;
  const b = { commessa: $('rCommessa').value, data: $('rData').value, presenti: Object.fromEntries(stato.presenti), automezzi: [...stato.mezziScelti], autisti: Object.fromEntries(stato.autisti), caposquadra: $('rCaposquadraScelta').value };
  for (const [k, id] of Object.entries(CAMPI_TESTO)) b[k] = $(id).value;
  ls.set(chiaveBozza(), b);
}
function caricaBozza() {
  const b = ls.get(chiaveBozza()) || {};
  $('rData').value = b.data || oggiISO();
  for (const [k, id] of Object.entries(CAMPI_TESTO)) $(id).value = b[k] || '';
  // Pausa mensa già scritta per i rapportini nuovi; in una bozza si rispetta quello che c'è, anche se vuoto
  if (!('mensaDalle' in b)) { $('rMensaDalle').value = PAUSA_DALLE; $('rMensaAlle').value = PAUSA_ALLE; }
  aggiornaTastoPausa();
  riempiResponsabili();
  stato.presenti = new Map(Object.entries(b.presenti || {}));
  $('rCaposquadraScelta').value = '';
  riempiCapisquadra();
  stato.mezziScelti = new Set(Array.isArray(b.automezzi) ? b.automezzi : []);
  stato.autisti = new Map(Object.entries(b.autisti || {}));
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
$('rCaposquadraScelta').addEventListener('change', () => { salvaBozza(); aggiornaFirmaCapo(); });

$('btnSvuota').onclick = async () => {
  if (await conferma('Svuotare il rapportino?', '<p>Verranno tolti tutti i dati scritti finora.</p>', 'Svuota')) svuotaModulo();
};

$('formRapportino').addEventListener('submit', async e => {
  e.preventDefault();
  const commessa = $('rCommessa').value;
  const data = $('rData').value;
  const errori = [];
  if (!commessa) errori.push('scegli la <b>commessa</b>');
  if (compilaUfficio() && !$('rCaposquadraScelta').value) errori.push('scegli il <b>caposquadra</b>');
  if (!data) errori.push('scrivi la <b>data</b>');
  if (stato.presenti.size === 0) errori.push('segna almeno un <b>operaio presente</b> (casella P)');
  if (data && data > oggiISO()) errori.push('la data è nel futuro: controlla');
  const mezziInviati = stato.automezzi.filter(m => stato.mezziScelti.has(m));
  mezziInviati.filter(m => !stato.autisti.get(m)).forEach(m => errori.push(`scegli l'<b>autista</b> di <b>${esc(m)}</b>`));
  if (errori.length) {
    await finestra('Manca qualcosa', `<ul>${errori.map(x => `<li>${x}</li>`).join('')}</ul>`);
    return;
  }

  // gli orari si rileggono dallo schermo prima dell'invio
  document.querySelectorAll('#listaOperai .operaio.presente').forEach(div => {
    if (stato.presenti.has(div.dataset.nome)) stato.presenti.set(div.dataset.nome, orariRiga(div));
  });
  // ordine dell'elenco, come sul modulo
  const operai = stato.elenchi.operai.filter(n => stato.presenti.has(n)).map(nome => ({ nome, ...stato.presenti.get(nome) }));
  const senzaOrari = operai.filter(o => !o.dalle && !o.alle);
  if (senzaOrari.length && !await conferma('Mancano gli orari', `
    <p>${senzaOrari.length === operai.length ? 'Nessun operaio ha' : `${senzaOrari.length} operai su ${operai.length} non hanno`} l'orario di lavoro in cantiere (Dalle … Alle …):</p>
    <ul>${senzaOrari.map(o => `<li>${esc(o.nome)}</li>`).join('')}</ul>
    <p>Se hai scritto gli orari in "Stessi orari per tutti i presenti", controlla che compaiano nelle righe.</p>`, 'Continua lo stesso')) return;
  const autistiMezzi = mezziInviati.map(mezzo => ({ mezzo, autista: stato.autisti.get(mezzo) }));
  const doppi = [...new Set(autistiMezzi.map(x => x.autista).filter((a, i, v) => v.indexOf(a) !== i))];
  if (doppi.length && !await conferma('Stesso autista su più automezzi', `
    <ul>${doppi.map(a => `<li><b>${esc(a)}</b>: ${autistiMezzi.filter(x => x.autista === a).map(x => esc(x.mezzo)).join(', ')}</li>`).join('')}</ul>
    <p>È giusto? Può capitare se si sono alternati alla guida.</p>`, 'Continua lo stesso')) return;
  const ok = await conferma('Inviare il rapportino?', `
    <p><b>${dataIT(data)}</b> · commessa <b>${esc(commessa)}</b></p>
    <p>${operai.length} operai presenti:</p>
    <ul>${operai.map(o => `<li>${esc(o.nome)}${o.dalle || o.alle ? ` <span class="aiuto">(${esc(o.dalle)}–${esc(o.alle)})</span>` : ''}</li>`).join('')}</ul>
    ${autistiMezzi.length ? `<p>Automezzi:</p><ul>${autistiMezzi.map(x => `<li>${esc(x.mezzo)} – Autista: ${esc(x.autista)}</li>`).join('')}</ul>` : ''}
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
    caposquadraNome: compilaUfficio() ? $('rCaposquadraScelta').value : stato.utente.nome,
    compilatoDaUfficio: compilaUfficio(),
    compilatoDaNome: stato.utente.nome,
    operai,
    automezzi: mezziInviati, // nell'ordine dell'elenco
    autistiMezzi,
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
  const uid = stato.utente.uid;
  ls.set(chiave, (ls.get(chiave) || 0) + 1);
  let consegnato = false;
  salvaCopia(uid, ref.id, dati);

  const scrittura = setDoc(ref, dati).then(() => {
    consegnato = true;
    togliCopia(uid, ref.id);
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
    $('inviatoTesto').textContent = 'Ora non c\'è connessione: il rapportino è conservato sul telefono e partirà da solo. Quando torna il campo tieni aperta l\'app: può volerci qualche minuto. Non rifare il rapportino e non premere Esci.';
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
    if (bozza) { ls.set(chiaveBozza(), bozza); caricaBozza(); togliCopia(uid, ref.id); }
    mostraVista('compila');
    await finestra('Invio non riuscito', `<p>${esc(messaggioErrore(err))}</p><p>Il rapportino è stato rimesso nel modulo: puoi riprovare.</p>`);
  });
}

// Copia di sicurezza: se l'app viene chiusa, Firebase non può più dire che un invio è stato rifiutato
// (per esempio perché l'utente nel frattempo è stato disattivato). La copia resta sul telefono finché
// l'arrivo non è sicuro e all'apertura successiva il rapportino riparte, sempre con lo stesso identificativo.
const chiaveCopie = uid => `rapp_copie_${uid}`;
function salvaCopia(uid, id, dati) {
  const { inviatoIl, ...copia } = dati; // l'ora di invio la mette il server
  ls.set(chiaveCopie(uid), { ...(ls.get(chiaveCopie(uid)) || {}), [id]: copia });
}
function togliCopia(uid, id) {
  const copie = ls.get(chiaveCopie(uid)) || {};
  delete copie[id];
  if (Object.keys(copie).length) ls.set(chiaveCopie(uid), copie);
  else ls.del(chiaveCopie(uid));
}
async function recuperaCopie(uid) {
  const ids = Object.keys(ls.get(chiaveCopie(uid)) || {});
  if (!ids.length) return;
  // prima parte la coda normale di Firebase: la maggior parte delle copie si toglie così
  try { await waitForPendingWrites(db); } catch (_) { return; }
  for (const id of ids) {
    if (stato.utente?.uid !== uid) return;
    const copia = (ls.get(chiaveCopie(uid)) || {})[id];
    if (!copia) continue;
    const ref = doc(db, 'rapportini', id);
    // Il responsabile può leggere: se il rapportino c'è già non va mai riscritto, perderebbe le ore.
    // Il caposquadra non può leggere, ma il database gli rifiuta la scrittura su un rapportino già esistente.
    try {
      if ((await getDocFromServer(ref)).exists()) { togliCopia(uid, id); continue; }
    } catch (err) {
      if (err?.code !== 'permission-denied') return; // niente rete: si riprova alla prossima apertura
    }
    try {
      await setDoc(ref, { ...copia, inviatoIl: serverTimestamp() });
      togliCopia(uid, id);
      avviso(`Rapportino del ${dataIT(copia.data)} recuperato e inviato`, 'ok');
    } catch (err) {
      if (err?.code !== 'permission-denied') return;
      // rifiutato: se l'utente è attivo vuol dire che il rapportino è già arrivato; altrimenti la copia resta
      const profilo = await getDocFromServer(doc(db, 'utenti', uid)).catch(() => null);
      if (profilo?.exists() && profilo.data().attivo === true) togliCopia(uid, id);
    }
  }
}

function controllaCoda() {
  const n = ls.get(chiaveCoda()) || 0;
  $('avvisoCoda').hidden = n === 0;
  $('avvisoCoda').textContent = `⏳ ${n} rapportino/i in attesa di connessione: partiranno da soli. Quando torna il campo tieni aperta l'app: può volerci qualche minuto. Non rifarli e non premere Esci.`;
  if (n > 0) {
    waitForPendingWrites(db).then(() => { ls.del(chiaveCoda()); $('avvisoCoda').hidden = true; }).catch(() => {});
  }
}

$('btnNuovo').onclick = () => mostraVista('compila');

// ================================================================
// RESPONSABILE
// ================================================================
function avviaResponsabile() {
  // elenco dei rapportini arrivati, aggiornato in tempo reale
  const off = onSnapshot(query(collection(db, 'rapportini'), where('stato', '==', 'arrivato')), snap => {
    stato.arrivati = snap.docs.map(d => ({ id: d.id, ...d.data() }))
      .sort((a, b) => (b.data || '').localeCompare(a.data || '') || ((b.inviatoIl?.seconds || 0) - (a.inviatoIl?.seconds || 0)));
    disegnaArrivati();
  }, err => avviso(`Rapportini non disponibili: ${messaggioErrore(err)}`, 'errore'));
  stato.annulla.push(off);

  if (stato.utente.sviluppatore) {
    const offU = onSnapshot(collection(db, 'utenti'), snap => disegnaUtenti(snap.docs.map(d => ({ uid: d.id, ...d.data() }))),
      err => avviso(`Elenco utenti non aggiornato: ${messaggioErrore(err)} (ricarica la pagina)`, 'errore'));
    stato.annulla.push(offU);
    // PIN degli utenti (li legge solo lo sviluppatore)
    const offP = onSnapshot(collection(db, 'accessi'), snap => {
      stato.pinUtenti = Object.fromEntries(snap.docs.map(d => [d.id, d.data().pin]));
      if (stato.utentiLista) disegnaUtenti(stato.utentiLista);
    }, err => avviso(`PIN non disponibili: ${messaggioErrore(err)}`, 'errore'));
    stato.annulla.push(offP);
  }

  // anni nell'elenco Convalidati: dal 2026 (inizio del rapportino digitale) all'anno in corso
  const annoOggi = Number(oggiISO().slice(0, 4));
  if (!$('annoConvalidati').options.length) {
    for (let a = annoOggi; a >= Math.min(2026, annoOggi); a--) $('annoConvalidati').add(new Option(a, a));
  }
  // la cartella Excel serve solo sul PC: sul telefono o sul tablet non si controlla nemmeno
  if (dispositivoMobile()) {
    stato.cartellaPronta = false;
    aggiornaInfoScrittura();
  } else {
    CartellaExcel.restore().then(s => (s === 'pronta' ? cartellaPronta() : mostraAvvisoCartella(s)));
    CartellaArchivio.restore().then(s => (s === 'pronta' ? archivioPronto() : mostraAvvisoArchivio(s)));
  }
}

function voceRapportino(r) {
  const conv = r.stato === 'convalidato';
  const extra = conv
    ? `<span class="etichetta ok">Convalidato ${esc(dataOraIT(r.convalidatoIl))}</span>`
    : `<span class="etichetta">Inviato ${esc(dataOraIT(r.inviatoIl))}</span>`;
  // spunta "scritto in Excel" con file e giorno (dal salvataggio della convalida)
  const giorno = r.giornoExcel || Number(String(r.data || '').slice(8, 10));
  const excel = !conv ? '' : r.fileExcel
    ? `<div class="spunta-excel" title="Cartella: ${esc(r.cartellaExcel || '')}"><span class="spunta">✓</span>Scritto in Excel · ${esc(r.fileExcel)} · giorno ${giorno}</div>`
    : '<div class="spunta-excel no">✗ Non scritto in Excel</div>';
  const nOperai = (r.operai?.length || 0) - (r.operaiTolti?.length || 0) + (r.operaiAggiunti?.length || 0);
  return `<button class="voce ${conv && r.fileExcel ? 'scritta' : ''}" data-id="${esc(r.id)}">
    <div class="voce-princ">
      <div class="voce-titolo">${esc(dataIT(r.data))} · ${esc(r.commessa)}${r.localita ? ` · ${esc(r.localita)}` : ''}</div>
      <div class="voce-sotto">Caposquadra ${esc(r.caposquadraNome)} · ${nOperai} operai${r.correttoDaNome ? ' · corretto dall\'ufficio' : ''}</div>
      ${excel}
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
// Elenco Convalidati per ANNO → MESE → SETTIMANA (lunedì-domenica), in base alla DATA del rapportino
const mesiAperti = new Set(); // mesi aperti dall'utente (restano aperti quando l'elenco si aggiorna)

function caricaConvalidati() {
  if (offConvalidati) offConvalidati();
  const anno = $('annoConvalidati').value;
  if (!anno) return;
  const mesi = Array.from({ length: 12 }, (_, i) => `${anno}-${String(i + 1).padStart(2, '0')}`);
  offConvalidati = onSnapshot(query(collection(db, 'rapportini'), where('stato', '==', 'convalidato'), where('mese', 'in', mesi)), snap => {
    convalidati = snap.docs.map(d => ({ id: d.id, ...d.data() }))
      .sort((a, b) => (b.data || '').localeCompare(a.data || '') || ((b.convalidatoIl?.seconds || 0) - (a.convalidatoIl?.seconds || 0)));
    disegnaConvalidati();
  }, err => avviso(messaggioErrore(err), 'errore'));
  stato.annulla.push(() => offConvalidati && offConvalidati());
}
$('annoConvalidati').addEventListener('change', () => { mesiAperti.clear(); caricaConvalidati(); });

// Settimana ISO (lunedì-domenica): numero e lunedì della settimana di una data "AAAA-MM-GG"
function settimanaDi(iso) {
  const [a, m, g] = iso.split('-').map(Number);
  const d = new Date(Date.UTC(a, m - 1, g));
  const giornoSett = d.getUTCDay() || 7; // lunedì = 1 … domenica = 7
  const lunedi = new Date(d); lunedi.setUTCDate(d.getUTCDate() - giornoSett + 1);
  const giovedi = new Date(d); giovedi.setUTCDate(d.getUTCDate() + 4 - giornoSett);
  const inizioAnno = new Date(Date.UTC(giovedi.getUTCFullYear(), 0, 1));
  const numero = Math.ceil(((giovedi - inizioAnno) / 86400000 + 1) / 7);
  return { numero, lunedi: lunedi.toISOString().slice(0, 10) };
}
const ggmm = d => d.toISOString().slice(8, 10) + '/' + d.toISOString().slice(5, 7);

function disegnaConvalidati() {
  const box = $('listaConvalidati');
  const anno = $('annoConvalidati').value;
  $('totaleConvalidati').textContent = convalidati.length ? `${convalidati.length} rapportini nel ${anno}` : '';
  if (!convalidati.length) { box.innerHTML = `<div class="nessuno">Nessun rapportino convalidato nel ${esc(anno)}</div>`; return; }

  // raggruppa: mese -> settimana (chiave = lunedì) -> rapportini
  const perMese = new Map();
  for (const r of convalidati) {
    const mese = String(r.data || '').slice(0, 7);
    if (!perMese.has(mese)) perMese.set(mese, new Map());
    const s = settimanaDi(r.data);
    const sett = perMese.get(mese);
    if (!sett.has(s.lunedi)) sett.set(s.lunedi, { numero: s.numero, voci: [] });
    sett.get(s.lunedi).voci.push(r);
  }
  const mesiOrdinati = [...perMese.keys()].sort().reverse();
  if (!mesiAperti.size) mesiAperti.add(mesiOrdinati[0]); // il mese più recente parte aperto

  box.innerHTML = mesiOrdinati.map(mese => {
    const settimane = [...perMese.get(mese).entries()].sort((a, b) => b[0].localeCompare(a[0]));
    const totale = settimane.reduce((s, [, x]) => s + x.voci.length, 0);
    const [a, m] = mese.split('-').map(Number);
    return `<details class="periodo-mese" data-mese="${mese}" ${mesiAperti.has(mese) ? 'open' : ''}>
      <summary><span class="periodo-nome">${MESI[m - 1]} ${a}</span><span class="periodo-conta">${totale} rapportin${totale === 1 ? 'o' : 'i'}</span></summary>
      ${settimane.map(([lunedi, s]) => {
        const l = new Date(lunedi + 'T00:00:00Z'), dom = new Date(l); dom.setUTCDate(l.getUTCDate() + 6);
        return `<div class="periodo-settimana">
          <div class="settimana-titolo"><b>Settimana ${s.numero}</b><span>lun ${ggmm(l)} – dom ${ggmm(dom)}</span><span class="periodo-conta">${s.voci.length}</span></div>
          <div class="elenco">${s.voci.map(voceRapportino).join('')}</div>
        </div>`;
      }).join('')}
    </details>`;
  }).join('');
}
// ricorda quali mesi sono aperti, così restano aperti quando arriva un nuovo rapportino
$('listaConvalidati').addEventListener('toggle', e => {
  const d = e.target.closest('.periodo-mese');
  if (!d) return;
  if (d.open) mesiAperti.add(d.dataset.mese); else mesiAperti.delete(d.dataset.mese);
}, true);
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

// Orario uguale per tutti gli operai del rapportino ('' se diversi o mancanti)
const orarioComune = (r, k) => { const s = new Set((r.operai || []).map(o => o[k]).filter(Boolean)); return s.size === 1 ? [...s][0] : ''; };

// Ore proposte per un turno del blocco (l'ufficio le può sempre cambiare), come sul cartaceo:
// totale = uscita - entrata - pausa; prima Ord. D fino a 8 ore; quello che avanza va in VIAGGIO fino al
// viaggio vero (deposito→cantiere + cantiere→deposito, solo se entrata e uscita sono gli orari del deposito);
// il resto è Straord. D. Il viaggio pagato ha un massimo per commessa: es. senza pausa 07-17 → 8 + 1 viaggio + 1 straord.
// Il massimo si sceglie nell'area sviluppatore (elenchi/correnti, campo viaggi). Se non è stato scelto vale 1 ora,
// tranne il deposito COMM.0, dove il viaggio non esiste.
const VIAGGIO_MAX = 1;
const VIAGGIO_VERO = 'vero'; // nessun limite: si propone tutto il viaggio calcolato dagli orari
const viaggioMassimo = c => {
  const scelto = stato.elenchi.viaggi?.[c];
  if (scelto !== undefined) return scelto;
  return /^COMM\.0$/i.test(String(c || '').trim()) ? 0 : VIAGGIO_MAX;
};
function proponiOreTurno(b, t, r) {
  const riga = b.righe[t];
  const totale = differenzaOre(riga.entrata, riga.uscita) ? Math.max(0, differenzaOre(riga.entrata, riga.uscita) - num(riga.pausa)) : 0;
  const ore = oreVuote();
  if (totale > 0) {
    let viaggio = 0;
    // Gli orari sono quelli degli operai di questo blocco; un blocco senza operai del caposquadra guarda tutto il rapportino
    const delBlocco = (r.operai || []).filter(o => b.operai.includes(o.nome));
    const gruppo = delBlocco.length ? { operai: delBlocco } : r;
    const [p, d, a, ri] = ['partenza', 'dalle', 'alle', 'rientro'].map(k => orarioComune(gruppo, k));
    if (t === 0 && p && d && a && ri && riga.entrata === p && riga.uscita === ri) {
      const v = differenzaOre(p, d) + differenzaOre(a, ri);
      if (v > 0 && v < totale) viaggio = v;
    }
    const massimo = viaggioMassimo(riga.commessa || r.commessa);
    // "Viaggio vero" (cantieri lontani): il viaggio resta intero e le ordinarie sono il lavoro in cantiere,
    // es. 05:00-19:00 con lavori 08-16 e 1 di pausa = Ord. D 7 + Viaggio 6.
    // Altrimenti prima le 8 ordinarie; quello che avanza è viaggio (fino al massimo), il resto straordinario,
    // es. 07:00-17:00 meno 1 di pausa = 9, viaggio vero 2 → Ord. D 8 + Viaggio 1
    const viaggioIntero = massimo === VIAGGIO_VERO;
    const ordinarie = Math.min(8, viaggioIntero ? totale - viaggio : totale);
    const resto = totale - ordinarie;
    const vg = viaggioIntero ? viaggio : Math.min(viaggio, resto, massimo);
    ore.or = fmtOre(ordinarie);
    if (vg > 0) ore.v = fmtOre(vg);
    if (resto - vg > 0) ore.sd = fmtOre(resto - vg);
  }
  b.ore[`t${t + 1}`] = ore;
}

function bloccoIniziale(r) {
  const tutti = (r.operai || []).map(o => o.nome);
  const comune = k => orarioComune(r, k);
  const pausa = differenzaOre(r.mensaDalle, r.mensaAlle);
  const b = {
    righe: [
      // come sul cartaceo: entrata = partenza dal deposito, uscita = rientro in deposito (se mancano, gli orari in cantiere)
      { commessa: r.commessa || '', entrata: comune('partenza') || comune('dalle'), uscita: comune('rientro') || comune('alle'), pausa: pausa ? fmtOre(pausa) : '' },
      { commessa: '', entrata: '', uscita: '', pausa: '' }
    ],
    ore: { t1: oreVuote(), t2: oreVuote() },
    operai: tutti
  };
  proponiOreTurno(b, 0, r);
  return b;
}

// Operai con orari diversi: un blocco per ogni orario, così ognuno ha le ore giuste.
// Un orario vuoto non è un orario diverso: se nessuna colonna ha due valori in contrasto, il blocco resta uno solo.
const CAMPI_ORARI = ['partenza', 'dalle', 'alle', 'rientro'];
const chiaveOrari = o => CAMPI_ORARI.map(k => o?.[k] || '').join('|');
function blocchiIniziali(r) {
  const inContrasto = CAMPI_ORARI.some(k => new Set((r.operai || []).map(o => o[k]).filter(Boolean)).size > 1);
  if (!inContrasto) return [bloccoIniziale(r)];
  const gruppi = new Map();
  (r.operai || []).forEach(o => {
    const chiave = chiaveOrari(o);
    if (!gruppi.has(chiave)) gruppi.set(chiave, []);
    gruppi.get(chiave).push(o);
  });
  return [...gruppi.values()].map(operai => bloccoIniziale({ ...r, operai }));
}
// Il blocco in cui rimettere un operaio tolto: quello di chi ha il suo stesso orario
function bloccoDellOrario(blocchi, r, nome) {
  const orariDi = n => (r.operai || []).find(o => o.nome === n);
  const o = orariDi(nome);
  if (!o) return blocchi[0];
  return blocchi.find(b => b.operai.some(n => orariDi(n) && chiaveOrari(orariDi(n)) === chiaveOrari(o)))
    || blocchi.find(b => b.righe[0].entrata === (o.partenza || o.dalle) && b.righe[0].uscita === (o.rientro || o.alle))
    || blocchi[0];
}

function apriRapportino(r) {
  if (!r) return;
  stato.aperto = {
    id: r.id, dati: r,
    blocchi: (r.blocchi?.length ? structuredClone(r.blocchi) : blocchiIniziali(r)),
    aggiunti: [...(r.operaiAggiunti || [])], // operai non segnati dal caposquadra, aggiunti dall'ufficio
    tolti: [...(r.operaiTolti || [])]        // operai segnati per sbaglio dal caposquadra, tolti dall'ufficio
  };
  disegnaFronte(r);
  disegnaBlocchi();
  aggiornaInfoScrittura();
  const conv = r.stato === 'convalidato';
  $('btnConvalida').textContent = conv ? 'Riscrivi su Excel' : 'Convalida e scrivi su Excel';
  $('btnSalvaOre').hidden = conv;
  $('btnArchivia').hidden = !conv || dispositivoMobile();
  $('infoArchivioRapportino').innerHTML = !conv ? '' : r.archivioFile
    ? `✓ In archivio: <b>${esc(r.archivioPercorso || '')}/${esc(r.archivioFile)}</b>`
    : '<b>Non ancora nell\'archivio.</b>';
  $('infoFileExcel').innerHTML = conv
    ? `Convalidato da ${esc(r.convalidatoDaNome || '')} il ${esc(dataOraIT(r.convalidatoIl))}.
       ${percorsoHTML(r.cartellaExcel, r.fileExcel, r.giornoExcel || Number(String(r.data).slice(8, 10)),
         (r.righeExcel || []).map(x => { const [nome, turno, riga] = x.split('|'); return { nome, turno: Number(turno), riga: Number(riga) }; }))}`
    : '';
  mostraVista('dettaglio');
}
$('btnTornaElenco').onclick = () => mostraVista(stato.aperto?.dati?.stato === 'convalidato' ? 'convalidati' : 'arrivati');

// Tabella operai della pagina 1 vista dal responsabile.
// "Solo i presenti": chi ha la X più chi è stato aggiunto dall'ufficio.
// "Tutti": l'elenco completo del file Excel come sul modulo cartaceo; chi non ha la X ha la casella vuota e "aggiungi".
let mostraTuttiOperai = ls.get('rapp_mostra_tutti') === true;

function tabellaOperaiFronte(r) {
  const { tolti, aggiunti } = stato.aperto;
  const segnati = r.operai || [];
  const rigaSegnato = (o, num) => {
    const tolto = tolti.includes(o.nome);
    return `<div class="operaio presente ${tolto ? 'tolto' : ''}">
      <div class="o-nome"><span class="o-num">${num}.</span><span class="o-testo">${esc(o.nome)}</span>
        ${tolto ? '<span class="nota-corr">tolto dall\'ufficio</span>' : ''}
        <button type="button" class="corr" data-${tolto ? 'rimetti' : 'togli-p1'}="${esc(o.nome)}">${tolto ? 'rimetti' : 'togli'}</button></div>
      <div class="o-p casella-p"></div>
      <div class="o-ora o-partenza"><span class="o-et">Partenza deposito</span><b>${esc(o.partenza) || '…'}</b></div>
      <div class="o-ora o-cantiere"><span class="o-et">Dalle</span><b>${esc(o.dalle) || '…'}</b><span class="o-et">Alle</span><b>${esc(o.alle) || '…'}</b></div>
      <div class="o-ora o-rientro"><span class="o-et">Rientro deposito</span><b>${esc(o.rientro) || '…'}</b></div>
    </div>`;
  };
  const rigaAggiunto = (n, num) => `<div class="operaio presente aggiunto">
      <div class="o-nome"><span class="o-num">${num}.</span><span class="o-testo">${esc(n)}</span>
        <span class="nota-corr">aggiunto dall'ufficio</span>
        <button type="button" class="corr" data-togli-operaio="${esc(n)}">togli</button></div>
      <div class="o-p casella-p"></div>
      <div class="o-ora o-partenza"><b>…</b></div><div class="o-ora o-cantiere"><b>…</b></div><div class="o-ora o-rientro"><b>…</b></div>
    </div>`;
  const rigaVuota = (n, num) => `<div class="operaio assente">
      <div class="o-nome"><span class="o-num">${num}.</span><span class="o-testo">${esc(n)}</span>
        <button type="button" class="corr" data-aggiungi-riga="${esc(n)}">aggiungi</button></div>
      <div class="o-p casella-p"></div>
      <div class="o-ora o-partenza"><b>…</b></div><div class="o-ora o-cantiere"><span class="tenue">Dalle … Alle …</span></div><div class="o-ora o-rientro"><b>…</b></div>
    </div>`;

  let righe;
  if (mostraTuttiOperai && stato.elenchi.operai.length) {
    // ordine del file Excel; in fondo eventuali nomi del rapportino che non sono più nell'elenco
    const nomi = [...stato.elenchi.operai, ...segnati.map(o => o.nome).filter(n => !stato.elenchi.operai.includes(n))];
    righe = nomi.map((n, i) => {
      const o = segnati.find(x => x.nome === n);
      return o ? rigaSegnato(o, i + 1) : aggiunti.includes(n) ? rigaAggiunto(n, i + 1) : rigaVuota(n, i + 1);
    }).join('');
  } else {
    righe = segnati.map((o, i) => rigaSegnato(o, i + 1)).join('') +
      aggiunti.map((n, i) => rigaAggiunto(n, segnati.length + i + 1)).join('');
  }

  const presenti = tuttiGliOperai();
  const liberi = stato.elenchi.operai.filter(n => !presenti.includes(n) && !segnati.some(o => o.nome === n));
  return `<div class="vista-operai">
      <button type="button" class="btn piccolo" data-mostra-tutti>${mostraTuttiOperai
        ? 'Mostra solo i presenti'
        : `Mostra tutti gli operai (${stato.elenchi.operai.length})`}</button>
      <span class="tenue">${presenti.length} presenti</span>
    </div>
    <div class="tab-operai sola-lettura">
      <div class="operaio testa">
        <div class="o-nome">COGNOME E NOME</div><div class="o-p">P</div>
        <div class="o-ora o-partenza">${esc(titoloColonna(intestazioneDi(r), 'colPartenza'))}</div>
        <div class="o-ora o-cantiere">${esc(titoloColonna(intestazioneDi(r), 'colCantiere'))}</div>
        <div class="o-ora o-rientro">${esc(titoloColonna(intestazioneDi(r), 'colRientro'))}</div>
      </div>
      ${righe}
    </div>
    ${!mostraTuttiOperai && liberi.length ? `<div class="correggi-operai">
      <select data-aggiungi-p1 aria-label="Aggiungi operaio dimenticato dal caposquadra">
        <option value="">+ Aggiungi operaio dimenticato dal caposquadra</option>
        ${liberi.map(n => `<option>${esc(n)}</option>`).join('')}
      </select></div>` : ''}`;
}

// Pagina 1 in sola lettura, con la stessa grafica del modulo
function disegnaFronte(r) {
  const val = v => esc(v) || '&nbsp;';
  const testo = v => `<div class="r-valore testo">${val(v)}</div>`;
  $('dettaglioFronte').innerHTML = `
    ${testataModulo(1, intestazioneDi(r))}
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
    ${tabellaOperaiFronte(r)}
    <div class="riquadro">
      <div class="r-riga">
        <div class="r-cella largo"><span class="r-et">AUTOMEZZI UTILIZZATI:</span>${testo(Array.isArray(r.automezzi)
          // i rapportini inviati prima della versione 6.6 non hanno l'autista
          ? r.automezzi.map(m => { const a = (r.autistiMezzi || []).find(x => x.mezzo === m)?.autista; return a ? `${m} – Autista: ${a}` : m; }).join('\n')
          : r.automezzi)}</div>
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
        <div class="r-cella largo"><span class="r-et">Firma Caposquadra:</span><span class="firma">${esc(r.caposquadraNome)} · inviato il ${esc(dataOraIT(r.inviatoIl))}</span>${r.compilatoDaUfficio
          ? `<span class="tenue">compilato dall'ufficio (${esc(r.compilatoDaNome || '')})</span>` : ''}</div>
        <div class="r-cella"><span class="r-et">Firma RC:</span><span class="firma">${r.convalidatoDaNome ? `${esc(r.convalidatoDaNome)} · ${esc(dataOraIT(r.convalidatoIl))}` : '<span class="tenue">alla convalida</span>'}</span></div>
      </div>
      ${correzioniOperai() ? `<div class="r-riga"><div class="r-cella"><span class="r-et">Correzioni operai:</span>
        <span class="firma">${r.correttoDaNome && !correzioniDaSalvare() ? `${esc(r.correttoDaNome)} · ${esc(dataOraIT(r.correttoIl))}` : `${esc(stato.utente.nome)} <span class="tenue">(da salvare)</span>`}</span>
        <span class="tenue">${correzioniOperai()}</span></div></div>` : ''}
    </div>`;
}

// Correzioni dell'ufficio agli operai segnati dal caposquadra (testo breve per la riga delle firme)
function correzioniOperai() {
  const { tolti, aggiunti } = stato.aperto;
  return [tolti.length ? `tolti: ${tolti.join(', ')}` : '', aggiunti.length ? `aggiunti: ${aggiunti.join(', ')}` : ''].filter(Boolean).join(' · ');
}
// Le correzioni a schermo sono diverse da quelle salvate nel rapportino?
function correzioniDaSalvare() {
  const r = stato.aperto.dati, uguali = (a, b) => JSON.stringify([...a].sort()) === JSON.stringify([...(b || [])].sort());
  return !uguali(stato.aperto.tolti, r.operaiTolti) || !uguali(stato.aperto.aggiunti, r.operaiAggiunti);
}
// Campi da salvare nel rapportino insieme alle ore
function campiCorrezioni() {
  const c = { operaiAggiunti: stato.aperto.aggiunti, operaiTolti: stato.aperto.tolti };
  if (correzioniDaSalvare()) { c.correttoDaNome = stato.utente.nome; c.correttoDaUid = stato.utente.uid; c.correttoIl = serverTimestamp(); }
  return c;
}
function ridisegnaRapportino() {
  disegnaFronte(stato.aperto.dati);
  disegnaBlocchi();
}

// Pagina 1: togli / rimetti un operaio segnato dal caposquadra, aggiungi un operaio dimenticato
$('dettaglioFronte').addEventListener('click', e => {
  const { tolti, aggiunti, blocchi } = stato.aperto;
  const d = e.target.dataset;
  if (d.mostraTutti !== undefined) {
    // un clic: tutti gli operai dell'elenco <-> solo i presenti (la scelta resta per i prossimi rapportini)
    mostraTuttiOperai = !mostraTuttiOperai;
    ls.set('rapp_mostra_tutti', mostraTuttiOperai);
    disegnaFronte(stato.aperto.dati);
    return;
  }
  if (d.aggiungiRiga !== undefined) {
    // operaio senza X trovato nell'elenco completo: aggiunto dall'ufficio e messo nel primo blocco ore
    if (!aggiunti.includes(d.aggiungiRiga)) aggiunti.push(d.aggiungiRiga);
    if (blocchi[0] && !blocchi[0].operai.includes(d.aggiungiRiga)) blocchi[0].operai.push(d.aggiungiRiga);
  } else if (d.togliP1 !== undefined) {
    if (!tolti.includes(d.togliP1)) tolti.push(d.togliP1);
    blocchi.forEach(b => { b.operai = b.operai.filter(n => n !== d.togliP1); });
  } else if (d.rimetti !== undefined) {
    stato.aperto.tolti = tolti.filter(n => n !== d.rimetti);
    const casa = bloccoDellOrario(blocchi, stato.aperto.dati, d.rimetti);
    if (casa && !casa.operai.includes(d.rimetti)) casa.operai.push(d.rimetti);
  } else if (d.togliOperaio !== undefined) {
    stato.aperto.aggiunti = aggiunti.filter(n => n !== d.togliOperaio);
    blocchi.forEach(b => { b.operai = b.operai.filter(n => n !== d.togliOperaio); });
  } else return;
  ridisegnaRapportino();
});
$('dettaglioFronte').addEventListener('change', e => {
  if (e.target.dataset.aggiungiP1 === undefined || !e.target.value) return;
  const n = e.target.value;
  if (!stato.aperto.aggiunti.includes(n)) stato.aperto.aggiunti.push(n);
  const b0 = stato.aperto.blocchi[0];
  if (b0 && !b0.operai.includes(n)) b0.operai.push(n);
  ridisegnaRapportino();
});

function opzioniCommessa(scelta) {
  const voci = [...stato.elenchi.commesse];
  if (scelta && !voci.includes(scelta)) voci.unshift(scelta);
  return '<option value=""></option>' + voci.map(c => `<option ${c === scelta ? 'selected' : ''}>${esc(c)}</option>`).join('');
}

// Operai del rapportino: quelli segnati dal caposquadra più quelli aggiunti dall'ufficio
function tuttiGliOperai() {
  const segnati = (stato.aperto.dati.operai || []).map(o => o.nome).filter(n => !stato.aperto.tolti.includes(n));
  return [...segnati, ...stato.aperto.aggiunti.filter(n => !segnati.includes(n))];
}

function disegnaBlocchi() {
  const { blocchi, aggiunti } = stato.aperto;
  const presenti = tuttiGliOperai();
  // per il menu "+ Aggiungi operaio": chi è nell'elenco del file Excel ma non ancora nel rapportino
  // (chi è stato segnato dal caposquadra e poi tolto non compare: si rimette con "rimetti" nella pagina 1)
  const segnatiCapo = (stato.aperto.dati.operai || []).map(o => o.nome);
  const liberi = stato.elenchi.operai.filter(n => !presenti.includes(n) && !segnatiCapo.includes(n));
  $('blocchi').innerHTML = blocchi.map((b, i) => `
    <div class="blocco" data-b="${i}">
      <div class="blocco-testa">
        <span>TURNO DI LAVORO ${blocchi.length > 1 ? i + 1 : ''}</span>
        <button class="btn piccolo" data-azione="proponi" title="Ricalcola Ord. D, Straord. D e Viaggio da entrata, uscita e pausa">Proponi ore</button>
        <button class="btn piccolo" data-azione="soloviaggio" title="Giornata di solo viaggio: tutte le ore da entrata a uscita, meno la pausa, vanno in Viaggio">Solo viaggio</button>
        ${blocchi.length > 1 ? `<button class="btn piccolo pericolo" data-azione="togli">Togli blocco</button>` : ''}
      </div>
      <div class="blocco-corpo">
        <div class="tabella-scorri"><table class="tab">
          <thead><tr><th></th><th>Commessa</th><th>Entrata</th><th>Uscita</th><th>Sub-ore</th><th>A detrarre pausa</th><th>Parziali</th><th>Totale</th></tr></thead>
          <tbody>${b.righe.map((r, t) => `<tr>
            <th>${t + 1}° turno</th>
            <td><select data-riga="${t}" data-k="commessa">${opzioniCommessa(r.commessa)}</select></td>
            <td><input type="time" data-riga="${t}" data-k="entrata" value="${esc(r.entrata)}"></td>
            <td><input type="time" data-riga="${t}" data-k="uscita" value="${esc(r.uscita)}"></td>
            <td class="calc" data-calc="sub${t}"></td>
            <td><input class="ore" inputmode="decimal" data-riga="${t}" data-k="pausa" value="${esc(r.pausa)}"></td>
            <td class="calc" data-calc="parz${t}"></td>
            ${t === 0 ? `<td class="calc totale-turni" rowspan="${b.righe.length}" data-calc="totTurni"></td>` : ''}
          </tr>`).join('')}
          </tbody>
        </table></div>
        <div class="blocco-sotto">
          <div class="box-operai">
            <div class="box-operai-titolo">OPERAI DI QUESTO BLOCCO</div>
            <div class="aiuto">Spunta:
              <button class="link" data-azione="tutti">tutti</button> ·
              <button class="link" data-azione="nessuno">nessuno</button></div>
            <div class="blocco-operai">${presenti.map(n => {
              const extra = aggiunti.includes(n);
              return `<label class="chip ${b.operai.includes(n) ? 'si' : ''} ${extra ? 'aggiunto' : ''}" ${extra ? 'title="Aggiunto dall\'ufficio"' : ''}>
                <input type="checkbox" data-operaio="${esc(n)}" ${b.operai.includes(n) ? 'checked' : ''}>${esc(n)}${extra
                  ? ` <button type="button" class="togli-mezzo" data-togli-operaio="${esc(n)}" aria-label="Togli ${esc(n)} dal rapportino">✕</button>` : ''}</label>`;
            }).join('')}</div>
            ${liberi.length ? `<select data-aggiungi-operaio aria-label="Aggiungi operaio">
              <option value="">+ Aggiungi operaio</option>
              ${liberi.map(n => `<option>${esc(n)}</option>`).join('')}
            </select>` : ''}
            ${aggiunti.length ? '<div class="aiuto">Con il bordo tratteggiato: aggiunti dall\'ufficio.</div>' : ''}
          </div>
          <div class="tabella-scorri tab-ore"><table class="tab">
            <thead><tr><th></th><th>1° turno</th><th>2° turno</th></tr></thead>
            <tbody>${VOCI_ORE.map(([k, et]) => `<tr><th>${et}</th>
              <td><input class="ore" inputmode="decimal" data-t="t1" data-k="${k}" value="${esc(b.ore.t1[k])}"></td>
              <td><input class="ore" inputmode="decimal" data-t="t2" data-k="${k}" value="${esc(b.ore.t2[k])}"></td></tr>`).join('')}
              <tr><th>Totale ore</th><td class="calc" data-calc="ore1"></td><td class="calc" data-calc="ore2"></td></tr>
            </tbody>
          </table></div>
        </div>
      </div>
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
$('blocchi').addEventListener('change', e => {
  // operaio non segnato dal caposquadra: entra nel rapportino e in questo blocco
  if (e.target.dataset.aggiungiOperaio !== undefined) {
    const n = e.target.value;
    if (!n) return;
    const b = stato.aperto.blocchi[parseInt(e.target.closest('.blocco').dataset.b)];
    if (!stato.aperto.aggiunti.includes(n)) stato.aperto.aggiunti.push(n);
    if (!b.operai.includes(n)) b.operai.push(n);
    ridisegnaRapportino();
    return;
  }
  if (e.target.tagName === 'SELECT') e.target.dispatchEvent(new Event('input', { bubbles: true }));
});
$('blocchi').addEventListener('click', e => {
  const togli = e.target.dataset.togliOperaio;
  if (togli !== undefined) {
    // toglie dal rapportino un operaio aggiunto dall'ufficio (da tutti i blocchi)
    e.preventDefault();
    stato.aperto.aggiunti = stato.aperto.aggiunti.filter(n => n !== togli);
    stato.aperto.blocchi.forEach(b => { b.operai = b.operai.filter(n => n !== togli); });
    ridisegnaRapportino();
    return;
  }
  const az = e.target.dataset.azione;
  if (!az) return;
  const i = parseInt(e.target.closest('.blocco').dataset.b);
  const bl = stato.aperto.blocchi;
  if (az === 'togli') bl.splice(i, 1);
  if (az === 'tutti') bl[i].operai = tuttiGliOperai();
  if (az === 'nessuno') bl[i].operai = [];
  if (az === 'soloviaggio') {
    // giornata di solo viaggio (es. trasferta a Milano): uscita - entrata - pausa, tutto in Viaggio; la pausa si toglie solo se c'è
    const ore = bl[i].righe.map(r => { const sub = differenzaOre(r.entrata, r.uscita); return sub ? Math.max(0, sub - num(r.pausa)) : 0; });
    if (!ore.some(v => v > 0)) { avviso('Per il solo viaggio servono entrata e uscita nel riquadro', 'errore'); return; }
    ore.forEach((v, t) => { bl[i].ore[`t${t + 1}`] = { ...oreVuote(), ...(v > 0 ? { v: fmtOre(v) } : {}) }; });
    avviso('Ore di solo viaggio: controllale prima di convalidare', 'ok');
  }
  if (az === 'proponi') {
    // ricalcola le ore dei due turni da entrata, uscita e pausa (sovrascrive quelle scritte)
    [0, 1].forEach(t => proponiOreTurno(bl[i], t, stato.aperto.dati));
    avviso('Ore proposte: controllale prima di convalidare', 'ok');
  }
  disegnaBlocchi();
  aggiornaRiepilogo();
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
    if (qualcosa && b.operai.length === 0) errori.push(`${nomeB}: spunta gli operai a cui vale, oppure togli il blocco`);
    if (!qualcosa && b.operai.length > 0) errori.push(`${nomeB}: scrivi le ore nel riquadro 1°/2° turno`);
  });
  const conOre = new Set(records.map(r => r.workerName));
  const senza = tuttiGliOperai().filter(n => !conOre.has(n));
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
    const campi = { blocchi: stato.aperto.blocchi, ...campiCorrezioni() };
    await updateDoc(doc(db, 'rapportini', stato.aperto.id), campi);
    Object.assign(stato.aperto.dati, campi, campi.correttoIl ? { correttoIl: new Date() } : {});
    disegnaFronte(stato.aperto.dati);
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

// ---------------------------------------------------------------- stampa (fronte e retro come il modulo)
// Fronte: la pagina 1 com'è a schermo (presenti o tutti gli operai, con le correzioni).
// Retro: per ogni blocco la tabella TURNO DI LAVORO, il riquadro ore 1°/2° turno, gli operai e Firma RC / Convalida.
function stampaRetro(r) {
  const vuoto = v => (v ? fmtOre(v) : '');
  const blocchi = stato.aperto.blocchi.map(b => {
    let tot = 0;
    const righe = b.righe.map(rg => {
      const sub = differenzaOre(rg.entrata, rg.uscita);
      const parz = sub ? Math.max(0, sub - num(rg.pausa)) : 0;
      tot += parz;
      return { ...rg, sub, parz };
    });
    const totOre = t => VOCI_ORE.reduce((s, [k]) => s + num(b.ore[t][k]), 0);
    return `<div class="retro-blocco">
      <table class="retro-tab retro-turno">
        <caption>TURNO DI LAVORO</caption>
        <thead><tr><th>Commessa</th><th>Entrata</th><th>Uscita</th><th>Sub-ore</th><th>A detrarre PAUSA</th><th>Parziali</th><th>Totale</th></tr></thead>
        <tbody>${righe.map((rg, t) => `<tr>
          <td>${esc(rg.commessa)}</td><td>${esc(rg.entrata)}</td><td>${esc(rg.uscita)}</td>
          <td>${vuoto(rg.sub)}</td><td>${esc(rg.pausa)}</td><td>${vuoto(rg.parz)}</td>
          ${t === 0 ? `<td rowspan="${righe.length}" class="retro-tot">${vuoto(tot)}</td>` : ''}</tr>`).join('')}</tbody>
      </table>
      <div class="retro-sotto">
        <div class="retro-operai"><b>Operai:</b> ${b.operai.map(esc).join(', ') || '—'}</div>
        <table class="retro-tab retro-ore">
          <thead><tr><th></th><th>1° TURNO</th><th>2° TURNO</th></tr></thead>
          <tbody>${VOCI_ORE.map(([k, et]) => `<tr><th>${et}</th><td>${vuoto(num(b.ore.t1[k]))}</td><td>${vuoto(num(b.ore.t2[k]))}</td></tr>`).join('')}
            <tr class="retro-totale"><th>TOTALE ORE</th><td colspan="2">${vuoto(totOre('t1') + totOre('t2'))}</td></tr></tbody>
        </table>
      </div>
      <table class="retro-tab retro-firma">
        <tr><th>Firma RC</th><td class="firma">${esc(r.convalidatoDaNome || '')}</td></tr>
        <tr><th>Convalida</th><td>${esc(r.convalidatoIl ? dataOraIT(r.convalidatoIl) : '')}</td></tr>
      </table>
    </div>`;
  }).join('');
  return `<section class="stampa-pagina modulo stampa-retro">${testataModulo(2, intestazioneDi(r))}${blocchi}</section>`;
}

$('btnStampa').onclick = () => {
  const r = stato.aperto.dati;
  $('stampa').innerHTML = `<section class="stampa-pagina modulo">${$('dettaglioFronte').innerHTML}</section>${stampaRetro(r)}`;
  const titolo = document.title;
  document.title = `Rapportino ${dataIT(r.data).replace(/\//g, '-')} ${r.commessa}`; // nome proposto per il PDF
  window.print();
  document.title = titolo;
};

// ---------------------------------------------------------------- scrittura su Excel
// Telefono o tablet: su Excel si scrive solo dal PC dell'ufficio
function dispositivoMobile() {
  return navigator.userAgentData?.mobile === true
    || /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent)
    || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1); // iPad che si presenta come Mac
}
function mostraAvvisoCartella(s) {
  stato.cartellaPronta = false;
  // la barra è solo per il responsabile sul PC (anche se il controllo finisce dopo un cambio di utente)
  if (stato.utente?.ruolo !== 'responsabile' || dispositivoMobile()) {
    $('avvisoCartella').hidden = true;
    aggiornaInfoScrittura();
    return;
  }
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
  aggiornaInfoScrittura();
}
function cartellaPronta() {
  stato.cartellaPronta = true;
  $('avvisoCartella').hidden = true;
  $('infoCartella').textContent = `Cartella collegata: ${CartellaExcel.name()}`;
  aggiornaInfoScrittura();
}

// Sotto "Convalida": dove scriverà l'app (nome della cartella e file Excel che contiene, per riconoscerla)
async function aggiornaInfoScrittura() {
  const el = $('infoScrittura');
  if (!el) return;
  if (dispositivoMobile()) { el.textContent = 'La scrittura su Excel si fa dal PC dell\'ufficio.'; return; }
  if (!stato.cartellaPronta) { el.innerHTML = `<b>Cartella Excel non collegata</b> (${stato.utente?.sviluppatore ? 'Elenchi e utenti → Cartella dei file Excel' : 'barra gialla in alto'}).`; return; }
  try {
    const files = await CartellaExcel.listExcelFiles();
    el.innerHTML = `Scrive nella cartella <b>${esc(CartellaExcel.name())}</b>, che contiene: ${files.length ? files.map(esc).join(', ') : 'nessun file Excel'}.`;
  } catch (_) { el.textContent = `Scrive nella cartella ${CartellaExcel.name()}.`; }
}
$('btnCartella').onclick = async () => {
  try {
    if ($('btnCartella').dataset.azione === 'conferma' && await CartellaExcel.confirmAccess()) return cartellaPronta();
    await CartellaExcel.choose();
    cartellaPronta();
  } catch (err) { if (err.name !== 'AbortError') avviso(`Cartella non collegata: ${err.message}`, 'errore'); }
};
$('btnCambiaCartella').onclick = async () => {
  if (dispositivoMobile()) { avviso('La cartella Excel si collega dal PC dell\'ufficio', 'errore'); return; }
  try { await CartellaExcel.choose(); cartellaPronta(); avviso(`Cartella collegata: ${CartellaExcel.name()}`, 'ok'); }
  catch (err) { if (err.name !== 'AbortError') avviso(err.message, 'errore'); }
};
async function richiediCartella(perConvalida = false) {
  if (dispositivoMobile()) {
    await finestra('Serve il PC dell\'ufficio', perConvalida
      ? '<p>Per scrivere su Excel usa il PC dell\'ufficio. Le ore le puoi salvare con "Salva senza convalidare".</p>'
      : '<p>I file Excel si usano solo dal PC dell\'ufficio.</p>');
    return false;
  }
  if (!CartellaExcel.supported()) {
    await finestra('Excel non raggiungibile', '<p>Per scrivere su Excel usa il PC dell\'ufficio con Microsoft Edge o Google Chrome.</p>');
    return false;
  }
  // rilegge ogni volta la cartella salvata: se è stata cambiata (anche da un'altra scheda) si usa quella nuova
  if (!window.__prova) {
    const s = await CartellaExcel.restore();
    if (s === 'pronta' || (s === 'da-confermare' && await CartellaExcel.confirmAccess().catch(() => false))) {
      cartellaPronta();
      return true;
    }
    mostraAvvisoCartella(s);
  } else if (stato.cartellaPronta) return true;
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

// Dove sono stati scritti i dati: cartella collegata, file, foglio, giorno e riga di ogni operaio.
// Il browser non rivela il percorso completo (es. \\NAS\...), solo il nome della cartella collegata.
function percorsoHTML(cartella, file, giorno, righe = []) {
  const inizio = 2 + (giorno - 1) * 69;
  return `<div class="percorso">
    <div><span>Cartella</span><b>${esc(cartella || '(cartella collegata)')}</b></div>
    <div><span>File</span><b>${esc(file || '')}</b></div>
    <div><span>Foglio</span><b>Presenze Master- OPERAI</b></div>
    <div><span>Giorno</span><b>${giorno} (righe da ${inizio} a ${inizio + 68})</b></div>
    ${righe.length ? `<details><summary>Riga di ogni operaio</summary>
      <table class="tab"><thead><tr><th>Riga</th><th>Operaio</th><th>Turno</th></tr></thead>
      <tbody>${righe.map(x => `<tr><td>${x.riga}</td><td>${esc(x.nome)}</td><td>${x.turno}°</td></tr>`).join('')}</tbody></table>
    </details>` : ''}
  </div>`;
}

const etichettaTurno =v => [v.commessa || '(senza commessa)', ...VOCI_ORE.filter(([k]) => v[k]).map(([k, et]) => `${et} ${fmtOre(v[k])}`)].join(' · ');

$('btnConvalida').onclick = async () => {
  const { dati } = stato.aperto;
  const { records, errori, avvisi } = calcolaRecord();
  if (errori.length) { await finestra('Da sistemare prima di scrivere', `<ul>${errori.map(e => `<li>${esc(e)}</li>`).join('')}</ul>`); return; }
  if (avvisi.length && !await conferma('Controlla', `<ul>${avvisi.map(a => `<li>${esc(a)}</li>`).join('')}</ul><p>Vuoi scrivere lo stesso?</p>`, 'Scrivi lo stesso')) return;
  if (!await richiediCartella(true)) return;

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
      if (!await CartellaExcel.hasBackupToday(fileName)) {
        try { await CartellaExcel.datedBackup(fileName, bytes); }
        catch (e) { throw new Error(`copia di sicurezza non riuscita (${e.message}): Excel non è stato modificato`); }
      }
      const risultato = await ExcelXlsm.writePresenze(bytes, giorno, daScrivere);
      await CartellaExcel.writeBytes(fileName, risultato.bytes);
    }
    // riga Excel di ogni turno scritto (stessa formula di excel.js: 69 righe per giorno, 2 per operaio)
    const righe = daScrivere.map(r => ({ nome: r.workerName, turno: r.turno, riga: 2 + (giorno - 1) * 69 + (r.workerIndex - 1) * 2 + (r.turno - 1) }))
      .sort((a, b) => a.riga - b.riga);

    await updateDoc(doc(db, 'rapportini', dati.id), {
      stato: 'convalidato',
      blocchi: stato.aperto.blocchi,
      ...campiCorrezioni(),
      convalidatoIl: serverTimestamp(),
      convalidatoDaUid: stato.utente.uid,
      convalidatoDaNome: stato.utente.nome,
      fileExcel: fileName,
      cartellaExcel: CartellaExcel.name(),
      giornoExcel: giorno,
      righeExcel: righe.map(x => `${x.nome}|${x.turno}|${x.riga}`)
    });
    // Copia nell'archivio: Excel è già scritto, quindi un errore qui non ferma niente e si può rifare dal rapportino aperto
    const finale = {
      ...dati, stato: 'convalidato', blocchi: stato.aperto.blocchi,
      operaiAggiunti: stato.aperto.aggiunti, operaiTolti: stato.aperto.tolti,
      convalidatoIl: new Date(), convalidatoDaUid: stato.utente.uid, convalidatoDaNome: stato.utente.nome,
      fileExcel: fileName, cartellaExcel: CartellaExcel.name(), giornoExcel: giorno
    };
    const esitoArchivio = await archiviaRapportino(finale).then(
      p => `<p class="ok-box">✓ Copia salvata nell'archivio: <b>${esc(p)}</b></p>`,
      e => `<p class="errore">⚠ Copia NON salvata nell'archivio (${esc(e.message)}). Excel è a posto: collega la cartella, riapri il rapportino dai Convalidati e premi "Salva nell'archivio".</p>`);
    await finestra('Convalidato', `<p>Scritti <b>${daScrivere.length}</b> turni.${saltati ? ` ${saltati} lasciati com'erano perché già presenti.` : ''}</p>
      ${percorsoHTML(CartellaExcel.name(), fileName, giorno, righe)}
      ${esitoArchivio}
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

// ---------------------------------------------------------------- archivio dei rapportini
function mostraAvvisoArchivio(s) {
  stato.archivioPronto = false;
  $('infoArchivio').textContent = 'Cartella dell\'archivio non collegata.';
  if (stato.utente?.ruolo !== 'responsabile' || dispositivoMobile() || !CartellaExcel.supported()) { $('avvisoArchivio').hidden = true; return; }
  $('avvisoArchivio').hidden = false;
  if (s === 'da-confermare') {
    $('avvisoArchivioTesto').textContent = `Conferma l'accesso alla cartella "${CartellaArchivio.name()}" dell'archivio rapportini.`;
    $('btnArchivioBarra').textContent = 'Consenti accesso';
    $('btnArchivioBarra').dataset.azione = 'conferma';
  } else {
    $('avvisoArchivioTesto').textContent = 'Scegli la cartella dell\'archivio rapportini: lì l\'app salva una copia di ogni rapportino convalidato.';
    $('btnArchivioBarra').textContent = 'Scegli cartella archivio';
    $('btnArchivioBarra').dataset.azione = 'scegli';
  }
}
function archivioPronto() {
  stato.archivioPronto = true;
  $('avvisoArchivio').hidden = true;
  $('infoArchivio').textContent = `Cartella dell'archivio: ${CartellaArchivio.name()}`;
}
async function scegliArchivio() {
  try {
    if ($('btnArchivioBarra').dataset.azione === 'conferma' && await CartellaArchivio.confirmAccess()) return archivioPronto();
    await CartellaArchivio.choose();
    archivioPronto();
    avviso(`Cartella dell'archivio: ${CartellaArchivio.name()}`, 'ok');
  } catch (err) { if (err.name !== 'AbortError') avviso(`Archivio non collegato: ${err.message}`, 'errore'); }
}
$('btnArchivioBarra').onclick = scegliArchivio;
$('btnCambiaArchivio').onclick = () => {
  if (dispositivoMobile()) { avviso('La cartella dell\'archivio si sceglie dal PC dell\'ufficio', 'errore'); return; }
  $('btnArchivioBarra').dataset.azione = 'scegli';
  scegliArchivio();
};
// Accesso alla cartella: la rilegge ogni volta e, se serve, chiede la conferma
async function archivioAccessibile() {
  if (window.__prova) return !!stato.archivioPronto;
  const s = await CartellaArchivio.restore();
  if (s === 'pronta') { archivioPronto(); return true; }
  if (s === 'da-confermare' && await CartellaArchivio.confirmAccess().catch(() => false)) { archivioPronto(); return true; }
  mostraAvvisoArchivio(s);
  return false;
}

// Dove va il file: Anno/MM-Mese/AAAA-MM-GG_commessa_caposquadra_codice.html
const nomeSicuro = t => String(t || '').replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();
function posizioneArchivio(r) {
  const [anno, mese] = String(r.data).split('-');
  const percorso = [anno, `${mese}-${MESI[Number(mese) - 1]}`];
  return { percorso, testo: percorso.join('/'), file: `${r.data}_${nomeSicuro(r.commessa)}_${nomeSicuro(r.caposquadraNome)}_${String(r.id).slice(0, 5)}.html` };
}
// Una pagina completa e autonoma: stessi stili della stampa e logo dentro il file, si apre con un doppio clic anche senza internet
const inBase64 = async url => {
  const b = await (await fetch(url)).blob();
  return new Promise((ok, no) => { const f = new FileReader(); f.onload = () => ok(f.result); f.onerror = () => no(f.error); f.readAsDataURL(b); });
};
async function paginaArchivio(r) {
  disegnaFronte(r);
  const corpo = `<section class="stampa-pagina modulo">${$('dettaglioFronte').innerHTML}</section>${stampaRetro(r)}`;
  const stili = (await (await fetch('style.css')).text()).replaceAll('@media print', '@media all');
  const logo = await inBase64(LOGO_BASE);
  return `<!doctype html>
<html lang="it" data-tema="chiaro"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Rapportino ${esc(dataIT(r.data))} ${esc(r.commessa)} - ${esc(r.caposquadraNome)}</title>
<style>${stili}
#stampa { max-width: 760px; margin: 0 auto; }
@media screen { #stampa { padding: 12px; } .stampa-pagina { margin-bottom: 24px; } }
</style></head><body><div id="stampa" class="stampa">${corpo.split(`src="${LOGO_BASE}"`).join(`src="${logo}"`)}</div></body></html>`;
}
// Salva il rapportino nell'archivio e lo segna nel database. Restituisce il percorso; se non riesce lancia l'errore.
async function archiviaRapportino(r) {
  if (!await archivioAccessibile()) throw new Error('cartella dell\'archivio non collegata');
  const pos = posizioneArchivio(r);
  await CartellaArchivio.salva(pos.percorso, pos.file, await paginaArchivio(r));
  const campi = { archivioFile: pos.file, archivioPercorso: pos.testo, archivioCartella: CartellaArchivio.name(), archiviatoIl: serverTimestamp() };
  await updateDoc(doc(db, 'rapportini', r.id), campi);
  Object.assign(r, campi);
  return `${CartellaArchivio.name()}/${pos.testo}/${pos.file}`;
}
$('btnArchivia').onclick = async () => {
  const r = stato.aperto?.dati;
  if (!r) return;
  try {
    const p = await archiviaRapportino(r);
    $('infoArchivioRapportino').innerHTML = `✓ In archivio: <b>${esc(r.archivioPercorso)}/${esc(r.archivioFile)}</b>`;
    avviso(`Salvato nell'archivio: ${p}`, 'ok');
  } catch (err) { avviso(`Non salvato nell'archivio: ${err.message}`, 'errore'); }
};

// Cancellazione dei mesi vecchi: nell'app restano il mese in corso e i 4 precedenti (a giugno si può cancellare gennaio)
const MESI_DA_TENERE = 5;
function mesiCancellabili() {
  const [a, m] = (window.__prova?.oggi || oggiISO()).split('-').map(Number);
  const primoDaTenere = a * 12 + (m - 1) - (MESI_DA_TENERE - 1);
  const elenco = [];
  for (let x = 2026 * 12; x < primoDaTenere; x++) elenco.push(`${Math.floor(x / 12)}-${String(x % 12 + 1).padStart(2, '0')}`);
  return elenco;
}
const nomeMese = m => `${MESI[Number(m.slice(5, 7)) - 1]} ${m.slice(0, 4)}`;
async function controllaMesiVecchi() {
  const box = $('mesiVecchi');
  const mesi = mesiCancellabili();
  box.innerHTML = '<div class="nessuno">Controllo in corso…</div>';
  try {
    const trovati = [];
    for (let i = 0; i < mesi.length; i += 30) {
      const snap = await getDocs(query(collection(db, 'rapportini'), where('stato', '==', 'convalidato'), where('mese', 'in', mesi.slice(i, i + 30))));
      snap.docs.forEach(d => trovati.push({ id: d.id, ...d.data() }));
    }
    const fermi = (stato.arrivati || []).filter(r => mesi.includes(r.mese)).length;
    const avvisoFermi = fermi ? `<div class="errore">${fermi} rapportini di quei mesi sono ancora da convalidare: non vengono cancellati.</div>` : '';
    if (!trovati.length) { box.innerHTML = `${avvisoFermi}<div class="nessuno">Nessun mese da cancellare: nell'app ci sono solo gli ultimi ${MESI_DA_TENERE} mesi.</div>`; return; }
    if (!await archivioAccessibile()) { box.innerHTML = '<div class="errore">Prima scegli la cartella dell\'archivio: senza non si può controllare cosa è stato salvato.</div>'; return; }
    const perMese = new Map();
    for (const r of trovati) {
      const salvato = !!r.archivioFile && await CartellaArchivio.esiste(String(r.archivioPercorso).split('/'), r.archivioFile);
      if (!perMese.has(r.mese)) perMese.set(r.mese, { ids: [], mancanti: [] });
      perMese.get(r.mese).ids.push(r.id);
      if (!salvato) perMese.get(r.mese).mancanti.push(r);
    }
    stato.mesiVecchi = perMese;
    box.innerHTML = avvisoFermi + [...perMese.keys()].sort().map(m => {
      const { ids, mancanti } = perMese.get(m);
      return `<div class="voce">
        <div class="voce-princ"><div class="voce-titolo">${nomeMese(m)}: ${ids.length} rapportini</div>
          <div class="voce-sotto">${mancanti.length
            ? `<b>${mancanti.length} non sono nell'archivio</b>: ${mancanti.slice(0, 8).map(r => `${esc(dataIT(r.data))} ${esc(r.commessa)}`).join(', ')}${mancanti.length > 8 ? '…' : ''}. Aprili dai Convalidati e premi "Salva nell'archivio".`
            : 'Tutti presenti nell\'archivio.'}</div></div>
        ${mancanti.length ? '' : `<button class="btn piccolo pericolo" data-cancella-mese="${esc(m)}">Cancella dal database</button>`}
      </div>`;
    }).join('');
  } catch (err) { box.innerHTML = `<div class="errore">${esc(messaggioErrore(err))}</div>`; }
}
$('btnControllaVecchi').onclick = controllaMesiVecchi;
$('mesiVecchi').addEventListener('click', async e => {
  const m = e.target.closest('[data-cancella-mese]')?.dataset.cancellaMese;
  if (!m) return;
  const { ids } = stato.mesiVecchi.get(m);
  if (!await conferma(`Cancellare ${nomeMese(m)}?`, `<p>Verranno cancellati dal database <b>${ids.length}</b> rapportini di ${nomeMese(m)}.</p>
    <p>Sono tutti nell'archivio, ma dall'app <b>non si potranno più vedere né riscrivere su Excel</b>. L'operazione non si può annullare.</p>`, 'Cancella')) return;
  try {
    for (let i = 0; i < ids.length; i += 400) {
      const lotto = writeBatch(db);
      ids.slice(i, i + 400).forEach(id => lotto.delete(doc(db, 'rapportini', id)));
      await lotto.commit();
    }
    avviso(`${nomeMese(m)}: cancellati ${ids.length} rapportini`, 'ok');
    controllaMesiVecchi();
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// ---------------------------------------------------------------- promemoria dello sviluppatore
// Il testo sta nel database (config/promemoria, solo lo sviluppatore), mai nel codice pubblicato.
// Si scrive in modo semplice: "## titolo", "- punto", "1. passo", "| cella | cella |" per le tabelle, **grassetto**.
function promemoriaInHtml(testo) {
  const riga = t => esc(t).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
  const fuori = [];
  let blocco = null, righe = [];
  const chiudi = () => {
    if (blocco === 'ul' || blocco === 'ol') fuori.push(`<${blocco}>${righe.map(x => `<li>${x}</li>`).join('')}</${blocco}>`);
    if (blocco === 'tab') fuori.push(`<div class="tabella-scorri"><table class="tab">${righe.map((c, i) => `<tr>${c.map(x => `<${i ? 'td' : 'th'}>${x}</${i ? 'td' : 'th'}>`).join('')}</tr>`).join('')}</table></div>`);
    blocco = null; righe = [];
  };
  const inBlocco = (tipo, valore) => { if (blocco !== tipo) chiudi(); blocco = tipo; righe.push(valore); };
  for (const l of String(testo || '').split(/\r?\n/)) {
    const t = l.trim();
    if (!t) chiudi();
    else if (t.startsWith('## ')) { chiudi(); fuori.push(`<h3>${riga(t.slice(3))}</h3>`); }
    else if (t.startsWith('- ')) inBlocco('ul', riga(t.slice(2)));
    else if (/^\d+\. /.test(t)) inBlocco('ol', riga(t.replace(/^\d+\. /, '')));
    else if (t.startsWith('|')) inBlocco('tab', t.replace(/^\||\|$/g, '').split('|').map(c => riga(c.trim())));
    else { chiudi(); fuori.push(`<p>${riga(t)}</p>`); }
  }
  chiudi();
  return fuori.join('');
}
let testoPromemoria = '';
async function caricaPromemoria() {
  try {
    const snap = await getDoc(doc(db, 'config', 'promemoria'));
    testoPromemoria = snap.exists() ? (snap.data().testo || '') : '';
    $('testoPromemoria').innerHTML = testoPromemoria ? promemoriaInHtml(testoPromemoria) : '<div class="nessuno">Il promemoria è vuoto: premi Modifica per scriverlo.</div>';
  } catch (err) { $('testoPromemoria').innerHTML = `<div class="errore">${esc(messaggioErrore(err))}</div>`; }
}
$('promemoria').addEventListener('toggle', () => { if ($('promemoria').open) caricaPromemoria(); });
$('btnModificaPromemoria').onclick = async () => {
  const scelta = finestra('Modifica il promemoria', `<textarea id="campoPromemoria" class="campo-promemoria" rows="18">${esc(testoPromemoria)}</textarea>
    <p class="aiuto">Non scrivere PIN o password. Titoli: ## titolo · Punti: - testo · Passi: 1. testo · Tabelle: | cella | cella | · Grassetto: **testo**</p>`,
    [{ testo: 'Annulla', valore: false }, { testo: 'Salva', valore: true, primario: true }]);
  const campo = $('campoPromemoria');
  if (!await scelta) return;
  try {
    await setDoc(doc(db, 'config', 'promemoria'), { testo: campo.value, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
    avviso('Promemoria salvato', 'ok');
    caricaPromemoria();
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
};

// ---------------------------------------------------------------- messaggi per i capisquadra
// Testi pronti da mandare con WhatsApp (un solo messaggio, il destinatario lo sceglie lo sviluppatore).
// Stanno in elenchi/messaggi (scrive solo lo sviluppatore); finché non si salva vale il messaggio di partenza.
const MESSAGGI_BASE = [{
  titolo: 'Rapportino nei giorni di solo viaggio',
  testo: [
    '*Rapportino nei giorni di solo viaggio*',
    '',
    'Quando la squadra fa solo il viaggio e non lavora in cantiere (per esempio il trasferimento a Milano), compilate il rapportino così:',
    '',
    '1. Scegliete la *commessa* e segnate con la *X* gli operai presenti.',
    '2. In *Arrivo in deposito* scrivete l\'ora in cui vi trovate in deposito per partire.',
    '3. In *Rientro in deposito* scrivete l\'ora di arrivo a destinazione, oppure l\'ora di rientro se tornate in giornata.',
    '4. Lasciate *vuoti* "Inizio lavori in cantiere" e "Fine lavori in cantiere".',
    '5. Premete *Applica a tutti i presenti*. *Non* usate il pulsante "Smart time entry".',
    '6. *Pausa:* se vi siete fermati a mangiare, scrivete l\'orario della sosta. Se non vi siete fermati, premete *Nessuna pausa*.',
    '7. In *Lavorazioni eseguite* scrivete: *Solo viaggio – trasferimento a …* (per esempio Milano).',
    '8. Premete *Invia*. Comparirà l\'avviso *"Mancano gli orari"*: è normale, premete *Continua lo stesso*.',
    '',
    'Le ore di viaggio le calcola l\'ufficio.',
    'Per dubbi contattate l\'ufficio.'
  ].join('\n')
}];
let messaggi = [];
function disegnaMessaggi() {
  $('listaMessaggi').innerHTML = messaggi.length ? messaggi.map((m, i) => `<div class="voce messaggio">
      <div class="voce-princ"><div class="voce-titolo">${esc(m.titolo)}</div>
        <details><summary class="voce-sotto">Mostra il testo</summary><div class="testo-messaggio">${esc(m.testo)}</div></details></div>
      <div class="azioni-messaggio">
        <button class="btn piccolo primario" type="button" data-msg-wa="${i}">WhatsApp</button>
        <button class="btn piccolo" type="button" data-msg-modifica="${i}">Modifica</button>
        <button class="btn piccolo pericolo" type="button" data-msg-elimina="${i}">Elimina</button>
      </div></div>`).join('')
    : '<div class="nessuno">Nessun messaggio: premi "+ Nuovo messaggio".</div>';
}
async function caricaMessaggi() {
  try {
    const snap = await getDoc(doc(db, 'elenchi', 'messaggi'));
    messaggi = snap.exists() ? (snap.data().voci || []) : MESSAGGI_BASE.map(m => ({ ...m }));
    disegnaMessaggi();
  } catch (err) { $('listaMessaggi').innerHTML = `<div class="errore">${esc(messaggioErrore(err))}</div>`; }
}
async function salvaMessaggi(nuovi, testoOk) {
  try {
    await setDoc(doc(db, 'elenchi', 'messaggi'), { voci: nuovi, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
    messaggi = nuovi;
    disegnaMessaggi();
    avviso(testoOk, 'ok');
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
}
async function modificaMessaggio(i) {
  const m = i === null ? { titolo: '', testo: '' } : messaggi[i];
  const scelta = finestra(i === null ? 'Nuovo messaggio' : 'Modifica il messaggio', `
    <label class="campo"><span>Titolo (lo vedi solo tu)</span><input id="campoTitoloMsg" maxlength="80" autocomplete="off" value="${esc(m.titolo)}"></label>
    <label class="campo"><span>Testo del messaggio</span><textarea id="campoTestoMsg" class="campo-promemoria" rows="16">${esc(m.testo)}</textarea></label>
    <p class="aiuto">Su WhatsApp le parole tra asterischi si vedono in grassetto: *così*. Non scrivere PIN o password.</p>`,
    [{ testo: 'Annulla', valore: false }, { testo: 'Salva', valore: true, primario: true }]);
  const titolo = $('campoTitoloMsg'), testo = $('campoTestoMsg');
  if (!await scelta) return;
  const voce = { titolo: titolo.value.trim() || 'Messaggio senza titolo', testo: testo.value.trim() };
  if (!voce.testo) { avviso('Il messaggio è vuoto: non è stato salvato', 'errore'); return; }
  const nuovi = [...messaggi];
  if (i === null) nuovi.push(voce); else nuovi[i] = voce;
  salvaMessaggi(nuovi, 'Messaggio salvato');
}
$('messaggi').addEventListener('toggle', () => { if ($('messaggi').open) caricaMessaggi(); });
$('btnNuovoMessaggio').onclick = () => modificaMessaggio(null);
$('listaMessaggi').addEventListener('click', async e => {
  const d = e.target.dataset;
  if (d.msgWa !== undefined) window.open('https://wa.me/?text=' + encodeURIComponent(messaggi[Number(d.msgWa)].testo), '_blank');
  else if (d.msgModifica !== undefined) modificaMessaggio(Number(d.msgModifica));
  else if (d.msgElimina !== undefined) {
    const i = Number(d.msgElimina);
    if (!await conferma('Eliminare il messaggio?', `<p>Il messaggio <b>${esc(messaggi[i].titolo)}</b> verrà eliminato.</p>`, 'Elimina')) return;
    salvaMessaggi(messaggi.filter((_, x) => x !== i), 'Messaggio eliminato');
  }
});

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
  // Viaggio massimo per commessa: assenze, ferie e altre voci nascoste non servono
  const scelte = [[1, '1 ora'], [0.5, 'Mezz\'ora'], [0, 'Nessuno'], [VIAGGIO_VERO, 'Viaggio vero']];
  disegnaGestioneOrarioDeposito();
  $('listaViaggi').innerHTML = e.commesse.filter(c => !e.nascoste.includes(c)).map(c => {
    const max = viaggioMassimo(c);
    return `<label class="riga-viaggio"><span>${esc(c)}</span><select data-viaggio-voce="${esc(c)}">${scelte.map(([v, t]) => `<option value="${v}" ${v === max ? 'selected' : ''}>${t}</option>`).join('')}</select></label>`;
  }).join('');
}

$('listaViaggi').addEventListener('change', async e => {
  const voce = e.target.dataset.viaggioVoce;
  if (voce === undefined) return;
  try { await updateDoc(doc(db, 'elenchi', 'correnti'), { viaggi: { ...stato.elenchi.viaggi, [voce]: e.target.value === VIAGGIO_VERO ? VIAGGIO_VERO : Number(e.target.value) } }); }
  catch (err) { avviso(messaggioErrore(err), 'errore'); mostraInfoElenchi(); }
});

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
    // merge: le scelte sul viaggio per commessa non si toccano
    await setDoc(doc(db, 'elenchi', 'correnti'), { operai, commesse, nascoste, file: master, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome }, { merge: true });
    avviso('Elenchi aggiornati', 'ok');
  } catch (err) { await finestra('Elenchi non aggiornati', `<p>${esc(messaggioErrore(err))}</p>`); }
};

// ---------------------------------------------------------------- intestazione del modulo
const ETICHETTE_INTESTAZIONE = {
  ragioneSociale: 'Ragione sociale', sedeLegale: 'Sede legale', sedeAmministrativa: 'Sede amministrativa',
  datiFiscali: 'P.IVA / C.C.I.A.A. / R.E.A.', telefono: 'Telefono e fax', contatti: 'E-mail, web e PEC',
  titolo: 'Titolo', sottotitolo: 'Sottotitolo', identificazione: 'Identificazione (es. Mod. 22.03)', edizione: 'Edizione',
  revisione: 'Revisione', dataRevisione: 'Data', norma1: 'Norma 1', norma2: 'Norma 2', norma3: 'Norma 3',
  colPartenza: 'Titolo colonna partenza', colCantiere: 'Titolo colonna cantiere', colRientro: 'Titolo colonna rientro'
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
        <span class="voce-num">${i + 1}.</span>
        <div class="voce-princ"><div class="voce-titolo">${esc(m)}</div></div>
        <button class="btn piccolo" data-modifica-mezzo="${i}" title="Modifica" aria-label="Modifica ${esc(m)}">✏️</button>
        <button class="btn piccolo pericolo" data-togli-mezzo="${i}">Togli</button></div>`).join('')
    : '<div class="nessuno">Nessun automezzo: aggiungili qui sotto</div>';
  const n = stato.automezzi.length;
  $('contaMezzi').textContent = n ? `(${n})` : '';
}
const salvaMezzi = voci => setDoc(doc(db, 'elenchi', 'automezzi'), { voci, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });

$('formMezzo').addEventListener('submit', async e => {
  e.preventDefault();
  const nuovo = $('nuovoMezzo').value.trim().replace(/\s+/g, ' ');
  if (!nuovo) return;
  if (stato.automezzi.some(m => m.toUpperCase() === nuovo.toUpperCase())) { avviso('Questo automezzo è già in elenco', 'errore'); return; }
  // il campo si svuota subito: se si sta già scrivendo il mezzo successivo, la conferma del salvataggio non lo cancella
  $('nuovoMezzo').value = '';
  try {
    await salvaMezzi([...stato.automezzi, nuovo]);
    avviso(`Aggiunto: ${nuovo}`, 'ok');
  } catch (err) {
    if (!$('nuovoMezzo').value) $('nuovoMezzo').value = nuovo;
    avviso(messaggioErrore(err), 'errore');
  }
});
$('listaMezzi').addEventListener('click', async e => {
  const matita = e.target.closest('[data-modifica-mezzo]');
  if (matita) {
    const vecchio = stato.automezzi[Number(matita.dataset.modificaMezzo)];
    if (vecchio === undefined) return;
    const scelta = finestra('Modifica automezzo', `<input id="campoMezzo" class="campo-mezzo" maxlength="80" autocomplete="off" value="${esc(vecchio)}">
      <p class="aiuto">I rapportini già inviati non cambiano. Chi lo ha già scelto in un rapportino non ancora inviato dovrà sceglierlo di nuovo.</p>`,
      [{ testo: 'Annulla', valore: false }, { testo: 'Salva', valore: true, primario: true }]);
    const campo = $('campoMezzo');
    campo.focus();
    campo.select();
    campo.onkeydown = ev => { if (ev.key === 'Enter') { ev.preventDefault(); $('finestraBottoni').lastChild.click(); } };
    if (!await scelta) return;
    const nuovo = campo.value.trim().replace(/\s+/g, ' ');
    if (!nuovo || nuovo === vecchio) return;
    // la posizione si ricalcola adesso: nel frattempo l'elenco può essere cambiato
    const pos = stato.automezzi.indexOf(vecchio);
    if (pos < 0) { avviso('Questo automezzo non è più in elenco', 'errore'); return; }
    if (stato.automezzi.some((m, x) => x !== pos && m.toUpperCase() === nuovo.toUpperCase())) { avviso('Questo automezzo è già in elenco', 'errore'); return; }
    try {
      await salvaMezzi(stato.automezzi.map((m, x) => (x === pos ? nuovo : m)));
      avviso(`Modificato: ${nuovo}`, 'ok');
    } catch (err) { avviso(messaggioErrore(err), 'errore'); }
    return;
  }
  const i = e.target.dataset.togliMezzo;
  if (i === undefined) return;
  const m = stato.automezzi[Number(i)];
  if (!await conferma('Togliere l\'automezzo?', `<p><b>${esc(m)}</b> non comparirà più sui telefoni. I rapportini già inviati non cambiano.</p>`, 'Togli')) return;
  try { await salvaMezzi(stato.automezzi.filter(x => x !== m)); }
  catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// ---------------------------------------------------------------- responsabili di commessa
// Elenco in elenchi/responsabili (i nomi stanno solo nel database): uno può essere spuntato come predefinito.
function disegnaGestioneResponsabili() {
  if (stato.utente?.ruolo !== 'responsabile') return;
  const { voci, predefinito } = stato.responsabili;
  $('listaResponsabili').innerHTML = voci.length
    ? voci.map((n, i) => `<div class="voce">
        <span class="voce-num">${i + 1}.</span>
        <div class="voce-princ"><div class="voce-titolo">${esc(n)}</div></div>
        <label class="spunta-predefinito"><input type="checkbox" data-predefinito="${i}" ${n === predefinito ? 'checked' : ''}> Predefinito</label>
        <button class="btn piccolo pericolo" data-togli-responsabile="${i}">Togli</button></div>`).join('')
    : '<div class="nessuno">Nessun responsabile: aggiungili qui sotto</div>';
  $('contaResponsabili').textContent = voci.length ? `(${voci.length})` : '';
}
const salvaResponsabili = (voci, predefinito) => setDoc(doc(db, 'elenchi', 'responsabili'), { voci, predefinito, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });

$('formResponsabile').addEventListener('submit', async e => {
  e.preventDefault();
  const nuovo = $('nuovoResponsabile').value.trim().replace(/\s+/g, ' ');
  if (!nuovo) return;
  const { voci, predefinito } = stato.responsabili;
  if (voci.some(n => n.toUpperCase() === nuovo.toUpperCase())) { avviso('Questo nome è già in elenco', 'errore'); return; }
  $('nuovoResponsabile').value = '';
  try {
    await salvaResponsabili([...voci, nuovo], predefinito);
    avviso(`Aggiunto: ${nuovo}`, 'ok');
  } catch (err) {
    if (!$('nuovoResponsabile').value) $('nuovoResponsabile').value = nuovo;
    avviso(messaggioErrore(err), 'errore');
  }
});
$('listaResponsabili').addEventListener('change', async e => {
  const i = e.target.dataset.predefinito;
  if (i === undefined) return;
  const { voci } = stato.responsabili;
  // la spunta è una sola: mettendola su un nome si toglie dagli altri; togliendola non c'è più un predefinito
  try { await salvaResponsabili(voci, e.target.checked ? voci[Number(i)] : ''); }
  catch (err) { avviso(messaggioErrore(err), 'errore'); disegnaGestioneResponsabili(); }
});
$('listaResponsabili').addEventListener('click', async e => {
  const i = e.target.dataset.togliResponsabile;
  if (i === undefined) return;
  const { voci, predefinito } = stato.responsabili;
  const n = voci[Number(i)];
  if (!await conferma('Togliere il responsabile?', `<p><b>${esc(n)}</b> non comparirà più nell'elenco. I rapportini già inviati non cambiano.</p>`, 'Togli')) return;
  try { await salvaResponsabili(voci.filter(x => x !== n), predefinito === n ? '' : predefinito); }
  catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// ---------------------------------------------------------------- utenti
function disegnaUtenti(utenti) {
  stato.utentiLista = utenti;
  utenti.sort((a, b) => (a.ruolo || '').localeCompare(b.ruolo || '') || (a.nome || '').localeCompare(b.nome || ''));
  $('listaUtenti').innerHTML = utenti.map(u => `
    <div class="voce">
      <div class="voce-princ">
        <div class="voce-titolo">${esc(u.nome)}</div>
        <div class="voce-sotto">Accesso: ${esc(u.codice)} · ${u.ruolo === 'responsabile' ? (u.sviluppatore ? 'Responsabile · Sviluppatore' : 'Responsabile') : 'Caposquadra'}</div>
      </div>
      <span class="etichetta ${u.attivo ? 'ok' : 'no'}">${u.attivo ? 'Attivo' : 'Disattivato'}</span>
      ${u.uid === stato.utente.uid ? '' : `<button class="btn piccolo" data-uid="${esc(u.uid)}" data-attivo="${u.attivo ? 1 : 0}">${u.attivo ? 'Disattiva' : 'Riattiva'}</button>
      <button class="btn piccolo pericolo" data-elimina="${esc(u.uid)}" data-nome="${esc(u.nome)}" data-codice="${esc(u.codice)}">Elimina</button>`}
      ${rigaPin(u)}
    </div>`).join('');
}
// Riga PIN: PIN nascosto con occhio, invio con WhatsApp, nuovo PIN (scritto o generato) e Cambia
function rigaPin(u) {
  const pin = (stato.pinUtenti || {})[u.uid];
  const id = esc(u.uid);
  return `<div class="voce-pin">
    <span class="pin-attuale" data-pin-di="${id}">${pin ? '••••••' : '<i>PIN non registrato</i>'}</span>
    ${pin ? `<button class="btn piccolo" data-occhio="${id}" title="Mostra / nascondi il PIN">👁</button>` : ''}
    <button class="btn piccolo whatsapp" data-whatsapp="${id}">WhatsApp</button>
    <input class="pin-nuovo" data-pin-campo="${id}" inputmode="numeric" maxlength="6" placeholder="6 cifre" autocomplete="off">
    <button class="btn piccolo" data-genera="${id}">🎲 Genera</button>
    <button class="btn piccolo" data-cambia="${id}">Cambia</button>
  </div>`;
}
// Messaggio WhatsApp: link per installare l'app, istruzioni e dati di accesso (si sceglie il contatto in WhatsApp)
// Data da cui si accettano solo rapportini digitali (scritta nel messaggio di benvenuto)
const INIZIO_DIGITALE = 'lunedì 9 novembre 2026 (09/11/2026)';
function messaggioInstallazione(nome, pin) {
  // il link è sempre quello ufficiale (anche se il messaggio parte dal vecchio indirizzo); nel simulatore quello locale
  const link = EMULATORE ? location.origin + location.pathname : 'https://rapportino.gccodelab.it/';
  return [
    `📋 *BENVENUTO NEL RAPPORTINO DI CANTIERE DIGITALE*`,
    `Vi.Cla. Future S.r.L.`,
    ``,
    `Ciao *${nome}*,`,
    ``,
    `📢 *AVVISO IMPORTANTE*`,
    `Da *${INIZIO_DIGITALE}* i rapportini si consegnano *SOLO IN DIGITALE*, con questa applicazione.`,
    `*Non saranno più accettati rapportini cartacei.*`,
    ``,
    `📝 *COM'È FATTO*`,
    `Il rapportino è *uguale a quello di carta*: commessa, data, operai con la X, orari, automezzi, pausa mensa, lavorazioni.`,
    `Solo che è *più semplice e veloce* da compilare.`,
    ``,
    `📲 *COME INSTALLARLA*`,
    `Apri questo link:`,
    link,
    ``,
    `▪️ *Android* (Chrome): tocca i tre puntini ⋮ in alto a destra e scegli *"Installa app"* (oppure "Aggiungi a schermata Home").`,
    `▪️ *iPhone* (Safari): tocca il tasto *Condividi* (il quadrato con la freccia) e scegli *"Aggiungi a Home"*.`,
    ``,
    `Sul telefono comparirà l'icona *Rapportino*.`,
    ``,
    `🔑 *I TUOI DATI PER ENTRARE*`,
    `Nome e cognome: *${nome}*`,
    `PIN: *${pin}*`,
    `_Non dare il PIN a nessuno._`,
    ``,
    `✏️ *PRIMA DI TUTTO: FAI PRATICA*`,
    `Dopo l'installazione entra nell'app e *prova a compilare i campi*: scegli la commessa, metti le X agli operai, scrivi gli orari, gli automezzi e le lavorazioni.`,
    ``,
    `⏱️ *ORARI UGUALI PER TUTTA LA SQUADRA*`,
    `Se avete fatto tutti gli stessi orari non serve scriverli operaio per operaio:`,
    `1️⃣ metti la X a tutti gli operai presenti (il riquadro compare quando ne hai segnati almeno 2);`,
    `2️⃣ nel riquadro *"Stessi orari per tutti i presenti"* scrivi partenza dal deposito, Dalle, Alle e rientro in deposito;`,
    `3️⃣ premi *"Applica a tutti i presenti"*: gli orari compaiono nella riga di ogni operaio con la X.`,
    `Se un operaio ha fatto un orario diverso, correggilo poi nella sua riga.`,
    ``,
    `⚠️ *ATTENZIONE: MENTRE FAI PRATICA NON PREMERE "INVIA AL RESPONSABILE"*`,
    `Quando hai finito di provare, premi *"Svuota"* per cancellare tutto.`,
    ``,
    `✅ *DAL 09/11/2026*`,
    `Compila il rapportino del giorno e premi *"Invia al responsabile"*: arriva subito in ufficio.`,
    `Controllalo bene prima di inviarlo: *dopo l'invio non si può più vedere né modificare.*`,
    ``,
    `❓ Per dubbi o problemi rivolgiti all'ufficio.`,
    ``,
    `_GC-CodeLab • Software Engineered & Designed by Giovanni Cavallo_`
  ].join('\n');
}
// Cambio PIN di un utente: ci si collega come lui con una connessione "usa e getta" (serve il PIN attuale)
// e si imposta quello nuovo. Per sé stessi si usa la sessione aperta.
async function cambiaPin(u, nuovo) {
  let attuale = (stato.pinUtenti || {})[u.uid];
  if (!attuale) {
    attuale = await chiediPin(`PIN attuale di ${u.nome}`, 'Il PIN attuale di questo utente non è registrato nell\'app. Scrivilo una volta (per i vecchi accessi: la vecchia password), poi l\'app lo ricorderà.');
    if (attuale === null) return false;
  }
  if (u.uid === stato.utente.uid) {
    await reauthenticateWithCredential(auth.currentUser, EmailAuthProvider.credential(emailDi(u.codice), attuale));
    await updatePassword(auth.currentUser, nuovo);
  } else {
    let app2;
    try {
      app2 = initializeApp(config, `cambio-pin-${Date.now()}`);
      const auth2 = initializeAuth(app2, { persistence: inMemoryPersistence });
      if (EMULATORE) connectAuthEmulator(auth2, 'http://127.0.0.1:9099', { disableWarnings: true });
      const cred = await signInWithEmailAndPassword(auth2, emailDi(u.codice), attuale);
      await updatePassword(cred.user, nuovo);
      await signOut(auth2);
    } finally {
      if (app2) deleteApp(app2).catch(() => {});
    }
  }
  await setDoc(doc(db, 'accessi', u.uid), { pin: nuovo, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
  stato.pinUtenti = { ...(stato.pinUtenti || {}), [u.uid]: nuovo };
  return true;
}
$('listaUtenti').addEventListener('click', async e => {
  const t = e.target.closest('[data-occhio],[data-whatsapp],[data-genera],[data-cambia]');
  if (!t) return;
  const uid = t.dataset.occhio || t.dataset.whatsapp || t.dataset.genera || t.dataset.cambia;
  const u = (stato.utentiLista || []).find(x => x.uid === uid);
  if (!u) return;
  const pin = (stato.pinUtenti || {})[uid];
  const campo = $('listaUtenti').querySelector(`[data-pin-campo="${CSS.escape(uid)}"]`);
  if (t.dataset.occhio) {
    const el = $('listaUtenti').querySelector(`[data-pin-di="${CSS.escape(uid)}"]`);
    el.textContent = el.textContent === pin ? '••••••' : pin;
  } else if (t.dataset.whatsapp) {
    if (!pin) { avviso('Prima imposta un PIN con "Genera" e "Cambia": poi lo puoi inviare', 'errore'); return; }
    window.open('https://wa.me/?text=' + encodeURIComponent(messaggioInstallazione(u.nome, pin)), '_blank');
  } else if (t.dataset.genera) {
    campo.value = String(crypto.getRandomValues(new Uint32Array(1))[0] % 1000000).padStart(6, '0');
  } else if (t.dataset.cambia) {
    const nuovo = campo.value.trim();
    if (!/^\d{6}$/.test(nuovo)) { avviso('Il nuovo PIN deve essere di 6 cifre (scrivilo o premi Genera)', 'errore'); return; }
    if (!await conferma('Cambiare il PIN?', `<p>Nuovo PIN di <b>${esc(u.nome)}</b>: <b>${nuovo}</b></p><p>Il vecchio PIN non funzionerà più. Ricordati di comunicarglielo (anche con il tasto WhatsApp).</p>`, 'Cambia')) return;
    t.disabled = true;
    try {
      if (await cambiaPin(u, nuovo)) { avviso(`PIN di ${u.nome} cambiato`, 'ok'); disegnaUtenti(stato.utentiLista); }
    } catch (err) {
      avviso(/invalid-credential|wrong-password/.test(err?.code || '') ? 'PIN attuale sbagliato: il PIN non è stato cambiato' : messaggioErrore(err), 'errore');
    } finally { t.disabled = false; }
  }
});
$('listaUtenti').addEventListener('click', async e => {
  const el = e.target.closest('[data-elimina]');
  if (el) {
    const ok = await conferma('Eliminare l\'utente?',
      `<p><b>${esc(el.dataset.nome)}</b> sparirà dall'elenco e non potrà più entrare nell'app.</p>
       <p>I rapportini che ha già inviato restano, con il suo nome.</p>
       <p>Lo stesso nome non si potrà riusare per un utente nuovo (aggiungi la seconda iniziale o un numero): se serve solo bloccarlo per un periodo, usa "Disattiva".</p>`, 'Elimina');
    if (!ok) return;
    try {
      await deleteDoc(doc(db, 'utenti', el.dataset.elimina));
      await deleteDoc(doc(db, 'accessi', el.dataset.elimina)).catch(() => {});
      // l'elenco si aggiorna subito, senza aspettare il database (con più schede aperte a volte tarda)
      disegnaUtenti((stato.utentiLista || []).filter(u => u.uid !== el.dataset.elimina));
      avviso('Utente eliminato', 'ok');
    } catch (err) { avviso(messaggioErrore(err), 'errore'); }
    return;
  }
  const b = e.target.closest('[data-uid]');
  if (!b) return;
  const attiva = b.dataset.attivo !== '1';
  try {
    await updateDoc(doc(db, 'utenti', b.dataset.uid), { attivo: attiva });
    disegnaUtenti((stato.utentiLista || []).map(u => (u.uid === b.dataset.uid ? { ...u, attivo: attiva } : u)));
  } catch (err) { avviso(messaggioErrore(err), 'errore'); }
});

// Il nuovo account si crea con una seconda connessione "usa e getta", così il responsabile resta collegato
$('formUtente').addEventListener('submit', async e => {
  e.preventDefault();
  const nome = $('nuNome').value.trim();
  const codice = normCodice(nome);
  const password = $('nuPassword').value.trim();
  const ruolo = $('nuRuolo').value;
  if (!codiceValido(codice)) { avviso('Scrivi nome e cognome (almeno 2 lettere)', 'errore'); return; }
  if (!/^\d{6}$/.test(password)) { avviso('Il PIN deve essere di 6 cifre', 'errore'); return; }
  $('btnCreaUtente').disabled = true;
  let app2;
  try {
    app2 = initializeApp(config, `nuovo-utente-${Date.now()}`);
    const auth2 = initializeAuth(app2, { persistence: inMemoryPersistence });
    if (EMULATORE) connectAuthEmulator(auth2, 'http://127.0.0.1:9099', { disableWarnings: true });
    const cred = await createUserWithEmailAndPassword(auth2, emailDi(codice), password);
    await setDoc(doc(db, 'utenti', cred.user.uid), { nome, codice, ruolo, attivo: true, creatoIl: serverTimestamp(), creatoDa: stato.utente.nome });
    await setDoc(doc(db, 'accessi', cred.user.uid), { pin: password, aggiornatoIl: serverTimestamp(), aggiornatoDa: stato.utente.nome });
    await signOut(auth2);
    $('formUtente').reset();
    await finestra('Utente creato', `<p><b>${esc(nome)}</b> può entrare con:</p><p>Nome e cognome: <b>${esc(nome)}</b><br>PIN: <b>${esc(password)}</b></p><p>Comunicaglieli di persona.</p>`);
  } catch (err) {
    avviso(messaggioErrore(err), 'errore');
  } finally {
    if (app2) deleteApp(app2).catch(() => {});
    $('btnCreaUtente').disabled = false;
  }
});

// ---------------------------------------------------------------- area sviluppatore (ingresso nascosto)
// 5 tocchi sul simbolo GC in fondo alla pagina + PIN. Per tutti gli altri il simbolo non fa nulla.
// Il PIN è salvato come impronta (SHA-256 con sale) in config/sviluppatore, leggibile solo dallo sviluppatore.
// La vera protezione resta nelle regole del database (solo il profilo sviluppatore può modificare elenchi e utenti).
let tocchiMarchio = [], pinSbagliati = 0, pinBloccatoFino = 0, ingressoSviluppatore = false;
async function impronta(pin, sale) {
  const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(`${sale}:${pin}`));
  return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join('');
}
async function chiediPin(titolo, testo) {
  const scelta = finestra(titolo, `<p>${testo}</p><input id="campoPin" class="campo-pin" type="password" inputmode="numeric" autocomplete="off" maxlength="12">`,
    [{ testo: 'Annulla', valore: false }, { testo: 'Conferma', valore: true, primario: true }]);
  const campo = $('campoPin');
  campo.focus();
  campo.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); $('finestraBottoni').lastChild.click(); } };
  return (await scelta) ? campo.value.trim() : null;
}
async function nuovoPin(primaVolta) {
  const p1 = await chiediPin(primaVolta ? 'Crea il PIN' : 'Nuovo PIN',
    `${primaVolta ? 'Primo ingresso nell\'area sviluppatore. ' : ''}Scegli un PIN da 4 a 12 cifre.`);
  if (p1 === null) return false;
  if (!/^\d{4,12}$/.test(p1)) { avviso('Il PIN deve avere da 4 a 12 cifre', 'errore'); return false; }
  const p2 = await chiediPin('Ripeti il PIN', 'Scrivi di nuovo lo stesso PIN.');
  if (p2 === null) return false;
  if (p1 !== p2) { avviso('I due PIN non sono uguali: riprova', 'errore'); return false; }
  const sale = [...crypto.getRandomValues(new Uint8Array(16))].map(x => x.toString(16).padStart(2, '0')).join('');
  await setDoc(doc(db, 'config', 'sviluppatore'), { pinImpronta: await impronta(p1, sale), sale, impostatoIl: serverTimestamp() });
  avviso('PIN salvato', 'ok');
  return true;
}
async function apriAreaSviluppatore() {
  if (!stato.utente?.sviluppatore) return;
  if (stato.areaAperta) { mostraVista('gestione'); return; }
  if (Date.now() < pinBloccatoFino) { avviso('Troppi PIN errati: riprova tra qualche minuto', 'errore'); return; }
  try {
    const snap = await getDoc(doc(db, 'config', 'sviluppatore'));
    const cfg = snap.exists() ? snap.data() : null;
    if (!cfg?.pinImpronta) {
      if (!await nuovoPin(true)) return;
    } else {
      const pin = await chiediPin('Area sviluppatore', 'Inserisci il PIN.');
      if (pin === null) return;
      if (await impronta(pin, cfg.sale) !== cfg.pinImpronta) {
        if (++pinSbagliati >= 5) { pinSbagliati = 0; pinBloccatoFino = Date.now() + 5 * 60 * 1000; avviso('PIN errato 5 volte: riprova tra 5 minuti', 'errore'); }
        else avviso('PIN errato', 'errore');
        return;
      }
    }
  } catch (err) { avviso(messaggioErrore(err), 'errore'); return; }
  pinSbagliati = 0;
  stato.areaAperta = true;
  document.querySelector('.scheda-btn[data-vista="gestione"]').hidden = false;
  mostraVista('gestione');
}
const cinqueTocchi = azione => () => {
  const ora = Date.now();
  tocchiMarchio = tocchiMarchio.filter(t => ora - t < 4000).concat(ora);
  if (tocchiMarchio.length >= 5) { tocchiMarchio = []; azione(); }
};
$('marchioGC').addEventListener('click', cinqueTocchi(apriAreaSviluppatore));
// Schermata di accesso: 5 tocchi segnano "accesso sviluppatore" (si vede nel sottotitolo); altri 5 lo annullano.
// Si entra con codice e password, poi l'app chiede il PIN (solo se il profilo è quello dello sviluppatore).
$('marchioAccesso').addEventListener('click', cinqueTocchi(() => {
  ingressoSviluppatore = !ingressoSviluppatore;
  $('sottotitoloAccesso').textContent = ingressoSviluppatore ? 'Accesso sviluppatore' : 'Mod. 22.03 digitale';
  if (ingressoSviluppatore) $('loginCodice').focus();
}));
$('btnCambiaPin').onclick = () => nuovoPin(false).catch(err => avviso(messaggioErrore(err), 'errore'));
$('btnChiudiArea').onclick = () => {
  stato.areaAperta = false;
  document.querySelector('.scheda-btn[data-vista="gestione"]').hidden = true;
  mostraVista('arrivati');
};

// Solo per le prove automatiche con il simulatore: cartella finta e lettura diretta (per verificare le regole)
if (EMULATORE) {
  window.__prova = {
    usaCartella: async handle => { await CartellaExcel.useHandle(handle); cartellaPronta(); },
    leggiPromemoria: async () => getDoc(doc(db, 'config', 'promemoria')).then(() => 'letto', e => e.code),
    usaArchivio: async handle => { await CartellaArchivio.useHandle(handle); archivioPronto(); },
    leggiRapportini: async () => (await getDocs(collection(db, 'rapportini'))).size,
    modificaRapportino: async (id, campi) => updateDoc(doc(db, 'rapportini', id), campi),
    scriviElenchi: async () => setDoc(doc(db, 'elenchi', 'automezzi'), { voci: ['PROVA'] }).then(() => 'scritto', e => e.code),
    modificaUtente: async uid => updateDoc(doc(db, 'utenti', uid), { attivo: false }).then(() => 'scritto', e => e.code),
    leggiPin: async () => getDocs(collection(db, 'accessi')).then(s => 'letti ' + s.size, e => e.code),
    testata: (pagina, t) => testataModulo(pagina, t),
    oreProposte: r => bloccoIniziale(r).ore.t1,
    blocchiProposti: r => blocchiIniziali(r)
  };
}
