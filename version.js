// Versione dell'app: a OGNI modifica si aumenta il numero e si aggiunge in cima allo STORICO cosa è cambiato.
// La usano la scritta in fondo alla pagina (con l'elenco "Novità") e il service worker (sw.js)
// per scaricare la versione nuova.
self.APP_STORICO = [
  { versione: '1.6', data: '03/11/2026', novita: [
    'Sotto \"Convalida\" si vede sempre in quale cartella scriverà l\'app e quali file Excel contiene',
    'Prima di ogni scrittura l\'app rilegge la cartella collegata (se è stata cambiata usa quella nuova)',
    'Copia di sicurezza giornaliera controllata dentro la cartella Backup: niente confusione tra cartelle con lo stesso nome'
  ] },
  { versione: '1.5', data: '03/11/2026', novita: [
    'Pagina 1 del responsabile: con un clic si vedono tutti gli operai dell\'elenco (come sul modulo cartaceo) o solo i presenti',
    'Nella vista completa, accanto a chi non ha la X c\'è \"aggiungi\"; la scelta della vista resta per i rapportini successivi'
  ] },
  { versione: '1.4', data: '03/11/2026', novita: [
    'Aggiornamenti più rapidi sul telefono: l\'app controlla la versione nuova ogni volta che torna in primo piano',
    'Quando c\'è una versione nuova compare la barra \"Aggiorna\"'
  ] },
  { versione: '1.3', data: '03/11/2026', novita: [
    'Il responsabile corregge gli errori del caposquadra nella pagina 1: togli (riga barrata) o rimetti un operaio',
    'Operaio dimenticato: si aggiunge anche dalla pagina 1 e compare in fondo alla tabella',
    'Nella riga delle firme resta la traccia: chi ha corretto, quando e cosa'
  ] },
  { versione: '1.2', data: '03/11/2026', novita: [
    'Dopo la convalida il messaggio dice dove sono stati scritti i dati: cartella, file, foglio, giorno e righe',
    'Si può aprire l\'elenco con la riga Excel di ogni operaio',
    'Le stesse informazioni restano visibili riaprendo un rapportino convalidato'
  ] },
  { versione: '1.1', data: '03/11/2026', novita: [
    'Pagina 2: box "Operai di questo blocco" nello spazio accanto al riquadro ore',
    'Il responsabile può aggiungere operai non segnati dal caposquadra (bordo tratteggiato, si tolgono con la ✕)',
    'Gli operai aggiunti vengono salvati con il rapportino e scritti su Excel come gli altri'
  ] },
  { versione: '1.0', data: '03/11/2026', novita: [
    'Prima versione online: collegata al database vero (Firebase, dati in Europa a Milano)',
    'Indirizzo: https://giannivicla74-jpg.github.io/rapportino-cantiere/'
  ] },
  { versione: '0.9', data: '03/11/2026', novita: [
    'Pagina 2 senza la riga di identificazione del modulo (è già nella pagina 1 accanto)'
  ] },
  { versione: '0.8', data: '03/11/2026', novita: [
    'Sul PC il rapportino aperto mostra pagina 1 e pagina 2 una accanto all\'altra',
    'Ogni pagina scorre per conto suo: si scrivono le ore guardando presenze e orari'
  ] },
  { versione: '0.7', data: '03/11/2026', novita: [
    'Logo Vi.Cla. Future originale, nitido, al posto di quello ritagliato dalla scansione'
  ] },
  { versione: '0.6', data: '03/11/2026', novita: [
    'Intestazione del modulo modificabile dall\'ufficio: sedi, telefono, sito, titolo, revisione, data e norme',
    'Si può cambiare il logo (e tornare a quello originale)',
    'Ogni rapportino conserva l\'intestazione con cui è stato compilato'
  ] },
  { versione: '0.5', data: '03/11/2026', novita: [
    'Numero di versione sempre visibile in fondo alla pagina',
    'Elenco "Novità" con le modifiche di ogni versione'
  ] },
  { versione: '0.4', data: '03/11/2026', novita: [
    'Automezzi utilizzati: elenco a comparsa, si possono scegliere più mezzi e toglierli con la ✕'
  ] },
  { versione: '0.3', data: '03/11/2026', novita: [
    'Automezzi utilizzati scelti da un elenco',
    'Il responsabile aggiunge e toglie gli automezzi in "Elenchi e utenti"'
  ] },
  { versione: '0.2', data: '03/11/2026', novita: [
    'Pagina 1 con la grafica del modulo cartaceo Mod. 22.03: logo Vi.Cla. Future, intestazione, riquadri e tabella operai',
    'Pagina 2 come il retro: turno di lavoro sopra, riquadro ore sotto'
  ] },
  { versione: '0.1', data: '03/11/2026', novita: [
    'Prima versione: il caposquadra compila e invia, poi non rivede più il rapportino',
    'Il responsabile riceve i rapportini, scrive le ore, convalida e scrive su Excel',
    'Accesso con codice e password; utenti creati dal responsabile',
    'Operai e commesse presi dal file Master_Operai; sul telefono solo le commesse vere',
    'Funziona anche senza campo: il rapportino parte appena torna la connessione'
  ] }
];
self.APP_VERSION = self.APP_STORICO[0].versione;
self.APP_VERSION_DATE = self.APP_STORICO[0].data;
