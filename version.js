// Versione dell'app: a OGNI modifica si aumenta il numero e si aggiunge in cima allo STORICO cosa è cambiato.
// La usano la scritta in fondo alla pagina (con l'elenco "Novità") e il service worker (sw.js)
// per scaricare la versione nuova.
self.APP_STORICO = [
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
