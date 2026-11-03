// Versione dell'app: a OGNI modifica si aumenta il numero e si aggiunge in cima allo STORICO cosa è cambiato.
// La usano la scritta in fondo alla pagina (con l'elenco "Novità") e il service worker (sw.js)
// per scaricare la versione nuova.
self.APP_STORICO = [
  { versione: '3.1', data: '03/11/2026', novita: [
    '\"Elenchi e utenti\" non compare più nella barra in alto',
    'Ingresso nascosto per lo sviluppatore: 5 tocchi sul simbolo GC in fondo alla pagina, poi il PIN',
    'Il PIN si crea al primo ingresso e si cambia dall\'area stessa; 5 PIN sbagliati bloccano per 5 minuti',
    'L\'area si richiude uscendo dall\'app, ricaricando la pagina o con \"Chiudi area sviluppatore\"'
  ] },
  { versione: '3.0', data: '03/11/2026', novita: [
    'Tema scuro: nella pagina 2 i messaggi rossi e gialli (es. \"Senza ore\") e i tasti ora si leggono bene',
    'Dentro il modulo bianco si usano sempre i colori del foglio chiaro'
  ] },
  { versione: '2.9', data: '03/11/2026', novita: [
    '\"Elenchi e utenti\" lo vede e lo usa solo lo sviluppatore (Giovanni Cavallo)',
    'Gli altri responsabili ricevono, completano, convalidano e scrivono su Excel come prima',
    'La cartella Excel gli altri responsabili la collegano dalla barra gialla in alto',
    'Il blocco vale anche nel database: senza permesso non si possono cambiare elenchi, automezzi, intestazione e utenti'
  ] },
  { versione: '2.8', data: '03/11/2026', novita: [
    'Schermata di accesso: tolto il collegamento \"Prima configurazione\" (nessuno lo vede più)',
    'La prima configurazione era già bloccata dal database dopo la prima volta; ora non compare nemmeno'
  ] },
  { versione: '2.7', data: '03/11/2026', novita: [
    'Elenchi e utenti: nuovo tasto \"Elimina\" accanto a Disattiva / Riattiva',
    'L\'utente eliminato sparisce dall\'elenco e non può più entrare; i rapportini che ha inviato restano',
    'Il codice di un utente eliminato non si può riusare per uno nuovo',
    'Non si può eliminare sé stessi'
  ] },
  { versione: '2.6', data: '03/11/2026', novita: [
    'Telefono in verticale: l\'intestazione \"COGNOME E NOME\" della tabella si legge intera (prima era tagliata)'
  ] },
  { versione: '2.5', data: '03/11/2026', novita: [
    'Telefono: nessun blocco della rotazione, l\'app segue il telefono',
    'In verticale la tabella operai è come sul foglio (tutto su una riga) e si scorre a destra e a sinistra con il dito',
    'Mentre si scorre, il nome dell\'operaio e la casella P restano fermi a sinistra',
    'Tolto l\'avviso \"Gira il telefono\"'
  ] },
  { versione: '2.4', data: '03/11/2026', novita: [
    'Firma in fondo alla pagina: \"GC-CodeLab • Software Engineered & Designed by Giovanni Cavallo\"'
  ] },
  { versione: '2.3', data: '03/11/2026', novita: [
    'Tasto per il tema in alto (◐ automatico, ☀ chiaro, 🌙 scuro), anche nella schermata di accesso',
    'Con il tema scuro lo sfondo è scuro e il modulo resta un foglio bianco: stanca meno gli occhi',
    'La scelta resta salvata su quel telefono o PC'
  ] },
  { versione: '2.2', data: '03/11/2026', novita: [
    'Brand GC-CodeLab: marchio, motto \"Development | Innovation | Solution\" e \"Sviluppato da Giovanni Cavallo\"',
    'La firma compare nella schermata di accesso, in fondo alle pagine e nelle Novità',
    'Il modulo Vi.Cla. Future e la stampa restano con il solo marchio dell\'azienda'
  ] },
  { versione: '2.1', data: '03/11/2026', novita: [
    'Sul telefono l\'app installata si apre in orizzontale: tabella operai completa come sul foglio',
    'Nel browser, con il telefono in verticale, compare l\'avviso \"Gira il telefono\" (si può chiudere)',
    'In orizzontale la barra in alto è più sottile e scorre via con la pagina'
  ] },
  { versione: '2.0', data: '03/11/2026', novita: [
    'Quando compila l\'ufficio, il caposquadra si sceglie dall\'elenco operai (prima usciva sempre il nome di chi era entrato)',
    'Sul rapportino resta scritto \"compilato dall\'ufficio\" con il nome di chi l\'ha compilato',
    'Per i capisquadra non cambia nulla: il loro nome resta automatico e non modificabile'
  ] },
  { versione: '1.9', data: '03/11/2026', novita: [
    'Tasto \"Stampa\" nel rapportino aperto: fronte e retro su due fogli A4 come il modulo cartaceo',
    'Retro con tabella Turno di lavoro, riquadro ore 1°/2° turno, operai e Firma RC / Convalida',
    'Dalla finestra di stampa si può salvare anche in PDF'
  ] },
  { versione: '1.8', data: '03/11/2026', novita: [
    'Convalidati organizzati per anno, mese e settimana (lunedì-domenica)',
    'Ogni rapportino va nel periodo della sua data; per ogni mese e settimana si vede quanti sono',
    'I mesi si aprono e chiudono con un clic; il più recente parte aperto'
  ] },
  { versione: '1.7', data: '03/11/2026', novita: [
    'Elenco Convalidati: spunta verde \"Scritto in Excel\" con file e giorno (bordo verde a sinistra)',
    'Se un rapportino non risulta scritto compare \"Non scritto in Excel\" in rosso',
    'Nel conteggio degli operai si tiene conto delle correzioni dell\'ufficio'
  ] },
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
