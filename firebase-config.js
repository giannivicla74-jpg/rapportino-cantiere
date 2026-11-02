// ============================================================
// COLLEGAMENTO AL PROGETTO FIREBASE
// Questi valori si copiano dalla console di Firebase (Impostazioni progetto > Le tue app > Config).
// NON sono una password: identificano solo il progetto. La protezione dei dati la fanno
// le regole di sicurezza (file firestore.rules) e gli accessi con codice e password.
// Finché qui c'è "DA-COMPILARE" l'app mostra la schermata "Configurazione mancante".
// ============================================================

export const firebaseConfig = {
  apiKey: 'DA-COMPILARE',
  authDomain: 'DA-COMPILARE',
  projectId: 'DA-COMPILARE',
  storageBucket: 'DA-COMPILARE',
  messagingSenderId: 'DA-COMPILARE',
  appId: 'DA-COMPILARE'
};

// Il codice utente diventa un indirizzo "finto" per l'accesso di Firebase (es. rossi -> rossi@rapportino.invalid).
// Nessuna email viene mai inviata: ".invalid" è un dominio riservato che non esiste.
export const DOMINIO_ACCESSI = 'rapportino.invalid';
