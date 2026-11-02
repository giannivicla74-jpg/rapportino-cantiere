# Rapportino di Cantiere (Mod. 22.03 digitale)

App web installabile (PWA):

- il **caposquadra** compila il rapportino dal telefono e lo invia;
- il **responsabile** lo riceve, scrive le ore e convalida, e l'app scrive presenze e ore nel file Excel delle presenze.

Come funziona:

- Dati e accessi: Firebase (Authentication + Firestore). Le regole di sicurezza stanno in `firebase/firestore.rules` nel repository di sviluppo.
- Operai e commesse si leggono dal file Excel dell'azienda e non sono scritti nel codice.
- I file Excel si scrivono direttamente dal browser (Edge o Chrome sul PC) con la File System Access API.
