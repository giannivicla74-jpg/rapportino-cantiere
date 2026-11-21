// ============================================================
// CARTELLA "ARCHIVIO RAPPORTINI" (File System Access API di Edge/Chrome)
// L'utente la sceglie una volta, separata da quella dei file Excel. Dentro l'app crea
// una cartella per anno e una per mese, con un file per ogni rapportino convalidato.
// ============================================================

const CartellaArchivio = (() => {
  const DB_NAME = 'rapportino-cartella-excel'; // stesso archivio del browser della cartella Excel, voce diversa
  const STORE = 'cartelle';
  const KEY = 'archivio';
  let dirHandle = null;
  let usaPonte = false;
  let ponteNome = '';

  function idb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function saveHandle(handle) {
    const db = await idb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(handle, KEY);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  }

  async function loadHandle() {
    const db = await idb();
    return new Promise((resolve, reject) => {
      const req = db.transaction(STORE).objectStore(STORE).get(KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }

  function viaPonte() {
    return usaPonte;
  }

  // 'pronta' (accesso già concesso o Ponte PC attivo), 'da-confermare' (serve un clic) o 'mancante'
  async function restore(forzaPonte = false) {
    if (typeof PonteLocale !== 'undefined') {
      const p = await PonteLocale.controlla(forzaPonte);
      if (p && p.ok && p.archivioPronta) {
        usaPonte = true;
        ponteNome = p.archivioNome || 'Archivio';
        return 'pronta';
      }
    }
    usaPonte = false;
    ponteNome = '';
    try {
      const handle = await loadHandle();
      if (!handle) return 'mancante';
      dirHandle = handle;
      return (await handle.queryPermission({ mode: 'readwrite' })) === 'granted' ? 'pronta' : 'da-confermare';
    } catch (_) {
      return 'mancante';
    }
  }

  // Da chiamare dopo un clic dell'utente
  async function confirmAccess() {
    if (usaPonte) return true;
    if (!dirHandle) return false;
    return (await dirHandle.requestPermission({ mode: 'readwrite' })) === 'granted';
  }

  async function choose() {
    const handle = await window.showDirectoryPicker({ id: 'archivio-rapportini', mode: 'readwrite' });
    usaPonte = false;
    ponteNome = '';
    dirHandle = handle;
    await saveHandle(handle);
    return handle.name;
  }

  // Solo per le prove automatiche
  async function useHandle(handle) {
    usaPonte = false;
    ponteNome = '';
    dirHandle = handle;
  }

  const name = () => (usaPonte ? ponteNome : (dirHandle ? dirHandle.name : ''));

  function requireDir() {
    if (!dirHandle) throw new Error('cartella dell\'archivio non collegata');
    return dirHandle;
  }

  async function cartellaDi(percorso, crea) {
    let dir = requireDir();
    for (const parte of percorso) dir = await dir.getDirectoryHandle(parte, { create: crea });
    return dir;
  }

  // Scrive il file e lo rilegge: se il contenuto non torna segnala l'errore
  async function salva(percorso, nomeFile, testo) {
    if (usaPonte) return PonteLocale.archivioSalva(percorso, nomeFile, testo);
    const dir = await cartellaDi(percorso, true);
    const fh = await dir.getFileHandle(nomeFile, { create: true });
    const w = await fh.createWritable();
    await w.write(testo);
    await w.close();
    const riletto = await (await fh.getFile()).text();
    if (riletto !== testo) throw new Error('il file salvato non corrisponde: riprova');
  }

  async function esiste(percorso, nomeFile) {
    if (usaPonte) return PonteLocale.archivioEsiste(percorso, nomeFile);
    try {
      const dir = await cartellaDi(percorso, false);
      const f = await (await dir.getFileHandle(nomeFile)).getFile();
      return f.size > 0;
    } catch (_) {
      return false;
    }
  }

  return { viaPonte, restore, confirmAccess, choose, useHandle, name, salva, esiste };
})();
