// ============================================================
// CARTELLA DEI FILE EXCEL (File System Access API di Edge/Chrome)
// L'utente sceglie una volta la cartella "Cartella di lavoro con attivazione macro di Excel":
// il browser ne ricorda l'accesso (salvato in IndexedDB) e l'app legge e scrive i file lì dentro.
// ============================================================

const CartellaExcel = (() => {
  const DB_NAME = 'rapportino-cartella-excel'; // diverso dal Lettore Presenze: le due app non si mescolano
  const STORE = 'cartelle';
  const KEY = 'excel';
  const BACKUP_KEEP = 20;
  let dirHandle = null;

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

  function supported() {
    return typeof window.showDirectoryPicker === 'function';
  }

  // Stato all'avvio: 'pronta' (accesso già concesso), 'da-confermare' (serve un clic) o 'mancante'
  async function restore() {
    try {
      const handle = await loadHandle();
      if (!handle) return 'mancante';
      dirHandle = handle;
      const perm = await handle.queryPermission({ mode: 'readwrite' });
      return perm === 'granted' ? 'pronta' : 'da-confermare';
    } catch (_) {
      return 'mancante';
    }
  }

  // Da chiamare dopo un clic dell'utente (il browser chiede la conferma solo in risposta a un clic)
  async function confirmAccess() {
    if (!dirHandle) return false;
    return (await dirHandle.requestPermission({ mode: 'readwrite' })) === 'granted';
  }

  async function choose() {
    const handle = await window.showDirectoryPicker({ id: 'cartella-excel', mode: 'readwrite' });
    dirHandle = handle;
    await saveHandle(handle);
    return handle.name;
  }

  // Solo per le prove automatiche: usa una cartella già pronta (es. lo spazio privato del browser)
  async function useHandle(handle) {
    dirHandle = handle;
  }

  function name() {
    return dirHandle ? dirHandle.name : '';
  }

  function requireDir() {
    if (!dirHandle) throw new Error('Cartella dei file Excel non collegata: premi "Collega cartella Excel"');
    return dirHandle;
  }

  async function exists(fileName) {
    try {
      await requireDir().getFileHandle(fileName);
      return true;
    } catch (_) {
      return false;
    }
  }

  // File Excel della cartella (senza copie di backup e senza i file temporanei di Excel "~$...")
  async function listExcelFiles() {
    const files = [];
    for await (const [entryName, entry] of requireDir().entries()) {
      if (entry.kind === 'file' && /\.(xlsm|xlsx)$/i.test(entryName) && !entryName.includes('_BACKUP') && !entryName.startsWith('~$')) {
        files.push(entryName);
      }
    }
    // Master per primo, poi i mesi in ordine naturale (10-Ottobre, 11-Novembre, 12-Dicembre...)
    return files.sort((a, b) => {
      const am = /master_operai/i.test(a), bm = /master_operai/i.test(b);
      if (am !== bm) return am ? -1 : 1;
      return a.localeCompare(b, 'it', { numeric: true, sensitivity: 'base' });
    });
  }

  function masterName(files) {
    return files.find(f => /master_operai/i.test(f) && /\.xlsm$/i.test(f)) || files.find(f => /master_operai/i.test(f)) || '';
  }

  async function readBytes(fileName) {
    const file = await (await requireDir().getFileHandle(fileName)).getFile();
    return new Uint8Array(await file.arrayBuffer());
  }

  // createWritable scrive su un file temporaneo e lo sostituisce solo a scrittura completata
  async function writeBytes(fileName, bytes, create = false) {
    const handle = await requireDir().getFileHandle(fileName, { create });
    const writable = await handle.createWritable();
    try {
      await writable.write(bytes);
      await writable.close();
    } catch (e) {
      try { await writable.abort(); } catch (_) {}
      throw e;
    }
  }

  // File aperto in Excel: Excel crea accanto il file di blocco "~$nome". Salvando da Excel si perderebbero
  // i dati scritti dall'app, quindi la scrittura viene fermata
  async function isOpenInExcel(fileName) {
    return exists('~$' + fileName);
  }

  // Copia con data e ora nella sottocartella "Backup", es. Backup/Presenze_10-Ottobre_2026_BACKUP_2026-10-23_04-55-12.xlsm.
  // Per ogni file restano le ultime BACKUP_KEEP copie
  async function datedBackup(fileName, bytes, label = 'BACKUP') {
    const backupDir = await requireDir().getDirectoryHandle('Backup', { create: true });
    const dot = fileName.lastIndexOf('.');
    const base = fileName.slice(0, dot), ext = fileName.slice(dot);
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    const stamp = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}`;

    const existing = [];
    for await (const [entryName] of backupDir.entries()) existing.push(entryName);
    let backupName = `${base}_${label}_${stamp}${ext}`;
    for (let n = 2; existing.includes(backupName); n++) backupName = `${base}_${label}_${stamp}_${n}${ext}`; // stesso secondo

    const writable = await (await backupDir.getFileHandle(backupName, { create: true })).createWritable();
    await writable.write(bytes);
    await writable.close();

    // Le copie più vecchie oltre le ultime BACKUP_KEEP vengono eliminate (il nome con data e ora si ordina da solo)
    const prefix = `${base}_${label}_`;
    const old = [...existing, backupName].filter(f => f.startsWith(prefix) && f.endsWith(ext)).sort();
    for (const f of old.slice(0, Math.max(0, old.length - BACKUP_KEEP))) {
      try { await backupDir.removeEntry(f); } catch (_) {}
    }
    return backupName;
  }

  // C'è già la copia di sicurezza di oggi per questo file, in QUESTA cartella? (guarda dentro Backup:
  // con un promemoria nel browser due cartelle con lo stesso nome si confondevano)
  async function hasBackupToday(fileName, label = 'BACKUP') {
    let backupDir;
    try { backupDir = await requireDir().getDirectoryHandle('Backup'); } catch (_) { return false; }
    const dot = fileName.lastIndexOf('.');
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    const prefix = `${fileName.slice(0, dot)}_${label}_${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_`;
    for await (const [entryName] of backupDir.entries()) if (entryName.startsWith(prefix)) return true;
    return false;
  }

  return { supported, restore, confirmAccess, choose, useHandle, name, exists, listExcelFiles, masterName, readBytes, writeBytes, isOpenInExcel, datedBackup, hasBackupToday };
})();
