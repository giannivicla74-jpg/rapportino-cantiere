// ============================================================
// PONTE PC LOCALE (opzionale, alternativa al collegamento cartella di Chrome)
// Se sul PC del responsabile è avviato Avvia_Ponte_Rapportino.bat,
// l'app parla con http://127.0.0.1:47821 per leggere/scrivere Excel e Archivio.
// Se il Ponte è spento, tutto continua a funzionare con la File System Access API.
// ============================================================

const PonteLocale = (() => {
  const BASE = 'http://127.0.0.1:47821';
  const CACHE_MS = 1500;
  let ultimo = null;
  let ultimoControllo = 0;

  function fetchTimeout(url, opts = {}, ms = 1000) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), ms);
    return fetch(url, { cache: 'no-store', ...opts, signal: ctrl.signal })
      .finally(() => clearTimeout(timer));
  }

  async function controlla(forza = false) {
    if (window.__prova?.senzaPonte) {
      ultimo = null;
      ultimoControllo = Date.now();
      return null;
    }
    const ora = Date.now();
    if (!forza && ultimoControllo && (ora - ultimoControllo) < CACHE_MS) return ultimo;
    try {
      const res = await fetchTimeout(`${BASE}/stato`, {}, 900);
      if (!res.ok) throw new Error('non raggiungibile');
      const dati = await res.json();
      ultimo = dati && dati.ok ? dati : null;
    } catch (_) {
      ultimo = null;
    }
    ultimoControllo = Date.now();
    return ultimo;
  }

  function ultimoStato() { return ultimo; }
  function attivo() { return Boolean(ultimo && ultimo.ok); }

  async function req(percorso, opts = {}, ms = 60000) {
    let res;
    try {
      res = await fetchTimeout(`${BASE}${percorso}`, opts, ms);
    } catch (_) {
      ultimo = null;
      ultimoControllo = Date.now();
      throw new Error('Ponte PC non raggiungibile: controlla che la finestra del Ponte sia aperta');
    }
    if (!res.ok) {
      let msg = `errore Ponte PC (${res.status})`;
      try {
        const j = await res.json();
        if (j && j.errore) msg = j.errore;
      } catch (_) {}
      throw new Error(msg);
    }
    return res;
  }

  async function excelLista() {
    const j = await (await req('/excel/lista')).json();
    return Array.isArray(j.files) ? j.files : [];
  }

  async function excelEsiste(nome) {
    const j = await (await req(`/excel/esiste?nome=${encodeURIComponent(nome)}`)).json();
    return Boolean(j.esiste);
  }

  async function excelLeggi(nome) {
    const res = await req(`/excel/leggi?nome=${encodeURIComponent(nome)}`);
    return new Uint8Array(await res.arrayBuffer());
  }

  async function excelScrivi(nome, bytes, create = false) {
    await req(`/excel/scrivi?nome=${encodeURIComponent(nome)}&crea=${create ? '1' : '0'}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: bytes
    });
  }

  async function excelHaBackupOggi(nome, label = 'BACKUP') {
    const j = await (await req(`/excel/backup-oggi?nome=${encodeURIComponent(nome)}&label=${encodeURIComponent(label)}`)).json();
    return Boolean(j.presente);
  }

  async function excelBackup(nome, bytes, label = 'BACKUP', tieni = 20) {
    const j = await (await req(`/excel/backup?nome=${encodeURIComponent(nome)}&label=${encodeURIComponent(label)}&tieni=${tieni}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream' },
      body: bytes
    })).json();
    return j.backupName;
  }

  async function archivioSalva(percorso, nome, testo) {
    const p = Array.isArray(percorso) ? percorso.join('/') : String(percorso || '');
    await req(`/archivio/salva?percorso=${encodeURIComponent(p)}&nome=${encodeURIComponent(nome)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
      body: testo
    });
  }

  async function archivioEsiste(percorso, nome) {
    const p = Array.isArray(percorso) ? percorso.join('/') : String(percorso || '');
    const j = await (await req(`/archivio/esiste?percorso=${encodeURIComponent(p)}&nome=${encodeURIComponent(nome)}`)).json();
    return Boolean(j.esiste);
  }

  return {
    controlla, ultimoStato, attivo,
    excelLista, excelEsiste, excelLeggi, excelScrivi, excelHaBackupOggi, excelBackup,
    archivioSalva, archivioEsiste
  };
})();
