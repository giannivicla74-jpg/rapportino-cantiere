// ============================================================
// LETTURA E SCRITTURA DEI FILE EXCEL (.xlsm) NEL BROWSER
// Stessa logica della versione con server (excel_xml_writer.mjs), ma lavora sui byte del file
// con JSZip: nessun programma da installare. I file vengono letti e scritti tramite la cartella
// che l'utente ha autorizzato (vedi cartella.js).
// ============================================================

const ExcelXlsm = (() => {
  const textOf = xml => [...xml.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(t => t[1]).join('');
  const decode = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

  async function open(bytes) {
    return JSZip.loadAsync(bytes);
  }

  async function readText(zip, name) {
    const f = zip.file(name);
    return f ? f.async('string') : '';
  }

  async function sharedStrings(zip) {
    const xml = await readText(zip, 'xl/sharedStrings.xml');
    return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => textOf(m[1]));
  }

  // Valore di una cella come testo (stringhe condivise, testo in linea o numero)
  function cellValue(sheetXml, shared, ref) {
    const cell = sheetXml.match(new RegExp(`<c r="${ref}"([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/c>)`));
    if (!cell || !cell[2]) return '';
    if (/t="s"/.test(cell[1])) {
      const v = cell[2].match(/<v>(\d+)<\/v>/);
      return v ? decode(shared[parseInt(v[1])] || '') : '';
    }
    if (/t="inlineStr"/.test(cell[1])) return decode(textOf(cell[2]));
    const v = cell[2].match(/<v>([^<]*)<\/v>/);
    return v ? decode(v[1]) : '';
  }

  async function toBytes(zip) {
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  // Riga Excel del turno di un operaio: ogni giorno occupa 69 righe, ogni operaio 2 (1° e 2° turno)
  function rowForRecord(dayNum, workerIndex, turno) {
    return 2 + (dayNum - 1) * 69 + (workerIndex - 1) * 2 + (turno - 1);
  }

  const HOUR_COLS = { or: 'D', sd: 'E', on: 'F', sn: 'G', v: 'H', rol: 'I', malattia: 'J' };

  // Elenco commesse dall'intervallo con nome "rif.cantieri" (es. 'Archivio - OPERAI'!$L$2:$L$32)
  async function readCommesse(bytes) {
    const zip = await open(bytes);
    const workbookXml = await readText(zip, 'xl/workbook.xml');
    const def = workbookXml.match(/<definedName name="rif\.cantieri"[^>]*>([^<]+)<\/definedName>/);
    if (!def) throw new Error('Intervallo "rif.cantieri" non trovato nel file Excel');

    const ref = def[1].replace(/&apos;/g, "'").match(/^'?(.+?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
    if (!ref) throw new Error(`Intervallo "rif.cantieri" non leggibile: ${def[1]}`);
    const [, sheetName, col, rowStart, , rowEnd] = ref;

    const sheet = [...workbookXml.matchAll(/<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"/g)]
      .find(m => m[1].replace(/&amp;/g, '&') === sheetName);
    if (!sheet) throw new Error(`Foglio "${sheetName}" non trovato`);
    const relsXml = await readText(zip, 'xl/_rels/workbook.xml.rels');
    const target = relsXml.match(new RegExp(`Id="${sheet[2]}"[^>]*Target="([^"]+)"`))[1];
    const sheetXml = await readText(zip, 'xl/' + target.replace(/^\/?xl\//, ''));
    const shared = await sharedStrings(zip);

    const commesse = [];
    for (let r = parseInt(rowStart); r <= parseInt(rowEnd); r++) {
      const val = cellValue(sheetXml, shared, `${col}${r}`).trim();
      if (val) commesse.push(val);
    }
    return commesse;
  }

  // Operai nell'ordine della tabella presenze (colonna B del giorno 1, righe 2..69, una ogni 2):
  // è l'ordine con cui l'app scrive, quindi la posizione (index) di ciascun nome è quella vera del file
  async function readWorkers(bytes) {
    const zip = await open(bytes);
    const sheetXml = await readText(zip, 'xl/worksheets/sheet1.xml');
    const shared = await sharedStrings(zip);
    const workers = [];
    for (let index = 1; index <= 34; index++) {
      const name = cellValue(sheetXml, shared, `B${2 + (index - 1) * 2}`).trim();
      if (name) workers.push({ index, name, fullName: name });
    }
    if (workers.length === 0) throw new Error('Nessun operaio trovato nella colonna B della tabella presenze');
    return workers;
  }

  // Turni da scrivere le cui righe in Excel contengono GIÀ dati diversi (es. lo stesso operaio in due rapportini
  // dello stesso giorno): scrivendo si perderebbero. Righe vuote o con dati identici non sono conflitti.
  async function findConflicts(bytes, day, records) {
    const dayNum = parseInt(day);
    const zip = await open(bytes);
    const sheetXml = await readText(zip, 'xl/worksheets/sheet1.xml');
    const shared = await sharedStrings(zip);
    const num = v => { const n = parseFloat(String(v ?? '').replace(',', '.')); return isNaN(n) ? 0 : n; };

    const conflicts = [];
    for (const rec of records) {
      const wIdx = parseInt(rec.workerIndex);
      const turno = parseInt(rec.turno) === 2 ? 2 : 1;
      if (isNaN(wIdx) || wIdx < 1 || wIdx > 34) continue;
      const row = rowForRecord(dayNum, wIdx, turno);

      const existing = { commessa: cellValue(sheetXml, shared, `C${row}`).trim() };
      const nuovo = { commessa: String(rec.cantiere ?? rec.commessa ?? '').trim() };
      for (const [k, col] of Object.entries(HOUR_COLS)) {
        existing[k] = num(cellValue(sheetXml, shared, `${col}${row}`));
        nuovo[k] = num(rec[k]);
      }

      const hasData = existing.commessa !== '' || Object.keys(HOUR_COLS).some(k => existing[k] !== 0);
      const same = existing.commessa.toUpperCase() === nuovo.commessa.toUpperCase() &&
        Object.keys(HOUR_COLS).every(k => Math.abs(existing[k] - nuovo[k]) < 0.001);
      if (hasData && !same) {
        conflicts.push({ workerIndex: wIdx, workerName: rec.workerName || '', turno, existing, nuovo });
      }
    }
    return conflicts;
  }

  // Forza il ricalcolo delle formule alla prossima apertura in Excel (Excel toglie l'opzione a ogni salvataggio)
  function withFullCalcOnLoad(wbXml) {
    if (/<calcPr\b/.test(wbXml)) {
      return wbXml.replace(/<calcPr\b([^>]*?)(\/?)>/, (_, attrs, selfClose) =>
        `<calcPr${attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${selfClose}>`);
    }
    // Nello schema calcPr segue definedNames (o sheets, se non ci sono nomi definiti)
    const anchor = wbXml.includes('</definedNames>') ? '</definedNames>' : '</sheets>';
    return wbXml.replace(anchor, `${anchor}<calcPr fullCalcOnLoad="1"/>`);
  }

  // Scrive commesse e ore dei turni nelle righe del giorno: restituisce i byte del file aggiornato
  async function writePresenze(bytes, day, records) {
    const dayNum = parseInt(day);
    if (isNaN(dayNum) || dayNum < 1 || dayNum > 31) throw new Error(`Giorno non valido: ${day}`);

    const zip = await open(bytes);
    let sheetXml = await readText(zip, 'xl/worksheets/sheet1.xml');

    function updateCell(rowXml, col, row, val, isString = false) {
      const cellRef = `${col}${row}`;
      const match = rowXml.match(new RegExp(`<c r="${cellRef}"([^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/c>)`));
      if (!match) return rowXml;
      const styleMatch = match[1].match(/s="(\d+)"/);
      const sAttr = styleMatch ? ` s="${styleMatch[1]}"` : '';

      let newCell;
      if (isString) {
        if (val && String(val).trim() !== '') {
          const cleanVal = String(val).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
          newCell = `<c r="${cellRef}"${sAttr} t="inlineStr"><is><t>${cleanVal}</t></is></c>`;
        } else {
          newCell = `<c r="${cellRef}"${sAttr}/>`;
        }
      } else {
        const n = parseFloat(val);
        newCell = (!isNaN(n) && n > 0) ? `<c r="${cellRef}"${sAttr}><v>${n}</v></c>` : `<c r="${cellRef}"${sAttr}/>`;
      }
      return rowXml.replace(match[0], newCell);
    }

    let countUpdated = 0;
    for (const rec of records) {
      const wIdx = parseInt(rec.workerIndex);
      const turno = parseInt(rec.turno) === 2 ? 2 : 1;
      if (isNaN(wIdx) || wIdx < 1 || wIdx > 34) continue;
      const row = rowForRecord(dayNum, wIdx, turno);

      const rowMatch = sheetXml.match(new RegExp(`<row r="${row}"[^>]*>([\\s\\S]*?)<\\/row>`));
      if (!rowMatch) continue;

      let rowContent = rowMatch[1];
      if (rec.cantiere !== undefined) rowContent = updateCell(rowContent, 'C', row, rec.cantiere, true);
      for (const [k, col] of Object.entries(HOUR_COLS)) rowContent = updateCell(rowContent, col, row, rec[k], false);

      sheetXml = sheetXml.replace(rowMatch[0], rowMatch[0].replace(rowMatch[1], rowContent));
      countUpdated++;
    }

    zip.file('xl/worksheets/sheet1.xml', sheetXml);
    const wbXml = await readText(zip, 'xl/workbook.xml');
    if (wbXml) zip.file('xl/workbook.xml', withFullCalcOnLoad(wbXml));

    return { bytes: await toBytes(zip), countUpdated };
  }

  // Mese e anno dal nome del file (es. "Presenze_11-Novembre_2026.xlsm" -> 11, 2026)
  function parseMonthAndYear(str) {
    const monthsMap = {
      gennaio: 1, febbraio: 2, marzo: 3, aprile: 4, maggio: 5, giugno: 6,
      luglio: 7, agosto: 8, settembre: 9, ottobre: 10, novembre: 11, dicembre: 12
    };
    const lower = String(str).toLowerCase();
    let foundMonth = null;
    let foundYear = null;
    for (const [name, n] of Object.entries(monthsMap)) {
      if (lower.includes(name)) { foundMonth = n; break; }
    }
    // Anno anche attaccato a "_" o "-" (es. "_2027"): \b non lo trovava, perché "_" conta come lettera
    const yearMatch = lower.match(/(?:^|[^0-9])(202[4-9]|203[0-9])(?![0-9])/);
    if (yearMatch) foundYear = parseInt(yearMatch[1]);
    if (!foundMonth) {
      const numMonthMatch = lower.match(/(?:^|[^0-9])(0?[1-9]|1[0-2])[-_/](202[4-9]|203[0-9])/);
      if (numMonthMatch) {
        foundMonth = parseInt(numMonthMatch[1]);
        foundYear = parseInt(numMonthMatch[2]);
      }
    }
    const now = new Date();
    return { month: foundMonth || now.getMonth() + 1, year: foundYear || now.getFullYear() };
  }

  // Prepara una copia del Master per un nuovo mese: date di ogni giorno, giorni oltre fine mese nascosti,
  // commesse e ore svuotate, ricalcolo all'apertura, anno nel Riepilogo. Restituisce i byte del nuovo file
  async function prepareNewMonth(bytes, month, year) {
    const mNum = parseInt(month);
    const yNum = parseInt(year);
    const d1 = new Date(yNum, mNum - 1, 1);
    const ref = new Date(1899, 11, 30);
    const baseSerial = Math.round((d1.getTime() - ref.getTime()) / 86400000);
    const daysInMonth = new Date(yNum, mNum, 0).getDate();

    const zip = await open(bytes);
    let sheetXml = await readText(zip, 'xl/worksheets/sheet1.xml');

    // Ogni giorno occupa 69 righe a partire dalla 2: giorno e posizione nel giorno a partire dal numero di riga.
    // Il foglio si percorre UNA volta sola (prima veniva riletto per ogni riga: migliaia di volte, molto lento)
    const LAST_ROW = 2 + 31 * 69 - 1;
    const dayOfRow = r => Math.floor((r - 2) / 69) + 1;

    // 1. Date in colonna A (righe 0..67 di ogni giorno): data del giorno, vuota oltre il fine mese
    sheetXml = sheetXml.replace(/<c r="A(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (m, rStr, attrs, inner) => {
      const r = parseInt(rStr);
      if (r < 2 || r > LAST_ROW || (r - 2) % 69 > 67) return m;
      const day = dayOfRow(r);
      const sMatch = attrs.match(/s="(\d+)"/);
      const sAttr = sMatch ? ` s="${sMatch[1]}"` : '';
      let formulaPart = '';
      if (inner) {
        const fMatch = inner.match(/<f[^>]*>[\s\S]*?<\/f>|<f[^>]*\/>/);
        if (fMatch) formulaPart = fMatch[0];
      }
      return day <= daysInMonth
        ? `<c r="A${r}"${sAttr}>${formulaPart}<v>${baseSerial + (day - 1)}</v></c>`
        : `<c r="A${r}"${sAttr}>${formulaPart}</c>`;
    });

    // Righe nascoste per i giorni oltre il fine mese, visibili tutte le altre
    sheetXml = sheetXml.replace(/<row r="(\d+)"([^>]*?)(\/?)>/g, (m, rStr, rowAttrs, selfClose) => {
      const r = parseInt(rStr);
      if (r < 2 || r > LAST_ROW) return m;
      let cleanAttrs = rowAttrs.replace(/\s*hidden="1"/g, '');
      if (dayOfRow(r) > daysInMonth) cleanAttrs += ' hidden="1"';
      return `<row r="${r}"${cleanAttrs}${selfClose}>`;
    });

    // 2. Svuota commesse (col C) e ore (col D:J) per tutte le righe operai (2:2139)
    sheetXml = sheetXml.replace(/<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g, (rowMatch, rowNumStr, rowAttrs, rowInner) => {
      const rowNum = parseInt(rowNumStr);
      if (rowNum < 2 || rowNum > 2139) return rowMatch;
      const updatedInner = rowInner.replace(/<c r="([C-J])\d+"([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g, (cMatch, col, cAttrs) => {
        const styleMatch = cAttrs.match(/s="(\d+)"/);
        return `<c r="${col}${rowNum}"${styleMatch ? ` s="${styleMatch[1]}"` : ''}/>`;
      });
      return `<row r="${rowNum}"${rowAttrs}>${updatedInner}</row>`;
    });
    zip.file('xl/worksheets/sheet1.xml', sheetXml);

    // 3. Ricalcolo completo all'apertura in Excel
    const wbXml = await readText(zip, 'xl/workbook.xml');
    if (wbXml) {
      zip.file('xl/workbook.xml', wbXml.replace(/<calcPr[^>]*\/>|<calcPr[^>]*>[\s\S]*?<\/calcPr>/,
        '<calcPr calcId="162913" fullCalcOnLoad="1" forceFullCalculation="1"/>'));
    }

    // 4. Anno nel Riepilogo Presenze (sheet3.xml, colonna B righe 2..35)
    const sheet3Xml = await readText(zip, 'xl/worksheets/sheet3.xml');
    if (sheet3Xml) {
      zip.file('xl/worksheets/sheet3.xml', sheet3Xml.replace(/<c r="B(\d+)"([^>]*)>(?:<v>[^<]*<\/v>)?<\/c>/g, (m, rNum, attrs) => {
        const r = parseInt(rNum);
        return (r >= 2 && r <= 35) ? `<c r="B${r}"${attrs}><v>${yNum}</v></c>` : m;
      }));
    }

    return toBytes(zip);
  }

  return { readCommesse, readWorkers, findConflicts, writePresenze, parseMonthAndYear, prepareNewMonth };
})();
