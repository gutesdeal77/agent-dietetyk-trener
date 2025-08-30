/********** PODSTAWY **********/
function ss(){ return SpreadsheetApp.getActive(); }

// 🔧 USTAWIENIA – podmień tylko to ID gdyby się zmienił folder
const FOLDER_ID          = '1SYaxKQP_dOz4e4nwF5mg9aS7kmP_7uzq'; // folder z CSV
const SCAN_SHEET         = 'Skany';      // [timestamp, ean, status]
const DB_SHEET           = 'DB';         // [ean, name, kcal_100g, unit, Domyślne_dni, Status]
const PANTRY_SHEET       = 'Spiżarka';   // [timestamp, ean, name, qty, unit, kcal_100g, expiry, status]
const PROCESSED_SUFFIX   = '.done';      // dopinamy do nazwy po imporcie
const TRASH_AFTER_IMPORT = false;        // true = do kosza zamiast .done

/********** MENU (opcjonalnie) **********/
function onOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('📦 EAN Importer')
      .addItem('Utwórz nagłówki/zakładki', 'ensureSheets')
      .addItem('Importuj CSV z Drive (API)', 'importCSVFromDrive')
      .addItem('Przetwórz „Skany” → „Spiżarka”', 'processInbox')
      .addSeparator()
      .addItem('Diag: lista plików', 'diagListCsvInFolder')
      .addItem('Diag: przetwórz 1 plik', 'diagProcessOneFile')
      .addToUi();
  } catch(e){ Logger.log('onOpen error: ' + e); }
}

/********** HELPERY **********/
const NORM = v => String(v||'').replace(/\D/g,'');
const NOW  = () => new Date();

function headersMap_(sh){
  const row = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn())).getValues()[0];
  const m = {}; row.forEach((name,i)=>{ if(name) m[name]=i+1; });
  return m;
}
function ensureHeaders_(sh, names){
  const have = sh.getLastRow()>=1 && sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].some(v=>v);
  if(!have) sh.getRange(1,1,1,names.length).setValues([names]);
}
function detectSep_(txt){
  const first = (txt.split(/\r?\n/)[0]||'');
  const cnt = s => (first.match(new RegExp('\\'+s,'g'))||[]).length;
  return cnt(';')>cnt(',') ? ';' : ',';
}

/********** ZAKŁADKI / NAGŁÓWKI **********/
function ensureSheets(){
  try{
    const s  = ss();
    const sk = s.getSheetByName(SCAN_SHEET)   || s.insertSheet(SCAN_SHEET);
    const db = s.getSheetByName(DB_SHEET)     || s.insertSheet(DB_SHEET);
    const sp = s.getSheetByName(PANTRY_SHEET) || s.insertSheet(PANTRY_SHEET);

    ensureHeaders_(sk, ['timestamp','ean','status']);
    ensureHeaders_(db, ['ean','name','kcal_100g','unit','Domyślne_dni','Status']);
    ensureHeaders_(sp, ['timestamp','ean','name','qty','unit','kcal_100g','expiry','status']);
    Logger.log('ensureSheets: OK');
  } catch(e){ Logger.log('ensureSheets error: ' + e); }
}

/********** DRIVE API – LISTA PLIKÓW W FOLDERZE **********/
// WYMAGA: Usługi → włącz „Drive API”
function listCsvFilesViaApi_(){
  // v3: q = " 'folderId' in parents and trashed=false "
  const q = `'${FOLDER_ID}' in parents and trashed=false`;
  const out = [];
  let pageToken;
  do{
    const res = Drive.Files.list({
      q,
      fields: 'files(id,name,modifiedTime,mimeType),nextPageToken',
      pageToken
    });
    (res.files||[]).forEach(f=>{
      if (/\.csv$/i.test(f.name) && !f.name.endsWith(PROCESSED_SUFFIX)) {
        out.push({id:f.id, name:f.name, modified:f.modifiedTime});
      }
    });
    pageToken = res.nextPageToken;
  } while(pageToken);
  return out;
}

/********** DRIVE API – POBIERANIE ZAWARTOŚCI PLIKU **********/
function fetchFileContent_(fileId){
  const url = 'https://www.googleapis.com/drive/v3/files/'+encodeURIComponent(fileId)+'?alt=media';
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer '+token }});
  return res.getContentText('UTF-8');
}

/********** DRIVE API – ZMIANA NAZWY / KOSZ **********/
function markFileProcessed_(fileId, name){
  if (TRASH_AFTER_IMPORT){
    Drive.Files.update({trashed:true}, fileId); // do kosza
  } else {
    Drive.Files.update({name: name+PROCESSED_SUFFIX}, fileId);
  }
}

/********** 1) IMPORT CSV → „Skany” **********/
function importCSVFromDrive(){
  ensureSheets();
  try{
    const files = listCsvFilesViaApi_();
    if(!files.length){ Logger.log('Import: brak świeżych CSV w folderze.'); return; }

    const sh = ss().getSheetByName(SCAN_SHEET);
    let imported = 0;

    files.forEach(f=>{
      const txt   = fetchFileContent_(f.id);
      const sep   = detectSep_(txt);
      const data  = Utilities.parseCsv(txt, sep);
      const out   = [];

      for(let i=1;i<data.length;i++){
        const r = data[i]; if(!r) continue;
        let ean = NORM(r[1]);
        if(!ean){
          const m = r.join(',').match(/\b\d{8,13}\b/);
          ean = m ? NORM(m[0]) : '';
        }
        if(ean) out.push([NOW(), ean, '']);
      }

      if(out.length){
        sh.getRange(sh.getLastRow()+1,1,out.length,out[0].length).setValues(out);
        imported += out.length;
      }
      markFileProcessed_(f.id, f.name);
      Logger.log(`Import: ${f.name} → ${out.length} wierszy`);
    });

    Logger.log(`Import CSV DONE → dodano ${imported} wierszy.`);
  } catch(e){
    Logger.log('importCSVFromDrive error: ' + e + ' | stack: ' + (e.stack||''));
  }
}

/********** 2) „Skany” → „Spiżarka” **********/
function processInbox(){
  ensureSheets();
  try{
    const s    = ss();
    const sk   = s.getSheetByName(SCAN_SHEET);
    const db   = s.getSheetByName(DB_SHEET);
    const sp   = s.getSheetByName(PANTRY_SHEET);
    const Hsk  = headersMap_(sk);
    const Hdb  = headersMap_(db);
    const Hsp  = headersMap_(sp);

    const IN = sk.getDataRange().getValues();
    if(IN.length<2){ Logger.log('Brak nowych skanów.'); return; }

    // DB → mapa
    const DB = db.getDataRange().getValues();
    const dbMap = new Map();
    for(let i=1;i<DB.length;i++){
      const row=DB[i];
      const e=NORM(row[(Hdb.ean||1)-1]); if(!e) continue;
      dbMap.set(e,{
        name:  row[(Hdb.name||2)-1] || '(brak danych)',
        kcal:  Number(row[(Hdb['kcal_100g']||3)-1] || 0),
        unit:  row[(Hdb.unit||4)-1] || 'g',
        days:  Hdb['Domyślne_dni'] ? Number(row[Hdb['Domyślne_dni']-1] || 0) : 0,
        status:Hdb['Status'] ? (row[Hdb['Status']-1] || '') : ''
      });
    }

    // klucze do deduplikacji
    const OUT = sp.getDataRange().getValues();
    const seen = new Set();
    for(let i=1;i<OUT.length;i++){
      const ts=OUT[i][(Hsp.timestamp||1)-1];
      const ee=NORM(OUT[i][(Hsp.ean||2)-1]);
      if(ts && ee) seen.add(ee+'|'+new Date(ts).toDateString());
    }

    const toOut=[], toMark=[];
    for(let r=1;r<IN.length;r++){
      const st = String(IN[r][(Hsk.status||3)-1]||'');
      if(st) continue;

      const ts  = IN[r][(Hsk.timestamp||1)-1] ? new Date(IN[r][(Hsk.timestamp||1)-1]) : NOW();
      const ean = NORM(IN[r][(Hsk.ean||2)-1]);
      if(!ean){ toMark.push([r+1,'err']); continue; }

      let meta = dbMap.get(ean);
      if(!meta){
        const off = getProductData_(ean);
        meta = { name:(off&&off.name)||'(nieznany)', kcal:(off&&off.kcal)||0, unit:'g', days:0, status:'' };
        db.appendRow([ean, meta.name, meta.kcal, meta.unit, '', '']);
        dbMap.set(ean, meta);
      }

      const addDays = meta.days || guessDaysByName_(meta.name);
      let expiry='';
      if(addDays>0){
        const d=new Date(ts); d.setDate(d.getDate()+addDays);
        expiry=Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      const key = ean+'|'+ts.toDateString();
      if(seen.has(key)){ toMark.push([r+1,'dup']); continue; }
      seen.add(key);

      toOut.push([ts, ean, meta.name, 1, meta.unit, meta.kcal, expiry, meta.status || 'neutralne']);
      toMark.push([r+1,'done']);
    }

    if(toOut.length){
      sp.getRange(sp.getLastRow()+1,1,toOut.length,toOut[0].length).setValues(toOut);
    }
    for(const [row,st] of toMark){
      sk.getRange(row,(Hsk.status||3)).setValue(st);
    }
    Logger.log(`processInbox: dodano ${toOut.length}, oznaczono ${toMark.length}`);
  } catch(e){
    Logger.log('processInbox error: ' + e + ' | stack: ' + (e.stack||''));
  }
}

/********** Open Food Facts (cache) **********/
function getProductData_(ean){
  const key='OFF_'+ean;
  const cached=CacheService.getScriptCache().get(key);
  if(cached) return JSON.parse(cached);

  const url=`https://world.openfoodfacts.org/api/v2/product/${encodeURIComponent(ean)}.json?lc=pl`;
  try{
    const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
    if(res.getResponseCode()>=200 && res.getResponseCode()<300){
      const data=JSON.parse(res.getContentText());
      if(data && data.status===1 && data.product){
        const p=data.product;
        const name=p.product_name_pl || p.product_name || `Produkt ${ean}`;
        const kcal=p.nutriments && p.nutriments['energy-kcal_100g']
                    ? Math.round(Number(p.nutriments['energy-kcal_100g'])) : 0;
        const meta={name, kcal, unit:'g'};
        CacheService.getScriptCache().put(key, JSON.stringify(meta), 43200);
        return meta;
      }
    }
  }catch(e){ Logger.log('OFF error: '+e); }
  return null;
}

/********** Heurystyka dat ważności **********/
function guessDaysByName_(name=''){
  const n=name.toLowerCase();
  if(/\b(sałata|mix sałat|rukola|szpinak)\b/.test(n)) return 3;
  if(/\b(jogurt|kefir|serek|twarożek)\b/.test(n)) return 7;
  if(/\b(mleko|śmietana)\b/.test(n)) return 5;
  if(/\b(wędlina|szynka|kiełbasa)\b/.test(n)) return 5;
  if(/\b(puszka|konserwa|tuńczyk|fasola|kukurydza)\b/.test(n)) return 180;
  return 0;
}

/********** DIAGNOSTYKA **********/
function diagListCsvInFolder(){
  try{
    const files = listCsvFilesViaApi_();
    Logger.log(`FOLDER_ID=${FOLDER_ID} | plików CSV bez .done: ${files.length}`);
    files.forEach(f=>Logger.log(`${f.name} | id=${f.id}`));
  } catch(e){
    Logger.log('diagListCsvInFolder error: ' + e + ' | stack: ' + (e.stack||''));
  }
}
function diagProcessOneFile(){
  try{
    const files = listCsvFilesViaApi_();
    if(!files.length){ Logger.log('Brak świeżych CSV do przetworzenia.'); return; }
    const f = files[0];
    const txt   = fetchFileContent_(f.id);
    const sep   = detectSep_(txt);
    const data  = Utilities.parseCsv(txt, sep);
    const out   = [];
    for(let i=1;i<data.length;i++){
      const r=data[i]; if(!r) continue;
      let ean=NORM(r[1]);
      if(!ean){
        const m=r.join(',').match(/\b\d{8,13}\b/);
        ean=m?NORM(m[0]):'';
      }
      if(ean) out.push([NOW(), ean, '']);
    }
    if(out.length){
      const sh=ss().getSheetByName(SCAN_SHEET);
      sh.getRange(sh.getLastRow()+1,1,out.length,out[0].length).setValues(out);
    }
    markFileProcessed_(f.id, f.name);
    Logger.log(`diagProcessOneFile: ${f.name} → dodano ${out.length} wierszy do „Skany”`);
  } catch(e){
    Logger.log('diagProcessOneFile error: ' + e + ' | stack: ' + (e.stack||''));
  }
}
