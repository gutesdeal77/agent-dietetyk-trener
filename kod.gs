/********** PODSTAWY **********/
function ss(){ return SpreadsheetApp.getActive(); }

// 🔧 USTAWIENIA – podmień tylko te ID gdyby się zmieniły foldery
const FOLDER_ID          = '1SYaxKQP_dOz4e4nwF5mg9aS7kmP_7uzq'; // folder z CSV Skaner z Lodówki
const CSV_FOLDER_ID      = FOLDER_ID;                            // alias dla funkcji CSV
const JSON_FOLDER_ID     = '1fQdNZCb3ar18J-k_KJVBWpGaXa0I3-22'; // folder z JSON E-PARAGONAMI (Biedronka)
const OCR_FOLDER_ID      = '1F7L3AC2VSSgfn39RUcZ2-XXwlV33qiZC';

const SCAN_SHEET         = 'Skany';      // [timestamp, ean, status]
const DB_SHEET           = 'DB';         // [ean, name, kcal_100g, unit, Domyślne_dni, Status]
const PANTRY_SHEET       = 'Spiżarka';   // [timestamp, ean, name, qty, unit, kcal_100g, expiry, status]
const RECEIPTS_SHEET     = 'Paragony';   // [źródło | plik | paragon_uid | data | sklep_nip | numer_dokumentu | płatność | lp | typ | nazwa_raw | ean | ilość | cena_jedn_brutto_zł | wartość_brutto_zł | vat_id | vat_stawka | status | uwagi]
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
      .addItem('Utwórz nagłówki „Paragony”', 'ensureParagonyHeaders_')
      .addItem('Importuj e-Paragony (JSON) z Drive', 'importReceiptsFromDrive_')
      .addItem('Diag: lista JSON w folderze', 'diagListJsonInFolder_')
      .addItem('Diag: lista OCR (PDF/JPG)', 'diagListOcrInFolder_')
      .addItem('Przelicz „Spiżarkę” (sumy)', 'recalcPantryTotals_')


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

function receiptsIndex_(){
  const sh   = ss().getSheetByName(RECEIPTS_SHEET);
  const last = sh.getLastRow();
  const byFile = new Set(last>1 ? sh.getRange(2,2,last-1,1).getValues().flat().filter(Boolean) : []);
  const byUid  = new Set(last>1 ? sh.getRange(2,3,last-1,1).getValues().flat().filter(Boolean) : []);
  return { byFile, byUid };
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
    ensureParagonyHeaders_(); 
    Logger.log('ensureSheets: OK');
  } catch(e){ Logger.log('ensureSheets error: ' + e); }
}

/********** DRIVE API – LISTA PLIKÓW W FOLDERZE **********/
// WYMAGA: Usługi → włącz „Drive API” (zaawansowane)
function listCsvFilesViaApi_(){
  // v3: q = " 'folderId' in parents and trashed=false "
  const q = `'${CSV_FOLDER_ID}' in parents and trashed=false`;
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

/********** DIAGNOSTYKA CSV **********/
function diagListCsvInFolder(){
  try{
    const files = listCsvFilesViaApi_();
    Logger.log(`CSV_FOLDER_ID=${CSV_FOLDER_ID} | plików CSV bez .done: ${files.length}`);
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

/********** PARAGONY: nagłówki **********/
function ensureParagonyHeaders_(){
  const sh = ss().getSheetByName(RECEIPTS_SHEET) || ss().insertSheet(RECEIPTS_SHEET);
  const cols = ['źródło','plik','paragon_uid','data','sklep_nip','numer_dokumentu','płatność','lp','typ','nazwa_raw','ean','ilość','cena_jedn_brutto_zł','wartość_brutto_zł','vat_id','vat_stawka','status','uwagi'];
  sh.getRange(1,1,1,cols.length).setValues([cols]);
  sh.setFrozenRows(1);
  try { const f = sh.getFilter(); if (f) f.remove(); } catch(e) {}
  sh.getRange(1,1,sh.getMaxRows(),cols.length).createFilter();
  sh.getRange('K:K').setNumberFormat('@');                // ean tekst
  sh.getRange('D:D').setNumberFormat('yyyy-mm-dd HH:mm'); // data-czas
  sh.getRange('M:N').setNumberFormat('0.00');             // kwoty
}

/********** PARAGONY: import JSON **********/
function importReceiptsFromDrive_(){
  ensureParagonyHeaders_();
  const sh = ss().getSheetByName(RECEIPTS_SHEET);
  const q = `'${JSON_FOLDER_ID}' in parents and trashed=false and (mimeType='application/json' or name contains '.json')`;
  const files = (Drive.Files.list({ q, pageSize: 1000 }).files || [])
                 .filter(f => !(f.name||'').endsWith(PROCESSED_SUFFIX));

  const idx = receiptsIndex_();        // lista już zaimportowanych plików/paragonów
  let batch = [];                      // zbierzemy wiersze i wstawimy hurtem

  files.forEach(f=>{
    if (idx.byFile.has(f.name)) return;  // plik już był → nie dubluj
    try{
      const txt = DriveApp.getFileById(f.id).getBlob().getDataAsString('UTF-8');
      const obj = JSON.parse(txt);
      const rows = parseReceiptJson_(obj, f.name);
      if (rows && rows.length){
        batch = batch.concat(rows);
        idx.byFile.add(f.name);         // żeby w tej samej sesji też nie dublować
      }
    }catch(e){ Logger.log('JSON error: '+(f.name||f.id)+' '+e); }
  });

  if (batch.length){
    sh.getRange(sh.getLastRow()+1,1,batch.length,batch[0].length).setValues(batch);
  }
  SpreadsheetApp.getActive().toast(`Import: ${batch.length} wierszy z ${files.length} plików`);

  // .done lub kosz po sukcesie
  files.forEach(f=>{
    try{
      if (TRASH_AFTER_IMPORT) {
        Drive.Files.update({ trashed: true }, f.id);
      } else {
        Drive.Files.update({ name: f.name + PROCESSED_SUFFIX }, f.id);
      }
    }catch(e){ Logger.log('rename/trash error: '+(f.name||f.id)+' '+e); }
  });
}

function parseReceiptJson_(obj, filename){
  const hdr = obj.header||[];
  let tin='', docNum='', dateISO='';
  for (let i=0;i<hdr.length;i++){
    if (hdr[i].headerData){
      const h = hdr[i].headerData;
      tin = h.tin || tin;
      docNum = h.docNumber || docNum;
      dateISO = h.date || dateISO;
    }
  }
  let pay='';
  (obj.body||[]).some(it=>{ if(it.payment){ pay = it.payment.name||''; return true; } return false; });

  let uid='';
  const re = /Numer[^>]*>(\d{10,})</;
  (obj.body||[]).forEach(it=>{
    if(it.addLine && typeof it.addLine.data==='string'){
      const m = it.addLine.data.match(re);
      if(m) uid = uid || m[1];
    }
  });

  const out=[], body = obj.body||[];
  let lp=0;
  body.forEach(it=>{
    if (it.sellLine){
      const s = it.sellLine;
      lp++;
      out.push(['e-paragon', filename, uid, new Date(dateISO), tin, docNum, pay,
                lp, 'sell', s.name||'', '', Number(s.quantity||0),
                (s.price||0)/100, (s.total||0)/100, s.vatId||'', '', '', '' ]);
    } else if (it.discountLine){
      const d = it.discountLine;
      out.push(['e-paragon', filename, uid, new Date(dateISO), tin, docNum, pay,
                '', 'discount', '', '', '', '', -(d.value||0)/100, d.vatId||'', '', '', '' ]);
    }
  });
  return out;
}

function diagListJsonInFolder_(){
  const q = `'${JSON_FOLDER_ID}' in parents and trashed=false and (mimeType='application/json' or name contains '.json')`;
  const res = Drive.Files.list({ q, pageSize: 200 });
  const files = res.files || [];

  const sh = ss().getSheetByName('DiagJSON') || ss().insertSheet('DiagJSON');
  sh.clear();
  sh.getRange(1,1,1,2).setValues([['name','id']]);
  if (files.length) {
    sh.getRange(2,1,files.length,2).setValues(files.map(f => [f.name, f.id]));
  }
  SpreadsheetApp.getActive().toast(`DiagJSON: ${files.length} plików`);
}
function diagListOcrInFolder_(){
  const q = `'${OCR_FOLDER_ID}' in parents and trashed=false and (mimeType='application/pdf' or mimeType contains 'image/')`;
  const res = Drive.Files.list({ q, pageSize: 200 });
  const files = res.files || [];
  const sh = ss().getSheetByName('DiagOCR') || ss().insertSheet('DiagOCR');
  sh.clear();
  sh.getRange(1,1,1,3).setValues([['name','id','mimeType']]);
  if(files.length) sh.getRange(2,1,files.length,3).setValues(files.map(f=>[f.name,f.id,f.mimeType]));
  SpreadsheetApp.getActive().toast(`DiagOCR: ${files.length} plików`);
}
function recalcPantryTotals_(){
  const sh = ss().getSheetByName('Spiżarka');
  const H  = headersMap_(sh);
  if(!H['ilość_łącznie'] || !H['kcal_łącznie']){
    SpreadsheetApp.getActive().toast('Brak kolumn: ilość_łącznie / kcal_łącznie'); return;
  }
  const last = sh.getLastRow(); if(last<2) return;
  const e = H['ean'], q = H['qty'], k100 = H['kcal_100g'], Q = H['ilość_łącznie'], K = H['kcal_łącznie'];

  sh.getRange(2, Q, last-1, 1)
  .setFormulaR1C1(`=IF(RC${e}="","",SUMIF(C${e},RC${e},C${q}))`);

sh.getRange(2, K, last-1, 1)
  .setFormulaR1C1(`=IF(RC${e}="","",RC${Q}*INDEX(C${k100},MATCH(RC${e},C${e},0)))`);

  SpreadsheetApp.getActive().toast('Przeliczono sumy w „Spiżarce”');
}


