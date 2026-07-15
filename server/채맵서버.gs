/************************************************************
 * 채맵 서버 (구글 앱스 스크립트, 무료) — v2
 * 지인 등록 접수 → 사진(여러 장, 분류) 드라이브 저장 →
 * 관리자 승인 시 카카오맵 링크에서 이름·주소 자동 수집 →
 * 채맵 앱이 읽어가는 승인 목록(JSON) 제공. 수정 요청도 반영.
 *
 * ▶ 코드 바꾼 뒤엔: 저장 → setup 실행 → 배포 관리 → ✏️ → 새 버전 → 배포
 ************************************************************/
const PROP = PropertiesService.getScriptProperties();
const HEADER = ['id','status','type','link','name','category','note','from',
  'photos','placeName','address','phone','hours','menu','editId','created'];

function setup(){
  let sheetId = PROP.getProperty('SHEET_ID');
  let ss;
  if(!sheetId){ ss = SpreadsheetApp.create('채맵DB'); PROP.setProperty('SHEET_ID', ss.getId()); }
  else ss = SpreadsheetApp.openById(sheetId);
  let sh = ss.getSheetByName('요청') || ss.getSheets()[0];
  sh.setName('요청');
  sh.getRange(1,1,1,HEADER.length).setValues([HEADER]); // 헤더 최신화(데이터 없을 때 안전)
  if(!PROP.getProperty('FOLDER_ID')) PROP.setProperty('FOLDER_ID', DriveApp.createFolder('채맵사진').getId());
  if(!PROP.getProperty('ADMIN_KEY')) PROP.setProperty('ADMIN_KEY', Utilities.getUuid().replace(/-/g,'').slice(0,12));
  if(!PROP.getProperty('SITE_PW')) PROP.setProperty('SITE_PW', '0828');
  Logger.log('준비 완료');
  Logger.log('ADMIN_KEY: ' + PROP.getProperty('ADMIN_KEY'));
  Logger.log('SITE_PW: ' + PROP.getProperty('SITE_PW'));
}

function sheet_(){ return SpreadsheetApp.openById(PROP.getProperty('SHEET_ID')).getSheetByName('요청'); }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function readAll_(){
  const vals = sheet_().getDataRange().getValues(); const head = vals.shift();
  return vals.map(r=>{ const o={}; head.forEach((h,i)=>o[h]=r[i]);
    try{ o.photos = JSON.parse(o.photos||'[]'); }catch(e){ o.photos=[]; } return o; });
}

function doGet(e){
  const p = e.parameter || {};
  if(p.action === 'photo' && p.id){
    return ContentService.createTextOutput(Utilities.base64Encode(DriveApp.getFileById(p.id).getBlob().getBytes()))
      .setMimeType(ContentService.MimeType.TEXT);
  }
  const rows = readAll_();
  if(p.action === 'list'){
    if(p.key !== PROP.getProperty('ADMIN_KEY')) return json_({error:'no-auth'});
    return json_({items: rows});
  }
  if(p.action === 'approved'){
    if(p.pw !== PROP.getProperty('SITE_PW')) return json_({error:'no-auth'});
    return json_({items: rows.filter(r=>r.status==='approved')});
  }
  return json_({ok:true, msg:'채맵 서버 작동 중'});
}

function savePhotos_(arr){ // [{cat, data}] → [{cat, id}]
  if(!arr || !arr.length) return [];
  const folder = DriveApp.getFolderById(PROP.getProperty('FOLDER_ID'));
  return arr.map((ph, i)=>{
    const m = String(ph.data||'').match(/^data:([^;]+);base64,(.+)$/);
    if(!m) return null;
    const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'p'+Date.now()+'_'+i));
    return {cat: ph.cat||'기타', id: f.getId()};
  }).filter(Boolean);
}

function doPost(e){
  let b = {};
  try{ b = JSON.parse(e.postData.contents); }catch(err){ return json_({error:'bad-json'}); }

  if(b.action === 'submit'){
    if(b.pw !== PROP.getProperty('SITE_PW')) return json_({error:'no-auth'});
    const id = 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2,6);
    const photos = savePhotos_(b.photos);
    const row = {};
    HEADER.forEach(h=>row[h]='');
    Object.assign(row, {id, status:'pending', type:b.type||'등록', link:b.link||'', name:b.name||'',
      category:b.category||'', note:b.note||'', from:b.from||'', photos:JSON.stringify(photos),
      phone:b.phone||'', hours:b.hours||'', menu:b.menu||'', editId:b.editId||'', created:new Date()});
    sheet_().appendRow(HEADER.map(h=>row[h]));
    return json_({ok:true, id});
  }

  if(b.key !== PROP.getProperty('ADMIN_KEY')) return json_({error:'no-auth'});
  const sh = sheet_(); const vals = sh.getDataRange().getValues(); const head = vals[0];
  const col = name => head.indexOf(name)+1;
  const findRow = id => { for(let i=1;i<vals.length;i++) if(vals[i][head.indexOf('id')]===id) return i+1; return 0; };
  const row = findRow(b.id);
  if(!row) return json_({error:'not-found'});

  if(b.action === 'delete'){ sh.deleteRow(row); return json_({ok:true}); }
  if(b.action === 'reject'){ sh.getRange(row, col('status')).setValue('rejected'); return json_({ok:true}); }
  if(b.action === 'approve'){
    const link = vals[row-1][head.indexOf('link')];
    const info = fetchKakao_(link);
    const type = vals[row-1][head.indexOf('type')];
    const editId = vals[row-1][head.indexOf('editId')];
    if(type==='수정' && editId){
      // 대상 행에 비어있지 않은 값만 덮어쓰기
      const tRow = findRow(editId);
      if(tRow){
        ['category','note','phone','hours','menu'].forEach(f=>{
          const v = vals[row-1][head.indexOf(f)];
          if(v!=='' && v!=null) sh.getRange(tRow, col(f)).setValue(v);
        });
        const newPhotos = vals[row-1][head.indexOf('photos')];
        if(newPhotos && newPhotos!=='[]') sh.getRange(tRow, col('photos')).setValue(newPhotos);
      }
      sh.getRange(row, col('status')).setValue('rejected'); // 수정요청 자체는 목록에서 처리완료로
      return json_({ok:true, applied:true});
    }
    if(info.name) sh.getRange(row, col('placeName')).setValue(info.name);
    if(info.address) sh.getRange(row, col('address')).setValue(info.address);
    sh.getRange(row, col('status')).setValue('approved');
    return json_({ok:true, info});
  }
  return json_({error:'bad-action'});
}

/* 카카오맵 링크 → 실제 장소 id → og:title(이름)·og:description(주소) 수집.
   kko.to 짧은 링크는 최종 페이지 HTML의 'place?id=숫자'에서 id를 뽑음(앱스스크립트가
   리다이렉트 헤더를 삼켜서 이 방식이 안정적) */
function fetchKakao_(link){
  try{
    const F = u => UrlFetchApp.fetch(u, {followRedirects:true, muteHttpExceptions:true, headers:{'User-Agent':'Mozilla/5.0'}}).getContentText();
    let id = (String(link).match(/place\.map\.kakao\.com\/(\d+)/)||[])[1]
          || (String(link).match(/[?&]id=(\d+)/)||[])[1] || '';
    if(!id){
      const h0 = F(link);
      id = (h0.match(/place\?id=(\d+)/)||[])[1]
        || (h0.match(/[?&]id=(\d+)/)||[])[1]
        || (h0.match(/"id"\s*:\s*"?(\d+)"?/)||[])[1] || '';
    }
    if(!id) return {name:'', address:''};
    const html = F('https://place.map.kakao.com/'+id);
    const g = re => { const m = html.match(re); return m ? m[1].trim() : ''; };
    let name = g(/<meta property="og:title" content="([^"]*)"/);
    let addr = g(/<meta property="og:description" content="([^"]*)"/);
    if(name==='카카오맵'){ name=''; addr=''; }
    return {name, address: addr};
  }catch(e){ return {name:'', address:''}; }
}
