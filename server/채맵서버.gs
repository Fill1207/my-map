/************************************************************
 * 채맵 서버 (구글 앱스 스크립트, 무료)
 * 역할: 지인 장소 등록 접수 → 사진은 구글 드라이브 저장 →
 *       관리자가 승인하면 카카오맵 링크에서 이름·주소를 자동 수집 →
 *       채맵 앱이 읽어가는 승인 목록(JSON) 제공
 *
 * ▶ 최초 1회: 위 메뉴에서 함수 'setup' 선택 후 실행(▶) → 권한 허용
 *   실행 로그에 나오는 ADMIN_KEY(관리자 열쇠)를 클로드에게 알려주세요.
 * ▶ 그 다음: 배포 → 새 배포 → 웹 앱 → "액세스 권한: 모든 사용자" → 배포
 *   나오는 웹 앱 URL을 클로드에게 알려주세요.
 ************************************************************/

const PROP = PropertiesService.getScriptProperties();

/* 최초 1회 실행: 저장용 스프레드시트·드라이브 폴더·관리자 열쇠 생성 */
function setup(){
  let sheetId = PROP.getProperty('SHEET_ID');
  if(!sheetId){
    const ss = SpreadsheetApp.create('채맵DB');
    const sh = ss.getActiveSheet();
    sh.setName('요청');
    sh.appendRow(['id','status','type','link','name','category','note','from',
                  'photoB','photoF','placeName','address','created']);
    sheetId = ss.getId();
    PROP.setProperty('SHEET_ID', sheetId);
  }
  let folderId = PROP.getProperty('FOLDER_ID');
  if(!folderId){
    const folder = DriveApp.createFolder('채맵사진');
    folderId = folder.getId();
    PROP.setProperty('FOLDER_ID', folderId);
  }
  let key = PROP.getProperty('ADMIN_KEY');
  if(!key){
    key = Utilities.getUuid().replace(/-/g,'').slice(0,12);
    PROP.setProperty('ADMIN_KEY', key);
  }
  // 사이트 비밀번호 — 이걸 아는 사람만 읽기/등록 가능 (지인만 보게)
  if(!PROP.getProperty('SITE_PW')) PROP.setProperty('SITE_PW', '0828');
  Logger.log('✅ 준비 완료');
  Logger.log('ADMIN_KEY(관리자 열쇠): ' + key);
  Logger.log('SITE_PW(사이트 비번): ' + PROP.getProperty('SITE_PW'));
  Logger.log('SHEET_ID: ' + sheetId);
}

function sheet_(){
  return SpreadsheetApp.openById(PROP.getProperty('SHEET_ID')).getSheetByName('요청');
}
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===== 읽기 (GET) =====
 * ?action=list&key=관리자열쇠   → 모든 요청(관리자용)
 * ?action=approved              → 승인된 장소만(앱이 지도에 표시)
 * ?action=photo&id=드라이브파일ID → 사진 이미지
 */
function doGet(e){
  const p = e.parameter || {};
  if(p.action === 'photo' && p.id){
    const file = DriveApp.getFileById(p.id);
    return ContentService.createTextOutput(Utilities.base64Encode(file.getBlob().getBytes()))
      .setMimeType(ContentService.MimeType.TEXT); // 앱에서 data URL로 표시
  }
  const rows = readAll_();
  if(p.action === 'list'){
    if(p.key !== PROP.getProperty('ADMIN_KEY')) return json_({error:'no-auth'});
    return json_({items: rows});
  }
  if(p.action === 'approved'){
    if(p.pw !== PROP.getProperty('SITE_PW')) return json_({error:'no-auth'}); // 비번 아는 사람만
    return json_({items: rows.filter(r=>r.status==='approved')});
  }
  return json_({ok:true, msg:'채맵 서버 작동 중'});
}

function readAll_(){
  const sh = sheet_();
  const vals = sh.getDataRange().getValues();
  const head = vals.shift();
  return vals.map(r=>{ const o={}; head.forEach((h,i)=>o[h]=r[i]); return o; });
}

/* ===== 쓰기 (POST, JSON 본문) =====
 * {action:'submit', ...}                    지인 등록 접수
 * {action:'approve'|'reject'|'delete', id, key}  관리자 처리
 */
function doPost(e){
  let b = {};
  try{ b = JSON.parse(e.postData.contents); }catch(err){ return json_({error:'bad-json'}); }

  if(b.action === 'submit'){
    if(b.pw !== PROP.getProperty('SITE_PW')) return json_({error:'no-auth'}); // 비번 아는 사람만 등록
    const id = 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2,6);
    const folder = DriveApp.getFolderById(PROP.getProperty('FOLDER_ID'));
    const save = (dataUrl, tag) => {
      if(!dataUrl) return '';
      const m = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
      if(!m) return '';
      const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id+'_'+tag);
      return folder.createFile(blob).getId();
    };
    const photoB = save(b.photoB, 'b');
    const photoF = save(b.photoF, 'f');
    sheet_().appendRow([id, 'pending', b.type||'등록', b.link||'', b.name||'',
      b.category||'', b.note||'', b.from||'', photoB, photoF, '', '', new Date()]);
    return json_({ok:true, id});
  }

  // 이하 관리자 전용
  if(b.key !== PROP.getProperty('ADMIN_KEY')) return json_({error:'no-auth'});
  const sh = sheet_(); const vals = sh.getDataRange().getValues(); const head = vals[0];
  const idCol = head.indexOf('id');
  for(let i=1;i<vals.length;i++){
    if(vals[i][idCol] === b.id){
      const row = i+1;
      if(b.action === 'delete'){ sh.deleteRow(row); return json_({ok:true}); }
      if(b.action === 'reject'){ sh.getRange(row, head.indexOf('status')+1).setValue('rejected'); return json_({ok:true}); }
      if(b.action === 'approve'){
        // 카카오맵 링크에서 이름·주소 자동 수집
        const link = vals[i][head.indexOf('link')];
        const info = fetchKakao_(link);
        if(info.name)    sh.getRange(row, head.indexOf('placeName')+1).setValue(info.name);
        if(info.address) sh.getRange(row, head.indexOf('address')+1).setValue(info.address);
        sh.getRange(row, head.indexOf('status')+1).setValue('approved');
        return json_({ok:true, info});
      }
    }
  }
  return json_({error:'not-found'});
}

/* 카카오맵 공유 링크 → 최종 페이지의 og:title(이름)·og:description(주소) 수집 */
function fetchKakao_(link){
  try{
    const res = UrlFetchApp.fetch(link, {followRedirects:true, muteHttpExceptions:true,
      headers:{'User-Agent':'Mozilla/5.0'}});
    const html = res.getContentText();
    const g = re => { const m = html.match(re); return m ? m[1].trim() : ''; };
    return {
      name: g(/<meta property="og:title" content="([^"]*)"/),
      address: g(/<meta property="og:description" content="([^"]*)"/)
    };
  }catch(err){ return {name:'', address:''}; }
}
