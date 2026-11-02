/**
 * 연구소 역량·교육 관리 서버 (Google Apps Script)
 * 제작자 KSW · © KSW. All rights reserved.
 *
 * 1) 새 구글 시트를 만들고 [확장 프로그램 → Apps Script]에 이 코드를 그대로 붙여넣기
 * 2) [배포 → 새 배포 → 유형: 웹 앱]
 *      다음 사용자 인증 정보로 실행: 나
 *      액세스 권한이 있는 사용자: 모든 사용자
 * 3) 나온 웹 앱 URL을 사이트(index.html)의 API_URL에 넣기
 *
 * 시트 탭(docs / users / log)은 첫 요청 때 자동으로 만들어져요.
 * 비밀번호·잠금 암호는 이 코드에 적지 않습니다. 사이트에서 처음 만든 관리자가 설정해요.
 * 다른 사람은 사이트에서 회원가입 → 관리자가 '계정 관리'에서 역할을 정해 승인하면 로그인할 수 있어요.
 */

const SHEETS = {
  docs:  ['path', 'data', 'updated', 'by'],
  users: ['id', 'name', 'role', 'hash', 'salt', 'active', 'created', 'lastLogin', 'team', 'rank'],
  log:   ['time', 'user', 'action', 'detail'],
};
// 역할: 연구소장·개발실장 = 부서 전체 역량평가·교육 이수 수정 / 팀장 = 자기 팀(소속) 인원과 그 팀이 쓰는 파트만 / 보기 = 조회만
const ROLES = ['관리자', '연구소장', '개발실장', '팀장', '보기'];
const LEVEL = { '관리자': 4, '연구소장': 3, '개발실장': 3, '편집': 3, '팀장': 2, '보기': 1 };   // '편집'은 예전 이름(연구소장·개발실장과 같음)
const NEED = { '관리자': 4, '편집': 2, '보기': 1 };
const DIRECTOR = ['연구소장', '개발실장', '편집'];
const TOKEN_TTL = 21600;           // 6시간 (CacheService 최대치)
const LOCK_SPIN = 10000;           // 엑셀 잠금 해시 반복 횟수

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(SHEETS[name]);
    s.setFrozenRows(1);
  } else if (s.getLastColumn() < SHEETS[name].length) {
    s.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);  // 새 칸 머리글 자동 추가
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }

function doGet(e) {
  return out_({ ok: true, msg: '연구소 관리 서버 작동 중' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: '잘못된 요청' }); }
  const write = ['put', 'import', 'userSave', 'userDelete', 'changePw', 'setup', 'setLockPw', 'login', 'signup'].indexOf(b.action) >= 0;
  const lock = LockService.getScriptLock();
  if (write) lock.waitLock(20000);
  try {
    switch (b.action) {
      case 'ping':       return out_({ ok: true, setup: needsSetup_() });
      case 'options':    return out_(options_());
      case 'setup':      return out_(setup_(b));
      case 'login':      return out_(login_(b));
      case 'signup':     return out_(signup_(b));
      case 'logout':     CacheService.getScriptCache().remove('t_' + b.token); return out_({ ok: true });
      case 'me':         return out_({ ok: true, user: auth_(b, '보기') });
      case 'load':       auth_(b, '보기'); return out_({ ok: true, docs: loadDocs_() });
      case 'put':        return out_(put_(b));
      case 'import':     return out_(import_(b));
      case 'users':      auth_(b, '관리자'); return out_({ ok: true, users: listUsers_(), lockSet: !!prop_('LOCK_PW') });
      case 'userSave':   return out_(userSave_(b));
      case 'userDelete': return out_(userDelete_(b));
      case 'changePw':   return out_(changePw_(b));
      case 'setLockPw':  return out_(setLockPw_(b));
      case 'lockHash':   return out_(lockHash_(b));
    }
    return out_({ ok: false, err: '알 수 없는 요청' });
  } catch (x) {
    const m = String(x && x.message || x);
    return out_({ ok: false, err: m, auth: m.indexOf('AUTH:') === 0 });
  } finally {
    if (write) lock.releaseLock();
  }
}

/* ---------- 계정 ---------- */
function hash_(pw, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}
function users_() {
  const s = sheet_('users'), v = s.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) if (v[i][0] !== '') out.push({ row: i + 1, id: String(v[i][0]), name: String(v[i][1]), role: String(v[i][2]), hash: String(v[i][3]), salt: String(v[i][4]), active: v[i][5] === true || v[i][5] === 'TRUE', created: String(v[i][6]), lastLogin: String(v[i][7]), team: String(v[i][8] || ''), rank: String(v[i][9] || '') });
  return out;
}
function needsSetup_() { return users_().filter(u => u.role === '관리자' && u.active).length === 0; }
function prop_(k, v) { const p = PropertiesService.getScriptProperties(); if (v === undefined) return p.getProperty(k); p.setProperty(k, v); }
function log_(user, action, detail) {
  try { sheet_('log').appendRow([now_(), user || '', action, String(detail || '').slice(0, 300)]); } catch (x) { }
}
function checkPw_(pw) { if (!pw || String(pw).length < 6) throw new Error('비밀번호는 6자 이상이어야 합니다'); }
function checkId_(id) { if (!/^[A-Za-z0-9._@-]{2,40}$/.test(String(id || ''))) throw new Error('아이디는 영문·숫자 2~40자로 정해 주세요'); }

function setup_(b) {
  if (!needsSetup_()) throw new Error('이미 관리자가 있습니다');
  checkId_(b.id); checkPw_(b.pw);
  const salt = Utilities.getUuid();
  sheet_('users').appendRow([b.id, b.name || b.id, '관리자', hash_(b.pw, salt), salt, true, now_(), '']);
  log_(b.id, 'setup', '최초 관리자 생성');
  return login_(b);
}
function login_(b) {
  const u = users_().filter(x => x.id === String(b.id || ''))[0];
  if (!u || hash_(String(b.pw || ''), u.salt) !== u.hash) { log_(b.id, 'login-fail', ''); return { ok: false, err: '아이디 또는 비밀번호가 맞지 않습니다' }; }
  if (u.role === '대기') return { ok: false, err: '가입 신청을 관리자가 아직 승인하지 않았습니다. 승인 후 로그인할 수 있습니다' };
  if (!u.active) return { ok: false, err: '사용이 중지된 계정입니다. 관리자에게 문의하세요' };
  const token = Utilities.getUuid();
  const info = { id: u.id, name: u.name, role: u.role, team: u.team };
  CacheService.getScriptCache().put('t_' + token, JSON.stringify(info), TOKEN_TTL);
  sheet_('users').getRange(u.row, 8).setValue(now_());
  log_(u.id, 'login', '');
  return { ok: true, token: token, user: info };
}
function auth_(b, need) {
  const raw = b && b.token ? CacheService.getScriptCache().get('t_' + b.token) : null;
  if (!raw) throw new Error('AUTH:로그인이 만료되었습니다. 다시 로그인해 주세요');
  const info = JSON.parse(raw);
  const u = users_().filter(x => x.id === info.id)[0];
  if (!u || !u.active) throw new Error('AUTH:사용이 중지된 계정입니다');
  info.role = u.role; info.name = u.name; info.team = u.team;
  if (!LEVEL[info.role] || LEVEL[info.role] < NEED[need]) throw new Error(need === '관리자' ? '권한이 없습니다 (관리자만 가능)' : '권한이 없습니다');
  return info;
}
function listUsers_() {
  return users_().map(u => ({ id: u.id, name: u.name, team: u.team, rank: u.rank, role: u.role, active: u.active, created: u.created, lastLogin: u.lastLogin }));
}
// 회원가입: '대기' 상태로만 만들어지고, 관리자가 역할을 정해 승인해야 로그인 가능
// 회원가입 화면에 보여 줄 부서·직급 목록 (로그인 전에도 볼 수 있는 이름 목록만)
function options_() {
  const v = sheet_('docs').getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === 'config/main') {
    try { const c = JSON.parse(v[i][1]); return { ok: true, teams: c.teamOrder || [], ranks: c.rankOrder || [] }; } catch (x) { }
  }
  return { ok: true, teams: [], ranks: [] };
}
function signup_(b) {
  if (needsSetup_()) throw new Error('먼저 관리자 계정을 만들어야 합니다');
  checkId_(b.id); checkPw_(b.pw);
  const name = String(b.name || '').trim(), team = String(b.team || '').trim(), rank = String(b.rank || '').trim();
  if (!team) throw new Error('부서를 선택하세요');
  if (!rank) throw new Error('직급을 선택하세요');
  if (!name) throw new Error('이름을 입력하세요');
  const all = users_();
  if (all.some(u => u.id === b.id)) throw new Error('이미 있는 아이디입니다. 다른 아이디를 정해 주세요');
  if (all.filter(u => u.role === '대기').length >= 30) throw new Error('승인 대기 중인 신청이 너무 많습니다. 관리자에게 문의하세요');
  const salt = Utilities.getUuid();
  sheet_('users').appendRow([b.id, name.slice(0, 30), '대기', hash_(b.pw, salt), salt, false, now_(), '', team.slice(0, 30), rank.slice(0, 30)]);
  log_(b.id, 'signup', name + ' ' + team + ' ' + rank);
  return { ok: true };
}
function userSave_(b) {
  const me = auth_(b, '관리자');
  const u = b.user || {};
  checkId_(u.id);
  if (ROLES.indexOf(u.role) < 0 && u.role !== '편집') throw new Error('역할이 올바르지 않습니다');
  if (u.role === '팀장' && !String(u.team || '').trim()) throw new Error('팀장은 소속 팀을 정해야 합니다');
  const s = sheet_('users'), all = users_(), ex = all.filter(x => x.id === u.id)[0];
  if (ex) {
    if (ex.id === me.id && (u.role !== '관리자' || u.active === false)) throw new Error('본인 계정의 관리자 권한은 해제할 수 없습니다');
    s.getRange(ex.row, 2, 1, 1).setValue(u.name || ex.name);
    s.getRange(ex.row, 3, 1, 1).setValue(u.role);
    s.getRange(ex.row, 6, 1, 1).setValue(u.active !== false);
    if (u.team !== undefined) s.getRange(ex.row, 9, 1, 1).setValue(String(u.team || ''));
    if (u.rank !== undefined) s.getRange(ex.row, 10, 1, 1).setValue(String(u.rank || ''));
    if (ex.role === '대기') log_(me.id, 'approve', u.id + ' → ' + u.role);
    if (u.pw) { checkPw_(u.pw); const salt = Utilities.getUuid(); s.getRange(ex.row, 4, 1, 2).setValues([[hash_(u.pw, salt), salt]]); }
    log_(me.id, 'user-update', u.id + ' ' + u.role + (u.pw ? ' 비밀번호 재설정' : '') + (u.active === false ? ' 중지' : ''));
  } else {
    checkPw_(u.pw);
    const salt = Utilities.getUuid();
    s.appendRow([u.id, u.name || u.id, u.role, hash_(u.pw, salt), salt, u.active !== false, now_(), '', String(u.team || ''), String(u.rank || '')]);
    log_(me.id, 'user-add', u.id + ' ' + u.role);
  }
  return { ok: true, users: listUsers_() };
}
function userDelete_(b) {
  const me = auth_(b, '관리자');
  if (b.id === me.id) throw new Error('본인 계정은 삭제할 수 없습니다');
  const ex = users_().filter(x => x.id === b.id)[0];
  if (ex) { sheet_('users').deleteRow(ex.row); log_(me.id, 'user-delete', b.id); }
  return { ok: true, users: listUsers_() };
}
function changePw_(b) {
  const me = auth_(b, '보기');
  const u = users_().filter(x => x.id === me.id)[0];
  if (hash_(String(b.old || ''), u.salt) !== u.hash) throw new Error('현재 비밀번호가 맞지 않습니다');
  checkPw_(b.pw);
  const salt = Utilities.getUuid();
  sheet_('users').getRange(u.row, 4, 1, 2).setValues([[hash_(b.pw, salt), salt]]);
  log_(me.id, 'change-pw', '');
  return { ok: true };
}

/* ---------- 데이터 ---------- */
function loadDocs_() {
  const v = sheet_('docs').getDataRange().getValues(), out = {};
  for (let i = 1; i < v.length; i++) if (v[i][0] !== '') { try { out[v[i][0]] = JSON.parse(v[i][1]); } catch (x) { } }
  return out;
}
const EDIT_PATHS = /^(grades|edu)\/[A-Za-z0-9_-]+$/;
const LEADER_PATHS = /^(grades|edu|parts)\/[A-Za-z0-9_-]+$/;
const PATH_OK = /^(config|parts|people|grades|courses|edu|history)\/[A-Za-z0-9_-]+$/;
function put_(b) {
  const me = auth_(b, '편집');
  const ops = b.ops || [];
  const s = sheet_('docs'), v = s.getDataRange().getValues(), idx = {};
  for (let i = 1; i < v.length; i++) if (v[i][0] !== '') idx[v[i][0]] = i + 1;
  const dels = [];
  let docs = null;
  ops.forEach(op => {
    if (!PATH_OK.test(op.path)) throw new Error('잘못된 경로: ' + op.path);
    if (DIRECTOR.indexOf(me.role) >= 0 && !EDIT_PATHS.test(op.path)) throw new Error('권한이 없습니다 (관리자만 수정 가능)');
    if (me.role === '팀장') {
      if (!LEADER_PATHS.test(op.path)) throw new Error('권한이 없습니다 (관리자만 수정 가능)');
      docs = docs || loadDocs_();
      const id = op.path.split('/')[1], team = String(me.team || '');
      const mine = Object.keys(docs).filter(k => k.indexOf('people/') === 0 && docs[k].team === team).map(k => docs[k]);
      if (op.path.indexOf('parts/') === 0) {
        const cur = docs[op.path];
        if (!cur || op.data === null || (op.data && op.data.name !== cur.name)) throw new Error('파트 추가·삭제·이름 변경은 관리자만 할 수 있습니다');
        if (!mine.some(p => p.cat === cur.name)) throw new Error('우리 팀이 쓰는 파트만 수정할 수 있습니다');
      } else {
        const p = docs['people/' + id];
        if (!p || p.team !== team) throw new Error('우리 팀 인원만 수정할 수 있습니다');
      }
    }
    if (op.data === null) { if (idx[op.path]) dels.push(idx[op.path]); return; }
    const json = JSON.stringify(op.data);
    if (json.length > 45000) throw new Error('데이터가 너무 큽니다: ' + op.path);
    if (idx[op.path]) s.getRange(idx[op.path], 2, 1, 3).setValues([[json, now_(), me.id]]);
    else { s.appendRow([op.path, json, now_(), me.id]); idx[op.path] = s.getLastRow(); }
  });
  dels.sort((a, c) => c - a).forEach(r => s.deleteRow(r));
  log_(me.id, 'put', ops.map(o => (o.data === null ? '-' : '') + o.path).join(', '));
  return { ok: true };
}
function import_(b) {
  const me = auth_(b, '관리자');
  const docs = b.docs || {};
  const rows = Object.keys(docs).filter(p => PATH_OK.test(p)).map(p => [p, JSON.stringify(docs[p]), now_(), me.id]);
  const s = sheet_('docs');
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  if (rows.length) s.getRange(2, 1, rows.length, 4).setValues(rows);
  log_(me.id, 'import', rows.length + '건');
  return { ok: true, count: rows.length };
}

/* ---------- 엑셀 잠금 ---------- */
function setLockPw_(b) {
  const me = auth_(b, '관리자');
  if (!b.pw) throw new Error('잠금 비밀번호를 입력하세요');
  prop_('LOCK_PW', String(b.pw));
  CacheService.getScriptCache().remove('lockHash');
  log_(me.id, 'lock-pw', '엑셀 잠금 비밀번호 변경');
  return { ok: true };
}
// 비밀번호는 서버 밖으로 내보내지 않고, 엑셀 시트 보호용 해시만 계산해서 돌려줌
function lockHash_(b) {
  auth_(b, '관리자');
  const pw = prop_('LOCK_PW');
  if (!pw) return { ok: true, prot: null };
  const c = CacheService.getScriptCache(), hit = c.get('lockHash');
  if (hit) return { ok: true, prot: JSON.parse(hit) };
  const salt = [];
  const u = Utilities.getUuid().replace(/-/g, '');
  for (let i = 0; i < 16; i++) salt.push(sb_(parseInt(u.substr(i * 2, 2), 16)));
  const pwb = [];
  for (let i = 0; i < pw.length; i++) { const ch = pw.charCodeAt(i); pwb.push(sb_(ch & 255), sb_(ch >> 8)); }
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_512, salt.concat(pwb));
  for (let i = 0; i < LOCK_SPIN; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_512, h.concat([sb_(i & 255), sb_((i >>> 8) & 255), sb_((i >>> 16) & 255), sb_((i >>> 24) & 255)]));
  const prot = { salt: Utilities.base64Encode(salt), hash: Utilities.base64Encode(h), spin: LOCK_SPIN };
  c.put('lockHash', JSON.stringify(prot), TOKEN_TTL);
  return { ok: true, prot: prot };
}
function sb_(x) { x = x & 255; return x > 127 ? x - 256 : x; }
