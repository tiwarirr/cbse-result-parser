/* ── SUBJECT NAMES  (source: cbseacademic.nic.in + cbse.gov.in, verified Mar 2026) ── */
const SN = {
  // ── CLASS X / IX  LANGUAGES ──
  '001':'English Elective',   '002':'Hindi Elective (Course-A)', '003':'Urdu Elective',
  '085':'Hindi Course-B',     '122':'Sanskrit',                  '184':'English Lang. & Lit.',

  // ── CLASS X  CORE ACADEMIC ──
  '041':'Mathematics (Standard)', '241':'Mathematics (Basic)',
  '086':'Science',                '087':'Social Science',
  '064':'Home Science',           '076':'NCC',

  // ── CLASS X  SKILL SUBJECTS (400-series, cbseacademic.nic.in) ──
  '401':'Retail',                     '402':'Information Technology',
  '403':'Security',                   '404':'Automotive',
  '405':'Intro to Financial Markets', '406':'Tourism',
  '407':'Beauty & Wellness',          '408':'Agriculture',
  '409':'Food Production',            '410':'Front Office Operations',
  '411':'Banking & Insurance',        '412':'Marketing & Sales',
  '413':'Health Care',                '414':'Apparel',
  '415':'Multimedia',                 '416':'Multi Skill Foundation',
  '417':'Artificial Intelligence',    '418':'Physical Activity Trainer',
  '419':'Data Science',               '420':'Electronics & Hardware',
  '421':'Pharma & Biotechnology',     '422':'Design Thinking',

  // ── CLASS XII  LANGUAGES ──
  '301':'English Core',  '302':'Hindi Core',   '303':'Urdu Core',
  '322':'Sanskrit Core', '118':'French',       '120':'German',
  '104':'Punjabi',       '105':'Bengali',      '106':'Tamil',
  '107':'Telugu',        '108':'Sindhi',       '109':'Marathi',
  '110':'Gujarati',      '112':'Malayalam',    '113':'Odia',
  '114':'Assamese',      '115':'Kannada',

  // ── CLASS XII  ACADEMIC ELECTIVES (careers360 + cbse.gov.in) ──
  '027':'History',                  '028':'Political Science',
  '029':'Geography',                '030':'Economics',
  '031':'Carnatic Music Vocal',     '032':'Carnatic Music Mel. Ins.',
  '033':'Carnatic Perc. (Mridangam)','034':'Hindustani Music Vocal',
  '035':'Hindustani Music Mel. Ins.','036':'Hindustani Perc. Ins.',
  '037':'Psychology',               '039':'Sociology',
  '041':'Mathematics',              '241':'Applied Mathematics',
  '042':'Physics',                  '043':'Chemistry',
  '044':'Biology',                  '045':'Biotechnology',
  '046':'Engineering Graphics',     '048':'Physical Education',
  '049':'Painting',                 '050':'Graphics',
  '051':'Sculpture',                '052':'Applied / Commercial Art',
  '054':'Business Studies',         '055':'Accountancy',
  '056':'Kathak Dance',             '057':'Bharatnatyam Dance',
  '058':'Kuchipudi Dance',          '059':'Odissi Dance',
  '060':'Manipuri Dance',           '061':'Kathakali Dance',
  '065':'Informatics Practices',    '066':'Entrepreneurship',
  '073':'Knowledge Trad. & Practices','074':'Legal Studies',
  '083':'Computer Science',

  // ── CLASS XII  SKILL SUBJECTS (800-series, cbseacademic.nic.in) ──
  '801':'Retail',                        '802':'Information Technology',
  '803':'Web Applications',              '804':'Automotive',
  '805':'Financial Markets Management',  '806':'Tourism',
  '807':'Beauty & Wellness',             '808':'Agriculture',
  '809':'Food Production',               '810':'Front Office Operations',
  '811':'Banking',                       '812':'Marketing',
  '813':'Health Care',                   '814':'Insurance',
  '816':'Horticulture',                  '817':'Typography & Computer Application',
  '818':'Geospatial Technology',         '819':'Electrical Technology',
  '820':'Electronics Technology',        '821':'Multimedia',
  '822':'Taxation',                      '823':'Cost Accounting',
  '824':'Office Procedures & Practices', '825':'Shorthand (English)',
  '826':'Shorthand (Hindi)',             '827':'Air Conditioning & Refrigeration',
  '828':'Medical Diagnostics',           '829':'Textile Design',
  '830':'Design',                        '831':'Salesmanship',
  '833':'Business Administration',       '834':'Food Nutrition & Dietetics',
  '835':'Mass Media Studies',            '836':'Library & Information Science',
  '837':'Fashion Studies',               '841':'Yoga',
  '842':'Early Childhood Care & Edu.',   '843':'Artificial Intelligence',
  '844':'Data Science',                  '845':'Physical Activity Trainer',
  '846':'Land Transportation',           '847':'Electronics & Hardware',
  '848':'Design Thinking & Innovation',
};
const sn = c => SN[c] || `Subj ${c}`;

/* ── STATE ── */
const raw = {X:null, XII:null};
const DB  = {X:[],   XII:[]};
let charts = {};
let activeCls = null;
let activeSec = 'summary';

/* ── UPLOAD ── */
function dragOver(e,c){e.preventDefault();document.getElementById('card-'+c).classList.add('dragover')}
function dragLeave(e,c){document.getElementById('card-'+c).classList.remove('dragover')}
function fileDrop(e,c){e.preventDefault();dragLeave(e,c);readFile(e.dataTransfer.files[0],c)}
function handleFile(e,c){readFile(e.target.files[0],c)}

function detectClass(text){
  const head = text.slice(0,3000);
  if(/SENIOR SCHOOL CERTIFICATE/i.test(head)) return 'XII';
  if(/SECONDARY SCHOOL EXAMINATION/i.test(head)) return 'X';
  return null;
}

function readFile(file, c){
  if(!file) return;
  const r = new FileReader();
  r.onload = ev => {
    const text = ev.target.result;
    const detected = detectClass(text);
    // If file dropped on wrong card, silently re-route to correct class
    const actualClass = detected || c;
    // Clear the other slot if we're re-routing (avoid stale data)
    if(detected && detected !== c){
      raw[c] = null;
      document.getElementById('card-'+c).className = 'upload-card';
      document.getElementById('status-'+c).textContent = '';
    }
    raw[actualClass] = text;
    const el = document.getElementById('status-'+actualClass);
    el.textContent = '✓ ' + file.name + (detected && detected !== c ? '  (auto-detected as Class '+detected+')' : '');
    el.style.color = actualClass==='X' ? 'var(--cx-dk)' : 'var(--cxii-dk)';
    document.getElementById('card-'+actualClass).className = 'upload-card loaded-'+actualClass;
    document.getElementById('btn-analyze').disabled = !(raw.X || raw.XII);
  };
  r.readAsText(file,'UTF-8');
}

/* ── PARSERS ── */
function normalize(t){ return t.replace(/\r\n/g,'\n').replace(/\r/g,'\n'); }

function parseX(text){
  const lines = normalize(text).split('\n'), out = [];
  for(let i=0;i<lines.length;i++){
    const ln = lines[i];
    const m = ln.match(/^(\d{8})\s+(F|M)\s+(.+?)\s{2,}((?:\d{3}\s+){4,6}\d{3})\s+(PASS|COMP|ABST|FAIL)/);
    if(!m) continue;
    const codes = m[4].trim().split(/\s+/);
    const cm = ln.match(/COMP\s+(\d{3})/);
    const toks = [...(lines[i+1]||'').matchAll(/(\d{2,3})\s+(A1|A2|B1|B2|C1|C2|D1|D2|E)\b/g)];
    out.push({rollNo:m[1],gender:m[2],name:m[3].trim(),result:m[5],cls:'X',
      compSub:cm?cm[1]:null, subjects:buildSubs(codes,toks,m[5])});
    i++;
  }
  return out;
}

function parseXII(text){
  const lines = normalize(text).split('\n'), out = [];
  for(let i=0;i<lines.length;i++){
    const ln = lines[i];
    let m = ln.match(/^(\d{8})\s+(F|M)\s+(.+?)\s{2,}((?:\d{3}\s+){4,6}\d{3})\s+(?:A1|A2|B1|B2|C1|C2|D1|D2|E)\s+(?:A1|A2|B1|B2|C1|C2|D1|D2|E)\s+(?:A1|A2|B1|B2|C1|C2|D1|D2|E)\s+(PASS|COMP|ABST|FAIL)/);
    if(!m) m = ln.match(/^(\d{8})\s+(F|M)\s+(.+?)\s{2,}((?:\d{3}\s+){4,6}\d{3})\s+.{0,50}?(ABST|PASS|COMP|FAIL)\b/);
    if(!m) continue;
    const codes = m[4].trim().split(/\s+/);
    const cm = ln.match(/COMP\s+((?:\d{3}[ \t]*)+)/);
    const toks = [...(lines[i+1]||'').matchAll(/(\d{2,3})\s+(A1|A2|B1|B2|C1|C2|D1|D2|E)\b/g)];
    out.push({rollNo:m[1],gender:m[2],name:m[3].trim(),result:m[5],cls:'XII',
      compSub:cm?cm[1].trim():null, subjects:buildSubs(codes,toks,m[5])});
    i++;
  }
  return out;
}

function buildSubs(codes, toks, result){
  if(result==='ABST' && toks.length===0)
    return codes.map(c=>({code:c,name:sn(c),marks:0,grade:'AB'}));
  return codes.slice(0,toks.length).map((c,j)=>({
    code:c, name:sn(c), marks:parseInt(toks[j][1]), grade:toks[j][2]
  }));
}

/* ── RUN ── */
function runAnalysis(){
  if(raw.X)   DB.X   = parseX(raw.X);
  if(raw.XII) DB.XII = parseXII(raw.XII);
  const tot = DB.X.length + DB.XII.length;
  if(!tot){alert('Could not parse any student records. Please check the file format.');return;}

  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  document.getElementById('btn-export').style.display = 'inline-block';

  const meta = parseMeta(raw.X || raw.XII || '');
  // Dynamic school name in header subtitle
  if(meta.school) document.getElementById('hschool').textContent = meta.school;

  document.getElementById('hmeta').innerHTML =
    [DB.X.length   ? `Class X: <strong style="color:var(--gold)">${DB.X.length}</strong>` : '',
     DB.XII.length ? `Class XII: <strong style="color:#22d3ee">${DB.XII.length}</strong>` : '']
    .filter(Boolean).join(' &nbsp;·&nbsp; ') +
    `<br><span style="opacity:.6">CBSE ${meta.year}` +
    (meta.code   ? ` &nbsp;·&nbsp; School ${meta.code}` : '') +
    (meta.region ? ` &nbsp;·&nbsp; ${meta.region}` : '') + `</span>`;

  buildClassTabs();
  const first = DB.X.length ? 'X' : 'XII';
  switchClass(first);

  // Save to localStorage after successful analysis
  saveToLocalStorage();
}

function saveToLocalStorage(){
  const meta = parseMeta(raw.X || raw.XII || '');
  const code = meta.code || 'CBSE';
  const year = meta.year || new Date().getFullYear();
  ['X','XII'].forEach(c=>{
    if(raw[c]) localStorage.setItem(`${code}-${year}-${c}`, raw[c]);
  });
  // Show clear button and backup button
  document.getElementById('btn-clear').style.display = 'inline-block';
  document.getElementById('btn-backup').style.display = 'inline-block';
}

function clearSavedData(){
  // Remove keys for this session
  const meta = parseMeta(raw.X || raw.XII || '');
  const code = meta.code || 'CBSE';
  const year = meta.year || new Date().getFullYear();
  ['X','XII'].forEach(c=>localStorage.removeItem(`${code}-${year}-${c}`));
  location.reload();
}

/* ── BACKUP EXPORT ── */
function exportBackup(){
  const pattern = /^(.+)-(\d{4})-(X|XII)$/;
  // Collect ALL sessions currently in localStorage
  const sessionsMap = {};
  for(let i=0;i<localStorage.length;i++){
    const k = localStorage.key(i);
    const m = k && k.match(pattern);
    if(!m) continue;
    const [,code,year,cls] = m;
    const key = `${code}-${year}`;
    if(!sessionsMap[key]) sessionsMap[key] = {code, year, classes:{}};
    sessionsMap[key].classes[cls] = localStorage.getItem(k);
  }
  if(!Object.keys(sessionsMap).length){
    alert('No saved data found to back up. Please analyse results first.'); return;
  }

  const meta = parseMeta(raw.X || raw.XII || '');
  const backup = {
    version: 1,
    exportedAt: new Date().toISOString(),
    generator: 'CBSE Result Dashboard',
    sessions: Object.values(sessionsMap).map(s=>({
      schoolCode: s.code,
      year: s.year,
      classes: s.classes
    }))
  };

  const blob = new Blob([JSON.stringify(backup, null, 2)], {type:'application/json'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  const school = (meta.school||meta.code||'CBSE').replace(/\s+/g,'_').substring(0,20);
  a.download = `CBSE_Backup_${school}_${meta.year||''}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

/* ── BACKUP IMPORT ── */
let _pendingImport = null; // holds parsed backup data while overlay is open

function handleImportFile(e){
  const file = e.target.files[0];
  if(!file) return;
  // Reset input so same file can be re-selected if needed
  e.target.value = '';
  const r = new FileReader();
  r.onload = ev => {
    let data;
    try { data = JSON.parse(ev.target.result); }
    catch{ alert('Invalid backup file — could not parse JSON.'); return; }

    if(!data.version || !Array.isArray(data.sessions) || !data.sessions.length){
      alert('Invalid backup file — missing sessions data.'); return;
    }
    _pendingImport = data;
    openImportOverlay(data);
  };
  r.readAsText(file,'UTF-8');
}

function openImportOverlay(data){
  const overlay = document.getElementById('import-overlay');
  const container = document.getElementById('import-sessions');
  const loadBtn = document.getElementById('import-btn-load');
  loadBtn.disabled = true;

  // Build session cards
  container.innerHTML = data.sessions.map((s,i)=>{
    const clsList = Object.keys(s.classes);
    const meta = parseMeta(Object.values(s.classes)[0] || '');
    const schoolName = meta.school || s.schoolCode;
    const classBadges = clsList.map(c=>
      `<span style="font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;
        background:${c==='X'?'var(--cx-lt)':'var(--cxii-lt)'};
        color:${c==='X'?'var(--cx-dk)':'var(--cxii-dk)'}">Class ${c}</span>`
    ).join('');
    return `<div class="import-session" data-idx="${i}" onclick="selectImportSession(this)">
      <div style="font-size:28px">🏫</div>
      <div class="import-session-info">
        <div class="import-session-name">${schoolName}</div>
        <div class="import-session-meta">School ${s.schoolCode} &nbsp;·&nbsp; ${s.year}</div>
        <div class="import-session-cls">${classBadges}</div>
      </div>
    </div>`;
  }).join('');

  // Auto-select if only one session
  if(data.sessions.length === 1){
    const card = container.querySelector('.import-session');
    if(card){ card.classList.add('selected'); loadBtn.disabled = false; }
  }

  const d = new Date(data.exportedAt);
  const dateStr = isNaN(d) ? '' : ` · Backed up ${d.toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric'})}`;
  document.getElementById('import-desc').textContent =
    `${data.sessions.length} session${data.sessions.length>1?'s':''} found in this backup${dateStr}. Select one to load.`;

  overlay.classList.add('show');
}

function selectImportSession(el){
  document.querySelectorAll('.import-session').forEach(c=>c.classList.remove('selected'));
  el.classList.add('selected');
  document.getElementById('import-btn-load').disabled = false;
}

function closeImportOverlay(){
  document.getElementById('import-overlay').classList.remove('show');
  _pendingImport = null;
}

function doImport(){
  if(!_pendingImport) return;
  const selected = document.querySelector('.import-session.selected');
  if(!selected){ alert('Please select a session to load.'); return; }
  const idx = parseInt(selected.dataset.idx);
  const session = _pendingImport.sessions[idx];
  if(!session){ alert('Session data not found.'); return; }

  // Write to localStorage
  Object.entries(session.classes).forEach(([cls, text])=>{
    localStorage.setItem(`${session.schoolCode}-${session.year}-${cls}`, text);
  });

  closeImportOverlay();
  // Reload to trigger tryRestoreFromLocalStorage
  location.reload();
}

function tryRestoreFromLocalStorage(){
  // Find any keys matching pattern: {code}-{year}-{X|XII}
  const pattern = /^(.+)-(\d{4})-(X|XII)$/;
  const found = {};
  for(let i=0;i<localStorage.length;i++){
    const k = localStorage.key(i);
    const m = k && k.match(pattern);
    if(m) found[m[3]] = {key:k, code:m[1], year:m[2], text:localStorage.getItem(k)};
  }
  if(!Object.keys(found).length) return;

  // Restore raw data for each class found
  let code='', year='';
  if(found.X)   { raw.X   = found.X.text;   code=found.X.code;   year=found.X.year; }
  if(found.XII) { raw.XII = found.XII.text;  code=found.XII.code; year=found.XII.year; }

  // Show restore banner
  const banner = document.getElementById('restore-banner');
  const msg    = document.getElementById('restore-msg');
  msg.textContent = `✓ Data restored from saved session — School ${code} · ${year}`;
  banner.classList.add('show');

  // Show backup button since we have saved data
  document.getElementById('btn-backup').style.display = 'inline-block';

  // Auto-analyse
  runAnalysis();
}

function parseMeta(t) {
  // Year — from exam title e.g. "EXAMINATION (MAIN)-2025"
  const yr = t.match(/EXAMINATION[^-\n]*[-\u2013](\d{4})/);
  // Region — e.g. "REGION: NOIDA"
  const rg = t.match(/REGION\s*:\s*([A-Z][A-Z0-9 ]+?)(?:\s{2,}|\n|PAGE)/i);
  // School code + name — e.g. "SCHOOL : - 60478   SUDITI GLOBAL ACADEMY..."
  const sc = t.match(/SCHOOL\s*:\s*-\s*(\d+)\s+(.+)/i);
  return {
    year:   yr ? yr[1] : new Date().getFullYear(),
    region: rg ? rg[1].trim() : '',
    code:   sc ? sc[1].trim() : '',
    school: sc ? sc[2].trim() : '',
  };
}

/* ── CLASS TABS ── */
function buildClassTabs(){
  const bar = document.getElementById('class-tabs');
  bar.innerHTML = '';
  ['X','XII'].forEach(c => {
    if(!DB[c].length) return;
    const b = document.createElement('button');
    b.id = 'ctab-'+c;
    b.className = 'ctab';
    b.innerHTML = `Class ${c} <span class="cpill cpill-${c}">${DB[c].length} students</span>`;
    b.onclick = ()=>switchClass(c);
    bar.appendChild(b);
  });
}

function switchClass(cls){
  activeCls = cls;
  document.querySelectorAll('.ctab').forEach(b=>b.classList.remove('ax','axii'));
  const bt = document.getElementById('ctab-'+cls);
  if(bt) bt.classList.add(cls==='X'?'ax':'axii');
  renderSec();
}

/* ── SECTION TABS ── */
function showSec(name, btn){
  activeSec = name;
  document.querySelectorAll('.stab').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  else {
    // sync desktop tabs when triggered from mobile select
    document.querySelectorAll('.stab').forEach(b=>{
      if(b.getAttribute('onclick') && b.getAttribute('onclick').includes("'"+name+"'")) b.classList.add('active');
    });
  }
  // sync mobile select
  const mob = document.getElementById('sec-select-mobile');
  if(mob) mob.value = name;
  document.querySelectorAll('.sec-panel').forEach(p=>p.classList.remove('active'));
  document.getElementById('sec-'+name).classList.add('active');
  renderSec();
}

function renderSec(){
  if(!activeCls) return;
  if(activeSec==='summary')  renderSummary(activeCls);
  if(activeSec==='subjects') renderSubjects(activeCls);
  if(activeSec==='merit')    renderMerit(activeCls);
  if(activeSec==='gender')   renderGender(activeCls);
  if(activeSec==='students') renderStudents(activeCls);
  if(activeSec==='search')   doSearch();
}

/* ── HELPERS ── */
function st(students){
  const pass=students.filter(s=>s.result==='PASS').length;
  const comp=students.filter(s=>s.result==='COMP').length;
  const abst=students.filter(s=>s.result==='ABST').length;
  const n=students.length;
  return {pass,comp,abst,n,pct:n?((pass/n)*100).toFixed(1):'0.0'};
}
function tm(s){ return s.subjects.reduce((a,b)=>a+(b.marks||0),0); }
function avg(a){ return a.length?a.reduce((x,y)=>x+y,0)/a.length:0; }
function pct(n,d){ return d?((n/d)*100).toFixed(1):'0.0'; }
function dc(id){ if(charts[id]){charts[id].destroy();delete charts[id];} }

// Colour palettes per class
const CC = {
  X:  {main:'#c9a84c',dk:'#a07830',boys:'#6b4226',girls:'#c9a84c',
       bar:i=>`hsl(42,${65-i*3}%,${38+i*2}%)`},
  XII:{main:'#0e7490',dk:'#0a5570',boys:'#0a5570',girls:'#22b8d8',
       bar:i=>`hsl(195,${65-i*3}%,${36+i*2}%)`},
};

function sc(lbl,val,det,type=''){
  return `<div class="scard ${type}"><div class="slbl">${lbl}</div>
    <div class="sval">${val}</div><div class="sdet">${det}</div></div>`;
}

function stripe(cls, students){
  const s = st(students);
  const meta = parseMeta(raw[cls]||'');
  const exam = cls==='X'?'All India Secondary School Examination':'Senior School Certificate Examination';
  const detail = [meta.code?`School ${meta.code}`:'', meta.school||'', meta.region?`Region: ${meta.region}`:''].filter(Boolean).join(' · ');
  return `<div class="cls-stripe s${cls}">
    <div class="cs-class">${cls}</div>
    <div class="cs-info">
      <div class="cs-exam">CBSE ${meta.year}</div>
      <div class="cs-name">${exam}</div>
      <div class="cs-detail">${detail}</div>
    </div>
    <div class="cs-space"></div>
    <div class="cs-big">
      <div class="cs-big-num">${s.pct}%</div>
      <div class="cs-big-lbl">Pass Rate</div>
    </div>
  </div>`;
}

/* ══════════════════════════════════════════════════════════════
   SUMMARY
══════════════════════════════════════════════════════════════ */
function renderSummary(cls){
  const sts = DB[cls];
  const s = st(sts);
  const boys  = sts.filter(x=>x.gender==='M');
  const girls = sts.filter(x=>x.gender==='F');
  const sb = st(boys), sg = st(girls);

  // Subject-wise result table — group by subject code across all students
  const subResultMap = {};
  sts.forEach(x=>{
    x.subjects.forEach(sub=>{
      if(!subResultMap[sub.code]) subResultMap[sub.code]={name:sub.name,list:[]};
      subResultMap[sub.code].list.push(x);
    });
  });
  // Deduplicate: each student counted once per subject (some students appear in multiple subject rows)
  // For per-subject pass/comp/absent we check if grade==='E' (failed that subject)
  const subjectRows = {};
  sts.forEach(x=>{
    x.subjects.forEach(sub=>{
      if(!subjectRows[sub.code]) subjectRows[sub.code]={name:sub.name,code:sub.code,total:0,pass:0,comp:0,abst:0};
      const r = subjectRows[sub.code];
      r.total++;
      if(sub.grade==='AB') r.abst++;
      else if(sub.grade==='E') r.comp++;  // failed this subject (Grade E)
      else r.pass++;
    });
  });
  const sRows = Object.values(subjectRows)
    .sort((a,b)=>b.total-a.total)
    .map(r=>{
      const pp = r.total ? ((r.pass/r.total)*100).toFixed(1) : '0.0';
      return `<tr>
        <td><strong>${r.name}</strong></td>
        <td style="font-family:'DM Mono',monospace;color:#bbb;font-size:12px">${r.code}</td>
        <td>${r.total}</td>
        <td><span class="badge bp">${r.pass}</span></td>
        <td><span class="badge bc">${r.comp}</span></td>
        <td><span class="badge ba">${r.abst}</span></td>
        <td style="font-family:'DM Mono',monospace;font-weight:700;color:var(--green)">${pp}%</td>
      </tr>`;
    }).join('');

  // Topper card — best-of-5-with-English scoring
  const engCode = detectEngCode(cls);
  const passPool = sts.filter(s=>s.result==='PASS'||s.result==='COMP');
  const topperData = passPool.map(x=>{
    const sc2 = computeScore(x,{mode:'bestNEng',n:5,engCode});
    return {...x,_pct:sc2.pct};
  }).sort((a,b)=>b._pct-a._pct);
  const topper = topperData[0];
  const topperHtml = topper ? `
    <div class="card" style="border-left:4px solid var(--gold);padding:16px 22px;margin-bottom:16px;display:flex;align-items:center;gap:22px;flex-wrap:wrap">
      <div style="font-size:30px;line-height:1">&#127942;</div>
      <div style="flex:1;min-width:180px">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#aaa;margin-bottom:4px">School Topper &mdash; Class ${cls}</div>
        <div style="font-family:'DM Serif Display',serif;font-size:20px;color:var(--ink)">${topper.name}</div>
        <div style="font-size:12px;color:#bbb;font-family:'DM Mono',monospace;margin-top:2px">Roll ${topper.rollNo} &nbsp;&middot;&nbsp; ${topper.gender==='F'?'Girl':'Boy'}</div>
      </div>
      <div style="text-align:right">
        <div style="font-family:'DM Mono',monospace;font-size:38px;font-weight:300;color:var(--green);line-height:1">${topper._pct.toFixed(2)}%</div>
        <div style="font-size:11px;color:#bbb;text-transform:uppercase;letter-spacing:.08em;margin-top:3px">Best 5 with English</div>
      </div>
    </div>` : '';

  const metaS = parseMeta(raw[cls]||'');
  document.getElementById('d-summary').innerHTML = `
    ${stripe(cls,sts)}
    <div class="stat-grid">
      ${sc('Total Students', s.n, `Class ${cls} \u00b7 CBSE ${metaS.year||''}`, cls==='X'?'cx':'cxii')}
      ${sc('Passed',         s.pass, s.pct+'% pass rate', 'green')}
      ${sc('Compartment',   s.comp, 'Need to clear 1+ sub', 'amber')}
      ${sc('Absent',        s.abst, 'Did not appear', 'red')}
      ${sc('Boys',    sb.n, sb.pct+'% passed', 'blue')}
      ${sc('Girls',   sg.n, sg.pct+'% passed', 'blue')}
    </div>
    ${topperHtml}

    <div class="two-col">
      <div class="card">
        <div class="card-title">Result Distribution — Class ${cls}</div>
        <div class="cwrap"><canvas id="c-sdnt-${cls}"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Boys vs Girls — Pass % — Class ${cls}</div>
        <div class="cwrap"><canvas id="c-sgdr-${cls}"></canvas></div>
      </div>
    </div>
    <div class="two-col">
      <div class="card">
        <div class="card-title">Grade Distribution — Class ${cls} (all subjects)</div>
        <div class="cwrap"><canvas id="c-sgrb-${cls}"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Subject-wise Result Breakdown — Class ${cls}</div>
        <div class="tbl-wrap" style="max-height:260px;overflow-y:auto">
          <table>
            <thead><tr><th>Subject</th><th>Code</th><th>Students</th><th>Pass</th><th>Failed/Comp</th><th>Absent</th><th>Pass %</th></tr></thead>
            <tbody>${sRows}</tbody>
          </table>
        </div>
      </div>
    </div>`;

  // Donut
  dc('c-sdnt-'+cls);
  charts['c-sdnt-'+cls] = new Chart(document.getElementById('c-sdnt-'+cls),{
    type:'doughnut',
    data:{labels:['Pass','Compartment','Absent'],
      datasets:[{data:[s.pass,s.comp,s.abst],
        backgroundColor:['#1a7a4a','#d4800a','#c0392b'],borderWidth:0,hoverOffset:6}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{legend:{position:'bottom',labels:{font:{family:'Nunito',size:12},padding:14}},
        tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${ctx.raw}`}}}}
  });

  // Gender pass bar
  dc('c-sgdr-'+cls);
  charts['c-sgdr-'+cls] = new Chart(document.getElementById('c-sgdr-'+cls),{
    type:'bar',
    data:{labels:['Boys','Girls'],
      datasets:[{data:[parseFloat(sb.pct),parseFloat(sg.pct)],
        backgroundColor:[CC[cls].boys,CC[cls].girls],borderRadius:6,barThickness:50}]},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{min:0,max:100,ticks:{callback:v=>v+'%',font:{family:'DM Mono'}},grid:{color:'#f0ede8'}},
              x:{grid:{display:false},ticks:{font:{family:'Nunito',weight:'700'}}}},
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.raw}%`}}}}
  });

  // Grade bar
  dc('c-sgrb-'+cls);
  const grades=['A1','A2','B1','B2','C1','C2','D1','D2','E'];
  const gc={}; grades.forEach(g=>gc[g]=0);
  sts.forEach(x=>x.subjects.forEach(sub=>{if(gc[sub.grade]!==undefined)gc[sub.grade]++;}));
  const gcols=['#1a7a4a','#2d8a3a','#1a4a7a','#2a5a9a','#8a6a00','#9a7a00','#8a4a00','#9a5a00','#c0392b'];
  charts['c-sgrb-'+cls] = new Chart(document.getElementById('c-sgrb-'+cls),{
    type:'bar',
    data:{labels:grades,datasets:[{data:grades.map(g=>gc[g]),backgroundColor:gcols,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{ticks:{font:{family:'DM Mono'}},grid:{color:'#f0ede8'}},
              x:{grid:{display:false},ticks:{font:{family:'DM Mono',size:12}}}},
      plugins:{legend:{display:false}}}
  });
}

/* ══════════════════════════════════════════════════════════════
   SUBJECTS
══════════════════════════════════════════════════════════════ */
function renderSubjects(cls){
  const sts = DB[cls];
  if(!sts.length){document.getElementById('d-subjects').innerHTML=nodata();return;}

  const sm={};
  const buckets = ['<=40','41-50','51-60','61-70','71-80','81-90','91-94','95-100'];
  function getBucket(m){
    if(m<=40) return '<=40';
    if(m<=50) return '41-50';
    if(m<=60) return '51-60';
    if(m<=70) return '61-70';
    if(m<=80) return '71-80';
    if(m<=90) return '81-90';
    if(m<=94) return '91-94';
    return '95-100';
  }

  sts.forEach(x=>x.subjects.forEach(sub=>{
    if(!sm[sub.code]) sm[sub.code]={code:sub.code,name:sub.name,marks:[],pass:0,n:0,abs:0,grades:{},mb:{}};
    const e=sm[sub.code];
    if(sub.grade!=='AB'){
      e.marks.push(sub.marks); e.n++;
      e.grades[sub.grade]=(e.grades[sub.grade]||0)+1;
      e.mb[getBucket(sub.marks)]=(e.mb[getBucket(sub.marks)]||0)+1;
      if(sub.grade!=='E') e.pass++;
    } else {
      e.abs++;
    }
  }));
  const subs=Object.values(sm).filter(s=>(s.marks.length+s.abs)>0).sort((a,b)=>avg(b.marks)-avg(a.marks));

  document.getElementById('d-subjects').innerHTML=`
    <div class="sec-h">
      <div class="sec-title">Subject Performance — Class ${cls}</div>
      <div class="sec-sub">Average marks, highest score and pass % per subject · ${sts.length} students</div>
    </div>
    <div class="card">
      <div class="card-title">Average Marks by Subject — Class ${cls}</div>
      <div class="cwrap-lg"><canvas id="c-subavg-${cls}"></canvas></div>
    </div>
    <div class="card" style="max-width: 100%; overflow-x: auto;">
      <div class="card-title">Detailed Marks Breakdown Table — Class ${cls}</div>
      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th rowspan="2" style="white-space:nowrap">Subject</th>
              <th rowspan="2" style="white-space:nowrap">Code</th>
              <th rowspan="2" title="Total Registered" style="white-space:nowrap">TOTAL</th>
              <th rowspan="2" title="Absent" style="white-space:nowrap">ABS</th>
              <th rowspan="2" title="Appeared" style="white-space:nowrap">APP</th>
              <th rowspan="2" title="Fail" style="white-space:nowrap">"E"(FAIL)</th>
              <th rowspan="2" title="Pass" style="white-space:nowrap">PASS</th>
              <th rowspan="2" style="white-space:nowrap">PASS%</th>
              <th rowspan="2" style="white-space:nowrap">AVG</th>
              <th rowspan="2" style="white-space:nowrap">MAX</th>
              <th rowspan="2" style="white-space:nowrap">MIN</th>
              <th colspan="8" style="text-align:center;background:var(--paper);color:#888;font-size:11px;letter-spacing:.06em">MARKS DISTRIBUTION</th>
            </tr>
            <tr>${buckets.map(b=>`<th style="white-space:nowrap;font-size:11px;text-align:center;padding:4px 6px">${b}</th>`).join('')}</tr>
          </thead>
          <tbody>${subs.map(s=>{
            const total=s.n+s.abs;
            const a=s.n?avg(s.marks).toFixed(1):'0.0';
            const mx=s.n?Math.max(...s.marks):0;
            const mn=s.n?Math.min(...s.marks):0;
            const fail=s.n-s.pass;
            const pp=s.n?pct(s.pass,s.n):'0.0';
            const ppColor=parseFloat(pp)>=90?'var(--green)':parseFloat(pp)>=75?'#1a7a8a':parseFloat(pp)>=50?'var(--amber)':'var(--red)';
            return `<tr>
              <td style="white-space:nowrap"><strong>${s.name}</strong></td>
              <td style="font-family:'DM Mono',monospace;color:#bbb;font-size:12px;text-align:center">${s.code}</td>
              <td style="font-family:'DM Mono',monospace;text-align:center">${total}</td>
              <td style="font-family:'DM Mono',monospace;text-align:center">${s.abs}</td>
              <td style="font-family:'DM Mono',monospace;text-align:center">${s.n}</td>
              <td style="font-family:'DM Mono',monospace;text-align:center;color:var(--red)">${fail}</td>
              <td style="font-family:'DM Mono',monospace;text-align:center;color:var(--green);font-weight:700">${s.pass}</td>
              <td style="font-family:'DM Mono',monospace;font-weight:700;text-align:center;color:${ppColor}">${pp}%</td>
              <td style="font-family:'DM Mono',monospace;text-align:center">${a}</td>
              <td style="font-family:'DM Mono',monospace;text-align:center">${mx}</td>
              <td style="font-family:'DM Mono',monospace;text-align:center">${mn}</td>
              ${buckets.map(b=>{
                const cnt=s.mb[b]||0;
                return `<td style="font-family:'DM Mono',monospace;font-size:12px;text-align:center;${cnt>0?'font-weight:700':'color:#ddd'}">${cnt||'—'}</td>`;
              }).join('')}
            </tr>`;
          }).join('')}</tbody>
        </table>
      </div>
    </div>`;

  const top=subs; // show all subjects, no artificial cap
  // Dynamic height: 32px per bar, minimum 300px
  const chartH = Math.max(300, top.length * 32);
  const canvasWrap = document.querySelector(`#c-subavg-${cls}`)?.parentElement;
  if(canvasWrap) canvasWrap.style.height = chartH + 'px';
  dc('c-subavg-'+cls);
  charts['c-subavg-'+cls]=new Chart(document.getElementById('c-subavg-'+cls),{
    type:'bar',
    data:{labels:top.map(s=>s.name.length>22?s.name.slice(0,20)+'…':s.name),
      datasets:[{label:'Avg Marks',data:top.map(s=>parseFloat(avg(s.marks).toFixed(1))),
        backgroundColor:top.map((_,i)=>CC[cls].bar(i)),borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      scales:{x:{ticks:{font:{family:'DM Mono'}},grid:{color:'#f0ede8'}},
              y:{ticks:{font:{family:'Nunito',size:12}},grid:{display:false}}},
      plugins:{legend:{display:false}}}
  });
}

/* ══════════════════════════════════════════════════════════════
   MERIT — CBSE dynamic scoring
══════════════════════════════════════════════════════════════ */

// Auto-detect English subject code for a class
function detectEngCode(cls){
  const engCodes = cls==='XII'
    ? ['301','302','001','002','118','120']  // prefer 301=English Core
    : ['184','001','002','085'];              // prefer 184=English Lang&Lit
  const allCodes = new Set();
  (DB[cls]||[]).forEach(s=>s.subjects.forEach(sub=>allCodes.add(sub.code)));
  return engCodes.find(c=>allCodes.has(c)) || '301';
}

// Compute score for a student given settings
function computeScore(s, settings){
  const subs = s.subjects.filter(sub=>sub.grade!=='AB');
  const sumM = arr => arr.reduce((a,b)=>a+b.marks,0);

  if(settings.mode==='all'){
    const total = sumM(subs);
    const maxM  = subs.length * 100;
    return { pct: maxM ? (total/maxM*100) : 0, used: subs.map(s=>s.code) };
  }

  const N = parseInt(settings.n)||5;

  if(settings.mode==='bestN'){
    const sorted = [...subs].sort((a,b)=>b.marks-a.marks);
    const best   = sorted.slice(0,N);
    return { pct: (sumM(best)/(N*100)*100), used: best.map(s=>s.code) };
  }

  if(settings.mode==='bestNEng'){
    const ec  = settings.engCode;
    const eng = subs.find(sub=>sub.code===ec);
    const rest= subs.filter(sub=>sub.code!==ec).sort((a,b)=>b.marks-a.marks);
    const best= eng ? [eng, ...rest.slice(0,N-1)] : rest.slice(0,N);
    return { pct: (sumM(best)/(N*100)*100), used: best.map(s=>s.code) };
  }
  return { pct:0, used:[] };
}

function renderMerit(cls){
  const sts=DB[cls];
  if(!sts.length){document.getElementById('d-merit').innerHTML=nodata();return;}

  const engCode = detectEngCode(cls);

  document.getElementById('d-merit').innerHTML=`
    <div class="sec-h">
      <div class="sec-title">Merit List — Class ${cls}</div>
      <div class="sec-sub">Percentage-based ranking · PASS &amp; COMP students</div>
    </div>

    <!-- Settings Panel -->
    <div class="card" style="margin-bottom:14px;background:#faf8f4;border:1px solid var(--border)">
      <div style="font-family:'DM Serif Display',serif;font-size:14px;color:var(--ink);margin-bottom:12px">⚙ Merit Scoring Settings</div>
      <div style="display:flex;flex-wrap:wrap;gap:16px;align-items:flex-end">
        <div>
          <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#aaa;margin-bottom:6px">Scoring Mode</div>
          <select class="sselect" id="mm-${cls}" onchange="onMeritModeChange('${cls}')">
            <option value="bestNEng">Best N with English</option>
            <option value="bestN">Best N (any subjects)</option>
            <option value="all">All Subjects</option>
          </select>
        </div>
        <div id="mn-wrap-${cls}">
          <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#aaa;margin-bottom:6px">N (count)</div>
          <input type="number" class="sinput" id="mn-${cls}" value="5" min="1" max="9"
            style="width:72px;padding:9px 12px;border-radius:50px"
            oninput="drawMerit('${cls}')">
        </div>
        <div id="me-wrap-${cls}">
          <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#aaa;margin-bottom:6px">English Subject Code</div>
          <input type="text" class="sinput" id="me-${cls}" value="${engCode}" maxlength="4"
            style="width:100px;padding:9px 12px;border-radius:50px;font-family:'DM Mono',monospace"
            oninput="drawMerit('${cls}')">
        </div>
        <div>
          <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#aaa;margin-bottom:6px">Gender</div>
          <select class="sselect" id="mg-${cls}" onchange="drawMerit('${cls}')">
            <option value="all">All Students</option>
            <option value="M">Boys only</option>
            <option value="F">Girls only</option>
          </select>
        </div>
        <div>
          <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#aaa;margin-bottom:6px">Show</div>
          <select class="sselect" id="mc-${cls}" onchange="drawMerit('${cls}')">
            <option value="10">Top 10</option>
            <option value="25">Top 25</option>
            <option value="50">Top 50</option>
            <option value="all">All</option>
          </select>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="tbl-wrap" id="mt-${cls}"></div>
    </div>`;

  drawMerit(cls);
}

function onMeritModeChange(cls){
  const mode = document.getElementById('mm-'+cls)?.value||'bestNEng';
  const nWrap = document.getElementById('mn-wrap-'+cls);
  const eWrap = document.getElementById('me-wrap-'+cls);
  if(nWrap) nWrap.style.display = mode==='all' ? 'none' : '';
  if(eWrap) eWrap.style.display = mode==='bestNEng' ? '' : 'none';
  drawMerit(cls);
}

function drawMerit(cls){
  const gf   = document.getElementById('mg-'+cls)?.value||'all';
  const cf   = document.getElementById('mc-'+cls)?.value||'10';
  const mode = document.getElementById('mm-'+cls)?.value||'bestNEng';
  const n    = document.getElementById('mn-'+cls)?.value||'5';
  const ec   = (document.getElementById('me-'+cls)?.value||'301').trim();

  const settings = { mode, n, engCode: ec };

  // Include PASS and COMP students
  let pool = DB[cls].filter(s=>s.result==='PASS'||s.result==='COMP');
  if(gf!=='all') pool = pool.filter(s=>s.gender===gf);

  // Compute scores
  pool = pool.map(s=>{
    const sc = computeScore(s, settings);
    return { ...s, _pct: sc.pct, _used: sc.used };
  });

  // Sort: % desc, then English marks desc as tie-breaker
  pool.sort((a,b)=>{
    if(Math.abs(b._pct - a._pct) > 0.001) return b._pct - a._pct;
    const engA = a.subjects.find(sub=>sub.code===ec)?.marks||0;
    const engB = b.subjects.find(sub=>sub.code===ec)?.marks||0;
    if(engB!==engA) return engB-engA;
    // final tie-break: highest single subject
    const maxA = Math.max(...a.subjects.map(s=>s.marks||0));
    const maxB = Math.max(...b.subjects.map(s=>s.marks||0));
    return maxB-maxA;
  });

  if(cf!=='all') pool = pool.slice(0, parseInt(cf));

  // Assign ranks (shared rank for tied students; next rank = position after the group)
  pool.forEach((s,i)=>{
    if(i>0 && Math.abs(s._pct-pool[i-1]._pct)<0.001) s._rank=pool[i-1]._rank;
    else s._rank = i+1;
  });

  const resBadge = r => r==='COMP'
    ? `<span class="badge bc" style="font-size:10px;margin-left:5px">COMP</span>`
    : `<span class="badge bp" style="font-size:10px;margin-left:5px">PASS</span>`;

  document.getElementById('mt-'+cls).innerHTML=`
    <table>
      <thead><tr>
        <th>Rank</th><th>Name</th><th>Roll No</th>
        <th>Gender</th><th>Result</th><th style="text-align:right">Score %</th><th>Subject Marks</th>
      </tr></thead>
      <tbody>${pool.length ? pool.map(s=>{
        const r=s._rank, bc=r===1?'r1':r===2?'r2':r===3?'r3':'';
        const usedSet = new Set(s._used);
        const subHtml = s.subjects.map(sub=>{
          const inScore = usedSet.has(sub.code);
          const style = inScore
            ? `font-weight:800;border:2px solid currentColor`
            : `opacity:0.35`;
          return `<span class="gr g-${sub.grade.toLowerCase()}" style="${style}">${sub.code}: ${sub.marks}</span>`;
        }).join(' ');
        return `<tr>
          <td><span class="rnk ${bc}">${r}</span></td>
          <td><strong>${s.name}</strong></td>
          <td style="font-family:'DM Mono',monospace;font-size:12px;color:#bbb">${s.rollNo}</td>
          <td>${s.gender==='F'?'👩 Girl':'👦 Boy'}</td>
          <td>${resBadge(s.result)}</td>
          <td style="font-family:'DM Mono',monospace;font-size:18px;font-weight:700;text-align:right;color:var(--green)">${s._pct.toFixed(2)}%</td>
          <td>${subHtml}</td>
        </tr>`;
      }).join('') : '<tr><td colspan="7" style="text-align:center;padding:32px;color:#ccc">No students found</td></tr>'}
      </tbody>
    </table>`;
}

/* ══════════════════════════════════════════════════════════════
   GENDER
══════════════════════════════════════════════════════════════ */
function renderGender(cls){
  const sts=DB[cls];
  if(!sts.length){document.getElementById('d-gender').innerHTML=nodata();return;}
  const boys=sts.filter(s=>s.gender==='M');
  const girls=sts.filter(s=>s.gender==='F');
  const sb=st(boys), sg=st(girls);

  document.getElementById('d-gender').innerHTML=`
    <div class="sec-h">
      <div class="sec-title">Gender Analysis — Class ${cls}</div>
      <div class="sec-sub">Comparative performance between boys and girls</div>
    </div>
    <div class="gstrip">
      <div class="gcard">
        <div class="gicon">👦</div>
        <div><div class="gname">Boys</div><div class="gcnt">${sb.n}</div>
          <div class="gpct">✓ ${sb.pct}% passed</div></div>
      </div>
      <div class="gcard">
        <div class="gicon">👩</div>
        <div><div class="gname">Girls</div><div class="gcnt">${sg.n}</div>
          <div class="gpct">✓ ${sg.pct}% passed</div></div>
      </div>
    </div>
    <div class="two-col">
      <div class="card">
        <div class="card-title">Pass % — Boys vs Girls — Class ${cls}</div>
        <div class="cwrap"><canvas id="c-gp-${cls}"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Average Total Marks — Class ${cls}</div>
        <div class="cwrap"><canvas id="c-ga-${cls}"></canvas></div>
      </div>
    </div>
    <div class="card">
      <div class="card-title">Subject-wise Average — Boys vs Girls — Class ${cls}</div>
      <div class="cwrap-lg"><canvas id="c-gs-${cls}"></canvas></div>
    </div>
    <div class="two-col">
      <div class="card">
        <div class="card-title">Grade Distribution — Boys — Class ${cls}</div>
        <div class="cwrap"><canvas id="c-ggb-${cls}"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Grade Distribution — Girls — Class ${cls}</div>
        <div class="cwrap"><canvas id="c-ggg-${cls}"></canvas></div>
      </div>
    </div>`;

  // pass %
  dc('c-gp-'+cls);
  charts['c-gp-'+cls]=new Chart(document.getElementById('c-gp-'+cls),{
    type:'bar',
    data:{labels:['Boys','Girls'],
      datasets:[{data:[parseFloat(sb.pct),parseFloat(sg.pct)],
        backgroundColor:[CC[cls].boys,CC[cls].girls],borderRadius:6,barThickness:50}]},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{min:0,max:100,ticks:{callback:v=>v+'%',font:{family:'DM Mono'}},grid:{color:'#f0ede8'}},
              x:{grid:{display:false},ticks:{font:{family:'Nunito',weight:'700'}}}},
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.raw}%`}}}}
  });

  // avg marks
  const ba=boys.filter(s=>s.result==='PASS').map(s=>tm(s));
  const ga=girls.filter(s=>s.result==='PASS').map(s=>tm(s));
  dc('c-ga-'+cls);
  charts['c-ga-'+cls]=new Chart(document.getElementById('c-ga-'+cls),{
    type:'bar',
    data:{labels:['Boys','Girls'],
      datasets:[{data:[parseFloat(avg(ba).toFixed(1)),parseFloat(avg(ga).toFixed(1))],
        backgroundColor:[CC[cls].boys,CC[cls].girls],borderRadius:6,barThickness:50}]},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{ticks:{font:{family:'DM Mono'}},grid:{color:'#f0ede8'}},
              x:{grid:{display:false},ticks:{font:{family:'Nunito',weight:'700'}}}},
      plugins:{legend:{display:false}}}
  });

  // subject comparison
  const subm={};
  sts.forEach(x=>x.subjects.forEach(sub=>{
    if(!subm[sub.code]) subm[sub.code]={name:sub.name,boys:[],girls:[]};
    if(sub.grade!=='AB'){
      if(x.gender==='M') subm[sub.code].boys.push(sub.marks);
      else subm[sub.code].girls.push(sub.marks);
    }
  }));
  const tsubs=Object.values(subm).filter(s=>s.boys.length>1&&s.girls.length>1)
    .sort((a,b)=>(b.boys.length+b.girls.length)-(a.boys.length+a.girls.length));
  // Dynamic height for gender subject chart
  const gChartH = Math.max(300, tsubs.length * 32);
  const gCanvasWrap = document.querySelector(`#c-gs-${cls}`)?.parentElement;
  if(gCanvasWrap) gCanvasWrap.style.height = gChartH + 'px';
  dc('c-gs-'+cls);
  charts['c-gs-'+cls]=new Chart(document.getElementById('c-gs-'+cls),{
    type:'bar',
    data:{labels:tsubs.map(s=>s.name.length>16?s.name.slice(0,14)+'…':s.name),
      datasets:[
        {label:'Boys Avg',data:tsubs.map(s=>parseFloat(avg(s.boys).toFixed(1))),backgroundColor:CC[cls].boys,borderRadius:3},
        {label:'Girls Avg',data:tsubs.map(s=>parseFloat(avg(s.girls).toFixed(1))),backgroundColor:CC[cls].girls,borderRadius:3}
      ]},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{ticks:{font:{family:'DM Mono'}},grid:{color:'#f0ede8'}},
              x:{grid:{display:false},ticks:{font:{family:'Nunito',size:11}}}},
      plugins:{legend:{position:'bottom',labels:{font:{family:'Nunito',size:12},padding:12}}}}
  });

  // grade distributions
  const grades=['A1','A2','B1','B2','C1','C2','D1','D2','E'];
  const gcols=['#1a7a4a','#2d8a3a','#1a4a7a','#2a5a9a','#8a6a00','#9a7a00','#8a4a00','#9a5a00','#c0392b'];
  [['b',boys,'c-ggb-'],['g',girls,'c-ggg-']].forEach(([,pool,id])=>{
    const gc={}; grades.forEach(g=>gc[g]=0);
    pool.forEach(x=>x.subjects.forEach(sub=>{if(gc[sub.grade]!==undefined)gc[sub.grade]++;}));
    dc(id+cls);
    charts[id+cls]=new Chart(document.getElementById(id+cls),{
      type:'bar',
      data:{labels:grades,datasets:[{data:grades.map(g=>gc[g]),backgroundColor:gcols,borderRadius:3}]},
      options:{responsive:true,maintainAspectRatio:false,
        scales:{y:{ticks:{font:{family:'DM Mono'}},grid:{color:'#f0ede8'}},
                x:{grid:{display:false},ticks:{font:{family:'DM Mono',size:11}}}},
        plugins:{legend:{display:false}}}
    });
  });
}

/* ══════════════════════════════════════════════════════════════
   SEARCH  (both classes, results grouped)
══════════════════════════════════════════════════════════════ */
function nodata(){ return '<div class="nodata">No data loaded for this class</div>'; }

/* ══════════════════════════════════════════════════════════════
   ALL STUDENTS TABLE
══════════════════════════════════════════════════════════════ */
function renderStudents(cls){
  const sts = DB[cls];
  if(!sts.length){document.getElementById('d-students').innerHTML=nodata();return;}

  // collect all subject codes in order of first appearance
  const subCols = [];
  const seenCodes = new Set();
  sts.forEach(s=>s.subjects.forEach(sub=>{
    if(!seenCodes.has(sub.code)){ seenCodes.add(sub.code); subCols.push({code:sub.code,name:sub.name}); }
  }));

  document.getElementById('d-students').innerHTML=`
    <div class="sec-h">
      <div class="sec-title">All Students — Class ${cls}</div>
      <div class="sec-sub">${sts.length} students · all subjects · click column headers to sort</div>
    </div>
    <div class="sbar">
      <input class="sinput" id="stf-q-${cls}" type="text" placeholder="🔍  Search by name or roll number…"
        oninput="drawStudents('${cls}')">
      <select class="sselect" id="stf-res-${cls}" onchange="drawStudents('${cls}')">
        <option value="all">All Results</option>
        <option value="PASS">Pass</option>
        <option value="COMP">Compartment</option>
        <option value="ABST">Absent</option>
        <option value="FAIL">Fail</option>
      </select>
      <select class="sselect" id="stf-gen-${cls}" onchange="drawStudents('${cls}')">
        <option value="all">All Genders</option>
        <option value="M">Boys</option>
        <option value="F">Girls</option>
      </select>
      <button class="btn-export" style="display:inline-block;padding:10px 20px;font-size:13px"
        onclick="exportCurrentView('${cls}')">⬇ Export Excel</button>
    </div>
    <div id="st-cnt-${cls}" style="font-size:12px;color:#aaa;font-family:'DM Mono',monospace;margin-bottom:8px"></div>
    <div class="card" style="padding:0">
      <div class="tbl-wrap" id="st-tbl-${cls}"></div>
    </div>`;

  drawStudents(cls);
}

// sort state per class
const sortState  = {};
// current filtered+sorted pool per class — used by exportCurrentView
const currentPool = {};

function drawStudents(cls){
  const resF = document.getElementById('stf-res-'+cls)?.value || 'all';
  const genF = document.getElementById('stf-gen-'+cls)?.value || 'all';
  const q    = (document.getElementById('stf-q-'+cls)?.value || '').toLowerCase().trim();

  // collect subject cols
  const subCols = [];
  const seenCodes = new Set();
  DB[cls].forEach(s=>s.subjects.forEach(sub=>{
    if(!seenCodes.has(sub.code)){ seenCodes.add(sub.code); subCols.push({code:sub.code,name:sub.name}); }
  }));

  let pool = DB[cls]
    .filter(s=> resF==='all' || s.result===resF)
    .filter(s=> genF==='all' || s.gender===genF)
    .filter(s=> !q || s.name.toLowerCase().includes(q) || s.rollNo.includes(q));

  // sorting
  const ss = sortState[cls] || {col:'total', dir:-1};
  pool = pool.slice().sort((a,b)=>{
    let av, bv;
    if(ss.col==='name')   { av=a.name; bv=b.name; return ss.dir*(av<bv?-1:av>bv?1:0); }
    if(ss.col==='roll')   { av=a.rollNo; bv=b.rollNo; return ss.dir*(av<bv?-1:av>bv?1:0); }
    if(ss.col==='result') { av=a.result; bv=b.result; return ss.dir*(av<bv?-1:av>bv?1:0); }
    if(ss.col==='total')  { av=tm(a); bv=tm(b); return ss.dir*(av-bv); }
    if(ss.col==='pct')    {
      const aS=a.subjects.filter(x=>x.grade!=='AB').length;
      const bS=b.subjects.filter(x=>x.grade!=='AB').length;
      av=aS?(tm(a)/(aS*100)*100):0; bv=bS?(tm(b)/(bS*100)*100):0;
      return ss.dir*(av-bv);
    }
    const ai=a.subjects.find(x=>x.code===ss.col); const bi=b.subjects.find(x=>x.code===ss.col);
    av=ai&&ai.grade!=='AB'?ai.marks:-1; bv=bi&&bi.grade!=='AB'?bi.marks:-1;
    return ss.dir*(av-bv);
  });

  // store for export
  currentPool[cls] = {pool, subCols};

  // count label
  const total = DB[cls].length;
  const cntEl = document.getElementById('st-cnt-'+cls);
  if(cntEl){
    const isFiltered = pool.length < total;
    cntEl.textContent = isFiltered
      ? `Showing ${pool.length} of ${total} students`
      : `${total} students`;
  }

  function sortBtn(col, label){
    const active = (sortState[cls]||{}).col===col;
    const dir = active ? (sortState[cls].dir===-1?'↓':'↑') : '';
    return `<span style="cursor:pointer;user-select:none;white-space:nowrap" onclick="setSortStudents('${cls}','${col}')">${label}${active?` <span style="color:var(--${cls==='X'?'cx':'cxii'})">${dir}</span>`:'  <span style="opacity:.25">↕</span>'}</span>`;
  }

  const resColors = {PASS:'var(--green)',COMP:'var(--amber)',ABST:'#999',FAIL:'var(--red)'};

  const html = `<table>
    <thead>
      <tr>
        <th style="text-align:center;width:44px">#</th>
        <th>${sortBtn('roll','Roll No')}</th>
        <th>${sortBtn('name','Name')}</th>
        <th style="text-align:center">Gender</th>
        <th style="text-align:center">${sortBtn('result','Result')}</th>
        <th style="text-align:center">${sortBtn('total','Total')}</th>
        <th style="text-align:center">${sortBtn('pct','%')}</th>
        ${subCols.map(sc=>`<th title="${sc.name}" style="text-align:center;min-width:72px;white-space:nowrap">
          <div style="font-size:10px;font-weight:600;letter-spacing:.04em;color:#aaa;margin-bottom:2px">${sc.code}</div>
          <div style="font-size:10px;opacity:.7;margin-bottom:4px;max-width:70px;overflow:hidden;text-overflow:ellipsis">${sc.name.length>10?sc.name.slice(0,9)+'…':sc.name}</div>
          ${sortBtn(sc.code, '↕')}
        </th>`).join('')}
      </tr>
    </thead>
    <tbody>
      ${pool.length ? pool.map((s,i)=>{
        const total = tm(s);
        const validSubs = s.subjects.filter(x=>x.grade!=='AB').length;
        const pctVal = validSubs ? ((total/(validSubs*100))*100).toFixed(1) : null;
        const resultBadge = `<span style="font-size:11px;font-weight:700;color:${resColors[s.result]||'#888'}">${s.result}</span>`;
        const compNote = s.compSub ? `<div style="font-size:10px;color:var(--amber);margin-top:2px">Comp: ${s.compSub}</div>` : '';
        return `<tr>
          <td style="text-align:center;font-family:'DM Mono',monospace;font-size:12px;color:#ccc">${i+1}</td>
          <td style="font-family:'DM Mono',monospace;font-size:12px;color:#bbb">${s.rollNo}</td>
          <td><strong>${s.name}</strong></td>
          <td style="text-align:center">${s.gender==='F'?'👩':'👦'}</td>
          <td style="text-align:center">${resultBadge}${compNote}</td>
          <td style="font-family:'DM Mono',monospace;font-size:18px;font-weight:600;text-align:center">${total}</td>
          <td style="font-family:'DM Mono',monospace;font-size:14px;font-weight:700;text-align:center;color:var(--green)">${pctVal !== null ? pctVal+'%' : '&mdash;'}</td>
          ${subCols.map(sc=>{
            const sub = s.subjects.find(x=>x.code===sc.code);
            if(!sub) return `<td style="color:#eee;text-align:center">—</td>`;
            if(sub.grade==='AB') return `<td style="text-align:center;color:#ccc;font-size:11px">AB</td>`;
            return `<td style="text-align:center">
              <div style="font-family:'DM Mono',monospace;font-weight:700">${sub.marks}</div>
              <div><span class="gr g-${sub.grade.toLowerCase()}" style="font-size:10px">${sub.grade}</span></div>
            </td>`;
          }).join('')}
        </tr>`;
      }).join('') : `<tr><td colspan="${7+subCols.length}" style="text-align:center;padding:40px;color:#ccc">No students match the current filters</td></tr>`}
    </tbody>
  </table>`;

  document.getElementById('st-tbl-'+cls).innerHTML = html;
}

function setSortStudents(cls, col){
  const cur = sortState[cls] || {col:'total', dir:-1};
  sortState[cls] = { col, dir: cur.col===col ? -cur.dir : -1 };
  drawStudents(cls);
}

/* ══════════════════════════════════════════════════════════════
   EXPORT CURRENT VIEW  —  exports the filtered+sorted table
   as it appears on screen right now
══════════════════════════════════════════════════════════════ */
function exportCurrentView(cls){
  if(!window.XLSX){ alert('Excel library not loaded yet — please wait a moment and try again.'); return; }
  if(!DB[cls] || !DB[cls].length){ alert('No data to export.'); return; }

  // Use the currently filtered+sorted pool (exactly what's on screen)
  const poolData = currentPool[cls];
  const exportPool = poolData ? poolData.pool : DB[cls];
  const subCols    = poolData ? poolData.subCols : (()=>{
    const seen=new Set(), cols=[];
    DB[cls].forEach(s=>s.subjects.forEach(sub=>{if(!seen.has(sub.code)){seen.add(sub.code);cols.push({code:sub.code,name:sub.name});}}));
    return cols;
  })();

  if(!exportPool.length){ alert('No students match the current filter — nothing to export.'); return; }

  const meta    = parseMeta(raw[cls] || '');
  const GOLD    = 'FFC9A84C', TEAL='FF0E7490', DARK='FF1A1A2E';
  const PASS_C  = 'FF1A7A4A', FAIL_C='FFC0392B', ABST_C='FF888888', COMP_C='FFD4800A';
  const RC = r  => r==='PASS'?PASS_C:r==='COMP'?COMP_C:r==='ABST'?ABST_C:FAIL_C;
  const H  = bg => ({
    font:{bold:true,color:{rgb:'FFFFFFFF'}}, fill:{fgColor:{rgb:bg}},
    alignment:{horizontal:'center',vertical:'center',wrapText:true}
  });

  // Assign rank only to PASS students within the current pool (in their current order)
  let rankCounter = 0;
  const ranks = exportPool.map(s => s.result==='PASS' ? ++rankCounter : '');

  // ── headers ──
  const fixHdrs = ['Rank','Roll No','Name','Gender','Result','Comp Subject','Total Marks','%'];
  const subHdrs = subCols.flatMap(s => [`${s.name} (${s.code})`, 'Grade']);
  const headers = [...fixHdrs, ...subHdrs];

  // ── data rows ──
  const rows = exportPool.map((s, i) => {
    const total = tm(s);
    const validSubs = s.subjects.filter(x=>x.grade!=='AB').length;
    const pctStr = validSubs ? ((total/(validSubs*100))*100).toFixed(1)+'%' : '';
    const fix = [
      ranks[i],
      s.rollNo,
      s.name,
      s.gender==='F' ? 'Female' : 'Male',
      s.result,
      s.compSub || '',
      total,
      pctStr
    ];
    const sdata = subCols.flatMap(sc => {
      const f = s.subjects.find(x => x.code===sc.code);
      return f ? [f.grade==='AB' ? 'AB' : f.marks, f.grade] : ['', ''];
    });
    return [...fix, ...sdata];
  });

  // ── worksheet ──
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

  ws['!cols'] = [
    {wch:6}, {wch:12}, {wch:28}, {wch:9}, {wch:10}, {wch:22}, {wch:13}, {wch:8},
    ...subCols.flatMap(() => [{wch:24}, {wch:8}])
  ];

  // ── header styles ──
  const hdrBg = cls==='X' ? GOLD : TEAL;
  headers.forEach((_, ci) => {
    const a = XLSX.utils.encode_cell({r:0, c:ci});
    if(ws[a]) ws[a].s = H(ci < fixHdrs.length ? DARK : hdrBg);
  });

  // ── data row styles ──
  const MEDAL = ['FFC9A84C','FFCCCCCC','FFCD7F32'];
  rows.forEach((row, ri) => {
    const res  = row[4];
    const rc   = RC(res);
    const rank = row[0];
    const rIdx = ri + 1;
    const medal = typeof rank==='number' && rank<=3 ? MEDAL[rank-1] : null;

    headers.forEach((_, ci) => {
      const a = XLSX.utils.encode_cell({r:rIdx, c:ci});
      if(!ws[a]) return;
      if(medal){
        const base = {fill:{fgColor:{rgb:medal}}, font:{bold:rank===1, color:{rgb:DARK}}};
        ws[a].s = {...base, alignment:{horizontal: ci===0||ci===6||ci===7 ? 'center':'left'}};
        return;
      }
      if(ci===0)  ws[a].s = {alignment:{horizontal:'center'}, font:{color:{rgb:'FFBBBBBB'}}};
      else if(ci===1) ws[a].s = {font:{color:{rgb:'FFAAAAAA'}}, alignment:{horizontal:'center'}};
      else if(ci===2) ws[a].s = {font:{bold:true}};
      else if(ci===3) ws[a].s = {alignment:{horizontal:'center'}};
      else if(ci===4) ws[a].s = {font:{bold:true, color:{rgb:rc}}, alignment:{horizontal:'center'}};
      else if(ci===5 && row[5]) ws[a].s = {font:{color:{rgb:COMP_C}, italic:true}};
      else if(ci===6) ws[a].s = {font:{bold:true}, alignment:{horizontal:'center'}};
      else if(ci===7) ws[a].s = {font:{bold:true, color:{rgb:'FF1A7A4A'}}, alignment:{horizontal:'center'}};
      else if(ci>=8 && ci%2===0) ws[a].s = {font:{bold:true}, alignment:{horizontal:'center'}};
      else if(ci>=8 && ci%2===1) ws[a].s = {font:{color:{rgb:rc}}, alignment:{horizontal:'center'}};
    });
  });

  // ── download ──
  const school = (meta.school||'CBSE').replace(/\s+/g,'_').substring(0,22);
  const isFiltered = exportPool.length < DB[cls].length;
  XLSX.utils.book_append_sheet(wb, ws, `Class ${cls}`);
  XLSX.writeFile(wb, `CBSE_Class${cls}_${isFiltered?'Filtered_':''}${school}_${meta.year||''}.xlsx`);
}
function exportExcel(){
  const meta = parseMeta(raw.X || raw.XII || '');
  const wb   = XLSX.utils.book_new();

  /* ─── shared constants ─── */
  const GRADES = ['A1','A2','B1','B2','C1','C2','D1','D2','E'];
  const GOLD='FFC9A84C', TEAL='FF0E7490', DARK='FF1A1A2E';
  const LGOLD='FFFFF8E8', LTEAL='FFE0F4F8', LGREY='FFF5F0E8';
  const PASS_C='FF1A7A4A', FAIL_C='FFC0392B', ABST_C='FF888888', COMP_C='FFD4800A';
  const GREEN_FILL='FFD6F0E0', AMBER_FILL='FFFFF0D6', RED_FILL='FFFCE8E8', BLUE_FILL='FFE8F0FA';

  /* ─── style helpers ─── */
  const H  = (bg,fg='FFFFFFFF') => ({
    font:{bold:true,color:{rgb:fg}}, fill:{fgColor:{rgb:bg}},
    alignment:{horizontal:'center',vertical:'center',wrapText:true},
    border:{bottom:{style:'medium',color:{rgb:'FFCCCCCC'}}}
  });
  const SH = (bg,fg=DARK) => ({         // section heading style
    font:{bold:true,sz:12,color:{rgb:fg}}, fill:{fgColor:{rgb:bg}},
    alignment:{horizontal:'left',vertical:'center'}
  });
  const C  = (bold=false,center=false) => ({
    font:{bold}, alignment:{horizontal:center?'center':'left',vertical:'center'}
  });
  const RC = r => r==='PASS'?PASS_C:r==='COMP'?COMP_C:r==='ABST'?ABST_C:FAIL_C;
  const PC = p => p>=90?PASS_C:p>=75?'FF1A7A8A':p>=50?COMP_C:FAIL_C;

  /* ─── data helpers ─── */
  const xavg  = arr => arr.length ? Math.round(arr.reduce((a,b)=>a+b,0)/arr.length*10)/10 : 0;
  const xpct  = (n,d) => d ? Math.round(n/d*1000)/10 : 0;
  const xtm   = s => s.subjects.filter(x=>x.grade!=='AB').reduce((a,b)=>a+b.marks,0);
  const xst   = arr => {
    const pass=arr.filter(s=>s.result==='PASS').length;
    const comp=arr.filter(s=>s.result==='COMP').length;
    const abst=arr.filter(s=>s.result==='ABST').length;
    const fail=arr.filter(s=>s.result==='FAIL').length;
    return {n:arr.length, pass, comp, abst, fail, pct:xpct(pass,arr.length)};
  };
  const buildSubMap = sts => {
    const sm={};
    sts.forEach(s=>s.subjects.forEach(sub=>{
      if(!sm[sub.code]) sm[sub.code]={code:sub.code,name:sub.name,
        marks:[],marksB:[],marksG:[],pass:0,passB:0,passG:0,n:0,nB:0,nG:0,grades:{}};
      const e=sm[sub.code];
      if(sub.grade!=='AB'){
        e.marks.push(sub.marks); e.n++;
        e.grades[sub.grade]=(e.grades[sub.grade]||0)+1;
        if(sub.grade!=='E') e.pass++;
        if(s.gender==='M'){e.marksB.push(sub.marks);e.nB++;if(sub.grade!=='E')e.passB++;}
        else              {e.marksG.push(sub.marks);e.nG++;if(sub.grade!=='E')e.passG++;}
      }
    }));
    return Object.values(sm).filter(s=>s.marks.length>0).sort((a,b)=>xavg(b.marks)-xavg(a.marks));
  };
  const getSubjectCols = sts => {
    const seen=new Map();
    sts.forEach(s=>s.subjects.forEach(sub=>{if(!seen.has(sub.code))seen.set(sub.code,sub.name);}));
    return [...seen.entries()].map(([code,name])=>({code,name}));
  };

  /* ─── write helpers ─── */
  function styleRange(ws, r0, c0, r1, c1, styleFn){
    for(let r=r0;r<=r1;r++) for(let c=c0;c<=c1;c++){
      const a=XLSX.utils.encode_cell({r,c});
      if(ws[a]) ws[a].s=styleFn(r-r0,c-c0,ws[a].v);
    }
  }
  function addSection(data, title, bg){
    // Returns rows: one blank separator, one bold title row, then data rows
    return [[], [title,...Array(data[0]?.length-1||0).fill('')], ...data];
  }

  /* ════════════════════════════════════════════════════════
     SHEET 1 — SUMMARY  (mirrors Summary tab)
  ════════════════════════════════════════════════════════ */
  {
    const rows=[];
    // ── title block ──
    rows.push(['CBSE RESULT SUMMARY — '+meta.year]);
    rows.push([`School: ${meta.school||'—'}   |   Code: ${meta.code||'—'}   |   Region: ${meta.region||'—'}`]);
    rows.push([]);

    // ── overall results ──
    rows.push(['OVERALL RESULT','Total','Pass','Compartment','Absent','Fail','Pass %']);
    ['X','XII'].forEach(cls=>{
      const s=xst(DB[cls]); if(!s.n) return;
      rows.push([`Class ${cls}`,s.n,s.pass,s.comp,s.abst,s.fail,s.pct+'%']);
    });
    if(DB.X.length && DB.XII.length){
      const all=[...DB.X,...DB.XII], s=xst(all);
      rows.push(['Combined',s.n,s.pass,s.comp,s.abst,s.fail,s.pct+'%']);
    }
    rows.push([]);

    // ── gender summary ──
    rows.push(['GENDER SUMMARY','Total','Pass','Compartment','Absent','Pass %']);
    ['X','XII'].forEach(cls=>{
      const sts=DB[cls]; if(!sts.length) return;
      const boys=sts.filter(s=>s.gender==='M'), girls=sts.filter(s=>s.gender==='F');
      const sb=xst(boys), sg=xst(girls);
      rows.push([`Class ${cls} — Boys`, sb.n, sb.pass, sb.comp, sb.abst, sb.pct+'%']);
      rows.push([`Class ${cls} — Girls`,sg.n, sg.pass, sg.comp, sg.abst, sg.pct+'%']);
    });
    rows.push([]);

    // ── grade distribution ──
    rows.push(['GRADE DISTRIBUTION (all subjects combined)', ...GRADES]);
    ['X','XII'].forEach(cls=>{
      const sts=DB[cls]; if(!sts.length) return;
      const gc={}; GRADES.forEach(g=>gc[g]=0);
      sts.forEach(x=>x.subjects.forEach(s=>{if(gc[s.grade]!==undefined)gc[s.grade]++;}));
      rows.push([`Class ${cls}`, ...GRADES.map(g=>gc[g])]);
    });
    rows.push([]);

    // ── subject-wise result (mirrors Summary tab table) ──
    rows.push(['SUBJECT-WISE RESULT BREAKDOWN','Code','Students','Pass','Failed/Fail in Sub','Absent','Pass %']);
    ['X','XII'].forEach(cls=>{
      const sts=DB[cls]; if(!sts.length) return;
      rows.push([`— Class ${cls} —`]);
      const subjectRows={};
      sts.forEach(x=>x.subjects.forEach(sub=>{
        if(!subjectRows[sub.code]) subjectRows[sub.code]={name:sub.name,code:sub.code,total:0,pass:0,comp:0,abst:0};
        const r=subjectRows[sub.code]; r.total++;
        if(sub.grade==='AB') r.abst++;
        else if(sub.marks<33) r.comp++;
        else r.pass++;
      }));
      Object.values(subjectRows).sort((a,b)=>b.total-a.total).forEach(r=>{
        const pp=r.total?xpct(r.pass,r.total):0;
        rows.push([r.name, r.code, r.total, r.pass, r.comp, r.abst, pp+'%']);
      });
    });

    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=[{wch:38},{wch:10},{wch:10},{wch:10},{wch:20},{wch:10},{wch:10}];

    // styles
    const secHdrRows=[];
    rows.forEach((r,i)=>{
      if(typeof r[0]==='string' && r[0]===r[0].toUpperCase() && r[0].length>3 && r[1]!==undefined && r[1]!=='')
        secHdrRows.push(i);
    });
    // title
    const t0=XLSX.utils.encode_cell({r:0,c:0});
    if(ws[t0]) ws[t0].s={font:{bold:true,sz:14},fill:{fgColor:{rgb:LGOLD}},alignment:{horizontal:'left'}};
    const t1=XLSX.utils.encode_cell({r:1,c:0});
    if(ws[t1]) ws[t1].s={font:{italic:true,sz:10,color:{rgb:'FF888888'}}};
    // section headers
    rows.forEach((row,ri)=>{
      if(!Array.isArray(row)||!row[0]) return;
      const v=String(row[0]);
      const isSecHdr = v===v.toUpperCase()&&v.length>4&&row.length>1&&row[1]!==undefined&&row[1]!=='';
      const isClsHdr = v.startsWith('— Class');
      if(isSecHdr){
        row.forEach((_,ci)=>{
          const a=XLSX.utils.encode_cell({r:ri,c:ci});
          if(ws[a]) ws[a].s=H(DARK);
        });
      } else if(isClsHdr){
        const a=XLSX.utils.encode_cell({r:ri,c:0});
        if(ws[a]) ws[a].s=SH(LGREY);
      } else if(ri>2 && row[1]!==undefined && row[1]!==''){
        // data rows — colour pass% column
        const pCol=row.length-1;
        const pp=parseFloat(row[pCol]);
        if(!isNaN(pp)){
          const a=XLSX.utils.encode_cell({r:ri,c:pCol});
          if(ws[a]) ws[a].s={font:{bold:true,color:{rgb:PC(pp)}},alignment:{horizontal:'center'}};
        }
        // centre numeric cols
        for(let c=1;c<row.length-1;c++){
          const a=XLSX.utils.encode_cell({r:ri,c});
          if(ws[a]&&typeof ws[a].v==='number') ws[a].s={alignment:{horizontal:'center'}};
        }
      }
    });

    XLSX.utils.book_append_sheet(wb, ws, 'Summary');
  }

  /* ════════════════════════════════════════════════════════
     SHEET 2 — SUBJECT ANALYSIS  (mirrors Subjects tab)
  ════════════════════════════════════════════════════════ */
  ['X','XII'].forEach(cls=>{
    const sts=DB[cls]; if(!sts.length) return;
    const subs=buildSubMap(sts);
    const hdrs=['Subject','Code','Appeared','Pass','Fail','Pass %',
                'Avg Marks','Highest','Lowest',...GRADES.map(g=>`Grade ${g}`)];
    const rows=subs.map(s=>[
      s.name, s.code, s.n, s.pass, s.n-s.pass,
      xpct(s.pass,s.n)+'%',
      xavg(s.marks), Math.max(...s.marks), Math.min(...s.marks),
      ...GRADES.map(g=>s.grades[g]||0)
    ]);
    const ws=XLSX.utils.aoa_to_sheet([hdrs,...rows]);
    ws['!cols']=[{wch:30},{wch:7},{wch:10},{wch:8},{wch:8},{wch:9},{wch:11},{wch:10},{wch:9},
                 ...GRADES.map(()=>({wch:9}))];
    hdrs.forEach((_,ci)=>{const a=XLSX.utils.encode_cell({r:0,c:ci});if(ws[a])ws[a].s=H(cls==='X'?GOLD:TEAL);});
    rows.forEach((row,ri)=>{
      const pp=parseFloat(row[5]);
      // pass% cell
      const pa=XLSX.utils.encode_cell({r:ri+1,c:5});
      if(ws[pa]) ws[pa].s={font:{bold:true,color:{rgb:PC(pp)}},alignment:{horizontal:'center'}};
      // centre all numeric
      [2,3,4,6,7,8,...GRADES.map((_,i)=>9+i)].forEach(ci=>{
        const a=XLSX.utils.encode_cell({r:ri+1,c:ci});
        if(ws[a]) ws[a].s={alignment:{horizontal:'center'}};
      });
    });
    XLSX.utils.book_append_sheet(wb, ws, `Cl ${cls} Subjects`);
  });

  /* ════════════════════════════════════════════════════════
     SHEET 3 — MERIT LIST  (mirrors Merit tab — ALL pass students)
  ════════════════════════════════════════════════════════ */
  ['X','XII'].forEach(cls=>{
    const sts=DB[cls]; if(!sts.length) return;
    const subCols=getSubjectCols(sts.filter(s=>s.result==='PASS'));
    const passed=sts.filter(s=>s.result==='PASS')
      .map(s=>({...s,total:xtm(s)}))
      .sort((a,b)=>b.total-a.total);

    const fixHdrs=['Rank','Roll No','Name','Gender','Total Marks'];
    const subHdrs=subCols.flatMap(s=>[s.name+' (Marks)',s.name+' (Grade)']);
    const hdrs=[...fixHdrs,...subHdrs];

    const rows=passed.map((s,i)=>{
      const fix=[i+1, s.rollNo, s.name, s.gender==='F'?'Female':'Male', s.total];
      const sdata=subCols.flatMap(sc=>{
        const f=s.subjects.find(x=>x.code===sc.code);
        return f?[f.grade==='AB'?'AB':f.marks, f.grade]:['',''];
      });
      return [...fix,...sdata];
    });

    const ws=XLSX.utils.aoa_to_sheet([hdrs,...rows]);
    ws['!cols']=[{wch:6},{wch:12},{wch:28},{wch:9},{wch:13},...subCols.flatMap(()=>[{wch:20},{wch:9}])];
    // header
    hdrs.forEach((_,ci)=>{
      const a=XLSX.utils.encode_cell({r:0,c:ci});
      if(ws[a]) ws[a].s=H(ci<5?DARK:(cls==='X'?GOLD:TEAL));
    });
    // medals for top 3
    const medals=[{fill:'FFC9A84C',ink:DARK},{fill:'FFCCCCCC',ink:'FF444444'},{fill:'FFCD7F32',ink:'FF444444'}];
    rows.forEach((row,ri)=>{
      const medal=medals[ri];
      hdrs.forEach((_,ci)=>{
        const a=XLSX.utils.encode_cell({r:ri+1,c:ci});
        if(!ws[a]) return;
        if(medal) ws[a].s={fill:{fgColor:{rgb:medal.fill}},font:{bold:ri===0,color:{rgb:medal.ink}},
                            alignment:{horizontal:ci===0||ci===4?'center':'left'}};
        else if(ci===0||ci===4) ws[a].s={alignment:{horizontal:'center'},font:{bold:ci===4}};
        else if(ci>=5&&ci%2===1) ws[a].s={alignment:{horizontal:'center'}};
      });
    });
    XLSX.utils.book_append_sheet(wb, ws, `Cl ${cls} Merit`);
  });

  /* ════════════════════════════════════════════════════════
     SHEET 4 — GENDER ANALYSIS  (mirrors Gender tab)
  ════════════════════════════════════════════════════════ */
  ['X','XII'].forEach(cls=>{
    const sts=DB[cls]; if(!sts.length) return;
    const boys=sts.filter(s=>s.gender==='M'), girls=sts.filter(s=>s.gender==='F');
    const sb=xst(boys), sg=xst(girls);
    const subs=buildSubMap(sts);
    const rows=[];

    // overall comparison
    rows.push([`GENDER ANALYSIS — CLASS ${cls}`]);
    rows.push([]);
    rows.push(['','Boys','Girls']);
    rows.push(['Total Students',     sb.n,                sg.n]);
    rows.push(['Passed',             sb.pass,             sg.pass]);
    rows.push(['Compartment',        sb.comp,             sg.comp]);
    rows.push(['Absent',             sb.abst,             sg.abst]);
    rows.push(['Pass %',             sb.pct+'%',          sg.pct+'%']);
    rows.push(['Avg Total Marks',
      xavg(boys.filter(s=>s.result==='PASS').map(s=>xtm(s))),
      xavg(girls.filter(s=>s.result==='PASS').map(s=>xtm(s)))
    ]);
    rows.push([]);

    // grade distribution by gender
    rows.push(['GRADE DISTRIBUTION','Boys',...Array(GRADES.length-1).fill(''),'Girls']);
    rows.push(['Grade',...GRADES,'','Grade',...GRADES]);
    const gcB={},gcG={}; GRADES.forEach(g=>{gcB[g]=0;gcG[g]=0;});
    boys.forEach(x=>x.subjects.forEach(s=>{if(gcB[s.grade]!==undefined)gcB[s.grade]++;}));
    girls.forEach(x=>x.subjects.forEach(s=>{if(gcG[s.grade]!==undefined)gcG[s.grade]++;}));
    rows.push(['Count',...GRADES.map(g=>gcB[g]),'','Count',...GRADES.map(g=>gcG[g])]);
    rows.push([]);

    // subject-wise boys vs girls average
    rows.push(['SUBJECT-WISE AVERAGE — BOYS vs GIRLS','Code','Boys Avg','Boys Pass%','Girls Avg','Girls Pass%','Difference (B−G)']);
    subs.forEach(s=>{
      const ba=xavg(s.marksB), ga=xavg(s.marksG);
      rows.push([
        s.name, s.code,
        ba, xpct(s.passB,s.nB)+'%',
        ga, xpct(s.passG,s.nG)+'%',
        Math.round((ba-ga)*10)/10
      ]);
    });

    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=[{wch:34},{wch:10},{wch:12},{wch:12},{wch:12},{wch:12},{wch:16}];

    rows.forEach((row,ri)=>{
      if(!row[0]) return;
      const v=String(row[0]);
      // section headers
      if(v===v.toUpperCase()&&v.length>5&&row[1]!==undefined){
        row.forEach((_,ci)=>{
          const a=XLSX.utils.encode_cell({r:ri,c:ci});
          if(ws[a]) ws[a].s=H(cls==='X'?GOLD:TEAL);
        });
        return;
      }
      // label column bold
      const a0=XLSX.utils.encode_cell({r:ri,c:0});
      if(ws[a0]&&v&&!v.startsWith('Count')&&!v.startsWith('Grade')) ws[a0].s={font:{bold:true}};
      // Boys column blue tint, Girls column
      [1,2].forEach(ci=>{
        const a=XLSX.utils.encode_cell({r:ri,c:ci});
        if(ws[a]&&typeof ws[a].v==='number') ws[a].s={alignment:{horizontal:'center'}};
      });
      // Difference column — green if +ve, red if -ve
      const da=XLSX.utils.encode_cell({r:ri,c:6});
      if(ws[da]&&typeof ws[da].v==='number'){
        const diff=ws[da].v;
        ws[da].s={font:{bold:true,color:{rgb:diff>0?PASS_C:diff<0?FAIL_C:'FF888888'}},
                  alignment:{horizontal:'center'}};
      }
    });

    XLSX.utils.book_append_sheet(wb, ws, `Cl ${cls} Gender`);
  });

  /* ════════════════════════════════════════════════════════
     SHEET 5 — ALL STUDENTS RAW DATA  (mirrors Search/full student list)
  ════════════════════════════════════════════════════════ */
  ['X','XII'].forEach(cls=>{
    const sts=DB[cls]; if(!sts.length) return;
    const subCols=getSubjectCols(sts);
    const fixHdrs=['Roll No','Name','Gender','Result','Compartment Sub','Total Marks'];
    const subHdrs=subCols.flatMap(s=>[s.name+' (Marks)',s.name+' (Grade)']);
    const hdrs=[...fixHdrs,...subHdrs];

    const rows=sts.map(s=>{
      const fix=[s.rollNo, s.name, s.gender==='F'?'Female':'Male',
                 s.result, s.compSub||'', xtm(s)];
      const sdata=subCols.flatMap(sc=>{
        const f=s.subjects.find(x=>x.code===sc.code);
        return f?[f.grade==='AB'?'AB':f.marks,f.grade]:['',''];
      });
      return [...fix,...sdata];
    });

    const ws=XLSX.utils.aoa_to_sheet([hdrs,...rows]);
    ws['!cols']=[{wch:12},{wch:28},{wch:9},{wch:12},{wch:18},{wch:13},
                 ...subCols.flatMap(()=>[{wch:22},{wch:9}])];
    // header
    hdrs.forEach((_,ci)=>{
      const a=XLSX.utils.encode_cell({r:0,c:ci});
      if(ws[a]) ws[a].s=H(ci<6?DARK:(cls==='X'?GOLD:TEAL));
    });
    // data rows
    rows.forEach((row,ri)=>{
      const res=row[3];
      hdrs.forEach((_,ci)=>{
        const a=XLSX.utils.encode_cell({r:ri+1,c:ci}); if(!ws[a]) return;
        if(ci===3) ws[a].s={font:{bold:true,color:{rgb:RC(res)}},alignment:{horizontal:'center'}};
        else if(ci===5) ws[a].s={font:{bold:true},alignment:{horizontal:'center'}};
        else if(ci>=6&&ci%2===0) ws[a].s={alignment:{horizontal:'center'}};
        else if(ci>=6&&ci%2===1) ws[a].s={font:{color:{rgb:RC(res)}},alignment:{horizontal:'center'}};
      });
    });

    XLSX.utils.book_append_sheet(wb, ws, `Cl ${cls} All Students`);
  });

  /* ── download ── */
  const school=(meta.school||'CBSE').replace(/\s+/g,'_').substring(0,28);
  XLSX.writeFile(wb, `CBSE_${school}_${meta.year||new Date().getFullYear()}.xlsx`);
}

// ── AUTO RESTORE ON PAGE LOAD ──
tryRestoreFromLocalStorage();

// Close import overlay on outside click or ESC
document.getElementById('import-overlay').addEventListener('click', e=>{
  if(e.target === e.currentTarget) closeImportOverlay();
});
document.addEventListener('keydown', e=>{
  if(e.key==='Escape') closeImportOverlay();
});
