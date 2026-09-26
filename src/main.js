import { supabase } from './supabaseClient.js';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';

/* ================= CONSTANTS ================= */
const DAYS = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY"];
const COLUMNS = [
  {type:'period', idx:0, label:"09:00 AM TO\n09:50 AM", hour:1},
  {type:'period', idx:1, label:"09:50 AM TO\n10:35 AM", hour:2},
  {type:'period', idx:2, label:"10:35 AM TO\n11:20 AM", hour:3},
  {type:'break',  label:"BREAK\n11:20-11:35", tag:"BREAK"},
  {type:'period', idx:3, label:"11:35 AM TO\n12:20 PM", hour:4},
  {type:'period', idx:4, label:"12:20 PM TO\n01:00 PM", hour:5},
  {type:'break',  label:"LUNCH\n01:00-01:50", tag:"LUNCH BREAK"},
  {type:'period', idx:5, label:"01:50 PM TO\n02:30 PM", hour:6},
  {type:'period', idx:6, label:"02:30 PM TO\n03:15 PM", hour:7},
  {type:'period', idx:7, label:"03:15 PM TO\n04:00 PM", hour:8},
  {type:'break',  label:"BREAK\n04:00-04:10", tag:"BREAK"},
  {type:'period', idx:8, label:"04:10 PM TO\n05:00 PM", hour:9},
];
const PERIOD_COUNT = 9;
const CLASS_LIST = ["II A","II B","III A","III B","IV A","IV B"];
const YEAR_OPTS = ["I Year","II Year","III Year","IV Year"];
const COLOR_PALETTE = ["#FDE68A","#BFDBFE","#BBF7D0","#FBCFE8","#DDD6FE","#FED7AA","#A7F3D0","#FCA5A5","#C7D2FE","#FDE2E4","#D9F99D","#99F6E4","#FFE4B5","#B0E0E6"];
const COLLEGE = { name:"KONGUNADU COLLEGE OF ENGINEERING AND TECHNOLOGY", caption:"(AUTONOMOUS)",
  address:"Namakkal - Trichy Main Road, Tholurpatti, Thottiam, Trichy", dept:"DEPARTMENT OF INFORMATION TECHNOLOGY",
  ac:"AC-02", rev:"Rev:2" };

/* ================= SUPABASE-BACKED STATE ================= */
/* `state` keeps the same in-memory shape the UI code reads from — it's
   populated from Supabase on load, and every mutating function below
   writes through to Supabase first, then updates this local mirror. */
let state = {
  faculty: [],
  academicYear: "2025-2026", semester:"ODD", programme:"B.Tech - Information Technology",
  subjectMaster: {},   // class -> [{id,code,title,shortcut}]
  allotment: {},       // class -> { acadYear -> [{id,subjectId,faculty,workload,color}] }
  ourTT: {},           // class -> [versions]
  ourActive: {},       // class -> versionId
  otherTT: []
};

async function loadAllFromSupabase(){
  const [settingsRes, facultyRes, subjRes, allotRes, verRes, otherRes] = await Promise.all([
    supabase.from('app_settings').select('*').single(),
    supabase.from('faculty').select('*').order('name'),
    supabase.from('subject_master').select('*').order('created_at'),
    supabase.from('allotments').select('*'),
    supabase.from('our_versions').select('*').order('created_at'),
    supabase.from('other_timetables').select('*').order('created_at')
  ]);
  if(settingsRes.error) throw settingsRes.error;
  if(facultyRes.error) throw facultyRes.error;
  if(subjRes.error) throw subjRes.error;
  if(allotRes.error) throw allotRes.error;
  if(verRes.error) throw verRes.error;
  if(otherRes.error) throw otherRes.error;

  const s = settingsRes.data || {};
  state.academicYear = s.academic_year || "2025-2026";
  state.semester = s.semester || "ODD";
  state.programme = s.programme || "B.Tech - Information Technology";

  state.faculty = (facultyRes.data||[]).map(f=>({
    id:f.id, name:f.name, designation:f.designation||"", responsibility:f.responsibility||"",
    dob:f.dob||"", doj:f.doj||"", aicteCode:f.aicte_code||"", auCode:f.au_code||"",
    aadhar:f.aadhar||"", pan:f.pan||""
  }));

  state.subjectMaster = {};
  (subjRes.data||[]).forEach(row=>{
    if(!state.subjectMaster[row.class]) state.subjectMaster[row.class]=[];
    state.subjectMaster[row.class].push({id:row.id, code:row.code||"", title:row.title||"", shortcut:row.shortcut||""});
  });

  state.allotment = {};
  (allotRes.data||[]).forEach(row=>{
    if(!state.allotment[row.class]) state.allotment[row.class]={};
    if(!state.allotment[row.class][row.academic_year]) state.allotment[row.class][row.academic_year]=[];
    state.allotment[row.class][row.academic_year].push({
      id:row.id, subjectId:row.subject_id, faculty:row.faculty_name||"", workload:row.workload||"", color:row.color||nextColor([])
    });
  });

  state.ourTT = {}; state.ourActive = {};
  (verRes.data||[]).forEach(row=>{
    if(!state.ourTT[row.class]) state.ourTT[row.class]=[];
    state.ourTT[row.class].push({ id:row.id, label:row.label, createdAt:row.created_at, meta:row.meta||{}, subjects:row.subjects||[], grid:row.grid||emptyGrid() });
    if(row.is_active) state.ourActive[row.class] = row.id;
  });

  state.otherTT = (otherRes.data||[]).map(row=>({ id:row.id, meta:row.meta||{}, subjects:row.subjects||[], grid:row.grid||emptyGrid() }));
}

async function sbCall(promise, label){
  const { data, error } = await promise;
  if(error){ alert((label||"Save")+" failed: "+error.message); throw error; }
  return data;
}

function uid(p){ return p+"_"+Math.random().toString(36).slice(2,9); }
function esc(s){ return (s==null?"":s).toString().replace(/[&<>"]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }
function nl2br(s){ return esc(s).replace(/\n/g,"<br>"); }
function acadParts(s){ const p=(s||"").split("-"); return [p[0]||"", p[1]||""]; }
function nowStr(){ const d=new Date(); return d.toLocaleDateString('en-GB')+" "+d.toLocaleTimeString('en-GB'); }
function reportId(){ return "TT-"+Date.now().toString(36).toUpperCase(); }
function nextColor(list){ const used=list.map(s=>s.color); for(const c of COLOR_PALETTE) if(!used.includes(c)) return c; return COLOR_PALETTE[list.length % COLOR_PALETTE.length]; }
function emptyGrid(){ const g={}; DAYS.forEach(d=> g[d]=new Array(PERIOD_COUNT).fill("")); return g; }
function mask(v){ if(!v) return ""; const s=v.toString(); return s.length<=4? "••••" : "••••"+s.slice(-4); }

function subjMasterList(cls){ if(!state.subjectMaster[cls]) state.subjectMaster[cls]=[]; return state.subjectMaster[cls]; }
function allotList(cls,ay){
  if(!state.allotment[cls]) state.allotment[cls]={};
  if(!state.allotment[cls][ay]) state.allotment[cls][ay]=[];
  return state.allotment[cls][ay];
}
function ourVersions(cls){ if(!state.ourTT[cls]) state.ourTT[cls]=[]; return state.ourTT[cls]; }
function activeOurRec(cls){
  const vs = ourVersions(cls); if(!vs.length) return null;
  const id = state.ourActive[cls];
  return vs.find(v=>v.id===id) || vs[vs.length-1];
}
function otherLabel(rec){ return (rec.meta.department||'Dept')+" "+(rec.meta.year||'')+" "+(rec.meta.section||''); }

/* aggregate every "live" timetable used for conflict + faculty aggregation:
   active version per class + every other-dept record */
function allTimetables(){
  const list=[];
  CLASS_LIST.forEach(c=>{ const rec=activeOurRec(c); if(rec) list.push({scope:'class', label:'Our Dept - '+c, key:c, rec}); });
  state.otherTT.forEach(r=> list.push({scope:'other', label:'Other Dept - '+otherLabel(r), key:r.id, rec:r}));
  return list;
}
function findSubjectByShortcut(rec, shortcut){
  return rec.subjects.find(s=> s.shortcut && s.shortcut.toLowerCase()===String(shortcut).toLowerCase());
}
function findConflicts(facultyName, day, periodIdx, excludeRec){
  if(!facultyName) return [];
  const out=[];
  allTimetables().forEach(t=>{
    const val = t.rec.grid[day][periodIdx]; if(!val) return;
    const subj = findSubjectByShortcut(t.rec, val);
    if(subj && subj.faculty && subj.faculty.trim().toLowerCase()===facultyName.trim().toLowerCase() && t.rec!==excludeRec){
      out.push(t.label);
    }
  });
  return out;
}
function facultyWeek(name, scope){
  scope = scope || 'all'; // 'all' | 'class' (Our Dept) | 'other' (Other Dept)
  const week={}; DAYS.forEach(d=> week[d]=new Array(PERIOD_COUNT).fill(null));
  allTimetables().forEach(t=>{
    if(scope!=='all' && t.scope!==scope) return;
    DAYS.forEach(d=> t.rec.grid[d].forEach((val,pIdx)=>{
      if(!val) return;
      const subj = findSubjectByShortcut(t.rec, val);
      if(subj && subj.faculty && subj.faculty.trim().toLowerCase()===name.trim().toLowerCase()){
        week[d][pIdx] = { shortcut: val, text: t.label, color: subj.color };
      }
    }));
  });
  return week;
}
function weekToAutotable(week){
  const head=[["Day"].concat(COLUMNS.map(c=> c.type==='period'? c.label.replace('\n',' ')+' (H'+c.hour+')' : c.tag))];
  const body = DAYS.map(day=>{ const row=[day]; COLUMNS.forEach(col=>{ if(col.type==='break'){row.push(col.tag);return;} const cd=week[day][col.idx]; row.push(cd? `${cd.shortcut} (${cd.text})`:''); }); return row; });
  return {head, body};
}
function facultyAllotmentRows(name){
  const rows=[];
  allTimetables().forEach(t=>{
    t.rec.subjects.forEach(s=>{
      if(s.faculty && s.faculty.trim().toLowerCase()===name.trim().toLowerCase()){
        rows.push({ source:t.label, code:s.code, title:s.title, workload:s.workload||0 });
      }
    });
  });
  return rows;
}

/* ================= UI STATE ================= */
let ui = {
  tab:"faculty",
  facMaskAll:true,
  facFilterText:"",
  facExportIncludeIds:false,
  subjClass: CLASS_LIST[0],
  allotClass: CLASS_LIST[0],
  allotYear: state.academicYear,
  ourClass: CLASS_LIST[0],
  otherActiveId: null,
  repClassSel: [],
  repFacSingle:"",
  repFacMulti: []
};

/* ================= RENDER ROOT ================= */
function render(){
  document.getElementById('app').innerHTML = `
    <div class="topbar">
      <div class="logo-badge">IT</div>
      <div class="title-block">
        <h1 style="font-size:18px;">Timetable Management System</h1>
        <div class="sub">Department of Information Technology &middot; Kongunadu College of Engineering and Technology</div>
      </div>
      <div class="row" style="gap:8px;">
        <button class="btn small" onclick="reloadFromDatabase()">⟳ Reload</button>
        <button class="btn small" onclick="exportBackup()">⬇ Backup JSON</button>
        <button class="btn small" onclick="doLogout()">Sign out</button>
      </div>
    </div>
    <div class="tabs">
      ${tabBtn("faculty","Faculty Master")}
      ${tabBtn("subjects","Subject Master")}
      ${tabBtn("allotment","Allotment")}
      ${tabBtn("ourdept","Our Department")}
      ${tabBtn("otherdept","Other Department")}
      ${tabBtn("reports","Reports")}
    </div>
    <div id="tabBody"></div>`;
  renderTabBody();
}
function tabBtn(id,label){ return `<button class="tab-btn ${ui.tab===id?'active':''}" onclick="setTab('${id}')">${label}</button>`; }
function setTab(id){ ui.tab=id; renderTabBody(); }
function renderTabBody(){
  const el = document.getElementById('tabBody');
  if(ui.tab==='faculty') el.innerHTML = viewFaculty();
  else if(ui.tab==='subjects') el.innerHTML = viewSubjects();
  else if(ui.tab==='allotment') el.innerHTML = viewAllotment();
  else if(ui.tab==='ourdept') el.innerHTML = viewOurDept();
  else if(ui.tab==='otherdept') el.innerHTML = viewOtherDept();
  else if(ui.tab==='reports') el.innerHTML = viewReports();
}

/* ================= FACULTY MASTER ================= */
function viewFaculty(){
  const filt = ui.facFilterText.trim().toLowerCase();
  const rows = state.faculty.filter(f=>{
    if(!filt) return true;
    return (f.name+" "+f.designation+" "+f.responsibility).toLowerCase().includes(filt);
  });
  return `
  <div class="panel">
    <h2>Faculty Master</h2>
    <div class="small-note">Faculty Name is the unique key used everywhere (allotment, timetable, class advisor, individual timetable). Stored only in this browser's local storage.</div>
    <div class="grid-cols-3" style="margin-top:12px;">
      <div class="field"><label>Name</label><input type="text" id="f_name"></div>
      <div class="field"><label>Designation</label><input type="text" id="f_desig" placeholder="AP / Dean / HoD"></div>
      <div class="field"><label>Responsibility</label><input type="text" id="f_resp" placeholder="Class Advisor II A"></div>
      <div class="field"><label>Date of Birth</label><input type="date" id="f_dob"></div>
      <div class="field"><label>Date of Joining</label><input type="date" id="f_doj"></div>
      <div class="field"><label>AICTE Code</label><input type="text" id="f_aicte"></div>
      <div class="field"><label>Anna University Code</label><input type="text" id="f_au"></div>
      <div class="field"><label>Aadhar</label><input type="text" id="f_aadhar"></div>
      <div class="field"><label>PAN</label><input type="text" id="f_pan"></div>
    </div>
    <div style="margin-top:12px;"><button class="btn primary" onclick="addFaculty()">Add Faculty</button></div>
  </div>
  <div class="panel">
    <h3>Faculty List</h3>
    <div class="row">
      <div class="field" style="max-width:280px;"><input type="text" placeholder="Filter by name / designation / responsibility" value="${esc(ui.facFilterText)}" oninput="ui.facFilterText=this.value; renderTabBody();"></div>
      <button class="btn small" onclick="toggleMaskAll()">${ui.facMaskAll?'Show':'Hide'} Aadhar/PAN</button>
    </div>
    <div class="tt-wrap" style="margin-top:10px;">
      <table class="datatable">
        <thead><tr><th>Name</th><th>Designation</th><th>Responsibility</th><th>DOB</th><th>DOJ</th><th>AICTE</th><th>AU Code</th><th>Aadhar</th><th>PAN</th><th></th></tr></thead>
        <tbody>
        ${rows.length===0? `<tr><td colspan="10" class="empty-note">No faculty found.</td></tr>` : rows.map(f=>`
          <tr>
            <td>${esc(f.name)}</td><td>${esc(f.designation)}</td><td>${esc(f.responsibility)}</td>
            <td>${esc(f.dob)}</td><td>${esc(f.doj)}</td><td>${esc(f.aicteCode)}</td><td>${esc(f.auCode)}</td>
            <td>${ui.facMaskAll? mask(f.aadhar) : esc(f.aadhar)}</td>
            <td>${ui.facMaskAll? mask(f.pan) : esc(f.pan)}</td>
            <td><button class="btn small danger" onclick="removeFaculty('${f.id}')">✕</button></td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>
  <div class="panel">
    <h3>Faculty Directory Report</h3>
    <div class="small-note">Filtered list above is what gets exported.</div>
    <div class="row" style="margin-top:8px;">
      <label style="display:flex;align-items:center;gap:6px;font-size:12.5px;"><input type="checkbox" ${ui.facExportIncludeIds?'checked':''} onchange="ui.facExportIncludeIds=this.checked;"> Include Aadhar/PAN in export</label>
    </div>
    <div class="row" style="margin-top:8px;">
      <button class="btn primary" onclick="exportFacultyDirectoryPDF()">⬇ PDF</button>
      <button class="btn primary" onclick="exportFacultyDirectoryExcel()">⬇ Excel</button>
    </div>
  </div>`;
}
function toggleMaskAll(){ ui.facMaskAll=!ui.facMaskAll; renderTabBody(); }
async function addFaculty(){
  const name = document.getElementById('f_name').value.trim();
  if(!name){ alert("Name is required."); return; }
  if(state.faculty.some(f=>f.name.toLowerCase()===name.toLowerCase())){ alert("Faculty already exists."); return; }
  const row = {
    name,
    designation: document.getElementById('f_desig').value.trim(),
    responsibility: document.getElementById('f_resp').value.trim(),
    dob: document.getElementById('f_dob').value || null,
    doj: document.getElementById('f_doj').value || null,
    aicte_code: document.getElementById('f_aicte').value.trim(),
    au_code: document.getElementById('f_au').value.trim(),
    aadhar: document.getElementById('f_aadhar').value.trim(),
    pan: document.getElementById('f_pan').value.trim()
  };
  const data = await sbCall(supabase.from('faculty').insert(row).select().single(), "Add faculty");
  state.faculty.push({ id:data.id, name:data.name, designation:data.designation||"", responsibility:data.responsibility||"",
    dob:data.dob||"", doj:data.doj||"", aicteCode:data.aicte_code||"", auCode:data.au_code||"", aadhar:data.aadhar||"", pan:data.pan||"" });
  renderTabBody();
}
async function removeFaculty(id){
  await sbCall(supabase.from('faculty').delete().eq('id', id), "Remove faculty");
  state.faculty = state.faculty.filter(f=>f.id!==id);
  renderTabBody();
}
function facultyRowsForExport(){
  const filt = ui.facFilterText.trim().toLowerCase();
  return state.faculty.filter(f=> !filt || (f.name+" "+f.designation+" "+f.responsibility).toLowerCase().includes(filt));
}
function exportFacultyDirectoryPDF(){
  const rows = facultyRowsForExport();
  const doc = new jsPDF({orientation:"landscape", unit:"mm", format:"a4"});
  doc.setFont("helvetica","bold"); doc.setFontSize(13);
  doc.text(COLLEGE.name, doc.internal.pageSize.getWidth()/2, 14, {align:"center"});
  doc.setFontSize(9); doc.text("FACULTY DIRECTORY — "+COLLEGE.dept, doc.internal.pageSize.getWidth()/2, 20, {align:"center"});
  const head = ["Name","Designation","Responsibility","DOB","DOJ","AICTE Code","AU Code"];
  if(ui.facExportIncludeIds){ head.push("Aadhar","PAN"); }
  const body = rows.map(f=>{
    const r=[f.name,f.designation,f.responsibility,f.dob,f.doj,f.aicteCode,f.auCode];
    if(ui.facExportIncludeIds) r.push(f.aadhar,f.pan);
    return r;
  });
  doc.autoTable({ head:[head], body, startY:26, styles:{fontSize:8,cellPadding:1.6}, headStyles:{fillColor:[238,241,247],textColor:20,fontStyle:'bold'}, margin:{left:14,right:14} });
  doc.save("Faculty_Directory.pdf");
}
function exportFacultyDirectoryExcel(){
  const rows = facultyRowsForExport();
  const head = ["Name","Designation","Responsibility","DOB","DOJ","AICTE Code","AU Code"];
  if(ui.facExportIncludeIds) head.push("Aadhar","PAN");
  const aoa = [head];
  rows.forEach(f=>{ const r=[f.name,f.designation,f.responsibility,f.dob,f.doj,f.aicteCode,f.auCode]; if(ui.facExportIncludeIds) r.push(f.aadhar,f.pan); aoa.push(r); });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = head.map(()=>({wch:16}));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Faculty");
  XLSX.writeFile(wb, "Faculty_Directory.xlsx");
}

/* ================= SUBJECT MASTER ================= */
function viewSubjects(){
  const list = subjMasterList(ui.subjClass);
  return `
  <div class="panel">
    <h2>Subject Master</h2>
    <div class="row" style="margin-top:8px;">
      ${CLASS_LIST.map(c=>`<button class="tab-btn ${ui.subjClass===c?'active':''}" onclick="setSubjClass('${c}')">${c}</button>`).join('')}
    </div>
  </div>
  <div class="panel">
    <h3>Add Subject — ${esc(ui.subjClass)}</h3>
    <div class="small-note">Leave Code blank for Library, Coaching, or Event entries.</div>
    <div class="grid-cols-3" style="margin-top:10px;">
      <div class="field"><label>Subject Code (optional)</label><input type="text" id="s_code"></div>
      <div class="field"><label>Subject Name</label><input type="text" id="s_title"></div>
      <div class="field"><label>Shortcut</label><input type="text" id="s_short"></div>
    </div>
    <div style="margin-top:10px;"><button class="btn primary" onclick="addSubjectMasterRow()">Add</button></div>
  </div>
  <div class="panel">
    <h3>Import from file — ${esc(ui.subjClass)}</h3>
    <div class="small-note">Excel/CSV with header row containing Code, Title/Name, Shortcut columns (Shortcut can be blank and filled in later).</div>
    <input type="file" id="subjImportFile" accept=".xlsx,.xls,.csv" style="margin-top:8px;max-width:320px;">
    <div style="margin-top:8px;"><button class="btn" onclick="importSubjects()">Import</button></div>
  </div>
  <div class="panel">
    <h3>Subjects — ${esc(ui.subjClass)}</h3>
    <div class="tt-wrap">
      <table class="datatable"><thead><tr><th>#</th><th>Code</th><th>Title</th><th>Shortcut</th><th></th></tr></thead>
      <tbody>${list.length===0? `<tr><td colspan="5" class="empty-note">No subjects yet.</td></tr>` : list.map((s,i)=>`
        <tr>
          <td>${i+1}</td>
          <td><input type="text" style="min-width:90px;" value="${esc(s.code)}" onchange="updSubjMaster('${s.id}','code',this.value)"></td>
          <td><input type="text" style="min-width:200px;" value="${esc(s.title)}" onchange="updSubjMaster('${s.id}','title',this.value)"></td>
          <td><input type="text" style="width:80px;font-weight:700;" value="${esc(s.shortcut)}" onchange="updSubjMaster('${s.id}','shortcut',this.value)"></td>
          <td><button class="btn small danger" onclick="delSubjMaster('${s.id}')">✕</button></td>
        </tr>`).join('')}</tbody>
      </table>
    </div>
  </div>`;
}
function setSubjClass(c){ ui.subjClass=c; renderTabBody(); }
async function addSubjectMasterRow(){
  const code = document.getElementById('s_code').value.trim();
  const title = document.getElementById('s_title').value.trim();
  const shortcut = document.getElementById('s_short').value.trim();
  if(!title){ alert("Subject name is required."); return; }
  const data = await sbCall(supabase.from('subject_master').insert({class:ui.subjClass, code, title, shortcut}).select().single(), "Add subject");
  subjMasterList(ui.subjClass).push({id:data.id, code:data.code||"", title:data.title||"", shortcut:data.shortcut||""});
  renderTabBody();
}
async function updSubjMaster(id,field,val){
  const list = subjMasterList(ui.subjClass); const s=list.find(x=>x.id===id); if(!s) return;
  const col = field==='code'?'code':field==='title'?'title':'shortcut';
  await sbCall(supabase.from('subject_master').update({[col]:val}).eq('id',id), "Update subject");
  s[field]=val;
}
async function delSubjMaster(id){
  await sbCall(supabase.from('subject_master').delete().eq('id',id), "Delete subject");
  state.subjectMaster[ui.subjClass] = subjMasterList(ui.subjClass).filter(s=>s.id!==id);
  renderTabBody();
}
function importSubjects(){
  const inp = document.getElementById('subjImportFile');
  const file = inp.files[0]; if(!file){ alert("Choose a file first."); return; }
  const reader = new FileReader();
  reader.onload = async e=>{
    try{
      const wb = XLSX.read(e.target.result, {type:'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1});
      if(!rows.length){ alert("File appears empty."); return; }
      const header = rows[0].map(h=>(h||"").toString().toLowerCase());
      const codeIdx = header.findIndex(h=>h.includes('code'));
      const titleIdx = header.findIndex(h=>h.includes('title')||h.includes('name'));
      const shortIdx = header.findIndex(h=>h.includes('short'));
      if(titleIdx===-1){ alert("Could not find a Title/Name column in the file header."); return; }
      const list = subjMasterList(ui.subjClass);
      const toInsert = [];
      for(let i=1;i<rows.length;i++){
        const row = rows[i]; if(!row || !row[titleIdx]) continue;
        const title = row[titleIdx].toString().trim();
        const code = codeIdx>=0 && row[codeIdx] ? row[codeIdx].toString().trim() : "";
        const shortcut = shortIdx>=0 && row[shortIdx] ? row[shortIdx].toString().trim() : "";
        if(code && list.some(s=>s.code && s.code.toLowerCase()===code.toLowerCase())) continue;
        toInsert.push({class:ui.subjClass, code, title, shortcut});
      }
      if(!toInsert.length){ alert("Nothing new to import."); return; }
      const data = await sbCall(supabase.from('subject_master').insert(toInsert).select(), "Import subjects");
      data.forEach(row=> list.push({id:row.id, code:row.code||"", title:row.title||"", shortcut:row.shortcut||""}));
      renderTabBody();
      alert(`Imported ${data.length} subject(s) into ${ui.subjClass}.`);
    }catch(err){ alert("Could not read that file: "+err.message); }
  };
  reader.readAsArrayBuffer(file);
}

/* ================= ALLOTMENT ================= */
function viewAllotment(){
  const subjects = subjMasterList(ui.allotClass);
  const rows = allotList(ui.allotClass, ui.allotYear);
  const facOptions = state.faculty.map(f=>`<option value="${esc(f.name)}">${esc(f.name)}</option>`).join('');
  return `
  <div class="panel">
    <h2>Subject-Faculty Allotment</h2>
    <div class="row" style="margin-top:8px;">
      ${CLASS_LIST.map(c=>`<button class="tab-btn ${ui.allotClass===c?'active':''}" onclick="setAllotClass('${c}')">${c}</button>`).join('')}
    </div>
    <div class="field" style="max-width:220px;margin-top:10px;">
      <label>Academic Year</label>
      <input type="text" value="${esc(ui.allotYear)}" placeholder="2025-2026" onchange="ui.allotYear=this.value; renderTabBody();">
    </div>
  </div>
  <div class="panel">
    <h3>Allotment — ${esc(ui.allotClass)} · ${esc(ui.allotYear)}</h3>
    ${subjects.length===0? `<div class="empty-note">No subjects in the Subject Master for ${esc(ui.allotClass)} yet — add them in the Subject Master tab first.</div>` : `
    <div class="tt-wrap">
      <table class="datatable">
        <thead><tr><th>Code</th><th>Subject</th><th>Shortcut</th><th>Faculty</th><th>Colour</th><th>Work Load</th></tr></thead>
        <tbody>
        ${subjects.map(subj=>{
          let row = rows.find(r=>r.subjectId===subj.id);
          if(!row){ row = {id:uid('al'), subjectId:subj.id, faculty:"", workload:"", color: nextColor(rows), _pending:true}; rows.push(row); }
          return `<tr>
            <td>${esc(subj.code)}</td><td>${esc(subj.title)}</td><td><b>${esc(subj.shortcut)}</b></td>
            <td><input type="text" list="allotFacList" style="min-width:170px;" value="${esc(row.faculty)}" onchange="updAllotRow('${row.id}','${subj.id}','faculty',this.value)"></td>
            <td><input type="color" value="${row.color}" style="width:36px;height:28px;border:none;background:none;" onchange="updAllotRow('${row.id}','${subj.id}','color',this.value)"></td>
            <td><input type="number" style="width:60px;" value="${row.workload}" onchange="updAllotRow('${row.id}','${subj.id}','workload',this.value)"></td>
          </tr>`;
        }).join('')}
        </tbody>
      </table>
    </div>
    <datalist id="allotFacList">${facOptions}</datalist>
    <div class="small-note" style="margin-top:8px;">Each field saves automatically as you edit it.</div>
    `}
  </div>`;
}
function setAllotClass(c){ ui.allotClass=c; renderTabBody(); }
async function updAllotRow(id,subjectId,field,val){
  const rows = allotList(ui.allotClass, ui.allotYear);
  const r = rows.find(x=>x.id===id); if(!r) return;
  r[field]=val;
  if(r._pending){
    const insertRow = { class:ui.allotClass, academic_year:ui.allotYear, subject_id:subjectId,
      faculty_name:r.faculty||null, workload: r.workload===""?null:r.workload, color:r.color };
    const data = await sbCall(supabase.from('allotments').insert(insertRow).select().single(), "Save allotment");
    r.id = data.id; delete r._pending;
  } else {
    const col = field==='faculty' ? 'faculty_name' : field;
    const patch = {}; patch[col] = field==='workload' ? (val===""?null:val) : val;
    await sbCall(supabase.from('allotments').update(patch).eq('id', r.id), "Save allotment");
  }
}

/* ================= OUR DEPARTMENT ================= */
function viewOurDept(){
  const cls = ui.ourClass;
  const versions = ourVersions(cls);
  const activeId = state.ourActive[cls];
  const activeRec = activeOurRec(cls);
  return `
  <div class="panel">
    <h2>Our Department / Class Timetable</h2>
    <div class="row" style="margin-top:8px;">
      ${CLASS_LIST.map(c=>`<button class="tab-btn ${ui.ourClass===c?'active':''}" onclick="setOurClass('${c}')">${c}</button>`).join('')}
    </div>
  </div>
  <div class="panel">
    <h3>Create New Version — ${esc(cls)}</h3>
    <div class="grid-cols-3">
      <div class="field"><label>Academic Year</label><input type="text" id="ov_ay" value="${esc(state.academicYear)}" placeholder="2025-2026"></div>
      <div class="field"><label>Semester</label><select id="ov_sem"><option>ODD</option><option>EVEN</option></select></div>
      <div class="field"><label>Hall Number</label><input type="text" id="ov_hall"></div>
      <div class="field"><label>W.E.F. Date</label><input type="text" id="ov_wef" placeholder="DD-MM-YYYY"></div>
      <div class="field"><label>Class Advisor</label><input type="text" id="ov_advisor" list="ourFacList"></div>
      <div class="field"><label>Version Label (optional)</label><input type="text" id="ov_label" placeholder="e.g. Odd 2025-26"></div>
    </div>
    <datalist id="ourFacList">${state.faculty.map(f=>`<option value="${esc(f.name)}">`).join('')}</datalist>
    <div class="small-note" style="margin-top:6px;">Subjects and faculty for this version will be pulled from the Allotment set for ${esc(cls)} at the Academic Year you enter above — make sure that allotment exists first.</div>
    <div style="margin-top:10px;"><button class="btn primary" onclick="createOurVersion('${cls}')">Create Version (54-box grid)</button></div>
  </div>
  <div class="panel">
    <h3>Versions</h3>
    ${versions.length===0? `<div class="empty-note">No versions yet for ${esc(cls)}.</div>` : `
    <div class="row">
      ${versions.map(v=>`<button class="version-chip ${v.id===(activeId||versions[versions.length-1].id)?'active':''}" onclick="setOurActive('${cls}','${v.id}')">${esc(v.label)} ${v.id===(activeId||versions[versions.length-1].id)?'· active':''}</button>`).join('')}
    </div>
    <div class="small-note" style="margin-top:6px;">The active version is what's used for conflict checks, individual faculty timetables, and reports.</div>
    `}
  </div>
  ${activeRec ? metaPanelOur(cls, activeRec) : ''}
  ${activeRec ? subjectPanelOur(cls, activeRec) : ''}
  ${activeRec ? gridPanel(activeRec, 'class', cls, cls) : ''}
  ${activeRec ? exportPanel('class', cls, cls) : ''}
  `;
}
function setOurClass(c){ ui.ourClass=c; renderTabBody(); }
async function persistOurVersion(cls, rec, fields){
  const patch = {}; fields.forEach(f=> patch[f]=rec[f]);
  await sbCall(supabase.from('our_versions').update(patch).eq('id', rec.id), "Save timetable");
}
async function createOurVersion(cls){
  const ay = document.getElementById('ov_ay').value.trim() || state.academicYear;
  const sem = document.getElementById('ov_sem').value;
  const hall = document.getElementById('ov_hall').value.trim();
  const wef = document.getElementById('ov_wef').value.trim();
  const advisor = document.getElementById('ov_advisor').value.trim();
  let label = document.getElementById('ov_label').value.trim();
  if(!label) label = "Version "+nowStr();
  const alRows = allotList(cls, ay);
  const subjMaster = subjMasterList(cls);
  const subjects = alRows.filter(r=>r.faculty).map(r=>{
    const sm = subjMaster.find(s=>s.id===r.subjectId) || {code:"",title:"",shortcut:""};
    return { id:uid('vs'), subjectId:r.subjectId, code:sm.code, title:sm.title, shortcut:sm.shortcut, faculty:r.faculty, workload:r.workload, color:r.color };
  });
  const meta = { academicYear: ay, semester: sem, programme: state.programme, hall, wef, advisor };
  const grid = emptyGrid();
  await supabase.from('our_versions').update({is_active:false}).eq('class', cls).eq('is_active', true);
  const data = await sbCall(supabase.from('our_versions').insert({ class:cls, label, meta, subjects, grid, is_active:true }).select().single(), "Create version");
  const rec = { id:data.id, label:data.label, createdAt:data.created_at, meta:data.meta, subjects:data.subjects, grid:data.grid };
  ourVersions(cls).push(rec);
  state.ourActive[cls] = rec.id;
  renderTabBody();
}
async function setOurActive(cls,id){
  await supabase.from('our_versions').update({is_active:false}).eq('class', cls).eq('is_active', true);
  await sbCall(supabase.from('our_versions').update({is_active:true}).eq('id', id), "Set active version");
  state.ourActive[cls]=id; renderTabBody();
}

function metaPanelOur(cls, rec){
  const m = rec.meta;
  return `
  <div class="panel">
    <h3>Header Details — ${esc(cls)} (${esc(rec.label)})</h3>
    <div class="grid-cols-3">
      <div class="field"><label>Academic Year</label><input type="text" value="${esc(m.academicYear)}" onchange="updOurMeta('${cls}','${rec.id}','academicYear',this.value)"></div>
      <div class="field"><label>Semester</label>
        <select onchange="updOurMeta('${cls}','${rec.id}','semester',this.value)">
          <option ${m.semester==='ODD'?'selected':''}>ODD</option><option ${m.semester==='EVEN'?'selected':''}>EVEN</option>
        </select>
      </div>
      <div class="field"><label>Programme</label><input type="text" value="${esc(m.programme)}" onchange="updOurMeta('${cls}','${rec.id}','programme',this.value)"></div>
      <div class="field"><label>Hall Number</label><input type="text" value="${esc(m.hall)}" onchange="updOurMeta('${cls}','${rec.id}','hall',this.value)"></div>
      <div class="field"><label>W.E.F.</label><input type="text" value="${esc(m.wef)}" onchange="updOurMeta('${cls}','${rec.id}','wef',this.value)"></div>
      <div class="field"><label>Class Advisor</label><input type="text" value="${esc(m.advisor)}" onchange="updOurMeta('${cls}','${rec.id}','advisor',this.value)"></div>
    </div>
  </div>`;
}
async function updOurMeta(cls,id,field,val){
  const rec = ourVersions(cls).find(v=>v.id===id); if(!rec) return;
  rec.meta[field]=val;
  await persistOurVersion(cls, rec, ['meta']);
}
function subjectPanelOur(cls, rec){
  return `
  <div class="panel">
    <h3>Subjects (from Allotment, editable for this version)</h3>
    <div class="tt-wrap">
      <table class="datatable">
        <thead><tr><th>Code</th><th>Title</th><th>Shortcut</th><th>Faculty</th><th>Colour</th><th>Work Load</th></tr></thead>
        <tbody>
        ${rec.subjects.length===0? `<tr><td colspan="6" class="empty-note">No allotted subjects found — check the Allotment tab for ${esc(cls)} / ${esc(rec.meta.academicYear)}.</td></tr>` :
        rec.subjects.map(s=>`
          <tr>
            <td><input type="text" style="width:80px;" value="${esc(s.code)}" onchange="updOurSubj('${cls}','${rec.id}','${s.id}','code',this.value)"></td>
            <td><input type="text" style="min-width:170px;" value="${esc(s.title)}" onchange="updOurSubj('${cls}','${rec.id}','${s.id}','title',this.value)"></td>
            <td><input type="text" style="width:70px;font-weight:700;" value="${esc(s.shortcut)}" onchange="updOurSubj('${cls}','${rec.id}','${s.id}','shortcut',this.value)"></td>
            <td><input type="text" list="ourFacList" style="min-width:160px;" value="${esc(s.faculty)}" onchange="updOurSubj('${cls}','${rec.id}','${s.id}','faculty',this.value)"></td>
            <td><input type="color" value="${s.color}" style="width:36px;height:28px;border:none;background:none;" onchange="updOurSubj('${cls}','${rec.id}','${s.id}','color',this.value)"></td>
            <td><input type="number" style="width:60px;" value="${s.workload||''}" onchange="updOurSubj('${cls}','${rec.id}','${s.id}','workload',this.value)"></td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
  </div>`;
}
async function updOurSubj(cls,recId,subjId,field,val){
  const rec = ourVersions(cls).find(v=>v.id===recId); if(!rec) return;
  const s = rec.subjects.find(x=>x.id===subjId); if(!s) return;
  s[field]=val;
  await persistOurVersion(cls, rec, ['subjects']);
  renderTabBody();
}

/* ================= OTHER DEPARTMENT ================= */
function viewOtherDept(){
  if(!ui.otherActiveId && state.otherTT.length) ui.otherActiveId = state.otherTT[0].id;
  const rec = state.otherTT.find(r=>r.id===ui.otherActiveId);
  return `
  <div class="panel">
    <h2>Other Department Timetable</h2>
    <div class="row" style="margin-top:8px;">
      ${state.otherTT.map(r=>`<button class="tab-btn ${ui.otherActiveId===r.id?'active':''}" onclick="setOtherActive('${r.id}')">${esc(otherLabel(r))}</button>`).join('')}
      <button class="btn small primary" onclick="addOtherTT()">+ New Timetable</button>
    </div>
  </div>
  ${rec? otherMetaPanel(rec):`<div class="panel"><div class="empty-note">No other-department timetable yet.</div></div>`}
  ${rec? otherSubjectPanel(rec):''}
  ${rec? gridPanel(rec,'other',rec.id, otherLabel(rec)):''}
  ${rec? exportPanel('other',rec.id, otherLabel(rec)):''}
  `;
}
async function addOtherTT(){
  const meta = { department:"", programme: state.programme, year: YEAR_OPTS[0], section:"", academicYear: state.academicYear, semester: state.semester||"ODD", hall:"", wef:"", advisor:"" };
  const data = await sbCall(supabase.from('other_timetables').insert({ meta, subjects:[], grid: emptyGrid() }).select().single(), "Create timetable");
  const rec = { id:data.id, meta:data.meta, subjects:data.subjects, grid:data.grid };
  state.otherTT.push(rec);
  ui.otherActiveId = rec.id; renderTabBody();
}
function setOtherActive(id){ ui.otherActiveId=id; renderTabBody(); }
async function removeOtherTT(id){
  if(!confirm("Delete this other-department timetable?")) return;
  await sbCall(supabase.from('other_timetables').delete().eq('id', id), "Delete timetable");
  state.otherTT = state.otherTT.filter(r=>r.id!==id);
  ui.otherActiveId = state.otherTT.length? state.otherTT[0].id : null;
  renderTabBody();
}
async function persistOtherTT(rec, fields){
  const patch = {}; fields.forEach(f=> patch[f]=rec[f]);
  await sbCall(supabase.from('other_timetables').update(patch).eq('id', rec.id), "Save timetable");
}
function otherMetaPanel(rec){
  const m = rec.meta;
  return `
  <div class="panel">
    <h3>Header Details</h3>
    <div class="grid-cols-3">
      <div class="field"><label>Department</label><input type="text" value="${esc(m.department)}" onchange="updOtherMeta('${rec.id}','department',this.value)"></div>
      <div class="field"><label>Programme</label><input type="text" value="${esc(m.programme)}" onchange="updOtherMeta('${rec.id}','programme',this.value)"></div>
      <div class="field"><label>Year</label><select onchange="updOtherMeta('${rec.id}','year',this.value)">${YEAR_OPTS.map(y=>`<option ${m.year===y?'selected':''}>${y}</option>`).join('')}</select></div>
      <div class="field"><label>Section</label><input type="text" value="${esc(m.section)}" onchange="updOtherMeta('${rec.id}','section',this.value)"></div>
      <div class="field"><label>Academic Year</label><input type="text" value="${esc(m.academicYear)}" onchange="updOtherMeta('${rec.id}','academicYear',this.value)"></div>
      <div class="field"><label>Semester</label><select onchange="updOtherMeta('${rec.id}','semester',this.value)"><option ${m.semester==='ODD'?'selected':''}>ODD</option><option ${m.semester==='EVEN'?'selected':''}>EVEN</option></select></div>
      <div class="field"><label>Hall Number</label><input type="text" value="${esc(m.hall)}" onchange="updOtherMeta('${rec.id}','hall',this.value)"></div>
      <div class="field"><label>W.E.F.</label><input type="text" value="${esc(m.wef)}" onchange="updOtherMeta('${rec.id}','wef',this.value)"></div>
      <div class="field"><label>Class Advisor</label><input type="text" list="otherFacList" value="${esc(m.advisor)}" onchange="updOtherMeta('${rec.id}','advisor',this.value)"></div>
    </div>
    <datalist id="otherFacList">${state.faculty.map(f=>`<option value="${esc(f.name)}">`).join('')}</datalist>
    <div style="margin-top:10px;"><button class="btn danger small" onclick="removeOtherTT('${rec.id}')">Delete this timetable</button></div>
  </div>`;
}
async function updOtherMeta(id,field,val){
  const rec=state.otherTT.find(r=>r.id===id); if(!rec) return;
  rec.meta[field]=val;
  await persistOtherTT(rec, ['meta']);
  renderTabBody();
}
function otherSubjectPanel(rec){
  return `
  <div class="panel">
    <h3>Subjects (manual entry)</h3>
    <div class="tt-wrap">
      <table class="datatable">
        <thead><tr><th>Code</th><th>Title</th><th>Shortcut</th><th>Faculty</th><th>Colour</th><th>Work Load</th><th></th></tr></thead>
        <tbody>
        ${rec.subjects.map(s=>`
          <tr>
            <td><input type="text" style="width:80px;" value="${esc(s.code)}" onchange="updOtherSubj('${rec.id}','${s.id}','code',this.value)"></td>
            <td><input type="text" style="min-width:170px;" value="${esc(s.title)}" onchange="updOtherSubj('${rec.id}','${s.id}','title',this.value)"></td>
            <td><input type="text" style="width:70px;font-weight:700;" value="${esc(s.shortcut)}" onchange="updOtherSubj('${rec.id}','${s.id}','shortcut',this.value)"></td>
            <td><input type="text" list="otherFacList" style="min-width:160px;" value="${esc(s.faculty)}" onchange="updOtherSubj('${rec.id}','${s.id}','faculty',this.value)"></td>
            <td><input type="color" value="${s.color}" style="width:36px;height:28px;border:none;background:none;" onchange="updOtherSubj('${rec.id}','${s.id}','color',this.value)"></td>
            <td><input type="number" style="width:60px;" value="${s.workload||''}" onchange="updOtherSubj('${rec.id}','${s.id}','workload',this.value)"></td>
            <td><button class="btn small danger" onclick="delOtherSubj('${rec.id}','${s.id}')">✕</button></td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>
    <div style="margin-top:10px;"><button class="btn primary small" onclick="addOtherSubj('${rec.id}')">+ Add Subject Row</button></div>
  </div>`;
}
async function addOtherSubj(recId){
  const rec = state.otherTT.find(r=>r.id===recId); if(!rec) return;
  rec.subjects.push({id:uid('os'), code:"", title:"", shortcut:"", faculty:"", workload:"", color:nextColor(rec.subjects)});
  await persistOtherTT(rec, ['subjects']);
  renderTabBody();
}
async function updOtherSubj(recId,sid,field,val){
  const rec = state.otherTT.find(r=>r.id===recId); if(!rec) return;
  const s = rec.subjects.find(x=>x.id===sid); if(!s) return;
  s[field]=val;
  await persistOtherTT(rec, ['subjects']);
  renderTabBody();
}
async function delOtherSubj(recId,sid){
  const rec = state.otherTT.find(r=>r.id===recId); if(!rec) return;
  rec.subjects = rec.subjects.filter(s=>s.id!==sid);
  await persistOtherTT(rec, ['subjects']);
  renderTabBody();
}

/* ================= SHARED: GRID PANEL ================= */
function getRec(scope,key){
  if(scope==='class') return activeOurRec(key);
  return state.otherTT.find(r=>r.id===key);
}
function gridPanel(rec, scope, key, displayLabel){
  const headerHtml = renderReportHeader(rec, scope, displayLabel);
  let rows="";
  DAYS.forEach((day,dayIdx)=>{
    let cells = `<td class="day-cell">${day}</td>`;
    COLUMNS.forEach(col=>{
      if(col.type==='break'){
        if(dayIdx===0) cells += `<td class="break-cell" rowspan="${DAYS.length}">${nl2br(col.label)}</td>`;
        return;
      }
      const val = rec.grid[day][col.idx]||"";
      const subj = val? findSubjectByShortcut(rec,val) : null;
      const bg = subj? subj.color : "transparent";
      const conflict = val? (findConflicts(subj?subj.faculty:null, day, col.idx, rec).length>0) : false;
      cells += `<td style="background:${bg};${conflict?'box-shadow:inset 0 0 0 2px var(--danger);':''}">
        <button class="cell-btn" onclick="openCellPicker('${scope}','${key}','${day}',${col.idx})">${esc(val)||'&nbsp;'}</button>
      </td>`;
    });
    rows += `<tr>${cells}</tr>`;
  });
  const headRow = `<tr><th>Timing</th>` + COLUMNS.map(col=> col.type==='period' ? `<th>${nl2br(col.label)}<br><span style="font-weight:400;">Hour ${col.hour}</span></th>` : `<th style="writing-mode:vertical-rl;transform:rotate(180deg);">${nl2br(col.tag)}</th>`).join('') + `</tr>`;
  return `
  <div class="panel">
    <h3>Timetable Grid (54 boxes)</h3>
    ${headerHtml}
    <div class="tt-wrap"><table class="grid"><thead>${headRow}</thead><tbody>${rows}</tbody></table></div>
    <div class="small-note" style="margin-top:8px;">Click a box, choose a shortcut. Empty hours stay blank. Red outline = same faculty already teaching elsewhere at that time.</div>
  </div>`;
}
function renderReportHeader(rec, scope, displayLabel){
  const m = rec.meta;
  const [ayFrom,ayTo] = acadParts(m.academicYear);
  const progLine = scope==='class' ? `<div><b>PROGRAMME:</b> ${esc(m.programme)}</div>` : `<div><b>DEPARTMENT:</b> ${esc(m.department||'')} &nbsp; <b>PROGRAMME:</b> ${esc(m.programme)}</div>`;
  const yearSecLine = scope==='class' ? `<div><b>YEAR/SEMESTER/SECTION:</b> ${esc(displayLabel)}</div>` : `<div><b>YEAR:</b> ${esc(m.year)} &nbsp; <b>SECTION:</b> ${esc(m.section)} &nbsp; <b>SEM:</b> ${esc(m.semester)}</div>`;
  return `
  <div class="report-head">
    <div class="accode">${COLLEGE.ac}<br>${COLLEGE.rev}</div>
    <div class="top-row"><div style="width:40px;"></div>
      <div class="collegename">
        <div class="cname">${COLLEGE.name}</div>
        <div class="caption">${COLLEGE.caption}</div>
        <div class="caption">${esc(COLLEGE.address)}</div>
        <div class="caption" style="font-weight:700;margin-top:3px;">TIME TABLE FOR Academic Year: ${esc(ayFrom)} - ${esc(ayTo)} (${esc(m.semester)} SEMESTER)</div>
        <div class="caption" style="font-weight:700;">${COLLEGE.dept}</div>
      </div><div style="width:40px;"></div>
    </div>
    <div class="report-meta">
      <div>${progLine}${yearSecLine}</div>
      <div style="text-align:right;"><div><b>W.E.F:</b> ${esc(m.wef)||'—'}</div><div><b>HALL NUMBER:</b> ${esc(m.hall)||'—'}</div></div>
    </div>
    <div class="report-advisor">Class Advisor: ${esc(m.advisor)||'—'}</div>
  </div>`;
}
function renderReportFooter(){
  return `<div class="footer-block"><div class="sig">Time Table In-Charge</div><div class="sig">Co-Ordinator</div><div class="sig">HoD</div><div class="sig">Principal</div></div>`;
}

/* ================= CELL PICKER ================= */
let pickerCtx=null;
function openCellPicker(scope,key,day,periodIdx){
  const rec = getRec(scope,key); if(!rec) return;
  pickerCtx = {scope,key,day,periodIdx};
  const current = rec.grid[day][periodIdx]||"";
  document.getElementById('modalBody').innerHTML = `
    <h3 style="margin-top:0;">${day} · Period box</h3>
    <div class="small-note">Choose a subject shortcut, or type a custom value.</div>
    <div class="chip-grid">
      ${rec.subjects.filter(s=>s.shortcut).map(s=>`<button class="chip" style="background:${s.color}" onclick="pickShortcut('${esc(s.shortcut).replace(/'/g,"\\'")}')">${esc(s.shortcut)}</button>`).join('') || '<span class="small-note">No subjects with shortcuts yet.</span>'}
    </div>
    <div class="field"><label>Custom value</label><input type="text" id="customCellVal" value="${esc(current)}"></div>
    <div class="row" style="margin-top:14px;justify-content:flex-end;">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn danger" onclick="clearCell()">Clear</button>
      <button class="btn primary" onclick="saveCellCustom()">Save</button>
    </div>`;
  document.getElementById('modalBg').classList.add('open');
}
function closeModal(){ document.getElementById('modalBg').classList.remove('open'); pickerCtx=null; }
function pickShortcut(sc){ setCell(sc); }
function saveCellCustom(){ setCell(document.getElementById('customCellVal').value.trim()); }
function clearCell(){ setCell(""); }
async function setCell(val){
  if(!pickerCtx) return;
  const { scope, key, day, periodIdx } = pickerCtx;
  const rec = getRec(scope,key); if(!rec) return;
  rec.grid[day][periodIdx]=val;
  const table = scope==='class' ? 'our_versions' : 'other_timetables';
  await sbCall(supabase.from(table).update({grid: rec.grid}).eq('id', rec.id), "Save grid");
  if(val){
    const subj = findSubjectByShortcut(rec, val);
    if(subj && subj.faculty){
      const conflicts = findConflicts(subj.faculty, day, periodIdx, rec);
      if(conflicts.length) alert("⚠ Faculty conflict: "+subj.faculty+" is already assigned at this time in:\\n"+conflicts.join("\\n"));
    }
  }
  closeModal(); renderTabBody();
}

/* ================= EXPORT (single timetable) ================= */
function exportPanel(scope,key,label){
  return `<div class="panel"><h3>Export</h3><div class="row">
    <button class="btn primary" onclick="exportPDF('${scope}','${key}')">⬇ Download PDF</button>
    <button class="btn primary" onclick="exportExcel('${scope}','${key}')">⬇ Download Excel</button>
  </div></div>`;
}
function gridToAutotableBody(rec){
  const head = ["Day"].concat(COLUMNS.map(c=> c.type==='period'? c.label.replace('\n',' ')+' (H'+c.hour+')' : c.tag));
  const body = DAYS.map(day=>{ const row=[day]; COLUMNS.forEach(col=>{ if(col.type==='break'){row.push(col.tag);return;} row.push(rec.grid[day][col.idx]||""); }); return row; });
  return {head:[head], body};
}
function subjectsToAutotableBody(rec){
  const head=[["SNO","SUBJECT CODE","SUBJECT TITLE","FACULTY NAME","WORK LOAD"]];
  const body = rec.subjects.map((s,i)=>[i+1,s.code,s.title,s.faculty,s.workload||'']);
  return {head,body};
}
function pdfHeaderBlock(doc, rec, scope, label, y0){
  const [ayFrom,ayTo] = acadParts(rec.meta.academicYear);
  doc.setFont("helvetica","bold"); doc.setFontSize(13);
  doc.text(COLLEGE.name, doc.internal.pageSize.getWidth()/2, y0, {align:"center"});
  doc.setFontSize(8); doc.setFont("helvetica","normal");
  doc.text(COLLEGE.caption, doc.internal.pageSize.getWidth()/2, y0+5, {align:"center"});
  doc.text(COLLEGE.address, doc.internal.pageSize.getWidth()/2, y0+9.5, {align:"center"});
  doc.setFont("helvetica","bold"); doc.setFontSize(9);
  const m = rec.meta;
  doc.text(`TIME TABLE FOR Academic Year: ${ayFrom} - ${ayTo} (${m.semester} SEMESTER)`, doc.internal.pageSize.getWidth()/2, y0+14.5, {align:"center"});
  doc.text(COLLEGE.dept, doc.internal.pageSize.getWidth()/2, y0+19, {align:"center"});
  doc.setFontSize(8); doc.text(COLLEGE.ac+"  "+COLLEGE.rev, doc.internal.pageSize.getWidth()-25, y0-2);
  doc.setFont("helvetica","normal"); doc.setFontSize(9);
  let line1 = scope==='class'? `PROGRAMME: ${m.programme}` : `DEPARTMENT: ${m.department||''}   PROGRAMME: ${m.programme}`;
  let line2 = scope==='class'? `YEAR/SEMESTER/SECTION: ${label}` : `YEAR: ${m.year}   SECTION: ${m.section}`;
  doc.text(line1, 14, y0+26); doc.text(line2, 14, y0+31);
  doc.text(`W.E.F: ${m.wef||'-'}`, doc.internal.pageSize.getWidth()-60, y0+26);
  doc.text(`HALL NUMBER: ${m.hall||'-'}`, doc.internal.pageSize.getWidth()-60, y0+31);
  doc.setFont("helvetica","bold"); doc.text(`Class Advisor: ${m.advisor||'-'}`, 14, y0+36); doc.setFont("helvetica","normal");
  return y0+42;
}
function pdfFooterBlock(doc, pageInfo){
  const w = doc.internal.pageSize.getWidth(), h = doc.internal.pageSize.getHeight(), y = h-16;
  doc.setFontSize(8); doc.setFont("helvetica","bold");
  const labels=["Time Table In-Charge","Co-Ordinator","HoD","Principal"]; const slotW=(w-28)/4;
  labels.forEach((lab,i)=>{ const x=14+i*slotW; doc.line(x,y,x+slotW-10,y); doc.text(lab,x,y+4); });
  doc.setFont("helvetica","normal"); doc.setFontSize(7);
  doc.text(`Report Gen date: ${nowStr()}`, 14, h-4);
  doc.text(`Page ${pageInfo.page}/${pageInfo.total}`, w/2, h-4, {align:"center"});
  doc.text(`ReportID: ${pageInfo.id}`, w-14, h-4, {align:"right"});
}
function drawTimetablePDFPage(doc, rec, scope, label, isFirst){
  if(!isFirst) doc.addPage();
  let y = pdfHeaderBlock(doc, rec, scope, label, 14);
  const gt = gridToAutotableBody(rec);
  doc.autoTable({ head:gt.head, body:gt.body, startY:y, styles:{fontSize:6.5,cellPadding:1.1,halign:'center',valign:'middle'}, headStyles:{fillColor:[238,241,247],textColor:20,fontStyle:'bold'}, margin:{left:14,right:14} });
  let y2 = doc.lastAutoTable.finalY+6;
  const st = subjectsToAutotableBody(rec);
  doc.autoTable({ head:st.head, body:st.body, startY:y2, styles:{fontSize:8,cellPadding:1.6}, headStyles:{fillColor:[238,241,247],textColor:20,fontStyle:'bold'}, margin:{left:14,right:14} });
}
function exportPDF(scope,key){
  const rec = getRec(scope,key); if(!rec) return;
  const label = scope==='class'? key : otherLabel(rec);
  const doc = new jsPDF({orientation:"landscape", unit:"mm", format:"a4"});
  const rid = reportId();
  drawTimetablePDFPage(doc, rec, scope, label, true);
  const total = doc.internal.getNumberOfPages();
  for(let p=1;p<=total;p++){ doc.setPage(p); pdfFooterBlock(doc,{page:p,total,id:rid}); }
  doc.save(`Timetable_${label.replace(/\s+/g,'_')}.pdf`);
}
function timetableSheetAOA(rec, scope, label){
  const m = rec.meta; const [ayFrom,ayTo] = acadParts(m.academicYear);
  const aoa=[];
  aoa.push([COLLEGE.name]); aoa.push([COLLEGE.caption]); aoa.push([COLLEGE.address]);
  aoa.push([`TIME TABLE FOR Academic Year: ${ayFrom} - ${ayTo} (${m.semester} SEMESTER)`]);
  aoa.push([COLLEGE.dept]);
  aoa.push([ scope==='class'? `PROGRAMME: ${m.programme}` : `DEPARTMENT: ${m.department||''}   PROGRAMME: ${m.programme}`, '', '', `W.E.F: ${m.wef||''}` ]);
  aoa.push([ scope==='class'? `YEAR/SEMESTER/SECTION: ${label}` : `YEAR: ${m.year}  SECTION: ${m.section}`, '', '', `HALL NUMBER: ${m.hall||''}` ]);
  aoa.push([`Class Advisor: ${m.advisor||''}`]); aoa.push([]);
  const head = ["Day"].concat(COLUMNS.map(c=> c.type==='period'? c.label.replace('\n',' ')+' (H'+c.hour+')' : c.tag));
  aoa.push(head);
  DAYS.forEach(day=>{ const row=[day]; COLUMNS.forEach(col=>{ if(col.type==='break'){row.push(col.tag);return;} row.push(rec.grid[day][col.idx]||""); }); aoa.push(row); });
  aoa.push([]); aoa.push(["SNO","SUBJECT CODE","SUBJECT TITLE","FACULTY NAME","WORK LOAD"]);
  rec.subjects.forEach((s,i)=> aoa.push([i+1,s.code,s.title,s.faculty,s.workload||'']));
  aoa.push([]); aoa.push(["Time Table In-Charge","Co-Ordinator","HoD","Principal"]);
  aoa.push([`Report Gen: ${nowStr()}`,"","", `ReportID: ${reportId()}`]);
  return aoa;
}
function sheetFromAOA(aoa){
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = new Array(14).fill({wch:14});
  ws['!merges'] = [{s:{r:0,c:0},e:{r:0,c:13}},{s:{r:1,c:0},e:{r:1,c:13}},{s:{r:2,c:0},e:{r:2,c:13}},{s:{r:3,c:0},e:{r:3,c:13}},{s:{r:4,c:0},e:{r:4,c:13}}];
  return ws;
}
function exportExcel(scope,key){
  const rec = getRec(scope,key); if(!rec) return;
  const label = scope==='class'? key : otherLabel(rec);
  const ws = sheetFromAOA(timetableSheetAOA(rec,scope,label));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, label.substring(0,28)||"Timetable");
  XLSX.writeFile(wb, `Timetable_${label.replace(/\s+/g,'_')}.xlsx`);
}

/* ================= REPORTS TAB ================= */
function viewReports(){
  return `
  <div class="panel">
    <h2>Reports</h2>
  </div>
  <div class="panel">
    <h3>Class-wise Report</h3>
    <div class="select-list">
      ${CLASS_LIST.map(c=>`<label><input type="checkbox" ${ui.repClassSel.includes(c)?'checked':''} onchange="toggleClassSel('${c}',this.checked)"> ${c} ${activeOurRec(c)? '':'<span class="small-note">(no version)</span>'}</label>`).join('')}
    </div>
    <div class="row" style="margin-top:12px;">
      <button class="btn primary" ${ui.repClassSel.length===0?'disabled':''} onclick="exportClassWisePDF()">⬇ Combined PDF (${ui.repClassSel.length})</button>
      <button class="btn primary" ${ui.repClassSel.length===0?'disabled':''} onclick="exportClassWiseExcel()">⬇ Combined Excel (${ui.repClassSel.length} sheets)</button>
    </div>
    <div class="small-note" style="margin-top:6px;">Uses each class's active version.</div>
  </div>
  <div class="panel">
    <h3>Individual Faculty Report</h3>
    <div class="select-list">
      ${state.faculty.map(f=>`<label><input type="checkbox" ${ui.repFacMulti.includes(f.name)?'checked':''} onchange="toggleFacSel('${esc(f.name).replace(/'/g,"\\'")}',this.checked)"> ${esc(f.name)}</label>`).join('') || '<span class="small-note">Add faculty in Faculty Master first.</span>'}
    </div>
    <div class="row" style="margin-top:12px;">
      <button class="btn primary" ${ui.repFacMulti.length===0?'disabled':''} onclick="exportFacultyPDFReport()">⬇ Combined PDF (${ui.repFacMulti.length})</button>
      <button class="btn primary" ${ui.repFacMulti.length===0?'disabled':''} onclick="exportFacultyExcelReport()">⬇ Combined Excel (${ui.repFacMulti.length} sheets)</button>
    </div>
    <div class="small-note" style="margin-top:6px;">Each faculty's report shows three grids — Our Department only, Other Department only, and a Combined view — plus a subject-allotment/workload table with a grand total.</div>
  </div>`;
}
function toggleClassSel(c,checked){ if(checked){ if(!ui.repClassSel.includes(c)) ui.repClassSel.push(c);} else ui.repClassSel=ui.repClassSel.filter(x=>x!==c); renderTabBody(); }
function toggleFacSel(name,checked){ if(checked){ if(!ui.repFacMulti.includes(name)) ui.repFacMulti.push(name);} else ui.repFacMulti=ui.repFacMulti.filter(x=>x!==name); renderTabBody(); }

function exportClassWisePDF(){
  const doc = new jsPDF({orientation:"landscape", unit:"mm", format:"a4"});
  const rid = reportId();
  let first=true;
  ui.repClassSel.forEach(c=>{
    const rec = activeOurRec(c); if(!rec) return;
    drawTimetablePDFPage(doc, rec, 'class', c, first); first=false;
  });
  const total = doc.internal.getNumberOfPages();
  for(let p=1;p<=total;p++){ doc.setPage(p); pdfFooterBlock(doc,{page:p,total,id:rid}); }
  doc.save("Class_Wise_Timetables.pdf");
}
function exportClassWiseExcel(){
  const wb = XLSX.utils.book_new();
  ui.repClassSel.forEach(c=>{
    const rec = activeOurRec(c); if(!rec) return;
    const ws = sheetFromAOA(timetableSheetAOA(rec,'class',c));
    let name=c.replace(/[\\/?*\[\]:]/g,'').substring(0,28), base=name, n=1;
    while(wb.SheetNames.includes(name)) name=base.substring(0,25)+"_"+(n++);
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  XLSX.writeFile(wb, "Class_Wise_Timetables.xlsx");
}

function drawFacultyPDFPage(doc, name, isFirst){
  if(!isFirst) doc.addPage();
  doc.setFont("helvetica","bold"); doc.setFontSize(13);
  doc.text(COLLEGE.name, doc.internal.pageSize.getWidth()/2, 14, {align:"center"});
  doc.setFontSize(8); doc.setFont("helvetica","normal");
  doc.text(COLLEGE.caption, doc.internal.pageSize.getWidth()/2, 19, {align:"center"});
  doc.setFont("helvetica","bold"); doc.setFontSize(10);
  doc.text("INDIVIDUAL FACULTY TIMETABLE", doc.internal.pageSize.getWidth()/2, 25, {align:"center"});
  doc.setFontSize(9); doc.text(`Faculty Name: ${name}`, 14, 33);

  let y = 37;
  const sections = [
    {label:"Our Department", scope:'class'},
    {label:"Other Department", scope:'other'},
    {label:"Combined (Our + Other)", scope:'all'}
  ];
  sections.forEach(sec=>{
    const {head, body} = weekToAutotable(facultyWeek(name, sec.scope));
    doc.setFont("helvetica","bold"); doc.setFontSize(8.5); doc.text(sec.label, 14, y);
    doc.autoTable({ head, body, startY:y+1.5, styles:{fontSize:6,cellPadding:0.9,halign:'center',valign:'middle'}, headStyles:{fillColor:[238,241,247],textColor:20,fontStyle:'bold'}, margin:{left:14,right:14} });
    y = doc.lastAutoTable.finalY + 6;
  });

  const allotRows = facultyAllotmentRows(name);
  const total = allotRows.reduce((sum,r)=> sum + (parseFloat(r.workload)||0), 0);
  doc.setFont("helvetica","bold"); doc.setFontSize(9); doc.text("Subject Allotment", 14, y); y+=2;
  doc.autoTable({ head:[["Source","Code","Subject","Work Load"]], body: allotRows.map(r=>[r.source,r.code,r.title,r.workload||'']), startY:y+2, styles:{fontSize:7.5,cellPadding:1.4}, headStyles:{fillColor:[238,241,247],textColor:20,fontStyle:'bold'}, margin:{left:14,right:14},
    foot:[["","","Total", total]], footStyles:{fontStyle:'bold', fillColor:[238,241,247]} });
}
function exportFacultyPDFReport(){
  const doc = new jsPDF({orientation:"landscape", unit:"mm", format:"a4"});
  const rid = reportId();
  ui.repFacMulti.forEach((name,i)=> drawFacultyPDFPage(doc, name, i===0));
  const total = doc.internal.getNumberOfPages();
  for(let p=1;p<=total;p++){ doc.setPage(p); pdfFooterBlock(doc,{page:p,total,id:rid}); }
  doc.save("Faculty_Reports.pdf");
}
function facultySheetAOA(name){
  const aoa=[]; aoa.push([COLLEGE.name]); aoa.push([COLLEGE.caption]); aoa.push(["INDIVIDUAL FACULTY TIMETABLE"]); aoa.push([`Faculty Name: ${name}`]); aoa.push([]);
  const head=["Day"].concat(COLUMNS.map(c=> c.type==='period'? c.label.replace('\n',' ')+' (H'+c.hour+')' : c.tag));
  const sections = [
    {label:"OUR DEPARTMENT", scope:'class'},
    {label:"OTHER DEPARTMENT", scope:'other'},
    {label:"COMBINED (OUR + OTHER)", scope:'all'}
  ];
  sections.forEach(sec=>{
    aoa.push([sec.label]); aoa.push(head);
    const week = facultyWeek(name, sec.scope);
    DAYS.forEach(day=>{ const row=[day]; COLUMNS.forEach(col=>{ if(col.type==='break'){row.push(col.tag);return;} const cd=week[day][col.idx]; row.push(cd? `${cd.shortcut} (${cd.text})`:''); }); aoa.push(row); });
    aoa.push([]);
  });
  aoa.push(["Subject Allotment"]); aoa.push(["Source","Code","Subject","Work Load"]);
  const allotRows = facultyAllotmentRows(name);
  allotRows.forEach(r=> aoa.push([r.source,r.code,r.title,r.workload||'']));
  const total = allotRows.reduce((sum,r)=> sum + (parseFloat(r.workload)||0), 0);
  aoa.push(["","","Total",total]);
  aoa.push([]); aoa.push(["Time Table In-Charge","Co-Ordinator","HoD","Principal"]); aoa.push([`Report Gen: ${nowStr()}`]);
  return aoa;
}
function exportFacultyExcelReport(){
  const wb = XLSX.utils.book_new();
  ui.repFacMulti.forEach((name,i)=>{
    const ws = sheetFromAOA(facultySheetAOA(name));
    let sn=name.replace(/[\\/?*\[\]:]/g,'').substring(0,28), base=sn, n=1;
    while(wb.SheetNames.includes(sn)) sn=base.substring(0,25)+"_"+(n++);
    XLSX.utils.book_append_sheet(wb, ws, sn);
  });
  XLSX.writeFile(wb, "Faculty_Reports.xlsx");
}

/* ================= BACKUP (export only — Supabase is the source of truth) ================= */
function exportBackup(){
  const blob = new Blob([JSON.stringify(state,null,2)], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download="ttms_backup.json"; a.click();
  URL.revokeObjectURL(url);
}
async function reloadFromDatabase(){
  try{ await loadAllFromSupabase(); render(); }
  catch(err){ alert("Could not reload from the database: "+err.message); }
}

/* ================= AUTH ================= */
async function doLogin(){
  const email = document.getElementById('login_email').value.trim();
  const password = document.getElementById('login_pw').value;
  const errEl = document.getElementById('login_err');
  errEl.textContent = "";
  const { error } = await supabase.auth.signInWithPassword({ email, password });
  if(error){ errEl.textContent = error.message; return; }
  boot();
}
async function doLogout(){
  await supabase.auth.signOut();
  boot();
}
function renderLogin(){
  document.getElementById('app').innerHTML = `
  <div class="login-wrap"><div class="login-card">
    <h1>Timetable Management System</h1>
    <div class="sub">Sign in with the account you created in your Supabase project.</div>
    <div class="field"><label>Email</label><input type="text" id="login_email"></div>
    <div class="field" style="margin-top:8px;"><label>Password</label><input type="text" id="login_pw" style="-webkit-text-security:disc;text-security:disc;"></div>
    <div id="login_err" class="login-err"></div>
    <button class="btn primary" style="width:100%;margin-top:10px;" onclick="doLogin()">Sign in</button>
  </div></div>`;
}
function renderLoading(msg){
  document.getElementById('app').innerHTML = `<div class="loading-note">${esc(msg||"Loading…")}</div>`;
}

/* ================= BOOT ================= */
async function boot(){
  const { data: { session } } = await supabase.auth.getSession();
  if(!session){ renderLogin(); return; }
  renderLoading("Loading your data…");
  try{
    await loadAllFromSupabase();
    render();
  }catch(err){
    renderLoading("Could not load data: "+err.message+" — check your Supabase connection and RLS policies.");
  }
}

/* ================= EXPOSE TO WINDOW (inline HTML handlers need these) ================= */
Object.assign(window, {
  ui, state,
  setTab, renderTabBody,
  addFaculty, removeFaculty, toggleMaskAll, exportFacultyDirectoryPDF, exportFacultyDirectoryExcel,
  setSubjClass, addSubjectMasterRow, updSubjMaster, delSubjMaster, importSubjects,
  setAllotClass, updAllotRow,
  setOurClass, createOurVersion, setOurActive, updOurMeta, updOurSubj,
  addOtherTT, setOtherActive, removeOtherTT, updOtherMeta, addOtherSubj, updOtherSubj, delOtherSubj,
  openCellPicker, closeModal, pickShortcut, saveCellCustom, clearCell,
  exportPDF, exportExcel,
  toggleClassSel, toggleFacSel, exportClassWisePDF, exportClassWiseExcel, exportFacultyPDFReport, exportFacultyExcelReport,
  exportBackup, reloadFromDatabase, doLogin, doLogout
});

/* ================= INIT ================= */
boot();
