
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  STATE
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
let jobId = null, students = [], patches = {}, activeStudent = null, pollTimer = null;

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  NAVIGATION
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const VT = {upload:'Upload & Convert',review:'Review & Edit',history:'Output History',guide:'Usage Guide'};
const VS = {upload:'Drop PDF transcripts to begin',review:'Inspect, edit and confirm student records',history:'Download generated Excel files',guide:'How to use TranscriptX'};

function showView(v){
  document.querySelectorAll('.view').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(x=>x.classList.remove('active'));
  document.getElementById('view-'+v).classList.add('active');
  document.getElementById('nav-'+v).classList.add('active');
  document.getElementById('tb-title').textContent=VT[v];
  document.getElementById('tb-sub').textContent=VS[v];
  if(v==='history') loadHistory();
  if(v==='review' && jobId) loadStudents();
}
function switchTab(t){
  ['students','errors','log'].forEach(n=>{
    document.getElementById('tab-'+n).classList.toggle('active',n===t);
    document.getElementById('tc-'+n).classList.toggle('hidden',n!==t);
  });
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  FILE HANDLING
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
let selFiles=[];
const dz=document.getElementById('dropzone');
const fi=document.getElementById('file-input');

dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over')});
dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');addFiles([...e.dataTransfer.files].filter(f=>f.name.endsWith('.pdf')))});
fi.addEventListener('change',()=>{addFiles([...fi.files]);fi.value=''});

function addFiles(files){files.forEach(f=>{if(!selFiles.find(x=>x.name===f.name))selFiles.push(f)});renderFiles()}
function removeFile(n){selFiles=selFiles.filter(f=>f.name!==n);renderFiles()}
function clearFiles(){selFiles=[];renderFiles()}
function renderFiles(){
  document.getElementById('files-list').innerHTML=selFiles.map(f=>`<span class="file-tag">ðŸ“„ ${f.name}<button onclick="removeFile('${f.name}')">âœ•</button></span>`).join('');
  document.getElementById('btn-convert').disabled=selFiles.length===0;
  document.getElementById('upload-note').textContent=selFiles.length?`${selFiles.length} file(s) ready`:'';
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  CONVERSION
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function startConversion(){
  if(!selFiles.length) return;
  const fd=new FormData();
  selFiles.forEach(f=>fd.append('files',f));
  fd.append('school_id',document.getElementById('school-id').value);

  document.getElementById('btn-convert').disabled=true;
  setChk('upload',true);
  document.getElementById('prog-card').classList.remove('hidden');
  setStatus('parsing','Parsingâ€¦');

  try{
    const r=await fetch('/api/upload',{method:'POST',body:fd});
    const d=await r.json();
    if(d.error){toast(d.error,'error');document.getElementById('btn-convert').disabled=false;return}
    jobId=d.job_id; patches={};
    startPoll();
  }catch(e){toast('Upload failed: '+e.message,'error');document.getElementById('btn-convert').disabled=false}
}

function startPoll(){clearInterval(pollTimer);pollTimer=setInterval(poll,1000)}

async function poll(){
  if(!jobId) return;
  try{
    const r=await fetch('/api/job/'+jobId);
    const d=await r.json();
    updateProg(d);
    if(d.status==='awaiting_confirmation'){
      clearInterval(pollTimer);
      setChk('parse',true);setChk('validate',true);
      setStatus('ready','Ready');setFoot('#4ade80','â— Ready');
      loadStudents(); showView('review'); loadFormB();
      toast(`Parsed ${d.student_count} students Â· ${d.error_count} with issues`,d.error_count?'warning':'success');
    }else if(d.status==='done'){
      clearInterval(pollTimer);
      setChk('excel',true);setStatus('done','Done');setFoot('#4ade80','â— Done');
      loadStudents();
      toast('âœ… Excel ready â€” downloadingâ€¦','success');
      setTimeout(()=>{window.location.href='/api/job/'+jobId+'/download'},600);
      loadHistory();
    }else if(d.status==='failed'){
      clearInterval(pollTimer);
      setStatus('fail','Failed');setFoot('#f87171','â— Error');
      toast('Conversion failed â€” check Parse Log','error');
    }
  }catch(e){}
}

function updateProg(d){
  const p=d.progress||0;
  document.getElementById('prog-bar').style.width=p+'%';
  document.getElementById('prog-pct').textContent=p+'%';
  const stages={parsing:'Parsing PDF pagesâ€¦',building:'Building Excelâ€¦',awaiting_confirmation:'Validation complete',done:'Complete',failed:'Failed'};
  document.getElementById('prog-stage').textContent=stages[d.status]||'Workingâ€¦';
  const lp=document.getElementById('log-panel');
  lp.innerHTML=(d.log||[]).slice(-30).map(e=>`<div class="le"><span class="lt">${e.t}</span><span class="l${e.level}">${esc(e.msg)}</span></div>`).join('');
  lp.scrollTop=lp.scrollHeight;
  if(d.error_count>0){
    document.getElementById('err-nav-badge').textContent=d.error_count;
    document.getElementById('err-nav-badge').classList.remove('hidden');
  }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  REVIEW
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function loadStudents(){
  if(!jobId) return;
  const r=await fetch('/api/job/'+jobId+'/students');
  const d=await r.json();
  students=d.students||[];
  const meta=d.meta||{}, errs=d.errors||{};

  document.getElementById('rev-empty').classList.add('hidden');
  document.getElementById('rev-content').classList.remove('hidden');
  document.getElementById('cbar').classList.remove('hidden');

  const tc=students.reduce((a,s)=>a+s.course_count,0);
  const ec=students.filter(s=>!s.valid).length;
  const wc=students.reduce((a,s)=>a+s.warnings.length,0);

  document.getElementById('stats-grid').innerHTML=`
    <div class="stat c-blue"><div class="stat-lbl">Students</div><div class="stat-val">${students.length}</div></div>
    <div class="stat c-blue"><div class="stat-lbl">Courses</div><div class="stat-val">${tc}</div></div>
    <div class="stat ${ec?'c-red':'c-green'}"><div class="stat-lbl">Errors</div><div class="stat-val">${ec}</div></div>
    <div class="stat ${wc?'c-amber':'c-green'}"><div class="stat-lbl">Warnings</div><div class="stat-val">${wc}</div></div>
    <div class="stat"><div class="stat-lbl">School ID</div><div class="stat-val">${meta.school_id||'â€“'}</div></div>
    <div class="stat"><div class="stat-lbl">Level</div><div class="stat-val">${meta.level||'â€“'}</div></div>`;

  document.getElementById('ci-students').textContent=students.length;
  document.getElementById('ci-courses').textContent=tc;
  document.getElementById('ci-errs').innerHTML=ec?`<span style="color:var(--red)">${ec} with errors</span>`:`<span style="color:var(--green)">All valid âœ“</span>`;
  document.getElementById('err-tab-badge').textContent=ec;

  let alertHtml='';
  if(ec>0) alertHtml+=`<div class="alert a-red"><div class="alert-icon">âš </div><div><div class="ab-title">${ec} record(s) have errors</div><div>Click any row to view and fix marks inline before generating Excel.</div></div></div>`;
  if(wc>0) alertHtml+=`<div class="alert a-amber"><div class="alert-icon">â„¹</div><div><div class="ab-title">${wc} parse warning(s)</div><div>Non-critical â€” see Parse Log tab for details.</div></div></div>`;
  if(!ec&&!wc) alertHtml=`<div class="alert a-green"><div class="alert-icon">âœ“</div><div><div class="ab-title">All records valid</div><div>Ready to generate Excel.</div></div></div>`;
  document.getElementById('rev-alerts').innerHTML=alertHtml;

  setChk('review',true);
  renderTable();
  renderErrors(errs);
  renderFullLog();
}

let filtered=[];
function filterStudents(){
  const q=document.getElementById('search').value.toLowerCase();
  const st=document.getElementById('flt-status').value;
  filtered=students.filter(s=>{
    const mq=!q||s.matricule.toLowerCase().includes(q)||s.name.toLowerCase().includes(q);
    const ms=st==='all'||(st==='valid'&&s.valid)||(st==='errors'&&!s.valid);
    return mq&&ms;
  });
  document.getElementById('count-lbl').textContent=`${filtered.length} of ${students.length}`;
  renderTable(filtered);
}

function renderTable(list){
  list=list||students; filtered=list;
  document.getElementById('count-lbl').textContent=`${list.length} of ${students.length}`;
  document.getElementById('students-tbody').innerHTML=list.map(s=>{
    const sc1=s.semester_counts['1']||0, sc2=s.semester_counts['2']||0;
    const pd=patches[s.matricule]?'âœ ':'';
    return `<tr style="cursor:pointer" onclick="openDetail('${s.matricule}')">
      <td class="mono" style="font-size:12px">${pd}${s.matricule}</td>
      <td>${s.name}</td><td>${s.level||'â€“'}</td><td>${sc1}</td><td>${sc2}</td>
      <td><strong>${s.course_count}</strong></td>
      <td>${s.valid?'':'<span class="dot-r"></span>'}${s.warnings.length?'<span class="dot-a"></span>':''}
        <span class="badge ${s.valid?'bg-green':'bg-red'}">${s.valid?'âœ“ Valid':'âœ— Errors'}</span></td>
      <td><button class="btn btn-sm btn-outline" onclick="event.stopPropagation();openDetail('${s.matricule}')">View â†’</button></td>
    </tr>`;
  }).join('')||'<tr><td colspan="8" style="text-align:center;padding:2rem;color:var(--g5)">No students match filter</td></tr>';
}

function renderErrors(errs){
  const c=document.getElementById('errors-list');
  const keys=Object.keys(errs);
  if(!keys.length){c.innerHTML=`<div class="alert a-green"><div class="alert-icon">âœ“</div><div><div class="ab-title">No validation errors</div></div></div>`;return}
  c.innerHTML=keys.map(mat=>{
    const s=students.find(x=>x.matricule===mat);
    return `<div class="alert a-red" style="margin-bottom:.75rem">
      <div class="alert-icon">âœ—</div>
      <div style="flex:1"><div class="ab-title">${mat} â€” ${s?s.name:''}</div>
        <ul>${errs[mat].map(e=>`<li>${esc(e)}</li>`).join('')}</ul>
        <button class="btn btn-sm btn-outline" style="margin-top:.5rem" onclick="openDetail('${mat}')">Edit inline â†’</button>
      </div>
    </div>`;
  }).join('');
}

function renderFullLog(){
  if(!jobId) return;
  fetch('/api/job/'+jobId).then(r=>r.json()).then(d=>{
    document.getElementById('full-log').innerHTML=(d.log||[]).map(e=>`<div class="le"><span class="lt">${e.t}</span><span class="l${e.level}">${esc(e.msg)}</span></div>`).join('');
  });
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  DETAIL PANEL
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function openDetail(mat){
  const s=students.find(x=>x.matricule===mat);
  if(!s) return;
  activeStudent=s;
  document.getElementById('dp-name').textContent=s.name;
  document.getElementById('dp-mat').textContent=`${s.matricule} Â· Level ${s.level||'?'} Â· ${s.specialty||''}`;
  document.getElementById('dp-badge').innerHTML=`<span class="badge ${s.valid?'bg-green':'bg-red'}">${s.valid?'âœ“ Valid':'âœ— Errors'}</span>`;
  document.getElementById('dp-save').classList.add('hidden');

  let html=`<div class="dp-meta">
    <div class="dm-item"><div class="dm-key">Faculty</div><div class="dm-val">${s.faculty||'â€“'}</div></div>
    <div class="dm-item"><div class="dm-key">Department</div><div class="dm-val">${s.department||'â€“'}</div></div>
    <div class="dm-item"><div class="dm-key">Academic Year</div><div class="dm-val">${s.academic_year||'â€“'}</div></div>
    <div class="dm-item"><div class="dm-key">Source Page</div><div class="dm-val">${s.source_page||'â€“'}</div></div>
  </div>`;

  if(s.errors.length) html+=`<div class="alert a-red mb3"><div class="alert-icon">âš </div><div><div class="ab-title">${s.errors.length} error(s)</div><ul>${s.errors.map(e=>`<li>${esc(e)}</li>`).join('')}</ul></div></div>`;
  if(s.warnings.length) html+=`<div class="alert a-amber mb3"><div class="alert-icon">â„¹</div><div><div class="ab-title">Parse warnings</div><ul>${s.warnings.map(w=>`<li>${esc(w)}</li>`).join('')}</ul></div></div>`;

  const sems=[...new Set(s.courses.map(c=>c.semester))].sort();
  for(const sem of sems){
    const sc=s.courses.filter(c=>c.semester===sem);
    const gpa=s.gpa_sem[String(sem)]||0;
    html+=`<div class="sem-sec">
      <div class="sem-hdr"><span>Semester ${sem}</span><span style="color:var(--blue);font-size:11px">GPA: ${gpa}</span></div>
      <div class="c-hdr"><span>Code</span><span>Title</span><span>CA /30</span><span>Exam /70</span><span>Total</span><span>Grade</span></div>`;
    for(const c of sc){
      const tc=c.total>=70?'mk-good':c.total>=50?'mk-pass':'mk-fail';
      const haserr=c.errors.length>0;
      const pk=`${c.code}_sem${c.semester}`;
      const pp=(patches[s.matricule]||{})[pk]||{};
      const cav=pp.ca!==undefined?pp.ca:c.ca;
      const exv=pp.exam!==undefined?pp.exam:c.exam;
      html+=`<div class="c-grid ${haserr?'has-err':''}">
        <span class="cc">${c.code}</span>
        <span class="ct" title="${esc(c.title)}">${esc(c.title)}</span>
        <input class="mark-inp" type="number" min="0" max="30" step=".5" value="${cav}"
          data-mat="${s.matricule}" data-code="${c.code}" data-sem="${c.semester}" data-field="ca"
          oninput="onEdit(this)">
        <input class="mark-inp" type="number" min="0" max="70" step=".5" value="${exv}"
          data-mat="${s.matricule}" data-code="${c.code}" data-sem="${c.semester}" data-field="exam"
          oninput="onEdit(this)">
        <span class="mark ${tc}" id="tot-${s.matricule.replace(/[^a-z0-9]/gi,'_')}-${c.code}-${c.semester}">${c.total.toFixed(1)}</span>
        <span class="badge ${gradeC(c.grade)}">${c.grade||'â€“'}</span>
      </div>`;
    }
    html+=`</div>`;
  }

  document.getElementById('dp-body').innerHTML=html;
  document.getElementById('overlay').classList.add('open');
}

function gradeC(g){if(!g)return 'bg-grey';if(g==='A')return 'bg-green';if(g.startsWith('B'))return 'bg-blue';if(g==='F')return 'bg-red';return 'bg-grey'}

function onEdit(inp){
  const mat=inp.dataset.mat, code=inp.dataset.code, sem=parseInt(inp.dataset.sem), field=inp.dataset.field;
  const val=parseFloat(inp.value), maxv=field==='ca'?30:70;
  inp.classList.remove('changed','invalid');
  if(isNaN(val)||val<0||val>maxv) inp.classList.add('invalid');
  else inp.classList.add('changed');

  if(!patches[mat]) patches[mat]={};
  const pk=`${code}_sem${sem}`;
  if(!patches[mat][pk]) patches[mat][pk]={};
  patches[mat][pk][field]=val;

  // live total
  const row=inp.closest('.c-grid');
  const ins=row.querySelectorAll('.mark-inp');
  const ca=parseFloat([...ins].find(i=>i.dataset.field==='ca').value)||0;
  const ex=parseFloat([...ins].find(i=>i.dataset.field==='exam').value)||0;
  const tid=`tot-${mat.replace(/[^a-z0-9]/gi,'_')}-${code}-${sem}`;
  const tel=document.getElementById(tid);
  if(tel){const t=ca+ex;tel.textContent=t.toFixed(1);tel.className=`mark ${t>=70?'mk-good':t>=50?'mk-pass':'mk-fail'}`}

  document.getElementById('dp-save').classList.remove('hidden');
  document.getElementById('ci-ovrd').classList.remove('hidden');
}

async function saveStudent(){
  if(!activeStudent||!jobId) return;
  const mat=activeStudent.matricule;
  const p=patches[mat]; if(!p) return;
  try{
    const r=await fetch(`/api/job/${jobId}/student/${encodeURIComponent(mat)}`,
      {method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});
    const d=await r.json();
    if(d.error){toast(d.error,'error');return}
    const idx=students.findIndex(s=>s.matricule===mat);
    if(idx>=0) students[idx]=d.student;
    activeStudent=d.student;
    delete patches[mat];
    document.getElementById('dp-save').classList.add('hidden');
    document.getElementById('dp-badge').innerHTML=`<span class="badge ${d.student.valid?'bg-green':'bg-red'}">${d.student.valid?'âœ“ Valid':'âœ— Errors'}</span>`;
    renderTable(); toast(`Saved changes for ${mat}`,'success');
    document.querySelectorAll('.mark-inp.changed').forEach(i=>i.classList.remove('changed'));
  }catch(e){toast('Save failed: '+e.message,'error')}
}

function closeDetail(e){
  if(e&&e.target!==document.getElementById('overlay')) return;
  document.getElementById('overlay').classList.remove('open');
  activeStudent=null;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  GENERATE EXCEL
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function dismissErrors(){if(!jobId)return;await fetch('/api/job/'+jobId+'/dismiss-errors',{method:'POST'});toast('Errors acknowledged â€” you can now generate Excel','warning')}

async function generateExcel(){
  if(!jobId) return;
  const btn=document.getElementById('btn-gen');
  btn.disabled=true; btn.innerHTML='<span class="spin"></span> Generatingâ€¦';
  const sid=document.getElementById('school-id').value;
  const ovs={};
  for(const [mat,pp] of Object.entries(patches)){
    ovs[mat]=[];
    for(const [key,vals] of Object.entries(pp)){
      const [code,sp]=key.split('_sem');
      ovs[mat].push({code,semester:parseInt(sp),...vals});
    }
  }
  try{
    const r=await fetch('/api/job/'+jobId+'/confirm',
      {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({school_id:parseInt(sid),overrides:ovs})});
    const d=await r.json();
    if(d.error){toast(d.error,'error');return}
    setStatus('building','Buildingâ€¦');startPoll();
  }catch(e){toast('Request failed: '+e.message,'error')}finally{
    btn.disabled=false;btn.innerHTML='âœ… Generate Excel';
  }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  HISTORY
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function loadHistory(){
  const r=await fetch('/api/history'), d=await r.json();
  const c=document.getElementById('hist-body');
  if(!d.files.length){c.innerHTML='<div style="text-align:center;padding:3rem;color:var(--g5)">ðŸ“ No files yet. Convert a PDF to see outputs here.</div>';return}
  c.innerHTML=`<table><thead><tr><th>File</th><th>Size</th><th>Created</th><th></th></tr></thead><tbody>
    ${d.files.map(f=>`<tr><td class="mono" style="font-size:12px">ðŸ“Š ${f.name}</td><td>${f.size_kb} KB</td><td>${f.created}</td>
      <td><a href="/api/history/${encodeURIComponent(f.name)}/download" class="btn btn-sm btn-primary" download>â¬‡ Download</a></td></tr>`).join('')}
  </tbody></table>`;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  MATRICULE CROSS-CHECK
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
let mcResults = null, mcPoll = null, mcRunning = false;
let mcYearOptions = [], mcSelYears = new Set(), mcYearLabel = 'recent 3 years';

async function loadMcYears(){
  try{
    const r = await fetch('/api/academic-years');
    const d = await r.json();
    mcYearOptions = d.years || [];
    mcSelYears = new Set(mcYearOptions.slice(0, 3));
    renderMcYears();
  }catch(e){}
}

function renderMcYears(){
  const box = document.getElementById('mc-years');
  if(!box) return;
  box.innerHTML = '';
  for(const y of mcYearOptions){
    const on = mcSelYears.has(y);
    const lb = document.createElement('label');
    lb.className = 'chk'; lb.style.padding = '2px 6px';
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.checked = on;
    cb.style.accentColor = 'var(--blue)';
    cb.addEventListener('change', () => {
      if(cb.checked) mcSelYears.add(y); else mcSelYears.delete(y);
      updateMcEstimate();
    });
    lb.appendChild(cb);
    lb.appendChild(document.createTextNode(' ' + y));
    box.appendChild(lb);
  }
  updateMcEstimate();
}

function setMcYearsPreset(preset){
  mcSelYears = preset === 'all' ? new Set(mcYearOptions) : new Set(mcYearOptions.slice(0, 3));
  renderMcYears();
}

function updateMcEstimate(){
  const est = document.getElementById('mc-est');
  if(!est) return;
  const n = students.length, yrs = mcSelYears.size;
  if(!n || !yrs){ est.textContent = ''; return; }
  const reqs = n * 6 * yrs;                       // phase-1 queries â‰ˆ 6 per student
  const mins = Math.max(1, Math.ceil(reqs / 66)); // â‰ˆ 8 concurrent @ 0.12s apart
  est.textContent = `â‰ˆ ${reqs.toLocaleString()} requests (${n} students Ã— ${yrs} yr) Â· ~${mins} min`;
  est.title = 'Rough phase-1 estimate; resolving students adds more queries.';
}

function setMcProg(checked, total){
  const bar = document.getElementById('mc-prog-bar');
  if(!bar) return;
  const pct = total ? Math.round(checked/total*100) : 0;
  bar.style.width = pct + '%';
  const lbl = document.getElementById('mc-prog-lbl');
  if(lbl) lbl.textContent = `${checked} / ${total}`;
}

async function runMatriculeCheck(){
  if(!jobId || mcRunning) return;
  const btn = document.getElementById('btn-mat-check');
  const status = document.getElementById('mc-status');
  const years = [...mcSelYears];
  if(!years.length){ toast('Select at least one academic year','warning'); return; }
  mcYearLabel = years.join(' Â· ');
  mcRunning = true;
  btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Checkingâ€¦';
  status.textContent = `Starting portal check across ${years.length} year(s)â€¦`;
  document.getElementById('mc-body').style.display = 'block';
  setMcProg(0, 0);
  document.getElementById('mc-results').innerHTML = '<div style="text-align:center;padding:1.5rem;color:var(--g4)">Querying the Landmark portalâ€¦</div>';

  try{
    const r = await fetch(`/api/job/${jobId}/matricule-check`,
      {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({years})});
    const d = await r.json();
    if(d.error){ toast(d.error,'error'); mcRunning=false; resetMcBtn(); return; }
    mcPoll = setInterval(pollMatCheck, 1500);
    pollMatCheck();
  }catch(e){
    toast('Failed to start check: '+e.message,'error');
    mcRunning = false; resetMcBtn();
  }
}

async function pollMatCheck(){
  try{
    const r = await fetch(`/api/job/${jobId}/matricule-check`);
    const d = await r.json();
    setMcProg(d.checked || 0, d.total || 0);
    if(d.status === 'running'){
      document.getElementById('mc-status').textContent =
        d.total ? `Checked ${d.checked} of ${d.total} studentsâ€¦` : 'Preparing queriesâ€¦';
      return;
    }
    clearInterval(mcPoll); mcPoll = null; mcRunning = false;
    resetMcBtn();
    if(d.status === 'failed'){
      document.getElementById('mc-results').innerHTML =
        `<div class="alert a-red"><div class="alert-icon">âœ—</div><div><div class="ab-title">Matricule check failed</div><div style="font-size:12px">${esc(d.error||'Unknown error')}</div></div></div>`;
      document.getElementById('mc-status').textContent = 'Failed';
      toast('Matricule check failed','error');
      return;
    }
    mcResults = d.results || [];
    renderMatriculeResults(mcResults);
    const cnt = s => mcResults.filter(x => x.status === s).length;
    const v = cnt('verified'), m = cnt('mismatch'), rv = cnt('review'), nf = cnt('not_found');
    document.getElementById('mc-status').textContent =
      `${v} verified Â· ${m} mismatch Â· ${rv} review Â· ${nf} not found`;
    const msgs = [];
    if(m) msgs.push(`${m} matricule mismatch(es)`);
    if(rv) msgs.push(`${rv} need(s) review`);
    if(msgs.length) toast(msgs.join(' â€” '),'warning');
    else if(v) toast(`All ${v} students verified âœ“`,'success');
  }catch(e){}
}

function resetMcBtn(){
  const btn = document.getElementById('btn-mat-check');
  btn.disabled = false; btn.innerHTML = 'â†» Re-check Matricules';
}

function renderMatriculeResults(results){
  const c = document.getElementById('mc-results');
  if(!results.length){
    c.innerHTML = '<div class="alert a-blue"><div class="alert-icon">â„¹</div><div>No students to check.</div></div>';
    return;
  }
  const by = s => results.filter(x => x.status === s);
  const verified = by('verified'), mismatched = by('mismatch'),
        review = by('review'), notfound = by('not_found'), skipped = by('skipped');
  const simPct = s => Math.round(s*100);

  let html = `<div class="flex gap-2 aic" style="margin-bottom:1rem;flex-wrap:wrap">
    <span class="badge bg-green">âœ“ ${verified.length} verified</span>
    ${mismatched.length?`<span class="badge bg-red">âœ— ${mismatched.length} mismatch</span>`:''}
    ${review.length?`<span class="badge bg-amber">? ${review.length} review</span>`:''}
    ${notfound.length?`<span class="badge bg-grey">? ${notfound.length} not found</span>`:''}
    ${skipped.length?`<span class="badge bg-grey">${skipped.length} skipped</span>`:''}
  </div>`;

  if(mismatched.length){
    html += `<div class="alert a-red mb3"><div class="alert-icon">âš </div>
      <div><div class="ab-title">${mismatched.length} student(s) have a different matricule in the portal</div>
      <div style="font-size:12px">Override if the portal matricule is correct.</div></div></div>`;
    for(const r of mismatched) html += renderMatchCard(r, 'red', simPct(r.confidence));
  }

  if(review.length){
    html += `<div class="alert a-amber mb3"><div class="alert-icon">â„¹</div>
      <div><div class="ab-title">${review.length} student(s) need manual review</div>
      <div style="font-size:12px">Low-confidence match â€” verify before overriding or generating Excel.</div></div></div>`;
    for(const r of review) html += renderMatchCard(r, 'amber', simPct(r.confidence));
  }

  if(verified.length){
    html += `<details style="margin-top:.5rem">
      <summary style="cursor:pointer;font-size:12px;color:var(--g5)">${verified.length} student(s) verified âœ“</summary>
      <div style="margin-top:.5rem;display:flex;flex-direction:column;gap:4px">`;
    for(const r of verified){
      const yStr = r.years && r.years.length ? r.years.map(y=>esc(y)).join(', ') : '';
      html += `<div style="font-size:12px;padding:.25rem .5rem;background:var(--g0);border-radius:4px;display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <span class="mono" style="font-weight:500">${esc(r.matricule)}</span>
        <span style="color:var(--g5)"> â€” ${esc(r.name)}</span>
        ${yStr?`<span class="text-sm muted" style="margin-left:6px">(${yStr})</span>`:''}
        <span class="badge bg-green" style="font-size:10px">${simPct(r.confidence)}%</span>
        ${r.matricule_matched?'':'<span class="badge bg-amber" style="font-size:10px">name-matched</span>'}
      </div>`;
    }
    html += `</div></details>`;
  }

  if(notfound.length){
    html += `<details style="margin-top:.5rem">
      <summary style="cursor:pointer;font-size:12px;color:var(--g5)">${notfound.length} student(s) not found in portal</summary>
      <div style="margin-top:.5rem;display:flex;flex-direction:column;gap:4px">`;
    for(const r of notfound){
      html += `<div style="font-size:12px;padding:.25rem .5rem;background:var(--g0);border-radius:4px">
        <span class="mono">${esc(r.matricule)}</span>
        <span style="color:var(--g5)"> â€” ${esc(r.name)}</span>
        <span style="color:var(--g4)"> (searched ${mcYearLabel})</span>
      </div>`;
    }
    html += `</div></details>`;
  }

  if(skipped.length){
    html += `<details style="margin-top:.5rem">
      <summary style="cursor:pointer;font-size:12px;color:var(--g5)">${skipped.length} student(s) skipped (no name)</summary>
      <div style="margin-top:.5rem;display:flex;flex-direction:column;gap:4px">`;
    for(const r of skipped){
      html += `<div style="font-size:12px;padding:.25rem .5rem;background:var(--g0);border-radius:4px">
        <span class="mono">${esc(r.matricule)}</span>
        <span style="color:var(--g4)"> â€” no name in transcript</span>
      </div>`;
    }
    html += `</div></details>`;
  }

  c.innerHTML = html;
}

function renderMatchCard(r, color, confidence){
  const borderColor = color==='red'?'var(--red)':color==='amber'?'var(--amber)':'var(--g3)';
  const btnClass = color==='red'?'btn-danger':'btn-primary';
  const yearTags = (years)=>years&&years.length?years.map(y=>`<span class="badge bg-grey" style="font-size:10px">${esc(y)}</span>`).join(' '):'';
  let candidateHtml = '';
  if(r.candidates && r.candidates.length){
    candidateHtml = `<details style="margin-top:.5rem">
      <summary style="cursor:pointer;font-size:11px;color:var(--g5)">${r.candidates.length} candidate(s) found</summary>
      <div style="margin-top:.4rem;display:flex;flex-direction:column;gap:3px">`;
    for(const cand of r.candidates){
      const isBest = cand.matricule === r.api_matricule;
      const matSim = cand.matricule_similarity >= 1 ? 'mat = PDF'
                   : cand.matricule_similarity >= 0.9 ? 'mat â‰ˆ PDF (OCR)' : 'mat â‰  PDF';
      candidateHtml += `<div style="font-size:11px;padding:.2rem .4rem;background:${isBest?'var(--blue-pale)':'var(--g0)'};border-radius:3px;display:flex;justify-content:space-between;align-items:center;gap:6px">
        <div><span class="mono" style="font-weight:${isBest?600:400}">${esc(cand.matricule)}</span>
        <span style="color:var(--g5)"> â€” ${esc(cand.name||'?')}</span>
        ${cand.years && cand.years.length ? `<span style="margin-left:6px">${yearTags(cand.years)}</span>` : ''}</div>
        <div class="flex gap-2 aic">
          <span style="color:var(--g5)">${Math.round(cand.confidence*100)}% ${isBest?'â† best':''} Â· ${matSim}</span>
          ${isBest?'':`<button class="btn btn-sm btn-outline" style="padding:.2rem .5rem;font-size:11px"
            onclick="event.stopPropagation();overrideMatricule('${esc(r.matricule)}','${esc(cand.matricule)}')">ðŸ”„ Select</button>`}
        </div>
      </div>`;
    }
    candidateHtml += `</div></details>`;
  }
  const cardId = `mc-card-${esc(r.matricule||'').replace(/[^a-z0-9]/gi,'_')}`;
  return `<div id="${cardId}" class="card" style="margin-bottom:.75rem;border-left:3px solid ${borderColor}">
    <div class="card-body" style="padding:.75rem 1rem">
      <div class="flex aic jbs" style="margin-bottom:.4rem">
        <div><strong>${esc(r.name)}</strong> <span class="badge ${color==='red'?'bg-red':color==='amber'?'bg-amber':'bg-grey'}" style="font-size:10px">${confidence}%</span></div>
        <div class="flex gap-2 aic">
          <button class="btn btn-sm btn-outline" onclick="document.getElementById('${cardId}').remove()">âœ• Dismiss</button>
          ${r.api_matricule?`<button class="btn btn-sm ${btnClass}" onclick="overrideMatricule('${esc(r.matricule)}','${esc(r.api_matricule)}')">
            ðŸ”„ Override â†’ ${esc(r.api_matricule)}</button>`:''}
        </div>
      </div>
      <div style="font-size:12px;display:flex;gap:1.5rem;align-items:center;flex-wrap:wrap">
        <span>PDF matricule: <span class="mono" style="color:var(--red);font-weight:600">${esc(r.matricule)}</span></span>
        <span>â†’</span>
        <span>Portal matricule: <span class="mono" style="color:var(--green);font-weight:600">${esc(r.api_matricule)}</span></span>
        ${r.api_name?`<span class="text-sm muted">Portal name: ${esc(r.api_name)}</span>`:''}
        ${r.matricule_matched?'<span class="badge bg-green" style="font-size:10px">matricule matches</span>':'<span class="badge bg-red" style="font-size:10px">matricule differs</span>'}
        ${r.reason?`<span class="text-sm muted">${esc(r.reason)}</span>`:''}
      </div>
      ${candidateHtml}
    </div>
  </div>`;
}

async function overrideMatricule(oldMat, newMat){
  if(!jobId) return;
  if(!confirm(`Override matricule "${oldMat}" â†’ "${newMat}"?\n\nThis updates the student's matricule for the Excel export.`)) return;

  try{
    const r = await fetch(`/api/job/${jobId}/matricule-override`,{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({old_matricule:oldMat,new_matricule:newMat})
    });
    const d = await r.json();
    if(d.error){ toast(d.error,'error'); return; }
    toast(`Matricule updated: ${oldMat} â†’ ${newMat}`,'success');
    await loadStudents();
    if(mcResults){
      for(const x of mcResults){
        if(x.matricule === oldMat){
          x.matricule = newMat; x.status = 'verified';
          x.mismatch = false; x.api_matricule = newMat;
          break;
        }
      }
      renderMatriculeResults(mcResults);
    }
  }catch(e){ toast('Override failed: '+e.message,'error'); }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  FORM B
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
let fbBusy = false;

async function loadFormB(){
  if(!jobId) return;
  try{
    const r = await fetch(`/api/job/${jobId}/form-b`);
    const d = await r.json();
    if(d.error) return;
    document.getElementById('fb-dept').textContent = d.department || 'â€”';
    document.getElementById('fb-level').textContent = d.level || 'â€”';
    const sel = document.getElementById('fb-catalog');
    const btn = document.getElementById('btn-fb-gen');
    const catalogs = d.catalogs || [];
    sel.innerHTML = `<option value="">None â€” from uploaded transcripts only</option>` +
      catalogs.map((c,i)=>{
        const label = `${c.department||'?'} Â· Level ${c.level||'?'}${i===0&&(d.department||d.level)?' (auto)':''}`;
        return `<option value="${esc(c.name)}">${esc(label)} â€” ${esc(c.name)}</option>`;
      }).join('');
    sel.disabled = false;
    btn.disabled = false;
    document.getElementById('fb-body').style.display = 'block';
    document.getElementById('fb-status').textContent =
      catalogs.length ? `from uploads Â· ${catalogs.length} official catalogue(s) available`
                      : 'from uploads Â· no official catalogue';
  }catch(e){}
}

async function generateFormB(){
  if(!jobId || fbBusy) return;
  const btn = document.getElementById('btn-fb-gen');
  const catalog = document.getElementById('fb-catalog').value;
  fbBusy = true;
  btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Cross-checkingâ€¦';
  document.getElementById('fb-status').textContent = 'Workingâ€¦';
  document.getElementById('fb-prog').style.display = 'block';
  document.getElementById('fb-prog-bar').style.width = '0%';
  document.getElementById('fb-prog-lbl').textContent = 'Matching coursesâ€¦';
  try{
    const r = await fetch(`/api/job/${jobId}/form-b/generate`,{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({catalog})
    });
    const d = await r.json();
    if(d.error){ toast(d.error,'error'); renderFormBError(d.error); return; }
    renderFormBResult(d);
    toast(`FORM B generated: ${d.catalog}`,'success');
  }catch(e){
    toast('FORM B failed: '+e.message,'error');
    renderFormBError(e.message);
  }finally{
    fbBusy = false;
    btn.disabled = false; btn.innerHTML = 'â†» Re-run Cross-Check';
  }
}

function renderFormBError(msg){
  document.getElementById('fb-prog').style.display = 'none';
  document.getElementById('fb-status').textContent = 'Failed';
  document.getElementById('fb-results').innerHTML =
    `<div class="alert a-red"><div class="alert-icon">âœ—</div><div><div class="ab-title">FORM B failed</div>
    <div style="font-size:12px">${esc(msg)}</div></div></div>`;
}

function renderFormBResult(d){
  document.getElementById('fb-prog').style.display = 'none';
  document.getElementById('fb-status').textContent =
    `Done â€” ${d.stats.catalog_count} courses from ${d.stats.student_count} students (uploads)`;
  const st = d.stats || {};
  const per = st.per_student || [];
  const totalMissing = per.reduce((a,p)=>a+p.missing.length,0);
  const totalUnexpected = per.reduce((a,p)=>a+p.unexpected.length,0);
  const badStudents = per.filter(p=>p.missing.length||p.unexpected.length||p.issues.length).length;

  let html = `<div class="flex gap-2 aic" style="margin-bottom:.75rem;flex-wrap:wrap">
    <span class="badge bg-blue">${st.catalog_count||0} courses</span>
    <span class="badge bg-blue">${st.student_count||0} students</span>
    <span class="badge bg-green">source: uploaded transcripts</span>
    <a class="btn btn-primary" style="text-decoration:none" href="/api/form-b/${encodeURIComponent(d.output_file)}/download">â¬‡ Download FORM B</a>
  </div>`;

  const cc = d.catalog_compare;
  if(cc){
    if(cc.error){
      html += `<div class="alert a-amber" style="margin-top:.5rem"><div class="alert-icon">â„¹</div>
        <div><div class="ab-title">Catalogue comparison skipped</div>
        <div style="font-size:12px">${esc(cc.name)} â€” ${esc(cc.error)}</div></div></div>`;
    }else{
      const ccPer = cc.per_student || [];
      const ccMissing = ccPer.reduce((a,p)=>a+p.missing.length,0);
      const ccUnexpected = ccPer.reduce((a,p)=>a+p.unexpected.length,0);
      html += `<details style="margin-top:.5rem" ${(ccMissing||ccUnexpected)?'open':''}>
        <summary style="cursor:pointer;font-size:12px;color:var(--g5)">Official catalogue comparison â€” ${esc(cc.name)} (${cc.catalog_count||0} courses)</summary>
        <div style="margin-top:.5rem;display:flex;flex-direction:column;gap:4px">
          <div class="flex gap-2 aic" style="flex-wrap:wrap">
            <span class="badge ${ccMissing?'bg-amber':'bg-green'}">${ccMissing} missing vs catalogue</span>
            <span class="badge ${ccUnexpected?'bg-amber':'bg-green'}">${ccUnexpected} not in catalogue</span>
          </div>
          <div style="max-height:260px;overflow:auto">
            <table style="font-size:12px">
              <thead><tr><th>Matricule</th><th>Name</th><th>Missing</th><th>Unexpected</th></tr></thead><tbody>`;
      for(const p of ccPer){
        const bad = p.missing.length||p.unexpected.length;
        html += `<tr${bad?' style="background:var(--red-pale)"':''}>
          <td class="mono">${esc(p.matricule)}</td><td>${esc(p.name)}</td>
          <td style="text-align:center">${p.missing.length}</td><td style="text-align:center">${p.unexpected.length}</td></tr>`;
      }
      html += `</tbody></table></div></div></details>`;
    }
  }

  html += `<details style="margin-top:.5rem" open>
    <summary style="cursor:pointer;font-size:12px;color:var(--g5)">Course coverage (per course Ã— students)</summary>
    <div style="margin-top:.5rem;max-height:260px;overflow:auto">
      <table style="font-size:12px">
        <thead><tr><th>Course</th><th>Description</th><th>Sem</th><th>Credit</th><th>Taken by</th></tr></thead><tbody>`;
  const courses = st.courses || [];
  for(const c of courses){
    html += `<tr><td class="mono">${esc(c.code)}</td><td>${esc(c.description)}</td>
      <td style="text-align:center">${c.semester}</td><td style="text-align:center">${c.credit}</td>
      <td style="text-align:center">${c.count} / ${st.student_count||0}</td></tr>`;
  }
  html += `</tbody></table></div></details>`;

  document.getElementById('fb-results').innerHTML = html;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  UI HELPERS
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const SC={parsing:'var(--blue)',ready:'var(--teal)',building:'var(--amber)',done:'var(--green)',fail:'var(--red)'};
function setStatus(s,l){
  document.getElementById('tb-right').innerHTML=`<span class="pill" style="background:${SC[s]||'#888'}22;color:${SC[s]||'#888'}">${l}</span>`;
}
function setFoot(col,txt){
  document.getElementById('foot-dot').style.color=col;
  document.getElementById('foot-txt').textContent=txt;
}
const CL={upload:'PDF uploaded',parse:'Pages parsed',validate:'Records validated',review:'Review confirmed',excel:'Excel generated'};
const CS={};
function setChk(k,v){
  CS[k]=v;
  document.getElementById('chk-'+k).innerHTML=(v?'âœ…':'â¬œ')+' '+CL[k];
  document.getElementById('chk-'+k).className='chk'+(v?' done':'');
}
function toast(msg,type){
  const t=document.createElement('div');
  t.className=`toast ${type||''}`;t.innerHTML=`<span>${msg}</span>`;
  document.getElementById('toasts').appendChild(t);
  setTimeout(()=>t.remove(),4200);
}
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;')}

setStatus('idle','Idle');
loadMcYears();
