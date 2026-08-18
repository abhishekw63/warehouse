/* online_b2b/tables.html — page script (separated from template). */
(function(){
  const $=s=>document.querySelector(s);
  const api=(url,body)=>B2B.postJSON(url,body);
  const toast=(m,k)=>{ if(window.B2B&&B2B.toast) B2B.toast(m,k||'ok'); };

  let TABLE = JSON.parse($('#ct-table').textContent||'null');
  let ROWS  = JSON.parse($('#ct-rows').textContent||'[]');

  function colorClass(key,val){
    const rules=(TABLE&&TABLE.color_rules)||{}; const cr=rules[key]; if(!cr||val==null) return '';
    const c=cr[val]||cr[String(val).trim()]; return c?('row-'+c):'';
  }
  function rowTint(data){ const rules=(TABLE&&TABLE.color_rules)||{};
    for(const k in rules){ const c=colorClass(k,data[k]); if(c) return c; } return ''; }
  function esc(s){return (s==null?'':String(s));}

  function render(){
    const grid=$('#ct-grid');
    if(!TABLE){ grid.innerHTML=''; $('#ct-empty').hidden=false; return; }
    const cols=TABLE.columns||[];
    let head='<thead><tr><th class="ct-rowix">#</th>'+cols.map(c=>`<th>${esc(c.label)}</th>`).join('')+'<th></th></tr></thead>';
    let body='<tbody>';
    ROWS.forEach((r,i)=>{
      body+=`<tr data-id="${r.id}" class="${rowTint(r.data||{})}"><td class="ct-rowix">${i+1}</td>`;
      cols.forEach(c=>{ const v=esc((r.data||{})[c.key]);
        const cc=colorClass(c.key,(r.data||{})[c.key]); const pill=cc?(' pill '+cc.replace('row-','pill-')):'';
        body+=`<td><input class="ct-cell${pill}" data-key="${c.key}" value="${v.replace(/"/g,'&quot;')}" placeholder="—"></td>`; });
      body+=`<td style="text-align:center"><button class="ct-del" title="Delete row">×</button></td></tr>`;
    });
    grid.innerHTML=head+body+'</tbody>';
    $('#ct-empty').hidden = ROWS.length>0;
    $('#ct-meta').textContent = ROWS.length+' row'+(ROWS.length===1?'':'s');
  }

  function rowData(tr){ const d={}; tr.querySelectorAll('.ct-cell').forEach(inp=>d[inp.dataset.key]=inp.value.trim()); return d; }
  function repill(tr,data){ tr.querySelectorAll('.ct-cell').forEach(inp=>{
    const cc=colorClass(inp.dataset.key,data[inp.dataset.key]);
    inp.classList.remove('pill','pill-red','pill-green','pill-amber');
    if(cc) inp.classList.add('pill',cc.replace('row-','pill-')); }); }
  $('#ct-grid').addEventListener('change',e=>{
    const inp=e.target.closest('.ct-cell'); if(!inp) return;
    const tr=inp.closest('tr'); const id=tr.dataset.id; const data=rowData(tr);
    const flash=()=>{ tr.className=rowTint(data); repill(tr,data); void tr.offsetWidth; tr.classList.add('flash'); setTimeout(()=>tr.classList.remove('flash'),900); };
    if(id&&id!=='null'){ api(`/b2b/tables/row/${id}/update/`,{data}).then(()=>{ ROWS=ROWS.map(r=>r.id==id?{...r,data}:r); flash(); toast('Saved'); }); }
    else { api('/b2b/tables/row/add/',{table_id:TABLE.id,data}).then(r=>{ if(r.ok){ tr.dataset.id=r.id; ROWS.push({id:r.id,data}); flash(); bumpCount(1); toast('Row added'); } }); }
  });
  $('#ct-grid').addEventListener('click',e=>{
    const b=e.target.closest('.ct-del'); if(!b) return;
    const tr=b.closest('tr'); const id=tr.dataset.id; if(!confirm('Delete this row?')) return;
    const done=()=>{ ROWS=ROWS.filter(r=>String(r.id)!==String(id)); render(); bumpCount(-1); toast('Deleted'); };
    if(id&&id!=='null') api(`/b2b/tables/row/${id}/delete/`,{}).then(done); else done();
  });
  $('#ct-add-row').addEventListener('click',()=>{ if(!TABLE) return; ROWS.push({id:null,data:{}}); render();
    const last=$('#ct-grid tbody tr:last-child .ct-cell'); if(last) last.focus(); });
  $('#ct-search').addEventListener('input',e=>{ const q=e.target.value.toLowerCase();
    document.querySelectorAll('#ct-grid tbody tr').forEach(tr=>{ tr.style.display=tr.textContent.toLowerCase().includes(q)?'':'none'; }); });

  document.querySelectorAll('#ct-tabs .tab').forEach(tab=>tab.addEventListener('click',()=>{
    if(tab.classList.contains('on')) return;
    document.querySelectorAll('#ct-tabs .tab').forEach(t=>t.classList.remove('on')); tab.classList.add('on');
    const gs=document.querySelector('.ct-grid-scroll'); if(gs) gs.style.opacity='0';
    fetch(`/b2b/tables/${tab.dataset.slug}/data/`).then(r=>r.json()).then(d=>{
      if(!d.ok){ if(gs) gs.style.opacity='1'; return; }
      TABLE=d.table; ROWS=d.rows; $('#ct-title').textContent=TABLE.name; $('#ct-search').value=''; render();
      if(gs) requestAnimationFrame(()=>{ gs.style.opacity='1'; });
    });
  }));
  $('#ct-del-table').addEventListener('click',()=>{ if(!TABLE||!confirm(`Delete the whole "${TABLE.name}" table and all its rows?`)) return;
    api(`/b2b/tables/${TABLE.id}/delete/`,{}).then(()=>location.href='/b2b/tables/'); });
  $('#ct-export').addEventListener('click',()=>{ if(!TABLE) return; const cols=TABLE.columns;
    const csv=[cols.map(c=>c.label).join(',')].concat(ROWS.map(r=>cols.map(c=>`"${String((r.data||{})[c.key]||'').replace(/"/g,'""')}"`).join(','))).join('\n');
    const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'})); a.download=TABLE.name+'.csv'; a.click(); });
  function bumpCount(n){ const c=document.querySelector(`#ct-tabs .tab[data-id="${TABLE.id}"] .tcount`); if(c) c.textContent=(parseInt(c.textContent||'0',10)+n); }

  const modal=$('#ct-modal');
  function addColRow(){ const d=document.createElement('div'); d.className='ct-colrow';
    d.innerHTML=`<input class="cc-label" placeholder="Column label"><input class="cc-key" placeholder="key (optional)"><button class="ct-x" type="button">×</button>`;
    d.querySelector('.ct-x').onclick=()=>d.remove(); $('#ct-cols').appendChild(d); }
  $('#ct-new-table').onclick=()=>{ $('#ct-tname').value=''; $('#ct-cols').innerHTML=''; addColRow(); addColRow(); modal.hidden=false; $('#ct-tname').focus(); };
  $('#ct-modal-x').onclick=$('#ct-cancel').onclick=()=>modal.hidden=true;
  $('#ct-addcol').onclick=addColRow;
  $('#ct-create').onclick=()=>{
    const name=$('#ct-tname').value.trim();
    const columns=[...document.querySelectorAll('.ct-colrow')].map(r=>{
      const label=r.querySelector('.cc-label').value.trim();
      let key=r.querySelector('.cc-key').value.trim().toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_|_$/g,'');
      if(label&&!key) key=label.toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_|_$/g,'');
      return label?{key,label,type:'text'}:null; }).filter(Boolean);
    if(!name||!columns.length){ toast('Add a name and at least one column','warn'); return; }
    api('/b2b/tables/create/',{name,columns}).then(r=>{ if(r.ok) location.href='/b2b/tables/?t='+r.slug; else toast(r.error||'Failed','error'); });
  };

  render();
})();
