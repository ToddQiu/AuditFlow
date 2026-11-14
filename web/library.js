'use strict';
let libraryData=null, libraryLimit=50, libraryLoading=false, librarySignature='';
function libraryOptions(id,entries){
  const select=document.querySelector(id),value=select.value;
  const signature=JSON.stringify(entries);
  if(select.dataset.options===signature)return;
  const first=select.options[0].text;
  select.innerHTML='<option value="">'+esc(first)+'</option>'+Object.entries(entries).map(([key,label])=>`<option value="${esc(key)}">${esc(label)}</option>`).join('');
  select.value=value;select.dataset.options=signature;
}
function libraryTime(value){return value?new Date(value).toLocaleString('zh-CN',{hour12:false}):'—';}
function librarySize(value){return value===undefined?'—':value<1024?value+' B':value<1048576?(value/1024).toFixed(1)+' KB':(value/1048576).toFixed(1)+' MB';}
async function loadLibrary(){
  if(libraryLoading)return;
  libraryLoading=true;
  try{
    libraryData=await api('/api/library');
    $('#library-root').textContent=libraryData.root;
    $('#library-scan-state').textContent=libraryData.scanning?'正在扫描与分类…':libraryData.last_scan?'上次扫描 '+libraryTime(libraryData.last_scan):'等待扫描';
    $('#library-scan').disabled=libraryData.scanning;
    $('#library-errors').textContent=libraryData.errors.join('；');
    libraryOptions('#library-category',libraryData.categories);
    libraryOptions('#library-status',libraryData.statuses);
    libraryOptions('#library-package',Object.fromEntries([...new Set(libraryData.files.map(f=>f.package||'根目录'))].sort().map(x=>[x,x])));
    const signature=JSON.stringify(libraryData.files);
    if(signature!==librarySignature&&!$('#library-table').contains(document.activeElement)){renderLibrary();librarySignature=signature;}
  }finally{libraryLoading=false;}
}
function renderLibrary(){
  if(!libraryData)return;
  const all=libraryData.files,q=$('#library-search').value.trim().toLowerCase(),category=$('#library-category').value,status=$('#library-status').value,pack=$('#library-package').value,attention=$('#library-attention').value;
  const rows=all.filter(f=>(!q||[f.relative,f.reason,libraryData.categories[f.category]].join(' ').toLowerCase().includes(q))&&(!category||f.category===category)&&(!status||f.status===status)&&(!pack||(f.package||'根目录')===pack)&&(!attention||(attention==='changed'?Boolean(f.change):attention==='unconfirmed'?!f.category_confirmed:f.availability!=='present')));
  $('#library-count').textContent=`共 ${all.length} 份资料 · 新增／更新 ${all.filter(f=>f.change&&f.availability==='present').length} · 待处理 ${all.filter(f=>f.status==='pending'&&f.availability==='present').length} · 当前显示 ${Math.min(rows.length,libraryLimit)} / ${rows.length}`;
  $('#library-more').hidden=rows.length<=libraryLimit;
  if(!rows.length){$('#library-table').innerHTML=`<div class="library-empty"><h2>${all.length?'没有符合筛选的资料':'资料库准备好了'}</h2><p>${all.length?'调整搜索或筛选条件查看其他文件。':'点击「打开资料文件夹」，把整包客户资料复制进去。扫描完成后，文件会出现在这里。'}</p></div>`;return;}
  const availability={missing:'已移除',unreadable:'读取受阻'};
  $('#library-table').innerHTML=`<table class="simple-table library-files"><thead><tr><th>文件与版本</th><th>分类与识别依据</th><th>处理状态</th><th>更新信息</th><th>操作</th></tr></thead><tbody>${rows.slice(0,libraryLimit).map(f=>{
    const index=all.indexOf(f),unavailable=f.availability!=='present',disabled=unavailable?'disabled':'';
    return `<tr data-library-index="${index}"><td><strong class="library-file-name">${esc(f.name)}</strong><p class="path small">${esc(f.relative)}</p><div class="library-tags">${unavailable?`<span class="status difference">${esc(availability[f.availability])}</span>`:f.change?`<span class="status needs_review">${f.change==='new'?'新增资料':'内容已更新'}</span>`:''}<span class="small subtle">${esc(f.ext.toUpperCase().slice(1)||'无扩展名')} · ${esc(librarySize(f.size))}</span></div><details><summary class="small">查看识别摘要</summary><p class="small">${esc(f.read_warning||'已读取部分表名／文字，用于资料分类。')}</p><p class="small subtle library-snippet">${esc(f.snippet||'没有文字摘要，可打开原件查看。')}</p></details></td><td><label class="sr-only" for="lib-cat-${index}">${esc(f.name)}的分类</label><select id="lib-cat-${index}" data-library-category ${disabled}>${Object.entries(libraryData.categories).map(([key,title])=>`<option value="${esc(key)}" ${key===f.category?'selected':''}>${esc(title)}</option>`).join('')}</select><p class="small subtle">${esc(f.reason||'待确认用途')}</p>${f.category_confirmed?'<span class="small">分类已确认</span>':`<button class="library-confirm" data-library-confirm ${disabled}>确认此分类</button>`}</td><td><label class="sr-only" for="lib-status-${index}">${esc(f.name)}的处理状态</label><select id="lib-status-${index}" data-library-status ${disabled}>${Object.entries(libraryData.statuses).map(([key,title])=>`<option value="${esc(key)}" ${key===f.status?'selected':''}>${esc(title)}</option>`).join('')}</select>${f.change==='updated'?'<p class="small error-text">内容变更，已重置待处理</p>':''}</td><td class="small library-modified">${esc(libraryTime(f.modified))}<p class="subtle">${esc(f.sha256?.slice(0,10)||'尚未读取')}…</p></td><td><div class="library-actions"><button data-library-open ${disabled}>打开原件</button><button data-library-reveal ${disabled}>所在文件夹</button>${f.can_use_tb?'<button data-library-intake>用于 TB／明细</button>':''}</div></td></tr>`;
  }).join('')}</tbody></table>`;
  $$('[data-library-index]').forEach(row=>{
    const file=all[Number(row.dataset.libraryIndex)];
    const update=async changes=>{libraryData=await api('/api/library/update',{path:file.path,sha256:file.sha256,...changes});renderLibrary();librarySignature=JSON.stringify(libraryData.files);};
    row.querySelector('[data-library-category]').onchange=e=>act(()=>update({category:e.target.value}));
    row.querySelector('[data-library-status]').onchange=e=>act(()=>update({status:e.target.value}));
    const confirm=row.querySelector('[data-library-confirm]');if(confirm)confirm.onclick=()=>act(()=>update({category:file.category}));
    row.querySelector('[data-library-open]').onclick=()=>act(()=>api('/api/library/open',{path:file.path}));
    row.querySelector('[data-library-reveal]').onclick=()=>act(()=>api('/api/library/open',{path:file.path,reveal:true}));
    const intake=row.querySelector('[data-library-intake]');if(intake)intake.onclick=()=>act(async()=>{showView('intake');await loadIntakeSource(file.path);notice('已选用资料库中的文件，请确认实体、期间及实际取数范围。');});
  });
}
function initializeLibrary(){
  for(const id of ['library-search','library-package','library-category','library-status','library-attention'])$('#'+id).addEventListener('input',()=>{libraryLimit=50;renderLibrary();});
  $('#library-more').onclick=()=>{libraryLimit+=50;renderLibrary();};
  $('#library-open-folder').onclick=()=>act(()=>api('/api/library/open-folder',{}));
  $('#library-scan').onclick=()=>act(async()=>{await api('/api/library/scan',{});await loadLibrary();});
  setInterval(()=>{if(!document.hidden)act(loadLibrary);},10000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)act(loadLibrary);});
}
