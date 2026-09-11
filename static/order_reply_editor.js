(() => {
  const frame = document.getElementById('replyEditor');
  const bar = document.getElementById('replyTools');
  if (!frame || !bar) return;
  let boundDocument = null;
  let cell = null, history = [], index = -1;
  const doc = () => frame.contentDocument;
  function rememberSelection() {
    const current=doc();
    if(!current) return;
    const selection=current.getSelection();
    const node=selection?.anchorNode;
    const element=node && (node.nodeType===1?node:node.parentElement);
    const selected=element?.closest('td,th');
    if(selected && selected.isConnected) cell=selected;
  }
  const snapshot = () => {
    const value = doc().body.innerHTML;
    if (history[index] === value) return;
    history = history.slice(0, index + 1); history.push(value);
    if (history.length > 60) history.shift();
    index = history.length - 1;
  };
  function grid(table) {
    const rows = [...table.rows], map = [], positions = new Map();
    rows.forEach((row, r) => {
      map[r] ||= []; let c = 0;
      [...row.cells].forEach(td => {
        while (map[r][c]) c++;
        const h = td.rowSpan || rows.length - r, w = td.colSpan;
        positions.set(td, {r, c, h, w});
        for(let y=r;y<r+h;y++) { map[y] ||= []; for(let x=c;x<c+w;x++) map[y][x]=td; }
        c += w;
      });
    });
    return {rows, map, positions};
  }
  function empty(tag) {
    const td = doc().createElement(tag); td.innerHTML='<br>';
    td.style.cssText='border:1px solid #999;padding:6px;min-width:80px'; return td;
  }
  function run(action) {
    initialize();
    rememberSelection();
    if(action==='copy-order') {
      if(!cell || !cell.isConnected) { alert('请先点击下方客户订单表格中的单元格'); return; }
      const table=cell.closest('table');
      const reply=doc().getElementById('nouya-current-reply');
      if(reply.contains(table)) { alert('这已经是顶部回复表格，无需再次复制'); return; }
      snapshot(); copyOrderTable(table,reply); snapshot(); return;
    }
    if(action==='full') { frame.closest('.or-compose').classList.toggle('reply-full'); return; }
    if(action==='preview') { const off=doc().body.contentEditable==='true'; doc().body.contentEditable=String(!off); bar.querySelector('[data-action="preview"]').textContent=off?'返回编辑':'预览'; return; }
    if(action==='undo'||action==='redo') {
      const next=index+(action==='undo'?-1:1);
      if(next>=0&&next<history.length) { index=next; doc().body.innerHTML=history[index]; cell=null; } return;
    }
    if(doc().body.contentEditable!=='true') { alert('请先返回编辑'); return; }
    if(!cell || !cell.isConnected) { alert('请先点击要操作的表格单元格'); return; }
    const table=cell.closest('table'), {rows,map,positions}=grid(table), p=positions.get(cell);
    snapshot();
    if(action==='width') {
      const width=Number(prompt('此列宽度（像素，80–800）','140'));
      if(width>=80&&width<=800) new Set(map.map(row=>row[p.c])).forEach(td=>{if(td) td.style.width=width+'px';});
    } else if(action==='fill') {
      const value=prompt('填写此列从当前行到末行的内容（空值不修改）');
      if(value) new Set(map.slice(p.r).map(row=>row[p.c])).forEach(td=>{if(td) td.textContent=value;});
    } else if(action==='left'||action==='right') {
      const boundary=p.c+(action==='right'?p.w:0), expanded=new Set();
      rows.forEach((row,r)=>{
        const cross=map[r][boundary];
        if(cross && positions.get(cross).c<boundary) {
          if(!expanded.has(cross)) { cross.colSpan++; expanded.add(cross); }
        } else {
          const next=[...row.cells].find(td=>positions.get(td).c>=boundary);
          row.insertBefore(empty(cell.tagName.toLowerCase()),next||null);
        }
      });
    } else if(action==='delete-col') {
      if(!confirm('删除当前列？可以撤销。')) return;
      new Set(map.map(row=>row[p.c])).forEach(td=>{if(td) {if(td.colSpan>1) td.colSpan--; else td.remove();}});
    } else {
      // Row insertion across merged regions is ambiguous: preserve the customer's layout.
      if([...positions.values()].some(v=>v.h>1||v.w>1)) { alert('此表有合并单元格，暂不支持增删行，以免破坏客户格式；仍可插列和编辑内容。'); return; }
      if(action==='delete-row') { if(confirm('删除当前行？可以撤销。')) cell.parentElement.remove(); }
      else {
        const row=doc().createElement('tr');
        for(let c=0;c<map[p.r].length;c++) row.append(empty('td'));
        const current=cell.parentElement;
        current.parentElement.insertBefore(row,action==='above'?current:current.nextSibling);
      }
    }
    snapshot();
  }
  bar.addEventListener('click', e=> { const button=e.target.closest('[data-action]'); if(button) run(button.dataset.action); });
  bar.addEventListener('mousedown', e=> {
    rememberSelection();
    // Keep the iframe caret while clicking a toolbar button.
    if(e.target.closest('button')) e.preventDefault();
  });
  function copyOrderTable(table, reply) {
    const copy=table.cloneNode(true);
    // Original ids must remain unique; never modify the source table.
    copy.removeAttribute('id');
    copy.querySelectorAll('[id]').forEach(node=>node.removeAttribute('id'));
    copy.id='nouya-reply-order-table';
    reply.append(copy);
    reply.scrollIntoView({block:'start'});
  }
  function initialize() {
    const current = doc();
    if (!current || !current.body || current === boundDocument) return;
    boundDocument = current;
    cell = null; history = []; index = -1;
    current.body.contentEditable = 'true';
    current.body.style.minHeight = '580px';
    current.body.style.padding = '12px';
    current.body.setAttribute('aria-label', '邮件正文编辑器');
    // Keep this reply separate from the original signature and quoted order.
    // The stable id survives draft sanitization, so reopening does not add another box.
    let reply = current.getElementById('nouya-current-reply');
    if (!reply) {
      reply = current.createElement('div');
      reply.id = 'nouya-current-reply';
      reply.innerHTML = '<p><br></p>';
      current.body.prepend(reply);
    }
    reply.style.minHeight = '';
    reply.setAttribute('aria-label', '本次回复内容');
    if (!reply.querySelector('table')) {
      const candidates=[...current.querySelectorAll('table')].filter(table=>{
        if(reply.contains(table)) return false;
        // Match direct cells only, not an outer email layout wrapping another table.
        const text=[...table.rows].slice(0,8).flatMap(row=>[...row.cells])
          .filter(td=>!td.querySelector('table')).map(td=>td.textContent).join(' ');
        return /物料|产品|品名|料号|material|description/i.test(text)
          && /数量|quantity|qty/i.test(text) && /交期|交货|delivery|备注|单价/i.test(text);
      });
      if(candidates.length===1) copyOrderTable(candidates[0],reply);
    }
    const selectCell = node => {
      const element = node && (node.nodeType === 1 ? node : node.parentElement);
      const selected = element?.closest('td,th');
      if (selected) cell = selected;
    };
    current.addEventListener('pointerdown', e => selectCell(e.target), true);
    current.addEventListener('click', e => selectCell(e.target), true);
    current.addEventListener('selectionchange', () => selectCell(current.getSelection()?.anchorNode));
    current.addEventListener('input', snapshot);
    current.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='z'){e.preventDefault();run(e.shiftKey?'redo':'undo');}});
    snapshot();
  }
  frame.addEventListener('load', initialize);
  // srcdoc may finish before this external script loads, especially from cache.
  initialize();
  // Read from the parent context as well: sandboxed WebKit documents may not
  // dispatch their script callbacks, even though the parent can read selection.
  setInterval(rememberSelection, 150);
})();
