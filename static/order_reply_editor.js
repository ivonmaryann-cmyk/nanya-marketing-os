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
  const compact=value=>String(value||'').toLowerCase().replace(/[\s\-_.:：()（）]/g,'');
  const headerField=value=>{
    const text=compact(value);
    if(/^(po|pono|ponumber|po号|客户订单号|订单号|采购订单号)$/.test(text)) return 'customer_order_number';
    if(/^(项次|序号|客户项次|item|itemno|linenumber)$/.test(text)) return 'line_no';
    if(/^(客户料号|客户产品编号|客户产品料号|物料编码|物料代码|料号|customerpart|customermaterial|materialcode)$/.test(text)) return 'customer_product_code';
    if(/^(客户规格|规格|品名规格|customerspec|specification|description)$/.test(text)) return 'customer_spec';
    if(/^(数量|采购量|订购数量|订单数量|qty|quantity)$/.test(text)) return 'quantity';
    if(/^(单价|含税单价|未税单价|税前单价|unitprice|price)$/.test(text)) return 'unit_price';
    if(/要求交货期|要求交期|客户需求日|需求交期/.test(text)) return '';
    if(/供应商复期|供应商回复|交期回复|交货回复|交期|交货期|deliveryreply/.test(text)) return 'delivery_reply';
    return '';
  };
  const directCells=row=>[...row.cells].filter(cell=>!cell.querySelector('table'));
  const textAt=(cells,index)=>index===undefined?'':String(cells[index]?.textContent||'').trim();
  const poFromText=text=>{
    const found=[...String(text||'').matchAll(/(?:建价\s*)?\bPO[-\s]?[A-Z0-9][A-Z0-9-]{3,}\b/ig)]
      .map(item=>item[0].replace(/^建价\s*/i,'').replace(/\s+/g,''));
    return [...new Set(found)];
  };
  function tableInfo(table) {
    const tableRows=[...table.rows], tableGrid=grid(table);
    let headerIndex=-1, fields={};
    // Customer tables often reserve the first rows for form numbers or titles.
    // Use the row with the most recognizable fields as the header.
    tableRows.slice(0,8).forEach((row,index)=>{
      const candidate={};
      (tableGrid.map[index]||[]).forEach((cell,cellIndex)=>{
        const field=headerField(cell.textContent);
        if(field&&!candidate[field]) candidate[field]=cellIndex;
      });
      if(Object.keys(candidate).length>Object.keys(fields).length) {
        fields=candidate;headerIndex=index;
      }
    });
    const tablePos=poFromText(table.textContent);
    const fallbackPo=fields.customer_order_number===undefined&&tablePos.length===1?tablePos[0]:'';
    const hasPo=fields.customer_order_number!==undefined||fallbackPo;
    const hasMaterial=fields.customer_product_code!==undefined;
    const hasQuantity=fields.quantity!==undefined;
    const score=(hasPo?3:0)+(hasMaterial?2:0)+(hasQuantity?2:0)
      +(fields.customer_spec!==undefined?1:0)+(fields.delivery_reply!==undefined?1:0);
    return {
      table,tableRows,tableGrid,headerIndex,fields,fallbackPo,score,
      valid:headerIndex>=0&&hasPo&&hasMaterial&&hasQuantity,
    };
  }
  let replyRowsById=new Map();
  function replyTables() {
    initialize();
    const current=doc();
    if(!current) return {rows:[],tableCount:0,hasPossibleOrderTable:false};
    const infos=[...current.querySelectorAll('table')].map(tableInfo);
    const selected=infos.find(info=>info.valid);
    if(!selected) return {
      rows:[],tableCount:0,
      hasPossibleOrderTable:infos.some(info=>info.score>=3),
    };
    const rows=[];
    let currentPo=selected.fallbackPo;
    selected.tableRows.slice(selected.headerIndex+1).forEach((row,rowIndex)=>{
      const sourceIndex=selected.headerIndex+rowIndex+1,cells=selected.tableGrid.map[sourceIndex]||[];
      if(!cells.length) return;
      const values={
        customer_order_number:textAt(cells,selected.fields.customer_order_number)||currentPo,
        line_no:textAt(cells,selected.fields.line_no),
        customer_product_code:textAt(cells,selected.fields.customer_product_code),
        customer_spec:textAt(cells,selected.fields.customer_spec),
        quantity:textAt(cells,selected.fields.quantity), unit_price:textAt(cells,selected.fields.unit_price),
      };
      if(!values.customer_order_number||!values.customer_product_code||!values.quantity) return;
      currentPo=values.customer_order_number;
      const rowId=row.dataset.nouyaReplyRowId||`reply-0-${rowIndex}`;
      row.dataset.nouyaReplyRowId=rowId;
      rows.push({row_id:rowId,...values});
      replyRowsById.set(rowId,{
        row,table:selected.table,headerRow:selected.tableRows[selected.headerIndex],
        headerIndex:selected.headerIndex,fields:selected.fields,
      });
    });
    return {rows,tableCount:1,hasPossibleOrderTable:true};
  }
  function deliveryCell(info,index) {
    const rowIndex=[...info.table.rows].indexOf(info.row);
    return (grid(info.table).map[rowIndex]||[])[index];
  }
  function appendTableCell(row, tagName) {
    const reference=directCells(row).at(-1);
    const cell=reference?reference.cloneNode(false):doc().createElement(tagName);
    cell.removeAttribute('id');cell.removeAttribute('data-nouya-delivery-reply');
    cell.rowSpan=1;cell.colSpan=1;cell.textContent='';
    return cell;
  }
  function ensureReplyColumn(info,useExisting=false) {
    const marked=[...info.headerRow.cells].find(cell=>cell.dataset.nouyaDeliveryReply==='true');
    if(marked) {
      const index=grid(info.table).positions.get(marked)?.c;
      if(index!==undefined) { info.fields.delivery_reply=index; return index; }
    }
    let index=info.fields.delivery_reply;
    if(index!==undefined&&useExisting) return index;
    index=(grid(info.table).map[info.headerIndex]||[]).length;
    const head=appendTableCell(info.headerRow,'th');head.textContent='交期回复';head.dataset.nouyaDeliveryReply='true';info.headerRow.append(head);
    info.fields.delivery_reply=index;
    [...info.table.rows].slice(info.headerIndex+1).forEach(row=>{
      if(directCells(row).length)row.append(appendTableCell(row,'td'));
    });
    return index;
  }
  function replyCell(info) {
    return deliveryCell(info,info.fields.delivery_reply);
  }
  function createReplyTable(rows) {
    const current=doc(), reply=current?.getElementById('nouya-current-reply');
    if(!reply) return;
    reply.querySelector('table[data-nouya-generated-order-table="true"]')?.remove();
    const table=current.createElement('table');table.style.cssText='border-collapse:collapse;border:1px solid #9caec4;font-size:14px';
    table.dataset.nouyaGeneratedOrderTable='true';
    const headers=[['customer_order_number','PO号'],['line_no','项次'],['customer_product_code','客户料号'],['customer_spec','客户规格'],['quantity','数量'],['delivery_reply','交期回复']];
    const head=current.createElement('thead'),headRow=current.createElement('tr');headers.forEach(([,label])=>{const cell=current.createElement('th');cell.style.cssText='border:1px solid #9caec4;padding:6px 8px;background:#eef4fb;text-align:left';cell.textContent=label;headRow.append(cell)});head.append(headRow);table.append(head);
    const body=current.createElement('tbody');rows.forEach(item=>{const row=current.createElement('tr');row.dataset.nouyaReplyRowId=item.row_id;headers.forEach(([field])=>{const cell=current.createElement('td');cell.style.cssText='border:1px solid #9caec4;padding:6px 8px;vertical-align:top';cell.textContent=item[field]||'';row.append(cell)});body.append(row)});table.append(body);reply.append(table);reply.scrollIntoView({block:'start'});
  }
  function applyReplyMatches(data) {
    if(data.generated_table) createReplyTable(data.generated_rows||[]);
    replyRowsById=new Map();replyTables();
    const infosByTable=new Map();
    replyRowsById.forEach(info=>{
      const infos=infosByTable.get(info.table)||[];
      infos.push(info);infosByTable.set(info.table,infos);
    });
    infosByTable.forEach(infos=>{
      const fieldIndex=infos[0].fields.delivery_reply;
      const useExisting=fieldIndex!==undefined&&infos.some(info=>!textAt([deliveryCell(info,fieldIndex)],0));
      const index=ensureReplyColumn(infos[0],useExisting);
      infos.forEach(info=>{info.fields.delivery_reply=index;});
    });
    (data.row_matches||[]).forEach(match=>{
      const info=replyRowsById.get(String(match.row_id));if(!info) return;
      const cell=replyCell(info);
      if(match.status==='matched'&&match.delivery_reply&&!String(cell?.textContent||'').trim()){cell.textContent=match.delivery_reply;cell.removeAttribute('title');}
      else if(match.reason) cell.title=match.reason;
    });
  }
  globalThis.nouyaReplyTables={collect:replyTables,apply:applyReplyMatches};
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
