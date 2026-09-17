(() => {
  const plainText = value => String(value ?? '').replace(/\r\n?/g, '\n');
  const matrixFromClipboard = value => {
    const rows = plainText(value).split('\n').map(row => row.split('\t'));
    if (rows.length > 1 && rows.at(-1).every(cell => cell === '')) rows.pop();
    return rows;
  };

  function createDialog(instance) {
    const dialog = document.createElement('dialog');
    dialog.className = 'template-bulk-dialog';
    dialog.innerHTML = '<form method="dialog"><header><h2>批量填写</h2><button value="cancel" aria-label="关闭">×</button></header><label>字段<select name="field"></select></label><label>填写内容<textarea name="value" rows="4"></textarea></label><p class="template-bulk-dialog-note">将只修改当前勾选的明细行。</p><footer><button value="cancel" class="template-bulk-cancel">取消</button><button value="apply" class="template-bulk-apply">应用</button></footer></form>';
    document.body.append(dialog);
    const form = dialog.querySelector('form'), field = form.elements.field, value = form.elements.value;
    form.addEventListener('submit', event => {
      if (event.submitter?.value !== 'apply') return;
      event.preventDefault();
      instance.applyBatch(field.value, value.value);
      dialog.close();
    });
    return {
      open() {
        const choices = instance.editableFields();
        if (!choices.length) return instance.notice('当前模板没有可批量填写的字段。');
        field.replaceChildren(...choices.map(item => {
          const option = document.createElement('option');
          option.value = item.field;
          option.textContent = item.label;
          return option;
        }));
        value.value = '';
        dialog.showModal();
        value.focus();
      },
    };
  }

  function injectStyle() {
    if (document.getElementById('templateBulkEditorStyle')) return;
    const style = document.createElement('style');
    style.id = 'templateBulkEditorStyle';
    style.textContent = '.template-bulk-selected{box-shadow:inset 0 0 0 2px #8bb9ff!important;background:#eef6ff!important}.template-bulk-active{box-shadow:inset 0 0 0 2px #2f76df!important;background:#e2efff!important}.template-bulk-status{min-height:18px;color:#597694;font-size:12px;font-weight:700}.template-bulk-dialog{width:min(420px,calc(100vw - 32px));padding:0;border:1px solid #d7e3f1;border-radius:8px;color:#294873;box-shadow:0 18px 50px rgba(18,46,84,.25)}.template-bulk-dialog::backdrop{background:rgba(22,42,71,.34)}.template-bulk-dialog form{padding:0}.template-bulk-dialog header,.template-bulk-dialog footer{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid #e1e9f3}.template-bulk-dialog footer{justify-content:flex-end;gap:9px;border-top:1px solid #e1e9f3;border-bottom:0}.template-bulk-dialog h2{margin:0;font-size:18px}.template-bulk-dialog header button{border:0;background:transparent;color:#6c82a0;font-size:24px;cursor:pointer}.template-bulk-dialog label{display:block;margin:16px 20px 0;color:#3f5d84;font-size:13px;font-weight:800}.template-bulk-dialog select,.template-bulk-dialog textarea{box-sizing:border-box;width:100%;margin-top:7px;border:1px solid #cbd9e9;border-radius:5px;padding:9px;background:#fff;color:#284b77;font:inherit;font-size:13px}.template-bulk-dialog-note{margin:10px 20px 0;color:#71839c;font-size:12px}.template-bulk-dialog footer button{border:0;border-radius:6px;padding:9px 14px;font:inherit;font-size:13px;font-weight:800;cursor:pointer}.template-bulk-cancel{border:1px solid #b9cce3!important;background:#fff;color:#526d8f}.template-bulk-apply{background:#397ce8;color:#fff}';
    document.head.append(style);
  }

  function init(options) {
    const rowsRoot = document.querySelector(options.rowsSelector);
    const toolbar = document.querySelector(options.toolbarSelector);
    if (!rowsRoot || !toolbar) return null;
    injectStyle();
    const excluded = new Set(options.excludedFields || []);
    const state = {anchor: null, active: null};
    const rowNodes = () => [...rowsRoot.querySelectorAll(options.rowSelector)];
    const controlFor = (row, field) => [...row.querySelectorAll('[data-field]')]
      .find(control => control.dataset.field === field && control.closest(options.cellSelector));
    const visibleFields = () => {
      const first = rowNodes()[0];
      if (!first) return [];
      return [...first.querySelectorAll('[data-field]')]
        .filter(control => control.closest(options.cellSelector))
        .map(control => control.dataset.field)
        .filter((field, index, fields) => fields.indexOf(field) === index);
    };
    const positionFor = control => {
      const row = control?.closest(options.rowSelector);
      if (!row || !rowsRoot.contains(row) || !control.closest(options.cellSelector)) return null;
      const rowIndex = rowNodes().indexOf(row), fieldIndex = visibleFields().indexOf(control.dataset.field);
      return rowIndex < 0 || fieldIndex < 0 ? null : {row: rowIndex, field: fieldIndex};
    };
    const cellFor = position => {
      const field = visibleFields()[position.field];
      return field === undefined ? null : controlFor(rowNodes()[position.row], field)?.closest(options.cellSelector);
    };
    const editable = control => Boolean(control && !control.disabled && !control.readOnly && !excluded.has(control.dataset.field));
    const selectedBounds = () => {
      if (!state.anchor || !state.active) return null;
      return {
        firstRow: Math.min(state.anchor.row, state.active.row), lastRow: Math.max(state.anchor.row, state.active.row),
        firstField: Math.min(state.anchor.field, state.active.field), lastField: Math.max(state.anchor.field, state.active.field),
      };
    };
    const clearHighlight = () => rowsRoot.querySelectorAll('.template-bulk-selected,.template-bulk-active').forEach(cell => {
      cell.classList.remove('template-bulk-selected', 'template-bulk-active');
    });
    const drawSelection = () => {
      clearHighlight();
      const bounds = selectedBounds();
      if (!bounds) return;
      for (let row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
        for (let field = bounds.firstField; field <= bounds.lastField; field += 1) {
          cellFor({row, field})?.classList.add('template-bulk-selected');
        }
      }
      cellFor(state.active)?.classList.add('template-bulk-active');
    };
    const status = document.createElement('span');
    status.className = 'template-bulk-status';
    toolbar.append(status);
    const notice = message => { status.textContent = message; };
    const setValue = (control, value) => {
      if (!editable(control)) return false;
      control.value = value;
      control.dispatchEvent(new Event('input', {bubbles: true}));
      control.dispatchEvent(new Event('change', {bubbles: true}));
      return true;
    };
    const instance = {
      notice,
      editableFields: () => visibleFields().map(field => ({field, label: options.fieldLabels?.[field] || field}))
        .filter(item => editable(controlFor(rowNodes()[0], item.field))),
      applyBatch(field, value) {
        const rows = rowNodes().filter(row => row.querySelector(options.selectionSelector)?.checked);
        if (!rows.length) return notice('请先勾选需要批量修改的明细行。');
        const changed = rows.reduce((count, row) => count + Number(setValue(controlFor(row, field), value)), 0);
        notice(changed ? `已批量修改 ${changed} 个单元格，保存后生效。` : '所选字段当前不可编辑。');
      },
      fillDown() {
        const bounds = selectedBounds();
        if (!bounds || bounds.firstRow === bounds.lastRow) return notice('请用 Shift + 单击选择至少两行后再向下填充。');
        let changed = 0;
        for (let field = bounds.firstField; field <= bounds.lastField; field += 1) {
          const source = controlFor(rowNodes()[bounds.firstRow], visibleFields()[field]);
          for (let row = bounds.firstRow + 1; row <= bounds.lastRow; row += 1) {
            changed += Number(setValue(controlFor(rowNodes()[row], visibleFields()[field]), source?.value || ''));
          }
        }
        notice(changed ? `已向下填充 ${changed} 个单元格，保存后生效。` : '选区内没有可编辑的单元格。');
      },
      copy(event) {
        const bounds = selectedBounds();
        if (!bounds) return false;
        const values = [];
        for (let row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
          values.push([...Array(bounds.lastField - bounds.firstField + 1)].map((_, offset) =>
            controlFor(rowNodes()[row], visibleFields()[bounds.firstField + offset])?.value || '').join('\t'));
        }
        event.clipboardData?.setData('text/plain', values.join('\n'));
        return true;
      },
      paste(event) {
        const start = state.active;
        if (!start) return false;
        const values = matrixFromClipboard(event.clipboardData?.getData('text/plain') || '');
        if (!values.length || !values[0].length) return false;
        const rows = rowNodes(), fields = visibleFields();
        let changed = 0, ignoredRows = 0, ignoredColumns = 0, readOnly = 0;
        values.forEach((sourceRow, offsetRow) => {
          const rowIndex = start.row + offsetRow;
          if (rowIndex >= rows.length) { ignoredRows += 1; return; }
          sourceRow.forEach((value, offsetField) => {
            const fieldIndex = start.field + offsetField;
            if (fieldIndex >= fields.length) { ignoredColumns += 1; return; }
            const control = controlFor(rows[rowIndex], fields[fieldIndex]);
            if (setValue(control, value)) changed += 1; else readOnly += 1;
          });
        });
        const ignored = [ignoredRows ? `${ignoredRows} 行超出模板` : '', ignoredColumns ? `${ignoredColumns} 格超出列范围` : '', readOnly ? `${readOnly} 格只读` : ''].filter(Boolean);
        notice(`已粘贴 ${changed} 个单元格${ignored.length ? `；${ignored.join('，')}未写入` : ''}。`);
        return true;
      },
    };
    const dialog = createDialog(instance);
    const fillButton = toolbar.querySelector('[data-bulk-fill-down]');
    const batchButton = toolbar.querySelector('[data-bulk-edit]');
    fillButton?.addEventListener('click', instance.fillDown);
    batchButton?.addEventListener('click', dialog.open);
    rowsRoot.addEventListener('pointerdown', event => {
      const control = event.target.closest('[data-field]');
      const position = positionFor(control);
      if (!position) return;
      if (event.shiftKey && state.anchor) state.active = position;
      else state.anchor = state.active = position;
      drawSelection();
    });
    rowsRoot.addEventListener('copy', event => {
      if (event.target.matches('input,textarea') && event.target.selectionStart !== event.target.selectionEnd) return;
      if (instance.copy(event)) event.preventDefault();
    });
    rowsRoot.addEventListener('paste', event => {
      if (instance.paste(event)) event.preventDefault();
    });
    rowsRoot.addEventListener('keydown', event => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'd') {
        event.preventDefault();
        instance.fillDown();
      }
    });
    return instance;
  }

  globalThis.nouyaTemplateBulkEditor = {init};
})();
