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
    style.textContent = '.template-bulk-selected{box-shadow:inset 0 0 0 2px #8bb9ff!important;background:#eef6ff!important}.template-bulk-active{box-shadow:inset 0 0 0 2px #2f76df!important;background:#e2efff!important}.template-filter-hidden{display:none!important}.template-bulk-status{min-height:18px;color:#597694;font-size:12px;font-weight:700}.template-bulk-dialog,.template-filter-dialog{width:min(420px,calc(100vw - 32px));padding:0;border:1px solid #d7e3f1;border-radius:8px;color:#294873;box-shadow:0 18px 50px rgba(18,46,84,.25)}.template-bulk-dialog::backdrop,.template-filter-dialog::backdrop{background:rgba(22,42,71,.34)}.template-bulk-dialog form,.template-filter-dialog form{padding:0}.template-bulk-dialog header,.template-bulk-dialog footer,.template-filter-dialog header,.template-filter-dialog footer{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid #e1e9f3}.template-bulk-dialog footer,.template-filter-dialog footer{justify-content:flex-end;gap:9px;border-top:1px solid #e1e9f3;border-bottom:0}.template-bulk-dialog h2,.template-filter-dialog h2{margin:0;font-size:18px}.template-bulk-dialog header button,.template-filter-dialog header button{border:0;background:transparent;color:#6c82a0;font-size:24px;cursor:pointer}.template-bulk-dialog label,.template-filter-dialog>form>label{display:block;margin:16px 20px 0;color:#3f5d84;font-size:13px;font-weight:800}.template-bulk-dialog select,.template-bulk-dialog textarea,.template-filter-dialog input[type=search]{box-sizing:border-box;width:100%;margin-top:7px;border:1px solid #cbd9e9;border-radius:5px;padding:9px;background:#fff;color:#284b77;font:inherit;font-size:13px}.template-bulk-dialog-note,.template-filter-note{margin:10px 20px 0;color:#71839c;font-size:12px}.template-bulk-dialog footer button,.template-filter-dialog footer button,.template-filter-clear,.template-filter-trigger{border:0;border-radius:6px;padding:9px 14px;font:inherit;font-size:13px;font-weight:800;cursor:pointer}.template-bulk-cancel,.template-filter-dialog footer button:not([value=apply]),.template-filter-clear,.template-filter-trigger{border:1px solid #b9cce3!important;background:#fff;color:#526d8f}.template-bulk-apply,.template-filter-dialog button[value=apply]{background:#397ce8;color:#fff}.template-filter-values{max-height:250px;overflow:auto;margin:12px 20px 0;border:1px solid #e0e8f2;border-radius:6px}.template-filter-values label{display:flex;gap:8px;align-items:center;margin:0;padding:8px 10px;border-bottom:1px solid #edf1f6;color:#405e84;font-size:13px;font-weight:600}.template-filter-values label:last-child{border-bottom:0}.template-filter-field{box-sizing:border-box!important;flex:0 0 160px!important;width:160px!important;min-width:160px!important;max-width:160px!important;height:36px;border:1px solid #b9cce3;border-radius:6px;padding:0 9px;background:#fff;color:#526d8f;font:inherit;font-size:13px;font-weight:700}.template-filter-trigger:hover{border-color:#397ce8!important;background:#f2f7ff;color:#286fd1}';
    document.head.append(style);
  }

  function init(options) {
    const rowsRoot = document.querySelector(options.rowsSelector);
    const toolbar = document.querySelector(options.toolbarSelector);
    if (!rowsRoot || !toolbar) return null;
    injectStyle();
    const excluded = new Set(options.excludedFields || []);
    const state = {anchor: null, active: null};
    const allRows = () => [...rowsRoot.querySelectorAll(options.rowSelector)];
    const rowNodes = () => allRows().filter(row => !row.hidden);
    const controlFor = (row, field) => row && [...row.querySelectorAll('[data-field]')]
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
        const rows = allRows().filter(row => row.querySelector(options.selectionSelector)?.checked);
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

    const filterableFields = () => visibleFields().filter(field => !excluded.has(field));
    const filters = new Map();
    const filterStatus = document.createElement('span');
    filterStatus.className = 'template-bulk-status';
    const clearFilters = document.createElement('button');
    clearFilters.type = 'button';
    clearFilters.className = 'template-filter-clear';
    clearFilters.textContent = '清除筛选';
    clearFilters.hidden = true;
    const filterDialog = document.createElement('dialog');
    filterDialog.className = 'template-filter-dialog';
    filterDialog.innerHTML = '<form method="dialog"><header><h2></h2><button value="cancel" aria-label="关闭">×</button></header><label>关键词<input type="search" name="query" placeholder="包含的内容"></label><div class="template-filter-values"></div><p class="template-filter-note">不勾选具体值时，仅按关键词筛选。</p><footer><button type="button" data-filter-clear>清除此列</button><button value="cancel">取消</button><button value="apply">应用</button></footer></form>';
    document.body.append(filterDialog);
    const filterForm = filterDialog.querySelector('form');
    const filterTitle = filterDialog.querySelector('h2');
    const filterQuery = filterForm.elements.query;
    const filterValues = filterDialog.querySelector('.template-filter-values');
    let activeFilterField = '';
    const valueFor = (row, field) => String(controlFor(row, field)?.value || '').trim();
    const specialFilterOptions = field => options.specialFilterOptions?.(field) || [];
    const filterMatches = row => [...filters.entries()].every(([field, filter]) => {
      const value = valueFor(row, field).toLocaleLowerCase();
      const specialMatches = [...(filter.specialValues || [])].some(value =>
        options.specialFilterMatches?.(row, field, value));
      const hasValueFilters = filter.values.size || (filter.specialValues || new Set()).size;
      return (!filter.query || value.includes(filter.query))
        && (!hasValueFilters || filter.values.has(value) || specialMatches);
    });
    const applyFilters = () => {
      let visible = 0;
      allRows().forEach(row => {
        const matches = filterMatches(row);
        row.hidden = !matches;
        row.classList.toggle('template-filter-hidden', !matches);
        const detail = row.nextElementSibling;
        if (detail?.classList.contains('oc-match-detail')) {
          if (!detail.dataset.templateFilterWasHidden) detail.dataset.templateFilterWasHidden = detail.hidden ? '1' : '0';
          detail.hidden = !matches || detail.dataset.templateFilterWasHidden === '1';
        }
        if (matches) visible += 1;
      });
      state.anchor = state.active = null;
      drawSelection();
      filterStatus.textContent = `筛选：显示 ${visible}/${allRows().length} 项`;
      const active = filters.size > 0;
      clearFilters.hidden = !active;
    };
    const openFilter = field => {
      activeFilterField = field;
      const current = filters.get(field) || {query: '', values: new Set(), specialValues: new Set()};
      filterTitle.textContent = `筛选：${options.fieldLabels?.[field] || field}`;
      filterQuery.value = current.query;
      const values = [...new Set(allRows().map(row => valueFor(row, field)).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-CN'));
      const specialValues = specialFilterOptions(field).map(item => {
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox'; checkbox.name = 'special-value'; checkbox.value = item.value;
        checkbox.checked = current.specialValues.has(item.value);
        label.append(checkbox, document.createTextNode(item.label));
        return label;
      });
      filterValues.replaceChildren(...specialValues, ...values.map(value => {
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox'; checkbox.name = 'value'; checkbox.value = value.toLocaleLowerCase();
        checkbox.checked = current.values.has(checkbox.value);
        label.append(checkbox, document.createTextNode(value));
        return label;
      }));
      filterDialog.showModal();
      filterQuery.focus();
    };
    filterForm.addEventListener('submit', event => {
      if (event.submitter?.value !== 'apply') return;
      event.preventDefault();
      const query = filterQuery.value.trim().toLocaleLowerCase();
      const values = new Set([...filterValues.querySelectorAll('input[name="value"]:checked')].map(input => input.value));
      const specialValues = new Set([...filterValues.querySelectorAll('input[name="special-value"]:checked')].map(input => input.value));
      if (query || values.size || specialValues.size) filters.set(activeFilterField, {query, values, specialValues}); else filters.delete(activeFilterField);
      filterDialog.close(); applyFilters();
    });
    filterForm.querySelector('[data-filter-clear]').addEventListener('click', () => {
      filters.delete(activeFilterField); filterDialog.close(); applyFilters();
    });
    clearFilters.addEventListener('click', () => { filters.clear(); applyFilters(); });
    const filterField = document.createElement('select');
    filterField.className = 'template-filter-field';
    filterField.setAttribute('aria-label', '选择筛选列');
    filterField.append(new Option('选择筛选列', ''));
    filterableFields().forEach(field => filterField.append(new Option(options.fieldLabels?.[field] || field, field)));
    const filterTrigger = document.createElement('button');
    filterTrigger.type = 'button';
    filterTrigger.className = 'template-filter-trigger';
    filterTrigger.textContent = '筛选';
    filterTrigger.addEventListener('click', () => {
      if (!filterField.value) return notice('请先选择需要筛选的列。');
      openFilter(filterField.value);
    });
    toolbar.append(filterField, filterTrigger, filterStatus, clearFilters);
    rowsRoot.addEventListener('input', () => { if (filters.size) applyFilters(); });
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
    applyFilters();
    return instance;
  }

  globalThis.nouyaTemplateBulkEditor = {init};
})();
