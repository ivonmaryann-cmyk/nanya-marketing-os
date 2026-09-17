(() => {
  // Reuse the field controls and modes; arrange one row per actual encoding segment.
  const parameters = document.getElementById('pn-parameters');
  const structures = JSON.parse(document.getElementById('pn-structure-data')?.textContent || '[]');
  if (parameters) {
    const board = !!document.getElementById('pn-copper-mode');
    const groups = board
      ? [['glue'], ['thickness'], ['board_spec'], ['copper_mode', '#pn-copper-standard', '#pn-copper-split'], ['structure'], ['marking'], ['size_mode', '#pn-sheet'], ['board_grade']]
      : [['glue'], ['glass_style'], ['glass_type'], ['pp_mode', 'pp_value'], ['pp_spec'], ['marking'], ['size_mode', '#pn-sheet', '#pn-roll'], ['pp_grade']];
    const grid = parameters.querySelector(':scope > .pn-grid');
    const wrap = document.createElement('div'); wrap.className = 'pn-check-wrap';
    const table = document.createElement('table'); table.className = 'pn-check-table';
    const head = table.createTHead().insertRow();
    ['参数', '识别结果 / 可修改', '生成编码', '位数 / 位置', '来源'].forEach(text => {
      const th = document.createElement('th'); th.scope = 'col'; th.textContent = text; head.append(th);
    });
    const body = table.createTBody(); let position = 1;
    groups.forEach((fields, index) => {
      const [name, width] = structures[index];
      const row = body.insertRow(); row.dataset.segment = index;
      const title = row.insertCell(); title.textContent = name;
      const editor = row.insertCell(); editor.className = 'pn-check-editor';
      fields.forEach(field => {
        const node = field.startsWith('#') ? parameters.querySelector(field) : parameters.querySelector(`[name="${field}"]`)?.closest('label');
        if (node) editor.append(node);
      });
      if (fields.length === 1) {
        const label = editor.querySelector('label');
        label?.querySelector('input,select')?.setAttribute('aria-label', name);
        label?.childNodes.forEach(node => { if (node.nodeType === Node.TEXT_NODE) node.textContent = ''; });
      }
      const codeCell = row.insertCell(); const code = document.createElement('code'); code.dataset.segmentCode = index; code.textContent = '—'; codeCell.append(code);
      const location = row.insertCell(); location.textContent = width + '位';
      const range = document.createElement('small'); range.textContent = position === position + width - 1 ? `第${position}位` : `第${position}–${position + width - 1}位`; location.append(range);
      row.dataset.position = range.textContent; row.dataset.width = width; position += width;
      const source = row.insertCell(); source.className = 'pn-check-source';
      const output = document.createElement('span'); output.dataset.rowSource = index; output.textContent = '待转换'; source.append(output);
      // Keep individual evidence attached to its control so switching modes remains accurate.
      editor.querySelectorAll('[data-evidence]').forEach(el => { el.hidden = true; });
    });
    wrap.append(table); grid.replaceWith(wrap);
    const summary = parameters.querySelector('.pn-live');
    parameters.prepend(summary);
    const total = document.createElement('small'); total.className = 'pn-total-width';
    total.textContent = `共 ${position - 1} 位 · ${structures.length} 个编码段`;
    summary.querySelector('h2').append(total);
    const code = summary.querySelector('#pn-live-code'); summary.querySelector('h2').after(code);
  }
  const dialog = document.querySelector('.pn-dialog');
  if (dialog && typeof dialog.showModal === 'function') {
    dialog.removeAttribute('open');
    dialog.showModal();
    dialog.addEventListener('cancel', () => { window.location.href = dialog.querySelector('a').href; });
  }
  const mode = document.getElementById('pn-size-mode');
  function sizeFields() {
    if (!mode) return;
    for (const [id, active] of [['pn-sheet', mode.value === 'sheet'], ['pn-roll', mode.value === 'roll']]) {
      const section = document.getElementById(id);
      if (!section) continue;
      section.hidden = !active;
      section.querySelectorAll('input,select').forEach(el => { el.disabled = !active; el.required = active; });
    }
  }
  mode?.addEventListener('change', sizeFields);
  sizeFields();
  const copper = document.getElementById('pn-copper-mode');
  function copperFields() {
    if (!copper) return;
    for (const type of ['standard', 'split']) {
      const section = document.getElementById('pn-copper-' + type);
      section.hidden = copper.value !== type;
      section.querySelectorAll('input,select').forEach(el => { el.disabled = section.hidden; el.required = !section.hidden; });
    }
  }
  copper?.addEventListener('change', copperFields);
  copperFields();
  const form = document.querySelector('.pn-convert[data-preview]');
  if (!form) return;
  const status = document.getElementById('pn-status');
  const save = document.getElementById('pn-save-result');
  let sequence = 0, timer, controller;
  let hasConverted = !parameters.hidden, sourceChanged = false;
  function updateSources() {
    parameters.querySelectorAll('[data-segment]').forEach(row => {
      const texts = Array.from(row.querySelectorAll('[data-evidence]'))
        .filter(el => !el.closest('[hidden]') || el.closest('[hidden]') === el)
        .filter(el => !el.parentElement.querySelector('input,select')?.disabled)
        .map(el => el.textContent).filter(Boolean);
      row.querySelector('[data-row-source]').textContent = texts.some(t => /业务修改/.test(t)) ? '已修改'
        : texts.some(t => /默认|通用项/.test(t)) ? '自动填写 · 请确认'
        : texts.some(t => /未识别/.test(t)) ? '待补充'
        : texts.length ? '规格提取' : '请核对';
    });
  }
  function describeChoices() {
    form.querySelectorAll('input[list]').forEach(input => {
      let description = input.parentElement.querySelector('.pn-choice-description');
      if (!description) {
        description = document.createElement('small');
        description.className = 'pn-choice-description';
        input.insertAdjacentElement('afterend', description);
      }
      const match = Array.from(input.list?.options || []).find(option => option.value === input.value.trim().toUpperCase());
      description.textContent = match?.textContent || (input.value ? '未命中启用映射，需核对码值' : '');
    });
  }
  function invalidate() {
    sequence += 1;
    controller?.abort();
    save.disabled = true;
    document.getElementById('pn-live-code').textContent = '等待重新校验';
    document.getElementById('pn-live-segments').replaceChildren();
    parameters.querySelectorAll('[data-segment-code]').forEach(el => { el.textContent = '—'; });
  }
  async function preview(extract) {
    if (!extract && sourceChanged) return;
    invalidate();
    const current = sequence;
    controller = new AbortController();
    status.textContent = extract ? '正在识别客户规格…' : '正在校验参数…';
    const values = Object.fromEntries(new FormData(form));
    try {
      const response = await fetch(form.dataset.preview, {method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({csrf:values.csrf,values,extract}),signal:controller.signal});
      if (!response.ok) throw new Error(response.status === 422 ? '请输入客户规格' : '校验失败，请刷新页面或稍后重试');
      const data = await response.json();
      if (current !== sequence) return;
      if (extract) {
        hasConverted = true;
        sourceChanged = false;
        form.querySelectorAll('#pn-parameters input,#pn-parameters select').forEach(el => {
          el.value = data.values[el.name] || '';
        });
        if (!copper?.value && copper) copper.value = 'split';
        sizeFields(); copperFields();
        describeChoices();
        document.getElementById('pn-parameters').hidden = false;
        form.querySelectorAll('[data-evidence]').forEach(el => {
          el.textContent = data.evidence[el.dataset.evidence] || '未识别，请补充核对';
        });
        updateSources();
      }
      status.textContent = (data.warnings || []).join('；');
      document.getElementById('pn-live-error').textContent = data.error || '请核对默认项及客户特殊要求，再确认保存。';
      document.getElementById('pn-live-code').textContent = data.result?.code || '参数尚不完整，暂不生成完整码值';
      const segments = document.getElementById('pn-live-segments');
      segments.replaceChildren();
      (data.result?.segments || []).forEach((part, index) => {
        const row = parameters.querySelector(`[data-segment="${index}"]`);
        row.querySelector('[data-segment-code]').textContent = part.code;
        const span = document.createElement('button'), label = document.createElement('small'), code = document.createElement('strong'), position = document.createElement('small');
        span.type = 'button'; span.className = 'pn-segment-link';
        label.textContent = part.label; code.textContent = part.code;
        position.textContent = `${row.dataset.width}位 · ${row.dataset.position}`;
        span.append(label,code,position); segments.append(span);
        span.addEventListener('click', () => {
          parameters.querySelectorAll('.pn-current-row').forEach(el => el.classList.remove('pn-current-row'));
          row.classList.add('pn-current-row'); row.querySelector('input:not(:disabled),select:not(:disabled)')?.focus();
        });
      });
      save.disabled = !data.result;
    } catch (error) {
      if (error.name !== 'AbortError' && current === sequence) status.textContent = error.message;
    }
  }
  document.getElementById('pn-extract').addEventListener('click', () => {clearTimeout(timer);preview(true);});
  form.addEventListener('input', event => {
    clearTimeout(timer); invalidate();
    if (['source_spec','customer','requirements'].includes(event.target.name)) {
      sourceChanged = hasConverted;
      status.textContent = hasConverted ? '输入已变化，请重新转换，避免使用旧结果。' : ''; return;
    }
    if (sourceChanged) return;
    const label = event.target.closest('label')?.querySelector('[data-evidence]');
    if (label) label.textContent = '业务修改';
    const row = event.target.closest('[data-segment]');
    if (row) row.querySelector('[data-row-source]').textContent = '已修改';
    describeChoices();
    timer = setTimeout(() => preview(false), 250);
  });
  form.addEventListener('submit', event => { if (save.disabled) event.preventDefault(); });
  if (!document.getElementById('pn-parameters').hidden) preview(false);
})();
