(() => {
  const normalize = (value) => String(value || "").trim().toLowerCase();

  document.querySelectorAll("table[data-rule-list]").forEach((table) => {
    const body = table.tBodies[0];
    if (!body) return;
    const rows = Array.from(body.querySelectorAll("tr.js-rule-row"));
    if (!rows.length) return;

    const toolbar = document.createElement("div");
    toolbar.className = "rule-list-toolbar";
    const sourceLabel = table.dataset.sourceLabel || "来源";
    const sourceAllLabel = table.dataset.sourceAllLabel || "全部来源";
    toolbar.innerHTML = `
      <label class="rule-list-filter rule-list-source-wrap">
        <span>${sourceLabel}</span><select class="rule-list-source"><option value="">${sourceAllLabel}</option></select>
      </label>
      <label class="rule-list-search">
        <span class="sr-only">搜索规则</span>
        <input type="search" placeholder="${table.dataset.searchPlaceholder || "搜索规则内容、客户或结果"}">
      </label>
      <label class="rule-list-filter rule-list-status-wrap">
        <span>状态</span><select class="rule-list-status"><option value="">全部状态</option></select>
      </label>
      <label class="rule-list-filter">
        <span>每页</span><select class="rule-list-size"><option>20</option><option>50</option><option>100</option></select>
      </label>
      <span class="rule-list-count" aria-live="polite"></span>
    `;

    const tableParent = table.parentElement;
    tableParent.insertBefore(toolbar, table);
    const pager = document.createElement("div");
    pager.className = "rule-list-pager";
    pager.innerHTML = `
      <button type="button" class="rule-list-prev" aria-label="上一页">上一页</button>
      <span class="rule-list-page" aria-live="polite"></span>
      <button type="button" class="rule-list-next" aria-label="下一页">下一页</button>
    `;
    tableParent.insertBefore(pager, table.nextSibling);

    const search = toolbar.querySelector("input[type=search]");
    const source = toolbar.querySelector(".rule-list-source");
    const status = toolbar.querySelector(".rule-list-status");
    const size = toolbar.querySelector(".rule-list-size");
    const count = toolbar.querySelector(".rule-list-count");
    const sourceWrap = toolbar.querySelector(".rule-list-source-wrap");
    const statusWrap = toolbar.querySelector(".rule-list-status-wrap");
    const prev = pager.querySelector(".rule-list-prev");
    const next = pager.querySelector(".rule-list-next");
    const pageLabel = pager.querySelector(".rule-list-page");
    let page = 1;

    const fillOptions = (select, values, preferredOrder = []) => {
      const available = Array.from(values).filter(Boolean);
      const ranked = preferredOrder.filter((value) => available.includes(value));
      available.filter((value) => !ranked.includes(value)).sort().forEach((value) => ranked.push(value));
      ranked.forEach((value) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
      });
    };
    const sources = new Set(rows.map((row) => row.dataset.source || ""));
    const statuses = new Set(rows.map((row) => row.dataset.status || ""));
    const sourceOrder = String(table.dataset.sourceOrder || "").split("|").filter(Boolean);
    fillOptions(source, sources, sourceOrder);
    fillOptions(status, statuses);
    sourceWrap.hidden = !Array.from(sources).some(Boolean);
    statusWrap.hidden = !Array.from(statuses).some(Boolean);

    const render = () => {
      const query = normalize(search.value);
      const sourceValue = source.value;
      const statusValue = status.value;
      const matched = rows.filter((row) => {
        return (!query || normalize(row.textContent).includes(query))
          && (!sourceValue || row.dataset.source === sourceValue)
          && (!statusValue || row.dataset.status === statusValue);
      });
      const pageSize = Number(size.value) || 20;
      const totalPages = Math.max(1, Math.ceil(matched.length / pageSize));
      page = Math.min(page, totalPages);
      const visible = new Set(matched.slice((page - 1) * pageSize, page * pageSize));
      rows.forEach((row) => { row.hidden = !visible.has(row); });
      count.textContent = `共 ${matched.length} 条`;
      pageLabel.textContent = `${page} / ${totalPages}`;
      prev.disabled = page <= 1;
      next.disabled = page >= totalPages;
      pager.hidden = matched.length <= pageSize;
      let empty = body.querySelector("tr.js-filter-empty");
      if (!matched.length && !empty) {
        empty = document.createElement("tr");
        empty.className = "js-filter-empty";
        const cell = document.createElement("td");
        cell.colSpan = table.rows[0]?.cells.length || 1;
        cell.className = "rc-empty cr-empty";
        cell.textContent = "没有找到符合条件的规则。";
        empty.appendChild(cell);
        body.appendChild(empty);
      } else if (empty) {
        empty.remove();
      }
    };

    [search, source, status, size].forEach((control) => {
      control.addEventListener(control === search ? "input" : "change", () => {
        page = 1;
        render();
      });
    });
    prev.addEventListener("click", () => { page -= 1; render(); });
    next.addEventListener("click", () => { page += 1; render(); });
    render();
  });
})();
