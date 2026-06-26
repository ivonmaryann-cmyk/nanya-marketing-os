(() => {
  const form = document.getElementById("single-quote-form");
  if (!form) return;

  const specInput = document.getElementById("single-spec");
  const submit = document.getElementById("single-quote-submit");
  const result = document.getElementById("single-quote-result");
  const valueEl = document.getElementById("single-quote-price");
  const metaEl = document.getElementById("single-quote-meta");
  const noteEl = document.getElementById("single-quote-note");

  const endpoint = form.dataset.endpoint;
  const submitText = form.dataset.submitText || submit.textContent || "计算单条";
  const loadingMeta = form.dataset.loadingMeta || "正在读取当前规则";
  const emptyTitle = form.dataset.emptyTitle || "未计算";
  const successFallback = form.dataset.successFallback || "计算成功";
  const missTitle = form.dataset.missTitle || "未命中";

  function setState(className, title, meta, note) {
    result.classList.remove("is-success", "is-error", "is-loading");
    if (className) result.classList.add(className);
    valueEl.textContent = title;
    metaEl.textContent = meta;
    noteEl.textContent = note;
  }

  function quoteValue(data) {
    if (data.price !== null && data.price !== undefined && data.price !== "") {
      const numeric = Number(data.price);
      return Number.isFinite(numeric) ? numeric.toFixed(2) : String(data.price);
    }
    return data.result || data.code || "";
  }

  form.addEventListener("submit", async event => {
    event.preventDefault();
    const spec = specInput.value.trim();
    if (!spec) {
      setState("is-error", emptyTitle, "请输入客户规格", "粘贴一条完整客户规格后再计算。");
      return;
    }

    submit.disabled = true;
    submit.textContent = "正在计算...";
    setState("is-loading", "计算中", loadingMeta, "请稍候。");

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ spec })
      });
      const data = await response.json();
      const value = quoteValue(data);
      if (response.ok && value) {
        const meta = [data.material_type || "规格", data.rule_version || "当前规则"].filter(Boolean).join(" · ");
        setState("is-success", value, meta, data.note || successFallback);
      } else {
        setState("is-error", missTitle, data.status || "计算失败", data.error || data.note || "无法计算该规格");
      }
    } catch (error) {
      setState("is-error", "请求失败", "请稍后重试", error.message || "网络或服务异常");
    } finally {
      submit.disabled = false;
      submit.textContent = submitText;
    }
  });
})();
