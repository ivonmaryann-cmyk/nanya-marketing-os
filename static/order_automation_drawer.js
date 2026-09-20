(() => {
  const drawer = document.getElementById('orderMailDrawer');
  if (!drawer) return;
  const content = drawer.querySelector('[data-drawer-content]');
  let loading = false;
  let currentDrawerUrl = '';

  const close = () => {
    drawer.hidden = true;
    content.replaceChildren();
    document.body.classList.remove('order-mail-drawer-open');
  };

  const showError = (message) => {
    content.innerHTML = `<section class="omd-panel"><header class="omd-head"><h2>邮件详情</h2><button type="button" class="omd-close" data-drawer-close aria-label="关闭邮件详情">×</button></header><p class="omd-load-error">${message}</p></section>`;
  };

  const open = async (url) => {
    if (loading) return;
    loading = true;
    currentDrawerUrl = url;
    drawer.hidden = false;
    document.body.classList.add('order-mail-drawer-open');
    content.innerHTML = '<section class="omd-panel"><p class="omd-loading">正在加载邮件详情…</p></section>';
    try {
      const response = await fetch(url, {headers: {'X-Requested-With': 'XMLHttpRequest'}});
      if (!response.ok) throw new Error(response.status === 404 ? '该邮件不存在或无权查看。' : '邮件详情加载失败，请稍后重试。');
      content.innerHTML = await response.text();
    } catch (error) {
      showError(error.message || '邮件详情加载失败，请稍后重试。');
    } finally {
      loading = false;
    }
  };

  document.querySelectorAll('.om-row').forEach((row) => {
    const detailLink = [...row.querySelectorAll('a.om-link')].find((link) => /\/order-automation\/cases\/\d+/.test(link.href));
    if (!detailLink) return;
    const drawerLocation = new URL(detailLink.href);
    drawerLocation.pathname = `${drawerLocation.pathname}/drawer`;
    const drawerUrl = drawerLocation.href;
    row.classList.add('om-row-clickable');
    row.addEventListener('click', (event) => {
      if (event.target.closest('a,button,input,select,textarea,label')) return;
      open(drawerUrl);
    });
  });

  drawer.addEventListener('click', (event) => {
    if (event.target === drawer || event.target.closest('[data-drawer-close]')) {
      close();
      return;
    }
  });

  drawer.addEventListener('submit', async (event) => {
    const form = event.target.closest('.omd-routing-form');
    if (!form) return;
    event.preventDefault();
    const panel = form.closest('.omd-panel');
    const error = form.querySelector('.omd-error');
    const submit = form.querySelector('[type="submit"]');
    submit.disabled = true;
    error.hidden = true;
    try {
      const response = await fetch(panel.dataset.routingUrl, {
        method: 'POST',
        headers: {'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest'},
        body: JSON.stringify(Object.fromEntries(new FormData(form))),
      });
      const result = await response.json();
      if (!response.ok || !result.ok) throw new Error(result.message || '分流保存失败，请稍后重试。');
      await open(currentDrawerUrl);
    } catch (requestError) {
      error.textContent = requestError.message || '分流保存失败，请稍后重试。';
      error.hidden = false;
      submit.disabled = false;
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !drawer.hidden) close();
  });
})();
