/* Shared: dark-mode toggle. */
(function () {
  const btn = document.getElementById('themeBtn');
  if (!btn) return;
  btn.addEventListener('click', () => {
    const dark = document.documentElement.dataset.theme !== 'dark';
    if (dark) document.documentElement.dataset.theme = 'dark'; else delete document.documentElement.dataset.theme;
    try { localStorage.setItem('credit-theme', dark ? 'dark' : 'light'); } catch (e) {}
  });
})();
