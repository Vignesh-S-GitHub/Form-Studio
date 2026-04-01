(function registerPwa() {
  if (!('serviceWorker' in navigator)) {
    return;
  }

  window.addEventListener('load', function () {
    navigator.serviceWorker.register('./sw.js').catch(function () {
      // Silent fail: app remains fully usable without SW.
    });
  });
})();
