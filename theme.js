// Applique le thème choisi dans Réglages (sombre / clair / auto) avant l'affichage, pour éviter un flash de couleur.
(function () {
  var mq = window.matchMedia("(prefers-color-scheme: light)");
  function pref() {
    try { return JSON.parse(localStorage.getItem("pc_theme")) || "dark"; } catch (e) { return "dark"; }
  }
  function apply() {
    var p = pref();
    var t = p === "auto" ? (mq.matches ? "light" : "dark") : p;
    document.documentElement.dataset.theme = t;
    var meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute("content", t === "light" ? "#ffffff" : "#0f1115");
  }
  apply();
  if (mq.addEventListener) mq.addEventListener("change", apply);
  window.applyTheme = apply;
})();
