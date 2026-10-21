/* Google Analytics 4 — estadísticas internas de visitas (no visible en el sitio).
   Para cambiar de propiedad, solo cambia GA_ID. */
(() => {
  "use strict";

  const GA_ID = "G-YFCXCFHX49";

  const loader = document.createElement("script");
  loader.async = true;
  loader.src = `https://www.googletagmanager.com/gtag/js?id=${GA_ID}`;
  document.head.appendChild(loader);

  window.dataLayer = window.dataLayer || [];
  window.gtag = function gtag() {
    window.dataLayer.push(arguments);
  };
  window.gtag("js", new Date());
  window.gtag("config", GA_ID);
})();
