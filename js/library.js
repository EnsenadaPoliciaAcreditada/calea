(() => {
  "use strict";

  const input = document.getElementById("docSearch");
  const rows = Array.from(document.querySelectorAll(".doc-row"));
  const countEl = document.getElementById("docCount");
  const emptyEl = document.getElementById("docEmpty");

  if (!input || !rows.length) return;

  const filter = () => {
    const q = input.value.trim().toLowerCase();
    let visible = 0;

    rows.forEach((row) => {
      const match = !q || row.dataset.search.includes(q);
      row.style.display = match ? "" : "none";
      if (match) visible += 1;
    });

    countEl.textContent = `${visible} documento${visible === 1 ? "" : "s"}`;
    emptyEl.classList.toggle("is-visible", visible === 0);
  };

  input.addEventListener("input", filter);
})();
