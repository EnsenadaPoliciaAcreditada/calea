(() => {
  "use strict";

  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  /* ---------------------------------------------------------------------
     Header: fondo sólido al hacer scroll
     --------------------------------------------------------------------- */
  const header = document.getElementById("siteHeader");
  const onScrollHeader = () => {
    header.classList.toggle("is-scrolled", window.scrollY > 24);
  };
  onScrollHeader();
  window.addEventListener("scroll", onScrollHeader, { passive: true });

  /* ---------------------------------------------------------------------
     Menú móvil
     --------------------------------------------------------------------- */
  const navToggle = document.getElementById("navToggle");
  const mainNav = document.getElementById("mainNav");

  const closeNav = () => {
    mainNav.classList.remove("is-open");
    navToggle.setAttribute("aria-expanded", "false");
    document.body.style.overflow = "";
  };

  navToggle.addEventListener("click", () => {
    const isOpen = mainNav.classList.toggle("is-open");
    navToggle.setAttribute("aria-expanded", String(isOpen));
    document.body.style.overflow = isOpen ? "hidden" : "";
  });

  mainNav.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", closeNav);
  });

  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeNav();
  });

  /* ---------------------------------------------------------------------
     Resalte del enlace de navegación activo según la sección visible
     --------------------------------------------------------------------- */
  const navLinks = Array.from(document.querySelectorAll(".nav-list a"));
  const sections = navLinks
    .map((link) => document.querySelector(link.getAttribute("href")))
    .filter(Boolean);

  if ("IntersectionObserver" in window && sections.length) {
    const navObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          const link = navLinks.find((a) => a.getAttribute("href") === `#${entry.target.id}`);
          if (!link) return;
          if (entry.isIntersecting) {
            navLinks.forEach((a) => a.classList.remove("is-active"));
            link.classList.add("is-active");
          }
        });
      },
      { rootMargin: "-45% 0px -50% 0px", threshold: 0 }
    );
    sections.forEach((section) => navObserver.observe(section));
  }

  /* ---------------------------------------------------------------------
     Revelado on-scroll (fade + translate) vía IntersectionObserver
     --------------------------------------------------------------------- */
  const revealEls = document.querySelectorAll(".reveal");

  if ("IntersectionObserver" in window) {
    const revealObserver = new IntersectionObserver(
      (entries, observer) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("is-visible");
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.15, rootMargin: "0px 0px -8% 0px" }
    );
    revealEls.forEach((el) => revealObserver.observe(el));
  } else {
    revealEls.forEach((el) => el.classList.add("is-visible"));
  }

  /* ---------------------------------------------------------------------
     Contadores animados (estadísticas)
     --------------------------------------------------------------------- */
  const counters = document.querySelectorAll(".js-counter");

  const animateCounter = (el) => {
    const target = Number(el.dataset.count || 0);
    if (reducedMotion || !target) {
      el.textContent = target;
      return;
    }
    const duration = 1400;
    const start = performance.now();

    const step = (now) => {
      const progress = Math.min((now - start) / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3);
      el.textContent = Math.round(eased * target);
      if (progress < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  };

  if ("IntersectionObserver" in window && counters.length) {
    const counterObserver = new IntersectionObserver(
      (entries, observer) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            animateCounter(entry.target);
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.6 }
    );
    counters.forEach((el) => counterObserver.observe(el));
  } else {
    counters.forEach((el) => (el.textContent = el.dataset.count));
  }

  /* ---------------------------------------------------------------------
     Organigrama — ramas expandibles con acordeón
     --------------------------------------------------------------------- */
  const orgChart = document.getElementById("orgChart");
  if (orgChart) {
    const branchButtons = Array.from(orgChart.querySelectorAll(".org-branch[data-branch]"));
    const panels = Array.from(orgChart.querySelectorAll(".org-detail-panel"));

    const closeAllPanels = () => {
      branchButtons.forEach((b) => b.setAttribute("aria-expanded", "false"));
      panels.forEach((p) => (p.hidden = true));
    };

    const openBranch = (slug) => {
      closeAllPanels();
      const btn = branchButtons.find((b) => b.dataset.branch === slug);
      const panel = panels.find((p) => p.dataset.panel === slug);
      if (!btn || !panel) return;
      btn.setAttribute("aria-expanded", "true");
      panel.hidden = false;
      panel.scrollIntoView({ behavior: reducedMotion ? "auto" : "smooth", block: "nearest" });
    };

    branchButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const isOpen = btn.getAttribute("aria-expanded") === "true";
        if (isOpen) {
          closeAllPanels();
        } else {
          openBranch(btn.dataset.branch);
        }
      });
    });

    orgChart.querySelectorAll(".org-detail-close").forEach((closeBtn) => {
      closeBtn.addEventListener("click", closeAllPanels);
    });
  }

  /* ---------------------------------------------------------------------
     Barra de progreso — directivas publicadas
     --------------------------------------------------------------------- */
  const progressFill = document.getElementById("docsProgressFill");
  if (progressFill && "IntersectionObserver" in window) {
    const progressObserver = new IntersectionObserver(
      (entries, observer) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            progressFill.style.width = `${progressFill.dataset.target}%`;
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.6 }
    );
    progressObserver.observe(progressFill);
  }

  /* ---------------------------------------------------------------------
     Timeline: línea de progreso dorada + activación de puntos
     --------------------------------------------------------------------- */
  const timeline = document.getElementById("timeline");
  const timelineFill = document.getElementById("timelineFill");
  const timelineItems = document.querySelectorAll(".timeline-item");

  if ("IntersectionObserver" in window && timelineItems.length) {
    const itemObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          entry.target.classList.toggle("is-visible", entry.isIntersecting || entry.target.classList.contains("is-visible"));
        });
      },
      { threshold: 0.5, rootMargin: "0px 0px -10% 0px" }
    );
    timelineItems.forEach((item) => itemObserver.observe(item));
  }

  if (timeline && timelineFill) {
    let ticking = false;
    const updateTimelineFill = () => {
      const rect = timeline.getBoundingClientRect();
      const viewportH = window.innerHeight;
      const total = rect.height;
      const visible = Math.min(Math.max(viewportH * 0.75 - rect.top, 0), total);
      const ratio = total > 0 ? visible / total : 0;
      timelineFill.style.transform = `scaleY(${ratio})`;
      ticking = false;
    };

    window.addEventListener(
      "scroll",
      () => {
        if (!ticking) {
          requestAnimationFrame(updateTimelineFill);
          ticking = true;
        }
      },
      { passive: true }
    );
    window.addEventListener("resize", updateTimelineFill);
    updateTimelineFill();
  }

  /* ---------------------------------------------------------------------
     Desplazamiento suave con offset de header fijo
     --------------------------------------------------------------------- */
  document.querySelectorAll('a[href^="#"]').forEach((link) => {
    link.addEventListener("click", (e) => {
      const id = link.getAttribute("href");
      if (id.length < 2) return;
      const target = document.querySelector(id);
      if (!target) return;
      e.preventDefault();
      const offset = header.offsetHeight + 12;
      const top = target.getBoundingClientRect().top + window.scrollY - offset;
      window.scrollTo({ top, behavior: reducedMotion ? "auto" : "smooth" });
      history.pushState(null, "", id);
    });
  });

  /* ---------------------------------------------------------------------
     Formulario de contacto — envío vía FormSubmit.co (sitio estático, sin backend propio)
     --------------------------------------------------------------------- */
  const contactForm = document.getElementById("contactForm");
  if (contactForm) {
    const note = document.getElementById("cfNote");
    const btn = contactForm.querySelector("button[type='submit']");
    const originalBtnText = btn.textContent;

    contactForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      btn.disabled = true;
      btn.textContent = "Enviando…";
      if (note) {
        note.textContent = "";
        note.classList.remove("field-note--error");
      }

      try {
        const ajaxUrl = contactForm.action.replace("formsubmit.co/", "formsubmit.co/ajax/");
        const response = await fetch(ajaxUrl, {
          method: "POST",
          headers: { Accept: "application/json" },
          body: new FormData(contactForm),
        });

        if (!response.ok) throw new Error("network");

        btn.textContent = "Mensaje enviado";
        if (note) note.textContent = "Gracias, tu mensaje fue enviado correctamente.";
        contactForm.reset();
      } catch (err) {
        btn.textContent = originalBtnText;
        if (note) {
          note.textContent = "No se pudo enviar el mensaje. Intenta de nuevo o escríbenos directamente a caleadspmensenada@gmail.com.";
          note.classList.add("field-note--error");
        }
      } finally {
        setTimeout(() => {
          btn.textContent = originalBtnText;
          btn.disabled = false;
        }, 2600);
      }
    });
  }

  /* ---------------------------------------------------------------------
     Año actual en el footer
     --------------------------------------------------------------------- */
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();
})();
