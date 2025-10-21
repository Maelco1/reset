
/**
 * planning-enhancements.js
 * - Computes a uniform square cell size based on the widest label (up to 6 chars)
 * - Propagates per-column color to cells via --slot-color
 * - Inserts visible column index badges above headers
 * - Re-runs when the planning DOM changes
 */
(function () {
  const TABLES_ROOT_SELECTOR = "#planning-tables"; // admin.html container
  const DAY_COL_CLASS = "planning-day-col";

  function $(sel, root = document) {
    return root.querySelector(sel);
  }
  function $all(sel, root = document) {
    return Array.from(root.querySelectorAll(sel));
  }

  function setRootCellSize(px) {
    const size = Math.max(24, Math.min(88, Math.round(px))); // sane bounds
    document.documentElement.style.setProperty("--cell-size", size + "px");
  }

  function measureCellSize(root) {
    // Base guess
    let S = 40;

    // Find all labels that render inside cells
    const labels = $all(".planning-slot-title, .planning-assignment-tag", root);
    if (!labels.length) {
      setRootCellSize(S);
      return S;
    }

    // Create a measuring element
    const measurer = document.createElement("span");
    measurer.style.position = "absolute";
    measurer.style.visibility = "hidden";
    measurer.style.whiteSpace = "nowrap";
    document.body.appendChild(measurer);

    let maxWidth = 0;
    let maxHeight = 0;

    for (const el of labels) {
      const text = el.textContent || "";
      // Limit to 6 visible chars (as per requirement)
      const t = text.slice(0, 6);
      // We mirror the same font used by cells
      const style = getComputedStyle(el);
      // We'll try sizes from small to larger to get approx width at current --cell-size
      measurer.style.fontFamily = style.fontFamily;
      measurer.style.fontWeight = "700";
      // We'll grow S until no label overflows when font-size = S/3.6
      // So first, compute each label's width per unit: width at 1px font-size
      measurer.style.fontSize = "16px";
      measurer.textContent = t;
      const base = measurer.getBoundingClientRect();
      const perPx = base.width / 16; // width per 1px font-size

      // For 6 chars typical latin, perPx ~ 0.6 * chars = 3.6, but we measure it
      // Width at font-size = S/3.6 is: width = perPx * (S/3.6)
      // For no overflow, need width <= S - 4 (some padding)
      // Solve perPx * (S/3.6) <= S - 4  => S >= 4 / (1 - perPx/3.6)
      // When perPx ~ chars*0.6 = up to 3.6, denominator can be small; we'll fallback to iterative.
      maxWidth = Math.max(maxWidth, base.width);
      maxHeight = Math.max(maxHeight, base.height);
    }

    // Iteratively find S that satisfies: for all labels, rendered width <= S - pad
    const PAD = 6;
    function fits(S_try) {
      const fs = S_try / 3.6; // font-size rule in CSS
      measurer.style.fontSize = fs + "px";
      for (const el of labels) {
        const text = (el.textContent || "").slice(0, 6);
        measurer.textContent = text;
        const rect = measurer.getBoundingClientRect();
        if (rect.width > (S_try - PAD)) return false;
      }
      return true;
    }

    S = Math.max(32, parseInt(getComputedStyle(document.documentElement).getPropertyValue("--cell-size")) || 40);

    // Grow until fits (cap to 88px for sanity)
    while (!fits(S) && S < 88) {
      S += 2;
    }

    // Ensure 80% height usability: fs ~= 0.8*S but capped by 1-line fit, already handled
    setRootCellSize(S);
    document.body.removeChild(measurer);
    return S;
  }

  function propagateColumnColors(root) {
    // For each planning table, read each header cell's --slot-color or data-color,
    // and set that CSS variable on all body cells in the same column.
    const tables = $all(".planning-table", root);
    for (const table of tables) {
      const headerRow = table.tHead ? table.tHead.rows[0] : null;
      if (!headerRow) continue;

      const headers = Array.from(headerRow.cells);
      headers.forEach((th, index) => {
        if (th.classList.contains(DAY_COL_CLASS)) return;
        // Read color from custom property or data attribute
        const computed = getComputedStyle(th).getPropertyValue("--slot-color").trim();
        const dataColor = th.getAttribute("data-color") || "";
        const inlineColor = (th.style && th.style.getPropertyValue("--slot-color") || "").trim();
        const color = inlineColor || dataColor || computed || "";

        if (color) {
          // Apply to header for visual coherence
          th.style.setProperty("--slot-color", color);
          // Apply to all body cells in same column
          for (const row of table.tBodies[0].rows) {
            const td = row.cells[index];
            if (td) td.style.setProperty("--slot-color", color);
          }
        }
      });
    }
  }

  function injectColumnNumbers(root) {
    const tables = $all(".planning-table", root);
    for (const table of tables) {
      const headerRow = table.tHead ? table.tHead.rows[0] : null;
      if (!headerRow) continue;
      let visibleIndex = 0;
      const headers = Array.from(headerRow.cells);
      headers.forEach((th) => {
        if (th.classList.contains(DAY_COL_CLASS)) return;
        visibleIndex += 1;
        // Remove previous badge if any
        const prev = th.querySelector(".planning-col-index");
        if (prev) prev.remove();
        const badge = document.createElement("span");
        badge.className = "planning-col-index";
        badge.textContent = "#" + String(visibleIndex);
        th.style.position = "relative";
        th.prepend(badge);
      });
    }
  }

  function enhance(root = document) {
    propagateColumnColors(root);
    injectColumnNumbers(root);
    measureCellSize(root);
  }

  function setupObserver() {
    const host = $(TABLES_ROOT_SELECTOR);
    if (!host) return;
    const obs = new MutationObserver(() => {
      // Debounce microtask
      queueMicrotask(() => enhance(host));
    });
    obs.observe(host, { childList: true, subtree: true });
    // First run
    enhance(host);
    // Recompute on window resize
    window.addEventListener("resize", () => enhance(host), { passive: true });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", setupObserver);
  } else {
    setupObserver();
  }
})();
