(function () {
  "use strict";

  const categories = ["Alle", "Power Query", "Regex", "VBA", "Formeln"];
  const categoryShort = { "Power Query": "PQ", Regex: "RX", VBA: "VB", Formeln: "fx" };
  const demos = Array.isArray(window.DEMO_CATALOG) ? window.DEMO_CATALOG : [];
  const catalogMeta = window.CATALOG_META || {};
  const core = window.CatalogCore;
  const state = { category: core.ALL_CATEGORY, query: "", newOnly: false, packageOnly: false };
  let activeDemo = null;

  const grid = document.querySelector("#demo-grid");
  const emptyState = document.querySelector("#empty-state");
  const resultCount = document.querySelector("#result-count");
  const searchInput = document.querySelector("#search-input");
  const newOnlyInput = document.querySelector("#new-only");
  const packageOnlyInput = document.querySelector("#package-only");
  const filterContainer = document.querySelector("#category-filters");
  const cardTemplate = document.querySelector("#demo-card-template");
  const dialog = document.querySelector("#detail-dialog");
  const counts = core.getCounts(demos);

  document.querySelector("#stat-total").textContent = counts.Alle;
  document.querySelector("#stat-new").textContent = counts.Neu;
  document.querySelector("#stat-packages").textContent = catalogMeta.packageDemoCount || demos.filter(demo => demo.package).length;
  document.querySelector("#catalog-version").textContent = catalogMeta.version || "Katalog";

  function makeFilterButtons() {
    categories.forEach(function (category) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "filter-button";
      button.dataset.category = category;
      button.setAttribute("aria-pressed", String(category === state.category));
      button.innerHTML = `<span>${category}</span><span class="filter-count">${counts[category] || 0}</span>`;
      button.addEventListener("click", function () {
        state.category = category;
        render();
      });
      filterContainer.appendChild(button);
    });
  }

  function setText(target, value) {
    target.textContent = value;
  }

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} Byte`;
    if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
    return `${(bytes / (1024 * 1024)).toLocaleString("de-DE", { maximumFractionDigits: 1 })} MB`;
  }

  function openDetails(demo) {
    activeDemo = demo;
    const meta = dialog.querySelector("#detail-meta");
    meta.replaceChildren();

    const category = document.createElement("span");
    category.className = `category-label category-${demo.category.toLowerCase().replace(/\s+/g, "-")}`;
    category.textContent = demo.category;
    meta.appendChild(category);

    if (demo.isNew) {
      const badge = document.createElement("span");
      badge.className = "new-badge dialog-new";
      badge.textContent = "✦ Neu";
      meta.appendChild(badge);
    }

    setText(dialog.querySelector("#detail-title"), demo.title);
    setText(dialog.querySelector("#detail-summary"), demo.summary);
    setText(dialog.querySelector("#detail-id"), `Demo ${demo.id}`);
    setText(dialog.querySelector("#detail-description"), demo.description);
    setText(dialog.querySelector("#detail-code"), demo.codeText);
    setText(dialog.querySelector("#code-label"), demo.category === "Power Query" ? "M-Code" : demo.category === "VBA" ? "VBA" : "Excel-Formel");
    dialog.querySelector("#code-details").open = false;
    dialog.querySelector("#copy-code").textContent = "Code kopieren";

    const sourceLink = dialog.querySelector("#source-link");
    sourceLink.href = demo.sourceUrl;
    sourceLink.setAttribute("aria-label", `${demo.title} auf GitHub öffnen`);

    const packageInfo = dialog.querySelector("#package-info");
    packageInfo.hidden = !demo.package;
    if (demo.package) {
      setText(dialog.querySelector("#package-name"), demo.package.fileName);
      setText(dialog.querySelector("#package-size"), `Beispieldatei · ${formatBytes(demo.package.sizeBytes)}`);
      const downloadLink = dialog.querySelector("#download-link");
      downloadLink.href = demo.package.downloadUrl;
      downloadLink.download = demo.package.fileName;
      downloadLink.setAttribute("aria-label", `${demo.package.fileName} herunterladen`);
    }

    dialog.showModal();
  }

  function createCard(demo) {
    const fragment = cardTemplate.content.cloneNode(true);
    const card = fragment.querySelector(".demo-card");
    const meta = fragment.querySelector(".card-meta");
    const title = fragment.querySelector("h3");
    const summary = fragment.querySelector(".card-summary");
    const badge = fragment.querySelector(".new-badge");
    const fileBadge = fragment.querySelector(".file-badge");
    const tagList = fragment.querySelector(".tag-list");
    const detailsButton = fragment.querySelector(".details-button");

    card.dataset.category = demo.category;
    if (demo.isNew) card.classList.add("is-new");

    const glyph = document.createElement("span");
    glyph.className = `category-glyph glyph-${demo.category.toLowerCase().replace(/\s+/g, "-")}`;
    glyph.textContent = categoryShort[demo.category];

    const id = document.createElement("span");
    id.className = "card-id";
    id.textContent = demo.id;
    meta.append(glyph, id);

    setText(title, demo.title);
    setText(summary, demo.summary);
    badge.hidden = !demo.isNew;
    fileBadge.hidden = !demo.package;

    demo.tags.slice(0, 2).forEach(function (tag) {
      const chip = document.createElement("span");
      chip.textContent = tag;
      tagList.appendChild(chip);
    });

    detailsButton.setAttribute("aria-label", `Details zu ${demo.title}`);
    detailsButton.addEventListener("click", function () { openDetails(demo); });
    card.addEventListener("dblclick", function () { openDetails(demo); });

    return fragment;
  }

  function render() {
    const filtered = core.filterDemos(demos, state);
    grid.replaceChildren(...filtered.map(createCard));

    filterContainer.querySelectorAll(".filter-button").forEach(function (button) {
      const isActive = button.dataset.category === state.category;
      button.setAttribute("aria-pressed", String(isActive));
    });

    emptyState.hidden = filtered.length !== 0;
    grid.hidden = filtered.length === 0;
    resultCount.textContent = filtered.length === 1 ? "1 Demo" : `${filtered.length} Demos`;
  }

  searchInput.addEventListener("input", function () {
    state.query = searchInput.value;
    render();
  });

  newOnlyInput.addEventListener("change", function () {
    state.newOnly = newOnlyInput.checked;
    render();
  });

  packageOnlyInput.addEventListener("change", function () {
    state.packageOnly = packageOnlyInput.checked;
    render();
  });

  document.querySelector("#reset-filters").addEventListener("click", function () {
    state.category = core.ALL_CATEGORY;
    state.query = "";
    state.newOnly = false;
    state.packageOnly = false;
    searchInput.value = "";
    newOnlyInput.checked = false;
    packageOnlyInput.checked = false;
    searchInput.focus();
    render();
  });

  document.querySelector("#close-dialog").addEventListener("click", function () { dialog.close(); });
  document.querySelector("#dialog-done").addEventListener("click", function () { dialog.close(); });
  document.querySelector("#copy-code").addEventListener("click", async function (event) {
    if (!activeDemo) return;
    try {
      await navigator.clipboard.writeText(activeDemo.codeText);
      event.currentTarget.textContent = "Kopiert ✓";
    } catch {
      event.currentTarget.textContent = "Kopieren nicht möglich";
    }
  });
  dialog.addEventListener("click", function (event) {
    const rect = dialog.getBoundingClientRect();
    const outside = event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom;
    if (outside) dialog.close();
  });

  document.addEventListener("keydown", function (event) {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
      event.preventDefault();
      searchInput.focus();
    }
  });

  makeFilterButtons();
  render();
})();
