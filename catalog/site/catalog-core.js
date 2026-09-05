(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.CatalogCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  const ALL_CATEGORY = "Alle";

  function normalize(value) {
    return String(value || "")
      .toLocaleLowerCase("de-DE")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
  }

  function searchableText(demo) {
    return normalize([
      demo.id,
      demo.title,
      demo.category,
      demo.summary,
      demo.description,
      demo.codeText,
      ...(demo.tags || []),
      ...(demo.learnings || [])
    ].join(" "));
  }

  function filterDemos(demos, options) {
    const category = options.category || ALL_CATEGORY;
    const query = normalize(options.query);

    return demos.filter(function (demo) {
      const categoryMatch = category === ALL_CATEGORY || demo.category === category;
      const newMatch = !options.newOnly || demo.isNew === true;
      const packageMatch = !options.packageOnly || Boolean(demo.package);
      const searchMatch = !query || searchableText(demo).includes(query);
      return categoryMatch && newMatch && packageMatch && searchMatch;
    });
  }

  function getCounts(demos) {
    return demos.reduce(function (counts, demo) {
      counts.Alle += 1;
      counts[demo.category] = (counts[demo.category] || 0) + 1;
      if (demo.isNew) counts.Neu += 1;
      return counts;
    }, { Alle: 0, Neu: 0 });
  }

  return { ALL_CATEGORY, normalize, searchableText, filterDemos, getCounts };
});
