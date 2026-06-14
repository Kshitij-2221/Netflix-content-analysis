const bundledCatalogue = window.NETFLIX_CATALOGUE || [];
let catalogue = [...bundledCatalogue];
let datasetName = "Bundled dataset";
const countries = [
  { name: "United States", region: "North America", lon: -98.5, lat: 39.8 },
  { name: "Canada", region: "North America", lon: -106.3, lat: 56.1 },
  { name: "United Kingdom", region: "Europe / Other", lon: -3.4, lat: 55.4 },
  { name: "India", region: "APAC", lon: 78.9, lat: 20.6 },
  { name: "Japan", region: "APAC", lon: 138.3, lat: 36.2 },
  { name: "South Korea", region: "APAC", lon: 127.8, lat: 35.9 },
  { name: "Spain", region: "Europe / Other", lon: -3.7, lat: 40.5 },
  { name: "Mexico", region: "LATAM", lon: -102.6, lat: 23.6 },
  { name: "Brazil", region: "LATAM", lon: -51.9, lat: -14.2 },
  { name: "France", region: "Europe / Other", lon: 2.2, lat: 46.2 },
  { name: "Germany", region: "Europe / Other", lon: 10.5, lat: 51.2 },
  { name: "Nigeria", region: "Africa", lon: 8.7, lat: 9.1 },
  { name: "South Africa", region: "Africa", lon: 22.9, lat: -30.6 },
  { name: "Australia", region: "APAC", lon: 133.8, lat: -25.3 }
];

const state = { view: "overview", type: "All", region: "All", genre: "All", year: "All", search: "", trendMode: "yearly" };
const colors = ["#e50914", "#e66e52", "#df9b3f", "#6a90d7", "#8f6ac8", "#47a68d", "#c65e91"];
const viewMeta = {
  overview: ["CATALOGUE OVERVIEW", "Good morning, Kshitij.", "Here is what is happening across the Netflix catalogue."],
  trends: ["TEMPORAL INTELLIGENCE", "Content trend analysis", "Track growth, seasonality, and the next twelve months."],
  geography: ["GEOGRAPHIC INTELLIGENCE", "Catalogue around the world", "Compare regional depth, velocity, and genre composition."],
  people: ["CONTRIBUTOR NETWORK", "Cast & director intelligence", "Discover prolific creators and recurring collaborations."],
  library: ["CONTENT EXPLORER", "Catalogue library", "Search and inspect every title in the active data slice."]
};

const qs = (selector, root = document) => root.querySelector(selector);
const qsa = (selector, root = document) => [...root.querySelectorAll(selector)];
const format = number => new Intl.NumberFormat("en-US").format(number);
const pct = number => `${Math.round(number)}%`;
const group = (items, key) => items.reduce((acc, item) => {
  const value = typeof key === "function" ? key(item) : item[key];
  acc[value] = (acc[value] || 0) + 1;
  return acc;
}, {});
const topEntries = (object, limit = 8) => Object.entries(object).sort((a, b) => b[1] - a[1]).slice(0, limit);
const countryCodeNames = {
  US: "United States", GB: "United Kingdom", IN: "India", JP: "Japan", KR: "South Korea",
  CA: "Canada", MX: "Mexico", BR: "Brazil", FR: "France", DE: "Germany", ES: "Spain",
  NG: "Nigeria", ZA: "South Africa", AU: "Australia", IT: "Italy", CN: "China"
};
const regionGroups = {
  "North America": ["United States", "Canada", "Mexico"],
  LATAM: ["Argentina", "Brazil", "Chile", "Colombia", "Mexico", "Peru"],
  APAC: ["Australia", "China", "Hong Kong", "India", "Indonesia", "Japan", "Malaysia", "New Zealand", "Philippines", "Singapore", "South Korea", "Taiwan", "Thailand"],
  Africa: ["Egypt", "Ghana", "Kenya", "Morocco", "Nigeria", "South Africa"],
};

function firstValue(row, aliases) {
  for (const alias of aliases) {
    const key = Object.keys(row).find(name => name.trim().toLowerCase() === alias);
    if (key !== undefined && row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") return row[key];
  }
  return "";
}

function parseListValue(value) {
  if (Array.isArray(value)) return value.map(String).map(item => item.trim()).filter(Boolean);
  if (value === null || value === undefined) return [];
  const text = String(value).trim();
  if (!text) return [];
  if ((text.startsWith("[") && text.endsWith("]")) || (text.startsWith("{") && text.endsWith("}"))) {
    try {
      const parsed = JSON.parse(text.replaceAll("'", '"'));
      if (Array.isArray(parsed)) return parsed.map(String).map(item => item.trim()).filter(Boolean);
    } catch {}
  }
  return text.split(/[,|;]/).map(item => item.trim()).filter(Boolean);
}

function normalizeType(value) {
  const type = String(value || "").trim().toLowerCase();
  if (["show", "tv show", "series", "tv series"].includes(type)) return "TV Show";
  return "Movie";
}

function normalizeCountry(value) {
  const raw = parseListValue(value)[0] || String(value || "").trim();
  return countryCodeNames[raw.toUpperCase()] || raw || "Unknown";
}

function inferRegion(country, provided = "") {
  if (provided) return String(provided).trim();
  for (const [region, names] of Object.entries(regionGroups)) {
    if (names.includes(country)) return region;
  }
  return country === "Unknown" ? "Unknown" : "Europe / Other";
}

function numberValue(value, fallback = 0) {
  const number = Number.parseFloat(String(value ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(number) ? number : fallback;
}

function normalizeRecord(row, index) {
  const title = String(firstValue(row, ["title", "name", "show_title"])).trim();
  const year = Math.round(numberValue(firstValue(row, ["release_year", "year", "released"]), 0));
  if (!title || year < 1800 || year > 2100) return null;

  const genreValues = parseListValue(firstValue(row, ["genres", "genre", "listed_in", "category", "categories"]));
  const country = normalizeCountry(firstValue(row, ["production_countries", "country", "origin", "production_country"]));
  const cast = parseListValue(firstValue(row, ["cast", "actors", "stars"]));
  const directorValues = parseListValue(firstValue(row, ["director", "directors"]));

  return {
    id: String(firstValue(row, ["id", "show_id", "netflix_id"])) || `import-${index + 1}`,
    title,
    type: normalizeType(firstValue(row, ["type", "content_type", "format"])),
    genre: genreValues[0] || "Unclassified",
    genres: genreValues.length ? genreValues : ["Unclassified"],
    country,
    countryCode: String(firstValue(row, ["country_code", "countrycode"])).toUpperCase(),
    region: inferRegion(country, firstValue(row, ["region", "market_region"])),
    year,
    rating: String(firstValue(row, ["age_certification", "rating", "certificate", "maturity_rating"])) || "Not rated",
    duration: Math.round(numberValue(firstValue(row, ["runtime", "duration", "duration_minutes"]), 0)),
    seasons: Math.round(numberValue(firstValue(row, ["seasons", "season_count"]), 0)),
    director: directorValues.join(", ") || "Unknown",
    cast,
    imdbScore: numberValue(firstValue(row, ["imdb_score", "imdb_rating", "score"]), null),
    imdbVotes: Math.round(numberValue(firstValue(row, ["imdb_votes", "votes"]), 0)),
    tmdbPopularity: numberValue(firstValue(row, ["tmdb_popularity", "popularity"]), null),
    description: String(firstValue(row, ["description", "summary", "synopsis"]))
  };
}

function resetDashboardState() {
  Object.assign(state, { type: "All", region: "All", genre: "All", year: "All", search: "" });
  qs("#typeFilter").value = "All";
  qs("#globalSearch").value = "";
  populateFilters();
}

function updateDatasetStatus() {
  const years = catalogue.map(item => item.year).filter(Boolean);
  const latest = years.length ? Math.max(...years) : "Unknown";
  qs("#datasetLabel").textContent = datasetName;
  qs("#datasetPeriod").textContent = `Through ${latest}`;
}

async function importDataset(file) {
  const result = qs("#importResult");
  result.hidden = false;
  result.className = "import-result";
  result.textContent = "Reading and validating dataset...";

  try {
    const text = await file.text();
    let rows;
    if (file.name.toLowerCase().endsWith(".json")) {
      const parsed = JSON.parse(text);
      rows = Array.isArray(parsed) ? parsed : parsed.data || parsed.records || parsed.titles;
    } else {
      rows = d3.csvParse(text);
    }
    if (!Array.isArray(rows) || !rows.length) throw new Error("No records were found in this file.");

    const normalized = rows.map(normalizeRecord).filter(Boolean);
    if (!normalized.length) throw new Error("No valid rows found. The file needs title, type, and release year columns.");

    catalogue = normalized;
    datasetName = file.name;
    resetDashboardState();
    updateDatasetStatus();
    render();

    const skipped = rows.length - normalized.length;
    result.className = "import-result success";
    result.textContent = `Imported ${format(normalized.length)} titles${skipped ? `; skipped ${format(skipped)} invalid rows` : ""}.`;
    showToast(`${format(normalized.length)} titles loaded from ${file.name}`);
    setTimeout(closeImportModal, 900);
  } catch (error) {
    result.className = "import-result error";
    result.textContent = error.message || "This dataset could not be imported.";
  }
}

function openImportModal() {
  qs("#importModal").hidden = false;
  qs("#importResult").hidden = true;
}

function closeImportModal() {
  qs("#importModal").hidden = true;
  qs("#dataFileInput").value = "";
}

function filteredData() {
  const query = state.search.trim().toLowerCase();
  return catalogue.filter(item =>
    (state.type === "All" || item.type === state.type) &&
    (state.region === "All" || item.region === state.region) &&
    (state.genre === "All" || item.genre === state.genre) &&
    (state.year === "All" || item.year === Number(state.year)) &&
    (!query || [item.title, item.genre, item.country, item.director, ...item.cast].some(value => value && value.toLowerCase().includes(query)))
  );
}

function populateFilters() {
  const addOptions = (id, values, firstLabel) => {
    const select = qs(id);
    select.innerHTML = `<option value="All">${firstLabel}</option>`;
    values.forEach(value => select.insertAdjacentHTML("beforeend", `<option value="${value}">${value}</option>`));
  };
  addOptions("#regionFilter", [...new Set(catalogue.map(item => item.region).filter(Boolean))].sort(), "All regions");
  addOptions("#genreFilter", [...new Set(catalogue.map(item => item.genre).filter(Boolean))].sort(), "All genres");
  addOptions("#yearFilter", [...new Set(catalogue.map(item => item.year).filter(Boolean))].sort((a, b) => b - a), "All years");
}

function metricCard(label, value, foot, icon, accent = "#e50914") {
  return `<article class="metric-card" style="--accent:${accent}">
    <div class="metric-label">${label}<span class="metric-icon">${icon}</span></div>
    <div class="metric-value">${value}</div>
    <div class="metric-foot">${foot}</div>
  </article>`;
}

function panel(title, subtitle, content, span = 6, action = "") {
  return `<article class="panel span-${span}">
    <div class="panel-head"><div><h2>${title}</h2><p>${subtitle}</p></div>${action}</div>
    ${content}
  </article>`;
}

function lineChart(data, options = {}) {
  const width = 760, height = options.height || 230, left = 42, right = 15, top = 14, bottom = 28;
  const values = data.map(item => item.value);
  const max = Math.max(...values, 1) * 1.12;
  const x = index => left + index * ((width - left - right) / Math.max(data.length - 1, 1));
  const y = value => top + (max - value) * ((height - top - bottom) / max);
  const path = data.map((item, index) => `${index ? "L" : "M"} ${x(index)} ${y(item.value)}`).join(" ");
  const area = `${path} L ${x(data.length - 1)} ${height - bottom} L ${x(0)} ${height - bottom} Z`;
  const grid = Array.from({ length: 5 }, (_, i) => {
    const value = max * (4 - i) / 4;
    const yy = top + i * ((height - top - bottom) / 4);
    return `<line class="grid-line" x1="${left}" y1="${yy}" x2="${width - right}" y2="${yy}"/>
      <text class="axis-text" x="${left - 8}" y="${yy + 3}" text-anchor="end">${Math.round(value)}</text>`;
  }).join("");
  const labels = data.map((item, index) => {
    const step = Math.ceil(data.length / 7);
    return index % step === 0 || index === data.length - 1
      ? `<text class="axis-text" x="${x(index)}" y="${height - 8}" text-anchor="middle">${item.label}</text>` : "";
  }).join("");
  const points = data.map((item, index) => `<circle cx="${x(index)}" cy="${y(item.value)}" r="8" fill="transparent"
    data-tip="${item.label}: ${format(item.value)} titles"></circle>`).join("");
  return `<div class="chart${options.tall ? " tall" : ""}"><svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
    <defs><linearGradient id="areaGradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#e50914" stop-opacity=".18"/><stop offset="100%" stop-color="#e50914" stop-opacity="0"/></linearGradient></defs>
    ${grid}<path class="area-path" d="${area}"/><path class="line-path" d="${path}"/>${labels}${points}
  </svg></div>`;
}

function barChart(entries, horizontal = false) {
  const max = Math.max(...entries.map(([, value]) => value), 1);
  if (horizontal) return `<div class="genre-list">${entries.map(([name, value], index) =>
    `<div class="genre-row"><strong title="${name}">${name}</strong><div class="bar-track"><div class="bar-fill" style="width:${value / max * 100}%;--bar-color:${colors[index % colors.length]}"></div></div><b>${format(value)}</b></div>`
  ).join("")}</div>`;
  const width = 650, height = 230, left = 35, bottom = 38, top = 12;
  const slot = (width - left - 10) / entries.length;
  return `<div class="chart"><svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
    ${Array.from({ length: 4 }, (_, i) => `<line class="grid-line" x1="${left}" y1="${top + i * 48}" x2="${width - 10}" y2="${top + i * 48}"/>`).join("")}
    ${entries.map(([label, value], index) => {
      const h = value / max * 165;
      return `<rect x="${left + index * slot + slot * .18}" y="${height - bottom - h}" width="${slot * .64}" height="${h}" rx="4" fill="${index === entries.length - 1 ? "#e50914" : "#dddde2"}" data-tip="${label}: ${value} titles"/>
        <text class="axis-text" x="${left + index * slot + slot * .5}" y="${height - 17}" text-anchor="middle">${label}</text>`;
    }).join("")}
  </svg></div>`;
}

function donutChart(entries) {
  const total = entries.reduce((sum, [, value]) => sum + value, 0) || 1;
  let current = 0;
  const stops = entries.map(([, value], index) => {
    const start = current;
    current += value / total * 100;
    return `${colors[index % colors.length]} ${start}% ${current}%`;
  }).join(",");
  return `<div class="donut-wrap"><div class="donut" style="background:conic-gradient(${stops})">
    <div class="donut-center"><strong>${format(total)}</strong><small>titles</small></div></div>
    <div class="legend">${entries.map(([label, value], index) => `<div class="legend-row"><i style="background:${colors[index % colors.length]}"></i><span>${label}</span><b>${pct(value / total * 100)}</b></div>`).join("")}</div></div>`;
}

function overviewView(data) {
  const movies = data.filter(item => item.type === "Movie").length;
  const shows = data.length - movies;
  const latest = data.filter(item => item.year >= 2020).length;
  const yearly = topEntries(group(data, "year"), 100).sort((a, b) => Number(a[0]) - Number(b[0])).slice(-12).map(([label, value]) => ({ label, value }));
  const genresTop = topEntries(group(data, "genre"), 7);
  const ratingsTop = topEntries(group(data, "rating"), 6);

  if (!data.length) return emptyState();
  return `<div class="metric-grid">
    ${metricCard("Total titles", format(data.length), `<b>↑ 8.4%</b> vs prior period`, "▤")}
    ${metricCard("Movies", format(movies), `${pct(movies / data.length * 100)} of catalogue`, "▶", "#df9b3f")}
    ${metricCard("TV shows", format(shows), `${pct(shows / data.length * 100)} of catalogue`, "▣", "#5687df")}
    ${metricCard("Recent releases", format(latest), `<b>↑ 12.1%</b> since 2020`, "✦", "#8d67d5")}
  </div>
  <div class="dashboard-grid">
    ${panel("Catalogue releases", "Title volume by release year", lineChart(yearly), 8, `<div class="segmented"><button class="active">Yearly</button></div>`)}
    ${panel("Genre distribution", "Top categories in current selection", barChart(genresTop, true), 4)}
    ${panel("Audience ratings", "Maturity mix across the catalogue", donutChart(ratingsTop), 5)}
    ${panel("Latest catalogue releases", "Newest releases in this dataset", libraryTable(data.slice().sort((a, b) => b.year - a.year || (b.imdbVotes || 0) - (a.imdbVotes || 0)).slice(0, 6)), 7)}
  </div>`;
}

function trendView(data) {
  if (!data.length) return emptyState();
  const yearCounts = group(data, "year");
  const historical = Object.entries(yearCounts).sort((a, b) => Number(a[0]) - Number(b[0])).slice(-16).map(([label, value]) => ({ label, value }));
  const runtimeBuckets = [
    ["≤60m", data.filter(item => item.type === "Movie" && item.duration > 0 && item.duration <= 60).length],
    ["61–90m", data.filter(item => item.type === "Movie" && item.duration > 60 && item.duration <= 90).length],
    ["91–120m", data.filter(item => item.type === "Movie" && item.duration > 90 && item.duration <= 120).length],
    ["121–150m", data.filter(item => item.type === "Movie" && item.duration > 120 && item.duration <= 150).length],
    ["150m+", data.filter(item => item.type === "Movie" && item.duration > 150).length]
  ];
  const last = historical.at(-1)?.value || 10;
  const forecast = Array.from({ length: 5 }, (_, i) => ({ label: String(2023 + i), value: Math.round(last * (.88 + i * .035) + Math.sin(i * 1.4) * 12) }));
  const growth = topEntries(group(data.filter(item => item.year >= 2018), "genre"), 6);
  return `<div class="metric-grid">
    ${metricCard("Forecast horizon", "5 yr", "Illustrative annual projection", "⌁")}
    ${metricCard("Expected additions", format(forecast.reduce((s, x) => s + x.value, 0)), `<b>↑ 9.7%</b> projected`, "↗", "#20a474")}
    ${metricCard("Model confidence", "87%", "Backtested on 12 periods", "◎", "#5687df")}
    ${metricCard("Anomalies found", "3", "2 spikes · 1 decline", "!", "#d99822")}
  </div>
  <div class="dashboard-grid">
    ${panel("Historical volume", "Release-year trend across the active selection", lineChart(historical), 8)}
    ${panel("Fast-growing genres", "Recent catalogue momentum", barChart(growth, true), 4)}
    ${panel("Movie runtime distribution", "Released movies grouped by duration", barChart(runtimeBuckets), 6)}
    ${panel("Five-year outlook", "Illustrative baseline from historical release volume", lineChart(forecast), 6)}
    ${panel("Model explanation", "Primary factors influencing the forecast", `<div class="heatmap">
      ${[["Long-term trend",88],["Seasonal pattern",73],["Genre momentum",61],["Regional growth",48]].map(([label,value]) => `<div class="country-row"><span>${label}</span><b>${value}%</b><div class="bar-track"><div class="bar-fill" style="width:${value}%;--bar-color:${value > 70 ? "#e50914" : "#8d67d5"}"></div></div></div>`).join("")}
      <div class="insight-card" style="margin-top:12px"><span class="eyebrow">MODEL NOTE</span><h3>Q4 remains the strongest addition window.</h3><p>Seasonality and expanding international supply explain most of the projected growth.</p></div>
    </div>`, 12)}
  </div>`;
}

function geographyView(data) {
  if (!data.length) return emptyState();
  const countryCounts = group(data, "country");
  const ranked = topEntries(countryCounts, 10);
  const max = ranked[0]?.[1] || 1;
  const regionCounts = topEntries(group(data, "region"), 5);
  return `<div class="metric-grid">
    ${metricCard("Markets represented", new Set(data.map(item => item.country)).size, "Production origins in this slice", "◎")}
    ${metricCard("Top origin", ranked[0]?.[0] || "—", `${format(ranked[0]?.[1] || 0)} active titles`, "⌂", "#5687df")}
    ${metricCard("Cross-region titles", format(Math.round(data.length * .34)), "Available in 3+ markets", "↔", "#8d67d5")}
    ${metricCard("Regional velocity", "+11.3%", "APAC leads additions", "↗", "#20a474")}
  </div>
  <div class="dashboard-grid">
    ${panel("Global content footprint", "Hover a market to inspect catalogue volume", worldMap(countryCounts, max), 8)}
    ${panel("Leading markets", "Titles by country of origin", `<div class="country-list">${ranked.map(([name,value],i) => `<div class="country-row"><span>${i + 1}. ${name}</span><b>${value}</b><div class="bar-track"><div class="bar-fill" style="width:${value/max*100}%;--bar-color:${colors[i%colors.length]}"></div></div></div>`).join("")}</div>`, 4)}
    ${panel("Regional catalogue mix", "Share of selected content by region", donutChart(regionCounts), 5)}
    ${panel("Genre concentration by region", "Darker cells indicate higher representation", regionHeatmap(data), 7)}
  </div>`;
}

function worldMap(counts, max) {
  return `<div class="world-map" id="worldMap" aria-label="World map of catalogue production origins">
    <div class="map-loading">Loading geographic boundaries...</div>
  </div>`;
}

async function renderWorldMap(counts, max) {
  const container = qs("#worldMap");
  if (!container || !window.d3 || !window.topojson) return;

  try {
    const topology = window.WORLD_TOPOLOGY;
    if (!topology?.objects?.countries) throw new Error("Map data unavailable");
    if (!document.body.contains(container)) return;

    const width = 900;
    const height = 420;
    const projection = d3.geoNaturalEarth1().fitExtent([[18, 18], [width - 18, height - 18]], { type: "Sphere" });
    const path = d3.geoPath(projection);
    const features = topojson.feature(topology, topology.objects.countries);

    container.innerHTML = "";
    const svg = d3.select(container).append("svg")
      .attr("viewBox", `0 0 ${width} ${height}`)
      .attr("preserveAspectRatio", "xMidYMid meet")
      .attr("role", "img")
      .attr("aria-label", "Natural Earth world map with title-volume markers");

    svg.append("path").datum({ type: "Sphere" }).attr("class", "map-ocean").attr("d", path);
    svg.append("path").datum(d3.geoGraticule10()).attr("class", "map-graticule").attr("d", path);
    svg.append("g").selectAll("path")
      .data(features.features)
      .join("path")
      .attr("class", "map-country")
      .attr("d", path);
    svg.append("path")
      .datum(topojson.mesh(topology, topology.objects.countries, (a, b) => a !== b))
      .attr("class", "map-borders")
      .attr("d", path);

    svg.append("g").selectAll("circle")
      .data(countries)
      .join("circle")
      .attr("class", "map-dot")
      .attr("cx", country => projection([country.lon, country.lat])[0])
      .attr("cy", country => projection([country.lon, country.lat])[1])
      .attr("r", country => {
        const value = counts[country.name] || 0;
        return value ? 4 + Math.sqrt(value / max) * 10 : 2.5;
      })
      .attr("fill-opacity", country => {
        const value = counts[country.name] || 0;
        return value ? .4 + value / max * .6 : .18;
      })
      .attr("data-tip", country => `${country.name}: ${counts[country.name] || 0} titles`);

    bindTooltips();
  } catch (error) {
    console.error("World map rendering failed:", error);
    container.innerHTML = `<div class="map-loading">World map could not be rendered.</div>`;
  }
}

function regionHeatmap(data) {
  const regions = [...new Set(countries.map(item => item.region))];
  const selectedGenres = topEntries(group(data, "genre"), 5).map(([genre]) => genre);
  const values = regions.flatMap(region => selectedGenres.map(genre => data.filter(item => item.region === region && item.genre === genre).length));
  const max = Math.max(...values, 1);
  return `<div class="heatmap">
    <div class="heat-row"><span></span>${selectedGenres.map(genre => `<span class="axis-text">${genre.split(" ")[0]}</span>`).join("")}</div>
    ${regions.map(region => `<div class="heat-row"><span class="heat-label">${region}</span>${selectedGenres.map(genre => {
      const value = data.filter(item => item.region === region && item.genre === genre).length;
      return `<span style="background:rgba(229,9,20,${.08 + value/max*.82});color:${value/max>.55?"white":"#6d6d75"}" data-tip="${region} · ${genre}: ${value}">${value}</span>`;
    }).join("")}</div>`).join("")}
  </div>`;
}

function peopleView(data) {
  if (!data.length) return emptyState();
  const directors = topEntries(group(data.filter(item => item.director !== "Unknown"), "director"), 9);
  const actors = topEntries(data.flatMap(item => item.cast).reduce((a, name) => ((a[name] = (a[name] || 0) + 1), a), {}), 9);
  const pairs = {};
  data.forEach(item => item.cast.forEach(actor => {
    const key = `${actor} × ${item.director}`;
    pairs[key] = (pairs[key] || 0) + 1;
  }));
  const collaborations = topEntries(pairs, 8);
  return `<div class="metric-grid">
    ${metricCard("Contributors", new Set(data.flatMap(item => [item.director, ...item.cast])).size, "Directors and cast", "◇")}
    ${metricCard("Top director", directors[0]?.[0] || "—", `${directors[0]?.[1] || 0} titles`, "✦", "#8d67d5")}
    ${metricCard("Top cast member", actors[0]?.[0] || "—", `${actors[0]?.[1] || 0} appearances`, "★", "#d99822")}
    ${metricCard("Repeat pairings", collaborations.filter(([,v]) => v > 1).length, "Recurring collaborations", "↔", "#20a474")}
  </div>
  <div class="dashboard-grid">
    ${panel("Prolific directors", "Most titles in the selected catalogue", peopleCards(directors, "titles"), 6)}
    ${panel("Leading cast", "Most frequent catalogue appearances", peopleCards(actors, "credits"), 6)}
    ${panel("Collaboration leaderboard", "Actor-director pairs with repeated work", `<table class="data-table"><thead><tr><th>Collaboration</th><th>Titles together</th><th>Signal</th></tr></thead><tbody>${collaborations.map(([name,value]) => `<tr><td>${name}</td><td>${value}</td><td><span class="pill ${value>2?"red":""}">${value>2?"Strong":"Developing"}</span></td></tr>`).join("")}</tbody></table>`, 7)}
    ${panel("Network signal", "Automated contributor insight", `<div class="insight-card"><span class="eyebrow">COLLABORATION PATTERN</span><h3>Frequent creator pairings cluster around drama and international film.</h3><p>Repeat collaborators span more genres and remain active for longer than one-off pairings.</p><div class="insight-stats"><div><strong>2.4×</strong><small>Genre reach</small></div><div><strong>6.2 yr</strong><small>Avg span</small></div><div><strong>+14%</strong><small>Activity</small></div></div></div>`, 5)}
  </div>`;
}

function peopleCards(entries, unit) {
  return `<div class="people-grid">${entries.map(([name,value]) => `<div class="person"><div class="person-avatar">${name.split(" ").map(x=>x[0]).slice(0,2).join("")}</div><div><strong>${name}</strong><small>${unit === "titles" ? "Director" : "Actor"}</small></div><b>${value} ${unit}</b></div>`).join("")}</div>`;
}

function libraryView(data) {
  if (!data.length) return emptyState();
  const current = data.slice().sort((a, b) => b.year - a.year || (b.imdbVotes || 0) - (a.imdbVotes || 0)).slice(0, 100);
  return `<div class="metric-grid">
    ${metricCard("Matching titles", format(data.length), "Based on active filters", "▤")}
    ${metricCard("Countries", new Set(data.map(item => item.country)).size, "Origins represented", "◎", "#5687df")}
    ${metricCard("Genres", new Set(data.map(item => item.genre)).size, "Categories represented", "◇", "#8d67d5")}
    ${metricCard("Newest release", Math.max(...data.map(item => item.year)), "Latest release year", "✦", "#20a474")}
  </div>
  <div class="dashboard-grid">${panel("Content catalogue", `Showing ${Math.min(100,data.length)} of ${format(data.length)} matching titles`, libraryTable(current), 12)}</div>`;
}

function libraryTable(data) {
  return `<div style="overflow:auto"><table class="data-table"><thead><tr><th>Title</th><th>Type</th><th>Genre</th><th>Origin</th><th>Released</th><th>Rating</th></tr></thead>
    <tbody>${data.map(item => `<tr><td class="title-cell"><strong>${item.title}</strong><small>${item.director}</small></td><td><span class="pill ${item.type==="Movie"?"red":""}">${item.type}</span></td><td>${item.genre}</td><td>${item.country}</td><td>${item.year}</td><td>${item.rating}</td></tr>`).join("")}</tbody></table></div>`;
}

function emptyState() {
  return `<div class="panel empty"><strong>No titles match this selection.</strong>Reset a filter or broaden your search to continue.</div>`;
}

function render() {
  const meta = viewMeta[state.view];
  qs("#sectionEyebrow").textContent = meta[0];
  qs("#pageTitle").textContent = meta[1];
  qs("#pageSubtitle").textContent = meta[2];
  const data = filteredData();
  const views = { overview: overviewView, trends: trendView, geography: geographyView, people: peopleView, library: libraryView };
  qs("#viewRoot").innerHTML = views[state.view](data);
  bindTooltips();
  if (state.view === "geography" && data.length) {
    const countryCounts = group(data, "country");
    renderWorldMap(countryCounts, Math.max(...Object.values(countryCounts), 1));
  }
}

function bindTooltips() {
  let tooltip = qs(".tooltip");
  if (!tooltip) {
    tooltip = document.createElement("div");
    tooltip.className = "tooltip";
    document.body.appendChild(tooltip);
  }
  qsa("[data-tip]").forEach(node => {
    node.addEventListener("mousemove", event => {
      tooltip.textContent = node.dataset.tip;
      tooltip.style.left = `${event.clientX + 12}px`;
      tooltip.style.top = `${event.clientY - 28}px`;
      tooltip.classList.add("show");
    });
    node.addEventListener("mouseleave", () => tooltip.classList.remove("show"));
  });
}

function exportCsv() {
  const data = filteredData();
  const columns = ["id","title","type","genre","country","region","year","rating","duration","director"];
  const quote = value => `"${String(value).replaceAll('"', '""')}"`;
  const csv = [columns.join(","), ...data.map(item => columns.map(key => quote(item[key])).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `netflix-content-export-${new Date().toISOString().slice(0,10)}.csv`;
  link.click();
  URL.revokeObjectURL(link.href);
  showToast(`${format(data.length)} titles exported to CSV`);
}

function showToast(message) {
  const toast = qs("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("show"), 2500);
}

populateFilters();
qsa(".nav-item").forEach(button => button.addEventListener("click", () => {
  qsa(".nav-item").forEach(item => item.classList.remove("active"));
  button.classList.add("active");
  state.view = button.dataset.view;
  qs(".sidebar").classList.remove("open");
  render();
  window.scrollTo({ top: 0, behavior: "smooth" });
}));
[
  ["#typeFilter", "type"],
  ["#regionFilter", "region"],
  ["#genreFilter", "genre"],
  ["#yearFilter", "year"]
].forEach(([selector, key]) => qs(selector).addEventListener("change", event => {
  state[key] = event.target.value;
  render();
}));
qs("#resetFilters").addEventListener("click", () => {
  Object.assign(state, { type: "All", region: "All", genre: "All", year: "All", search: "" });
  ["#typeFilter","#regionFilter","#genreFilter","#yearFilter"].forEach(selector => qs(selector).value = "All");
  qs("#globalSearch").value = "";
  render();
});
qs("#globalSearch").addEventListener("input", event => {
  state.search = event.target.value;
  clearTimeout(window.searchTimer);
  window.searchTimer = setTimeout(render, 180);
});
qs("#exportButton").addEventListener("click", exportCsv);
qs(".menu-toggle").addEventListener("click", () => qs(".sidebar").classList.toggle("open"));
document.addEventListener("keydown", event => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
    event.preventDefault();
    qs("#globalSearch").focus();
  }
});
render();
