/**
 * Sitemap import: Search Console list + direct URL, index recursion, gzip, reconciliation.
 */

function fetchSitemapUrls_(sourceUrl) {
  const collected = [];
  crawlSitemap_(sourceUrl, 0, collected, {});
  return collected;
}

function crawlSitemap_(url, depth, collected, seen) {
  if (depth > API_LIMITS.SITEMAP_MAX_DEPTH) return;
  const key = String(url).trim();
  if (!key || seen[key]) return;
  seen[key] = true;

  let xmlText;
  try {
    xmlText = fetchSitemapXml_(key);
  } catch (err) {
    throw new Error('Failed to fetch sitemap ' + key + ': ' + err.message);
  }

  let doc;
  try {
    doc = XmlService.parse(xmlText);
  } catch (err) {
    throw new Error('Malformed XML in sitemap ' + key + ': ' + err.message);
  }

  const root = doc.getRootElement();
  const local = root.getName();
  if (local === 'sitemapindex') {
    const children = childrenNamed_(root, 'sitemap');
    children.forEach(function (node) {
      const loc = textOf_(childNamed_(node, 'loc'));
      if (loc) crawlSitemap_(loc, depth + 1, collected, seen);
    });
    return;
  }
  if (local === 'urlset') {
    const urls = childrenNamed_(root, 'url');
    urls.forEach(function (node) {
      const loc = textOf_(childNamed_(node, 'loc'));
      if (!loc) return;
      collected.push({
        url: loc.trim(),
        lastmod: textOf_(childNamed_(node, 'lastmod'))
      });
    });
    return;
  }
  throw new Error('Sitemap ' + key + ' is not a urlset or sitemapindex (root <' + local + '>). Import aborted.');
}

function fetchSitemapXml_(url) {
  const resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'Accept-Encoding': 'gzip, deflate' }
  });
  const code = resp.getResponseCode();
  if (code >= 400) {
    throw new Error('HTTP ' + code);
  }
  const blob = resp.getBlob();
  const lower = url.toLowerCase();
  const contentType = String(blob.getContentType() || '').toLowerCase();
  const looksGzip = /\.gz($|\?)/.test(lower) ||
    contentType.indexOf('gzip') !== -1 ||
    contentType.indexOf('application/x-gzip') !== -1;
  if (looksGzip) {
    try {
      return Utilities.ungzip(blob).getDataAsString('UTF-8');
    } catch (err) {
      const asText = blob.getDataAsString('UTF-8');
      if (asText && asText.indexOf('<') !== -1) return asText;
      throw new Error('gzip decompress failed: ' + err.message);
    }
  }
  return blob.getDataAsString('UTF-8');
}

function childrenNamed_(element, name) {
  const ns = element.getNamespace();
  if (ns) {
    const byNs = element.getChildren(name, ns);
    if (byNs && byNs.length) return byNs;
  }
  const out = [];
  element.getChildren().forEach(function (child) {
    if (child.getName() === name) out.push(child);
  });
  return out;
}

function childNamed_(element, name) {
  if (!element) return null;
  const kids = childrenNamed_(element, name);
  return kids.length ? kids[0] : null;
}

function textOf_(element) {
  return element ? String(element.getText() || '').trim() : '';
}

function lastmodPriority_(lastmod) {
  if (!lastmod) return 3;
  const d = new Date(lastmod);
  if (isNaN(d.getTime())) return 3;
  const days = (Date.now() - d.getTime()) / 86400000;
  if (days <= 7) return 1;
  if (days <= 30) return 2;
  return 3;
}

function deriveSection_(url) {
  return firstPathSegment_(url);
}

function startSitemapImportFromGsc_(feedPath) {
  const property = getPropertyUrl_();
  if (!property) throw new Error('Set a property in Setup first.');
  const path = feedPath || '';
  const url = path.indexOf('http') === 0 ? path : guessSitemapUrl_(property, path);
  return loadImportSheet_(fetchSitemapUrls_(url), url);
}

function startSitemapImportFromUrl_(url) {
  if (!url) throw new Error('Paste a sitemap URL.');
  return loadImportSheet_(fetchSitemapUrls_(String(url).trim()), url);
}

function guessSitemapUrl_(property, path) {
  if (path && /^https?:/i.test(path)) return path;
  if (property.indexOf('sc-domain:') === 0) {
    const domain = property.substring('sc-domain:'.length);
    if (path) {
      if (path.indexOf('http') === 0) return path;
      return path.charAt(0) === '/' ? 'https://' + domain + path : path;
    }
    return 'https://' + domain + '/sitemap.xml';
  }
  if (path) {
    if (path.indexOf('http') === 0) return path;
    try {
      const base = property.replace(/\/$/, '');
      return path.charAt(0) === '/' ? (base.replace(/^(https?:\/\/[^\/]+).*/, '$1') + path) : path;
    } catch (e) {
      return path;
    }
  }
  return property.replace(/\/$/, '') + '/sitemap.xml';
}

function loadImportSheet_(entries, source) {
  const sheet = getSheet_(SHEETS.IMPORT);
  clearDataRows_(sheet);
  const property = getPropertyUrl_();
  const rows = entries
    .filter(function (e) { return e.url; })
    .filter(function (e, i, arr) {
      return arr.findIndex(function (x) { return x.url === e.url; }) === i;
    })
    .filter(function (e) { return !property || urlBelongsToProperty_(e.url, property); })
    .map(function (e) {
      return [e.url, e.lastmod || '', deriveSection_(e.url), true];
    });
  if (rows.length) {
    appendRows_(sheet, rows);
    sheet.getRange(2, 4, rows.length, 1).insertCheckboxes();
  }
  const sections = {};
  rows.forEach(function (r) {
    sections[r[2]] = (sections[r[2]] || 0) + 1;
  });
  return {
    source: source,
    found: entries.length,
    inProperty: rows.length,
    dropped: entries.length - rows.length,
    sections: Object.keys(sections).sort().map(function (k) {
      return { section: k, count: sections[k] };
    })
  };
}

function filterImport_(includeRegex, excludeRegex) {
  const sheet = getSheet_(SHEETS.IMPORT);
  const last = sheet.getLastRow();
  if (last < 2) return { kept: 0 };
  let includeRe = null;
  let excludeRe = null;
  try {
    if (includeRegex) includeRe = new RegExp(includeRegex);
    if (excludeRegex) excludeRe = new RegExp(excludeRegex);
  } catch (err) {
    throw new Error('Invalid regular expression: ' + err.message);
  }
  const range = sheet.getRange(2, 1, last - 1, IMPORT_HEADERS.length);
  const values = range.getValues();
  values.forEach(function (row) {
    const url = String(row[0] || '');
    let include = true;
    if (includeRe && !includeRe.test(url)) include = false;
    if (excludeRe && excludeRe.test(url)) include = false;
    row[3] = include;
  });
  range.setValues(values);
  const kept = values.filter(function (r) { return asBool_(r[3]); }).length;
  return { kept: kept, total: values.length };
}

function applySectionMap_(mapping) {
  const sheet = getSheet_(SHEETS.IMPORT);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const map = mapping || {};
  const range = sheet.getRange(2, 1, last - 1, 3);
  const values = range.getValues();
  values.forEach(function (row) {
    const derived = String(row[2] || '');
    if (map[derived] !== undefined) row[2] = map[derived];
  });
  range.setValues(values);
}

function previewReconciliation_(defaults) {
  defaults = defaults || {};
  const imported = readIncludedImport_();
  const existing = urlMapByUrl_();
  const existingSet = {};
  Object.keys(existing).forEach(function (u) { existingSet[u] = true; });
  const importSet = {};
  imported.forEach(function (e) { importSet[e.url] = true; });

  let neu = 0;
  imported.forEach(function (e) { if (!existingSet[e.url]) neu++; });
  let same = 0;
  imported.forEach(function (e) { if (existingSet[e.url]) same++; });
  let orphaned = 0;
  Object.keys(existing).forEach(function (u) { if (!importSet[u]) orphaned++; });

  return {
    newCount: neu,
    existingCount: same,
    orphanedCount: orphaned,
    defaults: {
      frequency: defaults.frequency || FREQUENCY.WEEKLY,
      tags: defaults.tags || '',
      active: defaults.active !== false
    }
  };
}

function commitSitemapImport_(defaults) {
  defaults = defaults || {};
  const frequency = defaults.frequency || FREQUENCY.WEEKLY;
  const tags = defaults.tags || '';
  const active = defaults.active !== false;
  compactUrlSheetIfNeeded_();
  const imported = readIncludedImport_();
  const existing = urlMapByUrl_();
  const importSet = {};
  imported.forEach(function (e) { importSet[e.url] = true; });

  const toAdd = [];
  imported.forEach(function (e) {
    if (existing[e.url]) return;
    toAdd.push({
      url: e.url,
      active: active,
      frequency: frequency,
      section: e.section,
      tags: tags,
      priority: lastmodPriority_(e.lastmod)
    });
  });
  if (toAdd.length) addUrlRows_(toAdd);

  const urlSheet = getSheet_(SHEETS.URLS);
  const last = lastUrlRow_(urlSheet);
  if (last >= 2) {
    const range = urlSheet.getRange(2, 1, last - 1, URL_HEADERS.length);
    const values = range.getValues();
    values.forEach(function (row) {
      const url = String(row[0] || '').trim();
      if (!url) return;
      row[URL_COL.ORPHANED - 1] = !importSet[url];
    });
    range.setValues(values);
    urlSheet.getRange(2, URL_COL.ORPHANED, last - 1, 1).insertCheckboxes();
    urlSheet.getRange(2, URL_COL.ACTIVE, last - 1, 1).insertCheckboxes();
  }

  clearDataRows_(getSheet_(SHEETS.IMPORT));
  refreshReservation_();
  return {
    added: toAdd.length,
    orphaned: Object.keys(existing).filter(function (u) { return !importSet[u]; }).length,
    existing: imported.length - toAdd.length
  };
}

function readIncludedImport_() {
  const sheet = getSheet_(SHEETS.IMPORT);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, IMPORT_HEADERS.length).getValues();
  return values.filter(function (r) { return r[0] && asBool_(r[3]); }).map(function (r) {
    return { url: String(r[0]).trim(), lastmod: String(r[1] || ''), section: String(r[2] || '') };
  });
}

function listGscSitemapsForUi_() {
  return listSubmittedSitemaps_(getPropertyUrl_());
}
