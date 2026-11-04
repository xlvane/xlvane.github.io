// The page frame shared by every page of the site: head (with the Content-Security-Policy), header
// with navigation, optional table of contents, and footer.
import { escapeHtml } from './markdown.mjs';

/** The address the site is published at (GitHub Pages of the organization xlvane). */
export const SITE = 'https://xlvane.github.io/';

/**
 * The canonical address of a page from its path in docs/: "index.html" is left out, so a folder has
 * one address ("guide/index.html" is "https://xlvane.github.io/guide/").
 */
export function canonicalUrl(path) {
  if (path.startsWith('/') || path.includes('..')) throw new Error(`not a path inside the site: ${path}`);
  return SITE + path.replace(/(^|\/)index\.html$/, '$1');
}

/** Content-Security-Policy of every page: nothing is loaded from other origins. */
export const CSP = "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'";

const NAV = [
  ['guide/index.html', 'Guide', 'guide'],
  ['demo/index.html', 'Demo', 'demo'],
  ['changelog.html', 'Changelog', 'changelog'],
];

/**
 * A complete page. `path` is its path in docs/ (for the canonical address); `root` is the relative path to the site root ("" or "../"); `active` marks the
 * navigation entry; `toc` is a list of { level, text, id } shown beside the content.
 */
export function page({ title, description, path, root, active = '', body, toc = [], version, scripts = [], wide = false }) {
  const url = canonicalUrl(path);
  const nav = NAV.map(([href, label, key]) => `<a href="${root}${href}"${key === active ? ' aria-current="page"' : ''}>${label}</a>`).join('');
  const contents =
    toc.length > 0
      ? `<nav class="toc" aria-label="On this page"><p class="toc-title">On this page</p><ul>${toc
          .map((h) => `<li class="toc-${String(h.level)}"><a href="#${escapeHtml(h.id)}">${escapeHtml(h.text)}</a></li>`)
          .join('')}</ul></nav>`
      : '';
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="${CSP}">
<title>${escapeHtml(title)}</title>
<meta name="description" content="${escapeHtml(description)}">
<link rel="canonical" href="${escapeHtml(url)}">
<meta property="og:type" content="website">
<meta property="og:site_name" content="xlvane">
<meta property="og:title" content="${escapeHtml(title)}">
<meta property="og:description" content="${escapeHtml(description)}">
<meta property="og:url" content="${escapeHtml(url)}">
<link rel="icon" href="${root}favicon.svg" type="image/svg+xml">
<link rel="stylesheet" href="${root}style.css">
${scripts.map((s) => `<script type="module" src="${s}"></script>`).join('\n')}
</head>
<body>
<a class="skip" href="#main">Skip to content</a>
<header class="site-header">
  <div class="bar">
    <a class="brand" href="${root}index.html"><span class="logo" aria-hidden="true"></span>xlvane</a>
    <span class="version" title="Version on npm">${escapeHtml(version)}</span>
    <nav class="main-nav" aria-label="Main">${nav}<a href="https://github.com/DevXord/xlvane">GitHub</a><a href="https://www.npmjs.com/package/xlvane">npm</a></nav>
  </div>
</header>
<div class="layout${toc.length > 0 ? ' with-toc' : ''}${wide ? ' wide' : ''}">
${contents}
<main id="main">
${body}
</main>
</div>
<footer class="site-footer">
  <p>xlvane ${escapeHtml(version)}, MIT License. Pages built from the package published on npm. <a href="https://github.com/DevXord/xlvane">Source on GitHub</a> · <a href="https://github.com/DevXord/xlvane/blob/main/SECURITY.md">Reporting a vulnerability</a></p>
</footer>
</body>
</html>
`;
}
