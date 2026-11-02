// The page frame shared by every page of the site: head (with the Content-Security-Policy), header
// with navigation, optional table of contents, and footer.
import { escapeHtml } from './markdown.mjs';

/** Content-Security-Policy of every page: nothing is loaded from other origins. */
export const CSP = "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'";

const NAV = [
  ['guide/index.html', 'Guide', 'guide'],
  ['demo/index.html', 'Demo', 'demo'],
  ['changelog.html', 'Changelog', 'changelog'],
];

/**
 * A complete page. `root` is the relative path to the site root ("" or "../"); `active` marks the
 * navigation entry; `toc` is a list of { level, text, id } shown beside the content.
 */
export function page({ title, description, root, active = '', body, toc = [], version, scripts = [], wide = false }) {
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
