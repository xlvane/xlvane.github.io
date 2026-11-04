// Checks the built site in docs/: every relative link and source points to a file that exists
// (and to an id that exists for #anchors), no page loads a script, style, image or font from another
// origin, every page carries the Content-Security-Policy and its canonical address, and robots.txt
// and sitemap.xml list the pages. Exits with 1 and a list on failure.
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { CSP, SITE, canonicalUrl } from './lib/layout.mjs';

const OUT = fileURLToPath(new URL('../docs/', import.meta.url));

function files(dir) {
  return readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    return statSync(path).isDirectory() ? files(path) : [path];
  });
}

// &amp; last, so that "&amp;lt;" becomes "&lt;" and not "<".
const decode = (s) => s.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const pages = files(OUT).filter((f) => f.endsWith('.html'));
const ids = new Map(pages.map((p) => [p, new Set([...readFileSync(p, 'utf8').matchAll(/\sid="([^"]*)"/g)].map((m) => decode(m[1])))]));
const problems = [];
let external = 0;
for (const page of pages) {
  const html = readFileSync(page, 'utf8');
  const where = relative(OUT, page);
  if (!html.includes(`content="${CSP}"`)) problems.push(`${where}: no Content-Security-Policy`);
  for (const m of html.matchAll(/<(script|link|img|iframe|source)\b[^>]*\s(?:src|href)="([^"]*)"/g)) {
    // A canonical link is not loaded; it must be the page's own address on the site.
    if (m[1] === 'link' && /\srel="canonical"/.test(m[0])) continue;
    if (/^(https?:)?\/\//i.test(m[2])) problems.push(`${where}: <${m[1]}> loads ${m[2]} from another origin`);
  }
  for (const m of html.matchAll(/\s(href|src)="([^"]*)"/g)) {
    const url = decode(m[2]);
    if (/^(https?:|mailto:)/i.test(url)) {
      external++;
      continue;
    }
    if (/^[a-z]+:/i.test(url)) {
      problems.push(`${where}: link with scheme ${url}`);
      continue;
    }
    const [rawPath, rawHash] = url.split('#');
    const path = decodeURIComponent(rawPath);
    const hash = rawHash === undefined ? undefined : decodeURIComponent(rawHash);
    const target = path === '' ? page : join(dirname(page), path);
    if (!existsSync(target)) {
      problems.push(`${where}: ${url} does not exist`);
      continue;
    }
    if (hash !== undefined && target.endsWith('.html') && !ids.get(target)?.has(hash)) problems.push(`${where}: ${url} has no id "${hash}"`);
  }
}
// robots.txt names the sitemap; the sitemap lists exactly the pages outside vendor/ and the
// canonical address of each page is its address in the sitemap.
const robots = existsSync(join(OUT, 'robots.txt')) ? readFileSync(join(OUT, 'robots.txt'), 'utf8') : '';
if (!robots.includes(`Sitemap: ${SITE}sitemap.xml`)) problems.push('robots.txt: missing or without the sitemap');
const sitemap = existsSync(join(OUT, 'sitemap.xml')) ? readFileSync(join(OUT, 'sitemap.xml'), 'utf8') : '';
const listed = [...sitemap.matchAll(/<loc>([^<]*)<\/loc>/g)].map((m) => decode(m[1])).sort();
const expected = pages.filter((p) => !relative(OUT, p).startsWith('vendor')).map((p) => canonicalUrl(relative(OUT, p).split(sep).join('/'))).sort();
if (JSON.stringify(listed) !== JSON.stringify(expected)) problems.push(`sitemap.xml: lists ${JSON.stringify(listed)}, expected ${JSON.stringify(expected)}`);
for (const page of pages) {
  const where = relative(OUT, page).split(sep).join('/');
  if (where.startsWith('vendor')) continue;
  const canonical = /<link rel="canonical" href="([^"]*)">/.exec(readFileSync(page, 'utf8'))?.[1];
  if (canonical === undefined || decode(canonical) !== canonicalUrl(where)) problems.push(`${where}: canonical ${String(canonical)} is not ${canonicalUrl(where)}`);
}
if (problems.length > 0) {
  console.error(problems.join('\n'));
  process.exit(1);
}
console.log(`${String(pages.length)} pages checked, ${String(external)} links to other sites not followed.`);
