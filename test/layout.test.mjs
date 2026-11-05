// node --test: the page frame (canonical address and Open Graph tags).
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { canonicalUrl, page } from '../scripts/lib/layout.mjs';

test('canonical addresses leave out index.html', () => {
  assert.equal(canonicalUrl('index.html'), 'https://xlvane.github.io/');
  assert.equal(canonicalUrl('guide/index.html'), 'https://xlvane.github.io/guide/');
  assert.equal(canonicalUrl('guide/mcp.html'), 'https://xlvane.github.io/guide/mcp.html');
  assert.equal(canonicalUrl('sitemap.xml'), 'https://xlvane.github.io/sitemap.xml');
  assert.throws(() => canonicalUrl('../x.html'));
  assert.throws(() => canonicalUrl('/x.html'));
});

test('every page has its canonical link and Open Graph tags, escaped', () => {
  const html = page({ title: 'A "b" & c', description: 'd <e>', path: 'guide/mcp.html', root: '../', version: '1.0.0', body: '' });
  assert.match(html, /<link rel="canonical" href="https:\/\/xlvane\.github\.io\/guide\/mcp\.html">/);
  assert.match(html, /<meta property="og:url" content="https:\/\/xlvane\.github\.io\/guide\/mcp\.html">/);
  assert.match(html, /<meta property="og:title" content="A &quot;b&quot; &amp; c">/);
  assert.match(html, /<meta property="og:description" content="d &lt;e&gt;">/);
  assert.match(html, /<meta property="og:type" content="website">/);
});

test('a page that moved gives the new page as its canonical address', () => {
  const html = page({ title: 't', description: 'd', path: 'guide/compat.html', canonicalPath: 'guide/exceljs.html', root: '../', version: '1.3.3', body: '' });
  assert.match(html, /<link rel="canonical" href="https:\/\/xlvane\.github\.io\/guide\/exceljs\.html">/);
  assert.match(html, /<meta property="og:url" content="https:\/\/xlvane\.github\.io\/guide\/exceljs\.html">/);
});
