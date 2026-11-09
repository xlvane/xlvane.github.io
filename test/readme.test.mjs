// node --test: the logo and the badges of a README do not reach the site.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { stripReadmeHeader } from '../scripts/lib/readme.mjs';
import { markdownToHtml } from '../scripts/lib/markdown.mjs';

const README = `<img src="https://raw.githubusercontent.com/DevXord/xlvane/main/assets/logo.svg" alt="" width="96" height="96">

# xlvane

[![ci](https://github.com/DevXord/xlvane/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/DevXord/xlvane/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/xlvane?color=3b4fd8&label=npm)](https://www.npmjs.com/package/xlvane)
[![node](https://img.shields.io/badge/node-%E2%89%A522-3b4fd8)](https://nodejs.org/)

> Read, inspect and write XLSX files.

Documentation: [xlvane.github.io](https://xlvane.github.io/).

Text with a [link](https://example.com/) and \`code\`.
`;

test('the logo line and the badge lines are removed, the rest is kept', () => {
  const md = stripReadmeHeader(README);
  assert.equal(md, '# xlvane\n\n> Read, inspect and write XLSX files.\n\nDocumentation: [xlvane.github.io](https://xlvane.github.io/).\n\nText with a [link](https://example.com/) and `code`.\n');
});

test('the page has no image from another server', () => {
  const { html } = markdownToHtml(stripReadmeHeader(README));
  assert.doesNotMatch(html, /<img\b/);
  assert.doesNotMatch(html, /img\.shields\.io|raw\.githubusercontent\.com|badge\.svg/);
  assert.match(html, /Read, inspect and write XLSX files/);
});

test('a README without a header is unchanged', () => {
  const md = '# x\n\nSome text.\n\n[a link](https://example.com/) on its own line\n';
  assert.equal(stripReadmeHeader(md), md);
});
