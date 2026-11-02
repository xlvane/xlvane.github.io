// Builds the site into docs/ (published by GitHub Pages from the main branch, /docs folder) from
// the packages published on npm and the changelog of the release tag:
//   node scripts/build.mjs --version 1.1.0
// The documentation pages are the Markdown files of the xlvane and @xlvane/compat packages, so the
// site shows exactly the documentation of that version. The demo uses the package's own browser
// code (dist/) and fflate's browser build, copied into docs/vendor/.
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, posix } from 'node:path';
import { fileURLToPath } from 'node:url';
import { page } from './lib/layout.mjs';
import { escapeHtml, markdownToHtml } from './lib/markdown.mjs';
import { fetchPackage } from './lib/npm.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const OUT = join(ROOT, 'docs');
const REPO = 'https://github.com/DevXord/xlvane';

const arg = (name) => {
  const i = process.argv.indexOf(`--${name}`);
  return i === -1 ? undefined : process.argv[i + 1];
};
const version = arg('version');
if (!version || !/^\d+\.\d+\.\d+$/.test(version)) {
  console.error('usage: node scripts/build.mjs --version X.Y.Z');
  process.exit(2);
}

/** Repository paths of the documents, and the page each becomes. */
const PAGES = [
  { repo: 'packages/xlvane/README.md', pkg: 'xlvane', file: 'README.md', out: 'guide/index.html', title: 'Guide' },
  { repo: 'packages/compat/README.md', pkg: '@xlvane/compat', file: 'README.md', out: 'guide/compat.html', title: '@xlvane/compat' },
  { repo: 'docs/errors.md', pkg: 'xlvane', file: 'docs/errors.md', out: 'guide/errors.html', title: 'Errors and warnings' },
  { repo: 'docs/inspect-json.md', pkg: 'xlvane', file: 'docs/inspect-json.md', out: 'guide/inspect-json.html', title: 'inspect() report' },
  { repo: 'docs/cli-json.md', pkg: 'xlvane', file: 'docs/cli-json.md', out: 'guide/cli-json.html', title: 'Command line JSON' },
  { repo: 'docs/exceljs-compat.md', pkg: 'xlvane', file: 'docs/exceljs-compat.md', out: 'guide/exceljs-compat.html', title: 'exceljs compatibility' },
  { repo: 'skills/xlvane/SKILL.md', pkg: 'xlvane', file: 'skills/xlvane/SKILL.md', out: 'guide/agent-skill.html', title: 'Agent Skill' },
  { repo: 'CHANGELOG.md', pkg: null, file: 'CHANGELOG.md', out: 'changelog.html', title: 'Changelog' },
];
const PAGE_BY_REPO = new Map(PAGES.map((p) => [p.repo, p.out]));
PAGE_BY_REPO.set('README.md', 'index.html');

const [core, compat] = await Promise.all([fetchPackage('xlvane', version), fetchPackage('@xlvane/compat', version)]);
const fflateVersion = JSON.parse(new TextDecoder().decode(core.files.get('package.json'))).dependencies.fflate.replace(/^\^/, '');
const fflatePkg = await fetchPackage('fflate', fflateVersion);
const changelogResponse = await fetch(`https://raw.githubusercontent.com/DevXord/xlvane/v${version}/CHANGELOG.md`);
if (!changelogResponse.ok) throw new Error(`CHANGELOG.md of tag v${version}: HTTP ${String(changelogResponse.status)}`);
const changelog = await changelogResponse.text();

const text = (files, path) => {
  const data = files.get(path);
  if (!data) throw new Error(`${path} is missing from the package.`);
  return new TextDecoder().decode(data);
};
const source = (p) => (p.pkg === null ? changelog : text(p.pkg === 'xlvane' ? core.files : compat.files, p.file));

/** A link of a document, mapped to a page of the site (relative to `from`) or to GitHub. */
function linker(repoPath, from) {
  const rel = (target) => posix.relative(posix.dirname(from), target) || posix.basename(target);
  return (url) => {
    if (url.startsWith('#') || /^(https?:|mailto:)/i.test(url)) {
      const gh = /^https:\/\/github\.com\/DevXord\/xlvane\/blob\/main\/([^#?]+)(#.*)?$/.exec(url);
      if (gh && PAGE_BY_REPO.has(gh[1])) return rel(PAGE_BY_REPO.get(gh[1])) + (gh[2] ?? '');
      if (gh && /^docs\/cli-json\/[\w.-]+\.json$/.test(gh[1])) return rel(`guide/${gh[1].slice(5)}`);
      return url;
    }
    const [path, hash = ''] = url.split('#');
    const target = posix.normalize(posix.join(posix.dirname(repoPath), path));
    if (PAGE_BY_REPO.has(target)) return rel(PAGE_BY_REPO.get(target)) + (hash ? `#${hash}` : '');
    if (/^docs\/cli-json\/[\w.-]+\.json$/.test(target)) return rel(`guide/${target.slice(5)}`);
    if (target.startsWith('..')) return null;
    return `${REPO}/blob/v${version}/${target}${hash ? `#${hash}` : ''}`;
  };
}

const write = (path, content) => {
  const full = join(OUT, path);
  mkdirSync(dirname(full), { recursive: true });
  writeFileSync(full, content);
};

rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
write('.nojekyll', '');
cpSync(join(ROOT, 'src/style.css'), join(OUT, 'style.css'));
cpSync(join(ROOT, 'src/favicon.svg'), join(OUT, 'favicon.svg'));

// Documentation pages.
const guideNav = PAGES.filter((p) => p.out.startsWith('guide/'));
for (const p of PAGES) {
  const root = p.out.includes('/') ? '../' : '';
  const { html, headings } = markdownToHtml(source(p), { link: linker(p.repo, p.out) });
  const sub = p.out.startsWith('guide/')
    ? `<nav class="guide-nav" aria-label="Guide">${guideNav.map((g) => `<a href="${posix.relative('guide', g.out)}"${g === p ? ' aria-current="page"' : ''}>${escapeHtml(g.title)}</a>`).join('')}</nav>`
    : '';
  const sourceLink = `<p class="source">Source: <a href="${REPO}/blob/v${version}/${p.repo}">${escapeHtml(p.repo)}</a> (version ${escapeHtml(version)})</p>`;
  write(
    p.out,
    page({
      title: `${p.title} - xlvane`,
      description: `${p.title} of xlvane ${version}, an XLSX library.`,
      root,
      active: p.out.startsWith('guide/') ? 'guide' : 'changelog',
      version,
      toc: headings.filter((h) => h.level === 2 || h.level === 3),
      body: `${sub}<article class="doc">${html}${sourceLink}</article>`,
    }),
  );
}
for (const [path, data] of core.files) if (/^docs\/cli-json\/[\w.-]+\.schema\.json$/.test(path)) write(`guide/${path.slice(5)}`, data);

// Landing page, with the quick start, the command line and the tested environments taken from the README.
const readme = text(core.files, 'README.md');
const section = (heading) => {
  const start = readme.indexOf(`\n## ${heading}\n`);
  if (start === -1) throw new Error(`README section "${heading}" not found.`);
  const end = readme.indexOf('\n## ', start + 4);
  return readme.slice(start, end === -1 ? undefined : end);
};
const firstCode = (md) => {
  const m = /```\w*\n([\s\S]*?)\n```/.exec(md);
  if (!m) throw new Error('code block not found');
  return m[1];
};
const firstParagraph = (md) => md.split('\n\n').map((s) => s.trim()).find((s) => s !== '' && !s.startsWith('#') && !s.startsWith('```')) ?? '';
const fill = (template, values) => template.replace(/\{\{(\w+)\}\}/g, (_, k) => {
  if (!(k in values)) throw new Error(`unknown placeholder ${k}`);
  return values[k];
});
const landing = fill(readFileSync(join(ROOT, 'src/index.html'), 'utf8'), {
  version: escapeHtml(version),
  quickstart: escapeHtml(firstCode(section('Quick start'))),
  cli: escapeHtml(firstCode(section('Command line in 5 minutes'))),
  environments: markdownToHtml(firstParagraph(section('Tested environments')), { link: linker('packages/xlvane/README.md', 'index.html') }).html,
});
write('index.html', page({ title: 'xlvane - read, inspect and write XLSX files', description: 'xlvane reads, inspects and writes XLSX files and treats input files as untrusted. MIT license, one runtime dependency.', root: '', version, body: landing }));

// Demo: the page, its script, and the browser code of xlvane and fflate.
write('demo/index.html', page({ title: 'Demo - xlvane', description: 'Inspect an XLSX file in the browser with xlvane; the file does not leave the computer.', root: '../', active: 'demo', version, wide: true, scripts: ['demo.js'], body: readFileSync(join(ROOT, 'src/demo/index.html'), 'utf8') }));
cpSync(join(ROOT, 'src/demo/demo.js'), join(OUT, 'demo/demo.js'));
const needed = new Set();
const queue = ['dist/index.js'];
while (queue.length > 0) {
  const path = queue.shift();
  if (needed.has(path)) continue;
  needed.add(path);
  for (const m of text(core.files, path).matchAll(/from "(\.\/[\w.-]+\.js)"/g)) queue.push(posix.join('dist', m[1]));
}
for (const path of needed) {
  // The bare specifier "fflate" points to the copy beside it (no import map, so the CSP needs no inline script).
  const js = text(core.files, path).replace(/from "fflate"/g, 'from "../fflate/browser.js"');
  if (/from "(?!\.\.?\/)/.test(js)) throw new Error(`${path} imports a module the browser cannot load.`);
  write(`vendor/xlvane/${posix.basename(path)}`, js);
}
const fflateEntry = JSON.parse(text(fflatePkg.files, 'package.json')).exports['.'].import.default;
write('vendor/fflate/browser.js', fflatePkg.files.get(fflateEntry.replace(/^\.\//, '')));
write('vendor/xlvane/LICENSE', core.files.get('LICENSE'));
write('vendor/fflate/LICENSE', fflatePkg.files.get('LICENSE'));
write(
  'vendor/README.txt',
  `xlvane ${version} (${core.integrity}): dist/*.js used by index.js, MIT license.\nfflate ${fflateVersion} (${fflatePkg.integrity}): ${fflateEntry}, MIT license.\nCopied from the packages published on npm by scripts/build.mjs.\n`,
);

if (!existsSync(join(OUT, 'index.html'))) throw new Error('index.html was not written');
console.log(`docs/: xlvane ${version}, @xlvane/compat ${version}, fflate ${fflateVersion}; ${String(PAGES.length + 2)} pages.`);
