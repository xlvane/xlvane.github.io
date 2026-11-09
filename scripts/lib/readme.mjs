// README files on npm and GitHub start with a logo and a row of badges, all loaded from other
// servers (raw.githubusercontent.com, img.shields.io, github.com). The site loads nothing from
// other origins (Content-Security-Policy, scripts/check.mjs), so the build removes those lines
// before the README becomes a page.

/** A line with the logo, an `<img>` whose source is on another server. */
const LOGO = /^\s*<img\s[^>]*\bsrc="https?:\/\/[^"]*"[^>]*>\s*$/i;
/** A line that holds only badges: linked images such as `[![ci](https://…)](https://…)`. */
const BADGES = /^\s*(?:\[!\[[^\]]*\]\([^)]*\)\]\([^)]*\)\s*)+$/;

/**
 * The README without the logo line and the badge lines, and without the blank lines they leave
 * behind. Other lines, including images and links in the text, are kept.
 */
export function stripReadmeHeader(md) {
  const out = [];
  for (const line of md.split('\n')) {
    if (LOGO.test(line) || BADGES.test(line)) continue;
    if (line.trim() === '' && (out.length === 0 || out[out.length - 1].trim() === '')) continue;
    out.push(line);
  }
  return out.join('\n');
}
