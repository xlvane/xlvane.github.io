# xlvane.github.io

The website of [xlvane](https://github.com/DevXord/xlvane): a landing page, the documentation of the published version, and a demo that inspects an XLSX file in the browser.

GitHub Pages publishes the `docs/` folder of the `main` branch. `docs/` is generated; edit `src/` and `scripts/` and build again.

```sh
node scripts/build.mjs --version 1.1.0   # downloads xlvane, @xlvane/compat and fflate from npm (integrity checked) and the CHANGELOG of tag v1.1.0
node scripts/check.mjs                   # every internal link and anchor exists, nothing is loaded from other origins
node --test test/*.test.mjs              # tests of the Markdown converter
```

- No dependencies: the generator, the Markdown converter and the tar reader are in `scripts/`.
- The documentation pages are the Markdown files of the packages published on npm, so they match that version.
- The demo uses the browser code of the published package (`docs/vendor/xlvane/`) and fflate's browser build (`docs/vendor/fflate/`); nothing is loaded from other sites, and every page sets a Content-Security-Policy that allows only this site.

`vendor.lock.json` holds the sha512 of every package version the site was built from. A version that is not in it (a new xlvane release, or a newer fflate within xlvane's range) stops the build until it is checked and accepted with `--accept`; a changed sha512 for a known version always stops it.

After each release of xlvane: check its provenance on npm, build with `--version X.Y.Z --accept`, check, commit and push.

License: MIT.
