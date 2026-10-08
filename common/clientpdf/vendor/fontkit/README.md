# fontkit (vendored)

- Package: `@pdf-lib/fontkit` 1.1.1 (fork of `fontkit` by Devon Govett), file `dist/fontkit.umd.min.js`, unmodified.
- Source: https://github.com/Hopding/fontkit (npm `@pdf-lib/fontkit@1.1.1`).
- Licence: MIT (see `LICENSE`).
- Used by `../../eo-pdf.js` (client-side PDF export) to parse TrueType / OpenType (CFF) / TrueType
  collection fonts, read glyph advances and write glyph subsets. Loaded on first PDF export only.

The bundle also contains these packages, all under permissive licences:
`restructure` (MIT), `unicode-properties` and `unicode-trie` (MIT), `brotli` (decoder, MIT),
`tiny-inflate` (MIT), `pako` (MIT and zlib), `dfa` (MIT), `clone` (MIT), `base64-js` (MIT),
`buffer` (MIT), `ieee754` (BSD-3-Clause).

Update: replace the file with the `dist/fontkit.umd.min.js` of a new version, keep this README and
the licence in step, run the client PDF tests (`tests/client-pdf/` in the product repository).
