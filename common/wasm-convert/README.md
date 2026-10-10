# Browser office conversion

Load the shared EOWasmFonts service and api.js. Call
`EOWasmConvert.convert(bytes, from, to, options)`; supported formats are
docx/odt/xlsx/ods/pptx/odp and doct/xlst/pptt editor binaries.

Each call creates and terminates one worker/module. Input and resource buffers
remain owned by the caller. Options accept signal, onProgress, resources,
fontNames and a same-origin runtime directory. Failures have a code and native
failures include nativeStatus. Returned Uint8Array has resources and stats.
Structural conversions preserve font metadata; fontNames requests physical
font files through EOWasmFonts when a conversion needs them. The editor loads
its screen fonts separately. No adapter or full font-directory scan is used.

Before an SDK asset build, run:
`python3 common/wasm-convert/install-runtime.py RELEASE_DIR`.
The installer validates the pinned manifest and all six artifact checksums,
rejects the PDF-only module, and installs the immutable version directory.
Generated JS/wasm and compression sidecars stay outside Git.

The example local-assets route must serve all local editor and worker files
with its CSP, including the module's correct application/wasm MIME and
Brotli/gzip sidecars. A deployment also needs the SDK wasm-convert asset glob
and the Controller module in each editor bundle. Those shared build hooks
require the lead's approval during parallel work.
