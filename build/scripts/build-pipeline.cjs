#!/usr/bin/env node
/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// Full grunt-free build pipeline for ElixirOffice sdkjs.
//
// Usage (from sdkjs/build/):
//   PRODUCT_VERSION=9.2.1 BUILD_ROOT=/path/to/deploy node scripts/build-pipeline.cjs
//
// Options (env vars):
//   PRODUCT_VERSION   default '0.0.0'
//   BUILD_ROOT        default ../deploy/sdkjs
//   BUILD_NUMBER      default '0'
//   COMPANY_NAME      default 'onlyoffice'
//   SDK_PLATFORM      '' | 'desktop' | 'mobile' — passed through to webpack configs
//   SDK_ADDONS        path.delimiter-separated addon directories
//   SKIP_DEVELOP      set to '1' to skip develop scripts generation
//   SKIP_BUNDLE       set to '1' to skip Phase 1 (deploy-assets + webpack ×4) entirely —
//                      for the fast-iteration workflow that only wants the
//                      develop/sdkjs/{module}/scripts.js manifest regenerated, without
//                      paying for a full bundle build. Implies build-develop runs with
//                      copyStandalone (deploy-assets, which normally supplies the
//                      processed device_scale.js, didn't run).
//
// Phase layout (wall-clock optimised):
//   Phase 1 — parallel: deploy-assets + webpack ×4 (word, cell, slide, visio) — skipped
//             entirely when SKIP_BUNDLE=1.
//             Each webpack config runs 2 compiler configs (min + all chunk) in parallel.
//   Phase 2 — sequential: build-develop (writes develop/sdkjs/{module}/scripts.js)

const { spawn } = require('child_process');
const path      = require('path');
const fs        = require('fs');
const { resolveBuildRoot } = require('../lib/env.cjs');

// This pipeline takes all configuration through env vars — it never reads process.argv.
// A stale caller still passing old Grunt CLI flags (--addon=, --desktop=true, --level=,
// --map, --mobile=true, --beta, --formatting=, --src=) would otherwise have those tokens
// silently ignored: the build exits 0 looking complete while quietly missing an addon or
// a whole platform's files. Fail loudly instead so a stale caller breaks visibly.
// Gated on require.main === module: when this file is require()'d by the test suite
// instead of run directly, process.argv reflects the test runner's own invocation
// (e.g. `node --test test/`), not arguments meant for this script.
if (require.main === module && process.argv.length > 2) {
    process.stderr.write(
        'build-pipeline: this script takes no CLI arguments — configuration is via env vars.\n' +
        `  Got: ${process.argv.slice(2).join(' ')}\n` +
        '  Old Grunt flag -> new env var:\n' +
        '    --addon=X       -> SDK_ADDONS=path/to/X (path.delimiter-separated for multiple)\n' +
        '    --desktop=true  -> SDK_PLATFORM=desktop\n' +
        '    --mobile=true   -> SDK_PLATFORM=mobile\n' +
        '    --map           -> SDK_SOURCE_MAPS=1\n' +
        '    --beta=X        -> BETA=X\n' +
        '    --level=*, --formatting=*, --src=* -> no equivalent (see build/DEVELOPER-GUIDE.md)\n'
    );
    process.exit(1);
}

const BUILD_DIR = path.resolve(__dirname, '..');

const SRC_ROOT = path.resolve(BUILD_DIR, '..');

const DEFAULT_BUILD_ROOT = path.resolve(BUILD_DIR, '..', 'deploy', 'sdkjs');

const BUILD_ROOT = resolveBuildRoot(BUILD_DIR);

// Guard against a BUILD_ROOT that resolves onto the sdkjs checkout, the
// build/ directory itself, or (bar the known default deploy dir) any other
// path inside the checkout — this is deleted wholesale below via fs.rmSync.
//
// Resolved through realpath (where possible) rather than the raw path, so a
// BUILD_ROOT that is or contains a symlink pointing back into the checkout
// can't slip past the path.relative string comparisons below — those compare
// logical paths and don't themselves follow symlinks. Falls back to the
// literal path when it (or an ancestor) doesn't exist yet — a not-yet-created
// BUILD_ROOT is common (it's often created fresh per build) and can't itself
// be a symlink back into the checkout.
function realOrSelf(p) {
    try {
        return fs.realpathSync(p);
    } catch (_) {
        return p;
    }
}

function assertSafeBuildRoot(root) {
    const realRoot = realOrSelf(root);
    const realSrc  = realOrSelf(SRC_ROOT);
    const realBuildDir = realOrSelf(BUILD_DIR);

    const relSourceFromRoot = path.relative(realRoot, realSrc);
    const rootContainsSource =
        relSourceFromRoot === '' ||
        (!relSourceFromRoot.startsWith('..') && !path.isAbsolute(relSourceFromRoot));

    if (realRoot === realSrc || realRoot === realBuildDir || rootContainsSource) {
        throw new Error(`Refusing to clean unsafe BUILD_ROOT: ${root}`);
    }

    const relRootFromSrc = path.relative(realSrc, realRoot);
    const insideSource =
        relRootFromSrc !== '' &&
        !relRootFromSrc.startsWith('..') &&
        !path.isAbsolute(relRootFromSrc);

    if (insideSource && realRoot !== realOrSelf(DEFAULT_BUILD_ROOT)) {
        throw new Error(`Refusing to clean in-tree BUILD_ROOT: ${root}`);
    }
}

assertSafeBuildRoot(BUILD_ROOT);

const PRODUCT_VERSION = process.env.PRODUCT_VERSION || '0.0.0';
const BUILD_NUMBER    = String(process.env.BUILD_NUMBER || process.env.GITHUB_RUN_NUMBER || '0');
const SKIP_DEVELOP    = process.env.SKIP_DEVELOP === '1';
const SKIP_BUNDLE     = process.env.SKIP_BUNDLE === '1';

// Only pass BUILD_ROOT through when the user actually set it — build-develop.cjs
// branches on process.env.BUILD_ROOT's presence to decide between a relative and
// an absolute output path. Injecting our own default here would force it down
// the "user set it" branch with the wrong (relative) path.
const CHILD_ENV = {
    ...process.env,
    PRODUCT_VERSION,
    BUILD_NUMBER,
};

// ---- output helpers (mirrors web-apps/build/scripts/build-pipeline.js) ----

const BOLD  = s => `\x1b[1m${s}\x1b[0m`;
const DIM   = s => `\x1b[2m${s}\x1b[0m`;
const GREEN = s => `\x1b[32m${s}\x1b[0m`;
const RED   = s => `\x1b[31m${s}\x1b[0m`;
const CYAN  = s => `\x1b[36m${s}\x1b[0m`;
const PAD   = 20;

function elapsed(ms) {
    return ms < 1000 ? `${ms}ms` : `${(ms / 1000).toFixed(1)}s`;
}

function banner(msg) {
    process.stdout.write(`\n${BOLD(CYAN('▶ ' + msg))}\n`);
}

// ---- task runner -----------------------------------------------------------

function task(label, cmd, args = [], opts = {}) {
    return { label, cmd, args, opts };
}

function runTask({ label, cmd, args, opts = {} }) {
    let child = null;
    // Set only by our own kill() below, i.e. when phase() is aborting the rest
    // of a batch after a sibling task already failed. A signal from anywhere
    // else (OOM killer, external SIGKILL, …) is a real failure, not an
    // intentional abort, and must not be swallowed as a non-fatal "killed".
    let killedByUs = false;
    const promise = new Promise(resolve => {
        const start = Date.now();
        const paddedLabel = label.padEnd(PAD);
        const stderrBuf = [];

        child = spawn(cmd, args, {
            env:   { ...CHILD_ENV, ...(opts.env || {}) },
            cwd:   opts.cwd || BUILD_DIR,
            stdio: ['ignore', 'pipe', 'pipe'],
        });

        child.stdout.on('data', chunk => {
            for (const line of chunk.toString().split('\n')) {
                if (line.trim()) process.stdout.write(`  ${DIM('[' + label + ']')} ${line}\n`);
            }
        });

        child.stderr.on('data', chunk => { stderrBuf.push(chunk.toString()); });

        child.on('error', err => {
            const ms = Date.now() - start;
            process.stdout.write(`  ${RED('✗')} ${paddedLabel} ${RED('FAILED')} ${DIM(elapsed(ms))}\n`);
            process.stderr.write(`  spawn error: ${err.message}\n`);
            resolve({ label, ms, code: 1 });
        });

        child.on('exit', (code, signal) => {
            const ms = Date.now() - start;
            if (signal && killedByUs) {
                process.stdout.write(`  ${DIM('○')} ${paddedLabel} ${DIM('killed ' + elapsed(ms))}\n`);
                if (stderrBuf.length) process.stderr.write(stderrBuf.join(''));
                resolve({ label, ms, code: -1 });
            } else if (signal) {
                process.stdout.write(`  ${RED('✗')} ${paddedLabel} ${RED('KILLED (' + signal + ')')} ${DIM(elapsed(ms))}\n`);
                if (stderrBuf.length) process.stderr.write(stderrBuf.join(''));
                resolve({ label, ms, code: 1 });
            } else if (code === 0) {
                process.stdout.write(`  ${GREEN('✓')} ${paddedLabel} ${DIM(elapsed(ms))}\n`);
                resolve({ label, ms, code: 0 });
            } else {
                process.stdout.write(`  ${RED('✗')} ${paddedLabel} ${RED('FAILED')} ${DIM(elapsed(ms))}\n`);
                if (stderrBuf.length) process.stderr.write(stderrBuf.join(''));
                resolve({ label, ms, code });
            }
        });
    });
    return { promise, kill: () => { killedByUs = true; child && child.kill('SIGTERM'); }, label };
}

async function phase(title, taskSpecs) {
    const count = taskSpecs.length;
    banner(`${title} — ${count} task${count !== 1 ? 's' : ''}`);

    const running = taskSpecs.map(runTask);
    let aborted = false;

    const results = await Promise.all(
        running.map(t =>
            t.promise.then(r => {
                if (r.code > 0 && !aborted) {
                    aborted = true;
                    running.forEach(o => { try { o.kill(); } catch (_) {} });
                }
                return r;
            })
        )
    );

    const failed = results.filter(r => r.code > 0);
    if (failed.length) {
        process.stderr.write(RED(`\n✗ ${failed.map(r => r.label).join(', ')} failed — aborting\n`));
        process.exit(1);
    }
    return results;
}

// ---- source maps -------------------------------------------------------

// Mirrors the old Gruntfile `copy-maps` task: NODE_ENV=development, or
// SDK_SOURCE_MAPS=1 on a production build, enables devtool:'source-map' in
// webpack.sdk.factory.mjs, which writes sdk-all(-min).js.map next to the
// bundles in the deploy directory. Left there, they're served to end users
// and leak full source paths. Neither var set means devtool:false and no
// .map files, so this is a no-op in that case.
function relocateSourceMaps() {
    const MODULES  = ['word', 'cell', 'slide', 'visio'];
    const mapsRoot = path.join(BUILD_DIR, 'maps');
    let moved = 0;

    for (const name of MODULES) {
        const moduleDir = path.join(BUILD_ROOT, name);
        if (!fs.existsSync(moduleDir)) continue;

        for (const file of fs.readdirSync(moduleDir)) {
            if (!file.endsWith('.map')) continue;
            fs.mkdirSync(mapsRoot, { recursive: true });
            // Flat `{module}-all(-min).js.map` naming — not a `{module}/` subdirectory —
            // matches the legacy Gruntfile.js `copy-maps` layout that
            // build/deserializer/download-maps.js (and whatever uploads these
            // externally) still expects (`${editor}${'-all(-min).js.map'}`).
            const legacyName = name + file.replace(/^sdk/, '');
            const src  = path.join(moduleDir, file);
            const dest = path.join(mapsRoot, legacyName);
            // BUILD_ROOT (moduleDir's ancestor) is very commonly a different mount
            // than this checkout's build/ dir — a bind-mounted volume in a dev
            // container, a separate Docker layer, an external deploy path passed
            // via the BUILD_ROOT env var. plain renameSync throws EXDEV across
            // filesystem boundaries; fall back to copy+unlink in that case.
            try {
                fs.renameSync(src, dest);
            } catch (err) {
                if (err.code !== 'EXDEV') throw err;
                fs.copyFileSync(src, dest);
                fs.unlinkSync(src);
            }
            moved++;

            // The bundle still carries `//# sourceMappingURL=<file>` pointing at
            // the now-relocated map — left in place, browsers/devtools request
            // that path from the deploy dir and 404 since the map no longer
            // lives there. Strip the comment; the map is still on disk under
            // mapsRoot for anyone who needs to attach it manually.
            const bundleFile = path.join(moduleDir, file.slice(0, -'.map'.length));
            if (fs.existsSync(bundleFile)) {
                const bundle = fs.readFileSync(bundleFile, 'utf8');
                const stripped = bundle.replace(/\n?\/\/# sourceMappingURL=.*$/, '');
                if (stripped !== bundle) fs.writeFileSync(bundleFile, stripped);
            }
        }
    }

    if (moved) {
        process.stdout.write(`  ${DIM(`relocated ${moved} source map(s) to ${mapsRoot}`)}\n`);
    }
}

// ---- pipeline --------------------------------------------------------------

const node = process.execPath;
// Resolve via Node's module resolution rather than a hardcoded node_modules/.bin
// path — this stays correct regardless of dependency hoisting (workspaces,
// pnpm, etc.), and a missing/misconfigured webpack fails as a clear
// MODULE_NOT_FOUND instead of a bare ENOENT from spawn().
const wpCli = require.resolve('webpack-cli/bin/cli.js');

const WEBPACK_CONFIGS = [
    'webpack.word.mjs',
    'webpack.cell.mjs',
    'webpack.slide.mjs',
    'webpack.visio.mjs',
];

async function main() {
    const wallStart = Date.now();

    process.stdout.write([
        BOLD('ElixirOffice sdkjs build pipeline'),
        `  BUILD_ROOT        ${BUILD_ROOT}`,
        `  PRODUCT_VERSION   ${PRODUCT_VERSION}`,
        `  BUILD_NUMBER      ${BUILD_NUMBER}`,
        `  SDK_PLATFORM      ${process.env.SDK_PLATFORM || '(default)'}`,
        `  SKIP_DEVELOP      ${SKIP_DEVELOP}`,
        `  SKIP_BUNDLE       ${SKIP_BUNDLE}`,
        '',
    ].join('\n'));

    // Clean deploy directory before building. Skipped under SKIP_BUNDLE — Phase 1
    // (the only phase that writes into BUILD_ROOT) doesn't run, so there's nothing
    // to clean, and wiping it would just discard whatever a prior full build left there.
    if (!SKIP_BUNDLE && fs.existsSync(BUILD_ROOT)) {
        fs.rmSync(BUILD_ROOT, { recursive: true, force: true });
    }

    // Phase 1: all independent work in parallel.
    //   - deploy-assets: copies CSS, fonts, images, themes, native JS (WHITESPACE compiled)
    //   - webpack ×4: each produces sdk-all-min.js + sdk-all.js for its module
    let p1 = [];
    if (!SKIP_BUNDLE) {
        const phase1Tasks = [
            task('deploy-assets', node, ['scripts/deploy-assets.cjs']),
            ...WEBPACK_CONFIGS.map(cfg => {
                const name = cfg.replace('webpack.', '').replace('.mjs', '');
                return task(`webpack:${name}`, node, [wpCli, '--config', cfg]);
            }),
        ];

        p1 = await phase('Phase 1 — parallel', phase1Tasks);

        relocateSourceMaps();
    }

    // Phase 2: develop scripts (fast, sequential is fine).
    let p2 = [];
    if (!SKIP_DEVELOP) {
        p2 = await phase('Phase 2 — develop', [
            // SKIP_STANDALONE=1 only when Phase 1 ran: deploy-assets already deployed
            // a processed device_scale.js there, and build-develop's copyStandalone
            // would overwrite it with an unprocessed raw copy. Under SKIP_BUNDLE,
            // deploy-assets never ran, so copyStandalone needs to run instead.
            task('build-develop', node, ['scripts/build-develop.cjs'], {
                env: SKIP_BUNDLE ? {} : { SKIP_STANDALONE: '1' },
            }),
        ]);
    }

    // Summary
    const all = [...p1, ...p2];
    const wallMs = Date.now() - wallStart;
    const longestLabel = Math.max(...all.map(r => r.label.length));

    process.stdout.write([
        '',
        BOLD('Summary'),
        ...all.map(r => {
            const mark = r.code === 0 ? GREEN('✓') : r.code < 0 ? DIM('○') : RED('✗');
            return `  ${mark} ${r.label.padEnd(longestLabel + 2)} ${DIM(elapsed(r.ms))}`;
        }),
        '',
        `  Wall clock: ${BOLD(elapsed(wallMs))}`,
        '',
    ].join('\n'));
}

if (require.main === module) {
    main().catch(err => {
        process.stderr.write(RED(`\nFatal: ${err.message || err}\n`));
        process.exit(1);
    });
} else {
    // Exposed for unit testing only (e.g. assertSafeBuildRoot) — running this
    // file directly as a script (require.main === module) is still the only
    // way the pipeline itself executes.
    module.exports = { assertSafeBuildRoot, SRC_ROOT, BUILD_DIR, DEFAULT_BUILD_ROOT };
}
