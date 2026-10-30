// SPDX-License-Identifier: AGPL-3.0-only
(function (window) {
    "use strict";

    // Ephemeral request coordination only. Callers must establish live layout
    // eligibility and perform actual cell/table page mapping and invalidation.
    const stores = new WeakMap();

    class CellWidowQualificationError extends Error {
        constructor(reason) {
            super("Cell widow replay qualification failed: " + reason);
            this.name = "CellWidowQualificationError";
        }
    }

    function object(value) {
        return value !== null && (typeof value === "object" || typeof value === "function");
    }

    function fail(state, reason) {
        if (state) {
            state.failed = true;
            for (const request of state.pending)
                state.records.get(request).status = "Invalid";
            state.pending.clear();
            state.pendingOwner = new WeakMap();
        }
        throw new CellWidowQualificationError(reason);
    }

    function live(state) {
        if (state.failed)
            fail(state, "the current generation has already failed");
    }

    function integer(state, value, name) {
        if (!Number.isSafeInteger(value) || value < 0)
            fail(state, "invalid " + name);
    }

    function spend(state) {
        live(state);
        if (state.workUsed >= state.maxSteps)
            fail(state, "document-pass work budget exhausted");
        state.workUsed++;
    }

    function frame(state, token) {
        live(state);
        const value = object(token) ? state.tokens.get(token) : null;
        if (!value || !value.active || value.generation !== state.generation)
            fail(state, "inactive or stale owner frame");
        return value;
    }

    function pending(state, request) {
        live(state);
        const value = object(request) ? state.records.get(request) : null;
        if (!value || request.generation !== state.generation || !state.pending.has(request))
            fail(state, "unknown, stale or terminal request");
        return value;
    }

    class CCellWidowReplayProtocol {
        constructor(options) {
            if (!options || !Number.isSafeInteger(options.generation) || options.generation < 0
                || !Number.isSafeInteger(options.maxRequests) || options.maxRequests < 1
                || !Number.isSafeInteger(options.maxDepth) || options.maxDepth < 0
                || !Number.isSafeInteger(options.maxSteps) || options.maxSteps < 1)
                fail(null, "explicit finite generation and operational bounds are required");
            stores.set(this, {
                generation: options.generation, maxRequests: options.maxRequests, maxDepth: options.maxDepth,
                maxSteps: options.maxSteps, workUsed: 0,
                used: 0, nextId: 1, failed: false, tokens: new WeakMap(),
                records: new WeakMap(), pending: new Set(), pendingOwner: new WeakMap(), seen: new WeakMap(),
                pageHighWater: new WeakMap()
            });
        }

        begin(generation) {
            const state = stores.get(this);
            integer(state, generation, "generation");
            if (generation < state.generation)
                fail(state, "generation moved backward");
            if (generation === state.generation) {
                live(state);
                return;
            }
            for (const request of state.pending)
                state.records.get(request).status = "Cancelled";
            state.pending.clear();
            state.pendingOwner = new WeakMap();
            state.seen = new WeakMap();
            state.pageHighWater = new WeakMap();
            state.generation = generation;
            state.used = 0;
            state.workUsed = 0;
            state.failed = false;
        }

        open(owner, admission) {
            const state = stores.get(this);
            live(state);
            if (!object(owner))
                fail(state, "owner identity must be an object");
            // Unknown work remains charged. The controller may admit only a
            // first forward cell page with no pending or cached replay.
            let productive = false;
            if (admission !== undefined) {
                if (!object(admission) || Array.isArray(admission) || typeof admission.retry !== "boolean")
                    fail(state, "invalid productive-work admission");
                integer(state, admission.page, "productive cell page");
                const previous = state.pageHighWater.get(owner);
                const forward = previous === undefined || admission.page > previous;
                productive = forward && !admission.retry && !state.pendingOwner.has(owner);
                if (forward)
                    state.pageHighWater.set(owner, admission.page);
            }
            if (!productive)
                spend(state);
            const token = Object.freeze({});
            state.tokens.set(token, {owner: owner, generation: state.generation, active: true,
                retry: !productive, contents: productive ? new WeakSet() : null});
            return token;
        }

        step(token, content) {
            const state = stores.get(this), value = frame(state, token);
            if (value.retry || state.pendingOwner.has(value.owner)) {
                value.retry = true;
                spend(state);
            } else if (!object(content) || value.contents.has(content)) {
                spend(state);
            } else {
                value.contents.add(content);
            }
        }

        close(token) {
            const state = stores.get(this);
            const value = object(token) ? state.tokens.get(token) : null;
            if (!value)
                fail(state, "unknown owner frame");
            // Cleanup remains possible after failure or a generation change.
            value.active = false;
        }

        request(token, options) {
            const state = stores.get(this), ownerFrame = frame(state, token), owner = ownerFrame.owner;
            if (!options || !object(options.paragraph))
                fail(state, "paragraph identity must be an object");
            for (const name of ["cutLine", "cellPage", "rowIndex", "tablePage"])
                integer(state, options[name], name);
            const signature = options.signature === undefined ? [] : options.signature;
            if (!Array.isArray(signature) || signature.length > 16 || !signature.every(Number.isFinite))
                fail(state, "invalid geometry signature");
            const ancestors = options.ancestors === undefined ? [] : options.ancestors;
            if (!Array.isArray(ancestors) || ancestors.length > state.maxDepth
                || !ancestors.every(object) || new Set([owner, ...ancestors]).size !== ancestors.length + 1)
                fail(state, "invalid or over-budget ancestor lineage");
            if (state.pendingOwner.has(owner))
                fail(state, "an owner already has an outstanding request");
            let paragraphs = state.seen.get(owner);
            if (!paragraphs) {
                paragraphs = new WeakMap();
                state.seen.set(owner, paragraphs);
            }
            let states = paragraphs.get(options.paragraph);
            if (!states) {
                states = new Set();
                paragraphs.set(options.paragraph, states);
            }
            const key = JSON.stringify([options.cutLine, options.cellPage, options.rowIndex, options.tablePage, signature]);
            if (states.has(key))
                fail(state, "a rejected split state repeated without progress");
            if (state.used >= state.maxRequests)
                fail(state, "document-pass request budget exhausted");
            const request = Object.freeze({
                id: state.nextId++, generation: state.generation, owner: owner, paragraph: options.paragraph,
                cutLine: options.cutLine, cellPage: options.cellPage, rowIndex: options.rowIndex,
                tablePage: options.tablePage, signature: Object.freeze(signature.slice())
            });
            const lineage = [owner, ...ancestors];
            ownerFrame.retry = true;
            state.records.set(request, {
                status: ancestors.length ? "Requested" : "Mapped",
                lineage: lineage, hop: 0, page: options.cellPage, row: options.rowIndex
            });
            states.add(key);
            state.used++;
            state.pending.add(request);
            state.pendingOwner.set(owner, request);
            return request;
        }

        map(request, mapping) {
            const state = stores.get(this), value = pending(state, request);
            if (value.status !== "Requested" || !mapping
                || mapping.from !== value.lineage[value.hop] || mapping.to !== value.lineage[value.hop + 1])
                fail(state, "mapping skipped or changed an ancestor");
            integer(state, mapping.page, "mapped page");
            integer(state, mapping.row, "mapped row");
            value.hop++;
            value.page = mapping.page;
            value.row = mapping.row;
            if (value.hop === value.lineage.length - 1)
                value.status = "Mapped";
        }

        mapping(request) {
            const state = stores.get(this), value = pending(state, request);
            return Object.freeze({owner: value.lineage[value.hop], page: value.page, row: value.row});
        }

        startReplay(request) {
            const state = stores.get(this), value = pending(state, request);
            if (value.status !== "Mapped")
                fail(state, "replay began before mapping completed");
            value.status = "Replaying";
        }

        consume(request, token) {
            const state = stores.get(this), value = pending(state, request);
            if (value.status !== "Replaying" || frame(state, token).owner !== request.owner)
                fail(state, "consumption requires the replaying original owner");
            value.status = "Consumed";
            state.pending.delete(request);
            state.pendingOwner.delete(request.owner);
        }

        cancel(request) {
            const state = stores.get(this), value = pending(state, request);
            value.status = "Cancelled";
            state.pending.delete(request);
            state.pendingOwner.delete(request.owner);
        }

        status(request) {
            const state = stores.get(this);
            const value = object(request) ? state.records.get(request) : null;
            if (!value)
                fail(state, "unknown request identity");
            return value.status;
        }

        // The interactive lifecycle records its original typed error separately.
        // Invalidate pending work without throwing a replacement cleanup error.
        invalidate() {
            const state = stores.get(this);
            state.failed = true;
            for (const request of state.pending)
                state.records.get(request).status = "Invalid";
            state.pending.clear();
            state.pendingOwner = new WeakMap();
        }

        abort(reason) {
            fail(stores.get(this), reason || "controller invariant failure");
        }

        snapshot() {
            const state = stores.get(this);
            return Object.freeze({generation: state.generation, requestsUsed: state.used, workUsed: state.workUsed, pending: state.pending.size, failed: state.failed});
        }
    }

    CCellWidowReplayProtocol.QualificationError = CellWidowQualificationError;
    window.AscWord = window.AscWord || {};
    window.AscWord.CCellWidowReplayProtocol = CCellWidowReplayProtocol;
})(window);
