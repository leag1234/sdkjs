// SPDX-License-Identifier: AGPL-3.0-only
(function (window) {
    "use strict";

    const Protocol = window.AscWord.CCellWidowReplayProtocol;
    const passes = new WeakMap(), capabilities = new WeakMap();
    const activeOwners = new WeakMap(), frames = new WeakMap();
    // Qualification bounds, not a promise that every document fits this budget.
    const LIMITS = Object.freeze({maxRequests: 1024, maxDepth: 32, maxSteps: 65536});

    function resetWidow(info) {
        info.WidowControlParagraph = null;
        info.WidowControlLine = -1;
        info.WidowControlReset = false;
    }

    function passFor(doc) {
        const generation = doc.GetRecalcId();
        let pass = passes.get(doc);
        if (!pass) {
            pass = {doc: doc, generation: generation,
                protocol: new Protocol({generation: generation, ...LIMITS}),
                owners: new WeakMap(), infos: new Set(), eligible: new WeakMap(), needsRestart: false, cacheHits: 0,
                layoutFailure: null, failureGeneration: -1, failureReported: false};
            passes.set(doc, pass);
        } else {
            if (generation === pass.generation && pass.needsRestart)
                fail(pass, "the current cell calculation was aborted");
            pass.protocol.begin(generation);
            if (generation !== pass.generation) {
                for (const info of pass.infos)
                    resetWidow(info);
                pass.generation = generation;
                pass.owners = new WeakMap();
                pass.infos.clear();
                pass.eligible = new WeakMap();
                pass.needsRestart = false;
                pass.cacheHits = 0;
            }
        }
        return pass;
    }

    function fail(pass, reason) {
        if (pass)
            pass.protocol.abort(reason);
        throw new Protocol.QualificationError(reason);
    }

    function rootContent(content) {
        return content && content.GetDocumentContentForRecalcInfo
            ? content.GetDocumentContentForRecalcInfo() : content;
    }

    function active(content) {
        const owner = rootContent(content);
        const value = owner ? activeOwners.get(owner) : null;
        if (!value)
            return null;
        if (value.doc.GetRecalcId() !== value.generation)
            fail(value.pass, "stale active cell generation");
        if (value.pass.protocol.snapshot().failed || value.pass.needsRestart)
            fail(value.pass, "active cell belongs to a failed pass");
        if (capabilities.get(value.doc) !== value.capability)
            return null;
        return value;
    }

    function liveRoot(table, doc, pass) {
        let element = table, parent = table.Parent, depth = 0;
        const seen = new Set();
        while (parent) {
            if (depth++ > LIMITS.maxDepth || seen.has(parent))
                fail(pass, "transparent owner lineage exceeded its depth or repeated");
            seen.add(parent);
            if (!Array.isArray(parent.Content) || parent.Content[element.Index] !== element)
                return null;
            if (parent instanceof CDocument)
                return parent === doc ? element : null;
            if (!(parent instanceof CDocumentContent) || !(parent.Parent instanceof CBlockLevelSdt)
                || parent.Parent.Content !== parent)
                return null;
            element = parent.Parent;
            parent = element.Parent;
        }
        return null;
    }

    function tablePolicy(table) {
        return table.Is_Inline()
            && !(table.HeaderInfo && table.HeaderInfo.Count > 0)
            && !table.Content.some(row => row.Get_Height().HRule === window.Asc.linerule_Exact
                || row.Content.some(cell => cell.GetVMerge() === vmerge_Continue));
    }

    function tableEligible(table, pass) {
        if (!pass.eligible.has(table))
            pass.eligible.set(table, tablePolicy(table));
        return pass.eligible.get(table);
    }

    function placement(value) {
        const page = value.owner.Pages[value.cellPage];
        if (!page)
            fail(value.pass, "completed cut lacks its current content page");
        const key = [value.generation, value.capability.page, value.tablePage, value.cellPage,
            page.X, page.Y, page.XLimit, page.YLimit];
        if (!key.every(Number.isFinite))
            fail(value.pass, "completed cut has invalid placement geometry");
        return key;
    }

    function prepareCachedCut(value) {
        const meta = value.meta, cache = meta.cache;
        if (!cache || meta.request)
            return;
        const key = placement(value);
        if (cache.table !== value.table || cache.placement.some((v, i) => v !== key[i])) {
            if (meta.info.WidowControlParagraph === cache.paragraph
                && meta.info.WidowControlLine === cache.cutLine)
                resetWidow(meta.info);
            meta.cache = null;
            value.cachedReplay = null;
            return;
        }
        const paragraph = cache.paragraph;
        if (!paragraph.Parent || rootContent(paragraph.Parent) !== value.owner
            || paragraph.Parent.Content[paragraph.Index] !== paragraph)
            fail(value.pass, "cached cut lost its live paragraph identity");
        if (meta.info.WidowControlParagraph && meta.info.WidowControlParagraph !== paragraph)
            fail(value.pass, "cached cut would replace another widow owner");
        value.cachedReplay = cache;
        meta.info.Set_WidowControl(paragraph, cache.cutLine);
    }

    function metaFor(pass, owner) {
        let meta = pass.owners.get(owner);
        if (!meta) {
            resetWidow(owner.RecalcInfo);
            meta = {info: owner.RecalcInfo, request: null, cache: null, destination: null};
            pass.owners.set(owner, meta);
            pass.infos.add(meta.info);
        } else if (meta.info !== owner.RecalcInfo) {
            fail(pass, "cell recalculation owner changed within a generation");
        }
        return meta;
    }

    function replayDestination(value, request, documentPage, tablePage, cellPage) {
        return Object.freeze({request: request, generation: value.generation,
            table: value.table, owner: value.owner, element: value.capability.element,
            documentPage: documentPage, tablePage: tablePage, cellPage: cellPage, row: value.row});
    }

    function liveParagraph(paragraph, owner, pass) {
        let element = paragraph, content = paragraph.Parent, depth = 0;
        const seen = new Set();
        while (content) {
            if (depth++ > LIMITS.maxDepth || seen.has(content))
                fail(pass, "pending paragraph lineage exceeded its depth or repeated");
            seen.add(content);
            if (!Array.isArray(content.Content) || content.Content[element.Index] !== element)
                return false;
            if (content === owner)
                return true;
            if (!(content instanceof CDocumentContent) || !(content.Parent instanceof CBlockLevelSdt)
                || content.Parent.Content !== content)
                return false;
            element = content.Parent;
            content = element.Parent;
        }
        return false;
    }

    function validatePendingPlacement(value) {
        const meta = value.meta, request = meta.request, destination = meta.destination;
        const protocol = value.pass.protocol;
        if (!request || protocol.status(request) !== "Replaying" || !destination
            || destination.request !== request || destination.generation !== value.generation
            || request.generation !== value.generation || destination.table !== value.table
            || destination.owner !== value.owner || destination.element !== value.capability.element
            || destination.documentPage !== value.capability.page
            || destination.tablePage !== value.tablePage || destination.cellPage !== value.cellPage
            || destination.row !== value.row || request.rowIndex !== value.row)
            fail(value.pass, "pending cut reached a different replay destination");
        const cell = value.owner.Parent, row = value.table.Content[value.row];
        if (!cell || cell.Content !== value.owner || cell.Row !== row || row.Index !== value.row
            || row.Content[cell.Index] !== cell || cell.GetTable() !== value.table
            || liveRoot(value.table, value.doc, value.pass) !== value.capability.element
            || value.table.GetAbsolutePage(value.tablePage) !== value.capability.page)
            fail(value.pass, "pending cut lost its live cell or page mapping");
        const mapping = protocol.mapping(request);
        if (meta.direction === "previous"
            ? mapping.owner !== value.doc || mapping.page !== destination.documentPage
                || mapping.row !== destination.element.Index
            : mapping.owner !== value.owner || mapping.page !== destination.cellPage
                || mapping.row !== destination.row)
            fail(value.pass, "pending cut no longer matches its authenticated ancestor mapping");
        const paragraph = request.paragraph;
        if (!liveParagraph(paragraph, value.owner, value.pass))
            fail(value.pass, "pending paragraph lost its cell membership");
        return paragraph;
    }

    const api = {
        // Fast recalculation has no scheduler capability. Inspect live membership
        // without starting/resetting a generation or changing the widow owner.
        shouldUseSlowParagraph: function (paragraph) {
            const cell = paragraph.IsTableCellContent(true);
            if (!cell)
                return false;
            const table = cell.GetTable(), doc = table.LogicDocument;
            if (!(doc instanceof CDocument) || !doc.IsDocumentEditor() || doc.MainDocument === false || doc.PrintSelection === true || table.ColumnsCount !== 1)
                return false;
            const owner = rootContent(paragraph.Parent), row = cell.Row;
            if (!paragraph.Parent || paragraph.Parent.Content[paragraph.Index] !== paragraph
                || owner !== cell.Content || owner.Parent !== cell
                || table.Content[row.Index] !== row || row.Content[cell.Index] !== cell
                || cell.IsVerticalText() || cell.IsInHeader()
                || !liveRoot(table, doc, null) || !tablePolicy(table))
                return false;
            const pass = passes.get(doc), meta = pass && pass.owners.get(owner);
            if (meta && meta.request && meta.request.paragraph === paragraph)
                return true;
            return paragraph.Pages.length > 1 && !!paragraph.Get_CompiledPr2().ParaPr.WidowControl;
        },

        beginGeneration: function (doc) {
            if (!(doc instanceof CDocument) || !doc.IsDocumentEditor() || doc.MainDocument === false || doc.PrintSelection === true)
                return false;
            const previous = passes.get(doc), generation = doc.GetRecalcId();
            const needsFull = previous && generation !== previous.generation
                && (previous.needsRestart || previous.protocol.snapshot().failed
                    || previous.protocol.snapshot().pending > 0);
            passFor(doc);
            return !!needsFull;
        },

        withDocument: function (doc, element, page, column, columns, callback) {
            const previous = capabilities.get(doc);
            if (!(doc instanceof CDocument) || !doc.IsDocumentEditor() || doc.MainDocument === false || doc.PrintSelection === true || columns !== 1 || column !== 0) {
                capabilities.delete(doc);
                try {
                    return callback();
                } finally {
                    if (previous)
                        capabilities.set(doc, previous);
                }
            }
            const pass = passFor(doc);
            const capability = {doc: doc, element: element, page: page, column: column,
                generation: pass.generation, pass: pass};
            capabilities.set(doc, capability);
            try {
                return callback();
            } catch (error) {
                api.failLayout(doc, error);
                throw error;
            } finally {
                if (previous)
                    capabilities.set(doc, previous);
                else
                    capabilities.delete(doc);
            }
        },

        calculateElement: function (doc, element, elementPage, page, column, columns) {
            return api.withDocument(doc, element, page, column, columns,
                function () { return element.Recalculate_Page(elementPage); });
        },

        enterCell: function (table, cell, tablePage, cellPage) {
            const doc = table.LogicDocument, capability = doc && capabilities.get(doc);
            if (!capability)
                return null;
            if (capability.generation !== doc.GetRecalcId() || capability.pass.generation !== capability.generation)
                fail(capability.pass, "stale scheduler capability");
            const pass = capability.pass;
            if (pass.needsRestart)
                fail(pass, "the current cell calculation was aborted");
            const row = cell.Row, owner = cell.Content;
            if (liveRoot(table, doc, pass) !== capability.element || !tableEligible(table, pass)
                || table.Content[row.Index] !== row || row.Content[cell.Index] !== cell
                || owner.Parent !== cell || cell.GetTable() !== table
                || cell.IsVerticalText() || cell.IsInHeader())
                return null;
            if (!Number.isSafeInteger(tablePage) || tablePage < 0
                || !Number.isSafeInteger(cellPage) || cellPage < 0)
                fail(pass, "negative or invalid cell/table page");
            if (activeOwners.has(owner))
                fail(pass, "reentrant frame for the same cell owner");
            const meta = metaFor(pass, owner);
            const value = {doc: doc, owner: owner, table: table, row: row.Index,
                tablePage: tablePage, cellPage: cellPage, generation: pass.generation,
                capability: capability, pass: pass, meta: meta, token: null, closed: false, result: undefined,
                cachedReplay: null, completedCut: null};
            if (meta.request) {
                const paragraph = validatePendingPlacement(value);
                meta.info.Set_WidowControl(paragraph, meta.request.cutLine);
            }
            value.token = pass.protocol.open(owner, {
                page: cellPage, retry: !!meta.request || !!meta.cache
            });
            const frame = Object.freeze({});
            frames.set(frame, value);
            activeOwners.set(owner, value);
            return frame;
        },

        leaveCell: function (frame, result) {
            if (!frame)
                return;
            const value = frames.get(frame);
            if (!value || value.closed)
                fail(value && value.pass, "unknown or closed cell frame");
            const protocol = value.pass.protocol, meta = value.meta;
            value.result = result;
            try {
                if (value.doc.GetRecalcId() !== value.generation || value.pass.generation !== value.generation) {
                    if (!protocol.snapshot().failed)
                        fail(value.pass, "stale cell frame cannot finalize a newer generation");
                    return;
                }
                if (result === undefined) {
                    value.pass.needsRestart = true;
                    if (meta.request && !protocol.snapshot().failed)
                        protocol.cancel(meta.request);
                    meta.request = null;
                    meta.destination = null;
                    meta.cache = null;
                    resetWidow(meta.info);
                } else if (meta.request && !protocol.snapshot().failed) {
                    const status = protocol.status(meta.request);
                    if (status === "Replaying") {
                        if (meta.info.WidowControlReset) {
                            fail(value.pass, "matched cut was not acknowledged before paragraph reset");
                        } else if ((result & 0xffff) === recalcresult2_End) {
                            fail(value.pass, "replay completed without consuming its cut line");
                        }
                    } else if ((result & 0xffff) !== recalcresult2_CellReplayRequested) {
                        fail(value.pass, "pending request was lost as ordinary cell completion");
                    }
                }
                if (result !== undefined && !protocol.snapshot().failed) {
                    if (value.cachedReplay)
                        fail(value.pass, "cached cut was not reapplied before cell completion");
                    if (value.completedCut) {
                        if ((result & 0xffff) !== recalcresult2_NextPage || meta.request)
                            fail(value.pass, "completed current cut did not produce its next-page fragment");
                        meta.cache = value.completedCut;
                    }
                }
            } finally {
                value.requestAtReturn = meta.request;
                activeOwners.delete(value.owner);
                protocol.close(value.token);
                value.closed = true;
            }
        },

        calculateCell: function (table, cell, tablePage, cellPage) {
            const frame = api.enterCell(table, cell, tablePage, cellPage);
            let result;
            try {
                result = cell.Content.Recalculate_Page(cellPage, true);
                if (frame && (!Number.isSafeInteger(result) || result < 0 || result > 0x3ffff
                    || ![0, 1, 2, 3, 4, 8].includes(result & 0xffff)))
                    fail(frames.get(frame).pass, "invalid eligible cell result");
                return {result: result, context: frame};
            } finally {
                api.leaveCell(frame, result);
            }
        },

        getOwner: function (content) {
            const value = active(content);
            return value ? value.owner : null;
        },

        tickContent: function (content) {
            const value = active(content);
            if (value) {
                value.pass.protocol.step(value.token, content);
                // The root page limits have been refreshed before this hook.
                // A child SDT shares the ledger but cannot prepare its root cache.
                if (content === value.owner)
                    prepareCachedCut(value);
            }
        },

        request: function (owner, paragraph, cutLine, paragraphPage, direction) {
            const value = active(owner);
            if (!value || value.owner !== owner)
                fail(value && value.pass, "request outside its active cell owner");
            if (!paragraph.Parent || rootContent(paragraph.Parent) !== owner
                || paragraph.Parent.Content[paragraph.Index] !== paragraph)
                fail(value.pass, "paragraph is not a live member of this cell");
            direction = direction === undefined ? "previous" : direction;
            if (direction !== "current" && direction !== "previous")
                fail(value.pass, "invalid replay direction");
            const page = owner.Pages[value.cellPage];
            if (!page || !Number.isSafeInteger(paragraphPage) || paragraphPage < 0)
                fail(value.pass, "missing owner page or invalid paragraph page");
            const request = value.pass.protocol.request(value.token, {
                paragraph: paragraph, cutLine: cutLine, cellPage: value.cellPage,
                rowIndex: value.row, tablePage: value.tablePage,
                signature: [page.X, page.Y, page.XLimit, page.YLimit, paragraphPage],
                ancestors: direction === "current" ? [] : [value.table, value.doc]
            });
            value.meta.request = request;
            value.meta.cache = null;
            value.cachedReplay = null;
            value.meta.direction = direction;
            value.meta.destination = direction === "current"
                ? replayDestination(value, request, value.capability.page, value.tablePage, value.cellPage) : null;
            if (direction === "current")
                value.pass.protocol.startReplay(request);
            value.meta.info.Set_WidowControl(paragraph, cutLine);
            return request;
        },

        // Paragraph clears its legacy reset flag before returning to content.
        // Authenticate consumption at the matched cut, while its owner is active.
        consumeCut: function (owner, paragraph, line) {
            const value = active(owner);
            if (!value || value.owner !== owner)
                fail(value && value.pass, "cut acknowledgment outside its active owner");
            const meta = value.meta, request = meta.request, cache = value.cachedReplay;
            if (request)
                validatePendingPlacement(value);
            if (meta.info.WidowControlParagraph !== paragraph
                || meta.info.WidowControlLine !== line || !meta.info.WidowControlReset)
                fail(value.pass, "cut acknowledgment does not match its legacy owner state");
            if (!request) {
                const key = placement(value);
                if (!cache || cache !== meta.cache || cache.paragraph !== paragraph
                    || cache.cutLine !== line || cache.placement.some((v, i) => v !== key[i]))
                    fail(value.pass, "cut acknowledgment lacks its authenticated cached outcome");
                value.cachedReplay = null;
                value.completedCut = cache;
                value.pass.cacheHits++;
                return;
            }
            if (value.pass.protocol.status(request) !== "Replaying"
                || request.paragraph !== paragraph || request.cutLine !== line)
                fail(value.pass, "cut acknowledgment does not match its authenticated request");
            value.pass.protocol.consume(request, value.token);
            meta.request = null;
            meta.destination = null;
            if (meta.direction === "current")
                value.completedCut = Object.freeze({table: value.table, paragraph: paragraph,
                    cutLine: line, placement: Object.freeze(placement(value))});
        },

        hasRequest: function (content) {
            const value = active(content);
            return !!(value && value.meta.request);
        },

        requiresRecalculation: function (frame) {
            if (!frame)
                return false;
            const value = frames.get(frame);
            if (!value || !value.closed)
                fail(value && value.pass, "recalculation query requires a completed cell frame");
            if (value.doc.GetRecalcId() !== value.generation)
                fail(value.pass, "stale cell recalculation query");
            if ((value.result & 0xffff) !== recalcresult2_CurPage || !value.requestAtReturn)
                return false;
            if (value.requestAtReturn !== value.meta.request)
                fail(value.pass, "cell frame no longer owns its returned request");
            return value.pass.protocol.status(value.requestAtReturn) === "Replaying";
        },

        rewind: function (frame) {
            const value = frame && frames.get(frame);
            if (!value || !value.closed || !value.meta.request
                || value.requestAtReturn !== value.meta.request || value.meta.direction !== "previous"
                || (value.result & 0xffff) !== recalcresult2_CellReplayRequested)
                fail(value && value.pass, "rewind lacks its completed cell request");
            if (value.doc.GetRecalcId() !== value.generation)
                fail(value.pass, "stale cell request reached its table");
            const tablePage = value.tablePage - 1;
            if (tablePage < 0 || value.cellPage < 1)
                fail(value.pass, "cell rewind would underflow its first table page");
            const documentPage = value.table.GetAbsolutePage(tablePage);
            if (!Number.isSafeInteger(documentPage) || documentPage < 0
                || documentPage !== value.capability.page - 1)
                fail(value.pass, "table-to-document rewind mapping is not the preceding page");
            const protocol = value.pass.protocol, request = value.meta.request;
            protocol.map(request, {from: value.owner, to: value.table, page: tablePage, row: value.row});
            protocol.map(request, {from: value.table, to: value.doc, page: documentPage, row: value.capability.element.Index});
            protocol.startReplay(request);
            value.meta.destination = replayDestination(value, request,
                documentPage, tablePage, value.cellPage - 1);
            return {result: recalcresult_PrevPage
                    | (value.result & (recalcresultflags_Column | recalcresultflags_Page)),
                tablePage: tablePage, documentPage: documentPage, rowIndex: value.row};
        },

        // Failure remains visible across generations until a full layout ends.
        failLayout: function (doc, error) {
            if (!(error instanceof Protocol.QualificationError) || !(doc instanceof CDocument))
                return false;
            const pass = passes.get(doc) || passFor(doc);
            if (pass.layoutFailure === error && pass.failureGeneration === doc.GetRecalcId())
                return true;
            pass.layoutFailure = error;
            pass.failureGeneration = doc.GetRecalcId();
            pass.failureReported = false;
            pass.needsRestart = true;
            pass.protocol.invalidate();
            if (doc.FullRecalc) {
                if (doc.FullRecalc.Id !== null && doc.FullRecalc.Id !== undefined)
                    clearTimeout(doc.FullRecalc.Id);
                doc.FullRecalc.Id = null;
                doc.FullRecalc.Continue = false;
            }
            if (doc.RecalcInfo && doc.RecalcInfo.Set_NeedRecalculateFromStart)
                doc.RecalcInfo.Set_NeedRecalculateFromStart(true);
            // An edit/Undo has already changed history when pagination runs.
            // Close only action bookkeeping; retain every history point/change.
            const action = doc.Action;
            if (action && (action.Start || action.UndoRedo)) {
                const editing = action.Start, undoRedo = action.UndoRedo;
                for (const field of ["Start", "UndoRedo", "Recalculate", "CancelAction", "CheckLock",
                    "UpdateStates", "UpdateSelection", "UpdateInterface", "UpdateRulers",
                    "UpdateUndoRedo", "UpdateTracks", "UpdatePlaceholders"])
                    action[field] = false;
                action.Depth = 0;
                action.PointsCount = 0;
                action.Additional = {};
                action.ScrollToTarget = true;
                if (action.Redraw) {
                    action.Redraw.Start = undefined;
                    action.Redraw.End = undefined;
                }
                if (doc.Api && doc.Api.checkChangesSize)
                    doc.Api.checkChangesSize();
                if (doc.private_UpdateUndoRedo)
                    doc.private_UpdateUndoRedo();
                if (doc.sendEvent) {
                    if (editing)
                        doc.sendEvent("asc_onUserActionEnd");
                    if (undoRedo)
                        doc.sendEvent("asc_onUndoRedo");
                }
            }
            const drawing = doc.DrawingDocument;
            if (drawing) {
                const waiting = drawing.m_arPrintingWaitEndRecalculate;
                drawing.m_arPrintingWaitEndRecalculate = null;
                if (waiting && doc.Api && doc.Api.sync_EndAction)
                    doc.Api.sync_EndAction(window.Asc.c_oAscAsyncActionType.BlockInteraction, waiting[0]);
                if (drawing.isDisableEditBeforeCalculateLA && doc.Api && doc.Api.sync_EndAction) {
                    drawing.isDisableEditBeforeCalculateLA = false;
                    doc.Api.sync_EndAction(window.Asc.c_oAscAsyncActionType.BlockInteraction,
                        window.Asc.c_oAscAsyncAction.Waiting);
                }
                drawing.OnEndRecalculate(false, true);
            }
            return true;
        },

        reportLayoutFailure: function (doc, error) {
            const pass = doc && passes.get(doc);
            if (!pass || pass.layoutFailure !== error || !(error instanceof Protocol.QualificationError)
                || !doc.Api || !doc.Api.sendEvent)
                return false;
            if (!pass.failureReported) {
                pass.failureReported = true;
                doc.Api.sendEvent("asc_onError", window.Asc.c_oAscError.ID.Unknown,
                    window.Asc.c_oAscError.Level.NoCritical);
            }
            return true;
        },

        isLayoutFailed: function (doc) {
            const pass = doc && passes.get(doc);
            return !!(pass && pass.layoutFailure);
        },

        assertLayoutReady: function (doc) {
            const pass = doc && passes.get(doc);
            if (pass && pass.layoutFailure)
                throw pass.layoutFailure;
        },

        completeLayout: function (doc) {
            const pass = doc && passes.get(doc);
            if (!pass || !pass.layoutFailure)
                return;
            if (pass.generation !== doc.GetRecalcId() || pass.needsRestart
                || pass.protocol.snapshot().failed || pass.protocol.snapshot().pending)
                fail(pass, "failed layout cannot complete without a fresh successful generation");
            pass.layoutFailure = null;
            pass.failureReported = false;
        },

        snapshot: function (doc) {
            const pass = passes.get(doc);
            return pass ? Object.freeze({...pass.protocol.snapshot(),
                cacheHits: pass.cacheHits,
                failed: pass.protocol.snapshot().failed || pass.needsRestart}) : Object.freeze({
                generation: doc.GetRecalcId(), requestsUsed: 0, workUsed: 0, pending: 0, failed: false, cacheHits: 0
            });
        }
    };

    window.AscWord.CellWidowReplay = Object.freeze(api);
})(window);
