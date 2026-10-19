# Table direction regression checks

Run the source checks from the SDK root:

    node --test tests/word/table-direction*.test.cjs
    python3 tests/code-style/check.py

The Node tests evaluate the actual property and geometry source. They cover the
three-state BidiVisual property, merge/default/copy semantics, internal record
boundaries, reflected cell metrics, table anchors and border drawing rules.
The interaction checks exercise the actual table methods and mapping helpers:
physical/paragraph key intent, nested container routing, fresh target caret
positions, document-order range collapse, physical vertical intervals, ruler
reflection, resize deltas/paint, exact-row shrink redistribution, stale geometry
and caret-preserving text selection initialization with later-page anchors.
They do not substitute for an editor or converter run.

For the real-engine check, open fixtures/rtl-table-direction.docx in a local
editor using the candidate production bundles, then load table-direction-engine.js
in that editor frame and call AscTestTableDirection(). Run in a disposable
document: the checks create table styles and history actions. Expect ten PASS
results, including public tblApply/Undo/Redo with actual cell-position assertions,
positive RTL ruler/outline edits and a recording overlay using the actual page
and rotated/scaled matrix methods. The overlay checks paint/hit geometry; it is
not a transformed-document hierarchy or native UI acceptance test.
The fixture is an original synthetic three-column document, CC0-1.0; its paragraph
bidi and logical cell order are separate from table direction. It is not an actual
Word output.

RTL column-border/ruler dragging, the southwest outline handle and leading
row/cell selection zones now map physical coordinates to the logical model.
Physical Left/Right intent is distinct from paragraph character direction;
container routing preserves nested table direction and document-order ranges.
Vertical navigation selects cells by physical intervals. Whole-table movement
records floating anchor changes before recalculation. Text selection starts
without temporarily selecting an entire cell or moving nested caret endpoints.

RTL gestures cancel if their resolved layout/style, target identity, parent
transform or direction changes during the preview. Ruler gestures also retain
an immutable starting markup identity. Pencil/eraser grid mutations remain
gated; the mode stays active but performs no mutations on RTL tables.

Source and real-engine checks do not replace trusted browser gestures or actual
Word comparison. Controlled browser evidence and remaining limits are recorded
in the product RTL table document and owned interaction work directory. In
particular, clipped overflowing text in short exact-height nested cells and
pagination/header-count changes need separate qualification.

The portable internal-property presence/value flags occupy previously unused
bits and add no bytes to the unframed record. The previous reader can continue
reading later fields, but ignores direction. Mixed-version collaborative history
or direction fidelity is not promised.

DOCX and ODT conversion must additionally be tested using a rebuilt native/full
Wasm converter. An old converter drops an explicit false override of an RTL style
and loses ODT table writing-mode. SDK-only PASS does not close these gates.

## release review r2 pointer and LTR regression scope
The four vertical functions retain the original LTR loops; mixed-stack physical
routing is entered only from RTL context. The fixed-field/revision snapshot
checks Inline/AllowOverlap/Distance before recalc, actual history writer position,
layout/document revisions and parent/transform identity without table traversal.
Ruler start captures a fresh shared token; commit revalidates without recapturing.
Source units include the real Rulers-to-Table empty-action path and external edit
negative. History-off inherited styles require normal style invalidation/recalc;
collaboration/cursor/save writes can conservatively cancel a preview. Product
tests/rtl-tables holds real paired LTR and runtime guards; core rebuild remains
separate. Exact profiles and independent receipt are in the Z handoff evidence.
