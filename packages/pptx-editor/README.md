# `@maxgent/ooxml-pptx-editor`

Local PPTX editing with explicit batch saves for Maxgent's `@maxgent/ooxml` fork: mutate an
in-memory `Presentation`, translate commands to OfficeCLI batches, and paint
through a `PptxEditorViewerHost` backed by a loaded `PptxPresentation`.

This package is published **separately** from `@maxgent/ooxml` so upstream
viewer syncs stay thin. Prefer the high-level `PptxEditorSession` +
`PptxEditorViewBinding` surface unless you are extending the editor itself.

| Package | Role |
| --- | --- |
| `@maxgent/ooxml` | Viewer / parser SDK (peer dependency) |
| `@maxgent/ooxml-pptx-editor` | Editor session, mutations, view binding |

## Architecture

```text
UI / host app
  │  apply(command) / undo() / redo() / save()
  ▼
PptxEditorSession          ← local history, saved position, save state, listeners
  └─ save(): replay the path from saved position to current position
        │  sendBatch(OfficeCliBatch), once per explicit save
        ▼
     your transport        ← confirmed | rejected | unknown

PptxEditorViewBinding
  │  session.subscribe → coalesced apply
  ▼
PptxEditorViewerHost       ← PptxViewer + PptxPresentation (mode: 'main')
  └─ replaceSlides + redraw
```

Data ownership:

| Layer | Owns |
| --- | --- |
| Session | Local `Presentation`, undo/redo, saved position, save lock |
| View host | Canvas, package media/theme plumbing, paint |
| Transport | Persistence / OfficeCLI side effects |

The binding never owns the canvas. It only pushes the session’s presentation
into a host that already loaded the same package.

## Install

```bash
pnpm add @maxgent/ooxml @maxgent/ooxml-pptx-editor
```

```ts
import { PptxPresentation, PptxViewer } from '@maxgent/ooxml/pptx';
import type { Presentation } from '@maxgent/ooxml/pptx';
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  PptxEditorViewerHost,
  UpdateTextMutation,
  createElementRef,
  OFFICECLI_BATCH_SEND_STATUSES,
} from '@maxgent/ooxml-pptx-editor';
```

Requires `@maxgent/ooxml >= 0.77.0-0`, which provides the internal main-thread
slide replacement hook. Inside this monorepo, depend on the workspace package:

```json
{
  "dependencies": {
    "@maxgent/ooxml-pptx-editor": "workspace:*"
  }
}
```

## Quick start

Minimal loop: load a viewer in **main** mode, export editor JSON with
`toEditorPresentation()`, open a session on that model, bind them, then apply
local commands. Call `save()` only from the save action.

```ts
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  PptxEditorViewerHost,
  UpdateTextMutation,
  createElementRef,
  OFFICECLI_BATCH_SEND_STATUSES,
  type OfficeCliBatch,
  type OfficeCliBatchSendResult,
} from '@maxgent/ooxml-pptx-editor';
import { PptxPresentation, PptxViewer } from '@maxgent/ooxml/pptx';
import type { Presentation } from '@maxgent/ooxml/pptx';

async function openEditor(args: {
  canvas: HTMLCanvasElement;
  source: string | ArrayBuffer;
  sendBatch: (batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>;
}) {
  // The viewer borrows this main-mode presentation. The caller owns it.
  const loadedPresentation = await PptxPresentation.load(args.source, {
    mode: 'main',
  });
  // Detached editor JSON from the same loaded package (not a second parser source).
  const presentation = await loadedPresentation.toEditorPresentation();
  const viewer = PptxViewer.fromPresentation(args.canvas, loadedPresentation);

  const session = new PptxEditorSession({
    presentation,
    sendBatch: args.sendBatch,
    createSaveId: () => crypto.randomUUID(),
  });

  const host = new PptxEditorViewerHost(viewer, loadedPresentation);
  const binding = new PptxEditorViewBinding({
    session,
    host,
    onRenderError: (cause) => {
      console.error('view apply failed', cause);
      // Host may be stale; recover with a full sync:
      binding.requestRender();
    },
  });
  await binding.whenIdle();

  return { viewer, loadedPresentation, session, binding };
}

function editFirstShapeText(
  session: PptxEditorSession,
  presentation: Presentation,
  nextText: string,
) {
  const slide = presentation.slides[0] as Presentation['slides'][number];
  const element = slide.elements[0] as Presentation['slides'][number]['elements'][number];
  const target = createElementRef(slide, element, 0);

  session.apply({
    id: 'edit-text-1',
    mutations: [new UpdateTextMutation({ target, value: nextText })],
  });

  // The model and view update locally. This function sends no request.
  console.log(session.getSnapshot().dirty);
}

async function onSaveClick(session: PptxEditorSession) {
  const result = await session.save();
  if (result.status === 'rejected' || result.status === 'unknown') {
    console.error(result.cause);
  }
  // Only confirmed (or unchanged) means the current position is saved.
}

// Supply a real sender. Never return confirmed before backend persistence.
```

Teardown:

```ts
binding.dispose();
session.dispose();
viewer.destroy();
```

Dispose the binding before (or with) the session so an in-flight drain does not
call `getSnapshot()` on a disposed session.

## Presentation model requirements

The session operates on `@maxgent/ooxml/pptx` `Presentation` JSON, not on the
zip package itself. Bootstrap that model from the loaded package:

```ts
const loadedPresentation = await PptxPresentation.load(source, { mode: 'main' });
const presentation = await loadedPresentation.toEditorPresentation();
```

Editable slides must expose complete `elementSources` parallel to `elements`
(same length). Text, transform, and style mutations support direct slide shapes
(`origin: 'slide'`). `RemoveElementMutation` also maps slide-origin pictures,
tables, and charts to OfficeCLI paths from their frontend element type.
OfficeCLI `zorder` is derived as the ordinal among
slide-origin entries before `presentationElementIndex` (1-based). This matches
true spTree position for top-level 1:1 shapes; groups / hidden nodes that expand
or skip break that equivalence. Master/layout decorations are
visible in the model but reject edit attempts with
`element.unsupportedOrigin`.

Stable identity:

| Concept | Id source |
| --- | --- |
| Slide | `slide.partName` when present, else `String(slide.index)` via `getSlideMutationId` |
| Element | OOXML `cNvPr` id when present, else `index:<n>` via `getElementMutationId` |

Build refs with `createElementRef(slide, element, elementIndex)` rather than
hand-writing ids.

```ts
import { createElementRef, ELEMENT_ORIGINS } from '@maxgent/ooxml-pptx-editor';

const target = createElementRef(slide, element, elementIndex);
// target.origin === ELEMENT_ORIGINS.SLIDE for editable slide elements
```

## Commands and mutations

A **command** groups one local edit and one history entry. A save combines multiple commands:

```ts
import type { Command } from '@maxgent/ooxml-pptx-editor';

const command: Command = {
  id: 'cmd-1',                 // identifier for the local edit
  mutations: [/* at least one */],
};
```

Built-in mutations:

| Class | Effect | OfficeCLI |
| --- | --- | --- |
| `UpdateTextMutation` | Replace shape plain text, whole-shape styles, or incremental paragraph/span edits (`text` and/or `style`) | `set` path + `{ text, bold, … }` and/or `range=` |
| `UpdateShapeMutation` | Patch shape position, size, rotation, flips, fill, or outline | `set` path + changed shape props |
| `InsertSlideMutation` | Insert an empty slide at a 0-based index | `add` under `/` with `type: 'slide'` and `index` |
| `RemoveSlideMutation` | Remove a slide; undoable locally before saving | `remove` at the current slide path |
| `AddElementMutation` | Insert a slide element at indexes | `add` under slide path |
| `RemoveElementMutation` | Remove a slide-origin shape, picture, table, or chart; undoable locally before saving | type-based stable `remove` path |

Low-level apply without a session:

```ts
import { applyCommand, applyMutation } from '@maxgent/ooxml-pptx-editor';

const { presentation, changedSlideIds, changedElements } = applyCommand(
  currentPresentation,
  command,
);
```

## Session API and migration

This is a breaking session API change. Migration is required:

1. Replace `createCommandId` with `createSaveId: () => crypto.randomUUID()`.
2. Replace `session.submit(command)` with synchronous `session.apply(command)`.
3. Remove per-edit `settled` waits. Undo and redo are synchronous local operations.
4. Call `await session.save()` from the save button.
5. Render save indicators from `dirty` and `saveStatus`. Disable editing when `canEdit` is false.

```ts
session.apply({ id: crypto.randomUUID(), mutations: [mutation] });
session.undo();
session.redo();
const result = await session.save();
```

`apply`, `undo`, and `redo` return a `PptxEditorSessionChange`. They send no requests.
`save` returns `unchanged` without transport when the current history position is saved.
Otherwise, it returns the sender result plus the batch `commandId`.
Local compilation errors reject the save promise before transport and retain the draft.

### Snapshot and events

```ts
interface PptxEditorSessionSnapshot {
  presentation: Presentation;
  dirty: boolean;
  saveStatus: 'idle' | 'saving' | 'failed' | 'unknown';
  saveCommandId?: string;
  saveError?: unknown;
  canEdit: boolean;
  undoDepth: number;
  redoDepth: number;
  canUndo: boolean;
  canRedo: boolean;
}
```

`dirty` compares history positions, not document bytes. An edit that sets the same
value can still be dirty. Returning to the saved position through undo or redo is clean.

Subscribe with `session.subscribe(listener)`. Events include a post-operation snapshot,
`changedSlideIds`, `changedElements`, and an optional command id.
Reasons are `local.applied`, `history.changed`, `save.changed`, and `presentation.resynced`.
The existing view binding consumes local changes without waiting for a save.
Snapshots and mutation inputs must be treated as immutable by callers.

### History and batch order

The session stores structurally shared presentation snapshots and mutation history.
Undo and redo move through local snapshots. A new edit discards the redo branch.
The saved position remains reachable until the next successful save.

At save time, the session finds the common history ancestor of the saved and current positions.
It collects inverse mutations from the saved branch, then forward mutations toward the current position.
The existing translator replays that sequence from the saved presentation.
Every mutation receives the resulting document state, the same batch id, and a global mutation index.
There is no command compression, reordering, final-document diff, or business-owned queue.

Unsaved deletes can be undone by restoring their snapshots, including populated slides.
After saving, operations without a faithful OfficeCLI inverse form an undo boundary.
This includes element deletion, slide deletion, and text operations whose existing inverter returns no inverse.
Shape restoration currently loses rich formatting or geometry, so saved element deletion also forms a boundary.
Later invertible edits remain undoable. Undoing saved text or transform edits creates a new dirty position.

History retains prior changed objects for the session lifetime; unchanged objects are shared.
There is no persistent draft store. Dispose the session when closing the editor.
`resync` also releases old history and explicitly discards the draft.

### Save outcomes

| Sender result | Required evidence | Session effect |
| --- | --- | --- |
| `confirmed` | All commands and the resulting file are durably saved | Advance saved position; preserve history |
| `rejected` | The authoritative file is unchanged | Retain draft and history; allow an explicit new save |
| `unknown` | Outcome or partial application cannot be ruled out | Retain draft; lock editing, undo, redo, and save |

During `saving`, editing, undo, redo, resync, and another save are blocked.
A thrown transport error or malformed result becomes `unknown`.
The session never retries automatically. A batch id does not guarantee backend idempotency.
Do not classify a timeout, partial batch, or uncertain upload as `rejected`.

After independently verifying an unknown batch, resolve that exact id:

```ts
session.resolveUnknown(saveCommandId, { status: 'confirmed' });
// Or, only after verifying that no authoritative change occurred:
session.resolveUnknown(saveCommandId, { status: 'rejected', cause });
```

Do not infer either result from `requestId` alone. If the outcome cannot be verified,
keep the session locked. Alternatively, explicitly discard the draft and reload:

```ts
session.resync(authoritativePresentation); // Discards all local history and draft.
```

Only resync after the previous request has stopped changing the file.
If authoritative media or package resources changed, reload the presentation and viewer too.

### Business and backend integration

The editor package owns local history, dirty tracking, batch generation, and save locking.
The business frontend owns the save button, error display, navigation guards, and transport adapter.
Flush any active text input into `session.apply` before calling `save`.
Slide navigation can wait for `binding.whenIdle()` instead of waiting for backend confirmation.

Keep the current `applyCommands` adapter: send `batch.commands` as `commands` and
`batch.commandId` as `requestId`. Map results according to the evidence table above.
No additional non-upload endpoint, Sandbox lifecycle change, or resident OfficeCLI service is required.

File conflict detection requires backend cooperation. Capture the loaded file version in the sender closure.
Send it as `expectedVersion`; the backend must atomically compare it before applying changes.
On success, return the new version and update that closure before returning `confirmed`.
A precondition conflict is `rejected` only if no change occurred.
Disable further save attempts until the user resolves that conflict; do not silently adopt the newer version.
The editor cannot provide cross-client conflict protection without that server-side check.

## View binding

`PptxEditorViewBinding` connects a session to any host that implements:

```ts
interface PptxEditorViewHost {
  applyPresentation(
    presentation: Presentation,
    options?: { readonly changedSlideIndexes?: readonly number[] },
  ): void | Promise<void>;
}
```

Create the standard host from a viewer that borrows the same main-mode
presentation:

```ts
const host = new PptxEditorViewerHost(viewer, loadedPresentation);
const binding = new PptxEditorViewBinding({
  session,
  host,
  syncOnBind: true, // default: push current session state immediately
  onRenderError: (cause) => {
    console.error(cause);
    binding.requestRender();
  },
});

await binding.whenIdle();
binding.requestRender(); // force a full apply
binding.dispose();
```

Behavior:

- Subscribes to slide changes and slide-count changes.
- Coalesces rapid mutations: while one apply is in flight, later changes merge
  into the next revision (latest snapshot wins).
- Passes `changedSlideIndexes` for incremental patches; uses a full apply when
  `requestRender()` is called or after a failed apply (host state unknown).
- Isolates host failures: errors go to `onRenderError`, the binding stays usable.
- Does **not** auto-retry. After a failure, call `requestRender()` or wait for
  the next session change (which escalates to a full apply).

`PptxEditorViewerHost` behavior:

- Keeps the loaded package’s media / theme plumbing; only swaps in-memory slide
  JSON used by the next paint.
- Invalidates find geometry; clears leftover highlight overlays even when the
  visible slide is not redrawn.
- Replaces the complete in-memory slide list when the slide count changes.
- Throws if the presentation is in `mode: 'worker'`.
- Does not own resources. Dispose the binding, viewer, and presentation
  explicitly.

## Element selection

`PptxEditorSelectionController` maps canvas pointer coordinates into slide EMUs,
hit-tests slide-origin elements from front to back, and exposes the selected
`ElementRef` used by mutations:

```ts
const selection = new PptxEditorSelectionController({
  session,
  host: viewer,
});

selection.subscribe(({ snapshot }) => {
  const selected = snapshot.selection;
  if (!selected) return hideSelectionOverlay();
  showSelectionOverlay(selected.element, selected.target);
});
```

Selection is transient UI state: it is not part of `Presentation`, command
history, or OfficeCLI transport. The controller follows optimistic element
updates, clears itself when the selected element disappears, and reconciles the
selection against the host's current slide before snapshot reads and pointer
input. Pass the viewer itself so `slideIndex` remains a live getter; do not copy
`viewer.slideIndex` into a plain host object, because that freezes its initial
numeric value. Dispose the controller before the session:

```ts
selection.dispose();
binding.dispose();
viewer.destroy();
loadedPresentation.destroy();
session.dispose();
```

The MVP hit test uses each rotated element frame, reverse render order, and a
4-CSS-pixel tolerance for lines. It skips layout/master decorations. Media is
not selectable and blocks selection of elements underneath it. Other elements
without a stable numeric OOXML id can be selected for UI feedback but report
`isOfficeCliTargetable: false`. In
edit mode, keep viewer text selection and hyperlink overlays disabled unless
the app forwards their pointer events into `selectAtClientPoint`.

## OfficeCLI transport

Mutations translate to an `OfficeCliBatch` via `toOfficeCliBatch` /
per-mutation `toOfficeCli`. The product envelope looks like:

```ts
{
  schemaVersion: /* OFFICECLI_BATCH_SCHEMA_VERSION */,
  officecliVersion: /* OFFICECLI_VERSION */,
  commandId: 'edit-text-1',
  commands: [
    { command: 'set', path: '/slide[1]/shape[@id=7]', props: { text: 'Hello' } },
  ],
}
```

Paths use stable slide ordinals and OOXML `cNvPr` ids. Shape updates target
direct slide shapes. Element removal selects `/shape`, `/picture`, `/table`, or
`/chart` from the frontend element type.

You can translate without submitting:

```ts
import { toOfficeCliBatch } from '@maxgent/ooxml-pptx-editor';

const batch = toOfficeCliBatch(presentation, command);
```

## Lower-level building blocks

Most apps should stay on `PptxEditorSession`. These are exported for tests and
custom pipelines:

| API | Role |
| --- | --- |
| `PptxEditorStore` | Optimistic presentation + pending commands + sync state |
| `UndoRedoStack` | Invert commands and drive undo/redo submissions |
| `SerialOfficeCliSubmitter` | Serial queue over `sendBatch` |
| `applyCommand` / `applyMutation` | Pure local apply |

## Current limitations

Document these in product code rather than papering over them:

1. **Main-thread presentation only.** The internal slide replacement hook throws
   in `mode: 'worker'`. Load `PptxPresentation` with `{ mode: 'main' }`.
2. **Slide content only.** The host installs slide models; presentation theme /
   size fields on the session snapshot are not pushed into the viewer.
3. **Slide-origin elements only.** Master/layout elements are not editable.
   Shape text, transform, and style mutations remain shape-only. Removal also
   maps pictures, tables, and charts from their frontend type. Media is rejected
   at translate time (`target.unsupportedElement`) because OfficeCLI has no
   stable `@id` selector for video/audio. A grouped,
   wrapped, or projected element may produce a path that OfficeCLI rejects.
   A rejected save retains the local draft. Saved deletions form an undo boundary
   because full package content cannot be restored faithfully by the existing inverter.
   OfficeCLI `zorder` is derived from `origin: 'slide'` ordinals before
   `presentationElementIndex`; this matches spTree position for top-level 1:1
   shapes, not for groups / hidden nodes that expand or skip.
   `UpdateTextMutation` can patch whole-shape text (`value` + `style`) and
   incremental `edits` (paragraph `text` and/or `style`, or span `style`).
   Paragraph text replacement and span edits require separate mutations.
   Selection-scoped text replacement (span rewrite) remains out of scope.
   Clear-to-inherit (`null`) style keys are
   resolve-then-set for OfficeCLI using paragraph/body/presentation defaults
   (explicit values, not true OOXML attribute removal).
   Character offsets use run-concatenated plain text (OfficeCLI `range` rules).
4. **Complete `elementSources` required** for any editable slide.
5. **Bootstrap via `toEditorPresentation()`.** Prefer
   `await loadedPresentation.toEditorPresentation()` (`mode: 'main'` only) so
   the session JSON comes from the same loaded package the viewer paints.
   Passing a separately parsed `Presentation` remains possible for tests, but
   product hosts should not maintain a second bootstrap source.

## Publishing / fork sync

- Published as `@maxgent/ooxml-pptx-editor` (independent semver, currently
  `0.1.0`).
- Peer-depends on `@maxgent/ooxml` — not wired into the umbrella `exports` map,
  so upstream sync of the viewer SDK does not fight editor packaging.
- Keep the Maxgent-only internal `replaceSlides` hook as the only editor patch in
  `packages/pptx`; viewer composition remains inside `packages/pptx-editor`.

```bash
pnpm --filter @maxgent/ooxml-pptx-editor build
pnpm --filter @maxgent/ooxml-pptx-editor publish --access public
```

## Testing

```bash
pnpm --filter @maxgent/ooxml-pptx-editor test
pnpm --filter @maxgent/ooxml-pptx-editor typecheck
```

Focused suites live under `test/` (`session/`, `rendering/`, `history/`,
`submission/`, `transport/`) and exercise local history, explicit save outcomes,
view coalescing, and OfficeCLI translation.

Real OfficeCLI batch replay coverage:

```bash
OFFICECLI_LIVE=1 pnpm --filter @maxgent/ooxml-pptx-editor test:officecli-live -- test/officecli-live/manual-save.live.test.ts
```

Rebuild PPTX WASM from current source before parser-backed live tests.
The live test compares editable content, geometry, and order with the saved file.
It does not assert equality of backend-generated package identifiers or all OOXML formatting.
