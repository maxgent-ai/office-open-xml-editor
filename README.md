# Office Open XML Viewer — Maxgent Fork

**Office document previews, with a PPTX editing SDK for your own application.**

This repository is Maxgent's fork of
[Yuki Yokotani's office-open-xml-viewer](https://github.com/yukiyokotani/office-open-xml-viewer).
The upstream project provides browser-based DOCX, XLSX, and PPTX previews using
Rust/WebAssembly parsers and Canvas renderers. We build on that foundation to
add PPTX editing, local undo/redo, and explicit saves through OfficeCLI.

**[Upstream preview demo](https://ooxml.silurus.dev)** ·
**[Editor API guide](packages/pptx-editor/README.md)** ·
**[Report a fork issue](https://github.com/maxgent-ai/office-open-xml-viewer/issues)** ·
**[MIT license](LICENSE)**

## Upstream and this fork

For preview examples, rendering capabilities, and viewer documentation, visit the
[upstream website](https://ooxml.silurus.dev) and
[upstream README](https://github.com/yukiyokotani/office-open-xml-viewer#readme).
The upstream demo showcases viewing; it does not include this fork's editor.

| Area | This fork |
| --- | --- |
| DOCX, XLSX, and PPTX previews | Inherited from upstream |
| PPTX editing | Separate `@maxgent/ooxml-pptx-editor` package |
| DOCX and XLSX editing | Not currently supported |
| Editor interface | You provide toolbars, text inputs, selection overlays, and save controls |
| File persistence | You provide a backend that applies OfficeCLI commands and saves the PPTX |

Upstream intentionally focuses on read-only viewing. Editor development lives in
this fork, with the viewer foundation kept in sync with upstream.
Use the `@maxgent` packages for editing integrations; upstream's `@silurus`
packages do not include the fork's editor hooks.

## Editing capabilities

The SDK supports these operations on eligible slide content:

| Capability | Available API |
| --- | --- |
| Change shape text and formatting | `UpdateTextMutation`: whole-shape text, paragraph text/style, and span style edits |
| Move, resize, rotate, and flip shapes | `UpdateShapeMutation` |
| Change shape fills and outlines | `UpdateShapeMutation` |
| Add shapes | `AddElementMutation` |
| Remove shapes, pictures, tables, and charts | `RemoveElementMutation` |
| Insert empty slides and remove slides | `InsertSlideMutation`, `RemoveSlideMutation` |
| Select elements on the canvas | `PptxEditorSelectionController` |
| Undo and redo local changes | `PptxEditorSession.undo()` / `.redo()` |
| Track unsaved changes and save state | `PptxEditorSession.getSnapshot()` / `.subscribe()` |
| Save accumulated changes in one batch | `PptxEditorSession.save()` |

Edits update the in-memory presentation and canvas without a network request.
An explicit save translates the pending editing history into one OfficeCLI batch.
The session tracks the saved history position and retains the local draft if saving fails.

### Current boundaries

- Editing requires a presentation loaded with `mode: 'main'`.
- Text and shape updates target direct slide shapes. Master and layout decorations remain read-only.
- Picture, table, and chart removal is supported; their content editing is not.
  Audio/video editing is not supported.
- Grouped or projected elements can produce unsupported OfficeCLI paths.
  Selection alone does not guarantee that an element can be saved.
- Text span replacement is not supported. Span formatting and paragraph text replacement use separate mutations.
- Shape insertion does not preserve every custom geometry or rich-text property.
  This is a scoped editor SDK, not a complete PowerPoint replacement or a general lossless round-trip API.
- Unsaved deletions can be undone. Saved deletions and other operations without a faithful inverse establish an undo boundary.
- Drafts live in memory. Persistent drafts and conflict detection across clients require host application support.

See the [editor guide](packages/pptx-editor/README.md#current-limitations)
for detailed targeting and formatting limits.

## Quick start

### Install

Install matching fork releases. This example uses `0.86.1-maxgent.1`, which
includes the local editing and explicit save API shown below.

```bash
pnpm add @maxgent/ooxml@0.86.1-maxgent.1 @maxgent/ooxml-pptx-editor@0.86.1-maxgent.1
```

The packages use ES modules. Your browser build must serve the viewer's `.wasm`
assets. See the [upstream bundler notes](https://github.com/yukiyokotani/office-open-xml-viewer#readme)
for setup details, using the `@maxgent/ooxml` package name in your integration.

### Open a deck and edit selected text

Provide a canvas and a `sendBatch` function connected to your persistence backend.
The function below returns actions that you can bind to your own controls.

```ts
import { PptxPresentation, PptxViewer } from '@maxgent/ooxml/pptx';
import {
  PptxEditorSession,
  PptxEditorViewBinding,
  PptxEditorViewerHost,
  PptxEditorSelectionController,
  UpdateTextMutation,
  type OfficeCliBatchSender,
} from '@maxgent/ooxml-pptx-editor';

export async function openEditor(
  canvas: HTMLCanvasElement,
  source: string | ArrayBuffer,
  sendBatch: OfficeCliBatchSender,
) {
  const presentation = await PptxPresentation.load(source, { mode: 'main' });
  const viewer = PptxViewer.fromPresentation(canvas, presentation);
  const session = new PptxEditorSession({
    presentation: await presentation.toEditorPresentation(),
    sendBatch,
    createSaveId: () => crypto.randomUUID(),
  });
  const binding = new PptxEditorViewBinding({
    session,
    host: new PptxEditorViewerHost(viewer, presentation),
    onRenderError: (cause) => console.error('Preview update failed', cause),
  });
  await binding.whenIdle();

  const selection = new PptxEditorSelectionController({ session, host: viewer });

  return {
    viewer,
    session,
    selection,
    replaceSelectedText(value: string) {
      const selected = selection.getSnapshot().selection;
      if (!session.getSnapshot().canEdit || !selected) return;
      if (selected.element.type !== 'shape' || !selected.isOfficeCliTargetable) return;
      session.apply({
        id: crypto.randomUUID(),
        mutations: [new UpdateTextMutation({ target: selected.target, value })],
      });
    },
    save: () => session.save(),
    destroy() {
      selection.dispose();
      binding.dispose();
      session.dispose();
      viewer.destroy();
      presentation.destroy();
    },
  };
}
```

Pass a URL or a file's `ArrayBuffer` as `source`. Select a shape on the canvas,
then call `replaceSelectedText()` from your text input.
Call `session.undo()` and `session.redo()` from your history controls.
Use `canUndo`, `canRedo`, and `canEdit` from the session snapshot to enable controls.

Before calling `save()`, apply any text still held in an active input.
Display `dirty` and `saveStatus` from the session snapshot, and handle both the
save result and promise rejection. `destroy()` releases resources when the editor closes.

### Connect saving

```text
Your controls → session.apply / undo / redo → local model → canvas
Your save button → session.save → sendBatch → OfficeCLI backend → saved PPTX
```

`sendBatch` receives `batch.commands` and `batch.commandId`.
An existing `applyCommands` adapter can send them as `commands` and `requestId`.
The backend must apply the batch to the source file and persist the result.
The current translator targets OfficeCLI `1.0.139`; keep the backend compatible
with the batch envelope's `officecliVersion`.
The editor package does not include a persistence server or a browser PPTX writer.

Return an `OfficeCliBatchSendResult` according to the actual persistence outcome:

| Status | Meaning | Session behavior |
| --- | --- | --- |
| `confirmed` | All commands and the resulting file are durably saved | Mark the current history position as saved |
| `rejected` | The authoritative file is unchanged | Keep the draft and allow a new save attempt |
| `unknown` | Partial application or an uncertain result is possible | Keep the draft and lock editing until reconciled |

Both `rejected` and `unknown` results include a `cause`.
A timeout is not evidence of rejection. The session does not retry automatically.
While saving, the session blocks edits, undo, redo, and additional saves.
Saving an unchanged session returns `unchanged` without calling the backend.

For unknown outcomes, file version checks, and recovery, follow the
[save integration guide](packages/pptx-editor/README.md#save-outcomes).

Existing integrations using `submit()` must migrate to `apply()` and explicit
`save()` calls. See the [migration steps](packages/pptx-editor/README.md#session-api-and-migration).

## Development

Use Node.js 24, the pnpm version declared in `package.json`, Rust, and `wasm-pack`.

```bash
git clone https://github.com/maxgent-ai/office-open-xml-viewer.git
cd office-open-xml-viewer
pnpm install --frozen-lockfile
pnpm build:wasm
pnpm storybook
```

Storybook runs at `http://localhost:6006`. Public stories are available in the
repository; private sample files are not distributed.

For editor changes:

```bash
pnpm --filter @maxgent/ooxml-pptx-editor test
pnpm --filter @maxgent/ooxml-pptx-editor typecheck
pnpm --filter @maxgent/ooxml-pptx-editor build
```

| Directory | Purpose |
| --- | --- |
| [`packages/pptx-editor`](packages/pptx-editor) | Editing sessions, mutations, selection, view binding, and OfficeCLI translation |
| [`packages/pptx`](packages/pptx) | PPTX parser and viewer, including the fork's editing hooks |
| [`packages/docx`](packages/docx), [`packages/xlsx`](packages/xlsx) | DOCX and XLSX previews |
| [`packages/core`](packages/core), [`packages/ooxml-common`](packages/ooxml-common) | Shared rendering and parsing |

## Contributing and support

Report fork-specific problems in
[Maxgent's issue tracker](https://github.com/maxgent-ai/office-open-xml-viewer/issues).
Include the package versions, reproduction steps, expected behavior, and a
redistributable sample when possible. Keep private documents out of public issues and PRs.

For changes, read [AGENTS.md](AGENTS.md), add focused regression coverage for bug
fixes, and open a PR against this fork's `main` branch. Consult the applicable
OOXML specification before changing document behavior.

## License and acknowledgments

Licensed under the [MIT License](LICENSE). See
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for third-party notices.

Thanks to [Yuki Yokotani](https://github.com/yukiyokotani) and the
[upstream contributors](https://github.com/yukiyokotani/office-open-xml-viewer/graphs/contributors)
for the parsing and rendering foundation that makes this fork possible.
