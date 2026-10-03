# Experimental legacy Office sources

Legacy binary Office files can be opened with the ordinary viewers and
loaders through optional model sources:

- `.ppt` with the PPTX loaders and `PptxViewer`

Each reader is a separate opt-in entry. It reads a supported binary subset
directly into the existing document, workbook or presentation model, without
generating an OOXML package, and the ordinary layout and Canvas renderers
draw the result. Importing the ordinary DOCX, XLSX or PPTX entries alone does
not import, fetch or initialize any legacy reader. Without a matching source,
legacy input continues to reject with
`OoxmlError.code === 'legacy-binary-format'`, so no migration is required for
applications that do not opt in.

These readers are experimental and deliberately narrow. They reject input
they cannot represent instead of showing a partial document, and none of them
executes macros, fields, formulas, actions or embedded objects.

## Enabling a source

Pass the source in the format-generic `modelSources` option. It is available
on the DOCX, XLSX and PPTX `load()` options, on the viewers, and on the Node
session options (`openDocxDocument`, `materializeDocxDocument`,
`openXlsxWorkbook`, `openPptxPresentation` and their siblings).

```typescript
import { PptxViewer } from '@silurus/ooxml/pptx';
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const canvas = document.querySelector('canvas') as HTMLCanvasElement;
const viewer = new PptxViewer(canvas, { modelSources: [legacyPptSource()] });
await viewer.load(pptOrPptxBytes);
```

| Input | Entry | Factory | Loaders and viewers |
| --- | --- | --- | --- |
| PPT | `@silurus/ooxml/legacy-ppt` | `legacyPptSource()` | `PptxViewer`, `PptxPresentation.load`, `openPptxPresentation` |

The factory accepts these optional settings:

```typescript
interface LegacySourceOptions {
  wasmUrl?: string; // absolute URL of the reader's WASM
  moduleUrl?: string; // absolute URL of the reader's source module
  maxInputBytes?: number; // defaults to, and may not exceed, 256 MiB
}
```

Creating a source fetches nothing. When a claimed file loads, the parser
Worker (or Node) imports the source's self-contained ES module, emitted as
`legacy-ppt-source-module*.mjs` next to the package files, and initializes
the reader's dedicated WASM, `legacy_ppt_direct_bg.wasm`. Serve these files
with the other package assets, and allow the module URL wherever a Content
Security Policy restricts `script-src` or Worker imports. Applications with a
custom asset pipeline pass absolute `moduleUrl` and `wasmUrl` values:

```typescript
const source = legacyPptSource({
  moduleUrl: new URL('/assets/legacy-ppt-source-module.js', location.href).href,
  wasmUrl: new URL('/assets/legacy_ppt_direct_bg.wasm', location.href).href,
});
```

Module URLs come only from application options; document content never
selects or rewrites them.

A source can report the document's own view preferences. Precedence is: an
explicit caller option, then the source's view default, then the renderer
default. Capabilities that a source lacks degrade the same way for every
format: resource metrics are reported without a ZIP usage snapshot, and
`toMarkdown()` rejects with an "... is unsupported for this source" error.
The legacy PPT reader lacks Markdown export and ZIP accounting.

To cancel a load, destroy the document or viewer, or start another load in
its place. Node sessions keep their `signal` option.

## Admission and failure behavior

A source's synchronous `claim(bytes)` accepts only its own [MS-CFB] family:
a compound file whose directory names `PowerPoint Document` (PPT). A container
that names more than one binary family (`WordDocument` for DOC and `Workbook`
or `Book` for XLS count too), or that carries `EncryptionInfo`, is not
claimed. Every input a
source does not claim takes the unchanged OOXML path, so:

- a legacy file without a matching source rejects with the typed
  `legacy-binary-format` error;
- encrypted OOXML packages keep their existing encryption errors;
- configuring `legacyPptSource()` does not enable DOC or XLS input.

Claimed input larger than `maxInputBytes` throws a `RangeError`. The limit is
resource policy, not an Office format limit. After a source claims a file, a
reader failure rejects the load: there is no fallback to another source or to
the OOXML path. Password-protected legacy binaries, pre-CFB Office formats and
unsupported binary structures are rejected by the readers. These checks are
structural, never filename-based.

## Experimental direct PPT source

Browser:

```typescript
import { PptxPresentation } from '@silurus/ooxml/pptx';
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const presentation = await PptxPresentation.load(legacyPptArrayBuffer, {
  modelSources: [legacyPptSource()],
});
const canvas = document.querySelector('canvas') as HTMLCanvasElement;

try {
  await presentation.renderSlide(canvas, 0, { width: 960 });
} finally {
  presentation.destroy();
}
```

Node:

```typescript
import { openPptxPresentation } from '@silurus/ooxml/node';
import { legacyPptSource } from '@silurus/ooxml/legacy-ppt';

const session = await openPptxPresentation(legacyPptBytes, {
  modelSources: [legacyPptSource()],
});

try {
  for await (const slide of session.slides()) {
    // Consume each ordinary shared slide model.
  }
} finally {
  await session.close();
}
```

The direct source is an experimental, bounded subset, not a full-fidelity
PowerPoint implementation. Besides Markdown and ZIP accounting, it currently
lacks audio/video media and embedded fonts. It rejects constructs it cannot
represent, including paragraph origins that cannot be resolved from the
stored text and master records. Its explicit unsupported diagnostics are
authoritative.

A slide's own `SlideShowSlideInfoAtom.fHidden` marks it hidden; hidden slides
and their content remain in the presentation. Display follows the existing
`PptxViewer` `hiddenSlideMode` option, whose default remains `'show'`.

The direct reader shows an embedded OLE object, such as an Excel or Graph
chart, as the presentation picture the file stores for it: an OLE shape is a
picture frame whose `pib` names the BLIP to display (MS-ODRAW 2.2.40 and
2.3.23.5), resolved through the shape's `ExObjRefAtom` to the document's
external object list (MS-PPT 2.7.7 and 2.10.1). The object storage is never
read or activated. PowerPoint's own PDF exports show that stored picture
unchanged for embedded objects drawn as content. Icon or thumbnail aspects,
linked objects and ActiveX controls, pictures without a supported BLIP, and
unsupported fill placements are rejected instead of being drawn without them.

PowerPoint displays GIF data that a producer stored in a PNG picture slot,
so the direct PPT reader identifies such a slot by its GIF87a/GIF89a
signature and emits it as `image/gif`; other mismatched content stays
rejected.

Classic linear gradients (MS-ODRAW `msofillShade` and `msofillShadeScale`)
keep their ordered shade colours, signed focus and 16.16 angle. Explicit
shade-array stops at either endpoint take precedence over scalar front or
back colours; a missing final endpoint uses the scalar back colour. Linear,
scaled, two-colour and translucent shades are projected, including on rotated
shapes and inside rotated or flipped groups. Path and title shades, other
shade types, custom fill rectangles and opacity combined with a shade-colour
array fail closed. Twenty controlled Office comparisons covered focus, angle,
endpoint conflicts and leaf flips; the projected positions were within one
DrawingML position unit of Office's serialized integers.

Pattern fills on unrotated shapes become tiled picture fills that follow
PowerPoint's own output: the 8x8 area of the stored 10x10 pattern bitmap,
one pattern pixel per point, white pixels in the fill colour and black
pixels in the background colour. Supported placements retain their shape transforms. Unsupported bitmap
sizes, translucent pattern colours and unprojected texture placements
remain rejected.

Picture colour settings follow how PowerPoint itself reads the binary
properties when it saves a binary deck as PPTX: "Black and White" becomes
DrawingML `grayscl` plus `biLevel` at 50%, and a transparent colour becomes a
`clrChange` to the same colour with zero alpha. The presentation renderer
applies these blip effects in document order for PPTX files as well.
Brightness/contrast (washout) on a picture becomes DrawingML `lum`: bright is
the brightness over 0x8000, and contrast is k - 1 or 1 - 1/k for the stored
16.16 slope k. PowerPoint writes exactly these values when it saves the
binary deck as PPTX. The renderer's `lum` formula reproduces PowerPoint's PDF
of a gray ramp under a bright × contrast grid, from both PPTX and .ppt, to
within 0.7 of 255. Brightness/contrast combined with a transparent colour or
black-and-white (whose order has no evidence), grayscale or black-and-white
alone, recolouring and adjustments on picture fills stay rejected until
Office output confirms their rendering.

A modern Office-saved PPT can retain paragraph properties in the OfficeArt
`metroBlob` alternative shape XML rather than in its classic text ruler. The
direct PPT source adopts that XML under the rule described below.

## Current implementation boundary

The repository contains only the direct readers; legacy input is never
converted to OOXML. The Rust crate `legacy-office-converter` builds one reader
per feature; this package has `direct-ppt`. Building it for `wasm32` without
it is a compile error. The `inspection` feature adds a native-only PPT
inspection example, and `fuzzing` exposes the fuzz entry points.

The local direct-render survey described below renders every installed
legacy sample through its source and pairs it with the Office-exported PDF.
The corpus is deliberately not redistributed. Broader binary-record coverage,
visual fidelity against Office, fuzzing and resource measurements remain part
of [issue #1472](https://github.com/yukiyokotani/office-open-xml-viewer/issues/1472).

## Best-effort fidelity evaluation

Loading without an error is a smoke test, **not reader completion**. The
target is useful best-effort preservation of the binary input's content and
display, with missing content and visual differences explicitly reported.
Pixel equality is not required for each incremental improvement. Pairing a
legacy file with its original OOXML is useful for investigation, but does not
prove fidelity: saving to an old format can itself change or remove features.
Use Office opening the actual legacy file as the visual reference. Office's
upgraded OOXML is useful for mapping binary records to XML, but Office's own
conversion can change layout and is not an absolute visual oracle. In
particular, rebuilt or down-saved corpus members must not silently be treated
as lossless copies of their original OOXML.

Compare the direct reader's Canvas output with Office-exported PDFs through
the local direct-render survey. Keep renderer self-regression tests against the
previous renderer separate from this fidelity comparison. Neither whole-corpus
Office equality nor Canvas display equality has been reached.

Treat the rendered model as a derived view of the original binary, keep the
binary as the authoritative source, and gate production use on a corpus
representative of the documents being ingested.

## Alternative shape XML

Modern Office can retain DrawingML in an OfficeArt `metroBlob`
([MS-ODRAW] §2.3.4.41). The source uses three outcomes: a package with no
alternative or a verified disagreement uses the binary projection; a verified
consistent alternative is adopted; an unreadable or unverifiable alternative
rejects the load. A checksum-only package does not require a theme.

Adopted alternatives retain their formatting and resource references, while
characters, identifiers and transforms come from the binary. Theme style
references must resolve to existing entries. Compatibility comparisons cover
only the properties documented beside the implementation; edits to other
properties can leave a stale alternative undetected. This is an experimental
compatibility boundary, not a claim of complete PowerPoint equivalence.

Per-entry and cumulative budgets bound input, records, decoded themes, text,
model storage and emitted JSON. Repeated hyperlink targets and split text
styles consume the model budget before they are copied. These ceilings are
library resource policy, not Office format restrictions.

The private Office corpus and its visual survey remain local development
material. Current completion requires fresh parser-backed tests, renderer
self-regression checks and separately adjudicated Office fidelity evidence.
Historical survey counts do not establish the current release's behavior.
