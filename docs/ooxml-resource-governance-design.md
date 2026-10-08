# OOXML resource governance and bounded processing

Status: accepted direction for implementation on Draft PR #1120. The standard
archive defaults were retained after the recorded calibration in
[`ooxml-resource-default-calibration.md`](ooxml-resource-default-calibration.md).

## Context

OOXML packages are ZIP archives. Compressed input size is not a useful memory
bound: a small package can contain a much larger XML part, and parsing that part
can create an XML arena, a format model, serialized output, and renderer state
at the same time. Issue #1102 demonstrates this amplification with one large
SpreadsheetML worksheet.

The package boundary follows ISO/IEC 29500-1 section 9 (Packages), and the ZIP
mapping follows ISO/IEC 29500-2 section 7.3.6 (ZIP package limitations). Resource
budgets are implementation admission policy layered around that package model;
they do not redefine valid OOXML content or infer trust from ZIP declarations.

The existing 512 MiB `maxZipEntryBytes` guard only caps one inflated ZIP entry.
It does not bound aggregate inflation, parser work, model growth, serialization,
layout state, canvas allocation, or the copies made at worker boundaries.

## Reconciliation with concurrent pull requests

PR #1119 identifies the same package-level failure mode and contributes the
correct fail-closed requirements: per-entry and aggregate actual-inflation
limits, archive-entry admission, one retained budget across operations, worker
wiring, and package poisoning after a proven violation. This branch adopts
those requirements, but implements them in the shared `ResourcePolicy` ->
`ResourceGovernor` -> `PackageSession` layers so accounting, typed errors,
usage snapshots, generated defaults, bounded readers, and later pull operations
have one owner. The PR is therefore complementary prior art rather than a patch
to stack or copy; its currently conflicting format-local wiring is superseded
by the shared control plane here.

One proposed public-policy detail is intentionally not copied from #1119.
Aggregate declared inflated bytes are recorded for diagnostics but do not reject a package:
central-directory declarations are untrusted, and charging entries that a lazy
viewer never reads would make the public policy depend on unused media. The
public aggregate limit therefore uses actual bytes observed across distinct
visited entries; actual decompressor output remains authoritative.

Archive entry count uses two distinct admission layers. The public
`resourceLimits.maxArchiveEntries` policy defaults to the calibrated standard
limit, accepts a lower positive limit, and can be disabled with `null`. A
separate implementation hard ceiling remains non-configurable and cannot be
disabled. ECMA-376 Part 2 §7.3 and Annex B define the ZIP-backed OPC package
model and its format constraints, but do not prescribe either of these lower
resource-safety limits; both are implementation admission policy rather than
OOXML conformance rules.

PR #1124 closes a separate ownership leak: XLSX and PPTX workers survived some
rejected loads. This branch adopts that behavior through the shared
`disposeRejectedLoad` helper and applies the same idempotent cleanup contract to
DOCX, XLSX, and PPTX. Focused rejected-load and destroy tests cover both a
partially constructed instance and failure before an instance exists. Successful
loads and the existing Viewer interfaces remain unchanged.

## Goals

- Keep the existing DOCX, XLSX, and PPTX Viewer constructors, `load(source)`,
  navigation methods, callbacks, and load/render error behavior compatible.
- Reject resource-policy violations deterministically before an uncontrolled
  browser- or worker-level failure whenever the library can measure the risk.
- Use one resource-policy vocabulary and one error/report shape across all three
  formats and across main and worker rendering modes.
- Reduce avoidable peak memory by consuming inflated structured parts through
  bounded, back-pressured stages.
- Keep OOXML interpretation specification-first. Streaming must preserve the
  processed infoset, relationships, inheritance, and ordering required by the
  relevant ECMA-376 parts; it must not introduce a reduced parser dialect.
- Preserve structured diagnostic information across worker boundaries and make
  debug usage visible from the main context.
- Keep the existing synchronous Node parsing helpers as materializing
  compatibility APIs, while adding an options-aware session/pull surface for
  server-side rendering jobs that need deterministic resource policy and
  bounded in-flight work.

## Non-goals

- This design does not promise constant total memory for every document.
- It does not infer final inflated size from compressed size.
- It does not guarantee that every engine-level OOM, GPU allocation failure, or
  terminated browser process can be converted to a JavaScript error.
- Worker mode is containment and lifecycle isolation, not a hard browser-process
  memory sandbox.
- Streaming does not mean that DOCX can paginate an arbitrary page without
  processing preceding layout state.
- Existing `ArrayBuffer` input necessarily keeps a complete compressed package
  in memory. Bounded inflation and model processing do not change that fact.
- Existing synchronous Node helpers that return a complete public model cannot
  promise fixed total memory. Fixed-in-flight-memory server workflows use the
  additive asynchronous session/pull surface instead.

## Delivery milestones

Each milestone has an observable exit gate. Later milestones may start only
after their shared contracts are stable; format implementation work may proceed
independently once M2 is complete.

### M0 — Baseline and contract freeze

- Rebase the Draft PR on current `main` and record the behavior of Issues #1102
  and #1088 and Draft PRs #1119 and #1124.
- Freeze the Viewer compatibility boundary, failure routing, public policy
  vocabulary, accounting semantics, and the format-owned responsibilities.
- Preserve baseline public API snapshots and focused resource-limit tests.
- Build declarations from current source before comparing separate DOCX, XLSX,
  and PPTX public baselines with `scripts/check-public-api.mjs`; stale `dist/`
  output must never make this gate pass.
- Close the rejected-load worker ownership gap for all three formats.

Exit: this document is internally consistent, cleanup regression tests pass for
DOCX/XLSX/PPTX, existing public API checks pass, and no implementation question
can silently change the public contract.

### M1 — Shared policy and governor

- Replace format-local option interpretation with one immutable normalizer and
  one per-session `ResourceGovernor` in core/common layers.
- Implement the two public limits, deprecated alias reconciliation, practical
  defaults, hidden hard quotas, structured errors, usage snapshots, and
  poison semantics.
- Enforce raw archive entry count, selected-entry declared/actual inflation,
  and actual distinct-entry aggregate inflation without trusting declarations.

Exit: adversarial Rust and TypeScript tests cover forged sizes, actual overrun,
repeat reads, conflict validation, worker error transport, and poisoned sessions.

### M2 — Package session and bounded pull control plane

- Introduce the random-access `PackageSession`, bounded entry reader, correlated
  pull/ack/release protocol, credits, leases, cancellation, and idempotent close.
- Ensure logical-operation counters survive multiple pulls and that stale or
  unknown transferred resources are disposed.
- Keep complete buffered package input for the existing Viewer source contract;
  do not claim URL range loading in this milestone.
- Define the same correlated session lifecycle for browser workers and Node.
  Node does not need a Worker hop to use backpressure, but it must use the same
  policy, usage, error, cancel, close, and format-owned chunk contracts.

Exit: protocol tests prove bounded in-flight data, backpressure, timeout/abort
convergence, late-response disposal, and parity across the direct and simulated
worker transports. Production Worker/WASM parity is established separately as
each real format pipeline adopts the substrate in M3-M5.

The M2 boundedness claim is deliberately limited to this substrate. Existing
format bootstrap calls still materialize their historical complete models until
M3-M5 migrate them; wrapping such a result in one nominal chunk would not make
it bounded.

### M3 — XLSX bounded worksheet pipeline

- Preserve SpreadsheetML dependency resolution and Part 3 MCE semantics while
  moving worksheet XML through complete-row batches.
- Compose the format driver with the shared `PullSessionHost` /
  `BoundedPullSession` control plane; do not introduce a second XLSX-specific
  ACK, generation, cancellation, poison, or credit state machine.
- Treat every emitted row batch as provisional. Validate the worksheet tail,
  ZIP CRC, and ancillary parts before preparing the terminal row-free sheet;
  finish the package operation and publish/cache the result only after the
  terminal chunk is acknowledged.
- Preserve sheet-local placeholder degradation for ordinary worksheet read,
  XML, and CRC failures. Resource-policy violations remain fatal, poison the
  package, discard provisional rows, and can never become placeholders.
- Remove the avoidable full worksheet string plus full serialized model overlap
  from the Viewer path; keep `getWorksheet()` as an explicitly materializing
  compatibility adapter.
- Apply non-configurable retained-model and renderer-index ceilings to that
  compatibility path. Raising archive limits must not expose the Window to an
  unbounded cell/object/index amplification path.
- Add an explicitly owned Node workbook session: `openXlsxWorkbook()` retains
  the validated index and archive, `worksheetRows()` yields the same bounded row
  batches, and `close()` releases the session. Async `materializeXlsxWorkbookIndex`,
  `materializeXlsxWorksheet`, and `materializeXlsxWorkbook` consume that same
  owned path when caller-owned materialization is required. The browser compatibility adapter owns provisional-row
  rollback. Lower-level Node session consumers see those batches directly and
  must discard them if the terminal worksheet is a `parseError` placeholder.

Exit: synthetic large worksheets cross multiple pulls, render identically, stop
at deterministic limits, and show a measured reduction in transient peak usage.

### M4 — PPTX slide-granular pipeline

- Keep presentation/theme/master/layout dependencies shared while parsing and
  retaining slides and their relationships as format-owned units.
- Make navigation, media leases, cleanup, and resource failures consistent with
  the common session contract.
- Add a Node presentation session that opens shared dependencies once and pulls
  one slide/resource unit at a time; `materializePptxPresentation` consumes that
  producer, while image/media extraction remains owned by the open session.
- Give image and media paths one count- and byte-bounded raw OPC-part owner per
  realm. Decoded bitmap/GPU caches remain separate because they own a different
  representation and lifetime.
- Project Markdown by draining the canonical slide producer sequentially into a
  bounded UTF-8 writer. The returned string is necessarily materialized, but a
  full `Presentation` must not coexist with it and an output crossing must be a
  typed serialization limit rather than an allocator failure.
- Preserve the pooled-canvas recycling contract from PR #1127: returning a
  slide canvas to the pool clears its zoom-derived CSS height before reuse.

Exit: multi-slide tests prove on-demand unit ownership, navigation compatibility,
bounded transient retention, and cleanup on rejection, reload, and destroy.

### M5 — DOCX sequential layout pipeline

- Move every DOCX package read onto the shared `PackageSession`. Keep Issue
  #1088's recoverable containment only for an optional object, isolated block,
  or non-primary story whose omission is structurally proven safe. Corrupt ZIP
  structure, a missing or malformed required XML part, invalid geometry or
  ownership, non-convergence, and resource-policy violations remain fatal; a
  resource violation also poisons the package and wins over optional-part
  fallback.
- Read `word/document.xml` through a two-pass bounded cursor. The first pass
  retains only the compact section, table-adjacency, and content-control plan
  needed by later blocks. The second pass converts one complete logical body
  block at a time through the existing specification-first semantic parser.
- Feed those blocks into a sealed, replayable layout-source store in the realm
  that owns layout, then use the single immutable acquisition -> normalization
  -> layout -> paint pipeline. The full-model compatibility adapter and the
  streamed Viewer adapter must converge on that same store; no streamed-only or
  legacy paginator is permitted.
- Keep the store contract independent of `DocxDocumentModel` and parser-owned
  object identity. It owns canonical `SourceRef`-keyed paragraph, table, and
  story records; resolved section, note, field, and font facts; and immutable
  image/math/paint manifests. A model adapter projects parser-private facts into
  those records before sealing. The streamed Viewer path separates the public
  and builder-owned graphs one bounded pull at a time, then destructively
  canonicalizes the builder-owned graph into the same store without constructing
  a third complete body graph. The mutable public compatibility model lives
  outside the store and cannot change sealed layout or paint inputs.
- Keep pagination after the store is sealed. Sequential section inheritance,
  fields, notes, bookmarks, convergence, total page count, and stable Viewer
  readiness do not permit random page access or final page metadata before the
  complete required part has been validated.
- Let Node build and paginate the same sealed source through an asynchronous
  document session. `materializeDocxDocument` consumes the canonical coordinator
  into a caller-owned model; server rendering can pull completed page render inputs only after the same Viewer readiness
  barrier, without a browser Worker and without a second semantic pipeline.
- Preserve source compatibility for lower-level APIs that synchronously expose
  a complete document model by materializing their stream. Such adapters do not
  receive a bounded-retention claim. The self-loaded Viewer path must avoid the
  simultaneous full document XML, whole-document XML arena, Rust model, and
  monolithic JSON representations. The streamed builder owns its accumulated
  records and seals them in place; sealing must not clone the complete retained
  logical source and create a second whole-document peak.
- Align recoverable parse/layout containment with Issue #1088. Resource limits,
  malformed required XML, invariant failures, and non-convergence remain fatal;
  already-produced chunks are never promoted to partial success after them.

Exit: pagination and visual behavior remain stable, recoverable failures yield
the defined partial result, resource failures are deterministic, synthetic
documents cross multiple acknowledged pulls, the Viewer path has no
whole-document XML or JSON materialization, measured transient retention tracks
the largest bounded unit rather than total document XML, and the DOCX
architecture audit passes.

### M6 — Containment and Ratatui-inspired diagnostics

- Normalize options before worker creation, terminate monolithic timed-out work,
  and converge all failure paths on idempotent disposal.
- Implement `debug: true` checkpoints and one polished, color-free final console
  card using a pure shared presentation model and a shared TUI console emitter.
  Browser DevTools receive typography-only `%c` styling to preserve the Unicode
  grid; Node and Worker consoles receive one plain argument.
- Document Worker mode as stronger lifecycle containment, not a memory sandbox
  and not the default.

Exit: no rejected load leaves an owned worker or transfer alive; debug output is
snapshot-tested, content-free, and equivalent across format and execution mode.

### M7 — Calibration and release-quality verification

- Measure public and synthetic documents across formats and the shared
  browser/WASM counter path, then adopt or revise the candidate defaults.
- Run Rust, rebuilt-WASM, focused/full TypeScript, typecheck, build, public API,
  visual, and high-water verification appropriate to the touched surfaces.
- Publish the limits as admission policy, not as a promise of exact memory use.

Exit: defaults have recorded evidence and acceptable false-rejection behavior;
all mandatory checks pass or any environment-only limitation is documented with
a reproducible local command.

### M8 — Independent critical review and Draft handoff

- Obtain independent GPT-5.6 Sol reviews of OOXML specification fidelity,
  responsibility boundaries, duplicate logic, API consistency, error semantics,
  and the claims made about bounded processing. Fable is intentionally excluded
  because it is unavailable for this delivery.
- Fix every accepted finding and re-run the affected gates.
- Push the reviewed branch and update Draft PR #1120 without merging it.

Exit: review findings and dispositions are recorded, the branch is clean and
reproducible, and the user can validate the Draft locally before any merge
decision.

## Compatibility boundary

The public Viewer surface is the compatibility boundary:

> Version note: the error-delivery statements and table in this section record
> the 0.76 contract in effect when resource governance was introduced. Version
> 0.77 supersedes that contract: every awaitable Viewer operation rejects on
> failure regardless of `onError`. See
> [the 0.77 error migration](migration-viewer-errors-0.77.md).

```ts
new DocxViewer(canvas, options).load(source)
new XlsxViewer(container, options).load(source)
new PptxViewer(canvas, options).load(source)
```

The existing `string | ArrayBuffer` sources remain accepted. Existing options
remain accepted. New resource and debug options are additive. A successful
`load()` resolves at the same user-visible readiness point as before, and the
existing `onError` versus rejection behavior remains unchanged.

For DOCX, “fatal required-part failure” means fatal to the document transaction:
no body batch or completed page is promoted to a partial success. It does not by
itself redefine the established Viewer routing for a fatal-document diagnostic
versus rejection. Resource-policy violations are different: they always reject
with `OoxmlResourceLimitError`. This preserves the external error route while
making Issue #1088's fatal/recoverable ownership boundary explicit.

The following are internal implementation contracts and may change:

- worker request/response envelopes;
- wasm-bindgen parser exports;
- parser-to-renderer model representation;
- JSON versus binary/chunk transport;
- archive-handle lifetime and caching;
- parsing and rendering batch sizes.

Lower-level `DocxDocument`, `XlsxWorkbook`, and `PptxPresentation` APIs should
remain source-compatible where a facade can preserve them without defeating the
resource design. They are not allowed to dictate the worker wire representation.
When a lower-level compatibility method explicitly returns or exposes a complete
model, its adapter may necessarily materialize that complete model; bounded
retention must not be claimed for that path.

New failures follow the existing ownership boundary:

| Failure | Direct factory | Viewer without `onError` | Viewer with `onError` |
| --- | --- | --- | --- |
| invalid resource option | `load()` rejects before worker creation | `load()` rejects | callback receives the error and `load()` resolves |
| detected resource limit | `load()` rejects with typed error | `load()` rejects | callback receives the error and `load()` resolves |
| worker crash during load | `load()` rejects and disposes the partial engine | `load()` rejects | callback receives the error and `load()` resolves |
| recoverable DOCX content/layout failure | partial-result contract defined by Issue #1088 | same partial result | same partial result plus its structured diagnostic |

Viewer construction remains non-throwing for stored load options. Validation is
performed when `load()` starts, before a worker or WASM archive is created.

## Public policy

The long-term public option is a plain object. Callers do not need a policy
factory.

```ts
export type OoxmlResourceLimit = number | null;

export interface OoxmlResourceLimits {
  /** Actual inflated bytes for any one archive entry, including media. */
  maxArchiveEntryBytes?: OoxmlResourceLimit;

  /** Sum of actual inflated bytes across distinct entries in one session. */
  maxTotalInflatedBytes?: OoxmlResourceLimit;

  /** Archive central-directory entries admitted before index allocation. */
  maxArchiveEntries?: OoxmlResourceLimit;
}

export interface LoadOptions {
  resourceLimits?: OoxmlResourceLimits;
  onResourceMetrics?: (metrics: OoxmlResourceMetrics) => void;
  debug?: boolean;

  /** @deprecated Use resourceLimits.maxArchiveEntryBytes. */
  maxZipEntryBytes?: number;
}
```

Semantics:

- `undefined` selects the library's standard default.
- A positive safe integer overrides that default. Byte-limit fields use bytes;
  `maxArchiveEntries` uses an entry count.
- `null` disables that configurable policy limit only. It does not disable
  non-configurable hard safety quotas.
- Invalid values in `resourceLimits` reject before a worker is created.
- The deprecated `maxZipEntryBytes` adapter preserves its historical input
  behavior and remains an all-entry limit, including media. If callers supply it
  together with `maxArchiveEntryBytes` and the two supplied values disagree,
  option normalization rejects instead of applying hidden precedence.
- Standard defaults are compatibility policy and may be revised deliberately
  between releases. Hard quotas protect implementation invariants and are not a
  supported tuning surface.

`maxTotalInflatedBytes` is stable under lazy loading and internal batch changes:
the session records the greatest actual inflated size observed for each distinct
entry and sums those maxima. Re-reading an entry does not consume this public
budget again. Repeated-inflation work and per-structured-part amplification are
measured separately and protected by internal operation/unit quotas.

The calibrated standard defaults are 128 MiB per archive entry, 256 MiB total,
and 4,096 archive entries; the evidence and its corpus limitations are recorded separately. They
are not memory guarantees. Adopting defaults below the old
512 MiB per-entry default intentionally narrows the set of documents that load
without overrides: source compatibility is preserved, but behavioral
compatibility for documents above the new defaults is not.

The approximately 267.7 MiB inflated worksheet reported in Issue #1102 is
therefore rejected by the standard per-entry default with a typed, catchable
resource error before the browser attempts the historical monolithic model.
That is the intentional safe default, not a claim that the document is invalid.
A caller may raise the policy limit after measuring its environment. The
materializing `getWorksheet()` compatibility path can still retain memory in
proportion to all returned cells; only the bounded pull/session path may claim
bounded in-flight row payloads, and neither path promises a fixed process heap.

Raising or disabling a configurable ZIP limit must not remove internal model
safety. M3 therefore calibrates non-configurable retained-worksheet ceilings;
the initial candidates are 250,000 cells, 100,000 row objects, 32 MiB of owned
UTF-8 string content, and 64 MiB of exact monolithic worksheet JSON. A workbook
compatibility cache initially admits 500,000 cells and 200,000 rows across
successfully committed unique sheets. These are logical implementation quotas,
not heap estimates, and apply only to paths that retain complete worksheets;
the Node workbook session's `worksheetRows()` stream may process more total rows
while remaining subject to package, per-unit, and in-flight limits. Crossings
are typed, non-configurable resource errors and are never converted to a
degraded sheet.

Renderer-derived indexes have an independent initial 250,000-entry ceiling.
Merge and styled-table ranges can expand into many coordinates even when the
worksheet contains few cells, so they are checked by range-area arithmetic
before allocation. The renderer must also reuse one cell lookup for ordinary
painting and conditional formatting rather than retaining duplicate keyed maps.
These hard values may be revised only from recorded browser/WASM evidence.

The defaults and hard archive ceilings have one language-neutral source
in `packages/ooxml-common/resource-policy.json`; generated TypeScript and Rust
constants are checked in CI so browser option normalization, parser-native
fallbacks, and effective hard caps cannot drift during calibration or later
releases. Generation rejects a default that exceeds its hard ceiling.

ChartEx has a non-configurable 524,288-element per-document allocation budget
from the same generated policy. Calibration measured a maximum of 2,454
parser-owned data-vector slots in any one document across 22,973 local ChartEx
parts, leaving more than 213 times headroom; this is availability policy, not a
document-format constraint. Each sparse numeric/string cache, derived category,
value, hierarchy and box-and-whisker bin vector is charged immediately before
allocation. Separate ChartEx parts and lazy operations share the owning package
governor, including DOCX native and streaming paths. A declared `ptCount` or
implicit point index beyond the allocatable range is rejected monotonically,
including the `unsignedInt` maximum. The count is neither compressed ZIP size
nor exact allocator bytes; formula-only source resolution remains under the
host's existing range policy. A crossing poisons the package with
`chartex-allocation:elements` at the parsing stage; it cannot silently omit a
series or select a DOCX picture fallback. No migration is required.

Archive entry count, XML nesting, relationships, model complexity, serialized
bytes, image dimensions, canvas pixels, and timeouts remain internal quotas or
separate existing options unless evidence shows that users can tune them
meaningfully. A resource error reports whether the violated limit was
configurable.

## Accounting model

Resource accounting belongs to a per-document package session, not to parser
call sites or renderer-specific helpers.

The design distinguishes the following concepts instead of forcing them into one
counter. The public `OoxmlResourceMetrics` snapshot currently reports source
bytes and package-governor archive counters; items described as hard guards are
enforced but are not presented as measured heap usage:

- compressed input bytes observed by the loader;
- central-directory entry count and declared expanded sizes;
- actual inflated bytes per archive entry;
- distinct-entry inflated total used by the public session limit;
- actual inflated bytes per structured part and indivisible parser unit;
- actual bytes delivered during each operation, counting repeated reads again;
- hard serialized/model projection ceilings at ownership boundaries;
- the observable WASM failure boundary (not allocator attribution or a stable
  public linear-memory counter);
- hard image dimensions, per-image pixels, active/cached decoded RGBA ownership,
  and decode concurrency before or around browser decoder allocation.

Those decoded-byte guards cover raster/ImageBitmap surfaces. Browser-managed
SVG/vector parse and decoded storage has neither a portable byte measure nor an
explicit close primitive. SVG caches are count-bounded and release their object
URLs and references, but that residual browser allocation cannot be charged to
the RGBA ownership counter.

Declared ZIP sizes are attacker-controlled. A selected entry declaration above
its per-entry limit is sufficient for early rejection, but a declaration below
a limit is not proof of safety. The whole archive's declared total is recorded
for diagnostics; it is not charged to the public distinct-inflation limit before
those entries are visited, because unused lazy media must not consume an
actual-byte session budget. Actual output is checked while it is inflated, using
`limit + 1` semantics where necessary to distinguish exact completion from
truncation.

After a proven package-policy violation, that package session is poisoned: no
later operation may continue reading from a partially trusted archive.

A logical operation may span many protocol pulls. Its internal work counters
live in the session ledger until that operation completes or aborts; they must
not reset at each WASM export or each chunk request.

## Internal responsibilities

### `ResourcePolicy`

An immutable normalized policy created before worker construction. It merges
standard defaults, the deprecated compatibility adapter, and caller overrides.
It contains no counters and performs no I/O.

### `ResourceGovernor`

The sole owner of counters, limit checks, high-water marks, the first violation,
and the final usage report. Parser and renderer code report observations to it;
they do not independently interpret public options.

### `PackageSession`

Owns the source bytes or random-access source, ZIP archive/index, governor, and
part handles. It provides bounded entry and structured-part readers, prevents
reads after abort/poison/close, and releases all retained archive resources
idempotently.

The input abstraction is random-access rather than a linear `ReadableStream`,
because ZIP discovery requires the end-of-central-directory record and seeks.
The initial implementation retains the complete compressed source, matching the
existing API. Range-backed URL input is a later independent optimization and
requires response-version pinning plus an owned entry reader; it is not a
prerequisite for bounded inflated-part processing. Agile-encrypted CFB input
remains on the buffered path until the crypto/container layer is redesigned.

The buffered implementation moves the input `Vec<u8>` into shared ownership
without copying its backing allocation. The ZIP crate validates the central
directory and resolves each entry data range; an owned Stored/Deflate decoder
then survives across pulls, incrementally verifies size and CRC, and never
restarts inflation from byte zero. This matches the compression methods enabled
by the repository's existing ZIP dependency. Unsupported or encrypted ZIP-entry
methods remain ordinary container errors rather than policy violations.

### Pull protocol

Consumers request work instead of receiving an unbounded push:

```text
open source -> manifest/session
pull resource chunk with credit -> data + usage checkpoint + leases
consume/acknowledge -> release transient chunk and leases
pull next chunk
complete or abort -> final report + close
```

Credit is a hard bound on the measured wire payload produced for one pull; it
is backpressure, not a CPU-time or security policy. A Deflate decoder may
consume more compressed input internally before producing the credited output,
so worker termination remains the containment mechanism for a decompressor that
does not yield. Resource limits remain independently enforced. A format unit is
never emitted above the caller's credit. If one indivisible unit cannot fit, the
pull fails deterministically with an insufficient-credit or hard-unit-limit
error instead of exceeding the bound or deadlocking the stream.

At most the internally allowed number of chunks may be in flight. Abort,
timeout, resource violation, parser trap, and explicit destruction all converge
on the same idempotent session close path. Worker-retained payloads carry leases
that must be released. Transferred objects in stale or unknown responses, such
as `ImageBitmap`, are disposed rather than merely ignored.

The protocol uses structured messages internally; it is not exported as a
Viewer API. Large payloads use transferable buffers where crossing a worker
boundary is unavoidable.

Every command and response carries separate correlation, session, operation,
and generation identities. Pull, acknowledgement, retained-lease release,
cancel, and close are correlated. A final data-bearing chunk is acknowledged
like every other chunk. Transferred payload ownership and worker-retained
numeric leases are distinct: acknowledging releases producer staging,
disposing releases main-side transfer resources, and lease release frees only
the explicitly retained worker resource.

The outer protocol standardizes correlation, sequencing, cancellation, usage,
leases, and close behavior. Format payloads and stream names remain opaque to
it. There is intentionally no common `ModelChunk<Row | Slide | Page>`.

## Format-specific processing

### XLSX

Workbook metadata, relationships, styles, theme, and shared strings are resolved
before dependent sheet rows. `sheetData` is consumed as complete row batches
while preserving ECMA-376 Part 3 Markup Compatibility processing. The renderer
or compatibility facade acknowledges each batch before another is produced.

The production worksheet projector reads the owned ZIP entry incrementally and
emits provisional complete-row batches without retaining a full worksheet string
or XML tree. The lower-level `getWorksheet()` compatibility adapter drains that
stream into a complete `Worksheet`; its retained rows, cells, string content,
serialized model, cache, and renderer indexes therefore have separate hard
ceilings. The explicitly owned Node workbook session can instead consume row
batches without retaining the complete worksheet model.

### PPTX

Presentation metadata, themes, slide masters, and layouts form the shared
dependency set. Slides and their relationships/media are independent processing
units after those dependencies are available. A slide may be parsed and cached
on demand without changing Viewer navigation semantics.

### DOCX

Styles, numbering, theme, settings, relationships, headers/footers, notes, and
other referenced stories must be available when their body content is
normalized. Body blocks can flow sequentially into the single immutable
acquisition -> normalization -> layout -> paint pipeline, releasing raw XML and
temporary parser arenas behind the layout frontier.

Pagination remains sequential because preceding layout determines later page
positions and total page count. Completed page/layout state, cross-references,
fields requiring total pages, and content needed by public document operations
may remain retained. Recoverable object/block failures and later layout failures
must follow Issue #1088's explicit containment boundaries rather than being
silently defaulted or routed through a legacy layout path.

The current Viewer readiness contract requires pagination to finish so page
count, per-page sizes, and bookmark destinations are stable. Streaming reduces
transient parser/model overlap but does not remove that readiness barrier.

## Abstraction boundary

The following are real shared abstractions:

- correlated pull-session lifecycle;
- resource ledger, reservations, and leases;
- random-access package source and bounded ZIP entry reader;
- timeout, cancel, close, error wire, and late-transfer disposal;
- common debug report/view rendering.

The following remain format-owned:

- semantic chunk types and batch boundaries;
- dependency graphs and readiness barriers;
- retained model/layout stores and eviction policy;
- DOCX pages, PPTX slides, and XLSX row batches/materialized worksheet caches.

Sharing the control plane must not create a second generic OOXML semantic model.

## Failure contract

A detected policy violation throws one `OoxmlResourceLimitError` family across
all formats and modes. Its details use one stable record with extensible
`resource` and `metric` string axes, rather than a closed public union that must
change whenever an internal hard quota is added. Known literals retain editor
completion, while newer worker literals remain decodable by an older host. An
archive-level entry-count violation therefore does not need a dummy part name.
The declared package total remains diagnostic rather than a rejection metric
because declarations are untrusted and lazily unvisited entries must not consume
an actual-byte policy. Shared fields include:

- stable code and stage;
- resource/metric identifier;
- limit and observed value;
- safe OOXML part address only when applicable;
- operation/format identifier;
- whether the limit is configurable;
- the last complete usage snapshot.

Worker serialization reconstructs the real error subclass on the main side.
Messages and debug reports never include document text, passwords, URLs, sheet
names, or local filesystem paths.

A residual failure that reaches a recognised trap-shaped boundary remains
`parser-crashed` unless the governor had already proved a specific limit
violation. In the current `panic = "abort"` build, Rust panic, allocation failure,
explicit `unreachable`, and stack overflow can terminate an export without
returning a typed `Result`; the WebAssembly JavaScript boundary may expose those
causes through the same `WebAssembly.RuntimeError` or another
implementation-dependent native error. The final trap reports that execution
stopped, not which Rust path caused it, so the discarded cause cannot be
reconstructed reliably afterward.

A panic hook may emit a debugging message before some Rust panics, but it is a
console side channel rather than a structured return value. Allocation aborts
need not pass through it, logging can itself fail while memory is exhausted, and
message/stack formats vary by toolchain and runtime. Error-class or message
sniffing must therefore not publish `parser-oom`: misclassifying a parser defect
as a large-file condition would hide actionable bugs. The stable contract is:

- a governor-observed crossing returns `OoxmlResourceLimitError` before the
  dangerous operation proceeds;
- a trap whose cause was not preserved remains `parser-crashed` and poisons the
  affected WASM instance when it reaches a recognised trap-shaped boundary;
- a future reliable `parser-oom` would require pre-trap structured cause
  retention across allocation paths, not post-trap inference.

The [WebAssembly JavaScript Interface error mappings](https://www.w3.org/TR/wasm-js-api-2/#error-condition-mappings-to-javascript)
permit stack exhaustion or OOM to surface through implementation-defined errors,
including an indistinguishable plain `Error`, or to terminate the process. The
guard recognises standard trap shapes plus observed `InternalError` and
`OOMError` names, but it cannot safely classify a plain `Error` without confusing
it with a graceful parser `Result::Err`. Conversion to `parser-crashed`, instance
poisoning, and recovery are therefore guaranteed only for recognised catchable
trap-shaped failures, not for every engine-level failure.

Canceling or timing out a legacy monolithic synchronous WASM call cannot stop it
cooperatively: the worker cannot process another message until the call returns.
During migration that path terminates the worker. Cooperative cancellation is
only claimed once the operation is divided into bounded pull exports.

Viewer reload intentionally keeps the old engine alive until the new load
succeeds. That compatibility behavior creates a temporary two-session memory
window. A per-session limit does not claim to bound this combined process peak.

## Debug reporting

Measurement and presentation are separate. `onResourceMetrics` receives a
versioned, machine-readable, content-free report without console output;
`debug: true` sends the same report to the built-in console renderer. Neither
changes limits or parser behavior, and observer failures are isolated from load
semantics. Workers send structured usage checkpoints to the main context. The
main context retains checkpoints quietly and, after option validation, emits
one report when the document/workbook/presentation factory succeeds or fails,
including failed loads for which no engine is returned. Viewer constructors use
that same engine-load scope: the report does not wait for first canvas paint,
because render errors are non-fatal in the existing Viewer contract and scroll
viewers paint asynchronously. The initial callback report is therefore not a
final package-session total: first paint or later lazy sheet, slide, image,
font, or other part access may increase usage or surface a separate render
error. Every successful browser engine and Viewer exposes the same asynchronous
`getResourceMetrics()` method; it probes the current package governor and
returns a new immutable snapshot. Collection for that method is always active,
while `debug` controls only console presentation. Bounded Node DOCX and
PPTX sessions report successful terminal metrics when their one-pass stream is
exhausted or the session is explicitly closed. XLSX can consume multiple
worksheets sequentially, so it reports success only when the reusable workbook
session is explicitly closed. Open-time and session-operation failures report
immediately.

The presentation is Ratatui-inspired: bordered blocks, compact rows, and gauges.
One pure report model renders deterministic Unicode. In browser DevTools, a
shared emitter applies only fixed-width typography, size, line height, disabled
ligatures, zero letter spacing, and preserved whitespace through `%c`; it never
sets foreground or background colors. Node and Worker consoles receive one plain
argument without CSS or ANSI escapes. The complete representation is
snapshot-tested. Console rendering and resource measurement remain separate
responsibilities. Reports include the largest
actual inflated entry specifically so `maxArchiveEntryBytes` can be calibrated,
alongside distinct session bytes for `maxTotalInflatedBytes`.

## Verification gates

- Public Viewer type and behavioral compatibility tests for all three formats.
- Declared-size forgery, actual overrun, entry-count, aggregate/operation, repeat
  read, poison, abort, timeout, and cleanup tests.
- Main/worker parity for successful results, typed errors, partial DOCX results,
  and final metrics.
- Specification-focused tests for streamed MCE and format dependency behavior.
- Synthetic large documents whose memory amplification can be varied without
  private fixtures.
- Peak/high-water benchmarks that distinguish input, inflate, parser/model,
  serialization, worker transfer, layout, and rendering stages.
- Rust tests and clippy, rebuilt WASM, focused and full TypeScript tests,
  typecheck, build, public API checks, and the DOCX architecture audit.
- Independent GPT-5.6 Sol reviews followed by fixes and re-verification.

The Draft PR is not merged until these gates are satisfied and the user has
reviewed the result locally.

## M8 review disposition

Fable was unavailable for M8, so the final review used two independent
GPT-5.6 Sol tracks: architecture/API consistency and OOXML specification/safety.
Both tracks initially requested changes and approved the implementation only
after the findings were fixed and re-verified.

The review fixes ensure that nested DOCX content is projected from the Part 3
MCE processed infoset, ordinary missing or malformed required parts preserve the
existing degraded-document route, resource-policy failures remain typed and
terminal, debug telemetry is bounded and non-fatal, and DOCX pull lifecycle
logic has one state machine. The streamed DOCX path retains the public model and
the immutable layout source but no third reachable document-scale paragraph
snapshot graph.

The final pass also aligned the three explicit Node session lifecycles, separated
#1119's adopted invariants from its intentionally unexposed tuning knobs, and
narrowed WASM recovery claims to recognised catchable trap-shaped failures.
Observed implementation-defined `InternalError` and `OOMError` names poison the
guarded instance without being mislabeled `parser-oom`; indistinguishable plain
errors and process termination remain outside the recovery guarantee.

Accepted non-blocking risks are explicit: the cloning and destructive
canonicalizers use different ownership strategies and are kept semantically
aligned by equivalence tests; physical collection of an abandoned worker graph
after vertical-layout fallback is controlled by the JavaScript runtime; and a
timed-out final debug probe reports the last complete checkpoint rather than
changing successful load semantics.
