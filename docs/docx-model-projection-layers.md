# DOCX model to layout boundary

The DOCX layout path has one public compatibility model and one immutable layout
source. The table records the cost in terms of a document with `N` body and story
nodes, `B` body blocks, and `S` section occurrences. It distinguishes a traversal
from a retained full-document copy.

| Stage | Work and cost | Consumer and reason to remain |
| --- | --- | --- |
| `packages/docx/model` and parser JSON | Rust resolves OOXML and serializes one `O(N)` JSON value. | The worker or main-thread loader parses it. This is the wire and public-model contract. |
| `src/types.ts` | Type declarations only; no runtime allocation. | Public API and parser-model type checking. |
| `src/parser-model.ts` normalization | One `O(N)` traversal. Caller-owned models use copy-on-write for changed run ancestry; exclusively owned models normalize in place. It preserves unavailable-drawing and revision sidecars, assigns math source keys, and resolves section placement. | The compatibility model and layout acquisition. Parser-only sidecars and public hand-built models require this boundary. |
| `src/layout-source-model-adapter.ts` | Shallow root, section, and header/footer defaults for hand-built public models, followed by one `O(N)` acquisition traversal and an immutable block repository. Caller-owned input detaches retained blocks from the mutable public model, copying at most one table at a time; exclusively owned input replaces blocks in its own graph. | Layout, paint resources, and the public compatibility model. The two graphs must have separate ownership so public mutation cannot change retained layout. |
| `src/layout/acquisition-input-projections.ts` | Type-only capability record; no model copy. Individual paragraph and table projections allocate the retained facts they actually use. | Parser-independent acquisition and the source store. These are semantic projections of parser-private facts, not another document model. |
| `src/parser-model.ts` body and section acquisition | Section indexing snapshots `O(S)` section records and referenced header/footer stories because it is also an independently consumed immutable contract. Body sequence acquisition traverses `O(B)` blocks and now streams adjacent-table grouping into one sequence array. The acquisition record is transient and makes no second deep copy. | `layout/body-layout-input.ts` resolves section owners. Grouping validates parser-owned logical table membership. |
| `src/layout/body-layout-input.ts` | Resolves `O(S)` section contexts and `O(B)` sequence entries, including continuous-section roles, in one sequence pass. Its final `snapshotPlainData` validates, freezes, and detaches the retained plain input in one traversal. | The source store and paginator. This is the trust and immutability boundary for the body input. |

There is no same-name `src/body-layout-input.ts` bridge. The parser JSON and public
API shape are unchanged. The final plain-data snapshot stays in TypeScript because
it protects the retained layout from mutable, hand-built public models; Rust serde
defaults cannot establish that JavaScript ownership guarantee.
