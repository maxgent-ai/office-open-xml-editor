# Paired DOCX / legacy DOC model diagnostics

Run this tool on models acquired from a proven Office export pair. It compares
retained parser facts and reports differences for investigation. It does not
approve rendering fidelity, prove that parsing retained every source fact, or
replace the previous renderer as the visual regression baseline.

```sh
node packages/legacy-converter/tools/legacy-doc-paired-model-compare.mjs \
  original-model.json exported-model.json report.json
node --test packages/legacy-converter/tools/legacy-doc-paired-model-compare.test.mjs
```

Each input is a model with a `body` array, or an envelope with exactly these keys:

```json
{
  "document": { "body": [], "section": {} },
  "resources": []
}
```

For image comparisons, put extracted resource records in the envelope's
`resources` array: `{ "path": "model-reference", "sha256": "64 lowercase hex digits",
"bytes": 123 }`, or `{ "path": "model-reference", "error": "extraction failed" }`.
Acquire and hash the actual resource bytes separately; the tool trusts this
inventory and never opens its resource paths. Raw models always take precedence
over envelope interpretation when their root has a `body` array.

The report separates ordered content, paragraph and run formatting, table facts,
notes, images, anchors, sections, document settings, and private acquisition
metadata. Contiguous equal text formatting is compared by UTF-16 offsets. Style
and table identities retain their equality relationships; note references retain
numbering facts while targeting the corresponding note. Cell margins use the
retained physical cell / row-exception / resolved-table precedence. Logical and
unsupported lexical margins remain unresolved; border conflict resolution,
font fallback, numbering, and layout are not reconstructed by this diagnostic.

`AGREEMENT_FOR_COMPARED_FACTS` applies only to the retained projected facts.
`DIFFERENT_FACTS_REQUIRE_CLASSIFICATION` requires source-side investigation.
`NOT_EXERCISED` means neither operand exercises that feature. Missing image bytes
or missing note targets produce `INCOMPLETE_RESOURCE_EVIDENCE` or
`INCOMPLETE_NOTE_EVIDENCE`, with `equal: null`. Other equality booleans refer only
to the projected facts; they are never an overall quality decision. Duplicate
note targets, duplicate resource paths, and partial parser models are rejected.

`rawDifferences` preserves all source-model differences. The separate
`representationDifferences` list is only the presence / identity subset, not a
complete explanation or a list of waived differences. Optional null / omission
and the normative paragraph snap-to-grid default are interpreted by the
projection; explicit false and zero survive. Unknown private fields retain their
raw identities, and unknown public subtrees remain inspectable in the full raw
ledger even if they use names recognized by the diagnostic projection.

Both input byte hashes are emitted to standard output. Bind those hashes to
immutable DOCX / DOC hashes, acquisition source and WASM revisions, resource
inventories, and observed Office export records in the caller's manifest. Prefer
an original DOCX / Word-saved DOCX / Word-saved DOC triad: a difference shared by
the two saves is useful evidence of a save or default-materialization change,
but cannot by itself prove its cause. Same-name files do not prove export lineage.
Record separately which exact inputs have Office-produced print PDFs.

The CLI admits regular JSON files up to 64 MiB, rejects changes during its bounded
read, limits traversal to depth 128 and one million work units, and caps each
ledger at 20,000 differences and the serialized report at 16 MiB. Reports with
changed sequence lengths preserve the whole affected sequence for review instead
of guessing alignment. Inputs exceeding these diagnostic budgets require a
separately reviewed investigation; the tool does not truncate them into success.
