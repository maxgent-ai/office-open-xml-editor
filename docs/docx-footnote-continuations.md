# Page-bottom footnote continuation

Paragraph footnotes continue across physical pages by default in
`DocxDocument.load`, `DocxViewer`, `DocxScrollViewer`, the Node
`openDocxDocument` session, and documents loaded through `modelSources` in
either environment. This is a library display default chosen to follow
Word's page allocation in the finite controls described below; it is not a
normative ECMA-376 algorithm and does not promise identical Word pagination or
paint. Passing `allowFootnoteContinuation: true` behaves the same as omitting it.

Migration: to keep the previous whole-note pagination, pass
`allowFootnoteContinuation: false`. That mode is unchanged: notes stay whole on
the page of their reference, with its existing limits. Callers that previously
omitted the option now get continued notes, which can change line placement and
page counts. They can also now see the continuation diagnostics described at the
end of this page (capacity errors for unsplittable notes, width-change and
changed-source-cut errors, and the note acquisition resource limits).

ECMA-376 Part 1 §17.11.1 describes continued footnotes, and §17.11.21 places
the note band at the page bottom. The layout retains complete shaped lines and
resumes in a new band on the following page. Each retained fragment belongs to
its destination page. The admission policy is checked by the existing regression tests and by public
Word 16.113.3 controls for the classes described below. Earlier Word-produced
controls (a short-note/font-size case and references on opposite pages) are not
retained, so whether their results are preserved is not established. Separately, library content-retention
policy permits splitting a note that cannot fit with its first reference line
on any fresh page.

Where a note is reserved is library policy, not a normative rule: §17.11.21 and
§17.18.34 only assign a note to the page that paints its reference. At each line
that first references notes on a page, the layout takes one note plan from that
line's remaining space (earlier new notes whole, only the last one split, at least
one real line plus any continuation notice), and later lines of the paragraph are
admitted only beside it. A later reference line can add space but never shrinks
an earlier note; if its own plan cannot fit, that line and the rest of the
paragraph move to the next page and the earlier notes stay as planned. A plan
also leaves room for the body lines that the paragraph's own `keepLines`
(§17.3.1.14) or widow/orphan control (§17.3.1.44) require to accompany the
reference line. When those required lines carry further references, the plan
first leaves each of those notes its minimum (its first line when it may continue,
otherwise the whole note), so the required lines are admitted together with their
notes. A page's continuation notice is reserved once, whichever of those notes
continues. Continued note lines are placed first on each following page; a
reference line whose own note cannot keep a real line beside them moves on to
the next page with room rather than being forced in; a framed paragraph whose
whole note does not fit beside them moves on the same way. A plan charges a
reference paragraph's authored space after (§17.3.1.33: spacing that follows
its last line) beside the note band, which follows the body; the page-edge
allowance for trailing space after is not applied at that boundary. If a whole
note misses beside the paragraph that closes with its reference only because of
that paragraph's authored space after, the note may continue even if it would fit
whole on a fresh page. When a head with at least one real note line fits, it is
sized beside the full charged body, so that spacing stays clear of the band. Word
16.113.3 split such notes in public exact-line controls (two- and four-line
notes) and kept them whole at the exact charged fit. When no real line fits
there (for example a one-line note), the library keeps the existing whole-note
fallback, which can still overlap that spacing. This is a known library
limitation, not Word's rule: in the public one-line control, Word moved the
final reference-bearing line and its note to the next page. On a page whose note
band already opens with continued lines of an earlier note, a later note whose
whole misses beside its reference may also continue there (beside that tail,
sized against the full charged body) instead of moving the reference line,
whether or not its paragraph has space after; it stays whole where the whole
fits. Word 16.113.3 behaved this way in public single-column exact-line controls
with two notes and in one automatic-line document; pages without such a tail
keep the earlier rule. Only when those constraints cannot be met even on a fresh page does the
plan fall back to the reference line alone. Word
16.113.3 controls with one column, exact 10pt lines and a first-line reference
showed this order (a following body line moved instead of shrinking the note
head); later reference lines follow the same policy without separate Word
evidence. With `allowFootnoteContinuation: false`, notes remain whole and this
order does not apply. The separator band charge for documents without a separator story is an
unchanged library policy.

Ordinary and continuation separator stories are independent. Explicit short and
full marks remain distinct from absent/default facts through the shared model,
including models supplied through `modelSources`. The rule's kind determines
its width even when a short mark appears in a continuation story. A bare empty
reserved story suppresses its rule while preserving the note-band gap. A bare
`separator` mark uses the existing short-rule policy (one third of text width);
§17.11.23 specifies a partial rule, not that numeric fraction. A bare
`continuationSeparator` mark uses the full main-story width (§17.11.1). A listed
(§17.11.9) footnote story whose one paragraph formats its single mark run keeps
that paragraph: its laid-out height under its own spacing and line rule replaces
the 6pt band for its role, while its mark keeps the rule width above. Word
16.113.3 controls agree in pagination: with an exact 24pt story, the reference
line and its whole note moved to the next page; with exact 6pt stories, the
reference line kept the one or two note lines that band leaves room for. The
band applies whether or not continuation is enabled. The band includes the
story's spacing before and after its line (§17.3.1.33), but the rule is drawn
at the midpoint of the mark's own line box, never in that spacing. That
midpoint is library policy, not a Word glyph-baseline observation. Bare stories
and endnote stories keep the 6pt band, with the rule at its midpoint.
General authored
separator text, paragraph borders or shading, and mark-run formatting that
differs from the paragraph mark are not yet retained; they use the existing DOCX
default rule. This is not permission for a native
producer to discard separator formatting: it must admit its formatting class
and special-character semantics before emitting the closed shared rule fact.
Native CP-empty, guard-only and paragraph-only stories are not implicitly
equivalent to DOCX default or bare-empty facts. Continuation notices and
endnote continuation are outside this footnote continuation contract.

Word-produced controls confirm marker selection separately from geometry: at
300pt main width, an 80-paragraph plain single-spaced note printed in Times New
Roman 12pt retains a listed Short continuation mark, while a listed Full mark
spans the main width. A missing continuation story prints a full-width rule and
Word adds a listed Full story on save. The parser preserves that original absence;
the layout applies its existing full-width continuation fallback. These are
bounded Word 16.113.3 compatibility-mode observations. The observed Short rule is
144pt, which differs from this library's one-third policy at that width. They do
not establish a universal Short constant, font-independent scaling or note
spacing.

Continuation currently partitions ordinary paragraph flow. Fitting tables and
framed paragraphs stay whole; notes requiring their splitting can still raise a
capacity error.
A long first note followed by another note can also exceed the admission
capacity. Width changes raise an explicit error because the continuation cursor
cannot reflow across different text widths. Page fields are resolved on the
destination page; stable source cuts continue, while a field-induced reflow that
changes the source cut raises an error before duplicated or omitted text can be
committed. Supporting those changing cuts requires source-offset reflow.

Full-note acquisition is bounded before shaping by one million source units
(strings count by UTF-16 length, and properties, containers and primitives also
count). A complete pagination execution also permits at most 32 million source
units of full-note acquisition across every footnote and convergence pass.
Destination-dependent reacquisitions are charged before shaping; immutable
whole-story cache hits are free. A new layout execution gets a fresh allowance.
The physical paginator retains its 10,000-page ceiling. These are library resource
policies, independent of OOXML admission behavior, and do not promise a wall-clock
or peak-memory limit. Continuation raises an explicit resource error beyond these
limits. Cancellation is checked between destination-page acquisitions.
