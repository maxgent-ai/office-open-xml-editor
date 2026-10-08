# Tab fitting differential check

This separate integration gate compares the current DOCX layout with a clean
checkout at the current `origin/main`, using identical parsed facts and text
metrics. Install dependencies and build WASM in both checkouts before running:

```sh
DOCX_TAB_BASELINE_CHECKOUT=<baseline-checkout> pnpm exec vitest run \
  --config packages/docx/tests/differential/tab-fitting.config.ts
```

It covers every no-float case from the float/tab matrix and generated ordinary
RTL/LTR lines with 1–3 tabs, all tab alignments, custom stops, prefixes, and
Latin/CJK/Thai content. The full indent generator varies left/right/firstLine/
hanging independently through negative, zero, and positive values, includes
omitted firstLine/hanging attributes to expose precedence, and varies bidi and
mirrorIndents. Stops fall before, inside, and after the authored reading band;
interleaved cells exercise 1–3 tabs. The reviewer's complete 1,728-case matrix
also covers four-tab counterexamples. Total: 240,370 cases (178 established
matrix, 5,184 original generated, 1,728 reviewer, 233,280 full-indent inputs).
Negative firstLine/hanging attributes are parser robustness inputs, beyond their
normative unsigned XML measure type; negative left/right indents are normative.
Both parsers currently discard mirrorIndents; those cases prove regression
parity, not inside/outside-indent fidelity.

Line partitions and placement geometry must agree unless
main allocates outside the paragraph band or the case is one of the five short
Word-backed RTL positional exceptions. Every candidate must preserve the text
and stay in band. An optional `DOCX_TAB_DIFFERENTIAL_REPORT` path saves counts
outside the checkout. This gate runs separately from `pnpm test` because it
requires the second checkout.
