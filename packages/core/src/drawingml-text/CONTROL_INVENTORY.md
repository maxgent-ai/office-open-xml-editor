# DrawingML text wrapping controls (issue #1562)

These controls use the same `a:txBody` runs, fonts, size, explicit insets,
`wrap="square"`, `noAutofit`, and left alignment in PowerPoint and Excel. The
synthetic cases cover plain Latin text, formatting seams, whitespace, CJK,
compound words, punctuation, nonbreaking space, and a tab. Microsoft Office for
Mac 16.113.2 exported the PDF observations below. `/` separates visual lines;
`␠` marks an ordinary space retained by our current painter; `⇥` marks a tab.
PDF text extraction maps U+00A0 to a space, so C10's extracted space is not a
claim about its Unicode identity.

The XLSX control uses `xdr:oneCellAnchor` with an explicit extent. Its first
box is 819150 × 1047750 EMU, or 64.5 × 82.5 pt. Excel's shape-size UI rounds
that to 0.90 × 1.15 in. The PowerPoint and Excel PDF borders have equal width
for every case, matching the declared EMUs exactly. Each box has 7.2 pt left
and right insets, so its effective text width is the outer width minus 14.4 pt.
The earlier `xdr:twoCellAnchor` control produced Excel PDF boxes 5 pt wider and
5 pt shorter; its line breaks are excluded from this comparison.

| ID | Text | Inner width (pt) | PowerPoint and Excel lines | Current PPTX | Current XLSX |
| --- | --- | ---: | --- | --- | --- |
| C00 | `abcdef` | 50.10 | `abcd / ef` | `abcd / ef` | `abcd / ef` |
| C01 | `abc` + `def` in separate runs | 50.10 | `abcd / ef` | `abcdef` | `abcd / ef` |
| C02 | `abc def` | 56.10 | `abc / def` | `abc␠ / def` | `abc␠d / ef` |
| C03 | `日本語Power` | 116.10 | `日本語 / Power` | `日本語 / Power` | `日本語Powe / r` |
| C04 | `non-managed` | 85.35 | `non- / managed` | `non- / managed` | `non-man / aged` |
| C05 | `abc␠␠` in one run | 41.85 | `abc` | `abc` | `abc␠ / ␠` |
| C06 | `abc` + `␠␠` in separate runs | 41.85 | `abc` | `abc` | `abc␠ / ␠` |
| C07 | `abc ` + `$` + `100` across runs | 67.35 | `abc / $100` | `abc␠ / $100` | `abc␠$1 / 00` |
| C08 | `日本語、次` | 71.10 | `日本語 / 、次` | `日本語 / 、次` | `日本語 / 、次` |
| C09 | `supercalifragilistic` | 79.35 | `supercal / ifragilisti / c` | same | same |
| C10 | `abc` + NBSP + `def` in one run | 56.10 | `abc d / ef` | `abc` + NBSP + `/ def` | `abc` + NBSP + `d / ef` |
| C11 | `abc⇥def` | 86.85 | `abc / d / ef` | one line | one line |

All twelve Office line sets agree across hosts at matched geometry. The shared
line breaker therefore owns these decisions; these controls justify no
host-specific break hook. In C05–C07, PDF extraction cannot expose trailing
spaces, so assertions about terminal-space *advance* must use geometry and
additional evidence rather than the extracted words alone. C08 documents the
observed boundary result; it does not by itself establish a general kinsoku
override. Other DrawingML features and fonts remain subject to their own
specification and Office evidence.
