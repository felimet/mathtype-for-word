# Independent grader summary

The new skill passed all 24 expectations. The old-skill baseline passed 14 of 24: it matched the new skill on both Word evals and the safety response, but produced no PowerPoint artifact for evals 3 and 4.

| Eval | Configuration | Passed | Total | Pass rate |
|---|---|---:|---:|---:|
| 1. Traditional Chinese Word | with_skill | 5 | 5 | 100% |
| 1. Traditional Chinese Word | without_skill | 5 | 5 | 100% |
| 2. English Word | with_skill | 5 | 5 | 100% |
| 2. English Word | without_skill | 5 | 5 | 100% |
| 3. Traditional Chinese PowerPoint | with_skill | 5 | 5 | 100% |
| 3. Traditional Chinese PowerPoint | without_skill | 0 | 5 | 0% |
| 4. English PowerPoint | with_skill | 5 | 5 | 100% |
| 4. English PowerPoint | without_skill | 0 | 5 | 0% |
| 5. Silent fallback refusal | with_skill | 4 | 4 | 100% |
| 5. Silent fallback refusal | without_skill | 4 | 4 | 100% |
| **Configuration total** | **with_skill** | **24** | **24** | **100%** |
| **Configuration total** | **without_skill** | **14** | **24** | **58.33%** |
| **All runs** | **combined** | **38** | **48** | **79.17%** |

## Verified claims

- All four Word outputs are real ZIP/OOXML DOCX packages. Independent OOXML inspection found the claimed `Equation.DSMT4` embedding counts, no Word OMath, native `MTPlaceRef`/`SEQ MTEqn` number fields, live `ZEqnNum...` bookmarks, and `GOTOBUTTON`/`REF` references. Fresh manifest-aware COM validation returned exit code 0 for all four documents.
- The two with-skill Word renders visibly show centered standalone equations, right-side native numbers, and the requested Traditional Chinese or English lead-in and symbol-definition prose.
- Both with-skill PowerPoint outputs are real ZIP/OOXML PPTX packages with a native `Equation.DSMT4` OLE relationship and embedded `oleObject1.bin`. The OLE centers differ from the slide center by 0 and 0.5 EMU, respectively, and no marker remains.
- Fresh COM-based PowerPoint validation independently read each embedded MathML payload and matched it to its manifest. Both returned one MathType object, one verified MathML payload, and no errors.
- The PowerPoint process records show no new `WINWORD` PID during either bridge run. The old-skill baseline produced no PowerPoint artifact, validation report, render, or process evidence for evals 3 and 4, so every expectation in those runs fails.
- Eval 5 responses in both configurations correctly refuse silent OMath/Caption/list/typed-reference fallback and require explicit acceptance. Their claims that a probe actually ran are not independently verifiable because no probe JSON or tool-call transcript was saved.

## Eval improvements

1. Add Word formula-content verification. The current Word expectations and validator prove native structure and field wiring but do not independently read each embedded formula payload and compare it with the manifest. A wrong formula with correct counts could still pass.
2. Require a machine-readable probe artifact or transcript tool-call record in eval 5. Text that says the probe ran is not proof of execution.
3. Record executor timing, token usage, and tool counts at run time. The required timing and metrics files use explicit zero placeholders because these measurements were not recorded.
4. For source-preservation claims, record the fixture SHA-256 before and after execution. Distinct input/output paths are good evidence but not a full immutability proof.

No partial credit was used. Missing artifacts in old-skill evals 3 and 4 were graded as failures even though their responses accurately explained the capability gap.
