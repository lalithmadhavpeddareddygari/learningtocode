# Integration testing — market research and proposal

Internal QE deck for the Embedded Coder **Code Efficiency** team: industry research on integration testing, and a proposal that sits between unit tests and full codegen + SIL system tests.

## Files

| File | Use |
| --- | --- |
| `Integration_Testing_Market_Research.pptx` | Present in PowerPoint / Teams |
| `index.html` | Present in a browser (arrow keys; Speaker notes). Jump with `?s=19` |
| `SPEAKER_NOTES.md` | Talking points, timing, likely questions |
| `generate_pptx.py` | Regenerates the `.pptx` (`python3 generate_pptx.py`) |

## Proposal in brief

1. Treat today’s SIL-heavy suite as **system** tests (necessary, sparse).
2. Add a true **integration** layer: snapshot/FileCheck-style checks on generated C or IR, static metric contracts, optional host kernel compile+run — **no SIL by default**.
3. **Pilot** on one optimization family (buffer reuse/copies or SIMD) before scaling.
