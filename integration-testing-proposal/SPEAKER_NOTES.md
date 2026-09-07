# Speaker notes — Integration testing for Code Efficiency

Use with `Integration_Testing_Market_Research.pptx` or `index.html`. Target **40–45 minutes** plus discussion.

## How to open

- **PowerPoint:** `Integration_Testing_Market_Research.pptx`
- **Browser (no PowerPoint):** open `index.html`, arrow keys or buttons, **Speaker notes** in the toolbar

## Story in one sentence

Industry treats integration testing as a **narrow, interface-level** layer between unit tests and full-system tests. For Embedded Coder Code Efficiency, that layer should assert **generated structure and static metrics** (and sometimes a tiny host run), not full **codegen + SIL**, which we keep as a sparse system net.

## Suggested timing

| Slides | Minutes | Goal |
| --- | --- | --- |
| 1–5 | 8 | Problem and inverted pyramid |
| 6–10 | 12 | Market research + analogues |
| 11–15 | 12 | Gap + proposal + CI |
| 16–19 | 10 | Pilot, metrics, ask |
| 20 | leftover | Q&A with sources up |

## Slide-by-slide

1. **Title** — You are QE, tasked by the Dev Manager. Outcome is a proposal we can pilot, not a literature review.
2. **Problem** — If you have numbers (SIL pack wall-clock, flakes, time-to-blame), say them here.
3. **Agenda** — Decision at the end is Phase 0–1 only.
4. **Team** — Three oracles: structure, metrics, legality. SIL is legality/numerics, not “did the opt fire.”
5. **Inverted pyramid** — Ice-cream cone anti-pattern. Keep SIL; stop writing it first.
6. **Definition** — Integration = interface defects. If it needs the whole stack, it is a system test.
7. **Consensus** — Google 80/15/5, Microsoft gates, ISTQB levels, compiler lit tests.
8. **LLVM** — FileCheck is the pattern to steal. `test-suite` ≈ our SIL.
9. **Automotive** — Customer SIL ≠ our SIL (different system under test). BTC migration suites and Ford CRL tests are integration-scale.
10. **Mapping** — Adopt Small/Medium/Large names internally.
11. **Gap** — Pause; invite “when did SIL miss an efficiency bug?”
12. **Four layers** — L2 is new investment.
13. **Four families** — Start A (snapshots) + C (static metrics).
14. **Buffer reuse example** — Review rule: failure must name pass or metric.
15. **CI** — Presubmit has no SIL. Nightly owns deep system.
16. **Pilot** — Recommend buffer reuse/copies unless SIMD is the hotter pain.
17. **Metrics** — Freeze 2–3 in Phase 0.
18. **Risks** — Goldens, SIL-in-disguise, false confidence.
19. **Ask** — Approve Phase 0–1; name a Dev counterpart; pick a family.
20. **Sources** — Leave up in Q&A.

## Likely questions

**Isn’t SIL already integration testing?**  
It is *system* testing of the generator (full workflow). Integration for us is “did these codegen stages honor their contract?” without executing the customer verify path.

**Will goldens be noisy?**  
Pattern CHECKs, not whole-file diffs. Review golden updates like code (LLVM).

**Do we drop SIL?**  
No. Sentinel SIL per family. L2 never the only numerical oracle for a pass that can change values.

**Why not only more unit tests?**  
Unit tests mock away the pipeline. Efficiency bugs often live in pass interaction, emission, and metrics plumbing.

**ISO 26262 / customers?**  
This is *our* developer test strategy for the tool. Customer B2B SIL remains their verification. Faster internal signal still helps qualification evidence quality.

## Ask to leave with

Approve **Phase 0–1** on one optimization family, with a named Dev pair, and permission to propose SIL demotion when L2 covers the same defect class.
