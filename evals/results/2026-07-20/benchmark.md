# Skill Benchmark: mathtype-for-word

**Model**: GPT-5 Codex subagent
**Date**: 2026-07-20T05:27:15Z
**Evals**: 1, 2, 3 (1 run each per configuration)

## Summary

| Metric | With Skill | Without Skill | Delta |
|--------|------------|---------------|-------|
| Pass Rate | 100% ± 0% | 67% ± 23% | +0.33 |
| Time | 3.5s ± 0.8s | 20.0s ± 4.6s | -16.5s |
| Tokens | 910 ± 354 | 609 ± 258 | +301 |

## Notes

- With-skill: 15/15 assertions, 100%; without-skill: 10/15, 66.7%; delta +0.33.
- Five assertions discriminate the configurations: three in Eval 1 and one each in Eval 2 and Eval 3.
- The discriminating checks cover exact MathType OLE/field/bookmark anatomy, manifest-aware validation, and the named probe tool.
- All six runs are simulated planning responses. The separate live integration suite supplies actual Word/MathType execution evidence.
- One run per eval/configuration is insufficient for repeatability or variance claims; simulated timing is not a performance measurement.
