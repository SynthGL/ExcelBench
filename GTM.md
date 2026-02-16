# WolfXL GTM Plan

## Objective
Position WolfXL as the fastest credible path for existing Python Excel workflows by focusing on one wedge:

- openpyxl-compatible API
- reproducible performance + fidelity benchmarks
- low migration friction

## Core Strategy
### T0 — Proof-First Developer GTM
1. Publish reproducible benchmark methodology and raw results.
2. Make migration friction near-zero (3-minute quickstart + import-swap examples).
3. Recruit 5-10 design partners and produce outcome-backed case studies.
4. Launch across developer channels with evidence-first messaging.

### Why this strategy
For developer tools, trust beats reach. Adoption grows when claims are reproducible and switching risk is low.

---

## Landing Page Draft
### Hero
- Headline: **Make Excel automation faster without rewriting your code.**
- Subheadline: WolfXL is an openpyxl-compatible engine backed by Rust, built for high-fidelity Excel reads/writes with reproducible benchmarks.
- CTA 1: **Get Started in 3 Minutes**
- CTA 2: **See Reproducible Benchmarks**

### Proof Strip
- Openpyxl-compatible API
- Reproducible benchmark suite
- Public fidelity + performance tracking

### How It Works
1. Swap imports.
2. Run your workflow (read/write/modify).
3. Measure impact on your files.

### Trust Section
- No black-box claims (methodology + commands + fixtures published)
- Known limitations documented
- Compatibility path for legacy `excelbench_rust` imports

### Install + Migration Snippets
```bash
pip install wolfxl
```

```python
# before
from openpyxl import load_workbook

# after
from wolfxl import load_workbook
```

---

## Launch Post Drafts
### Post A (X / LinkedIn)
```text
We just launched WolfXL: a Rust-backed, openpyxl-compatible Excel engine focused on speed + fidelity.

What it is:
- Drop-in style API for common openpyxl workflows
- Reproducible benchmark methodology (not cherry-picked screenshots)
- Public results + known limitations

Why we built it:
Excel automation at scale can get painfully slow, and rewrites are expensive.
WolfXL aims to reduce runtime without forcing teams to rebuild pipeline logic.

Try it:
[GitHub URL]
[PyPI URL]
[Benchmark methodology URL]
```

### Post B (HN / Reddit style)
```text
Show HN: WolfXL — openpyxl-compatible Excel I/O with reproducible fidelity + performance benchmarks

I built WolfXL to speed up Python Excel workloads without requiring a new programming model.

What’s different:
- API designed to feel familiar for openpyxl users
- Benchmark results are reproducible (fixtures, commands, hardware context disclosed)
- Compatibility transition path included

I’d love feedback on:
1) Real-world files where performance regresses
2) Fidelity edge cases we should prioritize
3) Missing migration docs/examples that would block adoption

Repo + docs:
[GitHub URL]
```

---

## Design Partner Outreach Template
```text
Subject: Can we benchmark your Excel pipeline with WolfXL? (hands-on support)

Hey [Name] — I’m building WolfXL, a faster openpyxl-compatible Excel engine.

I’m onboarding a small set of design partners and can help migrate one real workflow end-to-end.
Goal: measurable runtime improvement without changing your business logic.

What I’m asking:
- 1 representative workbook/pipeline
- 30–45 min kickoff
- feedback on migration friction + correctness

What you get:
- before/after benchmark report on your workload
- direct support on migration issues
- priority fixes for blockers you hit

If useful, I can send a 1-page technical brief and proposed test plan.

— [Your Name]
```

---

## 30-Day Execution Plan
### Week 1 — Trust Foundation
- Publish benchmark methodology page.
- Publish known limitations + compatibility notes.
- Finalize 3-minute quickstart.

### Week 2 — Activation Assets
- Publish 3 runnable examples:
  - read-heavy analytics
  - write-heavy reporting
  - modify-mode workflow
- Publish comparison notebook/script.

### Week 3 — Design Partner Sprint
- Onboard 5-10 teams.
- Capture before/after: runtime, memory, migration effort.

### Week 4 — Public Launch
- Publish launch thread + Show HN/Reddit post.
- Publish first 2-3 case studies.
- Share roadmap informed by partner feedback.

---

## KPI Dashboard (Weeks 1-4)
### North Star
- Activated users/week (install + first successful workflow)

### Acquisition
- PyPI installs/week
- GitHub stars/week
- Landing page -> install conversion

### Activation
- % first successful read
- % first successful write/save
- Median time-to-first-success

### Retention
- 7-day retention
- 28-day retention
- Active design partners after week 2

### Trust/Quality
- Reported fidelity issues/week
- Reported perf regressions/week
- Independent public validation posts

---

## Launch Week Checklist (Operational)
### Day 1
- Ship methodology + limitations pages.
- Verify all benchmark commands run clean from a fresh environment.

### Day 2
- Publish quickstart and migration examples.
- Internal dry run of partner onboarding flow.

### Day 3
- Send partner outreach to first 15 targets.
- Schedule 5 calls.

### Day 4
- Publish social launch post.
- Publish Show HN / Reddit post.

### Day 5
- Aggregate feedback.
- Prioritize top migration blockers and correctness gaps.

---

## Messaging Guardrails
- Never claim universal speedups; always include context + reproducible commands.
- Lead with compatibility + correctness first, speed second.
- Explicitly document where WolfXL is not faster or not yet feature-complete.

## Immediate Next Steps
1. Replace placeholders (`[GitHub URL]`, `[PyPI URL]`, `[Benchmark methodology URL]`).
2. Publish benchmark methodology + limitations pages before external posts.
3. Start outreach with a shortlist of 15 high-fit design partner prospects.
