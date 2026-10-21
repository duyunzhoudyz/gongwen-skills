<div align="center">

# gongwen-skills

**Agent Skills for writing Chinese official documents (公文) by the rules: they follow the national regulations and standards, adapt to the type of organization and the writer's role, never invent facts, respect confidentiality, and cut length instead of padding it.**

22 skills covering drafting, revision, review and standard-compliant Word layout, plus five industry packs for townships and communities, public hospitals, universities, K-12 schools and state-owned enterprises.

[简体中文](README.md) | English

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Skills](https://img.shields.io/badge/skills-22-brightgreen)
![Version](https://img.shields.io/badge/version-1.0.0-orange)
[![GitHub stars](https://img.shields.io/github/stars/duyunzhoudyz/gongwen-skills?style=flat&logo=github)](https://github.com/duyunzhoudyz/gongwen-skills/stargazers)
![Platforms](https://img.shields.io/badge/platforms-Claude%20Code%20%7C%20Codex%20%7C%20Cursor%20%7C%20Gemini%20CLI%20%7C%20Trae%20%7C%20Kimi%20%7C%20Qwen-lightgrey)
![Prompts](https://img.shields.io/badge/prompts-DeepSeek%20%7C%20Doubao%20%7C%20Kimi-lightgrey)

</div>

> **Who this is for.** gongwen-skills is a Chinese-language skill collection for people who write official documents in China: staff at government agencies, public institutions such as hospitals and schools, and state-owned enterprises. All skills, prompts and outputs are in Chinese, and the rules follow Chinese regulations and national standards. It does not translate documents or write English ones. This page is for developers and AI enthusiasts who want to understand, install or evaluate the project.

> **Confidentiality.** Never enter classified information into any internet-connected AI tool. For internal materials, use a model deployed on your organization's intranet; the prompt version and the offline package work with intranet models. See [使用须知](使用须知.md) (usage notice, in Chinese).

---

## Why

More and more people draft official documents with AI, and the same failures keep showing up:

- **Invented facts**: ask for a few local examples and get cases that never happened or companies that do not exist.
- **AI tone**: buzzwords, "not X but Y" reversals and dashes everywhere.
- **Hedging**: a request for instructions with no recommended option, a notice full of "suggest" and "you may consider", and a closing "for reference only".
- **Wrong format**: Markdown `**` and `#` pasted into Word, then the whole layout redone by hand to the national standard.
- **No sense of the organization**: hospital performance indicators, two-level document issuance at universities, the Party committee's prior review at state-owned enterprises.

gongwen-skills turns the experience of veteran writers into reusable, checkable workflows:

| | Typical AI writing | gongwen-skills |
|---|---|---|
| Data and examples | Made up when missing | Only from your materials; gaps are marked 【待补】 (to be filled) |
| Length | Grows and gets padded | Cuts by default; hits your word count when you set one and lists what material is missing |
| Style | AI tone and clichés | Removes AI tone and keeps legitimate official-writing patterns |
| Stance | Hedges and disclaimers | Clear wording; anything to verify goes in a note outside the document body |
| Format | Do it yourself | Generates GB/T 9704 compliant .docx; machine checks punctuation, numbers and document numbers |
| Organization type | One size fits all | 7 organization profiles, 5 role modes, 5 industry packs |

## Results

### Before and after: an AI-drafted notice

A simulated AI draft of a district education bureau's winter campus-safety notice, passed through `gongwen-direct` → `gongwen-humanize` → `gongwen-review` (full record in [examples/B-AI稿体检](examples/B-AI稿体检/)):

| Machine check | P0 must fix | P1 should fix | Length (Chinese characters) |
|---|---:|---:|---:|
| Original | 5 | 31 | 1092 |
| After | 0 | 0 | 807 |

Changes included: "suggest that schools set up..." became "schools shall set up..."; the document number, dates and number ranges were corrected to the national standards; a threat to have parents "detained by the police" was removed because the education bureau has no such authority; every figure in the text stayed unchanged, confirmed by a fact-diff script.

### Compared with similar projects

We ran this project and 8 similar Chinese official-document skills on GitHub on the same tasks: same model (claude-sonnet-5-5), same user messages and materials, blind-graded by two LLM judges (Claude and DeepSeek). The table below comes from two **held-out task sets** (no skill was changed based on them): 54 tasks, 552 runs, all materials fictional.

| Dimension | This project | Best of the others |
|---|---|---|
| Facts kept when removing AI tone | **682/682 kept, 0 fabrications** | 98.4%; judges flagged fabrication 15 times |
| Formatting: GB/T 9704 layout audit | **96.6** | 93.2 |
| Formatting: missing fonts | **standard font names in all 32 runs, no substitution** | one project substituted fonts in all 32 runs, 12 without telling the user |
| Writing: runs with fabrication (Claude judge) | **11/40** | 16/40 |
| Stopping on classified materials | **only project to stop** when the mark is on the first page; all missed a header-only mark (fixed here and re-tested) | none |
| Writing quality (mean rank, 1 = best) | 3.20, 2nd (Claude judge) | 3.15 |
| Formatting quality | 2.72, 3rd | 2.12 |
| AI-tone removal quality | 1.97, 2nd (DeepSeek); 2.19, 3rd (Claude) | 1.69 / 1.78 |

On the things that must not go wrong (no fabrication, facts kept, confidentiality, layout compliance, honest font handling) this project leads every other project tested; on overall document quality it is in the top tier, 0.1 to 0.6 ranks behind the best, and we keep improving. Projects, commits, per-track results, the iteration history and limitations are in the [competitive benchmark report](evals/results/compete-report.md) (in Chinese); scripts and tasks are in [evals/compete](evals/compete/).

### Benchmark

Before v0.1, the first 7 skills were evaluated with a real model: the same model (claude-opus-5-5), the same user messages and materials, one group with the skill loaded and one without, graded blind. 104 test cases, 621 checks, all materials fictional.

| Measure | With skill | Without skill |
|---|---:|---:|
| Pass rate, all checks (621) | **94.8%** | 59.7% |
| Pass rate, general checks (422) | **95.0%** | 74.6% |
| Runs with fabricated content, 34 anti-fabrication cases | **0** | 15 |
| Correctly stopped on classified or internal material, 7 cases | **7** | 2 |

- "General" checks are what an experienced writer should meet from the user message and materials alone. They exclude this project's own conventions (such as the drafting-notes section), which the no-skill group cannot know. Both figures are given; use the general one to judge the real gap.
- Without the skill, fabrication was mostly invented work examples, achievement claims and integrity statements, rarely statistics.
- Each case ran once, and the grader was the same model as the one being tested. Methods, per-skill results and limitations are in the [v0.1 benchmark report](evals/results/v0.1-benchmark.md) (in Chinese).

**v1.0 full benchmark (22 skills)**: the tested model is claude-sonnet-5-5; 372 test cases and 2,671 checks, all materials fictional, graded blind. Round 1 evaluated the pre-fix skills. After each skill was fixed from the defects round 1 exposed, round 2 re-ran the with-skill group; that is the release version.

| Measure | With skill (round 2) | Without skill |
|---|---:|---:|
| Pass rate, all checks (2,671) | **98.6%** | 54.1% |
| Pass rate, general checks (1,726) | **99.1%** | 68.7% |
| Runs with fabricated content, 136 anti-fabrication cases | **0** | 66 |
| Correctly stopped on material marked classified or internal, 21 cases | **21** | 3 |

- Round 1, before the fixes, with both groups under identical conditions: 97.1% vs 54.0% on all checks and 97.7% vs 68.5% on general checks. Across all cases, the with-skill group had fabrication in 19 runs in round 1. In round 2 the grader flagged 2, which on review quote the skill's own bundled reference material and do not count as fabrication.
- Trigger test: 439 of 440 phrasings were routed to the right skill (99.8%).
- Each of the 17 core skills passes at 97.4% or higher with zero fabrication and a passing trigger test, and is labeled Stable. The 5 industry packs meet the same numbers but are labeled Beta by convention until front-line testers report back.
- Limitations: each case ran once, and the grader is the same model as the one being tested. The round-2 with-skill group ran at a higher reasoning effort, so the round-2 gap is not a like-for-like comparison. The retest used the same cases the fixes were based on. The with-skill group still failed 37 checks (across 16 skills). There is no cross-model or human blind evaluation and no front-line tester feedback yet. Methods, per-skill results and limitations are in the [v1.0 benchmark report](evals/results/v1.0-benchmark.md) (in Chinese).

v1.0 revised several rules (first-round delivery, length) and the test cases with them, so it is a new baseline and is not directly comparable with v0.1.

## Quick start

| You use | How to install |
|---|---|
| DeepSeek, Doubao, Kimi or another web chatbot, or an intranet model | No install: copy a prompt (option 4) |
| Claude Code | Plugin marketplace (option 2) |
| Codex, Cursor, Gemini CLI, GitHub Copilot, Qwen Code, Kimi Code CLI, Trae, CodeBuddy | One-click script that detects installed tools (option 1), or option 3 |
| An agent that can install things itself | Ask it (option 5) |

**Option 1: one-click script (Windows, macOS, Linux).** Clone or download this repository, then run in its folder:

```powershell
# Windows
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

```bash
# macOS, Linux
bash install.sh
```

The script checks for Python 3.9+, asks before installing the layout dependencies, and installs the 17 core skills into every agent tool it finds. Add `-WhatIf` (Windows) or `--dry-run` (macOS, Linux) to preview without writing anything. To target one tool: `python tools/install.py --list-targets`, then `python tools/install.py --target qwen`.

**Option 2: Claude Code plugin marketplace.**

```text
/plugin marketplace add duyunzhoudyz/gongwen-skills
/plugin install gongwen@gongwen-skills
```

**Option 3: npx skills** (Codex, Cursor, Gemini CLI, GitHub Copilot and others).

```bash
npx skills add duyunzhoudyz/gongwen-skills
```

**Option 4: web chatbots and intranet models.** Open [`dist/prompts/`](dist/prompts/), copy one prompt file (one per skill, about 26,000 to 29,000 Chinese characters each), paste it at the start of a conversation, then state your request. Start with `gongwen.md`, the entry point. When the work moves to another skill, you get a handoff card to paste into a new conversation after the next prompt. For intranet use, download the offline package from [Releases](https://github.com/duyunzhoudyz/gongwen-skills/releases) or build it with `python tools/build_prompts.py --offline`.

**Option 5: let your agent install it.**

```text
请把 https://github.com/duyunzhoudyz/gongwen-skills 中 skills/ 目录下的 skill 安装到你的 skills 目录。其中 gongwen-grassroots、gongwen-hospital、gongwen-university、gongwen-k12、gongwen-soe 是行业包，只装我需要的 gongwen-hospital。装好后告诉我有哪些可用。
```

(It asks the agent to install the skills from `skills/`, taking only the hospital pack among the five industry packs, and to list what is available.)

### Industry packs

Industry packs install on top of the core skills. Without a pack, the core skills still handle that industry using the organization profile; with it, the industry's specialized documents go to the pack.

| Industry | Pack |
|---|---|
| Townships, sub-districts, villages and communities | `grassroots` |
| Public hospitals | `hospital` |
| Universities | `university` |
| Primary and secondary schools | `k12` |
| State-owned enterprises | `soe` |

```text
# Claude Code
/plugin install gongwen-hospital@gongwen-skills
```

```bash
powershell -ExecutionPolicy Bypass -File .\install.ps1 --pack hospital   # Windows
bash install.sh --pack hospital                                         # macOS, Linux
python tools/install.py --target auto --pack hospital                   # installer
```

Repeat `--pack` to choose several, or use `--pack all`.

> **Dependencies**: Word layout and Excel charts need Python 3.9+ with `python-docx` and `openpyxl` (`pip install python-docx openpyxl`). Without Python, the layout skill gives a manual WPS/Word settings checklist instead.
>
> **Uninstall**: add `--uninstall` to the script or installer command, review the list, then add `--yes`. Only skills installed by this project are removed.

## Skill map

17 core skills ship as one plugin; the 5 industry packs are optional. **Stable** means the skill met the real-model benchmark bar (pass rate of at least 80%, zero fabrication, correct triggering). **Beta** means it has been tested on samples and edge cases may still fail; the 5 industry packs also passed the v1.0 real-model benchmark, but by design stay Beta until frontline feedback and human blind review.

| Group | Skill | What it does | Maturity |
|---|---|---|---|
| Entry and setup | [`gongwen`](skills/gongwen/) | Entry point: works out unclear requests, plans the route through other skills, builds a 3-question profile, reminds you of seasonal writing tasks | Stable |
| Entry and setup | [`gongwen-profile`](skills/gongwen-profile/) | Full organization profile, house-style learning from past documents, leader preferences from edited drafts, banned-word list | Stable |
| Drafting | [`gongwen-draft`](skills/gongwen-draft/) | Statutory documents: notices, requests for instructions, reports on a single matter, letters, circulars, replies, opinions, internal memos, decisions, announcements | Stable |
| Drafting | [`gongwen-report`](skills/gongwen-report/) | Summaries, plans, work briefings, research reports, experience write-ups, performance reports, self-made implementation plans, congress work reports | Stable |
| Drafting | [`gongwen-speech`](skills/gongwen-speech/) | Leaders' speeches, host scripts, addresses, pledge and experience-sharing remarks; length set by speaking time | Stable |
| Drafting | [`gongwen-party`](skills/gongwen-party/) | Party-building and organizational-life materials; for personal materials without personal examples, an outline and filling guide only | Stable |
| Drafting | [`gongwen-brief`](skills/gongwen-brief/) | Government information items, bulletins, press releases, official-account posts; whether an item is likely to be adopted | Stable |
| Drafting | [`gongwen-minutes`](skills/gongwen-minutes/) | Meeting records, minutes, agenda papers, follow-up lists; records only what was decided | Stable |
| Drafting | [`gongwen-implement`](skills/gongwen-implement/) | Turns a higher-level document into task breakdowns, implementation plans and progress reports; flags copied text | Stable |
| Drafting | [`gongwen-rectify`](skills/gongwen-rectify/) | Rectification plans, ledgers (exportable to Excel), progress reports and self-inspection reports, reconciled against the ledger | Stable |
| Revision and review | [`gongwen-title`](skills/gongwen-title/) | Titles, subtitles, three sets of section headings, outlines, title-body consistency | Stable |
| Revision and review | [`gongwen-polish`](skills/gongwen-polish/) | Shorten, expand with existing facts only, change register, revise per a leader's comments, merge drafts | Stable |
| Revision and review | [`gongwen-humanize`](skills/gongwen-humanize/) | Removes AI tone while keeping official-writing patterns; facts untouched | Stable |
| Revision and review | [`gongwen-direct`](skills/gongwen-direct/) | Removes disclaimers and hedging; makes requests and replies take a clear position | Stable |
| Revision and review | [`gongwen-review`](skills/gongwen-review/) | Eight-gate review with a graded issue list, machine checks, confidentiality scan, banned-word check | Stable |
| Delivery tools | [`gongwen-format`](skills/gongwen-format/) | GB/T 9704-2012 Word output in six layouts (the bulletin layout follows common practice, as no national standard exists); repairs existing .docx; writes review comments into Word | Stable |
| Delivery tools | [`gongwen-data`](skills/gongwen-data/) | Table checks, merging, year-on-year and percentage-point calculations with formulas, local black-and-white Excel charts | Stable |
| Industry pack | [`gongwen-grassroots`](skills/gongwen-grassroots/) | Townships and communities: information items, inspection reports, notices to residents, emergency notices, ledgers | Beta |
| Industry pack | [`gongwen-hospital`](skills/gongwen-hospital/) | Public hospitals: accreditation self-assessment, performance analysis, quality reports, RCA and PDCA, department-head reports | Beta |
| Industry pack | [`gongwen-university`](skills/gongwen-university/) | Universities: audit self-assessment, discipline evaluation materials, grant-application form checks, internal memos | Beta |
| Industry pack | [`gongwen-k12`](skills/gongwen-k12/) | Schools: term plans and summaries, inspection self-assessment, parent-meeting speeches, student comments, letters to parents | Beta |
| Industry pack | [`gongwen-soe`](skills/gongwen-soe/) | State-owned enterprises: major-decision agenda papers, board resolutions, operating analysis, reform ledgers, ESG reports | Beta |

## Example requests

Users talk in plain Chinese; no skill names needed. A few examples:

| Request (English gloss) | Handled by |
|---|---|
| "My boss asked me to write something and I'm not sure what it is." | `gongwen` |
| "Write a request for instructions to the county natural resources bureau for temporary land use." (It will point out that the two bodies are not in a reporting line, so a letter is the right form.) | `gongwen-draft` |
| "Here is our department's work log for the year; write the year-end summary." | `gongwen-report` |
| "This was written by DeepSeek and sounds like AI. Fix it, but don't touch the numbers." | `gongwen-humanize` |
| "Check this before it goes for signature tomorrow, then lay it out as a red-header Word file." | `gongwen-review` → `gongwen-format` |
| "The national performance results are out and our CMI dropped; the director wants an analysis." | `gongwen-hospital` |

More examples are in the [Chinese README](README.md#直接这样说) and in each skill's README.

## Workflows

| Scenario | Route |
|---|---|
| Year-end summary, performance report, briefing | `gongwen` → `report` (fact ledger, outline, draft) → `direct` → `humanize` → `review` (pause for your decisions) → `format` |
| AI draft check-up | `direct` → `humanize` → `review`, then `format` if you need Word |
| Implementing a higher-level document | `implement` (key points, task table, plan, copy check) → `direct` → `review` (pause) → `format`; later, `implement` writes the progress report from the task table |

## Six guardrails

All skills follow these, ahead of any specific workflow ([shared/guardrails.md](shared/guardrails.md), in Chinese):

1. **Confidentiality first**: stop and warn on classification marks, "internal" labels or personal data.
2. **Facts only from the user**: facts only you can supply are marked 【待补】; doubtful items are listed in the drafting notes after the text; "next Wednesday" is resolved from today's date and noted, "last year" is kept as written, and the issue date is left for the signing day.
3. **Clean body, notes outside**: no disclaimers or AI self-talk in the document; items to verify go into drafting notes after it.
4. **Cut by default**: get to the point, respect official length limits, hit a requested length without padding.
5. **Traceable, machine-checkable rules**: every format, punctuation and number rule cites its clause; scripts check what scripts can.
6. **Assist, don't ghostwrite**: materials that need the author's own experience get an outline only when no personal examples are given.

## Organizations and roles

Seven built-in organization profiles (Party and government agencies, county and township level, communities, public hospitals, universities, K-12 schools, state-owned enterprises) and five role modes (teaching, efficient, translating plain language, reviewer, strict). A 3-question profile is created on first use; `gongwen-profile` builds the full version from your organization's own rules and past documents. Profiles live in `.gongwen/`, which is git-ignored because it may contain internal information.

## Rule sources

Regulations on the Handling of Official Documents of Party and Government Organs (2012); GB/T 9704-2012 (layout of official documents); GB/T 15834-2011 (punctuation); GB/T 15835-2011 (numerals in publications); Several Provisions on Rectifying Formalism and Reducing Burdens on Grassroots Units (2024); public explanations from the National Bureau of Statistics for data wording. Summaries with clause numbers are in [shared/rules/](shared/rules/) (in Chinese). The lint engine runs 12 check categories across 41 document types.

## FAQ

**Is it free? Do I need Claude?** It is free and open source under the MIT license. Skills install into Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot, Qwen Code, Kimi Code CLI and similar tools. Without any of these, use the prompt version: copy a file from `dist/prompts/` into a chat on DeepSeek, Doubao, Kimi or a model deployed on your intranet.

**Can I paste my organization's materials into it?** Not if they are classified or marked "internal" or "work secret". Never enter such material into an internet-connected AI tool; use an intranet model instead, and anonymize personal data first. The skills stop and warn when they see a classification mark, but that is only a last reminder. The responsibility stays with the user.

**Will it make up data?** The rules say to use only what you provide, mark facts only you can supply as 【待补】 and list doubtful items in the drafting notes. In the v0.1 benchmark, 34 cases deliberately withheld facts: with the skills there were 0 runs with fabrication, without them 15. Output varies between runs, so check every number, name and document number before use.

**Can it make a draft undetectable as AI-written?** No, and it does not claim to. Removing AI tone improves style, such as paragraph-opening commentary, buzzwords and too many dashes, so the text reads better.

**I work in a hospital, university, school, state-owned enterprise or township. Is there something for me?** Yes, five optional industry packs (`gongwen-hospital`, `gongwen-university`, `gongwen-k12`, `gongwen-soe`, `gongwen-grassroots`). General documents such as summaries and speeches still go through the core skills, which follow the organization profile.

**The Word fonts are wrong, or I have no Python.** Fonts are not bundled, and the tools never substitute another font. If a standard font (FangSong_GB2312, KaiTi_GB2312, Founder XiaoBiaoSong, SimHei, SimSun) is missing, the layout skill says which ones, explains where to obtain and how to install them, and refuses to render a preview until they are installed; obtain them through licensed channels. Without Python, the layout skill gives a manual WPS and Word settings checklist. Print and check before formal issue.

## Roadmap

v0.2 to v0.4 were planned as separate releases; they were finished early and ship together with v0.1 as v1.0.0, distinguished by maturity labels. See [CHANGELOG.md](CHANGELOG.md) (in Chinese).

- [x] v0.1: entry point, reports, titles, AI-tone removal, de-hedging, review, layout
- [x] v0.2: profiles, statutory documents, speeches, Party materials, polishing
- [x] v0.3: implementation, rectification, minutes, briefs, data and charts
- [x] v0.4: five industry packs
- [x] v1.0 (October 31, 2026): all 22 skills; real-model benchmark; one-click install; optional industry packs

Next: cross-model evaluation of the prompts on DeepSeek, Qwen and Kimi; tester feedback and human blind review to promote industry packs from Beta; layout testing in Word and WPS and font detection on macOS and Linux; Word tracked-changes output; merging per-skill check scripts into the lint engine.

Want to help test? Read the [tester handbook](docs/评测/测试者手册.md) (in Chinese): try the skills on fictional material, no programming needed.

## Community

- Douyin: (to be added)
- WeChat: (to be added)
- Douyin group and WeChat group: (QR codes to be added)

Rule corrections, new scenarios, industry material and bug reports go to [Issues](https://github.com/duyunzhoudyz/gongwen-skills/issues/new/choose); discussion goes to [Discussions](https://github.com/duyunzhoudyz/gongwen-skills/discussions). See [CONTRIBUTING.md](CONTRIBUTING.md) (in Chinese).

## Acknowledgements

This project borrows ideas and, where licenses allow, adapted rules from more than 30 open-source projects on Chinese de-AI writing, official-document rules and layout, data handling and education. The full list with licenses is in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md); a grouped table is in the [Chinese README](README.md#致谢).

### Contributors

Rule corrections, industry material and test cases are welcome; contributors will be credited here after their changes are merged, with their consent.

## License

[MIT](LICENSE). Rule summaries are paraphrased from public regulations and national standards and do not replace the originals.
