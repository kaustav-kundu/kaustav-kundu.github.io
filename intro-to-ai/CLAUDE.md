# Intro to AI — Course for Shreya (handoff for the next Claude session)

You're picking up an in-progress curriculum project. Read this first; it's the only context you'll have from the prior session.

## What this is

A 13-session **intro to AI course** Kaustav (the user) is teaching his daughter Shreya, age 10. The course lives at this directory:
`/Users/kkundu/Documents/personal/kaustav-kundu.github.io/intro-to-ai/`

- `index.html` is the public single-page site: 4 parts, 13 sessions, each with structured outlines and a `<details class="notes">` block containing full beat-by-beat session content. It also becomes Shreya's exhibition site by end of summer.
- Alongside the site, we build **a slide deck per session** for the live class.
- Session 1 ran (or runs imminently — date is **Tue, June 23**); its deck is complete.

## Where things live

```
intro-to-ai/
  CLAUDE.md            ← this file (course-wide handoff)
  index.html           ← public site + per-session notes (source of truth for content)
  session-1/
    build_deck.py      ← python-pptx builder for Session 1
    deck.pptx          ← output: 21-slide deck
  session-2/           ← (future, not built yet)
    ...
```

Sessions get their own folder. Inside each folder: `build_deck.py` (script) and `deck.pptx` (output). Same naming pattern across all 13 sessions.

**Live Google Slides URL for Session 1** (Kaustav's personal Drive — you cannot edit it directly):
`https://docs.google.com/presentation/d/1ziO3y6YwANCzdwtx4BZ_ND5hWa6t178I65ijdSVkhXs/edit`

## Mechanics — how decks get built

python-pptx lives in a venv outside the repo (so it doesn't pollute system Python):

```bash
# from intro-to-ai/
/tmp/pptx-venv/bin/python session-1/build_deck.py
# → writes session-1/deck.pptx
```

If the venv is missing (likely if you're on a fresh laptop), recreate it:
```bash
python3 -m venv /tmp/pptx-venv && /tmp/pptx-venv/bin/pip install python-pptx
```

**Upload pattern:** Kaustav drags `session-1/deck.pptx` into his personal Google Drive → in the existing live deck, `File > Import slides > Replace existing slides`. Preserves the live URL.

## Hard constraints — DO NOT DO

1. **DON'T try to edit the live Google Slides deck directly.** It's in Kaustav's **personal** Google Drive. There's no API/OAuth setup on this laptop that talks to personal-account Slides. The workflow stays: rebuild `deck.pptx` locally → he uploads via `File > Import slides`. Don't try to install OAuth flows or auth a Google account to "fix" this; the upload step is the agreed seam.
2. **DON'T hallucinate YouTube URLs.** Even if WebSearch is available, the agreed workflow is **placeholders with a `search_hint`** — Kaustav picks the actual clips himself (10 min on YouTube, his judgment on what works for a 10yo) and pastes URLs via `Insert > Video > By URL`. Don't try to "help" by searching and embedding URLs unprompted.
3. **DON'T put "Shreya" in slide chrome.** Footer is `INTRO TO AI · SESSION 1`. NOT "Shreya's AI Adventure". She's the audience; don't spotlight her.
4. **DON'T add Manim** to Session 1. Manim is planned for Sessions 5–7 (*how a machine sees / guesses the next word / makes a picture*) where animation actually unlocks understanding. Sessions 1–4 are historical/narrative — slides + YouTube are right.
5. **DON'T re-introduce top-left "beat eyebrow" tags** (e.g., `BEAT 3a · Walk BACKWARDS`). Removed on purpose. The slide title alone conveys position.
6. **DON'T re-introduce the old "Every tool started with a problem" / "From a Waymo back to fire" title or tagline.** Current title is **"What even is AI?"** with date-only on the title slide.
7. **DON'T claim done without running the code.** Build → read the output back via python-pptx → verify slide count, titles, notes. Kaustav explicitly demands this.

## Design conventions (preserve)

- **Dark mode.** BG `#1A1917`; body text cream `#F8F6F2`. Every slide starts with a full-bleed dark rectangle.
- **One accent color per slide.** Per-beat palette:
  - `P1` purple `#9B8FE8` — intelligence / Beat 1 / Beat 5
  - `P2` green `#5FC8A0` — problems AI solves / Beat 4 / Beat 6 / journal-automate
  - `P3` ochre `#E8A062` — backwards walk / Beat 3
  - `WHY` amber `#D49658` — ceremonial / Beats 2, 7
  - The video border, sidebar eyebrow, and sidebar accent rule all share that slide's color.
- **Video-dominant layout** on every video slide:
  - Title left-aligned at top, 44pt Georgia.
  - Video placeholder at `x=0.6", y=2.0", w=8.4", h=4.6"` (~63% slide width, near-16:9). Big `▶` glyph + label + search hint inside.
  - Sidebar at `x=9.2", w=3.55"`: small-caps eyebrow (11pt accent) → 0.6" colored rule → 3–4 bullets (13pt) → optional muted italic footnote for source citations.
  - Footer at `y=7.0"`, 10pt muted, left-aligned: `INTRO TO AI · SESSION 1`.
- **Inverted hero card at Beat 7** (the unanswered question): cream card on dark slide. This is the visual climax — don't change it.
- **Fonts:** Georgia (heading) + Helvetica Neue (body). 16:9 widescreen, 13.333 × 7.5 inches.
- **"Cards" = physical index cards.** When the spine talks about "writing on a card" (Beat 2 wishes, Beat 5C intelligence answers, Beat 7 the big question card), it means real physical cards Sharpie'd by Shreya and the parents. NOT slide elements. They get photographed → exhibition site.

## Session 1 — current state

21 slides:

1. **Title** — "What even is AI?" (120pt centered) + date only
2. **Beat 1A** — wow reel: Waymo + sidebar "what you might cue up"
3. **Beat 1B** — "AI is mimicking us." + AlphaGo move 37 + milestones sidebar
4. **Beat 2** — no-video twin-prompt card: "what do you wish you didn't have to do?"
5–11. **Beat 3a–g** — backwards walk reframed as "problem it solved": video call → telephone → recorded music → camera → printing press → writing → fire
12. **Beat 3h** — the calculator (seed for Beat 7's payoff)
13. **Beat 4a** — 🐋 Talking to whales (Project CETI)
14. **Beat 4b** — 👓 Helping a blind person see (Be My Eyes / Meta glasses)
15. **Beat 4c** — 🤖 Surgery robot (Johns Hopkins SRT-H)
16. **Beat 5a** — Experiences can't be outsourced
17. **Beat 5b** — Children are the great explorers (Gopnik)
18. **Beat 5c** — open question "So… what IS intelligence?" (no video)
19. **Beat 6** — what AI still can't do + Beat-2 callback
20. **Beat 7** — "If AI can do so much, why bother learning anything?" hero card + calculator/car analogy keys
21. **Wrap** — AI Journal: 3 automate / 3 keep human

## Accuracy guardrails — preserve these exact phrasings

- **Math milestones:** "silver-medal **STANDARD**" (2024 Google AI, 28/42); "gold-medal **STANDARD**" (2025 Gemini Deep Think, 35/42, officially graded). NEVER "won gold" — AI was not an official IMO contestant.
- **Whales:** "researchers **believe** they've found a phonetic alphabet"; "we're **beginning** to understand them." NEVER "we can talk to whales."
- **Surgery robot:** "**lifelike model**, NOT a living person." Frame as helping doctors.
- **Blind glasses:** keep the "**lasagna**" miss — it's the honesty beat.
- **Robot prices** (time-sensitive — re-verify before each session): Weave Isaac 0 ≈ $7,999 (~$450/mo); 1X NEO ≈ $20,000 (~$499/mo). Both still require remote human help.
- **Gopnik:** "best learning machines in the universe" is an accurate quote; "cultural technologies" framing is from her 2025 *Science* paper.
- **General principle:** hedge every AI-capability claim. The curriculum was peer-reviewed; earlier drafts contained factual errors.

## Reusable pattern for Sessions 2–13

When asked to build a deck for another session:

1. Pull that session's `<details class="notes">` block from `index.html` — that's the build-ready content.
2. Ask Kaustav whether he's written a **rebuild brief** (Markdown file, usually in `~/Downloads/`). Format: spine sentence + beat-by-beat content with video search hints, sidebar bullets, speaker notes, accuracy guardrails. Use it as the source of truth over `index.html` if provided.
3. Create `session-N/` directory. Copy `session-1/build_deck.py` → `session-N/build_deck.py`. Reuse the helpers: `add_blank`, `add_title`, `add_video_placeholder`, `render_sidebar`, `add_footer`, `slide_title`, `video_slide`. Replace the build loop with the new session's beats. The output path inside the script — `Path(__file__).parent / "deck.pptx"` — needs no change; it writes alongside the script.
4. Update the footer string in `add_footer()` to `INTRO TO AI · SESSION N`.
5. Keep every design convention above. Same palette, layout, dark mode.
6. After building, verify by reading the .pptx back via python-pptx (slide count, title text, speaker-notes preambles). NEVER report done without this.
7. Tell Kaustav the path (`session-N/deck.pptx`); he uploads.

## How Kaustav talks about this work

- **Direct & opinionated.** He wants strong recommendations, not menus. Lead with your pick + reasoning.
- **Hates over-engineering.** Simplest approach first. Minimal diffs. No unused imports, no premature abstractions.
- **Iterative.** He'll give targeted feedback in tight bursts. Don't rebuild more than asked.
- **Visual quality matters.** Public-facing site + class deck. Design choices need to be deliberate.
- **"Cards"** in conversation = physical index cards (see above), not slide cards.
