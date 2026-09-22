# Info chip — proposal to promote into the DCS Workbench design system

**Status:** implemented locally in `styles/workbench.css` as `.wb-info-chip`;
**not** yet in the design system. This file is the promotion proposal, written
so someone can apply it to `whywhyjoe/dcs-workbench-tools` without rediscovering
the reasoning. Nothing here has been applied to that repo.

## The gap

The design system ships two chip-shaped components, and neither fits a quiet
classification label:

| Component | Register | Treatment |
| --- | --- | --- |
| `.dcs-badge` | file type | 9px mono, uppercase, `.06em`, filled, `--ft-*` hue per type, `--radius-s` |
| `.dcs-chip` | **status** | 999px, 11px mono, `.04em`, filled, optional `-dot`, `ok`/`warn`/`err` signal variants, `cursor: pointer` |

Both are loud on purpose. `.dcs-badge` answers "what kind of file is this" with
colour; `.dcs-chip` answers "how is this doing" with a signal. Neither answers
"what kind of thing am I looking at" — a fact the reader wants available but
does not need pushed at them.

The readme's own rules predict the problem. "Colour is information. Signals mean
status" (§Colour) and "Signal colors are status, never decoration" (rule 2) both
say a classification must not borrow a signal colour. But with only two chips
available, anything chip-shaped that is *not* status ends up wearing a status
costume anyway. That is the gap.

## How it was found

The SP Workbench Pages view names the pages library it resolved — modern Site
Pages (BaseTemplate 119), classic publishing Pages (850), or a library merely
titled "Pages". First implementation reached for the badge register: 9px mono
uppercase, `--accent-soft` for modern, neutral for classic.

That read as a state the operator had to act on. It is not: it classifies the
view, and once read it should recede. Meanwhile the same view's page-kind chip
(`modern canvas page`, `classic publishing page`, `page with no readable body`)
had *already* solved this correctly — quiet sans, hairline, no fill — and had
been sitting in `styles/workbench.css` as a one-off since the drilldown shipped.

So the pattern was already in the system twice; it just had no name, and the
newer instance drifted to the wrong register precisely because there was nothing
to compose from. Rule 6 ("compose from existing classes") cannot be followed for
a component that does not exist.

## Proposed component

```css
/* Quiet classification: what a thing IS, not how it is doing. */
.dcs-chip-info {
  flex: none;
  font: 11px/1.5 var(--sans);
  color: var(--fg-faint);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 1px 8px;
  white-space: nowrap;
}
.dcs-chip-info[title] { cursor: help; }
```

Every value is an existing token; no new tokens are needed.

### Why it is a sibling of `.dcs-chip`, not a modifier

A `.dcs-chip-info` written as `.dcs-chip.dcs-chip-info` would have to unset the
999px radius, the fill, the mono face, the letter-spacing and `cursor: pointer`
— it overrides more than it keeps. It is a separate component that happens to be
chip-shaped, and it deliberately sits under the `--radius-l` 6px…999px scale at
**10px**, which is the one value in it not already assigned by the readme's
three-radii rule.

**Open question for the maintainer:** either add 10px to the radius scale as
"info chip", or restate the info chip at `--radius-l` (6px) and accept a
squarer shape. The 10px came from the existing page-kind chip; it was chosen by
eye before the radius rule was written, so it is legitimately up for revision.
Everything else in the component is already conformant.

### Rules to ship with it

1. **Sentence case.** It is a phrase, not a label — this is the readme's
   existing rule (§Voice: sentence case everywhere except panel/group titles and
   chip *labels*; an info chip carries a phrase, so it takes sentence case).
2. **Never colour a variant.** If a value needs a signal colour it is status,
   and it belongs in `.dcs-chip`. This is the rule the whole component exists to
   enforce.
3. **The chip is a summary; the sentence goes on `title=`.** Both live instances
   put the full explanation there ("classic publishing page — no canvas sections
   or columns, so the Structure tab does not apply").
4. **Non-interactive.** No hover state, no click. `cursor: help` only when a
   `title` is present — attribute-scoped so a chip without one does not lie
   about having more to say. This also fixes, for this component, the known
   `.dcs-chip` wart the readme lists at §Known issues 6 ("`.dcs-chip` is
   `cursor: pointer` even when the chip is not clickable").

## Where it would land in `dcs-workbench-tools`

| File | Change |
| --- | --- |
| `design-system/dcs-workbench.css` | add `.dcs-chip-info` in the feedback block, next to `.dcs-chip` |
| `design-system/assets/dcs-workbench.standalone.css` | same, to keep the standalone build byte-comparable |
| `design-system/readme.md` | add `Chip` → `ChipInfo` to the feedback component list (§Components); add the status-vs-info distinction to §Colour; note the radius decision in §Space, shape, elevation |
| `docs/05-design-systems.md` | one line under "Rules with teeth": classification is not status |

Do **not** hand-edit any vendored copy (Halo's) — per that repo's rule 7 it
re-syncs from the standalone build.

## Adoption in this repo, once promoted

`styles/workbench.css` keeps `.wb-info-chip` as a thin alias of
`.dcs-chip-info` (or the markup switches to the design-system class outright, if
the workbench is linking the sheet by then). Until promotion, `.wb-info-chip` is
the single local definition, and every instance composes it so they cannot
drift apart in the meantime: `.wb-lib-kind` in the Pages toolbar,
`.wb-detail-kind` in the page drilldown, `.wb-reduced-chip` for a
schema-rejected query, and the Lists Schema tab's
`.wb-schema-kind`/`-fields`/`-views`/`-cts`/`-versioning`/`-libkind` (five
quiet facts about a list, not states to act on). The Schema tab's "Copy
to…"/"New from schema…" apply dialog (`list-schema-dialog.js`) deliberately
does **not** add any info-chip instances of its own — its step list
(`.wb-schema-steps`) is the other register, STATUS, and takes signal colour
on purpose (running/done/failed/skipped), never composing `.wb-info-chip`.
