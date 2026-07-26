# DCSPad — visual design brief (for Claude Design)

## The prompt

> You are redesigning the visual layer of **DCSPad**, a SharePoint-native
> developer workbench (JSFiddle-style: code editors, live preview, console +
> network panels). It runs embedded in a SharePoint page, seated below
> SharePoint's black suite bar on a near-black surround. The current UI is a
> serviceable dark "DevTools-ish" theme that grew feature-by-feature; we want
> a deliberate design pass, not a re-architecture.
>
> Produce: (1) a **comp of the main workbench state** (match the layout in
> `01-main-after-run.png` — the layout is settled, the *treatment* is what
> changes), and (2) a **design-token sheet**: color palette as CSS custom
> properties, a syntax-highlighting palette for the code editors, a
> typography scale, and spacing/border/radius rules. Stay within the
> technical constraints below — everything you specify must be expressible
> as plain CSS custom properties + a Monaco theme.
>
> Priorities, in order:
> 1. **Contrast and readability.** The attached older-generation screenshot
>    (`reference-old-dcspad.png`) is included because its higher overall
>    contrast and **more colorful, more saturated syntax highlighting** read
>    better than the current build — carry that quality over. (Nothing else
>    about the old design is wanted.)
> 2. **Hierarchy.** Panel headers, tabs, section titles, the status bar and
>    the sidebar currently blur together in similar grays. Make scanning
>    effortless: where am I, what's interactive, what's status.
> 3. **A confident, current dark aesthetic** — this is a developer power
>    tool that lives inside corporate SharePoint; it should feel sharper
>    than its host, not decorated. No skeuomorphism, no gradients-as-paint;
>    accent color used with intent (currently teal `#4ec9b0` + blue
>    `#67a7f7` — keeping or replacing them is your call, but state it).
>
> Known sore points you should address explicitly:
> - Code editor syntax colors are muted (one-dark defaults on a darker
>   ground than one-dark expects) — this is the #1 complaint.
> - Sidebar rows (frameworks/snippets): checkbox, name, and row-action
>   icons (↑ ↓ ★ ✕) have been contrast-tuned twice by hand and still feel
>   unresolved.
> - The red error dots on the Console/Network tabs must stay legible but
>   sit more intentionally.
> - Buttons exist in three ad-hoc variants (primary Run, ghost icon
>   buttons, tiny `btn-xs`) — unify the system.
> - The empty-preview and empty-console states are dead air; make them
>   quietly branded.

## What DCSPad is (context for the comp)

Left sidebar (Resources: frameworks catalog + snippets), center editor
column (HTML/CSS/JS tabs, Monaco), right runtime column (Preview panel
over Console/Network panel), 40px topbar (logo · File · settings ……
Auto-run · Run · SP status chip), 24px status bar. Panels resize via
splitters; sidebar and the console panel collapse. On SharePoint it sits at
`inset: 53px 5px 5px` below the (desaturated) suite bar, corner radius 6px,
on a `#101216` surround. A boot splash (ASCII wordmark on the dark curtain)
crossfades into the app.

## Current design tokens (styles/app.css `:root`)

```css
--bg-0: #16181d;   /* deepest: page */
--bg-1: #1d2026;   /* panels */
--bg-2: #23262e;   /* heads / chrome */
--bg-3: #2b2f39;   /* hover / inputs */
--border: #33374233;
--border-strong: #3c4150;
--fg: #d6d9e0;  --fg-dim: #8b91a0;  --fg-faint: #5c6270;
--accent: #4ec9b0;   /* teal */
--accent-2: #67a7f7; /* blue */
--warn: #e2c08d;  --error: #f47067;  --ok: #7ecb89;
--mono: "Cascadia Code", "Consolas", "SF Mono", Menlo, monospace;
--sans: "Segoe UI", system-ui, sans-serif;
```

## Technical constraints (hard)

- Deliverables must compile to: CSS custom properties + plain CSS rules (no
  frameworks, no images/asset pipeline; inline SVG icons are fine), plus a
  **Monaco theme** for the editor (token→color list is
  enough: keywords, strings, numbers, comments, tags, attributes,
  properties, functions, operators, punctuation, plus editor ground:
  background, gutter, active line, selection, cursor, matching bracket).
- Dark theme only (the *preview iframe* has its own light/dark toggle —
  out of scope).
- The workbench must hold up from ~1000px wide to ultrawide; column widths
  are user-draggable, so treatments can't depend on fixed widths.
- System fonts only (see `--mono`/`--sans` above).
- Red error dots on Console/Network tabs and the SP Live/Mock chip
  (green/amber) are functional signals — restyle, don't remove.
- Accessibility: body text ≥ 4.5:1 against its ground; the current
  `--fg-dim`-on-`--bg-1` pairs sit near the line — improve, don't regress.

## Attachments

| file | what it shows |
| --- | --- |
| `screenshots/01-main-after-run.png` | main workbench, the state to comp |
| `screenshots/02-file-menu.png` | File dropdown |
| `screenshots/03-settings-menu.png` | settings dropdown (incl. text-size stepper) |
| `screenshots/04-error-state-dots.png` | console output, inspector tree, REPL, both error dots |
| `screenshots/05-collapsed-panels.png` | sidebar + console collapsed (docked tabs bar) |
| `reference-old-dcspad.png` | **older generation** — include for its contrast + colorful editor ONLY |

If the reference image is missing, this is what it shows (an earlier DCSPad
generation on another tenant): notably higher contrast between surfaces and
text; vivid multi-hue syntax highlighting (bright oranges/blues/greens on
near-black) that reads effortlessly; colored identity dots on the editor
tabs (HTML red, CSS blue, JS yellow, snippets green); a yellow RUN button;
a PREVIEW header carrying a green ✓ last-run timestamp; console filter
toggles as labeled colored checkboxes with COPY/CLEAR actions. The *layout*
of that generation is not wanted — only its contrast energy and editor
colorfulness.

(Hosted-on-SharePoint framing — suite bar above, dark surround — is
described above; the screenshots are the standalone build, which is
pixel-identical inside the frame.)

## Out of scope

Layout re-architecture, new features, light theme, the preview iframe's
content area (user-controlled), the splash ASCII art (keep-able as-is
unless a small treatment suggestion is cheap).
