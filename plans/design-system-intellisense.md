# Design-system and Alpine intelligence plan

Status: composable registry, Alpine core pack, generated CSS tokens, and
generated HTML class intelligence complete. Component snippets and icons remain.

## Inputs now available

The imported packages are unusually well suited to generated Monaco data:

- `bsp-design-system/colors_and_type.css` and `editorial.css` expose about 172
  documented CSS custom properties.
- `components.css` and `editorial.css` expose about 369 distinct component,
  element, modifier, state, and utility class names.
- `examples/components.html` and `examples/editorial-components.html` contain
  47 copy-ready `<pre>` specimens that can become snippets.
- `docs/TECHNICAL-REFERENCE.md` defines component composition, accessibility,
  and Alpine state contracts in one place.
- `bmo-icons.svg` is the curated 48-symbol functional sprite.
- `bsp-fluent-icon-lib/fluent-font-library.json` contains 2,654 icon records
  with names, sizes, styles, descriptions, metaphors, real filenames, and font
  codepoints. It is about 6.9 MB, so the browser should not load it directly.

The locations and important filenames are editable under `assets` in
`dcspad.config.json`. Local relative URLs are useful for review/development;
eventual SharePoint folder URLs can be placed in `hostedBaseUrl`.

## Generated artifact

The development-only generator under `tools/` reads the configured local
sources and emits a small, versioned set under `vendor/intelligence/`. The
current token/class phase uses one compact `bsp-design.json` payload plus
`manifest.json`; later phases can split independently if their payload size or
release cadence warrants it:

1. `bsp-html-data.json`
   - `<fluent-icon>` attributes and values.
   - Alpine `x-*`, `@*`, `:*`, and common magic-property documentation.
   - Design-system-specific HTML attributes where applicable.
2. `bsp-css-data.json`
   - Custom properties with nearby source comments as hover documentation.
   - Editorial properties identified as scoped to `.editorial`.
3. `bsp-classes.json`
   - Canonical BEM classes, grouped by block and annotated as base, element,
     modifier, utility, or state.
   - Retired vocabulary retained only as diagnostic guidance, never as a
     completion.
4. `bsp-snippets.json`
   - Cleaned component specimens converted to Monaco snippets with tab stops.
   - State-contract variants that include the required `x-data`, ARIA, and
     `x-cloak` wiring.
5. `bsp-icons.json`
   - A compact projection of the 2,654-record catalog: searchable display
     name, normalized token, description/metaphors, and only real variants.
   - Curated-sprite membership marked separately so completion can distinguish
     “available in `bmo-icons.svg`” from “requires full library/font.”
6. `manifest.json`
   - Schema version, source hashes/versions, file URLs, counts, and enabled
     intelligence-pack IDs.

The generator is allowed to use Node tooling: the design system's “buildless”
rule constrains its shipped pages, not development tools. DCSPad should load
only the generated compact data, never scrape CSS/examples or parse the 6.9 MB
icon catalog at runtime.

## Monaco integration

Before registering these sources, replace the PnP-only `setExtraLibs` switch
with a composable intelligence registry:

```text
pnpjs-2.15.0 ─┐
alpine-3 ──────┼─> active packs ─> flattened JS extraLibs
bsp-design ────┘                  + HTML/CSS/custom completion providers
```

Each enabled framework/catalog entry contributes its explicit `intelligence`
IDs from `dcspad.config.json`; multiple packs must coexist.

Provider responsibilities:

- JavaScript declarations: `Alpine`, `Alpine.data`, `Alpine.store`, magic
  properties usable in ordinary JavaScript, and any future design-system JS.
- HTML data: static Alpine directives, `<fluent-icon>`, attributes, values,
  hover documentation, and links.
- Contextual HTML completion: BEM classes inside `class=""`, icon tokens inside
  `<use href="">` and `<fluent-icon name="">`, plus complete component snippets.
- CSS completion: `var(--...)` tokens and canonical class selectors.
- Diagnostics: retired class names, an Alpine binding outside an `x-data`
  ancestor, missing `x-cloak` on common `x-show` overlays, icon variants that
  do not exist, and sprite-only names absent from the curated sprite.

## Delivery phases

1. **Complete:** composable intelligence-pack registry and explicit framework
   metadata.
2. **Complete:** Alpine global declarations, HTML directive data, contextual
   shorthand/magic completion, hovers, and editable snippets.
3. **Complete for requested scope:** generated design tokens and canonical
   class completion/hover. `<fluent-icon>` metadata moves with the icon phase.
4. Component/state snippets from the canonical specimens.
5. Compact icon index with contextual token/markup completion.
6. Optional deeper Alpine expression analysis (`x-data` members, `$refs`,
   `$store`) after the static layer is proven useful.

Every phase needs standalone and hosted tests. Hosted assets must remain
same-origin/versioned where workers or fetched metadata require it; configured
cross-folder SharePoint URLs must be tested for permissions and CORS before
being treated as production sources.
