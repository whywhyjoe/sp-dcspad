# PnPjs `pnp2` rollup support to-do

Status: **planned**.

## Goal

Allow the PnPjs 2.15 intelligence pack to work with a custom, compatible
rollup that exposes its global as `pnp2` instead of `pnp`. The rollup includes
additional PnP modules, but typing those additional exports is outside the
initial scope.

## Current behavior

- `probeGlobal` is a runtime fallback check. Setting it to `pnp2` correctly
  verifies that the custom rollup loaded.
- `intelligence: ["pnpjs-2.15.0"]` activates the existing PnPjs declaration
  graph independently of the script URL.
- The generated declaration bridge currently declares only `pnp`, so Monaco
  does not recognize otherwise-compatible calls made through `pnp2`.
- Entries in `dcspad.config.json` override catalog entries with matching IDs;
  they do not create new framework rows.

## To-do

- [ ] Add `pnp2` as a second alias of `typeof import("@pnp/pnpjs")` in the
  generated global declarations in `tools/build-monaco.mjs`.
- [ ] Add `pnp2` to the generated `Window` interface while retaining `pnp` for
  the classic bundle.
- [ ] Extend the build-time type probe to validate a representative
  `pnp2.sp.web...` call.
- [ ] Rebuild `vendor/monaco/pnpjs-types.json` with
  `npm run build:monaco` from `tools/`.
- [ ] Extend `tests/monaco.mjs` to verify completion/hover for `pnp2` and to
  ensure the declaration pack still unloads when no matching runtime is
  enabled.
- [ ] Test the custom rollup with this framework override:

```json
"pnpjs2": {
  "localUrl": "/path/to/pnp2-rollup.js",
  "cdnUrl": "",
  "probeGlobal": "pnp2",
  "intelligence": [
    "pnpjs-2.15.0"
  ]
}
```

## Catalog choice

For a straight replacement, keep the existing `pnpjs2` catalog ID and override
its URL and probe as shown above.

If classic `pnp` and custom `pnp2` need independent checkboxes, add a separate
catalog preset or imported catalog entry with a stable ID, then add a matching
configuration item for that ID. A configuration item alone will not create the
checkbox.

## Acceptance criteria

- Enabling the custom rollup makes `pnp2.sp.web` runnable in the preview.
- Monaco offers the same PnPjs 2.15 completions, signatures, hover, and
  diagnostics for `pnp2` that it offers for `pnp`.
- Existing projects using the classic `pnp` global remain unchanged.
- The configured runtime probe checks the global exposed by the selected
  rollup and does not trigger an incorrect fallback.
