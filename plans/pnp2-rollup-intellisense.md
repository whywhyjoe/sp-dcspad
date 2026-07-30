# PnPjs `pnp2` rollup support to-do

Status: **implemented on 2026-07-29**.

## Goal

Allow the PnPjs 2.15 intelligence pack to work with a custom, compatible
rollup that exposes its global as `pnp2` instead of `pnp`. The rollup includes
additional PnP modules, but typing those additional exports is outside the
initial scope.

## Implemented behavior

- `probeGlobal` is a runtime fallback check. Setting it to `pnp2` correctly
  verifies that the custom rollup loaded.
- `intelligence: ["pnpjs-2.15.0"]` activates the existing PnPjs declaration
  graph independently of the script URL.
- The generated declaration bridge declares both `pnp` and `pnp2` as aliases
  of the exact PnPjs 2.15.0 declaration graph.
- Entries in `dcspad.config.json` override catalog entries with matching IDs;
  they do not create new framework rows.

## To-do

- [x] Add `pnp2` as a second alias of `typeof import("@pnp/pnpjs")` in the
  generated global declarations in `tools/build-monaco.mjs`.
- [x] Add `pnp2` to the generated `Window` interface while retaining `pnp` for
  the classic bundle.
- [x] Extend the build-time type probe to validate a representative
  `pnp2.sp.web...` call.
- [x] Rebuild `vendor/monaco/pnpjs-types.json` with
  `npm run build:monaco` from `tools/`.
- [x] Extend `tests/monaco.mjs` to verify completion for `pnp2` and to
  ensure the declaration pack still unloads when no matching runtime is
  enabled.
- [x] Configure and test the mirrored custom rollup:

```json
"pnpjs2": {
  "localUrl": "./lib-mirror/pnp2.bundle.js",
  "cdnUrl": "",
  "probeGlobal": "pnp2",
  "intelligence": [
    "pnpjs-2.15.0"
  ]
}
```

## Catalog decision

The existing `pnpjs2` catalog ID now exclusively uses the mirrored
`pnp2.bundle.js`; there is no separate classic PnPjs entry or CDN fallback.
The maintained Alpine entry likewise exclusively uses the mirrored Alpine
3.15.2 file.

## Acceptance criteria

- Enabling the custom rollup makes `pnp2.sp.web` runnable in the preview.
- Monaco offers the same PnPjs 2.15 completions, signatures, hover, and
  diagnostics for `pnp2` that it offers for `pnp`.
- Existing projects using the classic `pnp` global remain unchanged.
- The configured runtime probe checks the global exposed by the selected
  rollup and does not trigger an incorrect fallback.
