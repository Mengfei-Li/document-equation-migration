# MathType Evidence Pack

This page explains what the current research-preview release can and cannot claim about the MathType route.

## Release-Facing Conclusion

The current strongest live route in this repository is:

```text
MathType OLE -> MathML -> OMML -> editable Word equations
```

For the current evidence set, that route should be treated as a **manual-review candidate**, not as an automated deliverable claim.

That means:

- editable OMML output can be produced on the current validated samples
- Word export has passed on the current validated samples
- conversion and replacement can complete for the intended MathType objects
- a human reviewer must still inspect layout drift and high-risk formulas before production use

This release does **not** claim lossless conversion, pixel-identical layout, or universal MathType coverage.

## What Has Been Validated

| Area | Current evidence | Release-facing meaning |
|---|---|---|
| Document-level conversion | Two maintainer-validated local MathType documents convert to editable OMML and export through Word. | The route works on the current real evidence set. |
| Layout preservation | The guarded layout-preservation option restores page count on both current real samples. | Layout stability improved enough to enter manual review, but not enough to claim print identity. |
| Visual review envelope | On the current guarded real-sample outputs, page counts match, unmatched pages are `0`, and the observed max changed ratio stays below `0.09`. | The current evidence supports a conservative manual-review envelope, not a universal pass/fail threshold. |
| Execution surface | The detector-first execution surface can route MathType work and drive the existing conversion path when external tools are explicitly allowed. | MathType is not an isolated script only; it is wired into the shared execution model. |
| Long-run behavior | Resume and chunk controls exist for long MathType runs. | Larger documents can be processed with a more stable recovery path. |
| Public detector fixtures | Public marker fixtures cover detection, routing, and dry-run evidence. | Public fixtures prove source discovery behavior, not full live conversion. |
| Public live-control fixture | A minimal attributed base64-encoded MTEF5 OLE payload is included for integrity, detection, temporary DOCX packaging, and optional external-tool conversion tests. | The public tree now has source-controlled material for live-conversion testing; the external conversion test is skipped by default unless explicitly enabled. |
| Local live-conversion control | A separate local research-control run proved that the external toolchain can complete end-to-end live conversion on the same payload class. | The live chain works locally, while public release claims remain manual-review gated. |

## What Users Can Claim Today

With the current evidence, users can reasonably say:

- the repository contains a research-preview MathType migration route
- the current route can produce editable OMML output on the current validated evidence set
- Word export has passed on the current validated real samples
- guarded layout preservation improves pagination stability on the current validated real samples
- the current best MathType result is a manual-review candidate

## What Users Should Not Claim Today

The current evidence does not justify these statements:

- "MathType conversion is lossless"
- "The output is pixel-identical to the source document"
- "All MathType documents are supported"
- "The guarded layout option is universally safe"
- "Public fixtures already prove the full live MTEF conversion path"
- "A successful run means the output is production-ready without review"

## How The Evidence Is Organized

The current release intentionally separates three evidence classes.

### 1. Public Repository Evidence

Publicly inspectable evidence lives in this repository:

- source code for the MathType route and shared execution surface
- detector and executor tests
- synthetic marker fixtures under `tests/fixtures/`
- an attributed `live_control` MathType fixture stored as base64 text under `tests/fixtures/mathtype_ole/`
- outward-facing documentation such as:
  - [Research-preview release notes](research-preview-release-notes.md)
  - [Dependencies](dependencies.md)
  - [Limitations](limitations.md)

This evidence is suitable for public inspection, CI, and discussion.

### 2. Local Non-Redistributable Validation Evidence

Some of the strongest current MathType evidence comes from real local documents that are not published in this repository.

Those documents are used to validate:

- document-level conversion completion
- Word export
- page-count recovery
- unmatched-page checks
- visual changed-ratio reporting

This is why the release remains research-preview and manual-review oriented. The project is being explicit about what is currently validated, but not pretending that every supporting input can be redistributed publicly.

### 3. Local Research-Control Live Fixtures

The project also uses a separate local research-control fixture class for proving the live MTEF toolchain itself.

That class exists because:

- public marker fixtures are intentionally minimal and safe to redistribute
- marker fixtures are good for detector and routing tests
- marker fixtures do not prove that a real binary MTEF payload will convert end to end

The public `live_control` fixture now covers the minimal attributed source-material side of that requirement. Full live conversion still requires the documented external tools and should be interpreted conservatively.

The repository includes an optional pytest path for that external conversion check. It is skipped by default and runs only when `DEM_RUN_EXTERNAL_MATHTYPE_TESTS=1` and the documented converter paths are provided.

## Manual-Review Gate

For the current research-preview release, a MathType output can enter manual review only when all of the following are true:

- Word export passes
- conversion and replacement complete for the intended MathType objects
- source and converted page counts are equal
- unmatched pages are `0`
- visual drift is documented for reviewer inspection

Current evidence supports this conservative envelope for the current validated samples. It does not establish a universal threshold for every future document.

## Why Layout Drift Still Matters

MathType OLE objects often carry fixed preview extents. Editable OMML does not preserve those extents in the same way, so layout can change even when the math remains editable and structurally useful.

In practice, the current route should be understood as:

- good for extracting editable equations
- promising for document migration
- still review-gated for layout-sensitive publishing workflows

Editable Word math and print-identical pagination are different goals.

## How To Reproduce The Publicly Inspectable Part

You can reproduce the public, inspectable portion of the evidence by:

1. reading [Dependencies](dependencies.md)
2. bootstrapping the third-party converter sources
3. running the detector-first scan / execution-plan flow
4. executing the MathType route with explicit external-tool approval
5. validating outputs with `dem validate-docx`
6. running the public test suite

Start with the commands in the [README](../README.md), then use the release notes and limitations pages to interpret the result conservatively.

## Current Gaps

The next evidence upgrade needs one or more of the following:

- broader independent real-sample coverage
- broader public live-conversion tests and CI jobs that run only when external converter prerequisites are explicitly available
- stronger evidence that the guarded layout option generalizes beyond the current validated samples
- stricter visual parity on real documents

Until then, the correct release posture is still:

**research preview, manual-review candidate, non-lossless, non-pixel-identical**
