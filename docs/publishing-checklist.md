# Publishing Checklist

Use this checklist before making a public GitHub release.

## Repository Hygiene

- [ ] No private documents.
- [ ] No generated DOCX/PDF/images.
- [ ] No generated cache, compiled bytecode, compiled Java classes, or scratch output directories.
- [ ] No JDK/JRE archives.
- [ ] No vendored third-party repositories unless licenses are reviewed.
- [ ] No local absolute paths in public documentation.
- [ ] No secrets or credentials.
- [ ] Third-party bootstrap and patch application are documented.
- [ ] No machine-specific temporary JDK, Office, or third-party source paths are required by the public instructions.
- [ ] `pyproject.toml`, `src/`, and `tests/` are included when the release surface is the packaged CLI.
- [ ] Minimal synthetic detector fixture binaries are limited to explicitly allowlisted test fixture paths.
- [ ] Fixture binaries are not generated conversion outputs and do not contain private, copyrighted, or real user documents.

## Documentation

- [ ] README explains status and limitations.
- [ ] README distinguishes automated deliverable candidates from manual-review candidates.
- [ ] `docs/research-preview-release-notes.md` is updated for the release tag.
- [ ] LICENSE is present.
- [ ] NOTICE describes third-party dependencies.
- [ ] SECURITY.md explains vulnerability reporting.
- [ ] CONTRIBUTING.md explains fixture and data-safety expectations.
- [ ] `docs/dependencies.md` documents the full JDK, Office `MML2OMML.XSL`, third-party bootstrap, and optional validation dependencies.
- [ ] Release notes do not claim lossless or pixel-identical MathType conversion.
- [ ] Release notes state that current guarded MathType outputs are manual-review candidates, not automated production-ready deliverables.
- [ ] Release notes distinguish public marker fixtures from live-convertible fixture evidence.

## Validation

- [ ] `python -m compileall .`
- [ ] `python -m pip install -r requirements.txt`
- [ ] Java dependency check uses a full JDK with `java.exe`, `javac.exe`, and `jdk.charsets`.
- [ ] Office `MML2OMML.XSL` is available through an installed Office path or `MML2OMML_XSL`.
- [ ] `powershell -ExecutionPolicy Bypass -File .\scripts\bootstrap_third_party.ps1`
- [ ] MathType converter source directories are available through `third_party/` or explicit environment variables.
- [ ] Smoke conversion with a redistributable fixture.
- [ ] If the smoke fixture is marker-based, describe it as detector/routing/dry-run evidence only.
- [ ] If a live-convertible fixture is used publicly, record source SHA-256, license, NOTICE attribution, and why the fixture is redistributable.
- [ ] Risk analysis generated for the smoke output.
- [ ] Any MathType `review-gated` output is described as a manual-review candidate, not automated deliverable output.
- [ ] For any manual-review candidate, keep validation evidence covering Word export, conversion/replacement counts, page count, unmatched pages, and visual changed-ratio metrics.

## Release

- [ ] Use a pre-1.0 tag such as `v0.1.0-research-preview`.
- [ ] Mark known limitations in the release notes.
- [ ] Use "research preview" and "manual-review candidate" language for MathType outputs unless the automated visual gate passes.
- [ ] Do not attach private documents or generated large artifacts.
