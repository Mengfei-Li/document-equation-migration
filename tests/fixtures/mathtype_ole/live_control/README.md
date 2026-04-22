# MathType Live-Control Fixture

This fixture contains a minimal public test payload for the MathType live-conversion path.

It is not a real user document and does not include generated DOCX, PDF, WMF, or binary output files. The OLE payload is stored as base64 text and is decoded only into temporary test files.

## Source

- Upstream project snapshot: `transpect/mathtype-extension`
- Local audited commit: `c1f788c7857802193220d370894ae52a2ce40d6c`
- Original path: `ruby/mathtype-0.0.7.5/spec/fixtures/input/mathtype5/equation1.bin`
- SHA-256: `9f53c650efc68c5c94952892a5432a7bbc6966558a5cc7de6f7c0581ead14d4e`

The nested `mathtype` gem fixture is used under its MIT license. See `NOTICE.md` for attribution.

## Scope

This fixture provides a real binary MTEF OLE payload suitable for exercising live-conversion tooling when the documented external converter prerequisites are available.

The repository's default CI tests only verify fixture integrity, source detection, and temporary DOCX packaging. They do not run the external MathType converter by default.
