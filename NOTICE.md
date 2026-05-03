# Notices

This repository contains original integration scripts for a MathType OLE to MathML / OMML conversion workflow.

The current public tree does not vendor large runtime bundles, sample exam documents, generated DOCX/PDF outputs, or third-party repositories.

Runtime integration currently expects these external projects when using the open-source conversion path:

- `transpect/mathtype-extension`
- `jure/mathtype_to_mathml`

Those projects and their dependencies are governed by their own licenses. If you vendor, redistribute, or modify them, preserve their notices and license files.

This repository includes one minimal base64-encoded MathType OLE test payload under:

- `tests/fixtures/mathtype_ole/live_control/word/embeddings/oleObject1.bin.b64`

The payload is derived from `ruby/mathtype-0.0.7.5/spec/fixtures/input/mathtype5/equation1.bin` in the `transpect/mathtype-extension` source tree at commit `c1f788c7857802193220d370894ae52a2ce40d6c`. That tree incorporates the `mathtype` Ruby gem fixture corpus. The nested gem files identify the material as MIT licensed with copyright notice:

- `Copyright (c) 2015 Jure Triglav`

The payload SHA-256 is:

- `9f53c650efc68c5c94952892a5432a7bbc6966558a5cc7de6f7c0581ead14d4e`

This repository includes a small patch file for `jure/mathtype_to_mathml` under `patches/`. The patch is distributed as part of this repository's MIT-licensed integration code, but it applies to a third-party codebase whose original license terms still apply.

This repository includes minimal Equation Editor 3.0 `Equation Native` stream controls derived from Apache POI test data under:

- `tests/fixtures/equation_editor_3_ole/apache_poi/`

Upstream provenance:

- Repository: `apache/poi`
- Commit pinned: `e6a04b49211e23c704fcdbe524d99d2f4486b083`
- Upstream source document paths:
  - `test-data/document/Bug61268.doc`
  - `test-data/document/Bug50936_1.doc`

The upstream source document SHA-256 values are:

- `Bug61268.doc`: `24b47ad892871c80495c64b440c24200dbcaa16c367faf780aec43e3bf8ddae6`
- `Bug50936_1.doc`: `9169e7a5f5aa865579ceb803bd91bd840134e9292c8a56e81c35913b47e218c7`

The public repository vendors only selected hex-encoded native streams, not the full upstream `.doc` files. The selected controls are redistributed under Apache-2.0 with source attribution recorded in:

- `tests/fixtures/equation_editor_3_ole/apache_poi/SOURCES.json`
- `tests/fixtures/equation_editor_3_ole/apache_poi/README.md`
- `THIRD_PARTY_LICENSES/Apache-POI-LICENSE.txt`
- `THIRD_PARTY_NOTICES/Apache-POI-NOTICE.txt`

The Apache POI notice is retained in `THIRD_PARTY_NOTICES/Apache-POI-NOTICE.txt`.

`MathType`, `Microsoft Word`, `Office`, and related product names are trademarks of their respective owners. This project is not affiliated with or endorsed by those owners.
