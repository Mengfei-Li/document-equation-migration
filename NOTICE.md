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

`MathType`, `Microsoft Word`, `Office`, and related product names are trademarks of their respective owners. This project is not affiliated with or endorsed by those owners.
