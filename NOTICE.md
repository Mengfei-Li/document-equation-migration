# Notices

This repository contains original integration scripts for a MathType OLE to MathML / OMML conversion workflow.

The current public tree does not vendor large runtime bundles, sample exam documents, generated DOCX/PDF outputs, or third-party repositories.

Runtime integration currently expects these external projects when using the open-source conversion path:

- `transpect/mathtype-extension`
- `jure/mathtype_to_mathml`

Those projects and their dependencies are governed by their own licenses. If you vendor, redistribute, or modify them, preserve their notices and license files.

This repository includes a small patch file for `jure/mathtype_to_mathml` under `patches/`. The patch is distributed as part of this repository's MIT-licensed integration code, but it applies to a third-party codebase whose original license terms still apply.

`MathType`, `Microsoft Word`, `Office`, and related product names are trademarks of their respective owners. This project is not affiliated with or endorsed by those owners.
