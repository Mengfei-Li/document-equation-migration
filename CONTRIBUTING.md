# Contributing

Contributions are welcome, especially reproducible conversion failures, minimal fixtures, and improvements to the MathML normalization rules.

## Before Opening an Issue

- Confirm the document contains MathType OLE objects, not only images.
- Re-run with the latest `main` branch.
- Include the generated `summary.csv`, `pipeline_summary.txt`, and `risk_analysis.txt` when possible.
- Do not upload copyrighted documents unless you have permission to share them publicly.

## Pull Requests

- Keep changes focused.
- Add or update a minimal fixture when fixing a conversion bug.
- Avoid adding large generated outputs to git.
- Do not vendor third-party binaries unless the license and size implications have been reviewed.

## Development Checks

```powershell
python -m compileall .
python -m pip install -r requirements.txt
```

Install `requirements-visual.txt` only when testing PDF visual comparison.

Full DOCX conversion tests require Windows, Java, Office's `MML2OMML.XSL`, and the external converter repositories.
