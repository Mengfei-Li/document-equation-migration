# Dependencies

## Required for Core Conversion

- Windows PowerShell.
- Python 3.11 or newer.
- Full Java JDK 17 or newer.
- Microsoft Office `MML2OMML.XSL`.
- `transpect/mathtype-extension`.
- `jure/mathtype_to_mathml`.

## MathType Live Conversion Prerequisites

MathType live conversion requires a full JDK, not only a JRE or a stripped runtime image.

The Java installation must provide:

- `java.exe`.
- `javac.exe` for the first-run compilation of `Ole2XmlCli.java`.
- the `jdk.charsets` module. The JRuby/Nokogiri extraction stage can require charsets such as `ISO-2022-JP`.

Known bad symptom:

```text
java.nio.charset.UnsupportedCharsetException: ISO-2022-JP
```

If this appears, use a full JDK distribution such as Temurin 17+ and set `JAVA_EXE` and `JAVAC_EXE` to that JDK's binaries, or pass the equivalent script parameters. Release documentation avoids machine-specific temporary JDK paths.

`MML2OMML.XSL` must come from a Microsoft Office installation. A common Office path is:

```text
C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
```

This repository does not vendor Office files.

The MathType converter sources are expected under `third_party/` after running the bootstrap script, or through explicit `MATHTYPE_EXTENSION_DIR` and `MATHTYPE_TO_MATHML_DIR` paths. Complete third-party repositories are not vendored into release artifacts without a license review.

`pandoc`, Word desktop, `PyMuPDF`, and `Pillow` are validation helpers:

- Missing `pandoc`: add `-SkipLatexPreview` when a LaTeX preview is not required.
- Missing Word desktop: Word/PDF deliverability cannot be claimed.
- Missing visual comparison packages: visual parity or a visual gate result cannot be claimed.

## Required for Validation

- Optional: `pandoc` for LaTeX preview generation.
- Microsoft Word desktop for optional PDF export through COM automation.
- Optional: `PyMuPDF` and `Pillow` for PDF visual comparison.

Install core Python dependencies with:

```powershell
python -m pip install -r requirements.txt
```

Install optional PDF visual comparison dependencies with:

```powershell
python -m pip install -r requirements-visual.txt
```

## Environment Variables

The scripts can use explicit parameters or these environment variables:

- `JAVA_EXE`: path to `java.exe`, or `java` if it is on `PATH`.
- `JAVAC_EXE`: path to `javac.exe`, or `javac` if it is on `PATH`.
- `MATHTYPE_EXTENSION_DIR`: path to `transpect/mathtype-extension`.
- `MATHTYPE_TO_MATHML_DIR`: path to `jure/mathtype_to_mathml`.
- `MML2OMML_XSL`: path to `MML2OMML.XSL`.

## Troubleshooting

| Symptom | Likely Cause | Resolution |
|---|---|---|
| `UnsupportedCharsetException: ISO-2022-JP` | Java runtime is missing `jdk.charsets`. | Use a full JDK 17+ and set `JAVA_EXE` / `JAVAC_EXE`. |
| `javac` is missing | A JRE or incomplete JDK is being used. | Install a full JDK and point `JAVAC_EXE` to `javac.exe`. |
| `MML2OMML.XSL` cannot be found | Office path is missing or non-standard. | Install Microsoft Office or set `MML2OMML_XSL`. |
| `mathtype-extension` or `mathtype_to_mathml` cannot be found | Third-party bootstrap has not run. | Run `scripts/bootstrap_third_party.ps1` or set the converter directory variables. |
| LaTeX preview fails because `pandoc` is missing | Optional preview dependency is unavailable. | Install `pandoc` or run with `-SkipLatexPreview`. |
| Word export is skipped or fails | Word desktop is unavailable or COM export failed. | Install Word desktop and rerun validation before claiming deliverability. |

## Third-Party Bootstrap

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\bootstrap_third_party.ps1
```

The script clones the required converter repositories under `third_party/` and applies `patches/mathtype_to_mathml-quality-fixes.patch`.

The patch is part of this repository because the public upstream `mathtype_to_mathml` output did not reproduce the quality fixes used by this integration pipeline.

## Why Large Runtimes Are Not Vendored

The public repository intentionally excludes JDK archives, portable JDK/JRE folders, generated outputs, and private sample documents. This keeps the git history small and avoids redistributing files whose licenses or ownership may not match this project.
