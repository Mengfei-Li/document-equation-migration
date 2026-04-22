# Security Policy

## Supported Versions

This project is in research preview. Security fixes target the `main` branch unless a release branch is explicitly announced.

## Reporting a Vulnerability

Please do not open public issues for suspected security vulnerabilities.

Report privately through GitHub's private vulnerability reporting feature if it is enabled for this repository. If it is not enabled, contact the repository owner directly.

## Document Safety

This project processes Word documents and embedded OLE objects. Treat all untrusted documents as potentially dangerous:

- Run conversions in a disposable working directory.
- Do not open untrusted converted documents in Word outside a sandbox.
- Do not upload private or copyrighted documents to public issues.
