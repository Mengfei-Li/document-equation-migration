# Security Policy

## Supported Versions

This project is in research preview. Security fixes target the `main` branch unless a release branch is explicitly announced.

## Reporting a Vulnerability

Suspected security vulnerabilities should not be reported through public issues.

Report privately through GitHub's private vulnerability reporting feature if it is enabled for this repository. If it is not enabled, contact the repository owner directly.

## Automated Secret Scanning

Pull requests, pushes, manual runs, and weekly scheduled runs are checked with Gitleaks against the full Git history.

Do not commit API keys, access tokens, passwords, private document samples, local machine paths, or private research artifacts. If a secret is committed, revoke or rotate it immediately; deleting it in a later commit is not enough because it remains visible in Git history.

Repository maintainers should also enable GitHub native secret scanning and push protection where available.

## Document Safety

This project processes Word documents and embedded OLE objects. Untrusted documents should be treated as potentially dangerous:

- Conversions should run in a disposable working directory.
- Untrusted converted documents should not be opened in Word outside a sandbox.
- Private or copyrighted documents should not be uploaded to public issues.
