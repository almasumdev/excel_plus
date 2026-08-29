# Security Policy

## Supported versions

Fixes land on the latest published version. Please reproduce on the newest
release before reporting.

| Version | Supported |
| ------- | --------- |
| Latest release | Yes |
| Anything older | No, please upgrade first |

## Reporting a vulnerability

Do not open a public issue for a security problem.

Use GitHub's private reporting on the
[Security tab](https://github.com/almasumdev/excel_plus/security/advisories/new),
or email dev.almasum@gmail.com. Include the package version, a description of
the issue, and a file or snippet that reproduces it if you have one.

You can expect an acknowledgement within a few days. If the report is confirmed,
a fix will be published and the advisory credited to you unless you would rather
stay anonymous.

## Scope

This package parses files. The realistic risk is a malformed or hostile workbook
reaching your process, so that is what to look at:

- A `.xlsx` file is a zip archive. A crafted archive can try to expand far beyond
  its compressed size, or name entries that escape the intended directory.
  Reports about resource exhaustion or path handling during decode are in scope.
- Worksheet XML is parsed with `package:xml`. External entity and
  billion-laughs style attacks against the XML layer are in scope.
- Formulas are evaluated by this package, not by a spreadsheet application.
  Evaluation is pure computation: it does not read files, open sockets, or shell
  out. A formula that can reach outside that boundary is a vulnerability.
- Legacy `.xls` is parsed from a binary compound document, where a malformed
  record can drive an out-of-range read. Those reports are in scope.

Out of scope:

- Macros. `.xlsm` files round-trip their `vbaProject.bin` untouched, and this
  package never executes it. What your application does with an extracted macro
  is your decision.
- Anything the produced file does once it is opened in Excel. If you write
  attacker-controlled text into a cell, treat CSV and formula injection in the
  consuming application as your responsibility: a leading `=`, `+`, `-` or `@`
  is a formula to Excel, and this package stores what you give it.
- Vulnerabilities in Excel, Google Sheets or LibreOffice themselves. Report
  those to the relevant vendor.
