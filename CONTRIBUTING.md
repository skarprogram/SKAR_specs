# Contributing to SKAR_specs

Thank you for helping improve SKAR_specs.

## Before opening a change

- Search existing issues and pull requests to avoid duplicate work.
- Open an issue before a large behavioral or architectural change.
- Never include a generated inventory report, debug log, device identifier, product key, or other personal or organizational data.
- Use GitHub's private vulnerability-reporting flow instead of a public issue for security concerns.

## Development

1. Fork the repository and create a focused branch.
2. Open `SKAR_specs.sln` in Visual Studio 2022.
3. Build both `Debug` and `Release` configurations.
4. Exercise the HTML, JSON, and text outputs on Windows.
5. Confirm that error paths degrade gracefully when WMI queries, Registry keys, network access, or external commands are unavailable.

The code is organized as one partial `Program` class. Keep collection logic in the corresponding `Program.Section.*.cs` file, data contracts in `Program.Models.cs`, and output-specific rendering in the relevant exporter file.

## Pull requests

Keep changes small and explain their user-visible effect. Include validation steps and screenshots only after redacting sensitive data. By submitting a contribution, you agree that it is licensed under the repository's GPL-3.0-only license.
