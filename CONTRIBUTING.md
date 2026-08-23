# Contributing to NIST NVD CVE 2.0 Converter

First off, thanks for taking the time to contribute! 🎉

The following is a set of guidelines for contributing to this project. These are mostly guidelines, not rules. Use your best judgment, and feel free to propose changes to this document in a pull request.

## How Can I Contribute?

### Reporting Bugs
- Use a clear and descriptive title for the issue.
- Describe the exact steps which reproduce the problem in as many details as possible.
- Explain which behavior you expected to see instead and why.

### Suggesting Enhancements
- Use a clear and descriptive title for the suggestion.
- Provide a step-by-step description of the suggested enhancement in as many details as possible.
- Explain why this enhancement would be useful to most users.

### Pull Requests
- Ensure the project continues to work on Windows, macOS, and Linux.
- Update the `README.md` if your change affects usage.
- Follow the existing code style (PEP 8).
- Run the documented dependency, compile, test, and `git diff --check` gates.
- Add a regression test for every repaired parser, feed, workbook, or pipeline contract.
- Do not commit generated feeds, workbooks, private fixtures, or credentials.

## Styleguide
- Use 4 spaces for indentation.
- Include docstrings for any new functions or classes.
- Keep dependency counts low.

## License
By contributing, you agree that your contributions will be licensed under its MIT License.
