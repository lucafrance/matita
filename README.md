# python-vba

Parser of the [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview).
Creates a json of the structure of the [Office VBA Object library](https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference/reference-object-library-reference-for-office) as described in the Office VBA Reference.

## Installation

Install the package with:

```powershell
python -m pip install .
```

## Usage

The `matita` package contains the parser and related modules. To run the parser, use the included script:

```powershell
python -m scripts
```

The output will be generated in `data/office-vba-api.json`.

## Limitations

The following objects are unsupported, because their name conflicts with reserved keywords in Python.
- [Break object (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.break)
- [Global object (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.global)

The following objects are unsupported, because non-scalar arguments are not implemented.
- [Report.Circle method (Access)](https://learn.microsoft.com/en-gb/office/vba/api/access.report.circle)
- [Report.Line method (Access)](https://learn.microsoft.com/en-gb/office/vba/api/access.report.line)

## Attribution

This project is based on the [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview) by Microsoft Corporation, [licensed](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/LICENSE) under [Creative Commons Attribution 4.0 International](https://creativecommons.org/licenses/by/4.0/).
