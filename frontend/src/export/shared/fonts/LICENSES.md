# Bundled fonts

Most of these faces are **metric-compatible substitutes** so that server-side /
worker export produces the same line breaks as the editor and embeds glyphs into
PDFs. The OFL/Apache faces below are freely redistributable.

| Family         | Substitutes for      | Author(s)                                   | License            |
|----------------|----------------------|---------------------------------------------|--------------------|
| Carlito        | Calibri              | Łukasz Dziedzic (Google)                    | SIL OFL 1.1        |
| Caladea        | Cambria              | Carolina Giovagnoli, Andrés Torresi (HT)    | SIL OFL 1.1        |
| Gelasio        | Georgia              | Eben Sorkin (Sorkin Type)                   | SIL OFL 1.1        |
| Arimo          | Arial / Helvetica    | Steve Matteson (Ascender / Google Croscore) | Apache License 2.0 |
| Cousine        | Courier New          | Steve Matteson (Ascender / Google Croscore) | Apache License 2.0 |
| StixTwoMath    | (math typesetting)   | The STIX Fonts project / Tiro Typeworks     | SIL OFL 1.1        |

`StixTwoMath-Regular.ttf` is **STIX Two Math** — the math font equations are
typeset and rendered with (real math glyphs + the Mathematical Alphanumeric block
for true italic/bold/blackboard letters). Single Regular face; not a metric clone.
Source: https://github.com/stipub/stixfonts (OFL 1.1).

Full license texts: SIL OFL 1.1 — https://openfontlicense.org ;
Apache 2.0 — https://www.apache.org/licenses/LICENSE-2.0

## ⚠ Times New Roman — proprietary (bundled on request)

`TimesNewRoman-*.ttf` is the **genuine Microsoft Times New Roman**, bundled at the
project owner's explicit request (replacing the OFL Tinos clone). It is **not**
freely redistributable — it is licensed by Microsoft/Monotype and shipping it
imposes that license on this repository. Source:
https://github.com/misuchiru03/font-times-new-roman . Ensure you have the rights
to redistribute it before publishing; otherwise revert to Tinos (Apache-2.0).
