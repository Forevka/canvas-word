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
| NotoSansSC     | (CJK fallback)       | Google                                      | SIL OFL 1.1        |
| NotoSansArabic | (Arabic fallback)    | Google                                      | SIL OFL 1.1        |
| NotoSansHebrew | (Hebrew fallback)    | Google                                      | SIL OFL 1.1        |

`StixTwoMath-Regular.ttf` is **STIX Two Math** — the math font equations are
typeset and rendered with (real math glyphs + the Mathematical Alphanumeric block
for true italic/bold/blackboard letters). Single Regular face; not a metric clone.
Source: https://github.com/stipub/stixfonts (OFL 1.1).

`NotoSansSC-Regular.ttf` is a **subset of Noto Sans SC** (Simplified Chinese) — the
default CJK fallback so Chinese text isn't `.notdef`/tofu in PDF export when no
embedder font is configured. It is subset to the GB2312 Level-1 common characters
(~3,755 hanzi) plus ASCII and CJK punctuation to keep the bundle small (~1.4 MB);
rarer/less-common hanzi outside that set are not covered. Single Regular face (all
styles map to it; no bold/italic faces). Source:
https://github.com/google/fonts/tree/main/ofl/notosanssc (OFL 1.1).

`NotoSansArabic-Regular.ttf` is **Noto Sans Arabic** v2.009 — the default Arabic
fallback so Arabic text isn't `.notdef`/tofu in PDF export when no embedder font is
configured. Full Arabic Unicode block coverage (Basic Arabic, Arabic Supplement,
Extended-A, Presentation Forms-A/B); contextual joining forms via OpenType GSUB.
~240 KB. Single Regular face (all styles map to it; no bold/italic faces). Source:
https://github.com/notofonts/noto-fonts (OFL 1.1).

`NotoSansHebrew-Regular.ttf` is **Noto Sans Hebrew** — the default Hebrew fallback so
Hebrew text isn't `.notdef`/tofu ("x") in PDF export when no embedder font is
configured. Full Hebrew Unicode block coverage (Hebrew block plus the Hebrew letters
and ligatures in the Alphabetic Presentation Forms range). ~27 KB. Single Regular
face (all styles map to it; no bold/italic faces). Source:
https://github.com/notofonts/noto-fonts (OFL 1.1).

Full license texts: SIL OFL 1.1 — https://openfontlicense.org ;
Apache 2.0 — https://www.apache.org/licenses/LICENSE-2.0

## ⚠ Times New Roman — proprietary (bundled on request)

`TimesNewRoman-*.ttf` is the **genuine Microsoft Times New Roman**, bundled at the
project owner's explicit request (replacing the OFL Tinos clone). It is **not**
freely redistributable — it is licensed by Microsoft/Monotype and shipping it
imposes that license on this repository. Source:
https://github.com/misuchiru03/font-times-new-roman . Ensure you have the rights
to redistribute it before publishing; otherwise revert to Tinos (Apache-2.0).
