# **Shou** – **Automated fake PDF Creation with embedded VBA macros**

**Shou** is a tool designed for automate MalDoc generation.

## Explanation:

We use an [MHT format](https://en.wikipedia.org/wiki/MHTML)

The technique was seen by the Japanese cert in 2023 (MaldocInPDF), it was simply automated (creation of the file, embedding of the macro...).

If the MHT file is renamed to .doc, it executes the macros, so it's a quickwin :)

Thanks to [@ttpreport](https://x.com/ttpreport) for the help, and permission to use his research :)


## To Do:

- VBA Obfuscation module
- More customisation VBA (Wscript.Shell.Exec / Run..)

## Quick Start Guide

1. **Installation:**
```shell
git clone https://github.com/1sis/Shou
cd Shou
pip install -r requirements.txt
```
Modify your registry key with "EnableVBOM.reg"


## Utilisation:

```python
python3 shou.py -f [SCRIPT.vba] -o [OUTPUT.doc]
```


## References:

- [OLETools – Detect VBA in Files](https://github.com/decalage2/oletools/)
- [Trustwave - Stealthy VBA Macro in PDF](https://www.trustwave.com/en-us/resources/blogs/spiderlabs-blog/stealthy-vba-macro-embedded-in-pdf-like-header-helps-evade-detection/)
- [JPCert - MaldocInPDF](https://blogs.jpcert.or.jp/en/2023/08/maldocinpdf.html)
- [Cyren - New tricks for Macro malware](https://web.archive.org/web/20240514103736/https://www.cyren.com/blog/articles/new-tricks-of-macro-malware)
- [TTP Report - Release activemaim](https://ttp.report/evasion/2023/11/02/releasing-activemaim-evade-macros-detection.html)