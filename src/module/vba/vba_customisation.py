#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: Isis (@0x1sis)

import re
import winreg
import tempfile

class Customisation:
    def __init__(self, vba_file):
        self.vba_file = vba_file

    @staticmethod
    def get_clsid_from_progid(prog_id):
        try:
            key_path = fr"SOFTWARE\Classes\{prog_id}\CLSID"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                clsid, _ = winreg.QueryValueEx(key, "")
                return clsid
        except FileNotFoundError:
            return None

    def modify_vba_file(self):
        with open(self.vba_file, "r", encoding="utf-8") as file:
            content = file.read()

            matches = re.findall(r'CreateObject\(["\'](?:new:)?([A-Za-z0-9.\-]+)["\']\)', content, re.IGNORECASE)

            for prog_id in matches:
                clsid = self.get_clsid_from_progid(prog_id.lower())
                if clsid:
                    content = content.replace(f'CreateObject("new:{prog_id}")', f'CreateObject("new:{clsid}")')

        temp_vba = tempfile.NamedTemporaryFile(delete=False, suffix=".vba", mode="w", encoding="utf-8")
        temp_vba.write(content)
        temp_vba.close()

        return temp_vba.name
