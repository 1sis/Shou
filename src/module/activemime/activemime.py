#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: Isis (@0x1sis)

# taken from https://github.com/ttpreport/activemaim

from .byte_iterator import ByteIterator
from types import SimpleNamespace


class ActiveMime:
    def __init__(self, raw_activemime):
        self.parse(raw_activemime)

    def parse(self, raw_activemime):
        iterator = ByteIterator(raw_activemime)
        self.cfb = SimpleNamespace()
        # taken from https://github.com/idiom/activemime-format
        self.cfb.header = iterator.next(0x0C)
        self.cfb.unknown_a = iterator.next(0x02)
        self.cfb.field_size = iterator.next(0x04)
        self.cfb.unknown_b = iterator.next(int.from_bytes(self.cfb.field_size, byteorder='little'))
        self.cfb.unknown_c = iterator.next(0x04)
        self.cfb.compressed_size = iterator.next(0x04)
        self.cfb.field_size_d = iterator.next(0x04)
        self.cfb.field_size_e = iterator.next(0x04)
        self.cfb.unknown_d = iterator.next(int.from_bytes(self.cfb.field_size_d, byteorder='little'))
        self.cfb.vba_tail_type = iterator.next(int.from_bytes(self.cfb.field_size_e, byteorder='little'))
        self.cfb.size = iterator.next(0x04)
        self.cfb.compressed_data = iterator.next(int.from_bytes(self.cfb.compressed_size, byteorder='little'))

        self.magic_tail = SimpleNamespace()  # all of this is a wild guess
        self.magic_tail.header = iterator.next(0x04)

        iterator.next(0x01)  # skipping delimiter
        self.magic_tail.unknown_a = iterator.next(0x01)

        # FF 01 01 00 00 00 56 99 00 00 00 00 FF FF
        # /// 56 - cmg
        # /// 99 - bEncrypt
        procedures_meta_cnt = int.from_bytes(iterator.next(0x01), byteorder='little')
        self.magic_tail.procedures_meta = list()
        for _ in range(procedures_meta_cnt):
            self.magic_tail.procedures_meta.append(
                iterator.next_until(b'\xFF\xFF')[:-2])  # macro metadata, sans delimiter

        self.magic_tail.unknown_b = iterator.next_until(b'\xFF\xFF')[:-2]  # always same, sans delimiter

        procedures_names_cnt = int.from_bytes(iterator.next(0x02), byteorder='little')
        self.magic_tail.procedures_names = list()
        for _ in range(procedures_names_cnt):
            procedure_name_unknown_1 = iterator.next(0x02)
            procedure_name_size = int.from_bytes(iterator.next(0x02), byteorder='little')
            procedure_name_full = iterator.next(procedure_name_size * 2).decode('UTF-16LE')  # x2 because utf-16
            procedure_name_project, procedure_name_module, procedure_name_procedure = procedure_name_full.split('.')
            self.magic_tail.procedures_names.append({
                'unknown_1': procedure_name_unknown_1,
                'project': procedure_name_project,
                'module': procedure_name_module,
                'procedure': procedure_name_procedure
            })

        self.magic_tail.unknown_c = iterator.next(0x03)

        procedures_links_cnt = int.from_bytes(iterator.next(0x02), byteorder='little')
        self.magic_tail.procedures_links = list()
        for _ in range(procedures_links_cnt):
            procedure_link_unknown_1 = iterator.next(0x02)
            procedure_link_size = int.from_bytes(iterator.next(0x02), byteorder='little')
            procedure_link_full = iterator.next(procedure_link_size * 2).decode('UTF-16LE')  # x2 because utf-16
            procedure_link_project, procedure_link_module, procedure_link_procedure = procedure_link_full.split('.')
            self.magic_tail.procedures_links.append({
                'unknown_1': procedure_link_unknown_1,
                'project': procedure_link_project,
                'module': procedure_link_module,
                'procedure': procedure_link_procedure,
                'unknown_2': iterator.next(0x02)
            })

        self.magic_tail.footer = iterator.next(iterator.length())

    def render(self):
        data = b''

        data += self.cfb.header
        data += self.cfb.unknown_a
        data += self.cfb.field_size
        data += self.cfb.unknown_b
        data += self.cfb.unknown_c
        data += self.cfb.compressed_size
        data += self.cfb.field_size_d
        data += self.cfb.field_size_e
        data += self.cfb.unknown_d
        data += self.cfb.vba_tail_type
        data += self.cfb.size
        data += self.cfb.compressed_data

        data += self.magic_tail.header
        data += b'\xFF'  # delimiter
        data += self.magic_tail.unknown_a
        data += len(self.magic_tail.procedures_meta).to_bytes(1, byteorder='little')

        for procedure_meta in self.magic_tail.procedures_meta:
            data += procedure_meta
            data += b'\xFF\xFF'  # delimiter

        data += self.magic_tail.unknown_b
        data += b'\xFF\xFF'  # delimiter
        data += len(self.magic_tail.procedures_names).to_bytes(2, byteorder='little')

        for procedure_name in self.magic_tail.procedures_names:
            data += procedure_name['unknown_1']
            procedure_name_full = f"{procedure_name['project']}.{procedure_name['module']}.{procedure_name['procedure']}"
            data += len(procedure_name_full).to_bytes(2, byteorder='little')
            data += procedure_name_full.encode('UTF-16LE')

        data += self.magic_tail.unknown_c
        data += len(self.magic_tail.procedures_links).to_bytes(2, byteorder='little')

        for procedure_link in self.magic_tail.procedures_links:
            data += procedure_link['unknown_1']
            procedure_link_full = f"{procedure_link['project']}.{procedure_link['module']}.{procedure_link['procedure']}"
            data += len(procedure_link_full).to_bytes(2, byteorder='little')
            data += procedure_link_full.upper().encode('UTF-16LE')
            data += procedure_link['unknown_2']

        data += self.magic_tail.footer

        return data