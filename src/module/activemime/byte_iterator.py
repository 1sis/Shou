#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Author: Isis (@0x1sis)

# Taken from https://github.com/ttpreport/activemaim

class ByteIterator:
    def __init__(self, data):
        self.data = data
        self.cursor = 0x00

    def next(self, size=0x01):
        result = self.data[self.cursor:(self.cursor + size)]
        self.cursor += size
        return result

    def next_peek(self, size=0x01):
        return self.data[self.cursor:(self.cursor + size)]

    def next_until(self, stop_bytes):
        result = b''
        while result[-len(stop_bytes):] != stop_bytes:
            if self.next_peek() == b'':
                raise Exception(f'Stop bytes {stop_bytes} were not encountered')
            else:
                result += self.next()
        return result

    def length(self):
        return len(self.data)