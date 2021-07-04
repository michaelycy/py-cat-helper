# -*- coding: UTF-8 -*-
# ! /usr/bin/python3

from base64 import b64decode
from os import path


def gen_tmp(base64_code, filepath):
    tmp = open(filepath, 'wb+')
    tmp.write(b64decode(base64_code))
    tmp.close()

    return path.abspath(filepath)
