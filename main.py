# -*- coding: utf-8 -*-

import datetime
import os
import re
import zipfile
import hashlib


def proceed_unzip(path):
    try:
        with zipfile.ZipFile(path + '.zip', 'r') as zip_ref:
            zip_ref.extractall(path + '_unlocked')
    except Exception as ex:
        raise RuntimeError(ex)


def proceed_zip(path):
    try:
        owd = os.getcwd()
        os.chdir(path + '_unlocked')

        zip_file = zipfile.ZipFile('../output.xlsx', "w")
        for (root, dirs, files) in os.walk('./'):
            for file in files:
                zip_file.write(os.path.join(root, file), compress_type=zipfile.ZIP_DEFLATED)

        zip_file.close()
        os.chdir(owd)
    except Exception as ex:
        raise RuntimeError(ex)


def proceed_unlock(path, filename):
    _ext = filename.split('.')[-1]
    _non_ext = ''.join(filename.split('.')[:-1])
    _hashpath = path + '/' + 'hashcode'
    _fullpath = _hashpath + '/' + _non_ext

    if _ext in ['xlsm', 'xlsx', 'xls']:
        os.rename('.'.join([_fullpath, _ext]), '.'.join([_fullpath, 'zip']))
    elif _ext in ['xlsb']:
        # TODO pyxlsb
        pass
    else:
        raise TypeError('File extension is not one of xlsx, xls, xlsm, xlsb.')

    proceed_unzip(_fullpath)

    _taskpath = _fullpath + '_unlocked/xl/worksheets'
    tasks = [f for f in os.listdir(_taskpath) if f[-4:] == '.xml']

    for task in tasks:
        with open(_taskpath + '/' + task, 'r', encoding='utf-8') as fr:
            text = fr.read()
            text_unlocked = re.sub('(<sheetProtection.+?>)', '', text)
            fr.close()

        with open(_taskpath + '/' + task, 'w', encoding='utf-8') as fw:
            fw.write(text_unlocked)
            fw.close()

    proceed_zip(_fullpath)


if __name__ == '__main__':
    path = './workplace'
    filename = 'test.xlsx'

    proceed_unlock(path, filename)
