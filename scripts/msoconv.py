#!/usr/bin/env python3
# -*- tab-width: 4; indent-tabs-mode: nil; py-indent-offset: 4 -*-
# Version: MPL 1.1 / GPLv3+ / LGPLv3+
#
# The contents of this file are subject to the Mozilla Public License Version
# 1.1 (the "License"); you may not use this file except in compliance with
# the License or as specified alternatively below. You may obtain a copy of
# the License at https://www.mozilla.org/MPL/
#
# Software distributed under the License is distributed on an "AS IS" basis,
# WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
# for the specific language governing rights and limitations under the
# License.

import os
import parser
from subprocess import Popen, PIPE, TimeoutExpired, DEVNULL, run
import signal
import string
import random
import multiprocessing
import shutil

def kill_mso():
    p = Popen(['ps', '-A'], stdout=PIPE)
    out, err = p.communicate()
    for line in out.splitlines():
        if b'WINWORD.EXE' in line or b'POWERPNT.EXE' in line or b'wine' in line:
            pid = int(line.split(None, 1)[0])
            print("Killing process: " + str(pid))
            os.kill(pid, signal.SIGKILL)

def launch_OfficeConverter(fileName, arguments):
    tmpName = ''.join(random.choice(string.ascii_letters) for x in range(8)) + '.pdf'
    inputName = os.path.join(arguments.input, fileName)
    outputName = os.path.join(arguments.output, fileName + ".pdf")
    shutil.copyfile(inputName, os.path.join('/tmp', tmpName))
    os.chdir('/tmp/')
    try:
        run(['xvfb-run', '-a', 'wine', 'OfficeConvert', '--format=pdf', fileName, "--output=" + tmpName],
                stdout=DEVNULL, stderr=DEVNULL, timeout=100)
        shutil.move(os.path.join('/tmp', tmpName), outputName)
        print("Converted " + inputName + " to " + outputName)

    except TimeoutExpired as e:
        print("TIMEOUT: Converting " + inputName + " to " + outputName)
        os.remove(os.path.join('/tmp', tmpName))

if __name__ == "__main__":
    parser = parser.CommonParser()
    parser.add_arguments(['--input', '--output', '--extension', '--wineprefix'])

    arguments = parser.check_values()

    os.environ["WINEPREFIX"] = arguments.wineprefix

    listFiles = []

    for fileName in os.listdir(arguments.input):
        if fileName.endswith('.' + arguments.extension):
            outputName = os.path.join(arguments.output, fileName + ".pdf")
            if not os.path.exists(outputName):
                listFiles.append(fileName)

    if listFiles:
        cpuCount = multiprocessing.cpu_count()
        chunkSplit = cpuCount * 16

        chunks = [listFiles[x:x+chunkSplit] for x in range(0, len(listFiles), chunkSplit)]
        for chunk in chunks:
            kill_mso()

            pool = multiprocessing.Pool(cpuCount)
            for fileName in chunk:
                pool.apply_async(launch_OfficeConverter, args=(fileName, arguments))

            pool.close()
            pool.join()


# vim:set shiftwidth=4 softtabstop=4 expandtab:
