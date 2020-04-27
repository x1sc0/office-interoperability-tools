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

def kill_mso():
    p = Popen(['ps', '-A'], stdout=PIPE)
    out, err = p.communicate()
    for line in out.splitlines():
        if b'WINWORD.EXE' in line or b'POWERPNT.EXE' in line or b'wine' in line:
            pid = int(line.split(None, 1)[0])
            print("Killing process: " + str(pid))
            os.kill(pid, signal.SIGKILL)

def X_is_running():
    p = Popen(["xset", "-q"], stdout=PIPE, stderr=PIPE)
    p.communicate()
    return p.returncode == 0

def launch_OfficeConverter(fileName, arguments):
    tmpName = ''.join(random.choice(string.ascii_letters) for x in range(8)) + '.pdf'
    baseName = os.path.splitext(fileName)[0]
    outputName = arguments.outdir + baseName + ".pdf"
    os.chdir(arguments.indir)
    try:
        run(['wine', 'OfficeConvert', '--format=pdf', fileName, "--output=" + tmpName],
                stdout=DEVNULL, stderr=DEVNULL, timeout=60)
        os.rename(tmpName, outputName)
        print("Converted " + arguments.indir + fileName + " to " + outputName)

    except TimeoutExpired as e:
        print("TIMEOUT: Converting " + arguments.indir + fileName + " to " + outputName)

if __name__ == "__main__":
    parser = parser.CommonParser()
    parser.add_arguments(['--indir', '--outdir', '--extension', '--wineprefix'])

    arguments = parser.check_values()

    kill_mso()

    #Check if X is running, otherwise, simulate one
    if not X_is_running():
        p = Popen(["Xvfb", ":0", "-screen", "0", "1024x768x16", "&"], stdout=PIPE, stderr=PIPE)
        p.communicate()
        os.environ["DISPLAY"] = ":0.0"

    os.environ["WINEPREFIX"] = arguments.wineprefix

    cpuCount = multiprocessing.cpu_count()
    chunkSplit = cpuCount * 16
    listFiles = []

    for fileName in os.listdir(arguments.indir):
        if fileName.endswith('.' + arguments.extension):
            baseName = os.path.splitext(fileName)[0]
            outputName = arguments.outdir + baseName + ".pdf"
            if not os.path.exists(outputName):
                listFiles.append(fileName)

    chunks = [listFiles[x:x+chunkSplit] for x in range(0, len(listFiles), chunkSplit)]
    for chunk in chunks:
        pool = multiprocessing.Pool(cpuCount)
        for fileName in chunk:
            pool.apply_async(launch_OfficeConverter, args=(fileName, arguments))

        pool.close()
        pool.join()

        kill_mso()


# vim:set shiftwidth=4 softtabstop=4 expandtab:
