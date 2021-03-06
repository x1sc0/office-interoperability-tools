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

import sys
import os
import parser
from subprocess import Popen, PIPE, DEVNULL
import signal
import string
import random
import shutil
import time

def display_in_used():
    p = Popen(["xdpyinfo", "-display", ":0"], stdout=PIPE, stderr=PIPE)
    p.communicate()
    return p.returncode == 0

def kill_mso(wineprefix):
    p = Popen(['ps', '-A'], stdout=PIPE)
    out, err = p.communicate()
    for line in out.splitlines():
        if b'WINWORD.EXE' in line or b'POWERPNT.EXE' in line or \
                b'wine' in line or b'explorer.exe' in line or \
                b'services.exe' in line or b'rpcss.exe' in line or \
                b'OSPPSVC.EXE' in line or b'plugplay.exe' in line:
            pid = int(line.split(None, 1)[0])
            try:
                os.kill(pid, signal.SIGKILL)
            except ProcessLookupError:  # process already dead
                pass

    userBackPath = os.path.join(wineprefix, 'drive_c', 'users.back')
    userPath = os.path.join(wineprefix, 'drive_c', 'users')
    if not os.path.exists(userBackPath):
        print("ERROR: Create a backup of " + userPath + " in " + userBackPath)
        print("Make sure MS Office is already working under Wine")
        print("it will replace " + userPath + " on every execution as a clean profile")
        print()
        sys.exit(1)
    shutil.rmtree(userPath)
    shutil.copytree(userBackPath, userPath)

def launch_OfficeConverter(fileName, arguments, count, total_count):
    tmpName = ''.join(random.choice(string.ascii_letters) for x in range(8)) + '.pdf'
    inputName = os.path.join(arguments.input, fileName)
    outputName = os.path.join(arguments.output, fileName + ".pdf")
    shutil.copyfile(inputName, os.path.join('/tmp', fileName))
    os.chdir('/tmp/')

    startTime = time.time()
    # Using timeout with popen fills up the memory and eventually the kernel throws a OSError: [Errno 12] Cannot allocate memory
    # pass the timeout parameter instead. 60 should be more than enough. Most files are converted within 10 seconds
    timeout = 60
    p = Popen(['timeout', str(timeout), 'wine', 'OfficeConvert', '--format=pdf', fileName, "--output=" + tmpName],
        stdout=DEVNULL, stderr=DEVNULL)
    p.communicate()
    endTime = time.time()
    diffTime = int(endTime - startTime)

    if os.path.exists(tmpName):
        shutil.move('/tmp/' + tmpName, outputName)
        print(str(count) + "/" + str(total_count) + " - SUCCESS: Converting " + inputName + " to " + outputName + ' after ' + str(diffTime) + ' seconds')
    else:
        if diffTime >= timeout - 1:
            print(str(count) + "/" + str(total_count) + " - TIMEOUT: Converting " + inputName + " to " + outputName + ' after ' + str(diffTime) + ' seconds')
        else:
            # Some files are corrupted
            print(str(count) + "/" + str(total_count) + " - ERROR: Converting " + inputName + " to " + outputName + ' after ' + str(diffTime) + ' seconds')

    os.remove(os.path.join('/tmp', fileName))

if __name__ == "__main__":
    parser = parser.CommonParser()
    parser.add_arguments(['--input', '--output', '--extension', '--wineprefix'])

    arguments = parser.check_values()

    if not display_in_used():
        print("Display :0 is no in used. Emulating it ")
        Popen(["Xvfb", ":0", "-screen", "0", "1024x768x16"])

    if "DISPLAY" not in os.environ:
        os.environ["DISPLAY"] = ":0.0"
    os.environ["WINEPREFIX"] = arguments.wineprefix

    listFiles = []

    for fileName in os.listdir(arguments.input):
        if fileName.endswith('.' + arguments.extension):
            outputName = os.path.join(arguments.output, fileName + ".pdf")
            if not os.path.exists(outputName):
                listFiles.append(fileName)

    if listFiles:
        chunkSplit = 64
        count = 0
        total_count = len(listFiles)

        chunks = [listFiles[x:x+chunkSplit] for x in range(0, len(listFiles), chunkSplit)]
        for chunk in chunks:
            kill_mso(arguments.wineprefix)

            for fileName in chunk:
                count += 1
                launch_OfficeConverter(fileName, arguments, count, total_count)

# vim:set shiftwidth=4 softtabstop=4 expandtab:
