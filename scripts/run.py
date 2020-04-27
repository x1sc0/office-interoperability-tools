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

import parser
import os
from subprocess import Popen, PIPE
import signal
import config

def kill_soffice():
    p = Popen(['ps', '-A'], stdout=PIPE)
    out, err = p.communicate()
    for line in out.splitlines():
        if b'soffice' in line:
            pid = int(line.split(None, 1)[0])
            print("Killing process: " + str(pid))
            os.kill(pid, signal.SIGKILL)

if __name__ == "__main__":
    parser = parser.CommonParser()
    parser.add_arguments(['--soffice', '--type', '--reference', '--component', '--indir', '--outdir'])

    parser.add_optional_arguments(['--wineprefix'])
    arguments = parser.check_values()

    kill_soffice()

    process = Popen([arguments.soffice, "--version"], stdout=PIPE, stderr=PIPE)
    stdout = process.communicate()[0].decode("utf-8")
    version = stdout.split(" ")[2].strip()

    sofficePath = os.path.dirname(arguments.soffice)

    os.environ["PYTHONPATH"] = sofficePath
    os.environ["URE_BOOTSTRAP"] = "file://" + sofficePath + "/fundamentalrc"

    scriptsPath = os.path.dirname(os.path.abspath(__file__))
    outDir = os.path.join(os.path.sep, arguments.outdir, arguments.type, arguments.component, version)

    if not os.path.exists(outDir):
        os.makedirs(outDir)

    #Step 1: Convert files with LibreOffice
    process = Popen(['python3', scriptsPath + '/unoconv.py',
        '--soffice=' + arguments.soffice,
        '--type=' + arguments.type,
        '--component=' + arguments.component,
        '--indir=' + arguments.indir,
        '--outdir=' + outDir])
    process.communicate()

    #Step 2: Convert files with MSO
    if arguments.type == 'ooxml':
        for extension in config.config['ooxml'][arguments.component]["import"]:
            process = Popen(['python3', scriptsPath + '/msoconv.py',
                '--extension=' + extension,
                '--wineprefix=' + arguments.wineprefix,
                '--indir=' + arguments.reference,
                '--outdir=' + outDir])
            process.communicate()

    #Step 3: Convert the roundtripped files from step 1 to PDF with MSO
    if arguments.type == 'ooxml':
        for extension in config.config['ooxml'][arguments.component]["export"]:
            process = Popen(['python3', scriptsPath + '/msoconv.py',
                '--extension=' + extension,
                '--wineprefix=' + arguments.wineprefix,
                '--indir=' + arguments.indir,
                '--outdir=' + outDir])
            process.communicate()

    #Step 4: Get rid of these files once converted to PDF
    for fileName in os.listdir(arguments.outdir):
        for extension in config.config[arguments.type][arguments.component]["export"]:
            ext = os.path.splitext(fileName)[1][1:]
            if ext == extension and ext != 'pdf':
                os.remove(arguments.indir + fileName)
