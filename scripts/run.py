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
import shutil

def kill_soffice():
    p = Popen(['ps', '-A'], stdout=PIPE)
    out, err = p.communicate()
    for line in out.splitlines():
        if b'soffice' in line:
            pid = int(line.split(None, 1)[0])
            print("Killing process: " + str(pid))
            os.kill(pid, signal.SIGKILL)

def convert_input_with_libreOffice(scriptsPath, inputPath, outputPath, typeName, componentName, testName):
    process = Popen(['python3', scriptsPath + '/unoconv.py',
        '--soffice=' + arguments.soffice,
        '--type=' + typeName,
        '--component=' + componentName,
        '--indir=' + inputPath,
        '--outdir=' + outputPath])
    process.communicate()

    if testName:
        count = 1
        assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, 'import', 'pdf']))))
        for extension in config.config[typeName][componentName]["export"]:
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension]))))
            count += 1
        for extension in config.config[typeName][componentName]["roundtrip"]:
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf']))))
            count += 1
        assert(len(os.listdir(outputPath)) == count)

def convert_reference_with_mso(scriptsPath, inputPath, referencePath, componentName, testName):
    for extension in config.config['ooxml'][componentName]["import"]:
        process = Popen(['python3', scriptsPath + '/msoconv.py',
            '--extension=' + extension,
            '--wineprefix=' + arguments.wineprefix,
            '--indir=' + inputPath,
            '--outdir=' + referencePath])
        process.communicate()

    if testName:
        assert(len(os.listdir(referencePath)) == 1)
        assert(os.path.exists(os.path.join(referencePath, '.'.join([testName, 'pdf']))))

def convert_output_to_pdf_with_mso(scriptsPath, outputPath, componentName, testName):
    for extension in config.config['ooxml'][arguments.component]["export"]:
        if extension not in config.config["ooxml"][arguments.component]["roundtrip"]:
            process = Popen(['python3', scriptsPath + '/msoconv.py',
                '--extension=' + extension,
                '--wineprefix=' + arguments.wineprefix,
                '--indir=' + outputPath,
                '--outdir=' + outputPath])
            process.communicate()

    if testName:
        for extension in config.config["ooxml"][componentName]["export"]:
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf']))))

def remove_non_pdf_files(outputPath, typeName, componentName, testName):
    for fileName in os.listdir(outputPath):
        for extension in config.config[typeName][componentName]["export"]:
            ext = os.path.splitext(fileName)[1][1:]
            if ext == extension:
                print("Removing " + os.path.join(outputPath, fileName))
                os.remove(os.path.join(outputPath, fileName))

    if testName:
        for extension in config.config[typeName][componentName]["export"]:
            assert(not os.path.exists(os.path.join(outputPath, '.'.join([testName, extension]))))
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf']))))

def replace_non_converted_files(scriptsPath, inputPath, outputPath, typeName, componentName, testName):

    if testName:
        #Remove directory
        shutil.rmtree(outputPath)
        os.makedirs(outputPath)

    failedPdfPath = os.path.join(scriptsPath, 'failed.pdf')
    for i in os.listdir(inputPath):
        if not i.startswith(".~lock."):
            importNamePath = os.path.join(outputPath, i + ".import.pdf")
            if not os.path.exists(importNamePath):
                print(importNamePath + " doesn't exists. Using failed.pdf")
                shutil.copyfile(failedPdfPath, importNamePath)

            for extension in config.config[arguments.type][arguments.component]["export"]:
                exportNamePath = os.path.join(outputPath, ".".join([i, extension, "pdf"]) )

                if not os.path.exists(exportNamePath):
                    print(exportNamePath + " doesn't exists. Using failed.pdf")
                    shutil.copyfile(failedPdfPath, exportNamePath)

    if testName:
        for extension in config.config[typeName][componentName]["export"]:
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf']))))

def compare_pdfs(scriptsPath, outputPath, referencePath, typeName, componentName, testName):
    process = Popen(['python3', scriptsPath + '/docompare.py',
        '--indir=' + outputPath,
        '--reference=' + referencePath])
    process.communicate()

    if testName:
        for extension in config.config[typeName][componentName]["export"]:
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf-pair-l', 'pdf']))))
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf-pair-p', 'pdf']))))
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf-pair-s', 'pdf']))))
            assert(os.path.exists(os.path.join(outputPath, '.'.join([testName, extension, 'pdf-pair-z', 'pdf']))))

def execute(arguments, isTest):
    scriptsPath = os.path.dirname(os.path.abspath(__file__))
    typeName = arguments.type
    componentName = arguments.component

    if isTest:
        inputPath = os.path.join( scriptsPath, "tests")
        outputPath = "/tmp/test-output/"
        referencePath = "/tmp/test-reference/"

        if os.path.exists(outputPath):
            shutil.rmtree(outputPath)
        if os.path.exists(referencePath):
            shutil.rmtree(referencePath)
        os.makedirs(referencePath)

        for fileName in os.listdir(inputPath):
            for extension in config.config[typeName][componentName]["import"]:
                isTest = fileName

    else:
        inputPath = arguments.indir
        outputPath = os.path.join(os.path.sep, arguments.outdir, arguments.type, arguments.component, version)
        referencePath = arguments.reference

    if not os.path.exists(outputPath):
        os.makedirs(outputPath)

    convert_input_with_libreOffice(scriptsPath, inputPath, outputPath, typeName, componentName, isTest)

    if typeName == 'ooxml':
        convert_reference_with_mso(scriptsPath, inputPath, referencePath, componentName, isTest)
        convert_output_to_pdf_with_mso(scriptsPath, outputPath, componentName, isTest)

    remove_non_pdf_files(outputPath, typeName, componentName, isTest)
    replace_non_converted_files(scriptsPath, inputPath, outputPath, typeName, componentName, isTest)
    compare_pdfs(scriptsPath, outputPath, referencePath, typeName, componentName, isTest)

    if isTest:
        shutil.rmtree(outputPath)
        shutil.rmtree(referencePath)

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

    execute(arguments, True)
    print("Tests Passed!!")
    execute(arguments, False)

