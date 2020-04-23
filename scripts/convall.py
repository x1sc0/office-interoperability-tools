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
#
# Major Contributor(s):
# Copyright (C) 2012 Red Hat, Inc., Michael Stahl <mstahl@redhat.com>
#  (initial developer)
#
# All Rights Reserved.
#
# For minor contributions see the git repository.
#
# Alternatively, the contents of this file may be used under the terms of
# either the GNU General Public License Version 3 or later (the "GPLv3+"), or
# the GNU Lesser General Public License Version 3 or later (the "LGPLv3+"),
# in which case the provisions of the GPLv3+ or the LGPLv3+ are applicable
# instead of those above.

import argparse
import os
import glob
import subprocess
import sys
import time
import uuid
import datetime
import tempfile
import config

import signal
import threading
try:
    from urllib.parse import quote
except ImportError:
    from urllib import quote

try:
    import pyuno
    import uno
    import unohelper
except ImportError:
    print("pyuno not found: try to set PYTHONPATH and URE_BOOTSTRAP variables")
    print("PYTHONPATH=/installation/opt/program")
    print("URE_BOOTSTRAP=file:///installation/opt/program/fundamentalrc")
    raise

try:
    from com.sun.star.document import XDocumentEventListener
except ImportError:
    print("UNO API class not found: try to set URE_BOOTSTRAP variable")
    print("URE_BOOTSTRAP=file:///installation/opt/program/fundamentalrc")
    raise

### utilities ###

def partition(list, pred):
    left = []
    right = []
    for e in list:
        if pred(e):
            left.append(e)
        else:
            right.append(e)
    return (left, right)

def getFiles(filesPath, extensions):
    auxNames = []
    for fileName in os.listdir(filesPath):
        for ext in extensions:
            if fileName.endswith(ext):
                fileNamePath = filesPath + fileName
                if os.path.isfile(fileNamePath):
                    auxNames.append(fileNamePath)

                    #Remove previous lock files
                    lockFilePath = filesPath + '.~lock.' + fileName + '#'
                    if os.path.isfile(lockFilePath):
                        os.remove(lockFilePath)
    return auxNames

### UNO utilities ###

class OfficeConnection:
    def __init__(self, args):
        self.args = args
        self.soffice = None
        self.socket = None
        self.xContext = None
        self.pro = None
    def setUp(self):
        (method, sep, rest) = self.args.soffice.partition(":")
        if sep != ":":
            raise Exception("soffice parameter does not specify method")
        if method == "path":
                socket = "pipe,name=pytest" + str(uuid.uuid1())
                self.soffice = self.bootstrap(rest, socket)
        elif method == "connect":
                socket = rest
        else:
            raise Exception("unsupported connection method: " + method)
        self.xContext = self.connect(socket)

    def bootstrap(self, soffice, socket):
        #Create temp directory for the user profile
        with tempfile.TemporaryDirectory() as tmpdirname:
            userPath = os.path.join(tmpdirname, 'libreoffice/4')
            os.makedirs(userPath)

            argv = [ soffice, "--accept=" + socket + ";urp",
                    "-env:UserInstallation=file://" + userPath,
                    "--quickstart=no", "--nofirststartwizard",
                    "--norestore", "--nologo", "--headless" ]
            self.pro = subprocess.Popen(argv)
            print(self.pro.pid)

    def connect(self, socket):
        xLocalContext = uno.getComponentContext()
        xUnoResolver = xLocalContext.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", xLocalContext)
        url = "uno:" + socket + ";urp;StarOffice.ComponentContext"
        print("OfficeConnection: connecting to: " + url)
        while True:
            try:
                xContext = xUnoResolver.resolve(url)
                return xContext
#            except com.sun.star.connection.NoConnectException
            except pyuno.getClass("com.sun.star.connection.NoConnectException"):
                print("NoConnectException: sleeping...")
                time.sleep(1)

    def tearDown(self):
        if self.soffice:
            if self.xContext:
                try:
                    print("tearDown: calling terminate()...")
                    xMgr = self.xContext.ServiceManager
                    xDesktop = xMgr.createInstanceWithContext(
                            "com.sun.star.frame.Desktop", self.xContext)
                    xDesktop.terminate()
                    print("...done")
#                except com.sun.star.lang.DisposedException:
                except pyuno.getClass("com.sun.star.beans.UnknownPropertyException"):
                    print("caught UnknownPropertyException while TearDown")
                    pass # ignore, also means disposed
                except pyuno.getClass("com.sun.star.lang.DisposedException"):
                    print("caught DisposedException while TearDown")
                    pass # ignore
            else:
                self.soffice.terminate()
            ret = self.soffice.wait()
            self.xContext = None
            self.socket = None
            self.soffice = None
            if ret != 0:
                raise Exception("Exit status indicates failure: " + str(ret))
#            return ret
    def kill(self):
        command = "kill " + str(self.pro.pid)
        killFile = open("killFile.log", "a")
        killFile.write(command + "\n")
        killFile.close()
        print("kill")
        print(command)
        os.system(command)

class PersistentConnection:
    def __init__(self, args):
        self.args = args
        self.connection = None
    def getContext(self):
        return self.connection.xContext
    def setUp(self):
        assert(not self.connection)
        conn = OfficeConnection(self.args)
        conn.setUp()
        self.connection = conn
    def preTest(self):
        assert(self.connection)
    def postTest(self):
        assert(self.connection)
    def tearDown(self):
        if self.connection:
            try:
                self.connection.tearDown()
            finally:
                self.connection = None
    def kill(self):
        if self.connection:
            self.connection.kill()

def simpleInvoke(connection, test):
    try:
        connection.preTest()
        test.run(connection.getContext(), connection)
    finally:
        connection.postTest()

def retryInvoke(connection, test):
    tries = 5
    while tries > 0:
        try:
            tries -= 1
            try:
                connection.preTest()
                test.run(connection.getContext(), connection)
                return
            finally:
                connection.postTest()
        except KeyboardInterrupt:
            raise # Ctrl+C should work
        except:
            print("retryInvoke: caught exception")
    raise Exception("FAILED retryInvoke")

def runConnectionTests(connection, invoker, tests):
    try:
        connection.setUp()
        for test in tests:
            invoker(connection, test)
    finally:
        pass
        #connection.tearDown()

class EventListener(XDocumentEventListener,unohelper.Base):
    def __init__(self):
        self.layoutFinished = False
    def documentEventOccured(self, event):
#        print(str(event.EventName))
        if event.EventName == "OnLayoutFinished":
            self.layoutFinished = True
    def disposing(event):
        pass

def mkPropertyValue(name, value):
    return uno.createUnoStruct("com.sun.star.beans.PropertyValue",
            name, 0, value, 0)

### tests ###

def loadFromURL(xContext, url, t, component):
    xDesktop = xContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", xContext)
    props = [("Hidden", True), ("ReadOnly", True)] # FilterName?
    loadProps = tuple([mkPropertyValue(name, value) for (name, value) in props])

    xListener = EventListener()
    xGEB = xContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.GlobalEventBroadcaster", xContext)
    xGEB.addDocumentEventListener(xListener)

    try:
        xDoc = None
        xDoc = xDesktop.loadComponentFromURL(url, "_blank", 0, loadProps)
        if component == "calc":
            try:
                if xDoc:
                    xDoc.calculateAll()
            except AttributeError:
                pass
            t.cancel()
            return xDoc
        elif component == "writer":
            time_ = 0
            t.cancel()
            while time_ < 30:
                if xListener.layoutFinished:
                    return xDoc
#                print("delaying...")
                time_ += 1
                time.sleep(1)
        else:
            t.cancel()
            return xDoc
        file = open("file.log", "a")
        file.write("layout did not finish\n")
        file.close()
        return xDoc
    except pyuno.getClass("com.sun.star.beans.UnknownPropertyException"):
        xListener = None
        raise # means crashed, handle it later
    except pyuno.getClass("com.sun.star.lang.DisposedException"):
        xListener = None
        raise # means crashed, handle it later
    except pyuno.getClass("com.sun.star.lang.IllegalArgumentException"):
        pass # means could not open the file, ignore it
    except:
        if xDoc:
            print("CLOSING")
            xDoc.close(True)
        raise
    finally:
        if xListener:
            xGEB.removeDocumentEventListener(xListener)

def handleCrash(file, disposed):
    print("File: " + file + " crashed")
    crashLog = open("crashlog.txt", "a")
    crashLog.write('Crash:' + file + ' ')
    if disposed == 1:
        crashLog.write('through disposed')
    crashLog.write('\n')
    crashLog.close()
#    crashed_files.append(file)
# add here the remaining handling code for crashed files

def alarm_handler(args):
    args.kill()

def writeExportCrash(fileName):
    exportCrash = open("exportCrash.txt", "a")
    exportCrash.write(fileName + '\n')
    exportCrash.close()

def exportDoc(xDoc, fileName, filterName, connection, timer):

    # note: avoid empty path segments in the url!
    fileURL = "file://" + fileName

    props = [ ("FilterName", filterName) ]

    saveProps = tuple([mkPropertyValue(name, value) for (name, value) in props])

    t = None
    try:
        args = [connection]
        t = threading.Timer(timer.getExportTime(), alarm_handler, args)
        t.start()
        xDoc.storeToURL(fileURL, saveProps)
    except pyuno.getClass("com.sun.star.beans.UnknownPropertyException"):
        if t.is_alive():
            writeExportCrash(filename)
        raise # means crashed, handle it later
    except pyuno.getClass("com.sun.star.lang.DisposedException"):
        if t.is_alive():
            writeExportCrash(filename)
        raise # means crashed, handle it later
    except pyuno.getClass("com.sun.star.lang.IllegalArgumentException"):
        pass # means could not open the file, ignore it
    except pyuno.getClass("com.sun.star.task.ErrorCodeIOException"):
        pass
    except:
        pass
    finally:
        if t.is_alive():
            t.cancel()

    print("xDoc.storeToURL " + fileURL + " " + filterName + "\n")

class ExportFileTest:
    def __init__(self, xDoc, filename, args, timer, isImport):
        self.xDoc = xDoc
        self.filename = filename
        self.args = args
        self.timer = timer
        self.isImport = isImport
        self.exportedFiles = []

    def run(self, connection):

        filterNames = { "ods": "calc8",
                "xls": "MS Excel 97",
                "xlsx": "Calc Office Open XML",
                "odt": "writer8",
                "doc": "MS Word 97",
                "docx": "Office Open XML Text",
                "rtf": "Rich Text Format",
                "odp": "impress8",
                "pptx": "Impress Office Open XML",
                "ppt": "MS PowerPoint 97",
                "pdf": self.args.component + "_pdf_Export"
                }

        fileBaseName = os.path.basename(self.filename)

        if self.isImport:
            fileNameNoExtension = os.path.splitext(fileBaseName)[0]
            filePath = self.args.outdir + '/' + fileNameNoExtension
            formats = config.config[self.args.type][self.args.component]["export"]
            for extension in formats:

                if extension == "pdf":
                    fileName = filePath + '.import.pdf'
                else:
                    fileName = filePath + "." + extension
                    self.exportedFiles.append(fileName)

                # Document has been roundtripped to PDF, no need to continue
                if os.path.exists(fileName) or os.path.exists(fileName + '.export.pdf'):
                    continue

                filterName = filterNames[extension]

                xExportedDoc = exportDoc(
                    self.xDoc, fileName, filterName, connection, self.timer)
                if xExportedDoc:
                    xExportedDoc.close(True)
        else:
            fileName = self.args.outdir + '/' + fileBaseName + '.export.pdf'

            if not os.path.exists(fileName):
                filterName = filterNames['pdf']

                xExportedDoc = exportDoc(
                    self.xDoc, fileName, filterName, connection, self.timer)
                if xExportedDoc:
                    xExportedDoc.close(True)

class LoadFileTest:
    def __init__(self, file, args, timer, isImport):
        self.file = file
        self.args = args
        self.timer = timer
        self.isImport = isImport
        self.exportedFiles = []

    def run(self, xContext, connection):
        print("Loading document: " + self.file)
        t = None
        args = None
        try:
            url = "file://" + quote(self.file)
            file = open("file.log", "a")
            file.write(url + "\n")
            file.close()
            xDoc = None
            args = [connection]
            t = threading.Timer(self.timer.getImportTime(), alarm_handler, args)
            t.start()
            xDoc = loadFromURL(xContext, url, t, self.args.component)
            print("doc loaded")
            t.cancel()
            if xDoc:
                exportTest = ExportFileTest(xDoc, self.file, self.args, self.timer, self.isImport)
                exportTest.run(connection)
                self.exportedFiles = exportTest.exportedFiles

        except pyuno.getClass("com.sun.star.beans.UnknownPropertyException"):
            print("caught UnknownPropertyException " + self.file)
            if not t.is_alive():
                print("TIMEOUT!")
            else:
                t.cancel()
                handleCrash(self.file, 0)
            connection.tearDown()
            connection.setUp()
            xDoc = None
        except pyuno.getClass("com.sun.star.lang.DisposedException"):
            print("caught DisposedException " + self.file)
            if not t.is_alive():
                print("TIMEOUT!")
            else:
                t.cancel()
                handleCrash(self.file, 1)
            connection.tearDown()
            connection.setUp()
            xDoc = None
        finally:
            if t.is_alive():
                t.cancel()
            try:
                if xDoc:
                    t = threading.Timer(10, alarm_handler, args)
                    t.start()
                    print("closing document")
                    xDoc.close(True)
                    t.cancel()
            except pyuno.getClass("com.sun.star.beans.UnknownPropertyException"):
                print("caught UnknownPropertyException while closing")
                connection.tearDown()
                connection.setUp()
            except pyuno.getClass("com.sun.star.lang.DisposedException"):
                print("caught DisposedException while closing")
                if t.is_alive():
                    t.cancel()
                else:
                    pass
                connection.tearDown()
                connection.setUp()
            print("...done with: " + self.file)

class NormalTimer:
    def __init__(self):
        pass

    def getImportTime(self):
        return 60

    def getExportTime(self):
        return 180

def runLoadFileTests(arguments, files, isImport):

    connection = PersistentConnection(arguments)
    exportedFiles = []
    try:
        files.sort()
        tests = []
        timer = NormalTimer()

        tests.extend( (LoadFileTest(file, arguments, timer, isImport) for file in files) )
        runConnectionTests(connection, simpleInvoke, tests)

        exportedFiles = [item for sublist in tests for item in sublist.exportedFiles]

    finally:
        connection.kill()

    return exportedFiles

components = ["writer", "calc", "impress"]
types = ["odf", "ooxml"]

class DefaultParser(argparse.ArgumentParser):

    def error(self, message):
        sys.stderr.write('error: %s\n' % message)
        self.print_help()
        sys.exit(2)

    def check_values(self):
        if arguments.dir:
            if not os.path.exists(arguments.dir):
                parser.error(arguments.dir + " doesn't exists")

        if arguments.outdir:
            if not os.path.exists(arguments.outdir):
                parser.error(arguments.outdir + " doesn't exists")

        if arguments.type:
            if arguments.type not in types:
                parser.error(arguments.type + " is an invalid type")

        if arguments.component:
            if arguments.component not in  components:
                parser.error(arguments.component + " is an invalid component")

if __name__ == "__main__":
    parser = DefaultParser()

    parser.add_argument('--soffice', required=True,
            help="""--soffice=method:location
            specify soffice instance to connect to
            supported methods: 'path', 'connect'""")
    parser.add_argument('--type', required=True,
            help="The mimetype to be tested. Options: 'ooxml' or 'odf'")
    parser.add_argument('--component', required=True,
            help="The component to be used. Options: 'calc', 'writer' or 'impress'")
    parser.add_argument('--dir', required=True,
            help="Path to the files directory")
    parser.add_argument('--outdir', required=True,
            help="Path to output directory")

    arguments = parser.parse_args()
    parser.check_values()

    extensions = config.config[arguments.type][arguments.component]["import"]
    importFiles = getFiles(arguments.dir, extensions)

    if importFiles:
        exportedFiles = runLoadFileTests(arguments, importFiles, True)
        print(exportedFiles)

    if exportedFiles:
        # Convert the roundtripped files to PDF
        runLoadFileTests(arguments, exportedFiles, False)

    # Remove the exported files to save some disk space. No longer needed
    for fileName in exportedFiles:
        if os.path.isfile(fileName):
            os.remove(fileName)


# vim:set shiftwidth=4 softtabstop=4 expandtab:
