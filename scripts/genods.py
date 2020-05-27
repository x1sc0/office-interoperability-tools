#! /usr/bin/python3
#
# This script generates a report from a list of results with numeric evaluations and from printed document files
#
# Copyright (C) 2013 Milos Sramek <milos.sramek@soit.sk>
# Licensed under the GNU LGPL v3 - http://www.gnu.org/licenses/gpl.html
# - or any later version.
#
import sys, os, getopt
import numpy as np
import re

from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style, TextProperties, ParagraphProperties, TableColumnProperties, TableCellProperties
from odf.text import P, A
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.office import Annotation

from docompare import create_overlayPDF

import requests
import ast

# we assume here this order in the testLabels list:[' PagePixelOvelayIndex[%]', ' FeatureDistanceError[mm]', ' HorizLinePositionError[mm]', ' TextHeightError[mm]', ' LineNumDifference']
testLabelsShort=['PPOI','FDE', 'HLPE', 'THE', 'LND']
testAnnotation = {
    'FDE': "Feature Distance Error / overlay of lines aligned verically and horizontally",
    'HLPE': "Horiz. Line Position Error / overlay of lines aligned only verically",
    'THE': "Text Height Error / page overlay with no alignment",
    'LND': "Line Number Difference / side by side view"
}

FDEMax = (0.01,0.5,1,2,4)        #0.5: difference of perfectly fitting AOO/LOO and MS document owing to different character rendering
HLPEMax = (0.01,5,10,15,20)        #
THEMax = (0.01,2, 4, 6,8)
LNDMax = (0.01,0.01,0.01,0.01,0.01)

def usage():
    print(sys.argv[0]+':')
    print("Usage: ", sys.argv[0], "[options]")
    print("\t--new Path to the new build")
    print("\t--old Path to the old build")
    print("\t--output outfile.odt ........ report {default: "+ofname+"}")
    print("\t--regression ....... Only display the regressions")
    print("\t--improvement ....... Only display the improvements")
    print("\t--odf ....... Check changes in ODF files")
    print("\t-h .................... this usage")

def parsecmd():
    global ofname, newPath, oldPath, checkRegressions, checkImprovements, checkOdf

    try:
        opts, args  = getopt.getopt(sys.argv[1:], "hvl:a:p:r:t:n:", ['help', 'new=', 'old=', 'output=', 'regression', 'improvement', 'odf'])
    except getopt.GetoptError as err:
        # print help information and exit:
        print(str(err)) # will print something like "option -a not recognized"
        usage()
        sys.exit(2)
    for o, a in opts:
        if o in ("-h", "--help"):
            usage()
            sys.exit()
        elif o == "--output":
            ofname = a
        elif o == "--new":
            newPath = a
        elif o == "--regression":
            checkRegressions = True
        elif o == "--improvement":
            checkImprovements = True
        elif o == "--odf":
            checkOdf = True
        elif o == "--old":
            oldPath = a
        else:
            assert False, "unhandled option"

    if os.path.isabs(oldPath) or os.path.isabs(newPath):
        print("ERROR: use relative pathes")
        sys.exit(1)

def check_tdf_bugs(listBugs):
    rest_url = 'https://bugs.documentfoundation.org/rest/bug?id='
    field = '&include_fields=id,status,summary'

    results = { 'import': [] , 'export': [] }
    tdfBugIds = set()

    for testName in listBugs:
        if re.search('fdo[0-9]*-[0-9].', testName):
            tdfBugIds.add(str(testName.split('fdo')[1].split('-')[0]))
        elif re.search('tdf[0-9]*-[0-9].', testName):
            tdfBugIds.add(str(testName.split('tdf')[1].split('-')[0]))

    for i in range(0, len(tdfBugIds), 200):
        subList = list(tdfBugIds)[i: i + 200]
        url = rest_url + ",".join(str(x) for x in subList) + field
        output = ast.literal_eval(requests.get(url).text)['bugs']
        for item in output:
            if item['status'] == 'NEW':
                if 'fileopen' in item['summary'].lower() or 'import' in item['summary'].lower():
                    results['import'].append(str(item['id']))
                elif 'filesave' in item['summary'].lower() or 'export' in item['summary'].lower():
                    results['export'].append(str(item['id']))

    return results

def create_doc_with_styles():
    textdoc = OpenDocumentSpreadsheet()

    # Create automatic styles for the column widths.
    # ODF Standard section 15.9.1
    nameColStyle = Style(name="nameColStyle", family="table-column")
    nameColStyle.addElement(TableColumnProperties(columnwidth="8cm"))
    textdoc.automaticstyles.addElement(nameColStyle)

    tagColStyle = Style(name="tagColStyle", family="table-column")
    tagColStyle.addElement(TableColumnProperties(columnwidth="5cm"))
    tagColStyle.addElement(ParagraphProperties(textalign="left")) #??
    textdoc.automaticstyles.addElement(tagColStyle)

    rankColStyle = Style(name="rankColStyle", family="table-column")
    rankColStyle.addElement(TableColumnProperties(columnwidth="1.5cm"))
    rankColStyle.addElement(ParagraphProperties(textalign="center")) #??
    textdoc.automaticstyles.addElement(rankColStyle)

    valColStyle = Style(name="valColStyle", family="table-column")
    valColStyle.addElement(TableColumnProperties(columnwidth="0.9cm"))
    valColStyle.addElement(ParagraphProperties(textalign="center")) #??
    textdoc.automaticstyles.addElement(valColStyle)

    linkColStyle = Style(name="linkColStyle", family="table-column")
    linkColStyle.addElement(TableColumnProperties(columnwidth="0.3cm"))
    linkColStyle.addElement(ParagraphProperties(textalign="center")) #??
    textdoc.automaticstyles.addElement(linkColStyle)

    # Create a style for the table content. One we can modify
    # later in the word processor.
    tablecontents = Style(name="tablecontents", family="paragraph")
    tablecontents.addElement(ParagraphProperties(numberlines="false", linenumber="0"))
    tablecontents.addElement(TextProperties(fontweight="bold"))
    textdoc.styles.addElement(tablecontents)


    TH = Style(name="THstyle",family="table-cell", parentstylename='Standard', displayname="Table Header")
    TH.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(TH)

    C0 = Style(name="C0style",family="table-cell", parentstylename='Standard', displayname="Color style 0")
    C0.addElement(TableCellProperties(backgroundcolor="#00FF00"))
    C0.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(C0)

    Csep = Style(name="Csepstyle",family="table-cell", parentstylename='Standard', displayname="Color style Sep ")
    Csep.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(Csep)

    C1 = Style(name="C1style",family="table-cell", parentstylename='Standard', displayname="Color style 1")
    C1.addElement(TableCellProperties(backgroundcolor="#00FF00"))
    C1.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(C1)

    C2 = Style(name="C2style",family="table-cell", parentstylename='Standard', displayname="Color style 2")
    C2.addElement(TableCellProperties(backgroundcolor="#FFFF00"))
    C2.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(C2)

    C3 = Style(name="C3style",family="table-cell", parentstylename='Standard', displayname="Color style 3")
    C3.addElement(TableCellProperties(backgroundcolor="#FFAA00"))
    C3.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(C3)

    C4 = Style(name="C4style",family="table-cell", parentstylename='Standard', displayname="Color style 4")
    C4.addElement(TableCellProperties(backgroundcolor="#FF0000"))
    C4.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(C4)

    C5 = Style(name="C5style",family="table-cell", parentstylename='Standard', displayname="Color style 5")
    C5.addElement(TableCellProperties(backgroundcolor="#FF0000"))
    C5.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(C5)

    CB = Style(name="CBstyle",family="table-cell", parentstylename='Standard', displayname="Color style blue")
    CB.addElement(TableCellProperties(backgroundcolor="#8888DD"))
    CB.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(CB)

    rankCellStyle = Style(name="rankCellStyle",family="table-cell", parentstylename='Standard', displayname="rankCellStyle")
    rankCellStyle.addElement(ParagraphProperties(textalign="center"))
    textdoc.styles.addElement(rankCellStyle)

    return textdoc

def loadResults(oldFile, newFile):
    results = { 'import':
            {
                'new': {'tests':{}},
                'old': {'tests':{}},
            },
            'export':
            {
                'new': {'tests':{}},
                'old': {'tests':{}},
            }
    }

    loadFile(newFile, results, 'new')

    if not checkOdf:
        loadFile(oldFile, results, 'old')

    return results

def loadFile(path, results, group):
    filePath = os.path.join(path, 'results.txt')
    results['import'][group]['path'] = path
    results['export'][group]['path'] = path

    with open(filePath, 'r') as f:
        lines = f.readlines()
        for line in lines:
            testName = os.path.split(line.split(":")[0].strip())[1]
            values = line.split(":")[1].strip().split(";")
            if 'import.pdf' in testName:
                results['import'][group]['tests'][testName] = values
            else:
                results['export'][group]['tests'][testName] = values

    return results

def valToGrade(data):
    """ get grade for individual observed measures
    """
    #if checking roundtrip, print is ['', '', '', '', ''] or viceversa
    if not data or  '' in data or data[-1] == 'timeout':
        return [8,8,8,8]

    if data[-1] == "empty":
        return [6,6,6,6]
    if data[-1] == "open":
        return [7,7,7,7]
    FDEVal=5
    for i in range(len(FDEMax)):
        if FDEMax[i] >float(data[0]):
            FDEVal=i
            break
    HLPEVal=5
    for i in range(len(HLPEMax)):
        if HLPEMax[i] > float(data[1]):
            HLPEVal=i
            break
    THEVal=5
    for i in range(len(THEMax)):
        if THEMax[i] > float(data[2]):
            THEVal=i
            break
    LNDVal=5
    for i in range(len(LNDMax)):
        if LNDMax[i] > abs(float(data[3])):
            LNDVal=i
            break
    return [FDEVal, HLPEVal, THEVal, LNDVal]

def addAnn(txt):
    ann=Annotation(width="10cm")
    annp = P(stylename="tablecontents", text=str(txt))
    ann.addElement(annp)
    return ann

def getRsltTable(testType, results, tdfBugs):

    targetAppsSel= results[testType].keys()

    # Start the table, and describe the columns
    table = Table(name=testType)
    table.addElement(TableColumn(numbercolumnsrepeated=1, stylename="nameColStyle"))
    table.addElement(TableColumn(stylename="linkColStyle"))
    if checkOdf:
        table.addElement(TableColumn(numbercolumnsrepeated=3,stylename="rankColStyle"))
    else:
        table.addElement(TableColumn(numbercolumnsrepeated=4,stylename="rankColStyle"))

    for i in targetAppsSel:
        for i in range(len(testLabelsShort)-1):
            table.addElement(TableColumn(stylename="valColStyle"))
            table.addElement(TableColumn(stylename="linkColStyle"))
        table.addElement(TableColumn(stylename="rankColStyle"))
        table.addElement(TableColumn(stylename="linkColStyle"))
    table.addElement(TableColumn(stylename="rankColStyle"))
    table.addElement(TableColumn(stylename="tagColStyle"))
    table.addElement(TableColumn(stylename="tagColStyle"))
    table.addElement(TableColumn(stylename="tagColStyle"))

    #First row: application names
    tr = TableRow()
    table.addElement(tr)
    tc = TableCell() #empty cell
    tr.addElement(tc)
    tc = TableCell() #empty cell
    tr.addElement(tc)
    tc = TableCell() #empty cell
    tr.addElement(tc)
    tc = TableCell() #empty cell
    tr.addElement(tc)
    tc = TableCell() #empty cell
    tr.addElement(tc)
    appcolumns=len(testLabelsShort)
    for a in targetAppsSel:
        tc = TableCell(numbercolumnsspanned=2*(appcolumns-1), stylename="THstyle")
        tr.addElement(tc)
        p = P(stylename="tablecontents",text="Target: " + results[testType][a]['path'])
        tc.addElement(p)
        for i in range(2*(appcolumns-1)-1): # create empty cells for the merged one
            tc = TableCell()
            tr.addElement(tc)
        tc = TableCell(stylename="Csepstyle")
        tr.addElement(tc)
    #Second row: test names
    tr = TableRow()
    table.addElement(tr)
    tc = TableCell(stylename="THstyle") #empty cell
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Test case")
    tc.addElement(p)
    tc = TableCell(stylename="THstyle") #empty cell
    tr.addElement(tc)
    if not checkOdf:
        tc = TableCell(stylename="THstyle") #empty cell
        tr.addElement(tc)
        p = P(stylename="tablecontents",text="P/R")
        tc.addElement(p)
        tc.addElement(addAnn("Negative: progression, positive: regression, 0: no change"))
    tc = TableCell(stylename="THstyle") #empty cell
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Max last")
    tc.addElement(p)
    tc.addElement(addAnn("Max grade for the last LO version"))
    tc = TableCell(stylename="THstyle") #empty cell
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Sum last")
    tc.addElement(p)
    tc.addElement(addAnn("Sum of grades for the last LO version"))
    tc = TableCell(stylename="THstyle") #empty cell
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Sum all")
    tc.addElement(p)
    tc.addElement(addAnn("Sum of grades for all tested versions"))

    for a in targetAppsSel:
        for tl in range(1, len(testLabelsShort)):   # we do not show the PPOI value
            tc = TableCell(numbercolumnsspanned=2,stylename="THstyle")
            tr.addElement(tc)
            p = P(stylename="tablecontents",text=testLabelsShort[-tl])
            tc.addElement(p)
            tc.addElement(addAnn(testAnnotation[testLabelsShort[-tl]]))
            tc = TableCell()    #the merged cell
            tr.addElement(tc)
        tc = TableCell(stylename="Csepstyle")
        tr.addElement(tc)
        tc = TableCell(stylename="THstyle")
        tr.addElement(tc)

    total = 0
    totalRegressions = 0
    totalEmpty = 0
    totalTimeOut = 0
    for testcase in results[testType]['new']['tests'].keys():
        try:
            agrades = np.array([valToGrade(results[testType][a]['tests'][testcase][1:]) for a in targetAppsSel])
            if np.array_equal(agrades[0], [8,8,8,8]):
                continue
            lastgrade=agrades[-1]
            maxgrade=agrades.max(axis=0)
            mingrade=agrades.min(axis=0)
        except KeyError:
            # In case a testcase is in the first csv but not in the second one
            continue
        total += 1

        #identify regressions and progressions
        progreg='x'

        if (lastgrade>mingrade).any():  #We have regression
            progreg=str(sum(lastgrade-mingrade))
        else:
            progreg=str(sum(lastgrade-maxgrade))

        if checkRegressions and (int(progreg) >= -1 and not np.array_equal(agrades[0], [7,7,7,7])):
            continue

        #Looking for improvements, we only care about fdo bugs
        if checkImprovements and ( int(progreg) < 1 or \
                ((not re.search('fdo[0-9]*-[0-9].', testcase) or \
                testType == 'import' and testcase.split('fdo')[1].split('-')[0] not in tdfBugs['import'] or \
                testType == 'export' and testcase.split('fdo')[1].split('-')[0] not in tdfBugs['export']) and
                (not re.search('tdf[0-9]*-[0-9].', testcase) or \
                testType == 'import' and testcase.split('tdf')[1].split('-')[0] not in tdfBugs['import'] or \
                testType == 'export' and testcase.split('tdf')[1].split('-')[0] not in tdfBugs['export']))):
            continue

        if checkOdf:
            allsum = sum([sum(valToGrade(results[testType][a]['tests'][testcase][1:])) for a in targetAppsSel] )
            if allsum <= 5:
                continue

        name = testcase.split("/",1)[-1].split(".")[0]

        #Avoid showing import regressions as export regressions
        if checkRegressions:
            if testType == "import":
                lImportReg.append(name)
            elif testType == "export" and not np.array_equal(agrades[0], [7,7,7,7]):
                if name in lImportReg or name in lExportReg:
                    continue
                lExportReg.append(name)

        if int(progreg) < 0:
            totalRegressions += 1
        elif np.array_equal(agrades[0], [6,6,6,6]):
            totalEmpty += 1
        elif np.array_equal(agrades[0], [7,7,7,7]):
            totalTimeOut += 1

        #testcase=testcase.split('/')[1]
        tr = TableRow()
        table.addElement(tr)
        tc = TableCell()
        tr.addElement(tc)
        p = P(stylename="tablecontents")
        #TODO: Fix doc link in roundtrip
        if re.search('fdo[0-9]*-[0-9].', testcase):
            ref = 'https://bugs.documentfoundation.org/show_bug.cgi?id=' + str(testcase.split('fdo')[1].split('-')[0])
            link = A(type="simple",href="%s"%ref, text=testcase)
        elif re.search('tdf[0-9]*-[0-9].', testcase):
            ref = 'https://bugs.documentfoundation.org/show_bug.cgi?id=' + str(testcase.split('tdf')[1].split('-')[0])
            link = A(type="simple",href="%s"%ref, text=testcase)
        elif re.search('ooo[0-9]*-[0-9].', testcase):
            ref = 'https://bz.apache.org/ooo/show_bug.cgi?id=' + str(testcase.split('ooo')[1].split('-')[0])
            link = A(type="simple",href="%s"%ref, text=testcase)
        elif re.search('abi[0-9]*-[0-9].', testcase):
            ref = 'https://bugzilla.abisource.com/show_bug.cgi?id=' + str(testcase.split('abi')[1].split('-')[0])
            link = A(type="simple",href="%s"%ref, text=testcase)
        elif re.search('kde[0-9]*-[0-9].', testcase):
            ref = 'https://bugs.kde.org/show_bug.cgi?id=' + str(testcase.split('kde')[1].split('-')[0])
            link = A(type="simple",href="%s"%ref, text=testcase)
        elif re.search('moz[0-9]*-[0-9].', testcase):
            ref = 'https://bugzilla.mozilla.org/show_bug.cgi?id=' + str(testcase.split('moz')[1].split('-')[0])
            link = A(type="simple",href="%s"%ref, text=testcase)
        elif re.search('gentoo[0-9]*-[0-9].', testcase):
            ref = 'https://bugs.gentoo.org/show_bug.cgi?id=' + str(testcase.split('gentoo')[1].split('-')[0])
            link = A(type="simple",href="%s"%ref, text=testcase)
        else:
            link = A(type="simple",href="%s"%('../' + testcase), text=testcase)
        p.addElement(link)
        tc.addElement(p)

        tComparison = TableCell(stylename="THstyle")
        tr.addElement(tComparison)

        if not checkOdf:
            tc = TableCell(valuetype="float", value=progreg)
            tr.addElement(tc)

        # max last
        lastmax = max([valToGrade(results[testType][a]['tests'][testcase][1:]) for a in targetAppsSel][-1])
        tc = TableCell(valuetype="float", value=str(lastmax))
        tr.addElement(tc)

        # sum last
        lastsum = sum([valToGrade(results[testType][a]['tests'][testcase][1:]) for a in targetAppsSel][-1])
        tc = TableCell(valuetype="float", value=str(lastsum))
        tr.addElement(tc)

        # sum all
        allsum = sum([sum(valToGrade(results[testType][a]['tests'][testcase][1:])) for a in targetAppsSel] )
        tc = TableCell(valuetype="float", value=str(allsum))
        tr.addElement(tc)

        for a in targetAppsSel:
            grades = valToGrade(results[testType][a]['tests'][testcase][1:])
            pdfPath = os.path.join(results[testType][a]['path'], testcase)

            if not checkOdf and a == 'new':
                oldFile = os.path.join(results[testType]['old']['path'], testcase).split('-pair.pdf')[0]
                if os.path.exists(oldFile):
                    newFile = os.path.join(results[testType]['new']['path'], testcase).split('-pair.pdf')[0]
                    if os.path.exists(newFile):
                        outputFile = oldFile + '-comparison.pdf'
                        if not os.path.exists(outputFile):
                            print("Creating " + outputFile)
                            create_overlayPDF(oldFile, newFile, outputFile)

                        p = P(stylename="tablecontents",text="")
                        comparisonLink = A(type="simple",href='../' + outputFile, text=">")
                        p.addElement(comparisonLink)
                        tComparison.addElement(p)

            # Add link only once
            linkAdded = False
            for grade in reversed(grades):   # we do not show the PPOI value
                if max(grades) > 1:
                    tc = TableCell(valuetype="float", value=str(grade), stylename='C'+str(int(grade))+'style')
                else:
                    tc = TableCell(valuetype="float", value=str(grade), stylename='CBstyle')
                tr.addElement(tc)
                tc = TableCell(stylename="THstyle")
                tr.addElement(tc)
                p = P(stylename="tablecontents")

                if not linkAdded and os.path.exists(pdfPath):
                    linkAdded = True
                    link = A(type="simple",href='../' + pdfPath, text=">")
                    p.addElement(link)
                    tc.addElement(p)
                    if checkOdf:
                        pComparison = P(stylename="tablecontents")
                        linkComparison = A(type="simple",href='../' + pdfPath, text=">")
                        pComparison.addElement(linkComparison)
                        tComparison.addElement(pComparison)

            tc = TableCell(stylename="THstyle")

            sumall = sum(valToGrade(results[testType][a]['tests'][testcase][1:]))
            if grades == [7,7,7,7]:
                p = P(stylename="tablecontents",text="timeout")
                if testType == "export":
                    gradesPrint = valToGrade(values[testcase][a.replace(testType, 'print')][1:])
                    if gradesPrint != [7,7,7,7]:
                        p = P(stylename="tablecontents",text="corrupted")
            elif grades == [6,6,6,6]:
                p = P(stylename="tablecontents",text="empty")
            elif sumall <= 8:
                if testType == "import":
                    goodDocuments.append(testcase)
                    p = P(stylename="tablecontents",text="good import")
                elif testType == "export":
                    if testcase in goodDocuments:
                        p = P(stylename="tablecontents",text="good import, good export")
                    elif testcase in badDocuments:
                        p = P(stylename="tablecontents",text="bad import, good export")
            elif sumall <= 20:
                if testType == "export":
                    if testcase in goodDocuments:
                        p = P(stylename="tablecontents",text="good import, bad export")
                        badDocuments.append(testcase)
                    elif testcase in badDocuments:
                        p = P(stylename="tablecontents",text="bad import, bad export")
                elif testType == "import":
                    badDocuments.append(testcase)
                    p = P(stylename="tablecontents",text="bad import")
            else:
                p = P(stylename="tablecontents",text="")

            tc.addElement(p)
            tr.addElement(tc)
            tc = TableCell(stylename="THstyle")
            tr.addElement(tc)

    tr = TableRow()
    table.addElement(tr)
    tr = TableRow()
    table.addElement(tr)
    tc = TableCell()
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Total compared bugs: " + str(total))
    tc.addElement(p)

    tr = TableRow()
    table.addElement(tr)
    tc = TableCell()
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Total number of regressions: " + str(totalRegressions))
    tc.addElement(p)

    tr = TableRow()
    table.addElement(tr)
    tc = TableCell()
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Total number of empty files: " + str(totalEmpty))
    tc.addElement(p)

    tr = TableRow()
    table.addElement(tr)
    tc = TableCell()
    tr.addElement(tc)
    p = P(stylename="tablecontents",text="Total number of Timeouts: " + str(totalTimeOut))
    tc.addElement(p)

    return table

if __name__ == "__main__":
    if not os.getcwd().endswith('office-interoperability-tools'):
        print('ERROR: call this script from office-interoperability-tools root directory')
        sys.exit(1)

    oldPath = ""
    newPath = ""
    ofname = ""
    checkRegressions = False
    checkImprovements = False
    checkOdf = False

    parsecmd()

    results = loadResults(oldPath, newPath)

    tdfBugs = None
    if checkImprovements:
        tdfBugs = check_tdf_bugs(results['import']['new'])

    textdoc = create_doc_with_styles()

    goodDocuments = []
    badDocuments = []
    lImportReg = []
    lExportReg = []
    importTable = getRsltTable("import", results, tdfBugs)
    textdoc.spreadsheet.addElement(importTable)
    exportTable = getRsltTable("export", results, tdfBugs)
    textdoc.spreadsheet.addElement(exportTable)

    textdoc.save(ofname)

