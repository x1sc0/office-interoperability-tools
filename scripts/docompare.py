#! /usr/bin/python
# docompare.py v1.0 - 2013-09-30
#
# This script generates csv files with numeric evaluations and pair views from printed document files
#
# Copyright (C) 2013 Milos Sramek <milos.sramek@soit.sk>
# Licensed under the GNU LGPL v3 - http://www.gnu.org/licenses/gpl.html
# - or any later version.
#
#from __future__ import print_function
import numpy as np
import cv2, Image
import sys, getopt, os, tempfile
try:
    import ipdb
except ImportError:
    pass
#if tifffile.__version__ == '2015.08.17' (installed on Ubuntu using apt)
#from tifffile import TIFFfile
#if tifffile.__version__ == '0.9.2' (installed on Ubuntu using pip)
from tifffile import TiffFile
from scipy import ndimage

class DoException(Exception):
    def __init__(self, what):
        self.what = what

def disp(iimg, label = "", gray=False):
    """ Display an image using pylab
    """
    try:
        import pylab
        dimage = iimg.copy()
        if iimg.ndim==3:
            dimage[...,0] = iimg[...,2]
            dimage[...,2] = iimg[...,0]

        pylab.imshow(dimage, interpolation='none')
        if gray:
            pylab.gray()
        #pylab.gca().format_coord = format_coord
        pylab.text(1500, -30, label)
        pylab.axis('off')
        pylab.show()
    except ImportError:
        print "Module pylab not available"

def distancetransf(image):
    if image.dtype=='bool':
        return ndimage.distance_transform_edt(1-image.astype(np.int8))
    else:
        return ndimage.distance_transform_edt(1-image)

def tmpname():
    f = tempfile.NamedTemporaryFile(delete=True)
    f.close()
    return f.name

def toBin(img, thr=200):
    #return (img < thr).astype(np.uint8)
    return (img < thr).max(axis=2).astype(np.uint8)

def GetLine(img, seglist, seg):
    """
    get the 'seg' line of the 'seglist' list of lines fron the image
    """
    return img[seglist[seg][0]:seglist[seg][0]+seglist[seg][1]].copy()


def GetLineSegments(itrim):
    """
    Segment an image in lines and interline spaces
    Returns lists of both (position width)
    """
        # sum along pixel lines
    asum = np.sum(itrim, axis=1)
    abin = asum > 0
    sp = []
    tx = []
    lastval=-1
    lastpos=-1
    for i in range(0, abin.size):
        if abin[i] != lastval:
            lastval = abin[i]
            if lastval:
                tx.append(np.array((i,0)))
                if i>0:
                    sp[-1][1] = i-sp[-1][0]
            else:
                sp.append(np.array((i,0)))
                if i>0:
                    tx[-1][1] = i-tx[-1][0]
    # set the last segment lenght
    if tx[-1][1] == 0:
        tx[-1][1] = itrim.shape[0] - tx[-1][0]
    if sp==[]:# empy if there is just one line in the image
        sp.append(np.array((0,0)))
    else:
        if sp[-1][1] == 0:
            sp[-1][1] = itrim.shape[0] - sp[-1][0]
    return tx, sp

def pdf2array(pdffile, res=300):
    """
    read pdf file and convert it to numpy array
    return list of pages and dimensions
        color order is BGR
    """
    tname = tmpname()+'.tif'
    cmd = 'cat %s 2>/dev/null | gs -dQUIET -dNOPAUSE -sDEVICE=tiff24nc -r%d -sOutputFile=%s - 2>/dev/null'%(pdffile, res, tname)
    os.system(cmd)

    if not os.path.exists(tname):
        return None, None

    imgfile = TiffFile(tname)
    pages = [p.asarray() for p in imgfile.pages[:10]]    # first ten pages only

    shapes = [p.shape for p in imgfile.pages[:10]]
    os.remove(tname)
    return pages, shapes

def makeSingle(pages, shapes):
    ''' merge pages to one image'''
    height=0
    width=0

    for s in shapes:
        height += s[0]
        width = max(width, s[1])
    #bigpage = np.zeros((height,width,3), dtype=pages[0].dtype)
    # white backgroud, in ordet to ignore page width differences within a document
    bigpage = 255*np.ones((height,width,3), dtype=pages[0].dtype)
    pos=0
    for p, s in zip(pages, shapes):
        bigpage[pos:pos+s[0],0:s[1],:]=p[0:s[0],0:s[1],:]
        pos += s[0]
    return bigpage

def getPagePixelOverlayIndex(iarray1, iarray2):
    """
    Compute the pixel overlap index without page or line alignment
    """
    (ystart, xstart), (ystop, xstop) = getBBox(iarray1, iarray2)
    itrim1 = iarray1[ystart:ystop, xstart:xstop].astype(np.uint8)
    itrim2 = iarray2[ystart:ystop, xstart:xstop].astype(np.uint8)
    #return 2.0* np.sum((itrim1+itrim2) > 1) / (np.sum(itrim1) + np.sum(itrim2))
    diff = itrim1 != itrim2
    ovl = 100*(1.0 - float(np.sum(diff)) / (np.sum(itrim2) + np.sum(itrim1)))
    rslt = 'PagePixelOvelayIndex[%%]: %2.1f : '%ovl
    return rslt

def mergeSingle(ml, tx0, sp0, tx1, sp1):
    """ merge one blob"""
    #print np.array(tx0)[:,1]
    #print np.array(tx1)[:,1]
    tx0[ml][1] += sp0[ml][1] + tx0[ml+1][1]
    tx0 = tx0[:ml+1]+tx0[ml+2:]
    sp0 = sp0[:ml]+sp0[ml+1:]
    #print np.array(tx0)[:,1]
    #print np.array(tx1)[:,1]
    #print
    return tx0, sp0

def pr(tx, p=1):
    print len(tx), np.array(tx)[:,p]

def mergeLocation(tx0, sp0, tx1, sp1):
    """
    find merge location in (tx0,sp0)
    """
    if len(tx0) < 2:
        return 9999
    txmin=min( np.array(tx0)[:,1] ) # minimal line heigh, used as detection threshold
    txx=[]
    for i in range(min(len(tx0)-1,len(tx1))):
        tx=tx0[i][1] + sp0[i][1] + tx0[i+1][1]
        txx.append(tx - tx1[i][1])
    cc = np.argmin(txx)
    if txx[cc] < txmin/3: #expected to be near 0
        return cc
    else:
        return 9999

def mergeBlobs(tx0, sp0, tx1, sp1):
    changed = True
    ml0 = mergeLocation(tx0, sp0, tx1, sp1)
    ml1 = mergeLocation(tx1, sp1, tx0, sp0)
    while ml0 < len(tx0) or ml1 < len(tx1):
        if ml0 < ml1:
            tx0, sp0 = mergeSingle(ml0, tx0, sp0, tx1, sp1)
        else:
            tx1, sp1 = mergeSingle(ml1, tx1, sp1, tx0, sp0)
        ml0 = mergeLocation(tx0, sp0, tx1, sp1)
        ml1 = mergeLocation(tx1, sp1, tx0, sp0)
    return tx0, sp0, tx1, sp1

def ovlLine(im1, im2, shift=0):
    """
    overlay two lines with different height
    """
    im=np.zeros((max(im1.shape[0],im2.shape[0]), im1.shape[1], 3))
    if im1.shape[0] < im2.shape[0]:
        im[:]=1
        im[shift:im1.shape[0]+shift,:,0] = 1-im1
        im[:,:,1] = 1-im2
        im[:,:,2] = 1-im2
    else:
        im[:,:,0] = 1-im1
        im[shift:,:,1] = 1-im2
        im[shift:,:,2] = 1-im2
    return im

def align(l1, l2, axis):
    if axis == 1: #horizontal alignment, we do not care about the right line end
    #cw = min(l2.shape[1],l1.shape[1])
    #l1 = l1[:,:cw]
    #l2 = l2[:,:cw]
        #compute correlation
        sc1 = np.sum(l1, axis=1-axis)
        sc2 = np.sum(l2, axis=1-axis)
        cor = np.correlate(sc1,sc2,"same")
        posErr =  np.argmax(cor)-sc1.shape[0]/2
        #place at right position
        if posErr > 0:
            l2c = l2.copy()
            l2c[:]=0
            l2c[:,posErr:] = l2[:,:-posErr]
            l2 = l2c
        elif posErr < 0:
            l1c = l1.copy()
            l1c[:]=0
            l1c[:,-posErr:] = l1[:,:posErr]
            l1=l1c
    else: #vertical alignment, we cate about both ends
        #compute correlation
        sc1 = np.sum(l1, axis=1-axis)
        sc2 = np.sum(l2, axis=1-axis)
        cor = np.correlate(sc1,sc2,"same")
        posErr =  np.argmax(cor)-sc1.shape[0]/2
        #place at right position
        if posErr > 0:
            l2c=l2.copy()
            l2c[:]=0
            l2c[posErr:,:] = l2[:-posErr,:]
            l2 = l2c
        elif posErr < 0:
            l1c=l1.copy()
            l1c[:]=0
            l1c[-posErr:,:]=l1[:posErr,:]
            l1 = l1c
    return posErr, l1, l2

def alignLineIndex(l1, l2, halign=True):
    """
    compute several line similarity measures with horizontal and vertical alignment
    """

    #make the same width
    #ipdb.set_trace()
    if l1.shape[1] > l2.shape[1]:
        l2=np.pad(l2,((0,0),(0,l1.shape[1] - l2.shape[1])),'constant',constant_values=0)
    else:
        l1=np.pad(l1,((0,0),(0,l2.shape[1] - l1.shape[1])),'constant',constant_values=0)

    #make the same height
    if l1.shape[0] > l2.shape[0]:
        l2=np.pad(l2,((0,l1.shape[0] - l2.shape[0]), (0,0)),'constant',constant_values=0)
    else:
        l1=np.pad(l1,((0,l2.shape[0] - l1.shape[0]), (0,0)),'constant',constant_values=0)

    # align in the horizontal direction
    horizPosErr = 0
    if halign:
        horizPosErr,l1,l2=align(l1,l2,1)

    # align in the vertical direction
    vertPosErr,ll1,ll2=align(l1,l2,0)

        #overlap index
    diff = ll1 != ll2
    if np.sum(ll2) + np.sum(ll1) == 0:
        ovlapindex = 1.0
    else:
        ovlapindex = 1.0 - float(np.sum(diff)) / (np.sum(ll2) + np.sum(ll1))


    ld1 = distancetransf(ll1)
    ld2 = distancetransf(ll2)
    # if one of the images has only 1 value ignore negative distances
    if ll1.all() or ll2.all():
        ld1[np.where(ld1<0)]=0
        ld2[np.where(ld2<0)]=0

    overlayedLines = ovlLine(ll1, ll2)
    return overlayedLines, (abs(horizPosErr), ovlapindex, np.average(abs(ld1-ld2)), np.max(abs(ld1-ld2)))

def lineIndexPage(iarray0, iarray1):
    """
    compute similarity measures for each page line, compute statistics and return it in a string form
    """
    # crop the first image and segment it to lines
    (ystart0, xstart0), (ystop0, xstop0) = getBBox(iarray0)
    itrim0 = iarray0[ystart0:ystop0, xstart0:xstop0].astype(np.uint8)
    tx0, sp0 = GetLineSegments(itrim0)

    # crop the second image and segment it to lines
    (ystart1, xstart1), (ystop1, xstop1) = getBBox(iarray1)
    itrim1 = iarray1[ystart1:ystop1, xstart1:xstop1].astype(np.uint8)
    tx1, sp1 = GetLineSegments(itrim1)

    # detect merged lines in one set and merge them in the other
    tx0, sp0, tx1, sp1 = mergeBlobs(tx0, sp0, tx1, sp1)

        # create arrays with aligned lines
    vh_lines=[] # horizontally aligned lines
    v_lines=[]  # original lines from iarray1
    indices=[]
    for i in range(min(len(tx0),len(tx1))):
        l0= GetLine(itrim0, tx0, i)
        l1 = GetLine(itrim1, tx1, i)
        cline, ind = alignLineIndex(l0, l1)
        vh_lines.append(cline)
        indices.append(ind)
        #cline, ind = alignLineIndex(GetLine(itrim0, tx0, i), GetLine(itrim1, tx1, i), halign=False)
        cline, ind = alignLineIndex(l0, l1, halign=False)
        v_lines.append(cline)
    indices = np.array(indices)

    #create a page view to display vertically and horizontally adjusted overlays, taking line spaces from the source (first) page
    # height of the output page: sum of overlayed blobs + sum of spaces from image 1
    outheight = ystart0 + sum([b.shape[0] for b in vh_lines])+ sum(np.array(sp0)[:,1]) +            (iarray0.shape[0] - ystop0)+10
        #outheight = ystart0 + sum([b.shape[0] for b in vh_lines])+ sum(np.array(sp0)[:,1])
    vh_page = np.zeros((outheight, iarray0.shape[1], 3), dtype=np.uint8)
    vh_page[:] = 1

    # create page of horizontally aligned lines
    ar=ystart0    #the actual row
    for i in range(min(len(tx0),len(tx1))):
        #we cut  vh_lines[i] from right, if it is too long
        awidth= vh_lines[i].shape[1]
        if xstart0+vh_lines[i].shape[1] > vh_page.shape[1]:
            awidth = vh_page.shape[1] - xstart0
        vh_page[ar:ar+vh_lines[i].shape[0],xstart0:xstart0+vh_lines[i].shape[1]] = vh_lines[i][:,:awidth,:]
        if i < len(sp0):
            ar += vh_lines[i].shape[0] + sp0[i][1]
        else:
            ar += vh_lines[i].shape[0]

    #create a page view to display vertically adjusted overlays, taking line spaces from the source (first) page
    # height of the output page: sum of overlayed blobs + sum of spaces from image 1
    outheight = ystart0 + sum([b.shape[0] for b in v_lines])+ sum(np.array(sp0)[:,1]) + (iarray0.shape[0] - ystop0) + 10
    #outheight = ystart0 + sum([b.shape[0] for b in v_lines])+ sum(np.array(sp0)[:,1])
    v_page = np.zeros((outheight, iarray0.shape[1], 3), dtype=np.uint8)
    v_page[:] = 1

    # create page of the original lines from iarray1
    ar=ystart0    #the actual row
    for i in range(min(len(tx0),len(tx1))):
        #we cut  vh_lines[i] from right, if it is too long
        awidth= vh_lines[i].shape[1]
        if xstart0+vh_lines[i].shape[1] > vh_page.shape[1]:
            awidth = vh_page.shape[1] - xstart0
        v_page[ar:ar+v_lines[i].shape[0],xstart0:xstart0+v_lines[i].shape[1]] = v_lines[i][:,:awidth,:]
        #ar += v_lines[i].shape[0] + sp0[i][1]
        if i < len(sp0):
            ar += v_lines[i].shape[0] + sp0[i][1]
        else:
            ar += v_lines[i].shape[0]

    #height error in pixels
    heightErr=abs(float(itrim0.shape[0]-itrim1.shape[0]))
    #normalized height error in pixels
    # normalize the error to standard a4 height - overestimates error of small pages, so commented out
    #heightErr = heightErr * (dpi*a4height/i2mm)/float(itrim0.shape[0]+itrim1.shape[0])

    #If changing the rslt format, adjust the gentestviews.sh script accordingly
    linersltDist = 'FeatureDistanceError[mm]: %2.1f '
    #linersltDist = linersltDist%(px2mm(np.max(indices[:,3])))
    linersltDist = px2mm(np.max(indices[:,3]))
    linersltHPos = 'HorizLinePositionError[mm]: %2.2f '
    #linersltHPos = linersltHPos%(px2mm(np.max(indices[:,0])))
    linersltHPos = px2mm(np.max(indices[:,0]))
    pagerslt = ' TextHeightError[mm]: %2.2f | LineNumDifference: %2d'
    #pagerslt = pagerslt%((px2mm(heightErr), len(tx1)-len(tx0)))

    return 1-vh_page, 1-v_page, linersltDist, linersltHPos, px2mm(heightErr), len(tx1)-len(tx0)

def annotateImg(img, color, size, position, text):
    cv2.putText(img, text, position, cv2.FONT_HERSHEY_PLAIN, size, color, thickness = 2)
    return img

def mergeSide(img1, img2):
    ''' place two images side-by-side'''
    offset=10
    ishape = img1.shape
    nshape=(max(img1.shape[0], img2.shape[0]), img1.shape[1]+img2.shape[1]+offset, 3) #shape for numpy
    big=np.zeros(nshape, dtype=np.uint8)
    big[:]=200
    big[:img1.shape[0],:img1.shape[1]]=img1
    big[:img2.shape[0],img1.shape[1]+offset:]=img2
    return big

def genside (img1, img2, height, width, name1, name2, txt1, txt2):
    """
    create a side-by-side view
    img1, img2: images
    name1, name2: their names
    txt1, txt2: some text
    """
    if len(img1.shape)==2:
        cimg1 = np.zeros((img1.shape[0], img1.shape[1], 3), dtype=np.uint8)
        cimg1[...,0] = img1
        cimg1[...,1] = img1
        cimg1[...,2] = img1
    else:
        cimg1 = img1
    if len(img2.shape)==2:
        cimg2 = np.zeros((img2.shape[0], img2.shape[1], 3), dtype=np.uint8)
        cimg2[...,0] = img2
        cimg2[...,1] = img2
        cimg2[...,2] = img2
    else:
        cimg2 = img2

    if annotated:
        cimg1=annotateImg(cimg1, (0,0,255), 2, (100, 70), 'Source: '+name1)
        #cimg1=annotateImg(cimg1, (0,0,255), 2, (100, 130), txt1)
        cimg2=annotateImg(cimg2, (0,0,255), 2, (100, 70), 'Target: '+name2)
        #cimg2=annotateImg(cimg2, (0,0,255), 2, (100, 130), txt2)
        cimg = mergeSide(cimg1, cimg2)
    if annotated:
        cimg=annotateImg(cimg, (0,255,0), 2, (100, 130), txt1)
        return cimg

def genoverlay(img1, title, name1, name2, stattxt, img2=None):
    """
    create an overlayed view
    img1, img2: images
    title: kind of title to print
    name1, name2: their names
    txt: text to print below the title
    """

    if img2 is None:
        outimg = 255*(1-img1)
    else:
        s=np.maximum(img1.shape,img2.shape)
        outimg=np.zeros((s[0], s[1], 3), dtype=np.uint8)
        #outimg[:img1.shape[0], :img1.shape[1],0] = (255*(1-img1))
        #outimg[:img2.shape[0], :img2.shape[1],1] = (255*(1-img2))
        #outimg[:img2.shape[0], :img2.shape[1],2] = (255*(1-img2))
        outimg[:img1.shape[0], :img1.shape[1],0] = img1
        outimg[:img2.shape[0], :img2.shape[1],1] = img2
        outimg[:img2.shape[0], :img2.shape[1],2] = img2
        outimg = 255*(1-outimg)
    if annotated:
        outimg = annotateImg(outimg, (0, 0, 255), 2, (100, 50), title)
        txt = "cyan: %s %s"%(sourceid,name1)
        outimg = annotateImg(outimg, (0, 255, 255), 2, (100, 80), txt)
        txt = "red: %s %s"%(targetid,name2)
        outimg = annotateImg(outimg, (255, 0, 0), 2, (100, 110), txt)
        #outimg=annotateImg(outimg, 'blue', mm2px(4), mm2px(4), txt)
        outimg = annotateImg(outimg, (0, 0, 255), 1.3, (100, 140), stattxt)
    return outimg

def mm2px(val):
    """
    convert 'val' im mm to pixels
    """
    global dpi, i2mm
    return int(val*dpi/i2mm)

def getBBox(img1, img2=None):
    """
    get a bounding box of nonzero pixels of one or two images
    """
    if img2 is None:
        B = np.argwhere(img1)
        if B.shape[0] == 0:
            return (0,0),(0,0)
        return B.min(0), B.max(0) + 3
    else:
        min1, max1 = getBBox(img1)
        min2, max2 = getBBox(img2)
        return np.minimum(min1, min2), np.maximum(max1, max2)

def saveRslt(overlayStyle, title, img0, img1, name0, name1, rslt, rsltText, outfile):
    exifcmd = 'exiftool -overwrite_original -Custom1="%s" %s >/dev/null'
    oname = "%s-%s.pdf"%(outfile, overlayStyle)
    if overlayStyle=='s':
        s=np.minimum(img0.shape, img1.shape)
        outimg=genside(img0, img1, s[0], s[1], name0, name1, rsltText.replace('*',' '), '')
    else:
        outimg = genoverlay(img0, title, name0, name1, rslt, img2=img1)
    Image.fromarray(outimg).save(oname)
    os.system(exifcmd%(rsltText, oname))

def px2mm(val):
    """
    convert 'val' in mm to pixels
    """
    global dpi, i2mm
    return val*i2mm/dpi

#testLabelsShort=['PPOI','FDE', 'HLPE', 'THE', 'LND']
def valToGrade(data):
    """ get grade for individual observed measures """
    # error grading
    FDEMax = (0.01,0.5,1,2,4)        #0.5: difference of perfectly fitting AOO/LOO and MS document owing to different character rendering
    HLPEMax = (0.01,5,10,15,20)        #
    THEMax = (0.01,2, 4, 6,8)
    LNDMax = (0.01,0.01,0.01,0.01,0.01)
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
    #print (FDEVal, HLPEVal, THEVal, LNDVal)
    return max((FDEVal, HLPEVal, THEVal, LNDVal))

def usage(desc):
    global outfile, binthr, badThr
    print sys.argv[0]+':',  desc
    print "Usage: ", sys.argv[0], "[options]", "source.pdf target.pdf"
    print "\t-o ................ file to store output view to {default:"+outfile+'}'
    #print "\t-p ................ report the results of individual pages {default: whole document}"
    #print "\t-fp ............... use only the first page {default: all pages}"
    print "\t-d ................ resolution of the converted pdf document {default: "+str(dpi)+"}"
    print "\t-a ................ annotate each page with file name {default: no annotation}"
    print "\t-s ................ side-by-side pages {default: all}"
    print "\t-p ................ page overlay without alignment {default: all}"
    print "\t-l ................ page overlay with vertical line alignment {default: all}"
    print "\t-z ................ page overlay with vertical and horizontal line alignment {default: all}"
    print "\t-t ................ binarization threshold {default: %d}"%binthr
    print "\t-b ................ run within the git bisect process (names of the output files will be modified) {default: normal}"
    print "\t-c ................ Only creates an overlay PDF {default: normal}"
    print "\t-g int ............ threshold used in bisecting to distinguish good/bad (1-5) {default: %d}"%badThr
    print "\t--source-id string  additional info to be printed in rendered pdf {default: none}"
    print "\t--target-id string  additional info to be printed in rendered pdf {default: none}"
    print "\t-v ................ be verbose"
    print "\t-h ................ this usage"

def parsecmd(desc):
    global verbose, dpi, Names, annotated, overlayStyle, outfile, binthr, bisecting, createOverlayPDF, badThr, sourceid, targetid
    try:
            opts, Names = getopt.getopt(sys.argv[1:], "hvalpszbgc:o:d:t:", ["help", "verbose", "source-id=", "target-id="])
    except getopt.GetoptError as err:
        # print help information and exit:
        print str(err) # will print something like "option -a not recognized"
        usage(desc)
        sys.exit(2)
    for o, a in opts:
        if o in ("-v", "--verbose"):
            verbose = True
        elif o in ("-h", "--help"):
            usage(desc)
            sys.exit()
        elif o in ("-a"):
            annotated = True
        elif o in ("-b"):
            bisecting = True
        elif o in ("-c"):
            createOverlayPDF = True
        elif o in ("-s"):
            overlayStyle = 's'  #side-by-side
        elif o in ("-p"):
            overlayStyle = 'p'  #page overlay, no alighment
        elif o in ("-l"):
            overlayStyle = 'l'  #page overlay, vertical line alignment
        elif o in ("-z"):
            overlayStyle = 'z'  #page overlay, vertical and horizontal line alignment
        elif o in ("-g"):
            badThr = int(a)
        elif o in ("-o"):
            outfile = a
        elif o in ("-d"):
            dpi = int(a)
        elif o in ("-t"):
            binthr = int(a)
        elif o in ("--source-id"):
            sourceid=a
        elif o in ("--target-id"):
            targetid=a
        else:
            assert False, "unhandled option"
        # ...

#global definitions
a4width=210
a4height=297
i2mm=25.4    # inch to mm conversion
annotated = False
bisecting = False
createOverlayPDF = False
badThr = 3
overlayStyle = 'a'  #output all versions by default
progdesc="Compare two pdf documents and return some statistics"
verbose = False
dpi = 200
binthr = 200
exifcmd = 'exiftool -overwrite_original -Custom1="%s" %s >/dev/null'
outfile = "paircmp.pdf"
sourceid="source"
targetid="target"

def mainfunc():

    global outfile
    parsecmd(progdesc)
    if Names == []:
        usage(progdesc)
        sys.exit(1)

    #if bisecting, use the same name as target (name will be modified)
    if bisecting:
        outfile = Names[1]

    # get rid of extension, if there is one
    if outfile[-4:] == '.pdf':
        outfile = outfile[:-4]
    try:
        #load documents
        pages1, shapes1 = pdf2array(Names[0], dpi)
        if pages1 == None:
            raise DoException("failed to open %s."%(Names[0]))

        badpagetxt=""
        pages2, shapes2 = pdf2array(Names[1], dpi)

        if createOverlayPDF:
            img1 = makeSingle(pages1, shapes1)
            img2 = makeSingle(pages2, shapes2)
            bimg1 = toBin(img1,binthr)
            bimg2 = toBin(img2,binthr)
            outimg = genoverlay(bimg1, '', Names[0], Names[1], '', img2=bimg2)
            rsltText="-:-:-:-:-:-:-:-:-:open"  #dummy result string 10 dashes necessary
            Image.fromarray(outimg).save(outfile+'.pdf')
            os.system(exifcmd%(rsltText, outfile+'.pdf'))
            sys.exit(0)

        if pages2 == None:
            img1 = makeSingle(pages1, shapes1)
            outimg = genoverlay(toBin(img1,binthr), "target file '%s' cannot be loaded, test failed"%(Names[1]), Names[0], Names[1], "")
            rsltText="-:-:-:-:-:-:-:-:-:open"  #dummy result string 10 dashes necessary
            Image.fromarray(outimg).save(outfile+badpagetxt+'-p.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-p.pdf'))
            Image.fromarray(outimg).save(outfile+badpagetxt+'-l.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-l.pdf'))
            Image.fromarray(outimg).save(outfile+badpagetxt+'-z.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-z.pdf'))
            Image.fromarray(outimg).save(outfile+badpagetxt+'-s.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-s.pdf'))
            raise DoException("failed to open %s."%(Names[1]))


        if bisecting:
            badpagetxt="-bad"

        # create single image for each
        img1 = makeSingle(pages1, shapes1)
        img2 = makeSingle(pages2, shapes2)

        msg = ''
        if not (img1 != img1[0,0,0]).any():
            if not (img2 != img2[0,0,0]).any():
                msg= "Source and Target pdfs '%s' are empty, nothing to test"%(Names[0])
                rsltText = "-:0:-:0:-:0:-:0:-:0"  #dummy result string 10 dashes necessary
                outimg = genoverlay(toBin(img1,binthr), msg, Names[0], Names[1], "")
            else:
                msg= "Source pdf '%s' is empty, Target pdf not, nothing to test"%(Names[0])
                rsltText = "-:-:-:-:-:-:-:-:-:empty"  #dummy result string 10 dashes necessary
                outimg = genoverlay(toBin(img2,binthr), msg, Names[0], Names[1], "")
        elif not (img2 != img2[0,0,0]).any():
            msg = "Target pdf '%s' is empty, test failed"%(Names[1])
            rsltText = "-:-:-:-:-:-:-:-:-:empty"  #dummy result string 10 dashes necessary
            outimg = genoverlay(toBin(img1,binthr), msg, Names[0], Names[1], "")

        if msg:
            Image.fromarray(outimg).save(outfile+badpagetxt+'-p.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-p.pdf'))
            Image.fromarray(outimg).save(outfile+badpagetxt+'-l.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-l.pdf'))
            Image.fromarray(outimg).save(outfile+badpagetxt+'-z.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-z.pdf'))
            Image.fromarray(outimg).save(outfile+badpagetxt+'-s.pdf')
            os.system(exifcmd%(rsltText, outfile+badpagetxt+'-s.pdf'))
            raise DoException(msg)


        #crop to common size
        s1 = img1.shape
        s2 = img2.shape
        s=np.minimum(s1,s2)
        #img1 = img1[:s[0],:s[1]]
        #img2 = img2[:s[0],:s[1]]
        bimg1 = toBin(img1,binthr)
        bimg2 = toBin(img2,binthr)

        plainOvlRslt = getPagePixelOverlayIndex(bimg1, bimg2)
        lineVHOvlPage, lineVOvlPage, lineOvlDistRslt, lineOvlHPosRslt, pageHeightRslt, pageLinesRslt = lineIndexPage(bimg1, bimg2)

        le1 = 'FeatureDistanceError[mm]: %2.1f '%lineOvlDistRslt
        le2 = ': HorizLinePositionError[mm]: %2.2f '%lineOvlHPosRslt
        le3 = ': TextHeightError[mm]: %2.2f '%pageHeightRslt
        le4 = ': LineNumDifference: %2d'%pageLinesRslt
        grade = valToGrade((lineOvlDistRslt, lineOvlHPosRslt, pageHeightRslt, pageLinesRslt))

        if bisecting:
            if grade < badThr:
                outfile = outfile+"-good"
            else:
                outfile = outfile+"-bad"

        #pagerslt = pagerslt%((px2mm(heightErr), len(tx1)-len(tx0)))

        rsltText = plainOvlRslt + le1 + le2 + le3 + le4
        # command to write statistics to the pdf file, to be used in report creation

        #options: s, p, l z
        if overlayStyle == 'p' or overlayStyle == 'a':
            saveRslt('p',  'Page overlay, no alignment', bimg1, bimg2, Names[0], Names[1], plainOvlRslt+le3, rsltText, outfile)

        if overlayStyle == 'l' or overlayStyle == 'a':
            saveRslt('l', 'Page overlay, vertically aligned lines', lineVOvlPage, None, Names[0], Names[1], le2, rsltText, outfile)

        if overlayStyle == 'z' or overlayStyle == 'a':
            saveRslt('z', 'Page overlay, vertically and horizontally aligned lines', lineVHOvlPage, None, Names[0], Names[1], le1, rsltText, outfile)

            # side-by-side
        if overlayStyle == 's' or overlayStyle == 'a':
            saveRslt('s', '', img1, img2, Names[0], Names[1], le1, rsltText, outfile)

        # set exit status
        if grade < badThr:
            sys.exit(0)
        else:
            sys.exit(1)

    except DoException, e:
        print("\n" + e.what + " (" + Names[0] + ", " + Names[1] + ")")
        sys.exit(0)

if __name__=="__main__":
    mainfunc()
