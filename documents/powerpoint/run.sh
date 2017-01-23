export LOMASTERPROG="/home/xisco/libreoffice/instdir/program/soffice"
export LO52PROG="/home/xisco/bibisect/bibisect-linux-64-5.3/instdir/program/soffice"
export FTPATH="/home/xisco/office-interoperability-tools"
export GDCONVERT="$FTPATH/gdconvert/gdconvert"
export WINEPROG="/usr/bin/wine"
export WINEPREFIX="/home/xisco/.wineprefixes/msoffice2010/"

#set input type here
export iformat="pptx"

#set output type here
export oformat="pptx"

export sourcedir="orig"	#copy test files here (docx in this case, may have subdirectories)

# tested applications 
# applications are defined in officeconfig.sh
export rtripapps="LOMASTER LO52"

# reference application to be used for printing
export sourceapp="MSWINE"	# MS Office running under Wine on Linux

../../scripts/convall.sh
../../scripts/printall.sh
../../scripts/compareall.sh
../../scripts/gencsv.sh
../../scripts/genods.py -i all.csv -o rslt.ods
