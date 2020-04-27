source ./config.sh

#set input type here
export iformat="ods"

#set output type here
export oformat="ods xls xlsx"

export sourcedir="orig"	#copy test files here (odt in this case, may have subdirectories)

# tested applications
# applications are defined in officeconfig.sh
export rtripapps="LOMASTER"

# reference application to be used for printing
export sourceapp="old"

../../scripts/convall.sh
../../scripts/printall.sh --odf
../../scripts/compareall.sh --odf
../../scripts/gencsv.sh