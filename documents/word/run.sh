source ./config.sh

#set input type here
export iformat="doc docx rtf"

#set output type here
export oformat="doc docx rtf odt"

export sourcedir="orig"	#copy test files here (docx in this case, may have subdirectories)

# tested applications
# applications are defined in officeconfig.sh
export rtripapps="LOMASTER"

# reference application to be used for printing
export sourceapp="MSWINE"	# MS Office running under Wine on Linux

../../scripts/convall.sh
../../scripts/printsource.sh
../../scripts/printall.sh
../../scripts/compareall.sh
../../scripts/gencsv.sh
