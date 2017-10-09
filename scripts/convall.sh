#!/bin/bash
#set -o xtrace #be verbose
. $FTPATH/officeconf.sh

checkLO

line="$(date) - Running convall.sh"
if [ ! -e ./log ]; then
    echo > ./log
else
    echo >> ./log
fi
echo $line >> ./log


for rtapp in `echo $rtripapps`; do
	startOOoServer $rtapp
	let canconvert=canconvert$rtapp
	if [ $canconvert -eq 1 ]
	then
        echo $(date) - Processing $rtapp >> ./log
		echo Processing $rtapp
        folder=$rtapp'-'$(ver$rtapp)
		#create directory if it does not exit
		if [ ! -d $folder ]; then
			mkdir $folder
			rm -f `find $folder -name \*.pdf`
			for fmt in $iformat; do
				rm -f `find $folder -name \*.$fmt`
			done
		fi
		for ifmt in $iformat; do
			files=`find $sourcedir -name \*.$ifmt`
			#echo xx $files
			for ifile in $files; do
				for ofmt in $oformat; do
					ofile=${ifile/.$ifmt/.$ofmt}
					ofile=${ofile/$sourcedir/$folder}
					# convert to the same file type
					# input: orig/bullets.doc
					# output: LO52/bullets.doc
					if [ ! -e "$ofile" ] || [ "$ofile" -ot "$ifile" ];
					then
						if ! timeout 60s $FTPATH/scripts/doconv.sh -f $ofmt -a $rtapp -i $ifile -o $ofile; then
							echo Timeout Reached
						fi
					#else
						#echo "$ofile is up to date"
					fi

					# keep type to enable processing of multiple formats
					ofile2=`dirname $ofile`/`basename $ofile`.$rtapp.pdf

					# convert to pdf
					# input: orig/bullets.doc
					# output: LO52/bullets.doc.LO52.pdf

                    if [ $ofmt == $ifmt ]; then
					    if [ ! -e "$ofile2" ] || [ "$ofile2" -ot "$ifile" ]; then
						    if ! timeout 90s $FTPATH/scripts/doconv.sh -f pdf -a $rtapp -i $ifile -o $ofile2; then
							    echo Timeout Reached
						    fi
					    #else
						    #echo "$ofile is up to date"
					    fi
                    fi
				done
			done
		done
	else
		echo "$0 warning: $rtapp conversion executable not defined. Is this the right system?" 2>&1
	fi
	killOOoServer
done
