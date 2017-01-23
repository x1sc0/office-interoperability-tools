#!/bin/bash
#set -o xtrace #be verbose

. $FTPATH/officeconf.sh 
checkLO


# trap ctrl-c and call ctrl_c()
trap killWINEOFFICE INT

function killWINEOFFICE() {
	if pgrep WINWORD.EXE > /dev/null; then
		echo "Killing WORD (WINWORD.EXE)"
		ps -ef | grep WINWORD.EXE | grep -v grep | awk '{print $2}' | xargs kill
	fi
	if pgrep POWERPNT.EXE > /dev/null; then
		sleep 5s
		#powerpoint might take some seconds to export to pdf. wait 5 seconds
		#before killing it
		if pgrep POWERPNT.EXE > /dev/null; then
			echo "Killing POWERPOINT (POWERPNT.EXE)"
			ps -ef | grep POWERPNT.EXE | grep -v grep | awk '{print $2}' | xargs kill
		fi
	fi
exit 1
}

if ! xset q &>/dev/null; then
    echo "No X server. Simulating one..."
    export DISPLAY=:0.0
    Xvfb :0 -screen 0 1024x768x16 &
fi

let canprint=canprint$sourceapp
if [ $canprint -eq 1 ] 
then

	if [ ! -d "$sourceapp" ]; then
		mkdir $sourceapp
		rm -f `find $sourceapp -name \*.pdf`
		for fmt in $iformat; do
			rm -f `find $sourceapp -name \*.$fmt`
		done
	fi


	for ifmt in $iformat; do
		for ofmt in $oformat; do
			if [ $ofmt != $ifmt ]; then
				printfmt=$ofmt
				ext=$ofmt
			else
				printfmt=pdf
				ext=$ifmt.$printfmt
			fi
			for i in `find $sourcedir -name \*.$ifmt`; do
				(
				dir=`dirname $i`
				ofile=${i/.$ifmt/.$ext}
				ofile=${ofile/$dir/$sourceapp}
				auxoutput=${i/.$ifmt/.$printfmt}
				if [ ! -e "$ofile" ] || [ "$ofile" -ot "$ifile" ];
				then
					# keep type to enable processing of multiple formats

					echo Printing $i to $ofile
					print$sourceapp $printfmt $i &>/dev/null
					if [ ! -e $auxoutput ];
					then
						killWINEOFFICE
						# delete in the case it is there from the previous test
						# missing file will be in report indicated by grade 7
						echo Failed to create $auxoutput
					else
						mv $auxoutput $ofile
					fi
				fi
				)
			done
		done
	done

	for a in $sourceapp $rtripapps; do
		echo Printing in $a
        apptype=`echo $a|cut -b -2`
	    if [ $apptype == "LO" -o $apptype == "AO" -o $apptype == "OO" -o $apptype == "BB" ]
	    then
            a=$a'-'$(ver$a)
        fi
		for ofmt in $oformat; do
			for i in `find $a -name \*.$ofmt`; do
				(
				pdffile=$i.pdf
				auxpdf=${i/.$ofmt/.pdf}
				if [ ! -e "$pdffile" ] || [ "$pdffile" -ot "$i" ];
				# files already have LO51 in their name, no renaming necessary
				#if [ ! -e "$ofile" ];
				then
					echo Printing $i to $pdffile
					# apps in general cannot create specific file but just $auxpdf
					print$sourceapp pdf $i &>/dev/null
					# convert to pdf
					# input: orig/bullets.doc
					# output: orig/bullets.doc.pdf
					# output: LO52/bullets.doc.pdf
					#rename to contain $ofmt in file name
					if [ ! -e $auxpdf ];
					then
						killWINEOFFICE
						# delete in the case it is there from the previous test
						# missing file will be in report indicated by grade 7
						echo Failed to create $pdffile
						rm -f $pdffile
					else
						mv $auxpdf $pdffile
					fi
				#else
					#echo "$ofile is up to date"
				fi
				)
			done
		done
	done

else
	echo "$0 error: $sourceapp print command not defined. Is this the right system?" 2>&1
	exit 1
fi
