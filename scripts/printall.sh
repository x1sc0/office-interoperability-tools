#!/bin/bash
#set -o xtrace #be verbose

OfficeFormats="doc docx rtf ppt pptx"

. $FTPATH/officeconf.sh
checkLO

line="$(date) - Running printall.sh"
if [ ! -e ./log ]; then
    echo > ./log
else
    echo >> ./log
fi
echo $line >> ./log

# trap ctrl-c and call ctrl_c()
trap killWINEOFFICE INT

if ! xset q &>/dev/null; then
    echo "No X server. Simulating one..."
    export DISPLAY=:0.0
    Xvfb :0 -screen 0 1024x768x16 &
fi

odf=false
case "$1" in
    --odf)
		odf=true
        shift
        ;;
esac

let canprint=canprint$sourceapp
if [ "$odf" == false ] && [ $canprint -ne 1 ]
then
	echo "$0 error: $sourceapp print command not defined. Is this the right system?" 2>&1
	exit 1
fi

if [ ! -d "$sourceapp" ]; then
    mkdir $sourceapp
    rm -f `find $sourceapp -name \*.pdf`
    for fmt in $iformat; do
        rm -f `find $sourceapp -name \*.$fmt`
    done
fi

for a in $rtripapps; do
    echo $(date) - Printing in $a >> ./log
    echo Printing in $a
    apptype=`echo $a|cut -b -2`
    if [ $apptype == "LO" -o $apptype == "AO" -o $apptype == "OO" -o $apptype == "BB" ]
    then
        aVer=$a'-'$(ver$a)
    fi
    for ofmt in $oformat; do
        for i in `find $aVer -name \*.$ofmt`; do
            (
            pdffile=$i.export.pdf
            auxpdf=${i/.$ofmt/.pdf}
            if [ ! -e "$pdffile" ] || [ "$pdffile" -ot "$i" ];
            # files already have LO51 in their name, no renaming necessary
            #if [ ! -e "$ofile" ];
            then
                echo Printing $i to $pdffile

                if [ "$odf" == false ]; then
                    if [[ $OfficeFormats =~ $ofmt ]];
                    then
                        # apps in general cannot create specific file but just $auxpdf
                        print$sourceapp $i &>/dev/null
                        # convert to pdf
                        # input: orig/bullets.doc
                        # output: orig/bullets.doc.pdf
                        # output: LO52/bullets.doc.export.pdf
                        #rename to contain $ofmt in file name
                        if [ ! -e $auxpdf ];
                        then
                            echo Failed to create $pdffile
                            killWINEOFFICE
                            # delete in the case it is there from the previous test
                            # missing file will be in report indicated by grade 7
                            rm -f $pdffile
                        else
                            mv $auxpdf $pdffile
                        fi
                    else
                        if ! timeout 90s $FTPATH/scripts/doconv.sh -f pdf -a $a -i $i -o $pdffile; then
                            echo Timeout Reached
                        fi
                    fi
                else
					if [ ! -e "$pdffile" ] || [ "$pdffile" -ot "$i" ];
					then
						if ! timeout 60s $FTPATH/scripts/doconv.sh -f pdf -a $rtripapps -i $i -o $pdffile; then
							echo Timeout Reached
                            rm -r /tmp/lu*
						fi
					fi
                fi
            fi
            )
        done
    done
done
