#!/bin/bash
#set -o xtrace #be verbose

. $FTPATH/officeconf.sh

line="$(date) - Running printsource.sh"
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
        for i in `find $sourcedir -name \*.$ifmt`; do
            (
            dir=`dirname $i`
            ofile=${i/.$ifmt/.pdf}
            ofile=${ofile/$dir/$sourceapp}
            auxoutput=${i/.$ifmt/.pdf}
            if [ ! -e "$ofile" ] || [ "$ofile" -ot "$ifile" ];
            then
                # keep type to enable processing of multiple formats
                echo Printing $i to $ofile
                print$sourceapp $i &>/dev/null
                if [ ! -e $auxoutput ];
                then
                    echo Failed to create $ofile
                    killWINEOFFICE
                    # delete in the case it is there from the previous test
                    # missing file will be in report indicated by grade 7
                else
                    mv $auxoutput $ofile
                fi
            fi
            )
        done
    done
else
    echo "$0 error: $sourceapp print command not defined. Is this the right system?" 2>&1
    exit 1
fi


