#!/bin/bash
#set -o xtrace #be verbose

. $FTPATH/officeconf.sh

line="$(date) - Running compareall.sh"
if [ ! -e ./log ]; then
    echo > ./log
else
    echo >> ./log
fi
echo $line >> ./log

dpi=300		#dpi to render pdfs
threshold=166	#threshold to identify foreground

function usage
{
	echo "$0: Compare pdf files and generate pair pdfs" 1>&2
	echo "Usage: $0 applist" 1>&2
	#not functional
	#echo "    -d int ............ dpi for rendering of pdf files {$dpi}" 1>&2
	#echo "    -t int ............ threshold to identify foreground {$threshold}" 1>&2
	echo "If applist is not specified, all applications specified in config.sh will be processed." 1>&2
}

function cmp ()
{
	#echo 1 $1
	refpdf=`basename $1` 	#source document with suffix
	refpdf=`basename $refpdf .pdf`	#source document without suffix
    ddd=`dirname $1`
    subdir=${ddd/\.\//}	# nice subdir path
    spdf=$sourceapp/$subdir/$refpdf	#source document with nice path
	for app in `echo $rtripapps`; do
        appfolder=$app'-'$(ver$app)
        for ofmt in $oformat; do
            tpdf=$appfolder/$subdir/$1
            tpdf="${tpdf/.pdf/.$ofmt.pdf}"
            docompare $spdf $tpdf $count
        done
        for ifmt in $iformat; do
            tpdf=$appfolder/$subdir/$1
            tpdf="${tpdf/.pdf/.$ifmt.$app.pdf}"
            docompare $spdf $tpdf $count
        done
    done
}

function docompare ()
{
    if [ ! -e "$1.pdf" ] || [ ! -e "$2-pair-l.pdf" ] || [ "$2-pair-l.pdf" -ot "$1" ];
    then
        echo $3 - Creating pairs for $1.pdf and $2
        time timeout 240s  docompare.py -t $threshold -d $dpi -a -o $2-pair $1.pdf $2
        echo

        if [ ! -e "$2-pair-l.pdf" ] || [ "$2-pair-l.pdf" -ot "$1" ];
        then
            rm /tmp/*.tif 2>/dev/null
        fi
    fi
}
# read the options
TEMP=`getopt -o h --long help -n 'test.sh' -- "$@"`
eval set -- "$TEMP"

# extract options and their arguments into variables.
while true ; do
    case "$1" in
        -h|--help)
		usage; exit 1;;
	-g)
		shift
		dpi=$1
		shift
		;;
	-t)
		shift
		threshold=$1
		shift
		;;
        --) shift ; break ;;
        *) echo "Internal error!" ; exit 1 ;;
    esac
done
shift $(expr $OPTIND - 1 )

# file names: bullets.docx.pdf


if [[ $# -gt 0 ]]
then
	while test $# -gt 0;
	do
		if [ -d "$1" ]; then
            echo $(date) - Processing $1 >> ./log
  			echo Processing $1
			#for pdfdoc in $pdfs; do cmp `basename $pdfdoc` $1; done
			count=0
			for pdfdoc in $pdfs; do ((count++)); cmp $pdfdoc $1 $count; done
		else
  			echo Directory $1 does not exist
		fi
  		shift
	done
else
    cd $sourceapp
    pdfs=`find . -name \*.pdf | grep -v pair | sort -n -k 1.7,1.9`
    cd ..
    count=0
    for pdfdoc in $pdfs; do ((count++)); cmp $pdfdoc $count; done
fi
