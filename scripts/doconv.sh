#!/bin/bash
#set -o xtrace #be verbose

# To use:
# In the 'orig' directory run
# for i in *.docx; do doconv.sh -a LO43 -i $i -o ../LO43/$i; done

function usage
{
	echo "$0 convert office document to different format" 1>&2
	echo "Usage: $0 [switches] " 1>&2
	echo "Switches:" 1>&2
	echo "    -h --help ......... this usage" 1>&2
	echo "    -a ................ application to use (LOxx, AOxx, MSxx,....)" 1>&2
	echo "    -f ................ output format extension (odt, docx, ...)" 1>&2
	echo "    -i ................ input file" 1>&2
	echo "    -o ................ output file" 1>&2
	exit 1
}


# read configuration
. $FTPATH/officeconf.sh
rtripapp=""
ifile=xx.odt
ofile=xx.pdf
oformat=odt
failed=$FTPATH/failed

while [ $# -gt 0 ]
do
	case "$1" in
		-h* | --help*)
			usage
			shift
			;;
		-a)
			shift
			rtripapp=$1
			shift
			;;
		-i)
			shift
			ifile=$1
			shift
			;;
		-o)
			shift
			ofile=$1
			shift
			;;
		-f)
			shift
			oformat=$1
			shift
			;;
		*)
			break
			;;
	esac
done

fromtype=`echo $ifile|awk -F "." '{print $NF}'`
totype=`echo $ofile|awk -F "." '{print $NF}'`
apptype=`echo $rtripapp|cut -b -2`

if [ -e .~lock.auxttt.xx.$totype# ]
then
    rm .~lock.auxttt.xx.$totype#
fi

let canconvert=canconvert$rtripapp
if [ $canconvert -eq 1 ]
then
	if [ $apptype == "LO" -o $apptype == "AO" -o $apptype == "OO" -o $apptype == "BB" ]
	then
        #Remove the user profile
        rm -Rf /tmp/test/user

		echo "$0 Converting $ifile to $ofile using $rtripapp" 2>&1
		cp $ifile auxttt.$fromtype
		conv$rtripapp xx.$totype auxttt.$fromtype
		if [ -e auxttt.xx.$totype ]
		then
			mv auxttt.xx.$totype $ofile
			rm -f auxttt.$fromtype
		else
			cp $failed.$totype $ofile
			echo $0: $ifile failed to open by conv$rtripapp
			exit 1
		fi
	elif [ $apptype == "AW" -o $apptype == "CW" ]
	then
		echo "$0 Converting $ifile to $ofile using $rtripapp" 2>&1
		cp $ifile auxttt.$fromtype
		conv$rtripapp $totype auxttt.$fromtype
		if [ -e auxttt.$totype ]
		then
			mv auxttt.$totype $ofile
			rm -f auxttt.$fromtype
		else
			cp $failed.$totype $ofile
			echo $0: $ifile failed to open by conv$rtripapp
			exit 1
		fi
	#elif [ $apptype == "AO" ]
	#then
		#echo "$0 Converting $ifile to $ofile using $rtripapp" 2>&1
		#cp $ifile auxttt.$fromtype
		##echo conv$rtripapp xx.$totype auxttt.$fromtype
		#conv$rtripapp xx.$totype auxttt.$fromtype
		#mv auxttt.xx.$totype $ofile
		#rm auxttt.$fromtype
	elif [ $apptype == "MS" ]
	then
		echo "$0: Converting $ifile to $ofile using $rtripapp" 2>&1
		case $rtripapp in
			MS07)	msapp=$MS07PROG;;
			MS10)	msapp=$MS10PROG;;
			MS13)	msapp=$MS13PROG;;
			MSWINE)	msapp="$WINEPROG OfficeConvert" ;;
		esac
		$msapp --format=$oformat --output=$ofile $ifile &>/dev/null
	elif [ $apptype == "GD" ]
	then
		echo "$0: Converting $ifile to $ofile using $rtripapp 2>&1"
		cp $ifile auxttt.$fromtype
		$GDCONVERT $totype auxttt.$fromtype
		if [ -e auxttt.$totype ]
		then
			mv auxttt.$totype $ofile
			rm -f auxttt.$fromtype
		else
			cp $failed.$totype $ofile
			echo $0: $ifile failed to open by $GDCONVERT
		fi
	else
		echo "$0: Unsupported application $rtripapp" 2>&1
	fi
else
	echo "$0 error: $rtripapp conversion executable not defined. Is this the right system?" 2>&1
	exit 1
fi




