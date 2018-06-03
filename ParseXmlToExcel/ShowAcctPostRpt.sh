#!/bin/bash
#Created:20180110
#Description:Running the tool that parse AcctPostingReport.xml to Excel file.
#Author:shenlj

. /summit/set_prod.sh

echo "******************* `date +'%Y%m%d %H:%M:%S'` start excute shell $0 **********************"

#check command exit error
checkerror()
{
	if [ $? != 0 ];then
		call_from=$(caller 0)
		echo "Executing command failed,check error at [$call_from]"
		exit 1;
	fi
}

#get yesterday date 
yestd=`date +"%m/%d/%Y" -d "-1 days"`
yestd_fmt=`date +"%Y%m%d" -d "-1 days"`
#yestd=11/30/2017
#yestd_fmt=20171130

mkdir -p /sumdata/report/AcctPostingReport/$yestd_fmt
checkerror
acctpost -PDATE $yestd -O /sumdata/report/AcctPostingReport/$yestd_fmt/AcctPostingReport.xml -XML -EXTERNAL -SEC -BOOK
checkerror

impf_default=/sumdata/report/AcctPostingReport/$yestd_fmt/AcctPostingReport.xml
expf_default=/sumdata/report/AcctPostingReport/$yestd_fmt/AcctPostingReport_$yestd_fmt.xlsx

#read -p "Please input Xml file:" import_file
#if [ ! $import_file ];then
#	import_file=$impf_default
#fi


#read -p "Please input Excel file:" export_file
#if [ ! $export_file ];then
#	export_file=$expf_default
#fi


base_path=$(cd `dirname $0` ; pwd)
echo "Excuting Infomation [import Xml file path:$impf_default"
echo "Export Excel file path:$expf_default"
echo "The tool jar path:$base_path]"

echo "Start running the tool ..."
echo "*--------------The Java Console Infomation------------------------------*"
cd $base_path
java -jar AccountPostingReports.jar $impf_default $expf_default
checkerror
echo "*-----------------------------------------------------------------------*"

#Transmit the Excel file to destination
ADDR=10.200.184.86
FTP_USR=fuser
FTP_PWD=`echo U2FsdGVkX18SDLgU3pVtNqSc2lLYu7dQmkiQSXbLCME= | openssl aes-128-cbc -d -k 123 -base64`
#-------test--------
#ADDR=10.112.19.76
#FTP_USR=ima
#FTP_PWD=`echo U2FsdGVkX1/uyfVXYy/eXCqwMU7V/G3jM+/szgLVqeE= | openssl aes-128-cbc -d -k 123 -base64`
SOURCEPATH=/sumdata/report/AcctPostingReport/$yestd_fmt

echo "Start Transmit the Excel file..."
if [ ! -d $SOURCEPATH ]
	then 
		echo "Source file $SOURCEPATH doesn't exist!" 
		exit 1
	else
		cd $SOURCEPATH
fi
DESTPATH=AcctPosting/
ftp -n << EOF
	open $ADDR
	user $FTP_USR $FTP_PWD
	bin	
	prompt
	cd ${DESTPATH}
	mput *.xlsx
	by
EOF
checkerror

echo "******************* `date +'%Y%m%d %H:%M:%S'` end excute shell $0 **********************"
