#!/bin/bash
source env/bin/activate
for filename in var/pdf/*;
    do
    	if [ -f "$filename" ]; then
		python parse.py -i ${filename} -o var/xlsx/output.xlsx
		mv ${filename} var/backup/
	fi
done
