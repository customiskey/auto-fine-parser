#!/bin/bash

# Copyright (C) 2023  Custom is Key B.V. - Allaert Euser
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.

source env/bin/activate
for filename in var/pdf/*;
    do
    	if [ -f "$filename" ]; then
		python parse.py -i ${filename} -o var/xlsx/output.xlsx
		mv ${filename} var/backup/
	fi
done
