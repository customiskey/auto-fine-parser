#!/usr/bin/env python3

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

# Default imports
import getopt
from sys import argv, exit

# Specific imports for project
from PyPDF2 import PdfReader
import re
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import pandas as pd
import os
from datetime import datetime

debugging = 0

def parse_arguments(argv):
   global debugging
   inputfile = ''
   outputfile = ''
   try:
      opts, args = getopt.getopt(argv, "hdi:o:",["ifile=","ofile="])
   except getopt.GetoptError:
      print_help()
      exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print_help()
         exit()
      elif opt == '-d':
         debugging = 1
         print_debug("debugging is ON")
      elif opt in ("-i", "--ifile"):
         inputfile = arg
      elif opt in ("-o", "--ofile"):
         outputfile = arg
   if inputfile == '':
      print_help()
      exit(2)
   return inputfile, outputfile

def print_help():
   print('parse.py -i <inputfile> -o <outputfile>')

def print_debug(output):
   global debugging
   if debugging != 0:
      print(output)

def extract_txt_from_pdf(pdffile):
   txt = ''
   reader = PdfReader(pdffile)
   for i in range(len(reader.pages)):
      page = reader.pages[i]
      txt += page.extract_text()
   return txt

def normalize_bedrag(txt):
   try:
      txt = txt.split()[1]
      txt = txt.split(',')[0] + '.' + txt.split(',')[1]
   except:
      txt = '0'
   x = float(txt)
   return x

def normalize_cjibnr(txt):
   try:
      txt = txt.split('\n')[1]
   except:
      txt = 'FAULT'
   return txt

def normalize_datum(txt):
   pattern = r'[0-9]'
   abcpattern = r'[a-zA-Z]'
   try:
      txt = txt.split('\n')[1]

      # write regex to recognize numbers in front of month
      # change to yyyy-mm-dd
   except:
      txt = 'FAULT'
   
   months = {
      'januari': 1,
      'februari': 2,
      'maart': 3,
      'april': 4,
      'mei': 5,
      'juni': 6,
      'juli': 7,
      'augustus': 8,
      'september': 9,
      'oktober': 10,
      'november': 11,
      'december': 12
   }
   try:
      dd = ''
      for letter in txt[0:2]:
         if not re.search(abcpattern, letter):
            dd += letter
      mm = months[re.sub(pattern, '', txt).strip()]
      yy = txt.split()[1]
      print_debug(dd)
      print_debug(mm)
      print_debug(yy)
      dt_object = datetime.strptime(str(dd)+'/'+str(mm)+'/'+str(yy), '%d/%m/%Y')

   except:
      dt_object = datetime.strptime('01/01/1970', '%d/%m/%Y')
   return dt_object

def normalize_feitcode(txt):
   try:
      txt = txt.strip().split(')')[0]
   except:
      txt = 'FAULT'
   return txt

def normalize_kenteken(txt):
   try:
      txt = txt.split('\n')[1]
   except:
      txt = 'FAULT'
   return txt

def normalize_plaats(txt):
   try:
      txt = txt.split('\n(')[1].split(' )')[0]
      txt = txt.replace('\n', ' ')         
   except:
      txt = 'FAULT'
   return txt

def normalize_omschrijving(txt):
   try:
      txt = txt.replace('\n', '')
      txt = txt[:-1].strip()
   except:
      txt = 'FAULT'
   return txt

def normalize(id,txt):
   result = ''
   switch = {
      'cjibnr': normalize_cjibnr(txt),
      'Datum Bekeuring': normalize_datum(txt),
      'Feitcode overtreding': normalize_feitcode(txt),
      'Hoogte bedrag': normalize_bedrag(txt),
      'Kenteken': normalize_kenteken(txt),
      'Locatie bekeuring': normalize_plaats(txt),
      'Omschrijving overtreding': normalize_omschrijving(txt)
   }
   return switch[id]

def extract_values_from_txt(txt_string):
   # identifiers voor en na text
   # example:
   #  Datum beschikking     
   #  31januari 2023
   #  verkeersboete

   identifiers = {
      'cjibnr': ['-nummer','\nDatum'],
      'Datum Bekeuring':['Wanneer',' om'],
      'Feitcode overtreding':['feitcode', '\nWanneer'],
      'Hoogte bedrag':['Door u te betalen', '\nDit bedrag'],
      'Kenteken':['Kenteken','Dit'],
      'Locatie bekeuring':['Waar','\nKenteken'],
      'Omschrijving overtreding': ['Omschrijving overtreding', 'feitcode']
      }
   
   fine_details_dict = {}
   for id in identifiers:
      word_start = identifiers[id][0]
      word_end = identifiers[id][1]
      searchregex = re.compile(word_start + '(.*?)' + word_end, re.DOTALL)
      patternmatch = searchregex.search(txt_string)
      if patternmatch:
         patternresult = patternmatch.group(1)
         fine_details_dict[id] = normalize(id,patternresult)
      else:
         fine_details_dict[id] = ''
   fine_details_dict['Datum Invoer'] = datetime.today()
   return fine_details_dict

def make_clickable(val):
   return f'<a target="_blank" href="{val}">{val}</a>'

def append_xlsx(fine_details_dict, xlsxfile):
   df = pd.DataFrame(fine_details_dict)
   df_excel = pd.read_excel(xlsxfile)
   result = df_excel.append(df)
   #for uri in result['URI'].values:
   #   result['URI'] = make_clickable(result['URI'])
   result.to_excel(xlsxfile, index=False)


def append_data_to_excel(excel_name, sheet_name, data):
    with pd.ExcelWriter('xlsx/outputnew.xlsx') as writer:
        columns = []
        for k, v in data.items():
            columns.append(k)
        df = pd.DataFrame(data, index= None)
        df_source = None
        if os.path.exists(excel_name):
            df_source = pd.DataFrame(pd.read_excel(excel_name, sheet_name=sheet_name, engine='openpyxl'))
        if df_source is not None:
            df_dest = df_source.append(df)
        else:
            df_dest = df
        df_dest.to_excel(writer, sheet_name=sheet_name, index = False, columns=columns)

def main():
   data = {}
   inputfile, outputfile = parse_arguments(argv[1:])
   txt_string = extract_txt_from_pdf(inputfile)
   print_debug(txt_string)
   fine_details_dict = extract_values_from_txt(txt_string)
   fine_details_dict['URI'] = "\\\\fs01\\office\\data_exchange\\cjib\\backup\\" + inputfile.split('/')[2]
   for column in fine_details_dict:
      data[column] = [fine_details_dict[column]]
   
   append_xlsx(data, outputfile)
   
   print_debug(data)
   return

if __name__ == "__main__":
   main()


