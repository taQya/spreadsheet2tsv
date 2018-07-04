#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from xml.etree.ElementTree import *

SS = '{urn:schemas-microsoft-com:office:spreadsheet}'
INPUT_PATH = './'
OUTPUT_PATH = './'
TAB = '\t'
ENTER = '\n'

def read(filepath, sheetname):
	tree = parse(filepath)
	sheetdata = []
	for node in tree.iter(SS + "Worksheet"):
		if node.attrib.get(SS + "Name") == sheetname:
			for row in node.iter(SS + "Row"):
				single_row = []
				for r in row:
					data = r.find(SS + "Data")
					if data is None or data.text is None:
						single_row.append("")
					else:
						single_row.append(data.text)

				sheetdata.append(single_row)

	return sheetdata

def write_tsv(file_path, src):
	with open(file_path, "w") as f:
		for row in src:
			f.write(TAB.join(row).encode("utf-8"))
			f.write(ENTER.encode("utf-8"))

def main():
	args = sys.argv
	if len(args) != 3:
		print str('Invalid parameter.')
		print str('python spreadsheet2tsv.py ${SPREADSHEET_NAME} ${SHEETNAME}')
		sys.exit()

	xmlfile = args[1]
	sheetname = args[2]
	sheetdata = read(INPUT_PATH + xmlfile, sheetname)

	output_filename = xmlfile.replace(".xml", "") + "_" + sheetname + ".tsv"

	write_tsv(OUTPUT_PATH + output_filename, sheetdata)

if __name__ == '__main__':
    main()