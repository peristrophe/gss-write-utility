#!/usr/bin/env python

import os
import sys
import csv
import writegss

key = '/path/to/service-account-key.json'
worksheet = writegss.WorksheetUtil('1_R1jkvv4WW7jomwMCyYJA-UCNVVFJIzdscFlF9xGXB4', sheet_index=0, keyfile=key)

sampledatafile = 'sampledata.csv'

sys.stderr.write('>>> module = {0}\n'.format(csv.__name__))
sys.stderr.write('READ SAMPLE DATA: {0}\n'.format(os.path.abspath(sampledatafile)))
    
with open(sampledatafile, 'r') as f:
    reader = csv.reader(f)
    header = next(reader)
    style = [ { 'bold': True } ]*len(header)
    
    #sys.stderr.write('File Records: {0}\n'.format(len(list(reader))))
    worksheet.write_records_with_prepare(reader, headers=header, fg=style)
