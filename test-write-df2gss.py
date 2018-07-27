#!/usr/bin/env python

import os
import sys
import pandas
import writegss

key = '/path/to/service-account-key.json'
worksheet = writegss.WorksheetUtil('1_R1jkvv4WW7jomwMCyYJA-UCNVVFJIzdscFlF9xGXB4', sheet_index=0, keyfile=key)

sampledatafile = 'sampledata.csv'

sys.stderr.write('>>> module = {0}\n'.format(pandas.__name__))
sys.stderr.write('READ SAMPLE DATA: {0}\n'.format(os.path.abspath(sampledatafile)))

df = pandas.read_csv(sampledatafile)
sys.stderr.write('File Records: {0}\n'.format(len(df)))
header = list(df.columns)
style = [ { 'bold': True } ]*len(header)
worksheet.write_records_with_prepare(df, headers=header, fg=style)
