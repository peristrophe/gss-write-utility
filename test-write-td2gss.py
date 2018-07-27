#!/usr/bin/env python

import os
import sys
import tdclient
import writegss

query = """
SELECT {0} FROM your_table
""".strip()

key = '/path/to/service-account-key.json'
worksheet = writegss.WorksheetUtil('1_R1jkvv4WW7jomwMCyYJA-UCNVVFJIzdscFlF9xGXB4', sheet_index=0, keyfile=key)

header = ['id', 'first_name', 'middle_name', 'last_name', 'birthday', 'sex', 'address', 'zipcode']
style = [ { 'bold': True } ]*len(header)

sys.stderr.write('>>> module = {0}\n'.format(tdclient.__name__))
sys.stderr.write('{0} THROWING QUERY BELOW {1}\n{2}\n{3}\n'.format('#'*10, '#'*20, query.format(', '.join(header)), '#'*52))

with tdclient.Client() as td:
    job = td.query('your_database', query.format(', '.join(header)), type='presto')
    job.wait()
    sys.stderr.write('Result Records: {0}\n'.format(job.num_records))
    worksheet.write_records_with_prepare(job, headers=header, fg=style)
