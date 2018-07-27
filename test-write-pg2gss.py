#!/usr/bin/env python

import os
import sys
import psycopg2 as pg
import writegss

dsn = 'postgresql://{0}:{1}@{2}:{3}/{4}'.format(
    os.getenv('PGUSER'), os.getenv('PGPASSWORD'),
    os.getenv('PGHOST'), os.getenv('PGPORT'), os.getenv('PGDATABASE')
)

query = """
SELECT {0} FROM your_table
""".strip()

key = '/path/to/service-account-key.json'
worksheet = writegss.WorksheetUtil('1_R1jkvv4WW7jomwMCyYJA-UCNVVFJIzdscFlF9xGXB4', sheet_index=0, keyfile=key)

header = ['id', 'first_name', 'middle_name', 'last_name', 'birthday', 'sex', 'address', 'zipcode']
style = [ { 'bold': True } ]*len(header)

sys.stderr.write('>>> module = {0}: API Level = {1}\n'.format(pg.__name__, pg.apilevel))
sys.stderr.write('{0} THROWING QUERY BELOW {1}\n{2}\n{3}\n'.format('#'*10, '#'*20, query.format(', '.join(header)), '#'*52))

with pg.connect(dsn) as conn:
    with conn.cursor('SelectQuery') as cur:
        cur.execute(query.format(', '.join(header)))
        worksheet.write_records_with_prepare(cur, headers=header, fg=style)
