from collections.abc import MutableSequence
from functools import singledispatch
from itertools import islice

from writegss.abstracts import *

@singledispatch
def paging_records(obj, size):
    pass


@paging_records.register(tuple)
@paging_records.register(MutableSequence)
def _(seq, size):
    picked = 0
    while picked < len(seq):
        try:
            records = list(islice(seq, picked, picked+size))
        except IndexError:
            if len(records) < 1: raise StopIteration

        picked += len(records)
        yield records


@paging_records.register(DatabaseApiCursor)
def _(cursor, size):
    while True:
        records = cursor.fetchmany(size)
        if len(records) < 1: raise StopIteration
        yield records


@paging_records.register(TdClientJob)
def _(job, size):
    picked = 0
    while picked < job.num_records:
        try:
            records = list(islice(job.result(), picked, picked+size))
        except ValueError:
            raise StopIteration
        except:
            if len(records) < 1: raise StopIteration

        picked += len(records)
        yield records


@paging_records.register(PandasDataframe)
def _(df, size):
    df = df.fillna('')
    picked = 0
    while picked < len(df):
        try:
            records = list(islice(df.itertuples(index=False), picked, picked+size))
        except:
            if len(records) < 1: raise StopIteration

        picked += len(records)
        yield records


@paging_records.register(CSVReader)
def _(reader, size):
    picked = 0
    while True:
        records = list(islice(reader, size))
        if len(records) < 1: raise StopIteration
        records = list(map(lambda rec: list(map(lambda x: x if x != 'null' else None, rec)), records))

        picked += len(records)
        yield records
