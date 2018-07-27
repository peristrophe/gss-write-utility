from abc import ABC, abstractmethod

### PEP 249: Python Database API Specification
### https://www.python.org/dev/peps/pep-0249/
### e.g.) psycopg2, MySQLdb, ...
class DatabaseApiCursor(ABC):

    @abstractmethod
    def fetchmany(self):
        pass

    @classmethod
    def __subclasshook__(cls, C):
        if any('fetchmany' in B.__dict__ for B in C.__mro__):
            return True
        else:
            return NotImplemented


class TdClientJob(ABC):

    @abstractmethod
    def fetchmany(self):
        pass

    @classmethod
    def __subclasshook__(cls, C):
        if any(repr(B) == "<class 'tdclient.job_model.Job'>" for B in C.__mro__):
            return True
        else:
            return NotImplemented


class PandasDataframe(ABC):

    @abstractmethod
    def fetchmany(self):
        pass

    @classmethod
    def __subclasshook__(cls, C):
        if str(C) == "<class 'pandas.core.frame.DataFrame'>":
            return True
        else:
            return NotImplemented


class CSVReader(ABC):

    @abstractmethod
    def fetchmany(self):
        pass

    @classmethod
    def __subclasshook__(cls, C):
        if str(C) == "<class '_csv.reader'>":
            return True
        else:
            return NotImplemented
