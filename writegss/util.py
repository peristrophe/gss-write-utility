import os
import re
import datetime
from math import floor
from copy import deepcopy

import httplib2
from apiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials 

from writegss.dispatcher import *

### perhaps a namedtuple is better ...?
COLOR_PALLET = {
    'WHITE':          { 'red': 1.000,   'green': 1.000,   'blue': 1.000 },
    'BLACK':          { 'red': 0.000,   'green': 0.000,   'blue': 0.000 },
    'ORIGIN-RED':     { 'red': 1.000,   'green': 0.000,   'blue': 0.000 },
    'ORIGIN-GREEN':   { 'red': 0.000,   'green': 1.000,   'blue': 0.000 },
    'ORIGIN-BLUE':    { 'red': 0.000,   'green': 0.000,   'blue': 1.000 },
    'ORIGIN-CYAN':    { 'red': 0.000,   'green': 1.000,   'blue': 1.000 },
    'ORIGIN-MAGENTA': { 'red': 1.000,   'green': 0.000,   'blue': 1.000 },
    'ORIGIN-YELLOW':  { 'red': 0.000,   'green': 1.000,   'blue': 1.000 },
    'RED':            { 'red': 0.800,   'green': 0.200,   'blue': 0.200 },
    'GREEN':          { 'red': 0.200,   'green': 0.800,   'blue': 0.200 },
    'BLUE':           { 'red': 0.200,   'green': 0.200,   'blue': 0.800 },
    'CYAN':           { 'red': 0.200,   'green': 0.800,   'blue': 0.800 },
    'MAGENTA':        { 'red': 0.800,   'green': 0.200,   'blue': 0.800 },
    'YELLOW':         { 'red': 0.800,   'green': 0.800,   'blue': 0.200 },
    'GRAY':           { 'red': 0.500,   'green': 0.500,   'blue': 0.500 },
    'LIGHT-RED':      { 'red': 0.956,   'green': 0.800,   'blue': 0.800 },
    'LIGHT-GREEN':    { 'red': 0.850,   'green': 0.917,   'blue': 0.827 },
    'LIGHT-BLUE':     { 'red': 0.811,   'green': 0.886,   'blue': 0.952 },
    'LIGHT-CYAN':     { 'red': 0.815,   'green': 0.878,   'blue': 0.890 },
    'LIGHT-MAGENTA':  { 'red': 0.917,   'green': 0.819,   'blue': 0.862 },
    'LIGHT-YELLOW':   { 'red': 1.000,   'green': 0.949,   'blue': 0.800 },
    'LIGHT-GRAY':     { 'red': 0.700,   'green': 0.700,   'blue': 0.700 },
    'DARK-RED':       { 'red': 0.600,   'green': 0.000,   'blue': 0.000 },
    'DARK-GREEN':     { 'red': 0.219,   'green': 0.462,   'blue': 0.113 },
    'DARK-BLUE':      { 'red': 0.043,   'green': 0.325,   'blue': 0.580 },
    'DARK-CYAN':      { 'red': 0.074,   'green': 0.309,   'blue': 0.360 },
    'DARK-MAGENTA':   { 'red': 0.454,   'green': 0.105,   'blue': 0.278 },
    'DARK-YELLOW':    { 'red': 0.749,   'green': 0.564,   'blue': 0.000 },
    'DARK-GRAY':      { 'red': 0.300,   'green': 0.300,   'blue': 0.300 }
}

CP = COLOR_PALLET ### Aliasing short symbol for accessibility


class CellAddressUtil:

    LABELD = re.compile('\d+')
    LABELA = re.compile('[A-Z]+')

    def __init__(self, address='A1'):
        self.address = address

    def __repr__(self):
        return self.address

    @property
    def row(self):
        return self._row

    @row.setter
    def row(self, value):
        if value < 1:
            raise AttributeError('Attempted to set invalid cell address: row = {0}'.format(value))

        self._row = int(value)

    @row.deleter
    def row(self):
        self._row = 1

    @property
    def column(self):
        return self._column

    @column.setter
    def column(self, value):
        if value < 1:
            raise AttributeError('Attempted to set invalid cell address: column = {0}'.format(value))

        self._column = int(value)

    @column.deleter
    def column(self):
        self._column = 1

    @property
    def address(self):
        return '{0}{1}'.format(self.to_column_name(self.column), self.row)

    @address.setter
    def address(self, value: str):
        if re.match('^[A-Z]+\d+$', value) is None:
            raise AttributeError('Attempted to set invalid format of cell address: address = {0}'.format(value))

        self.row = int(self.LABELD.search(value)[0])
        self.column = self.to_column_number(self.LABELA.search(value)[0])

    @address.deleter
    def address(self):
        self.row = 1
        self.column = 1

    @property
    def grid(self):
        return (self.row, self.column)

    @grid.setter
    def grid(self, values):
        if len(values) < 2:
            raise AttributeError('Invalid size of values: {0}'.format(values))

        self.row = int(values[0])
        self.column = int(values[1])

    @grid.deleter
    def grid(self):
        self.row = 1
        self.column = 1

    @staticmethod
    def to_column_name(column_number):
        if not isinstance(column_number, int):
            raise TypeError('Invalid type argument. Column number must be Integer.')

        if column_number <= 0:
            raise ValueError('Invalid argument. Column number must be over than zero.')

        syms = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        shift_dim = lambda x, base, e=0: shift_dim(x-base**e, base, e+1) if x-base**e >= 0 else [x, e]
        rebase = lambda x, base: rebase(floor(x/base), base) + [x%base] if floor(x/base) != 0 else [x]

        index, dim = shift_dim(column_number, len(syms))
        digits = ([0]*dim + rebase(index, len(syms)))[-dim:]

        return ''.join(map(lambda x: syms[x], digits))

    @staticmethod
    def to_column_number(column_name):
        if not isinstance(column_name, str):
            raise TypeError('Invalid type argument. Column name must be String.')

        if re.match('^[A-Z]+$', column_name) is None:
            raise ValueError('Invalid argument. Column name must be composed of uppercase alphabet only.')

        syms = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

        offset = sum(map(lambda x: len(syms)**x, range(len(column_name))))
        digits = list(map(lambda x: syms.index(x), column_name))
        decimal = sum(map(lambda d: d[1]*(len(syms)**d[0]), enumerate(reversed(digits))))

        return decimal + offset

    def to_a1range(self, other):
        if not isinstance(other, __class__):
            raise ValueError('Invalid argument. Must be instance of {0} class.'.format(__class__.__name__))

        top_left = __class__()
        bottom_right = __class__()
        top_left.grid = (min(self.row, other.row), min(self.column, other.column))
        bottom_right.grid = (max(self.row, other.row), max(self.column, other.column))
        return '{0}:{1}'.format(top_left.address, bottom_right.address)

    ### Suppose use case when you want to relative cell address with keep current address.
    ### e.g.) destination = current_cell.shift(row=3, col=1).address
    def shift(self, *, row=0, col=0):
        dup = __class__(self.address)
        dup.row += int(row)
        dup.column += int(col)
        return dup


class WorksheetUtil:
    '''
    Google Sheets API wrapper class for use more simply.
    Class concept, hundling just one worksheet.
    '''

    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
    scopes = ['https://www.googleapis.com/auth/spreadsheets']

    default_keyfile  = os.getenv('GSS_SERVICE_ACCOUNT_KEY')
    default_interval = datetime.timedelta(seconds=40*60)
    default_pagesize = 2000

    @property
    def selected(self):
        return self._selected

    @selected.setter
    def selected(self, address):
        if '_selected' in list(dir(self)):
            self._selected.address = address
        else:
            self._selected = CellAddressUtil(address)

    @selected.deleter
    def selected(self):
        self._selected.address = 'A1'

    @property
    def was_exceeded_interval(self):
        return datetime.datetime.today() > self.authorized_at + self.interval

    def __init__(self, drive_file_id, **kwargs):
        sheet_id    = kwargs['sheet_id'] if 'sheet_id' in kwargs.keys() else None
        sheet_title = kwargs['sheet_title'] if 'sheet_title' in kwargs.keys() else None
        sheet_index = kwargs['sheet_index'] if 'sheet_index' in kwargs.keys() else None
        keyfile  = kwargs['keyfile'] if 'keyfile' in kwargs.keys() else None
        interval = kwargs['interval'] if 'interval' in kwargs.keys() else None
        pagesize = kwargs['pagesize'] if 'pagesize' in kwargs.keys() else None

        self.spreadsheet_id = drive_file_id
        self.get_credentials(keyfile)
        self.authorization()
        self.cache_metadata(sheet_id, sheet_title, sheet_index)
        self.selected = 'A1'
        self.interval = interval if isinstance(interval, datetime.timedelta) else self.default_interval
        self.pagesize = pagesize if pagesize is not None else self.default_pagesize

    def get_credentials(self, keyfile=None):
        json_path = keyfile or self.default_keyfile
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(json_path, self.scopes)

    def authorization(self):
        http = self.credentials.authorize(httplib2.Http())
        self.service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=self.discoveryUrl)
        self.authorized_at = datetime.datetime.today()

    def refresh_credentials(self):
        if self.was_exceeded_interval: self.authorization()

    def cache_metadata(self, sheet_id=None, sheet_title=None, sheet_index=None):
        if sheet_id is sheet_title is sheet_index is None:
            raise ValueError('Too few arguments. Please specify value about target sheet info as \'sheet_id\' or \'sheet_title\' or \'sheet_index\'.')

        metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        self.target = next(filter(lambda x:
                                  x['properties']['title'] == sheet_title or
                                  x['properties']['index'] == sheet_index or
                                  x['properties']['sheetId'] == sheet_id,
                           metadata['sheets']))['properties']

    def _header_initiation_message(self, headers, **kwargs):
        '''
        :param List headers:
        :param List bg:
        :param List fg:
        :param Dict tab:
        :param int rows:
        :param int cols:
        '''
        background = foreground = []
        tab = None
        rowsize = 1
        colsize = len(headers) or 26

        for k in list(kwargs.keys()):
            if   k == 'bg'  : background = list(kwargs['bg'])
            elif k == 'fg'  : foreground = list(kwargs['fg'])
            elif k == 'tab' : tab = kwargs['tab']
            elif k == 'rows': rowsize = int(kwargs['rows']) if len(headers) else 10000
            elif k == 'cols': colsize = int(kwargs['cols'])
            else: continue

        properties = dict()
        properties['sheetId'] = self.target['sheetId']
        properties['title'] = self.target['title']
        properties['index'] = self.target['index']
        properties.update({ 'gridProperties': { 'rowCount': rowsize, 'columnCount': colsize } })
        if tab is not None:
            properties.update({ 'tabColor': tab })

        message = []
        message.append({
            'updateSheetProperties': {
                'properties': properties,
                'fields': '*'
            }
        })

        for i, value in enumerate(headers):
            operation = {
                'repeatCell': {
                    'range': {
                        'sheetId': properties['sheetId'],
                        'startRowIndex': 0, 'endRowIndex': 1,
                        'startColumnIndex': i, 'endColumnIndex': i+1
                    },
                    'cell': {
                        'userEnteredValue': {
                            'stringValue': value
                        }
                    },
                    'fields': '*'
                }
            }

            if len(background) >= i+1 and background[i] is not None:
                operation['repeatCell']['cell'].update({ 'userEnteredFormat': { 'backgroundColor': background[i] } })

            if len(foreground) >= i+1 and foreground[i] is not None:
                if 'userEnteredFormat' in operation['repeatCell']['cell'].keys():
                    operation['repeatCell']['cell']['userEnteredFormat'].update({ 'textFormat': foreground[i] })
                else:
                    operation['repeatCell']['cell'].update({ 'userEnteredFormat': { 'textFormat': foreground[i] } })

            message.append(operation)

        termination = {
            'repeatCell': {
                'range': {
                    'sheetId': properties['sheetId'],
                    'startRowIndex': 1 if len(headers) > 0 else 0
                },
                'fields': '*'
            }
        }

        if rowsize > 1:
            message.append(termination)

        return message

    def send_update_command(self, values, targetrange, input_opt):
        self.refresh_credentials()
        return self.service.spreadsheets().values().update(
                   spreadsheetId=self.spreadsheet_id,
                   range=targetrange,
                   valueInputOption=input_opt,
                   body={'values': values}
               ).execute()

    def send_batchupdate_command(self, command_message):
        previous = deepcopy(self.target)
        self.refresh_credentials()
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id, body={'requests': command_message}).execute()
        self.cache_metadata(sheet_title=previous['title'])
        return result

    def moveto(self, index:int):
        if index == self.target['index']:
            return True

        properties = deepcopy(self.target)
        properties['index'] = index
        message = [ { 'updateSheetProperties': { 'properties': properties, 'fields': '*' } } ]

        return self.send_batchupdate_command(message)

    def resize_worksheet(self, addrows=0, addcols=0, rows=None, cols=None):
        if addrows == addcols == 0:
            return True

        properties = deepcopy(self.target)
        properties['gridProperties']['rowCount'] = self.target['gridProperties']['rowCount']+addrows if rows is None else rows
        properties['gridProperties']['columnCount'] = self.target['gridProperties']['columnCount']+addcols if cols is None else cols
        message = [ { 'updateSheetProperties': { 'properties': properties, 'fields': '*' } } ]

        if self.target['gridProperties']['rowCount'] == 1 and properties['gridProperties']['rowCount'] > 1:
            message.append({
                'repeatCell': {
                    'range': {
                        'sheetId': properties['sheetId'],
                        'startRowIndex': 1
                    },
                    'fields': '*'
                }
            })

        return self.send_batchupdate_command(message)

    def prepare_worksheet(self, **kwargs):
        '''
        :param List headers
        :param List bg
        :param List fg
        :param Dict tab
        :param int rows
        :param int cols
        '''
        hd = list(kwargs['headers']) if 'headers' in kwargs.keys() else []
        del kwargs['headers']
        message = self._header_initiation_message(hd, **kwargs)

        result = self.send_batchupdate_command(message)
        self.selected = 'A2'

        return result

    def append_records(self, values, input_opt='USER_ENTERED'):
        '''
        :param List values:
        :param str input_opt:
        '''
        pr = paging_records(values, self.pagesize)

        for records in pr:
            startcell = 'A{0}'.format(str(self.target['gridProperties']['rowCount']+1))
            fromaddr = CellAddressUtil(startcell)

            row = len(records)
            col = max(map(lambda x: len(x), records))
            targetrange = '{0}!{1}'.format(self.target['title'], fromaddr.shift(row=row, col=col).to_a1range(fromaddr))

            self.resize_worksheet(addrows=row)
            result = self.send_update_command(records, targetrange, input_opt)

    def write_records(self, values, input_opt='USER_ENTERED'):
        '''
        :param List values:
        :param str input_opt:
        '''
        pr = paging_records(values, self.pagesize)
        for records in pr:
            row = len(records)
            col = max(map(lambda x: len(x), records))
            targetrange = '{0}!{1}'.format(self.target['title'], self.selected.shift(row=row, col=col).to_a1range(self.selected))

            result = self.send_update_command(values, targetrange, input_opt)
            self.selected.row += row

    def write_records_with_prepare(self, values, **kwargs):
        '''
        :param List values:
        :param List headers:
        :param List bg:
        :param List fg:
        :param Dict tab:
        :param str input_opt:
        '''
        input_opt = kwargs['input_opt'] if 'input_opt' in kwargs.keys() else 'USER_ENTERED'

        self.prepare_worksheet(**kwargs)

        pr = paging_records(values, self.pagesize)
        for records in pr:
            self.append_records(records, input_opt)
