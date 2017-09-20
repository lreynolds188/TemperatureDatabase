from openpyxl import load_workbook
import utilities
import sys
import matplotlib.pyplot as pyplot

conn = utilities.ConnectDatabase()
cursor = conn.cursor()
wb_filename = 'assets\World Temperature.xlsx'


def LoadWorkbook(_filename):
    try:
        sys.stdout.write('Loading World_Temperature workbook... ')
        sys.stdout.flush()

        _workbook = load_workbook(_filename)

        print('World_Temperature workbook loaded successfully.\n')
        return _workbook
    except:
        print('Error N1: Failed to load World_Temperature workbook.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)


def AddWorkbookSheet(_workbook):
    try:
        sys.stdout.write('Creating new worksheet... ')
        sys.stdout.flush()

        _worksheet = _workbook.create_sheet('Comparison')
        _worksheet.title = "Comparison"

        print('Worksheet created successfully.\n')
        return _worksheet
    except:
        print('Error N2: Failed to create worksheet.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)


def SelectAvgStateTempData():
    try:
        sys.stdout.write('Selecting data from Temperatures_by_State... ')
        sys.stdout.flush()

        _query = "SELECT State, AVG(AverageTemperature) FROM GlobalLandTemperaturesByState WHERE AverageTemperature != 'NULL' GROUP BY State ORDER BY AverageTemperature"
        cursor.execute(_query)

        print('Data select successful.\n')
    except:
        print('Error E3: Data select failed.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)


def InsertDataIntoWorkbook(_worksheet):
    try:
        sys.stdout.write('Inserting data into World_Temperature workbook... ')
        sys.stdout.flush()

        _stateTempData = cursor.fetchall()

        for rows in _stateTempData:
            _worksheet.append(rows)

        print('Data successfully inserted.\n')
    except:
        print('Error E4: Data insert failed.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)


def PlotData():
    pyplot.plot


worldTemp_workbook = LoadWorkbook(wb_filename)
worldTemp_worksheet = AddWorkbookSheet(worldTemp_workbook)
SelectAvgStateTempData()
InsertDataIntoWorkbook(worldTemp_worksheet)
worldTemp_workbook.save(wb_filename)
utilities.CloseDatabaseConnection(conn)