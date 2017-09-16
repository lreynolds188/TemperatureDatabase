from openpyxl import load_workbook
import sqlite3
import sys

db_filename = 'database.db';


def ConnectDatabase():
    try:
        sys.stdout.write('Attempting database connection... ')
        sys.stdout.flush()
        connection = sqlite3.connect(db_filename)
        return connection
    except:
        print('Error A1: Database connection failed.')
        exit(1)
    finally:
        print('Database connection successful.')


def CreateCountryWorkbook():
    filename_country = 'Assets/GlobalLandTemperaturesByCountry.xlsx'

    # Workbook 1
    sys.stdout.write('Loading Country workbook... ')
    sys.stdout.flush()
    try:
        workbook_country = load_workbook(filename_country)
        sheet_ranges_country = workbook_country.sheetnames
        sheet_country = workbook_country[sheet_ranges_country[0]]
        return sheet_country;
    except:
        print('Error A1: Failed to load Country workbook.')
        exit(1)
    finally:
        print('Country workbook loaded successfully.')


def CreateMajorCityWorkbook():
    filename_major_city = 'Assets/GlobalLandTemperaturesByMajorCity.xlsx'

    # Workbook 2
    sys.stdout.write('Loading Major_City workbook... ')
    sys.stdout.flush()
    try:
        workbook_major_city = load_workbook(filename_major_city)
        sheet_ranges_major_city = workbook_major_city.sheetnames
        sheet_major_city = workbook_major_city[sheet_ranges_major_city[0]]
        return sheet_major_city;
    except:
        print('Error A2: Failed to load Major_City workbook.')
        exit(1)
    finally:
        print('Major_City workbook loaded successfully.')


def CreateStateWorkbook():
    filename_state = 'Assets/GlobalLandTemperaturesByState.xlsx'

    sys.stdout.write('Loading State workbook... ')
    sys.stdout.flush()
    try:
        workbook_state = load_workbook(filename_state)
        sheet_ranges_state = workbook_state.sheetnames
        sheet_state = workbook_state[sheet_ranges_state[0]]
        return sheet_state;
    except:
        print('Error A3: Failed to load State workbook.')
        exit(1)
    finally:
        print('State workbook loaded successfully.\n')


def CreateDatabase():
    conn = ConnectDatabase()
    cursor = conn.cursor()

    try:
        sys.stdout.write('Creating database... ')
        sys.stdout.flush()
        cursor.execute('''DROP TABLE IF EXISTS GlobalLandTemperaturesByCountry''')
        cursor.execute(
            '''CREATE TABLE GlobalLandTemperaturesByCountry (EventDate VARCHAR(25), AverageTemperature DECIMAL(25), AverageTemperatureUncertainty DECIMAL(25), Country VARCHAR(25))''')
        cursor.execute('''DROP TABLE IF EXISTS GlobalLandTemperaturesByMajorCity''')
        cursor.execute(
            '''CREATE TABLE GlobalLandTemperaturesByMajorCity (EventDate VARCHAR(25), AverageTemperature DECIMAL(25), AverageTemperatureUncertainty DECIMAL(25), City VARCHAR(25), Country VARCHAR(25), Latitude VARCHAR(25), Longitude VARCHAR(25))''')
        cursor.execute('''DROP TABLE IF EXISTS GlobalLandTemperaturesByState''')
        cursor.execute(
            '''CREATE TABLE GlobalLandTemperaturesByState (EventDate VARCHAR(25), AverageTemperature DECIMAL(25), AverageTemperatureUncertainty DECIMAL(25), State VARCHAR(25), Country VARCHAR(25))''')
        conn.commit()
        conn.close()
    except:
        print('Error A5: Database creation failed.')
        exit(1)
    finally:
        print('Database creation successful... Database connection closed.\n')


def InsertCountryWorksheet(_worksheet):
    conn = ConnectDatabase()
    cursor = conn.cursor()

    try:
        sys.stdout.write('Inserting Country data into database... ')
        sys.stdout.flush()
        for row in _worksheet.iter_rows():
            _date = row[0].value
            _averageTemp = row[1].value
            _averageTempUncertain = row[2].value
            _country = row[3].value
            _query = "INSERT INTO GlobalLandTemperaturesByCountry (EventDate, AverageTemperature, AverageTemperatureUncertainty, Country) VALUES (?, ?, ?, ?)"
            _values = (_date, _averageTemp, _averageTempUncertain, _country)

            cursor.execute(_query, _values)

        conn.commit()
        conn.close()
    except:
        print('Error A7: Country data insert failed.')
        exit(1)
    finally:
        print('Data insert successful... Database connection closed.\n')


def InsertMajorCityWorksheet(_worksheet):
    conn = ConnectDatabase()
    cursor = conn.cursor()

    try:
        sys.stdout.write('Inserting Major_City data into database... ')
        sys.stdout.flush()
        for row in _worksheet.iter_rows():
            _date = row[0].value
            _averageTemp = row[1].value
            _averageTempUncertain = row[2].value
            _city = row[3].value
            _country = row[4].value
            _latitude = row[5].value
            _longitude = row[6].value
            _query = "INSERT INTO GlobalLandTemperaturesByMajorCity (EventDate, AverageTemperature, AverageTemperatureUncertainty, City, Country, Latitude, Longitude) VALUES (?, ?, ?, ?, ?, ?, ?)"
            _values = (_date, _averageTemp, _averageTempUncertain, _city, _country, _latitude, _longitude)

            cursor.execute(_query, _values)

        conn.commit()
        conn.close()
    except:
        print('Error A9: Database connection failed.')
        exit(1)
    finally:
        print('Data insert successful... Database connection closed.\n')


def InsertStateWorksheet(_worksheet):
    conn = ConnectDatabase()
    cursor = conn.cursor()

    try:
        sys.stdout.write('Inserting State data into database... ')
        sys.stdout.flush()
        for row in _worksheet.iter_rows():
            _date = row[0].value
            _averageTemp = row[1].value
            _averageTempUncertain = row[2].value
            _state = row[3].value
            _country = row[4].value
            _query = "INSERT INTO GlobalLandTemperaturesByState (EventDate, AverageTemperature, AverageTemperatureUncertainty, State, Country) VALUES (?, ?, ?, ?, ?)"
            _values = (_date, _averageTemp, _averageTempUncertain, _state, _country)

            cursor.execute(_query, _values)

        conn.commit()
        conn.close()
    except:
        print('Error A11: State data insert failed.')
        exit(1)
    finally:
        print('Data insert successful... Database connection closed.\n')


CreateDatabase();

country_worksheet = CreateCountryWorkbook();
majorCity_worksheet = CreateMajorCityWorkbook();
state_worksheet = CreateStateWorkbook();

InsertCountryWorksheet(country_worksheet);
InsertMajorCityWorksheet(majorCity_worksheet);
InsertStateWorksheet(state_worksheet);
