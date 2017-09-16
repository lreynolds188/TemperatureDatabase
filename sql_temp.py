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
        print('Error B1: Database connection failed.')
        exit(1)
    finally:
        print('Database connection successful.')


def CreateDatabase():
    conn = ConnectDatabase()
    cursor = conn.cursor()

    try:
        sys.stdout.write('Creating database... ')
        sys.stdout.flush()
        cursor.execute('''DROP TABLE IF EXISTS SouthernCities''')
        cursor.execute(
            '''CREATE TABLE SouthernCities (City VARCHAR(25), Country VARCHAR(25), Latitude DECIMAL(25), Longitude DECIMAL(25))''')
        conn.commit()
        conn.close()
    except:
        print('Error B2: Database connection failed.')
        exit(1)
    finally:
        print('Database creation successful... Database connection closed.\n')


def InsertSouthernCitiesWorksheet():
    conn = ConnectDatabase()
    cursor = conn.cursor()

    try:
        sys.stdout.write('Loading Southern_Cities data... ')
        sys.stdout.flush()
        cursor.execute(
            "SELECT City, Country, Latitude, Longitude FROM GlobalLandTemperaturesByMajorCity WHERE Latitude LIKE '%s'")
        rows = cursor.fetchall()
    except:
        print('Error B4: Failed to load Major_City workbook.')
        exit(1)
    finally:
        print('Southern_Cities data loaded successfully.')

    try:
        sys.stdout.write('Inserting Southern_Cities data into database... ')
        sys.stdout.flush()
        for row in rows:
            _city = row[0]
            _country = row[1]
            _latitude = row[2]
            _longitude = row[3]

            _query = "INSERT INTO SouthernCities (City, Country, Latitude, Longitude) VALUES (?, ?, ?, ?)"
            _values = (_city, _country, _latitude, _longitude)

            cursor.execute(_query, _values)

        conn.commit()
        conn.close()
    except:
        print('Error B5: Failed to load Major_City workbook.')
        exit(1)
    finally:
        print('Data insert successful... Database connection closed.\n')


def CalculateMaxMinAvg():
    conn = ConnectDatabase()
    cursor = conn.cursor()

    try:
        sys.stdout.write('Loading max, min, avg temperature data... ')
        sys.stdout.flush()
        cursor.execute(
            "SELECT MAX(AverageTemperature) FROM GlobalLandTemperaturesByState WHERE State = 'Queensland' AND EventDate LIKE '2000%'")
        maxTemp = cursor.fetchone()
        cursor.execute(
            "SELECT MIN(AverageTemperature) FROM GlobalLandTemperaturesByState WHERE State = 'Queensland' AND EventDate LIKE '2000%'")
        minTemp = cursor.fetchone()
        cursor.execute(
            "SELECT AVG(AverageTemperature) FROM GlobalLandTemperaturesByState WHERE State = 'Queensland' AND EventDate LIKE '2000%'")
        avgTemp = cursor.fetchone()
    except:
        print('Error B4: Failed to load Major_City workbook.')
        exit(1)
    finally:
        print('Max, min, avg data loaded successfully.')

    print('Maximum temperature', maxTemp[0])
    print('Minimum temperature', minTemp[0])
    print('Average temperature', avgTemp[0])


CreateDatabase()
InsertSouthernCitiesWorksheet()
CalculateMaxMinAvg()
