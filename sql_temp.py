import sys
import utilities

conn = utilities.ConnectDatabase()
cursor = conn.cursor()


def CreateDatabase():
    try:
        sys.stdout.write('Creating database... ')
        sys.stdout.flush()
        cursor.execute('''DROP TABLE IF EXISTS SouthernCities''')
        cursor.execute('''CREATE TABLE SouthernCities (City VARCHAR(25), Country VARCHAR(25), Latitude DECIMAL(25), Longitude DECIMAL(25))''')
        conn.commit()
        print('Database creation successful.\n')
    except:
        print('Error S1: Database connection failed.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)


def InsertSouthernCitiesWorksheet():
    try:
        sys.stdout.write('Inserting Southern_Cities data into database... ')
        sys.stdout.flush()
        rows = cursor.fetchall()
        for row in rows:
            _city = row[0]
            _country = row[1]
            _latitude = row[2]
            _longitude = row[3]

            _query = "INSERT INTO SouthernCities (City, Country, Latitude, Longitude) VALUES (?, ?, ?, ?)"
            _values = (_city, _country, _latitude, _longitude)

            cursor.execute(_query, _values)
        conn.commit()
        print('Data insert successful.\n')
    except:
        print('Error S2: Failed to load Major_City workbook.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)


def SelectSouthernCitiesData(_cursor):
    try:
        cursor = _cursor
        sys.stdout.write('Loading Southern_Cities data... ')
        sys.stdout.flush()
        cursor.execute("SELECT City, Country, Latitude, Longitude FROM GlobalLandTemperaturesByMajorCity WHERE Latitude LIKE '%s'")
        print('Southern_Cities data loaded successfully.\n')
    except:
        print('Error S3: Failed to load Major_City workbook.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)


def CalculateMaxMinAvg():
    try:
        sys.stdout.write('Loading max, min, avg temperature data... ')
        sys.stdout.flush()
        cursor.execute("SELECT MAX(AverageTemperature) FROM GlobalLandTemperaturesByState WHERE State = 'Queensland' AND EventDate LIKE '2000%'")
        maxTemp = cursor.fetchone()
        cursor.execute("SELECT MIN(AverageTemperature) FROM GlobalLandTemperaturesByState WHERE State = 'Queensland' AND EventDate LIKE '2000%'")
        minTemp = cursor.fetchone()
        cursor.execute("SELECT AVG(AverageTemperature) FROM GlobalLandTemperaturesByState WHERE State = 'Queensland' AND EventDate LIKE '2000%'")
        avgTemp = cursor.fetchone()
        print('Max, min, avg data loaded successfully.\n')
    except:
        print('Error S4: Failed to load Major_City workbook.\n')
        utilities.CloseDatabaseConnection(conn)
        exit(1)

    print('Maximum temperature', maxTemp[0])
    print('Minimum temperature', minTemp[0])
    print('Average temperature', avgTemp[0], '\n')


CreateDatabase()
InsertSouthernCitiesWorksheet()
CalculateMaxMinAvg()
utilities.CloseDatabaseConnection(conn)