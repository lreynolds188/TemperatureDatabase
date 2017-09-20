import sqlite3
import sys

def ConnectDatabase():
    try:
        db_filename = 'assets\database.db';
        sys.stdout.write('Attempting database connection... ')
        sys.stdout.flush()
        conn = sqlite3.connect(db_filename)
        print('Connection successful.\n')
        return conn
    except:
        print('Error Util1: Database connection failed.\n')
        exit(1)


def CloseDatabaseConnection(_connection):
    try:
        conn = _connection
        sys.stdout.write('Closing database connection... ')
        sys.stdout.flush()
        conn.close()
        print('Connection closed.')
    except:
        print('Error Util2: Failed to close database connection.\n')