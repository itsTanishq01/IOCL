import mysql.connector
from mysql.connector import Error

def DropTable():
    # Replace these variables with your MySQL root user connection details
    root_host = 'localhost'
    root_user = 'root'
    root_password = 'root'

    # Change the database name to 'Report'
    database_name = 'Report'

    # Table name to drop
    table_name = 'Input_Table'

    try:
        # Connect to MySQL as root
        root_conn = mysql.connector.connect(host=root_host, user=root_user, password=root_password)
        if root_conn.is_connected():
            cursor = root_conn.cursor()

            # Use the database
            cursor.execute(f"USE {database_name}")

            # Drop the table if it exists
            cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
            print(f'Table {table_name} dropped successfully.')

    except Error as e:
        print(f"Error: {e}")

    finally:
        # Close the cursor and connection
        if root_conn.is_connected():
            cursor.close()
            root_conn.close()
            print('MySQL connection closed')

if __name__ == "__main__":
    DropTable()
