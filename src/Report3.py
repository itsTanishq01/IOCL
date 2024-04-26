import pandas as pd
import mysql.connector
from mysql.connector import Error
from tkinter import Tk, filedialog

def get_excel_file_path():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.asksaveasfilename(
        title="Save Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    root.destroy()  # Close the Tkinter window

    return file_path

def export():
    # Replace these variables with your MySQL root user connection details
    root_host = 'localhost'
    root_user = 'root'
    root_password = 'root'

    # Change the database name to 'Report'
    database_name = 'Report'

    # Build the SQL query
    sql_query = '''
        SELECT
            Month,
            Company_Name,
            Ro_Market_Type,
            ROUND(SUM(Ms_Cy_Sales), 2) AS Total_Volume,
            ROUND((SUM(Ms_Cy_Sales) / 
                  (SELECT SUM(Ms_Cy_Sales) 
                   FROM input_table t2 
                   WHERE t2.Company_Name = t1.Company_Name 
                     AND t2.Month = t1.Month)) * 100, 2) AS Percentage
        FROM
            input_table t1
        GROUP BY
            Month, Company_Name, Ro_Market_Type
        ORDER BY
            Month, Company_Name, Ro_Market_Type
    '''

    try:
        # Connect to MySQL as root
        root_conn = mysql.connector.connect(host=root_host, user=root_user, password=root_password)
        if root_conn.is_connected():
            cursor = root_conn.cursor()

            # Use the database
            cursor.execute(f"USE {database_name}")

            # Execute the SQL query
            cursor.execute(sql_query)

            # Fetch the results into a DataFrame
            result_df = pd.DataFrame(cursor.fetchall(), columns=[desc[0] for desc in cursor.description])

            # Get the Excel file path from the user
            excel_file_path = get_excel_file_path()

            # Export the DataFrame to Excel
            result_df.to_excel(excel_file_path, index=False)

            print(f'Data exported to Excel: {excel_file_path}')

    except Error as e:
        print(f"Error: {e}")

    finally:
        # Close the cursor and connection
        if root_conn.is_connected():
            cursor.close()
            root_conn.close()
            print('MySQL connection closed')

if __name__ == "__main__":
    export()
