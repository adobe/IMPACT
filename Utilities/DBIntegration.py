# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.
import mysql.connector
def insertLogDatainDB(dataList):
    mysql_config = {
    'host': '',
    'user': '',
    'password': '',
    'database': '',
    'port':''
}
# Connect to the MySQL database
    conn = mysql.connector.connect(**mysql_config)
    cursor = conn.cursor()

    data_to_insert = dataList
    print("Logging : ",data_to_insert)
    insert_query = "INSERT INTO impactdashboard (ProductName,RunAsUser,RunMethod,locale,RunMode,AppType)	 VALUES (%s, %s, %s, %s,%s,%s)"

    # Execute the query with the data
    cursor.execute(insert_query, data_to_insert)

    # Commit the changes to the database
    conn.commit()

    # Close the cursor and conn
    cursor.close()
    conn.close()
    print("Logged !")
