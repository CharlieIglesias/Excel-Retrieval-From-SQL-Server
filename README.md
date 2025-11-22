# SQL → Excel Scheduled Read of Team Task Data

## Overview
This Excel macro automates reading departmental data from the secure **Task_Data** SQL Server database into an Excel table. It is used daily by 60–80 team members to populate Excel with centralised operational data, which is then used to calculate productivity, utilisation, and capacity metrics. The solution ensures safe, concurrent usage and reliable data retrieval, supporting accurate reporting across the department.

## Key Features & Benefits
- **Automated SQL read:** Pulls all records from the SQL Server table `TeamTaskData` into an Excel table automatically.  
- **Time-controlled scheduling:** Runs at a predefined time each day (9:00 AM), reducing manual intervention.  
- **Safe and concurrent usage:** Checks that no other macro is running and uses a **Running** status to prevent conflicts.  
- **Secure and centralised access:** Retrieves data from the central **Task_Data** SQL Server database for consistency and accuracy.  

## Usage
1. Open the Excel workbook containing the macro.  
2. The macro automatically schedules a daily read at the defined time (or first opening after 9 AM).  
3. When triggered, it:  
   - Checks that no other macro is running, marking the workbook as **Running**.  
   - Waits 10 seconds to ensure the workbook is ready.  
   - Pulls the `TeamTaskData` table from SQL Server into the Excel table.  
   - Waits another 10 seconds and then clears the **Running** status and logs the last read date.  
4. The Excel table is now ready for downstream productivity and utilisation calculations.  

## Technologies / Tools
- Excel VBA (Macros)  
- Microsoft SQL Server (**Task_Data**)  

## Impact
- Eliminates manual copy-paste for retrieving centralised data.  
- Ensures integrity and consistency when reading from SQL Server.  
- Supports daily departmental reporting on productivity, utilisation, and capacity efficiently.  
- Provides a robust foundation for real-time departmental performance monitoring and analysis.
