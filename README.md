# SAP_AfO_Refresh

A Python script converted into an executable to refresh SAP Analysis for Office Excel Workbooks.

This project is based on the work of:

- Ivan Bondarenko
  - [SAP BO Analysis for Office (BOAO) Automation](https://github.com/IvanBond/SAP-BOA-Automation)
- Arnold Souza
  - [SAP Refresh](https://github.com/ArnoldSouza/SapRefresh)


## Scope

This tool was created to update BOA queries within existing Excel files, which can be used by any user without a Python installation. The parameters for these queries are stored in a singular configuration Excel file, enabling its use across multiple Excel files containing BOA queries.

The primary objective of this tool is to refresh BOA queries. While it offers the capability to execute VBS scripts and other executable files post-refresh, it lacks the functionality to run macros in Excel files, which have historically been prone to errors in VBA-based solutions.

Users have the option to launch the tool by running the EXE file manually or scheduling it for automatic execution, such as through Windows Task Scheduler. When manually initiated, the tool presents a user interface, while in automatic mode, it operates silently in the background utilizing predefined settings.

## Tool Contents

### a. EXE File (mandatory)

The code has been converted into a single EXE file to allow users without Python installed on their PC to use the tool. When the EXE file is started, a temporary Python installation for the duration of the runtime is made in the temporary folder. Starting the EXE file from a local drive ensures faster execution compared to starting it from a network drive.

### b. Configuration File (mandatory)

The configuration file contains all information about the queries to be refreshed and the variable inputs that must be changed. It can be created initially using the EXE file from Excel files containing queries. Additionally, a few settings can be configured in the file, such as sending an email after errors occur during the refresh or starting a script after the refresh is completed. For automated refresh, it must be stored in the same folder as the EXE file.

### c. Password File (recommended)

If the tool is used to update queries automatically, user credentials must be provided. Using the EXE file, you can encode this password file using a secret key, ensuring that only the tool can read the password data. It must be saved in the local user directory or the folder containing the EXE file for automated refresh.

### d. Holiday File (optional)

If queries need to be run on a specific working day, this file is necessary to correctly calculate working days (standard working days from Monday to Friday are considered - this can be changed in the EXE source code). It must be saved in the local user directory or the folder containing the EXE file for automated refresh.

## Pre-requisites

- Configuration File and EXE must be stored on a local path (Netdrive or own PC).
- Ensure all other Excel files are closed (including not visible ones via Task Manager) and no other automations will run during tool usage. The tool may also close all other running Excel instances during a refresh.
- Never click into the command window of the tool, as it can stop code execution. If execution stops, you can continue by pressing Enter.
- Use clean and working files (No old and unused queries, queries without crosstab, no queries that have been deleted or are test queries. Result size should never exceed the limits).
- If running the tool via Task Scheduler, follow instructions in section 6b/c for settings and 7b for scheduling a task.

## Create Configuration

This process generates the initial "Configuration.xlsx" file in the folder where the EXE is located for automating the refresh of your queries. During this procedure, all queries in the selected Excel files will be refreshed once to read the variables and filters in use.

To expedite this process, limit the scope of each query by selecting a small amount of data, e.g., via a variable. You can adjust all variables later in the created configuration file.

There are three ways to create configuration files:

### a. For a Single Query File

### b. For Multiple Files in a Folder

Ensure that only files containing queries, and only those you want to refresh, are in the folder to prevent errors. Only Excel files will be considered. A configuration file can be in the same folder; if the filename contains *config*, it is excluded.

### c. For Multiple Query Files Based on a Configuration File

For this option, you need a configuration file. Enter "Filename", "Filepath", and "Fullpath" on Sheet "Files". Then, a configuration file for all files will be created in the folder where the EXE is located. It is also possible to use a file containing only file links; see the file "File_Input.xlsx" in this folder as an example.

Ensure you create and use the configuration file on the same PC (drives may be named differently, which may cause errors). At a later stage, it is also possible to copy content from two configuration files into a single file if necessary.

## Contents Configuration File

Each configuration file contains five sheets: "Variables_Filters", "Queries", "Files", "Settings", and "Time Values". These sheets cannot be renamed, and no columns in the existing tables can be renamed or reordered. If necessary, it is possible to add further sheets and columns in tables on their right end.

### a. Sheet "Variables_Filters"

Used to define variables and filters for all queries. Column "Command" specifies whether the command is related to a filter or a variable. The value to be used for the next refresh can be set in column "Value" (Excel Formulas can also be used). For time-dependent values, the column "Variable to use as Value" can be filled with a variable name from Sheet "Time Values"; if filled with a variable, the column "Value" will be overwritten.

### b. Sheet "Queries"

Defines if and when the queries should be refreshed. To determine whether a query should be refreshed or not, use column "Refresh on WD". Enter "99" if the query should always be refreshed, and enter a number of a working day if the query should only be refreshed on a certain working day. A query will only be refreshed if the current date matches the correct working day.

To account for company-specific workdays, ensure you provide the path to the "Holidays.csv" in the configuration file or place it in the same folder as the EXE file.

The column "Last refreshed" will contain the time and date of the last refresh and an error description if the refresh failed.

It is also possible to define if the query results should be saved as a CSV file after the refresh is completed. If a CSV is required, enter a path on your local machine in column "Save as CSV".

### c. Sheet "Files"

Contains all files that must be refreshed.

### d. Sheet "Settings"

Refer to section 6b.

### e. Sheet "Time Values"

The values on this sheet will be refreshed in the background before refreshing queries. Time format to be passed to queries is MM.DD.YYYY. As mentioned, the variable names from column "Variable" can be used on Sheet "Variables_Filters" in Column "Variable to use as Value". You can also use more variables in one cell, e.g., LM_MM_YYYY - LM_MM_YYYY.


## Tool Settings

Ensure you provide a path to the password files or credentials if you want to use automated refresh. Alternatively, you can place the password text file in the same folder as the EXE file or the local user directory, and it will be recognized automatically. See the last section for all additional settings.

### a. Settings in the User Interface

#### Settings:
- **Close Running Excel Instances:** Will close all other Excel files during initialization. If deactivated, there is still the possibility that Excel will be closed during the refresh. (Recommendation: Activated)
- **Force Close BOA-Popups:** Will monitor for any pop-ups during query refresh. Pop-ups with the title “Messages” will be recognized as BOA Error messages and will be closed. The error messages will be logged. (Recommendation: Activated)
- **Suppress All BOA Messages:** Will implement a temporary Macro in the Excel files, which triggers the Analysis Plugins “SuppressMessages” function. This is only necessary in special cases, as it is possible to hide non-critical messages in the BOA Analysis Plugin settings. (Recommendation: Deactivated)
- **Run Refresh All in Excel Files After Refresh:** Triggers a Refresh All action in the Excel file. After the Refresh All, there is a waiting time of 60 seconds, which can only be changed in the tool's source code. (Recommendation: Optional)
- **Capture Query Runtimes:** Will capture the query runtime for initial refresh and refresh after submitting the variables in a CSV file. (Recommendation: Optional)
- **Close Excel After Refresh:** Closes Excel after the refresh of all queries is completed, or errors occurred. (Recommendation: Activated)
- **Restart Excel After Refresh Error:** If an error occurs during refreshing a query, this setting will, if activated, save the file and restart Excel. If deactivated, there will be one retry to refresh the query but without restarting Excel in the first place. (Recommendation: Deactivated)
- **Save Query Results in CSV:** Will save each query as a CSV file. Make sure to set correct formats in the Excel files; the tool will read all values from the Excel as a string. (Recommendation: Optional)
- **Always Refresh Queries with Errors During Refresh:** If a query failed to refresh, the column “Last refreshed” in the sheet “Queries” in the configuration file will contain “Error”. All these queries will be refreshed, regardless of any entry in the column “Refresh on WD” which usually steers if a query is to be refreshed or not. (Recommendation: Optional)

### b. Settings in the Configuration File

Every configuration file contains a sheet “Settings” on which basic settings as shown below can be defined.

### c. Settings via Program Arguments

The settings mentioned in section 6a can also be passed to the tool when it is started without the interface. In this case, certain arguments must be passed to the program (see an example for Argument “Scheduled” in section 7b).

The argument “Scheduled” (“arg0”) is a prerequisite to use any other argument and to use the tool for an automated refresh of queries.

Setting                                                     | Argument short             | Argument long               
------------------------------------------------------------|------------------------|--------------------------
Scheduled                                                  | “arg0”                 | “scheduled”              
Activate “Supress all BOA messages”                       | “arg1”                 | “supressboamessages”     
Deactivate “Force close BOA-pop-ups”                      | “arg2”                 | “dontforcecloseboamessages” 
Deactivate “Close running Excel instances”                 | “arg3”                 | “dontcloseexcel”         
Deactivate “Close Excel after refresh”                    | “arg4”                 | “dontkillexcelaftercompletion” 
Deactivate “Run Refresh all in Excel files after refresh” | “arg5”                 | “dontrefresh”            
Activate “Capture query runtimes”                         | “arg6”                 | “captureruntimes”        
Activate “Save Query results in CSV”                      | “arg7”                 | “saveresultscsv”         
Activate “Always refresh queries with errors during refresh” | “arg8”              | “refresherrorqueries”    
Activate “Restart Excel after Refresh Error”              | “arg9”                 | “restartexcelonerror”    

## Refresh Queries

To use automated refresh with the task scheduler, the EXE file and the Configuration File must be in the same folder on a local drive and not on SharePoint or any other Online Storage (Name of the configuration File has to start with Config* - make sure there is only one file like that, the first found file will be used).

### a. Refresh Manually via the Interface

### b. Automated Refresh with Task Scheduler

Select the EXE file as an action in the task scheduler, and add “Scheduled” as an argument; if you don't do this, it won't work as the initial pop-up will show up.

### Possible / Known Sources of Errors:

- Make sure all other Excel files are closed (also not visible ones via the Task Manager must be closed)
- Analysis Addin doesn’t load – open File manually refresh a query and save file.
- Never click into the command window of the tool, this will stop the code execution.
- Use clean files (No old and unused queries, queries without crosstab etc.)
- When the tool cannot open files it is possible that there are hyperlinks on the paths in the configuration file – remove those hyperlinks. Another reason is that you are using paths copied from your browser; this usually does not work, e.g., the filename must end with a valid file extension like .xlsx and not with any other parts from the URL.
- Errors 2147023174 “The RPC server is unavailable” & 2147023170 “The remote procedure call failed” no solution yet – usually works after a couple of minutes after the error.
- Although BOA messages can be hidden using an implemented macro it is safer to activate message suppression in the Analysis Addin settings. By default, it is set to suppress all messages – in case you adjusted this setting, think about setting it back to suppress messages.
- Multiple automations running in parallel (also VBA macros etc.) will likely cause errors – make sure only one is running at a time.
- Make sure all mandatory variables in each query are populated (otherwise variable prompt shows up).
- Cached files are a frequent source of error and cannot be handled programmatically. Therefore, turn on deleting cached files.

### Additional Settings

1. Load and use holiday file
    - Select and use a different Holiday.csv file
2. Create new secret file
    - Creates a secret text file which is the key to decrypt any encoded password file. Make sure the secret file is stored where only you have access to.
3. Encode password file
    - First select a password file which you want to encrypt, then an encoded password file is created using the secret file from “Path secret file”
4. Load and use encoded passwords
    - Selected Encoded Password File will be used

Please note that settings 2-4 potentially must be removed in future versions (encryption).

