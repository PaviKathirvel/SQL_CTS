The file is briefing about the scripts 'Stop_Services.ps1' & 'Start_Services.ps1' and how to run them for MSSQL DB Stop-Start Services Automatically.

Script Location: \\na.jnj.com\ncsusdfsroot\NCSUSGRPDATA\sqlsrvr\MSSQLDML\Scripts\StopStartServices\

Stop-Services
Description:
    This script expects two inputs from the user :  Servers_List, CR number 
		Servers_List is the list of servers stored in a CSV file located in CentralServer(ITSUSRAWSP10439) at 'D:\StopStartServices\StopServices_Servers.csv'
		CR Number is the change control number using for the stop-start services process
	Validates the correctness of the entered CR number. 
		Dynamically picks the value of affected CIs listed in the CR. 
		Compares the servers enlisted in CSV file with the affected CIs of the CR.
		Throws the warning if any unmatched servers are listed in the csv file and procceeds with only matched servers further.
    Checks for the production server in the matched servers and rules them out from the list if any of the prod server exists.
	The blackout duration is set to 1 hr by default. Prompts the user to change the default time if needed.
	Keeps the servers in blackout for the default/modified duration.
	Connects to the servers one at a time, gathers the list of SQL services which are in running state.
	Stops the services and change the starttype to 'manual'.
	Saves all the data(servers and the respective stopped services) in a CSV file presents at the location 'D:\StopStartServices\CSVOutput\'
	Sends out a mail along with the log file to the user who executes the script and the DL in CC.
	
Usage or how to run:
	1. Login to the MSSQL Central server(ITSUSRAWSP10439).
	2. Open the csv file 'D:\StopStartServices\StopServices_Servers.csv' and provide the list of servers where the services are need to be stopped.
	3. Open the PowerShell Console in admin mode and Execute the script \\na.jnj.com\ncsusdfsroot\NCSUSGRPDATA\sqlsrvr\MSSQLDML\Scripts\StopstartServices\Stop_Services.ps1
	4. Enter the valid CR number when the script prompts for the CR Number.
	5. Script will prompts for blackout duration. Default duration is 1 hr. Provide 'y' or 'Y' to change. Provide 'n' or 'N' to not change.
	6. If 'y' is given, then provide the  balckout duration. The format is 'HH:mm' (hrs and minutes).

Notes:
    This version of script is not supporting some special cased servers like prod, AG, clustered.

Start-Services
Description:
	This script expect one input from the user :  CR number 
	Validates the correctness of the entered CR number by checking the output file created during the stop services. 
	Connects to the servers listed in the output file, one at a time
	Starts the services which were stopped by the automation script and change the starttype to 'automatic'.
	Saves all the data(servers and the respective started services) in a CSV file presents at the location 'D:\StopStartServices\CSVOutput\'
	Sends out a mail along with the log file to the user who executes the script and the DL in CC.
	
Usage or how to run:
	1. Login to the MSSQL Central server(ITSUSRAWSP10439).
	2. Open the PowerShell Console in admin mode and Execute the script \\na.jnj.com\ncsusdfsroot\NCSUSGRPDATA\sqlsrvr\MSSQLDML\Scripts\StopstartServices\Start_Services.ps1
	3. Enter the valid CR number when the script prompts for the CR Number.

Notes:
	This version of script is not supporting some special cased servers like prod, AG, clustered.

Script Author: Pavithra K