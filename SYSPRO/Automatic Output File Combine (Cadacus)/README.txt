This script was tested running continuosly on the SYSPRO Application Server via Windows Task Scheduler.
It iterates through the DocumentFiles.txt file and checks to see if the specified document exists.
If the document exists, it calls SYSPRO and runs Output File Combine for specified Communication path.

To add another document to the list, get the full path from SYSPRO/Cadacus Document Setup and paste it on the next line in DocumentFiles.txt

Script Files:
	AutoOFC.ps1
	DocumentFiles.txt
