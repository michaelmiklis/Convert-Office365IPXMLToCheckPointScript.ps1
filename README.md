# Convert-Office365IPXMLToCheckPointScript.ps1

This PowerShell script converts the Microsoft Office 365 IP address XML file into a script file for importing
into a CheckPoint Firewall using dbedit.

More details on Microsoft Office 365 IP address ranges and urls can be found here:

https://support.office.com/en-gb/article/office-365-urls-and-ip-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2


Usage
========
Execute the following commandline in a PowerShell Window:

	Convert-Office365IPXMLToCheckPointScript.ps1 -Uri "https://go.microsoft.com/fwlink/?LinkId=533185"
	
If you want to save the output of the PowerShell Script to a file use the following commandline:

	Convert-Office365IPXMLToCheckPointScript.ps1 -Uri "https://go.microsoft.com/fwlink/?LinkId=533185" | Out-File myFile.txt
	
The Checkpoint import will fail if there are duplicate entries in the file - to avoid this use the following commandline:

	Convert-Office365IPXMLToCheckPointScript.ps1 -Uri "https://go.microsoft.com/fwlink/?LinkId=533185" | Select-Object -Unique | Out-File myFile.txt