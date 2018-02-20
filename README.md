# Convert-Office365IPXMLToCheckPointScript.ps1

This PowerShell script converts the Microsoft Office 365 IP address XML file into a script file for importing
into a CheckPoint Firewall using dbedit.

More details on Microsoft Office 365 IP address ranges and urls can be found here:

https://support.office.com/en-gb/article/office-365-urls-and-ip-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2


Usage
========
Execute the following commandline in a PowerShell Window:

	Convert-Office365IPXMLToCheckPointScript.ps1 -Uri "https://go.microsoft.com/fwlink/?LinkId=533185"
	
