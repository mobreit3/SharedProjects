# SharedProjects
Merritt's Shared Projects

This is the PowerShell Folder under my Shared Projects Repository.

The following powershell program is called convert2010.ps1

The purpose of this program is it will scan a directory on your local hard drive. It will find any Old Microsoft Office 2000/2003 Format for Word (.doc) and Excel (.xls) files in the specified directory. It will automatically convert it to the newer Microsoft Office 2010+ file format. The new files will have the new file extensions: .docx for Word 2010 and later and .xlsx for Excel 2010 and later.

WHY WOULD YOU NEED THIS?
For me it was to convert hundreds of old Word Files that were saved in an old Microsoft Sharepoint Site. The newer versions of Sharepoint (i.e. 2010+) will only allow you to utilize the MS Office .docx and .xlsx files. This was a way to automate the tedious process of opening and performing a "save as" to all these indivdual documentation files prior to migrating them from the OLD sharepoint site to the NEWER sharepoint 2013 version.   


System Requirements: 

Your Workstataion will need a version of Microsoft Office 2010 (Both Word and Excel) installed locally in order for this powershell program to work. It will also requires Windows 7 or later.

Then the following variable ($Sourcefolder).
It should be changed to reflect the location of the OLD Micosoft Office 2000/2003 files on your HardDrive.

$Sourcefolder = “c:\doclocation1\*” 

So C:\doclocation1 should be the Drive including the directory/path where your OLD Microsoft Word and Excel Files are stored.


To Run the Program perform the following steps:
1.Download convert2010.ps1
2. Modify $Sourcefolder variable as necessary
3. Right click on convert2010.ps1 and select "Run with Powershell". 

