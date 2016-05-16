$MSword = new-object -comobject word.application
 $MSword.Visible = $True
 $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat],”wdFormatDocumentDefault”);

#Set the where to find the files and their file type
$Sourcefolder = “c:\doclocation1\*”
$fileType = “*doc”
write-output "Source Folder: " $Sourcefolder;
write-output "File Type: " $fileType;
write-output "Begin Word Document Processing....";
$count = 0

Get-ChildItem -path $Sourcefolder -include $fileType -recurse | foreach-object {
 $opendoc = $MSword.documents.open($_.FullName)
 $savename = ($_.fullname).substring(0,($_.FullName).lastindexOf(“.”)) 
 $opendoc.Convert()
 $opendoc.saveas([ref]”$savename”, [ref]$saveFormat);
 write-output $savename;
 $count++;
 $opendoc.close();
 }

#Clean up
 $MSword.quit() 
 write-output "End Word Document Processing....";
 write-output "Total file Count:" $count;
 
 

$MSexcel = new-object -comobject excel.application
 $MSexcel.Visible = $True
 $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Excel.XlFileFormat],”xlWorkbookDefault”);

#Set the where to find the files and their file type
$Sourcefolder = “c:\doclocation1\*”
$fileType = “*xls”
write-output "Source Folder: " $Sourcefolder;
write-output "File Type: " $fileType;
write-output "Begin Excel Document Processing....";
$count = 0

Get-ChildItem -path $Sourcefolder -include $fileType -recurse | foreach-object {
 $opendoc = $MSexcel.Workbooks.open($_.FullName)
 $savename = ($_.fullname).substring(0,($_.FullName).lastindexOf(“.”)) 
 $opendoc.saveas($savename, $saveFormat);
 write-output $savename;
 $count++;
 $opendoc.close();
 }

#Clean up
 $MSexcel.quit() 
 write-output "End Excel Document Processing....";
 write-output "Total file Count:" $count;
