#############################
# Created: 4/25/2018
# Author: Morgan Atwood
# Purpose: To report the status of printers on the print server
#############################


#Printer Server
$printserver = 

#Html format
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@



Get-WmiObject -Class Win32_Printer -ComputerName $printserver -ErrorAction Stop | Select-Object Name,` 
    @{Name="PrinterStatus";Expression={Switch($_.PrinterStatus) 
                                        { 
                                            1 {"Toner Low"} 
                                            2 {"Unknown";} 
                                            3 {"Ready"} 
                                            4 {"Printing"} 
                                            5 {"Warming Up"} 
                                            6 {"Stopped printing"} 
                                            7 {"Offline"} 
                                        }}} ,DriverName,PortName |Sort-Object Name | ConvertTo-Html -Head $Header | Out-File -FilePath C:\Temp\PrinterReport.html

$file = Get-Content 'C:\Temp\PrinterReport.html' | Out-String

### EMAIL REPORT ###
$emailaddress = ""
$from = ""
$smtpServer = ""
$Subject = "Printer Summary Report"

send-mailmessage $emailaddress -Subject $subject -Body $file -BodyAsHtml -from $from -SmtpServer $smtpServer -priority High 
