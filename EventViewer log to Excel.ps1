# Getting the  EventLogs for Source named Application Error from Application folder
$CFLog = Get-EventLog -LogName Application -Source "Application Error" 
$CFLog | Select-Object -Property *

#Excel Com object
$Excel=new-object -ComObject Excel.Application
$Excel.Visible = $True  # setting visibility to true

#adding and and selecting new excel page
$xl = $Excel.Workbooks.Add()
#$myXl = $Excel.Selection
$myXlSheet = $Excel.Worksheets.Item("sheet1")
$myXlSheet.activate()

$myXlSheet.Cells.Item(1,1) = "Time Generated"
$myXlSheet.Cells.Item(1,2) = "Machine Name"
$myXlSheet.Cells.Item(1,3) = "Entry Type"
$myXlSheet.Cells.Item(1,4) = "Source"
$myXlSheet.Cells.Item(1,5) = "Message"

$row1 = 2;


foreach($item in $CFLog){
IF($item){
    $myXlSheet.Cells.Item($row1,1) = $item.TimeGenerated
    $myXlSheet.Cells.Item($row1,2) = $item.MachineName
    $myXlSheet.Cells.Item($row1,3) = $item.EntryType
    $myXlSheet.Cells.Item($row1,4) = $item.Source
    $myXlSheet.Cells.Item($row1,5) = $item.Message
    }
    $row1++
    }


# Set the width of the columns automatically
$myXlSheet.columns.item("A:J").EntireColumn.AutoFit() | out-null

$Excel.DisplayAlerts = 'True'
$ext=".xlsx"
$path="D:\Logs\EventLogs$ext"
$myXlSheet.SaveAs($path) 
$myXlSheet.Close
$Excel.DisplayAlerts = 'True'
$Excel.Quit()
