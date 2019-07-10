$filePath = "C:\Work\Designs\Miscellaneous\Jordan Landry\XXXX - SM Test Folder"
$sheetName = "Test"
$fileName = "blah.xltm"
$objExcel = New-Object -ComObject Exel.Application
$objExcel.Visible = $false
$WorkBook = $objExcel.Workbooks.Open($filePath)

$WorkSheet = $WorkBook.sheets.item($filename)
    if($sheetName -eq"")
    {
    $worksheet = $WorkBook.sheets.Item(4)
    }
    else
    $worksheet = $WorkBook.sheets.Item($sheetName)
    }
    
    $intRowMax = ($worksheet.UsedRange.Rows).count
        

