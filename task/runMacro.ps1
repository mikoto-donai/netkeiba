
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $False
$workbook = $excel.workbooks.open((Convert-Path ..) + "\bin\netkeiba.xlsm")
$excel.Run("controller.main")
$workbook.Save()
$excel.Quit()

$workbook = $null
$excel = $null
[GC]::Collect()
exit