
$excel = new-object -comobject Excel.Application

$path = Get-ChildItem 'C:\ResourceCheckerExport_2018_03_09.xls'
$path = $path.FullName
$WorkBook = $excel.Workbooks.Open($path)

while ($true) {
$name = Read-Host 'Digite o login:'
$WorkSheet = $WorkBook.sheets.Item(1)
$usuarios = $WorkSheet.usedrange.columns(3).cells.find("$name").row
$brd = $WorkSheet.Range("A$usuarios")
$brd = $brd.Text
$command = "cmd.exe /C C:\Windows\System32\msra.exe /OfferRA $brd"
Invoke-Expression -Command:$command
}