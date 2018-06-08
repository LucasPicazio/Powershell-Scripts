$tabela = Get-ChildItem 'C:\Users\n173983\AppData\Local\Temp' | Sort {$_.LastWriteTime} | select -last 1 

$tabelap = $tabela.FullName

$excel = New-Object -com excel.application

$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

$planilha  = $excel.workbooks.open("$tabelap")

$index = $tabelap.IndexOf('_')

$sdata = Get-Date -UFormat "_%Y_%m_%d.xls"

$tabelap = $tabelap.Substring(0,$index)

$tabelap = "$tabelap$sdata"

$planilha.SaveAs("$tabelap",$xlFixedFormat)

$excel.Quit()
#--------------------------- Seção acima usada para pegar arquivo mudar nome e colocar como read only ---------------------------#

$arquivo = Get-Item $tabelap

$arquivo.IsReadOnly = $true


#--------------------------- Seção acima usada colocar file como read only ---------------------------#
$destination = @("\\csao11p20012a\depto\Correio\DBA\$($arquivo.Name)","\\csao11p20011c\IT\PCHardw\INFRA\Controles\ResourceChecker\$($arquivo.Name)", "\\csao11p20011d\rwapps\CSHG\CSHGDsl\SUPORTE\RemoteTools\$($arquivo.Name)")

$caminhoteste  = "\\csao11p20012a\depto\Correio\DBA"

if(!(Test-Path -Path $caminhoteste)){
    New-Item -ItemType directory -Path $caminhoteste
}


foreach ($dir in $destination)
{
    Copy-Item -Path $tabelap -Destination $dir
}

Rename-Item -Path "\\csao11p20011d\rwapps\CSHG\CSHGDsl\SUPORTE\RemoteTools\$($arquivo.Name)" -NewName "rs$sdata"
#--------------------------- Seção acima usada para colocar arquivos na rede ---------------------------#


$data = Get-Date -UFormat "%d/%m/%Y"

Add-Type -assembly "Microsoft.Office.Interop.Outlook"

$Outlook = New-Object -comobject Outlook.Application

$namespace = $Outlook.GetNameSpace("MAPI")

$email = $Outlook.createItem(0)

$email.to = "list.csbg-usr-support@credit-suisse.com; list.csb-it-dba@credit-suisse.com "
$email.subject = "Resource Checker $data"
$email.body = "Bom dia,

Seguem os caminhos do Resource Checker de hoje:

Suporte: \\csao11p20011c\IT\PCHardw\INFRA\Controles\ResourceChecker\ResourceCheckerExport$sdata

DBA: \\csao11p20012a\depto\Correio\DBA\ResourceCheckerExport$sdata

Atenciosamente,

Lucas Picazio 
IT Infrastructure Servers/User Support 
+55 11 3701 8596 (*551 8596) 
"
$email.send()
pause

#--------------------------- Seção acima usada para mandar e-mail ---------------------------#