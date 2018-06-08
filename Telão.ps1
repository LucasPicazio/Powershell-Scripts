<#
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$global:Outlook = New-Object -comobject Outlook.Application
$global:namespace = $global:Outlook.GetNameSpace("MAPI")
$global:emails = $global:namespace.Folders.item(3).folders.item(2).folders.item(3).folders.item(5).items
$global:t = get-date -UFormat %m/%d/%y
$global:EMAIL  = $global:emails | where { $global:_.receivedtime -lt $global:t} | select -Last 1

$global:reunioes = ([regex]::Matches($global:email.body, "Responsável Técnico" )).count


function tabela ([String] $global:corte1) {

$global:index = $global:corte1.IndexOf(":")
$global:index = $global:index + 2
$global:corte1 = $global:corte1.Substring($global:index)
$global:partes = $global:corte1.split("`n`r", 2)
$global:nome = $global:partes[0]



$global:partes = $global:partes[1].split("`n`r", 2)
$global:partes = $global:partes[1].split("`n`r", 2)
$global:partes = $global:partes[1].split("`n`r", 2)


$global:x = 3
$global:palavra = "às"
if($global:partes[1] -match "das"){

$global:x = 4
$global:palavra = "das"

}

$global:index = $global:partes[1].IndexOf("$global:palavra")
$global:index = $global:index + $global:x
$global:corte1 = $global:partes[1].Substring($global:index)
$global:partes = $global:corte1.split(" ", 2)
$global:hora = $global:partes[0]




$global:index = $global:partes[1].IndexOf("–")
$global:index = $global:index + 2
$global:corte1 = $global:partes[1].Substring($global:index)
$global:partes = $global:corte1.split("`n`r", 2)
$global:sala = $global:partes[0]




$global:partes = $global:partes[1].split("`n`r", 2)
$global:partes = $global:partes[1].split("`n`r", 2)
$global:partes = $global:partes[1].split("`n`r", 2)

$global:index = $global:partes[1].IndexOf(":")
$global:index = $global:index + 2
$global:corte1 = $global:partes[1].Substring($global:index)
$global:partes = $global:corte1.split("`n`r", 2)
$global:resp = $global:partes[0]
 
 
 $global:hora  | Add-Content 'C:\temp\telao.txt'
 $global:nome  | Add-Content 'C:\temp\telao.txt'
 $global:sala  | Add-Content 'C:\temp\telao.txt'
 $global:resp  | Add-Content 'C:\temp\telao.txt'


 }

 $global:index = $global:email.Body.IndexOf(":")
 $global:corte1 = $global:email.Body.Substring($global:index)
 $global:i = 0
 $path = "C:\temp"
 New-Item -Path $path -Name "telao.txt" -force

 while ($global:i -lt ($global:reunioes)) {

    tabela($global:corte1)
    $global:index = $global:corte1.IndexOf("Responsável Técnico")
    $global:corte1 = $global:corte1.Substring($global:index)
    $global:i++
}

Copy-Item -Path 'C:\temp\telao.txt' -Destination '\\csao11p20011d\Rwapps\CSHG\CSHGDsl\SUPORTE\lpicazio' #>