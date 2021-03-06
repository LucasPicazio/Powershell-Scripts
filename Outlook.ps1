function inicializa {
    function Get-XamlObject {
	[CmdletBinding()]
	param (
		[Parameter(Position = 0,
				   Mandatory = $true,
				   ValuefromPipelineByPropertyName = $true,
				   ValuefromPipeline = $true)]
		[Alias("FullName")]
		[System.String[]]$Path
	)

	BEGIN
	{
		Set-StrictMode -Version Latest

		$wpfObjects = @{ }
		Add-Type -AssemblyName presentationframework, presentationcore

	} #BEGIN

	PROCESS
	{
		try
		{
			foreach ($xamlFile in $Path)
			{
				#Change content of Xaml file to be a set of powershell GUI objects
				$inputXML = Get-Content -Path $xamlFile -ErrorAction Stop
				$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace 'x:Class=".*?"', '' -replace 'd:DesignHeight="\d*?"', '' -replace 'd:DesignWidth="\d*?"', ''
				[xml]$xaml = $inputXMLClean
				$reader = New-Object System.Xml.XmlNodeReader $xaml -ErrorAction Stop
				$tempform = [Windows.Markup.XamlReader]::Load($reader)

				#Grab named objects from tree and put in a flat structure
				$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
				$namedNodes | ForEach-Object {

					$wpfObjects.Add($_.Name, $tempform.FindName($_.Name))

				} #foreach-object
			} #foreach xamlpath
		} #try
		catch
		{
			throw $error[0]
		} #catch
	} #PROCESS

	END
	{
		Write-Output $wpfObjects
	} #END
}
    function Next {
        $global:i++
        $wpf.from.Text = $emails[$i].sendername
        $wpf.subject.text = $emails[$i].subject
        $wpf.body.Text = $emails[$i].body
    }
    
    Add-Type -assembly "Microsoft.Office.Interop.Outlook"
    $Outlook = New-Object -comobject Outlook.Application
    $namespace = $Outlook.GetNameSpace("MAPI")
    $emails = $namespace.Folders.item(3).folders.item(2).folders.item(3).items
   
    $path = 'C:\Outlook Replyer'

    $wpf = Get-ChildItem -Path $path -Filter *.xaml -file | Where-Object { $_.Name -ne 'App.xaml' } | Get-XamlObject
    $wpf
    #-------------------------------------------Botões---------------------------------------------------------------
    $global:i = 1
    $ticket = Get-Random -max 20529 <#Criação de ticket#>

    $wpf.Confirmadata.add_click({
    
    $datad = $wpf.caixadata.Text
    $wpf.janela1.Close()
    $emails = $emails | sort receivedtime -Descending
    $wpf.from.Text = $emails[$i].sendername
    $wpf.subject.text = $emails[$i].subject
    $wpf.body.Text = $emails[$i].body
    $wpf.janela2.ShowDialog()

    })

    $wpf.Next.add_click({ 
    next
    })

    $wpf.OK.add_click({
    $resposta = $emails[$i].forward()
    if($wpf.geral.ischecked) {
        next
        $resposta.to = "lucas.picazio@credit-suisse.com"
        $resposta.body = "O ticket $ticket foi aberto para um de nossos analistas verificar"
        $resposta.send() 
     }
     if($wpf.FileRecord.ischecked) {
        next
        $resposta.body = "O ticket $ticket foi aberto para um de nossos analistas realizar a gravação"
        $resposta.send() 
     }
     if($wpf.shareddrive.ischecked) {
        next
        $resposta.body = "O ticket $ticket foi aberto para um de nossos analistas liberar o acesso, assim que aprovado"
        $resposta.send() 
     }
     if($wpf.download.ischecked) {
        next
        $resposta.body = "O ticket $ticket foi aberto para um de nossos analistas realizar o download"
        $resposta.send() 
     }
     if($wpf.aprovado.ischecked) {
     #function atualiza e function aprova
        next
        $resposta.body = "O ticket $ticket foi atualizado para o analista verificar."
        $resposta.send() 
     }
     if($wpf.comentario.ischecked) {
        next
        #function atualiza
        $resposta.body = "O ticket $ticket foi atualizado com seu comentario para o analista dar um retorno sobre o processo"
        $resposta.send() 
     }
   
    

})
    
    $wpf.geral.ischecked = $true
    $wpf.janela1.showdialog()
  

    
}



inicializa


<#$excel = new-object -comobject Excel.Application


$pathexcel = "C:\Users\Lucas Picazio\Desktop\usuarios.xlsx"
$WorkBook = $excel.Workbooks.Open($pathexcel)
$WorkSheet = $WorkBook.sheets.Item(1)
$usuarios = $WorkSheet.usedrange.columns(1).cells | Where-Object {$_.text -like '*carlos*'}
$x = $usuarios.row
$analista = $WorkSheet.Range("B$x")
$analista.Text #>

