$ie = new-object -ComObject "InternetExplorer.Application"

$ie.Silent = $true
$ie.navigate("https://www.youtube.com/")

while ($ie.Busy) {Start-Sleep -Milliseconds 100}
$doc = $ie.Document
$btn = $doc.body.getElementsByTagName('input') | Where-Object {$_.title -like 'Pesquisar'}
$btn.value = 'vitas'
$btn = $doc.getElementByid('search-btn') 
$btn.click()
while ($ie.Busy) {Start-Sleep -Milliseconds 500}
$btn = $doc.body.getElementsByTagName("img") | Where-Object {$_.src -eq 'https://i.ytimg.com/vi/B5-X_3_Kpww/hqdefault.jpg?sqp=-oaymwEXCPYBEIoBSFryq4qpAwkIARUAAIhCGAE=&rs=AOn4CLC0W4xXQNnxnhOKcSDCkEBceJ7dVA' }
$btn.click()
while ($ie.Busy) {Start-Sleep -Milliseconds 500}
$ie.visible = $true
