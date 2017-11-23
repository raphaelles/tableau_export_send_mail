$CsvFile = import-csv -Path "C:\Script\AutomacaoDisparoEmail\uf.csv" -Delimiter "|"

$Server = "http://localhost"
$UserName = "stit.santos"
$Password = "Webmotors1"
$Folder = "C:\RelatoriosExportados\IndicadoresWebmotors\"
$Tabcmd = "E:\Program Files\Tableau\Tableau Server\10.3\extras\Command Line Utility\tabcmd.exe"
$UrlPanel = "/views/EstoqueConcorrente-Redshift/"
$UrlWorksheet = New-Object System.Collections.ArrayList
$UrlWorksheet.Add("AvaliaoClassificadosporPortal")
$UrlWorksheet.Add("ClassificadosporPortaleTipoPessoa")
$UrlWorksheet.Add("ClassificadosPorPraa")
$UrlAnalitico = "AnaliticoAbrangnciadeClassificadosPopulao"
$NameExport = "Webmotors_Estoque_Concorrente_"

$EmailFrom = "marcus.martins@webmotors.com.br"
$EmailTo = "stit.santos@webmotors.com.br"
##"ivan.azevedo@webmotors.com.br"
$DateToday = (Get-Date).ToString("dd/MM")
$EmailSubject = "Relatório Estoque Concorrente"
$EmailBody = "Relatório Estoque Concorrente - $DateToday<BR>"
$SmtpServer = "smtp.office365.com"
##"10.11.4.176"
##"smtp.office365.com"
$SmtpAuthUsername = "marcus.martins@webmotors.com.br"
$SmtpAuthPassword = "ALkb7yqawx#"


foreach ($Line in $CsvFile)
{
  $UF = $Line.UF
  $FileNameCsv = $Folder + $NameExport + ($UF -replace ",", "_") + ".csv"
  $FullUrlCsv = $UrlPanel + $UrlAnalitico + ".csv" + "?:Refresh&DataAtual=Sim&UF=" + $UF
  $UF
  $FileNameCsv
  $TesteArquivo = Test-Path $FileNameCsv
  IF ($TesteArquivo -eq "True")
      { Remove-Item $FileNameCsv }
  $FullUrlCsv
     & $Tabcmd get -s $Server -u $UserName -p $Password  $FullUrlCsv  -f $FileNameCsv


   $MailMessage = New-Object System.Net.Mail.MailMessage 
   $MailMessage.from = ($EmailFrom) 
   $MailMessage.To.add($EmailTo)
   $MailMessage.Subject = $EmailSubject + " " + $UF
   $MailMessage.Body = $EmailBody
   $MailMessage.IsBodyHtml = $True
   $Attachment1 = New-Object System.Net.Mail.Attachment($FileNameCsv)
   $MailMessage.Attachments.Add($Attachment1)

  foreach ($Element in $UrlWorksheet)
    {
        
      $FileNamePdf = $Folder + $NameExport + $Element + "_" + ($UF -replace ",", "_") + ".pdf"
      $FileNamePng = $Folder + $NameExport + $Element + "_" + ($UF -replace ",", "_") + ".png"
      
      $FullUrlPdf = $UrlPanel + $Element + ".pdf" + "?:Refresh&DataAtual=Sim&UF=" + $UF
      $FullUrlPng = $UrlPanel + $Element + ".png" + "?:Refresh&DataAtual=Sim&Size=1600,900&UF=" + $UF
      
    $Element
    $FileNamePdf
    $FileNamePng
    $TesteArquivo = Test-Path $FileNamePdf
     IF ($TesteArquivo -eq "True")
         { Remove-Item $FileNamePdf }
    $TesteArquivo = Test-Path $FileNamePng
     IF ($TesteArquivo -eq "True")
         { Remove-Item $FileNamePng }
    $FullUrlPdf
    $FullUrlPng
      & $Tabcmd get -s $Server -u $UserName -p $Password  $FullUrlPdf  -f $FileNamePdf
      & $Tabcmd get -s $Server -u $UserName -p $Password  $FullUrlPng  -f $FileNamePng

   $Attachment2 = New-Object System.Net.Mail.Attachment($FileNamePng)
   $Attachment3 = New-Object System.Net.Mail.Attachment($FileNamePdf)
   $MailMessage.Attachments.Add($Attachment2)
   $MailMessage.Attachments.Add($Attachment3)
    }  
    


   $Smtp = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
   $Smtp.Credentials = New-Object System.Net.NetworkCredential("$SmtpAuthUsername", "$SmtpAuthPassword")
   $Smtp.EnableSsl = $True
   $Smtp.Send($MailMessage)


}
& $Tabcmd logout
