$csvfile = import-csv -path "uf.csv"

$server = "http://localhost"
$username = "stit.santos"
$password = "Webmotors1"
$Folder = "C:\RelatoriosExportados\IndicadoresWebmotors\"
$tabcmd = "E:\Program Files\Tableau\Tableau Server\10.3\extras\Command Line Utility\tabcmd.exe"
$URLPanel = "/views/EstoqueConcorrente-Redshift/"
$URLWorksheet = New-Object System.Collections.ArrayList
$URLWorksheet.Add("AvaliaoClassificadosporPortal")
$URLWorksheet.Add("ClassificadosporPortaleTipoPessoa")
$URLWorksheet.Add("ClassificadosPorPraa")
$URLAnalitico = "AnaliticoAbrangnciadeClassificadosPopulao"
$NameExport = "Webmotors_Estoque_Concorrente_"
$EmailFrom = "stit.santos@webmotors.com.br"
$EmailTo = "stit.santos@webmotors.com.br"
$DateToday = (Get-Date).ToString("dd/MM")
$EmailSubject = "Relatório Estoque Concorrente"
$EmailBody = "Relatório Estoque Concorrente - $DateToday<BR>"
$SMTPServer = "smtp.office365.com"
$SMTPAuthUsername = "stit.santos@webmotors.com.br"
$SMTPAuthPassword = "Webmotors3"


foreach ($line in $csvfile)
{
  $UF = $line.UF
  $FileNameCsv = $Folder + $NameExport + $UF + ".csv"
  $FullURLCsv = $URLPanel + $URLAnalitico + ".csv" + "?:Refresh&DataAtual=Sim&UF=" + $UF
  $UF
  $FileNameCsv
  $FullURLCsv
     & $tabcmd get -s $server -u $username -p $password  $FullURLCsv  -f $FileNameCsv

function send_email {
   $mailmessage = New-Object system.net.mail.mailmessage 
   $mailmessage.from = ($EmailFrom) 
   $mailmessage.To.add($EmailTo)
   $mailmessage.Subject = $EmailSubject + " " + $UF
   $mailmessage.Body = $EmailBody
   $attachment1 = New-Object System.Net.Mail.Attachment($FileNameCsv)
   $mailmessage.Attachments.Add($attachment1)

  foreach ($Element in $URLWorksheet)
    {
        
      $FileNamePdf = $Folder + $NameExport + $Element + "_" + $UF + ".pdf"
      $FileNamePng = $Folder + $NameExport + $Element + "_" + $UF + ".png"
      
      $FullURLPdf = $URLPanel + $Element + ".pdf" + "?:Refresh&DataAtual=Sim&UF=" + $UF
      $FullURLPng = $URLPanel + $Element + "?:Refresh&DataAtual=Sim&Size=1600,900&UF=" + $UF
      
    $UF
    $Element
    $FileNamePdf
    $FileNamePng
    $FullURLPdf
    $FullURLPng
    print $URLWorksheet[$i]
      & $tabcmd get -s $server -u $username -p $password  $FullURLPdf  -f $FileNamePdf
      & $tabcmd get -s $server -u $username -p $password  $FullURLPng  -f $FileNamePng

   $attachment2 = New-Object System.Net.Mail.Attachment($FileNamePng)
   $attachment3 = New-Object System.Net.Mail.Attachment($FileNamePdf)
   $mailmessage.Attachments.Add($attachment2)
   $mailmessage.Attachments.Add($attachment3)
    }  
    

   $mailmessage.IsBodyHTML = $true
   $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
   $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPAuthUsername", "$SMTPAuthPassword")
   $SMTPClient.EnableSsl = $true
   $SMTPClient.Send($mailmessage)
   }
send_email    

}
& $tabcmd logout