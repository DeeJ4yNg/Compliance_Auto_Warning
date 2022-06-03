$GLOBAL:criptRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -parent

$fileName = "$criptRoot\test.csv"
$Database = 'database'
$Server = 'abc\db'

function Export_Csv 
{
    $SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Data Source=$Server;Integrated Security=true;Initial Catalog=$Database"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $set = New-Object data.dataset

    # Regional Successful rate by month
    $Regions = ("Region1","Region2","Region3","","Region4")
    foreach ($Region in $Regions)
    {

        $SqlQuery = "SQL Query"

        $SqlCmd.CommandText = $SqlQuery
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        $SqlAdapter.Fill($set)
    }

    # All computers successful rate by month
    $SqlQuery_all = "#SQL Query"
    $SqlCmd.CommandText = $SqlQuery_all
    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $SqlAdapter.Fill($set)

    # Consuming Data
    $Table = $Set.Tables[0]
    $Table | Export-CSV $fileName
}

function SendMail 
{   
    $csv = import-csv $fileName
    $outstanding = $csv | Where-Object {$_.SuccessfulRate -le 50} | out-string
    $subject="Daily Notification"
    $message = 'Outstanding Regions as below, please check.'
    $body = "Dear all,

$message
$outstanding
Attachment:

"
    $Outlook = New-Object -comobject Outlook.Application
    $mail = $Outlook.CreateItem(0) # 1 means Meeting
    $mail.Subject=$subject
    $mail.Body=$body
    $mail.Attachments.Add($fileName)
    $mail.Recipients.Add('dee.w.wu@kpmg.com')
    #$mail.Recipients.Add('irevern.long@kpmg.com')
    #$mail.To = 'dee.w.wu@kpmg.com'
    $mail.Send()
}

Export_Excel
Start-Sleep -Seconds 5
SendMail




