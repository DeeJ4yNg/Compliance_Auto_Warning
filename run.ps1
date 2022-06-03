$fileName = "RegionalStatus.csv"
$Database = 'Database'
$Server = 'abc\db01'

function Export_CSV
{
    # Accessing Data Base
    $SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Data Source=$Server;Integrated Security=true;Initial Catalog=$Database"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $set = New-Object data.dataset

    # Regional Successful rate by month
    $Regions = ("Region1","Region2","Region3","Region4")
    foreach ($Region in $Regions)
    {

        $SqlQuery = "#SQL Query"

        $SqlCmd.CommandText = $SqlQuery
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        $SqlAdapter.Fill($set)
    }

    
    $SqlQuery_all = "#SQL Query"
    $SqlCmd.CommandText = $SqlQuery_all
    $SqlCmd.Connection = $SqlConnection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $SqlAdapter.Fill($set)
    $Table = $Set.Tables[0]
    $Table | Export-CSV $fileName
}

function SendMail 
{   
    $csv = import-csv $fileName
    $outstanding = $csv | Where-Object {$_.SuccessfulRate -le 50}
    $total = $csv | Where-Object {$_.SubOU -le 'All'}
    $HTMLmessage1 = "<p>Dear all,</p>
    <p>Total Successful rate by month.</p>"
    $HTMLmessage2 = "<p>Outstanding Regions as below.</p>"
    $HTMLmessage3 = "<p>Regards,</p>"

    #HTML table all
    $HtmlTable1 = "<table border='1' align='Left' cellpadding='2' cellspacing='0' style='color:black;font-family:arial,helvetica,sans-serif;text-align:left;'>
    <caption><b>Total Compliance Status</b></caption>
    <tr style ='font-size:12px;font-weight: normal;background: #FFFFFF'>
    <th align=left><b>Month</b></th>
    <th align=left><b>OU</b></th>
    <th align=left><b>Fail</b></th>
    <th align=left><b>NumTotal</b></th>
    <th align=left><b>SuccessfulRate</b></th>
    </tr>
    "
    foreach ($row in $total)
    {
        $HtmlTable1 += "<tr style='font-size:12px;background-color:#FFFFFF'>
    <td>" + $row.Month + "</td>
    <td>" + $row.OU + "</td>
    <td>" + $row.Fail + "</td>
    <td>" + $row.NumTotal + "</td>
    <td>" + $row.SuccessfulRate + "</td>
    </tr>"
    }
    $HtmlTable1 += "</table>"

    #HTML table outstanding
    $HtmlTable2 = "<table border='1' align='Left' cellpadding='2' cellspacing='0' style='color:black;font-family:arial,helvetica,sans-serif;text-align:left;'>
    <caption><b>Outstanding Regions</b></caption>
    <tr style ='font-size:12px;font-weight: normal;background: #FFFFFF'>
    <th align=left><b>Month</b></th>
    <th align=left><b>OU</b></th>
    <th align=left><b>Fail</b></th>
    <th align=left><b>NumTotal</b></th>
    <th align=left><b>SuccessfulRate</b></th>
    </tr>"
    foreach ($row in $outstanding)
    {
        $HtmlTable2 += "<tr style='font-size:12px;background-color:#FFFFFF'>
    <td>" + $row.Month + "</td>
    <td>" + $row.OU + "</td>
    <td>" + $row.Fail + "</td>
    <td>" + $row.NumTotal + "</td>
    <td>" + $row.SuccessfulRate + "</td>
    </tr>"
    }
    $HtmlTable2 += "</table>"

    $subject="Daily Notification"
    $body = "<body>
    <p>Dear all,</p>
    <p>Please check attachment for more region status.</p>
    $HtmlTable1
    $HtmlTable2
    </body>
    "
    $Outlook = New-Object -comobject Outlook.Application
    $mail = $Outlook.CreateItem(0)
    $mail.Subject=$subject
    $mail.HTMLBody=$body
    $mail.Attachments.Add($fileName)
    $mail.Recipients.Add('somebody@gmail.com')
    $mail.Send()
}

Export_CSV
Start-Sleep -Seconds 5
SendMail




