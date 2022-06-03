$GLOBAL:criptRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -parent

$fileName = "$criptRoot\RegionalStatus.csv"
$Database = 'CM_CHN'
$Server = 'cnhkgsms01\casdb01'

function Export_Excel 
{
    # Accessing Data Base
    $SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Data Source=$Server;Integrated Security=true;Initial Catalog=$Database"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $set = New-Object data.dataset

    # Regional Successful rate by month
    $Regions = ("BJ","SH","GZ","SZ","HK","FS","NJ","NJ2","CD","HZ","QD","XM","TJ","SY","XM","Servicing","OTHERS","Staging","HSBC_KYC")
    foreach ($Region in $Regions)
    {

        $SqlQuery = "select DISTINCT SPDSF.Month_D, SPPT.SubOU, SPDSF.Fail, SPPT.NumTotal, 
SuccessfulRate=cast(cast(100.0 * (NumTotal - Fail)/NumTotal AS decimal(18, 2)) AS varchar(5))
from 
(
select 
SPDS.ArticleID, SPDS.Title, SPID.Month,
NumTotal=count(*), 
SuccessfulRate = cast(cast(100.0 * SUM(CASE WHEN SPDS.StateDescription in ('Update is installed','Update is not required') THEN 1 ELSE 0 END) / COUNT(*) AS decimal(18, 2)) AS varchar(5)),
SPDS.Month_D,SPDS.AvailableDate, SPDS.Num_AvailableDays, SPDS.SubOU
from [_05_Security_Patches_Deployment_Status_(Win10_20H2)] SPDS
join [_04_Security_Patches_Info_List_Details] SPID on SPDS.CI_ID = SPID.CI_ID
where SPID.OS='Win10 20H2' and SPID.PatchStatus = 'Available'
and SPDS.SubOU = '$Region'
group by SPDS.CI_ID, SPDS.BulletinID, SPDS.ArticleID, SPDS.Title,SPDS.Month_D,SPDS.AvailableDate, SPDS.Num_AvailableDays,SPID.PatchStatus,SPID.Month, SPDS.SubOU
) 
SPPT join
(
select DISTINCT SPIL.Month_D,COUNT(distinct(spdsf.computername)) as fail 
from (SELECT SP.*, AD.operatingsystem,AD.SubOU as OU
FROM [_05_Security_Patches_Deployment_Status_(Win10_20H2)] SP join _05_AD_CMDB_StaffList AD on SP.ComputerName = AD.AD_Machine
where SP.StateID in ('0','2') and SP.SubOU = '$Region'
) SPDSF 
join [_04_Security_Patches_Info_List_Details] SPIL on SPDSF.ci_id = SPIL.CI_ID
group by SPIL.Month_D) SPDSF on SPPT.Month_D = SPDSF.Month_D
group by SPDSF.Month_D,SPDSF.Fail,SPPT.NumTotal, SPPT.SubOU"

        $SqlCmd.CommandText = $SqlQuery
        $SqlCmd.Connection = $SqlConnection
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
        $SqlAdapter.Fill($set)
    }

    # All computers successful rate by month
    $SqlQuery_all = "select DISTINCT SPDSF.Month_D,SubOU='All', SPDSF.Fail, SPPT.Total_all as NumTotal, SuccessfulRate = cast(cast(100.0 * (Total_all - Fail)/Total_all AS decimal(18, 2)) AS varchar(5))
from [_05_Security_Patches_Deployment_Status_(Win10_20H2)(PivotTable)] SPPT 
join 
(
select DISTINCT SPIL.Month_D,COUNT(distinct(spdsf.computername)) as fail 
from (SELECT SP.*, AD.operatingsystem,AD.SubOU as OU
FROM [_05_Security_Patches_Deployment_Status_(Win10_20H2)] SP join _05_AD_CMDB_StaffList AD on SP.ComputerName = AD.AD_Machine
where SP.StateID in ('0','2') 
) SPDSF
join [_04_Security_Patches_Info_List_Details] SPIL on SPDSF.ci_id = SPIL.CI_ID
group by SPIL.Month_D) SPDSF on SPPT.Month_D = SPDSF.Month_D
group by SPDSF.Month_D,SPDSF.Fail,SPPT.Total_all"
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
    <th align=left><b>Month_D</b></th>
    <th align=left><b>SubOU</b></th>
    <th align=left><b>Fail</b></th>
    <th align=left><b>NumTotal</b></th>
    <th align=left><b>SuccessfulRate</b></th>
    </tr>
    "
    foreach ($row in $total)
    {
        $HtmlTable1 += "<tr style='font-size:12px;background-color:#FFFFFF'>
    <td>" + $row.Month_D + "</td>
    <td>" + $row.SubOU + "</td>
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
    <th align=left><b>Month_D</b></th>
    <th align=left><b>SubOU</b></th>
    <th align=left><b>Fail</b></th>
    <th align=left><b>NumTotal</b></th>
    <th align=left><b>SuccessfulRate</b></th>
    </tr>"
    foreach ($row in $outstanding)
    {
        $HtmlTable2 += "<tr style='font-size:12px;background-color:#FFFFFF'>
    <td>" + $row.Month_D + "</td>
    <td>" + $row.SubOU + "</td>
    <td>" + $row.Fail + "</td>
    <td>" + $row.NumTotal + "</td>
    <td>" + $row.SuccessfulRate + "</td>
    </tr>"
    }
    $HtmlTable2 += "</table>"

    $subject="Patch Daily Notification"
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
    $mail.Recipients.Add('dee.w.wu@kpmg.com')
    $mail.Recipients.Add('irevern.long@kpmg.com')
    #$mail.To = 'dee.w.wu@kpmg.com'
    $mail.Send()
}

Export_Excel
Start-Sleep -Seconds 5
SendMail




