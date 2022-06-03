select 
SPDS.ArticleID, SPDS.Title, SPID.Month,
NumTotal=count(*), 
NumInstalled=isnull(sum(case when SPDS.StateDescription = 'Update is installed' then 1 else 0 end), 0),
NumNotRequired=isnull(sum(case when SPDS.StateDescription = 'Update is not required' then 1 else 0 end), 0),
NumRequired=isnull(sum(case when SPDS.StateDescription = 'Update is required' then 1 else 0 end), 0),
NumNnknown=isnull(sum(case when SPDS.StateDescription = 'Detection state unknown' then 1 else 0 end), 0),
SuccessfulRate = cast(cast(100.0 * SUM(CASE WHEN SPDS.StateDescription in ('Update is installed','Update is not required') THEN 1 ELSE 0 END) / COUNT(*) AS decimal(18, 2)) AS varchar(5))
from [_05_Security_Patches_Deployment_Status_(Win10_20H2)] SPDS
join [_04_Security_Patches_Info_List_Details] SPID on SPDS.CI_ID = SPID.CI_ID
where SPID.OS='Win10 20H2' and SPID.PatchStatus = 'Available'
and SPDS.SubOU = 'BJ'
group by SPDS.CI_ID, SPDS.BulletinID, SPDS.ArticleID, SPDS.Title,SPDS.Month_D,SPDS.AvailableDate, SPDS.Num_AvailableDays,SPID.PatchStatus,SPID.Month, SPDS.SubOU




-----------------------------------------------------Main-----------------------------------------------------
select DISTINCT SPDSF.Month_D,SPDSF.Fail, SPPT.NumTotal, 
SuccessfulRate=cast(cast(100.0 * (NumTotal - Fail)/NumTotal AS decimal(18, 2)) AS varchar(5))
from 
(
select 
SPDS.ArticleID, SPDS.Title, SPID.Month,
NumTotal=count(*), 
NumInstalled=isnull(sum(case when SPDS.StateDescription = 'Update is installed' then 1 else 0 end), 0),
NumNotRequired=isnull(sum(case when SPDS.StateDescription = 'Update is not required' then 1 else 0 end), 0),
NumRequired=isnull(sum(case when SPDS.StateDescription = 'Update is required' then 1 else 0 end), 0),
NumNnknown=isnull(sum(case when SPDS.StateDescription = 'Detection state unknown' then 1 else 0 end), 0),
SuccessfulRate = cast(cast(100.0 * SUM(CASE WHEN SPDS.StateDescription in ('Update is installed','Update is not required') THEN 1 ELSE 0 END) / COUNT(*) AS decimal(18, 2)) AS varchar(5)),
SPDS.Month_D,SPDS.AvailableDate, SPDS.Num_AvailableDays
from [_05_Security_Patches_Deployment_Status_(Win10_20H2)] SPDS
join [_04_Security_Patches_Info_List_Details] SPID on SPDS.CI_ID = SPID.CI_ID
where SPID.OS='Win10 20H2' and SPID.PatchStatus = 'Available'
and SPDS.SubOU = 'BJ'
group by SPDS.CI_ID, SPDS.BulletinID, SPDS.ArticleID, SPDS.Title,SPDS.Month_D,SPDS.AvailableDate, SPDS.Num_AvailableDays,SPID.PatchStatus,SPID.Month, SPDS.SubOU
) 
SPPT join
(
select DISTINCT SPIL.Month_D,COUNT(distinct(spdsf.computername)) as fail 
from (SELECT SP.*, AD.operatingsystem,AD.SubOU as OU
FROM [_05_Security_Patches_Deployment_Status_(Win10_20H2)] SP join _05_AD_CMDB_StaffList AD on SP.ComputerName = AD.AD_Machine
where SP.StateID in ('0','2') and SP.SubOU = 'BJ'
) SPDSF 
join [_04_Security_Patches_Info_List_Details] SPIL on SPDSF.ci_id = SPIL.CI_ID
group by SPIL.Month_D) SPDSF on SPPT.Month_D = SPDSF.Month_D
group by SPDSF.Month_D,SPDSF.Fail,SPPT.NumTotal

--cast(cast(100.0 * SUM(CASE WHEN SPDS.StateDescription in ('Update is installed','Update is not required') THEN 1 ELSE 0 END) / COUNT(*) AS decimal(18, 2)) AS varchar(5))