
-----------------------------------------------------Main-----------------------------------------------------
select DISTINCT SPDSF.Month,SPDSF.Fail, SPPT.NumTotal, 
SuccessfulRate=cast(cast(100.0 * (NumTotal - Fail)/NumTotal AS decimal(18, 2)) AS varchar(5))
from 
(
select 
SPDS.ID, SPDS.Title, SPID.Month,
NumTotal=count(*), 
SuccessfulRate = cast(cast(100.0 * SUM(CASE WHEN SPDS.StateDescription = 'Succeeded' THEN 1 ELSE 0 END) / COUNT(*) AS decimal(18, 2)) AS varchar(5)),
SPDS.Month,SPDS.AvailableDate, SPDS.Num_AvailableDays
from [Table01] SPDS
join [Table02] SPID on SPDS.ID = SPID.ID
and SPDS.SubOU = 'xx'
group by column01, column02
) 
SPPT join
(
select DISTINCT SPIL.Month,COUNT(distinct(spdsf.computername)) as fail 
from (SELECT SP.*, AD.operatingsystem,AD.OU
FROM [Table01] SP join [Table03] AD on SP.ComputerName = AD.Machine
where SP.StateID in ('0','2') and SP.SubOU = 'xx'
) SPDSF 
join [Table02] SPIL on SPDSF.ID = SPIL.ID
group by SPIL.Month) SPDSF on SPPT.Month = SPDSF.Month
group by SPDSF.Month,SPDSF.Fail,SPPT.NumTotal
