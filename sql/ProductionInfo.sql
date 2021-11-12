SELECT ld.CreateDate, ld.LaborQty, UPPER(ld.EmployeeNum)
FROM EpicorLive10.dbo.LaborDtl ld 
WHERE ld.JobNum=? AND ld.OprSeq =? AND ld.LaborQty <> 0
ORDER BY ld.CreateDate ASC;

SELECT UPPER(ld.EmployeeNum)[Employee], SUM(ld.LaborQty)[TotalQty]
FROM EpicorLive10.dbo.LaborDtl ld 
WHERE ld.JobNum=? AND ld.OprSeq =? AND ld.LaborQty <> 0
GROUP BY UPPER(ld.EmployeeNum)