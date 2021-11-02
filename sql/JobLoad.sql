SELECT src.*
FROM (SELECT r.cell_c, r.Description[Machine], jo.WIStartDate[Start], jh.JobNum,  
		jh.DrawNum[Drawing#], jh.PartNum[Part#], jh.RevisionNum[Part Rev], jh.PartDescription[Description],
		(jo.EstProdHours + jo.EstSetHours)[T Hrs],ROUND((((jo.RunQty - jo.QtyCompleted) * jo.ProdStandard) / 60) + jo.EstSetHours,2)[E Rem Hrs], 
		jo.SetupPctComplete[Set%], jo.RunQty[Run Qty],  jh.ProjectID[Project ID], 
		jo.QtyCompleted[Completed], (jo.RunQty - jo.QtyCompleted)[Remaining],
		jo.Character01[FA Type], jh.PhaseID [Phase ID], jo.OprSeq [Op#], ROW_NUMBER()
		OVER (PARTITION BY r.cell_c, r.Description
				ORDER BY r.cell_c, r.Description, jo.WIStartDate ASC) [R1]
	FROM EpicorLive10.dbo.JobOper jo 
	LEFT OUTER JOIN EpicorLive10.dbo.JobOpDtl jdt ON jo.JobNum  = jdt.JobNum AND jdt.OprSeq = jo.OprSeq 
	LEFT OUTER JOIN EpicorLive10.dbo.Resource r ON jdt.ResourceID = r.ResourceID 
	LEFT OUTER JOIN EpicorLive10.dbo.JobHead jh ON jo.JobNum = jh.JobNum 
	WHERE jo.OpComplete = 0 AND jo.OprSeq = jo.PrimaryProdOpDtl AND jo.OpCode IN ('SWISS','CNC') AND (jo.WIStartDate IS NULL OR jo.WIStartDate > '2021-01-01') 
		AND jo.JobNum NOT LIKE ('%MNT%') AND jh.JobComplete = 0 AND jh.JobEngineered = 1 AND r.Description IS NOT NULL AND jh.JobReleased = 1) src
WHERE src.R1 = 1




