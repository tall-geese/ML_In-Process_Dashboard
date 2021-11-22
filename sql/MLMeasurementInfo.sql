-- Unique Employees that have ONE or MORE pass Inspections
	SELECT DISTINCT src2.Employee
	FROM (SELECT src.RoutineName, src.ObsID, MAX(src.EmpID)[Employee]
		FROM (SELECT r.RunName, rt.RoutineName, frd.ObsID, COALESCE(dt.ItemName,'???')[EmpID],
				CASE
					WHEN frd.Value > fp.UpperToleranceLimit THEN -1
					WHEN frd.Value < fp.LowerToleranceLimit THEN -1
					WHEN frd.Value IS NULL THEN -1
					ELSE frd.Value 
				END AS [Result]
			FROM MeasurLink7.dbo.Run r 
			LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureRunData frd ON frd.RunID  = r.RunID AND fr.FeatureID = frd.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON frd.FeatureID = f.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND frd.FeatureID = dt.FeatureID AND frd.ObsID = dt.StartObsID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
			WHERE r.RunName = ? AND (rt.RoutineName = ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)
			UNION ALL
			SELECT r.RunName, rt.RoutineName, afrd.ObsID, COALESCE(dt.ItemName,'???')[EmpID],
				CASE 
					WHEN afrd.DefectCount = 1 THEN -1
					WHEN afrd.DefectCount IS NULL THEN -1
					ELSE afrd.DefectCount
				END AS [Result]
			FROM MeasurLink7.dbo.Run r 
			LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
			LEFT OUTER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON afrd.RunID  = r.RunID AND fr.FeatureID = afrd.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON afrd.FeatureID = f.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd.FeatureID = dt.FeatureID AND afrd.ObsID = dt.StartObsID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
			WHERE r.RunName = ? AND (rt.RoutineName = ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)) src
		GROUP BY src.RoutineName, src.ObsID
		HAVING MIN(src.Result) >= 0 )src2;
		  
		  
------------------------------------------------------------------------------
-- Pivot the Count of Valid Inspections for each Employees, for each Routine
	SELECT Pvt.*
	FROM (SELECT src.RoutineName, src.ObsID, MAX(src.EmpID)[Employee]
		FROM (SELECT r.RunName, rt.RoutineName, frd.ObsID, COALESCE(dt.ItemName,'???')[EmpID],
				CASE
					WHEN frd.Value > fp.UpperToleranceLimit THEN -1
					WHEN frd.Value < fp.LowerToleranceLimit THEN -1
					WHEN frd.Value IS NULL THEN -1
					ELSE frd.Value 
				END AS [Result]
			FROM MeasurLink7.dbo.Run r 
			LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureRunData frd ON frd.RunID  = r.RunID AND fr.FeatureID = frd.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON frd.FeatureID = f.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND frd.FeatureID = dt.FeatureID AND frd.ObsID = dt.StartObsID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
			WHERE r.RunName = ? AND (rt.RoutineName = ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)
			UNION ALL
			SELECT r.RunName, rt.RoutineName, afrd.ObsID, COALESCE(dt.ItemName,'???')[EmpID],
				CASE 
					WHEN afrd.DefectCount = 1 THEN -1
					WHEN afrd.DefectCount IS NULL THEN -1
					ELSE afrd.DefectCount
				END AS [Result]
			FROM MeasurLink7.dbo.Run r 
			LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
			LEFT OUTER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON afrd.RunID  = r.RunID AND fr.FeatureID = afrd.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON afrd.FeatureID = f.FeatureID 
			LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd.FeatureID = dt.FeatureID AND afrd.ObsID = dt.StartObsID 
			LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
			WHERE r.RunName = ? AND (rt.RoutineName = ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)) src
		GROUP BY src.RoutineName, src.ObsID
		HAVING MIN(src.Result) >= 0 )src2
	PIVOT(COUNT(ObsId)
		  FOR Employee IN ({Features})
		  ) AS Pvt
	ORDER BY Pvt.RoutineName DESC
		  
		  