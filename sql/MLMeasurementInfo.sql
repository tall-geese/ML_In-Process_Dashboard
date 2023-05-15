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
			WHERE r.RunName = ? AND (rt.RoutineName LIKE ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)
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
			WHERE r.RunName = ? AND (rt.RoutineName LIKE ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)) src
		GROUP BY src.RoutineName, src.ObsID
		HAVING MIN(src.Result) >= 0 )src2;
		  
		  
	
--------------------
--------------------
--------------------
--------------------		
-- Pivot the Count of Valid Inspections for each Employees, for each Routine
SELECT Pvt.*
FROM (SELECT src3.RoutineName, src3.ObsNo, MAX(src3.EmpID)[Employee]
	FROM (SELECT r.RunID, rt.RoutineName, src2.ObsNo, COALESCE(dt.ItemName,'???')[EmpID], f.FeatureID,   --Variable Query
			CASE
				WHEN frd2.Value > fp.UpperToleranceLimit THEN -1
				WHEN frd2.Value < fp.LowerToleranceLimit THEN -1
				WHEN frd2.Value IS NULL THEN -1
				ELSE frd2.Value 
			END AS [Result]
		FROM .dbo.Run r 
		LEFT OUTER JOIN .dbo.Routine rt ON r.RoutineID = rt.RoutineID 
		LEFT OUTER JOIN .dbo.FeatureRun fr ON r.RunID = fr.RunID 
		LEFT OUTER JOIN (SELECT src.RunID, src.ObsNo
						FROM (SELECT frd.RunID, frd.ObsNo 
							FROM .dbo.FeatureRunData frd 
							INNER JOIN .dbo.Run r2 ON r2.RunID = frd.RunID
							INNER JOIN .dbo.Routine rt2 ON r2.RoutineID = rt2.RoutineID 
							WHERE r2.RunName = ? AND (rt2.RoutineName LIKE ? OR rt2.RoutineName LIKE '%_IP_%')
							GROUP BY frd.RunID, frd.ObsNo
							UNION ALL	
							SELECT afrd.RunID, afrd.ObsNo 
							FROM .dbo.AttFeatureRunData afrd
							INNER JOIN .dbo.Run r2 ON r2.RunID = afrd.RunID
							INNER JOIN .dbo.Routine rt2 ON r2.RoutineID = rt2.RoutineID 
							WHERE r2.RunName = ? AND (rt2.RoutineName LIKE ? OR rt2.RoutineName LIKE '%_IP_%')
							GROUP BY afrd.RunID, afrd.ObsNo) src
						GROUP BY src.RunId, src.ObsNo) src2 ON r.RunID = src2.RunID
		LEFT OUTER JOIN .dbo.FeatureRunData frd2 ON r.RunID = frd2.RunID AND fr.FeatureID = frd2.FeatureID AND src2.ObsNo = frd2.ObsNo
		LEFT OUTER JOIN .dbo.Feature f ON fr.FeatureID = f.FeatureID 
		LEFT OUTER JOIN .dbo.DataTraceability dt ON r.RunID = dt.RunID AND frd2.FeatureID = dt.FeatureID AND frd2.ObsID = dt.StartObsID
		LEFT OUTER JOIN .dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
		WHERE r.RunName = ? AND (rt.RoutineName LIKE ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)
		UNION ALL
		SELECT r.RunID, rt.RoutineName, src2.ObsNo, COALESCE(dt.ItemName,'???')[EmpID], f.FeatureID,   --Attribute Query
			CASE 
				WHEN afrd2.DefectCount = 1 THEN -1
				WHEN afrd2.DefectCount IS NULL THEN -1
				ELSE afrd2.DefectCount
			END AS [Result]
		FROM .dbo.Run r 
		LEFT OUTER JOIN .dbo.Routine rt ON r.RoutineID = rt.RoutineID 
		LEFT OUTER JOIN .dbo.FeatureRun fr ON r.RunID = fr.RunID 
		LEFT OUTER JOIN (SELECT src.RunID, src.ObsNo
						FROM (SELECT frd.RunID, frd.ObsNo 
							FROM .dbo.FeatureRunData frd 
							INNER JOIN .dbo.Run r2 ON r2.RunID = frd.RunID
							INNER JOIN .dbo.Routine rt2 ON r2.RoutineID = rt2.RoutineID 
							WHERE r2.RunName = ? AND (rt2.RoutineName LIKE ? OR rt2.RoutineName LIKE '%_IP_%')
							GROUP BY frd.RunID, frd.ObsNo
							UNION ALL	
							SELECT afrd.RunID, afrd.ObsNo 
							FROM .dbo.AttFeatureRunData afrd
							INNER JOIN .dbo.Run r2 ON r2.RunID = afrd.RunID
							INNER JOIN .dbo.Routine rt2 ON r2.RoutineID = rt2.RoutineID 
							WHERE r2.RunName = ? AND (rt2.RoutineName LIKE ? OR rt2.RoutineName LIKE '%_IP_%')
							GROUP BY afrd.RunID, afrd.ObsNo) src
						GROUP BY src.RunId, src.ObsNo) src2 ON r.RunID = src2.RunID
		LEFT OUTER JOIN .dbo.AttFeatureRunData afrd2 ON r.RunID = afrd2.RunID AND fr.FeatureID = afrd2.FeatureID AND src2.ObsNo = afrd2.ObsNo
		LEFT OUTER JOIN .dbo.Feature f ON fr.FeatureID = f.FeatureID 
		LEFT OUTER JOIN .dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd2.FeatureID = dt.FeatureID AND afrd2.ObsID = dt.StartObsID
		WHERE r.RunName = ? AND (rt.RoutineName LIKE ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)) src3
	GROUP BY src3.RoutineName, src3.ObsNo
	HAVING MIN(src3.Result) >= 0) src4
PIVOT(COUNT(ObsNo)
	  FOR Employee IN ({Employees})
	  ) AS Pvt
ORDER BY Pvt.RoutineName DESC
			
--------------------
--------------------
--------------------
--------------------





--BACKUP of original query
	--For Routines the collect by random method, or when inspections get skipped over by the Import Templates...
	--If a row is missing inspection data for all but one feature, it is still treated as a pass
------------------------------------------------------------------------------
-- Pivot the Count of Valid Inspections for each Employees, for each Routine


	-- SELECT Pvt.*
	-- FROM (SELECT src.RoutineName, src.ObsID, MAX(src.EmpID)[Employee]
		-- FROM (SELECT r.RunName, rt.RoutineName, frd.ObsID, COALESCE(dt.ItemName,'???')[EmpID],
				-- CASE
					-- WHEN frd.Value > fp.UpperToleranceLimit THEN -1
					-- WHEN frd.Value < fp.LowerToleranceLimit THEN -1
					-- WHEN frd.Value IS NULL THEN -1
					-- ELSE frd.Value 
				-- END AS [Result]
			-- FROM MeasurLink7.dbo.Run r 
			-- LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.FeatureRunData frd ON frd.RunID  = r.RunID AND fr.FeatureID = frd.FeatureID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON frd.FeatureID = f.FeatureID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND frd.FeatureID = dt.FeatureID AND frd.ObsID = dt.StartObsID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
			-- WHERE r.RunName = ? AND (rt.RoutineName LIKE ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)
			-- UNION ALL
			-- SELECT r.RunName, rt.RoutineName, afrd.ObsID, COALESCE(dt.ItemName,'???')[EmpID],
				-- CASE 
					-- WHEN afrd.DefectCount = 1 THEN -1
					-- WHEN afrd.DefectCount IS NULL THEN -1
					-- ELSE afrd.DefectCount
				-- END AS [Result]
			-- FROM MeasurLink7.dbo.Run r 
			-- LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON afrd.RunID  = r.RunID AND fr.FeatureID = afrd.FeatureID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON afrd.FeatureID = f.FeatureID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd.FeatureID = dt.FeatureID AND afrd.ObsID = dt.StartObsID 
			-- LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
			-- WHERE r.RunName = ? AND (rt.RoutineName LIKE ? OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)) src
		-- GROUP BY src.RoutineName, src.ObsID
		-- HAVING MIN(src.Result) >= 0 )src2
	-- PIVOT(COUNT(ObsId)
		  -- FOR Employee IN ({Employees})
		  -- ) AS Pvt
	-- ORDER BY Pvt.RoutineName DESC
		  
		  