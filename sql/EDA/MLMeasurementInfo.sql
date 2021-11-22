
--These are the routines that have been created for the job so far....
SELECT r.RunName,r.RunID, r.RoutineID, rt.RoutineID, rt.RoutineName 
FROM MeasurLink7.dbo.Run r
LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID  = rt.RoutineID 
--WHERE r.RunName = '003864-32-1'
WHERE r.RunID =11714


--Find all of the Routines that Exist in the Database for this Part
SELECT rt.RoutineName, p.PartName 
FROM MeasurLink7.dbo.RoutineFeatures rf 
LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON rf.RoutineID  = rt.RoutineID 
LEFT OUTER JOIN MeasurLInk7.dbo.Feature f ON rf.FeatureID  = f.FeatureID 
LEFT OUTER JOIN MeasurLink7.dbo.Part p ON f.PartID  = p.PartID 
WHERE p.PartName = 'P20-440-026F_A' AND (rt.RoutineName LIKE '%IP_%' OR rt.RoutineName LIKE '%FA_MINI%')   --The desired FAI Routine would change here depending on setup Type
GROUP BY rt.RoutineName, p.PartName 

--What employees have taken inspections for the above Routines?
	--Variable Features
SELECT src.*
FROM (SELECT r.RunName, rt.RoutineName, f.FeatureID, f.FeatureName, f.FeatureType, frd.ObsID, frd.Value, dt.StartObsID, dt.ItemName, fp.LowerToleranceLimit, fp.UpperToleranceLimit,
		CASE
			WHEN frd.Value > fp.UpperToleranceLimit THEN 'F'
			WHEN frd.Value < fp.LowerToleranceLimit THEN 'F'
			WHEN frd.Value IS NULL THEN 'F'
			ELSE 'P'
		END AS [Result]
	FROM MeasurLink7.dbo.Run r 
	LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureRunData frd ON frd.RunID  = r.RunID AND fr.FeatureID = frd.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON frd.FeatureID = f.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND frd.FeatureID = dt.FeatureID AND frd.ObsID = dt.StartObsID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
	WHERE r.RunName = 'NV14408' AND rt.RoutineName = '1642652_D_IP_1XSHIFT' AND f.FeatureType = 1 AND dt.TraceabilityListID = 143)src
ORDER BY src.ObsID ASC
	--Can either be the Run ID of the RunName and RoutineName
--WHERE src.Result <> 'F'

--SAME but for ATTRIBUTE
SELECT src.*
FROM (SELECT r.RunName, rt.RoutineName, f.FeatureID, f.FeatureName, f.FeatureType, afrd.ObsID, afrd.DefectCount, dt.StartObsID, COALESCE(dt.ItemName,'???')[EmpID], fp.LowerToleranceLimit, fp.UpperToleranceLimit,
		COALESCE(afrd.DefectCount,1)[Result]
	FROM MeasurLink7.dbo.Run r 
	LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
	LEFT OUTER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON afrd.RunID  = r.RunID AND fr.FeatureID = afrd.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON afrd.FeatureID = f.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd.FeatureID = dt.FeatureID AND afrd.ObsID = dt.StartObsID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
	WHERE r.RunName = 'SB0019' AND rt.RoutineName = '50-01-0002_A_IP_BENCH' AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL))src  --Can either be the Run ID of the RunName and RoutineName
ORDER BY src.ObsID ASC

--Very TEMP
SELECT src.*
FROM (SELECT r.RunName, rt.RoutineName, f.FeatureID, f.FeatureName, f.FeatureType, afrd.ObsID, afrd.DefectCount, dt.TraceabilityListID,
		COALESCE(afrd.DefectCount,1)[Result]
	FROM MeasurLink7.dbo.Run r 
	LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
	LEFT OUTER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON afrd.RunID  = r.RunID AND fr.FeatureID = afrd.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON afrd.FeatureID = f.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd.FeatureID = dt.FeatureID AND afrd.ObsID = dt.StartObsID
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
	WHERE r.RunName = 'SD0938' AND rt.RoutineName = 'DRW-00815-02_RAG_FI_VIS' AND f.FeatureType = 2 )src  --Can either be the Run ID of the RunName and RoutineName
ORDER BY src.ObsID ASC





------------------------------------------------------------------------------
--Finished Query In Progress
--Testing GROUP BY method where we can remove an entire observation
SELECT MAX(src.RunName)[Run], MAX(src.RoutineName)[Routine], MAX(src.EmpID)[Employee], src.ObsID, MIN(src.Result)
FROM (SELECT r.RunName, rt.RoutineName, f.FeatureID, f.FeatureName, f.FeatureType, frd.ObsID, frd.Value, dt.StartObsID, COALESCE(dt.ItemName,'???')[EmpID], fp.LowerToleranceLimit, fp.UpperToleranceLimit,
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
	WHERE r.RunName = 'NV14408' AND rt.RoutineName = '1642652_D_IP_1XSHIFT' AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL))src  --Can either be the Run ID of the RunName and RoutineName
GROUP BY src.ObsID
HAVING MIN(src.Result) >= 0

--Attribute Feature Testing
SELECT MAX(src.RunName)[Run], MAX(src.RoutineName)[Routine], MAX(src.EmpID)[Employee], src.ObsID, MAX(src.Result)[Result]
FROM (SELECT r.RunName, rt.RoutineName, f.FeatureID, f.FeatureName, f.FeatureType, afrd.ObsID, afrd.DefectCount, dt.StartObsID, COALESCE(dt.ItemName,'???')[EmpID], fp.LowerToleranceLimit, fp.UpperToleranceLimit,
		COALESCE(afrd.DefectCount,1)[Result]
	FROM MeasurLink7.dbo.Run r 
	LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
	LEFT OUTER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON afrd.RunID  = r.RunID AND fr.FeatureID = afrd.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON afrd.FeatureID = f.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd.FeatureID = dt.FeatureID AND afrd.ObsID = dt.StartObsID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
	WHERE r.RunName = 'NV14408' AND rt.RoutineName = '1642652_D_IP_1XSHIFT' AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL))src  --Can either be the Run ID of the RunName and RoutineName
GROUP BY src.ObsID
HAVING MAX(src.Result) = 0



------------------------------------------------------------------------------
-- Working Test for the one Job and One Routine
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
			WHERE r.RunName = 'NV14408' AND rt.RoutineName = '1642652_D_IP_1XSHIFT' AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)
			UNION ALL
		--Attribute Feature Testing
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
			WHERE r.RunName = 'NV14408' AND rt.RoutineName = '1642652_D_IP_1XSHIFT' AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)) src
		GROUP BY src.RoutineName, src.ObsID
		HAVING MIN(src.Result) >= 0 )src2
	PIVOT(COUNT(ObsId)
		  FOR Employee IN ([0277], [0423])
		  ) AS Pvt
		  
		  
------------------------------------------------------------------------------
-- Working Test for All Routines for the Job and All Employees
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
			WHERE r.RunName = 'NV14408' AND (rt.RoutineName = '1642652_D_FA_VIS' OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 1 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)
			UNION ALL
		--Attribute Feature Testing
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
			WHERE r.RunName = 'NV14408' AND (rt.RoutineName = '1642652_D_FA_VIS' OR rt.RoutineName LIKE '%_IP_%') AND f.FeatureType = 2 AND (dt.TraceabilityListID = 143 OR dt.TraceabilityListID IS NULL)) src
		GROUP BY src.RoutineName, src.ObsID
		HAVING MIN(src.Result) >= 0 )src2
	PIVOT(COUNT(ObsId)
		  FOR Employee IN ([0277], [0423], [???], [0207])
		  ) AS Pvt
	ORDER BY Pvt.RoutineName DESC
		  
		  