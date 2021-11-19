
--These are the routines that have been created for the job so far....
SELECT r.RunName,r.RunID, r.RoutineID, rt.RoutineID, rt.RoutineName 
FROM MeasurLink7.dbo.Run r
LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID  = rt.RoutineID 
WHERE r.RunName = '003864-32-1'


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





--Testing GROUP BY method where we can remove an entire observation
SELECT MAX(src.RunName), MAX(src.RoutineName)[Routine], MAX(src.ItemName)[Employee], src.ObsID, MIN(src.Result)
FROM (SELECT r.RunName, rt.RoutineName, f.FeatureID, f.FeatureName, f.FeatureType, frd.ObsID, frd.Value, dt.StartObsID, dt.ItemName, fp.LowerToleranceLimit, fp.UpperToleranceLimit,
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
	WHERE r.RunName = 'NV14408' AND rt.RoutineName = '1642652_D_IP_1XSHIFT' AND f.FeatureType = 1 AND dt.TraceabilityListID = 143)src  --Can either be the Run ID of the RunName and RoutineName
GROUP BY src.ObsID
HAVING MIN(src.Result) >= 0

--Attribute Feature Testing
SELECT MAX(src.RunName), MAX(src.RoutineName)[Routine], MAX(src.ItemName)[Employee], src.ObsID, MIN(src.Result)
FROM (SELECT r.RunName, rt.RoutineName, f.FeatureID, f.FeatureName, f.FeatureType, afrd.ObsID, afrd.DefectCount, dt.StartObsID, dt.ItemName, fp.LowerToleranceLimit, fp.UpperToleranceLimit,
		COALESCE(afrd.DefectCount,1)[Result]
	FROM MeasurLink7.dbo.Run r 
	LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
	LEFT OUTER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON afrd.RunID  = r.RunID AND fr.FeatureID = afrd.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON afrd.FeatureID = f.FeatureID 
	LEFT OUTER JOIN MeasurLink7.dbo.DataTraceability dt ON r.RunID = dt.RunID AND afrd.FeatureID = dt.FeatureID AND afrd.ObsID = dt.StartObsID 
	LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON fp.FeatureID = f.FeatureID AND fr.FeaturePropID = fp.FeaturePropID 
	WHERE r.RunName = 'NV14408' AND rt.RoutineName = '1642652_D_IP_1XSHIFT' AND f.FeatureType = 2 AND dt.TraceabilityListID = 143)src  --Can either be the Run ID of the RunName and RoutineName
GROUP BY src.ObsID
HAVING MIN(src.Result) >= 0
	
