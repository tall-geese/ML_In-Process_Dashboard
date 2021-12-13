SELECT fp.LowerToleranceLimit, fp.UpperToleranceLimit 	
FROM MeasurLink7.dbo.Run r
LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID  = rt.RoutineID 
LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON f.FeatureID = fr.FeatureID 
LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON f.FeatureID  = fp.FeatureID 
WHERE r.RunName = ? AND rt.RoutineName = ? AND f.FeatureName = ?