
SELECT f.FeatureName, f.FeaturePropID 
FROM MeasurLink7.dbo.Run r 
LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID  = rt.RoutineID 
LEFT OUTER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID  = fr.RunID 
LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON fr.FeatureID  = f.FeatureID 
WHERE r.RunName = ? AND rt.RoutineName = ?