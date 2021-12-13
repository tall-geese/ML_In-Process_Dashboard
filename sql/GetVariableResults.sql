SELECT frd.Value 
FROM MeasurLink7.dbo.Run r
INNER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID  = rt.RoutineID 
INNER JOIN MeasurLink7.dbo.FeatureRun fr ON r.RunID = fr.RunID 
INNER JOIN MeasurLink7.dbo.Feature f ON f.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7.dbo.FeatureRunData frd ON fr.RunID  = frd.RunID AND fr.FeatureID = frd.FeatureID 
WHERE r.RunName = ? AND rt.RoutineName = ? AND f.FeatureName = ? AND frd.Value IS NOT NULL;