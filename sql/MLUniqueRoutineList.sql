SELECT DISTINCT rt.RoutineName
FROM MeasurLink7.dbo.Part p 
LEFT OUTER JOIN MeasurLink7.dbo.Feature f ON f.PartID = p.PartID 
LEFT OUTER JOIN MeasurLink7.dbo.RoutineFeatures rf ON f.FeatureID = rf.FeatureID 
LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON rf.RoutineID = rt.RoutineID 
WHERE p.PartName = ? AND rt.RoutineName IS NOT NULL AND (rt.RoutineName LIKE ? OR rt.RoutineName LIKE '%_IP_%')