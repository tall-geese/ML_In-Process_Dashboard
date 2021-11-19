
--FeaturePrp Investigation
--	How many revisions have a different value from one to the Next? 
-- OR How many Characteristics have revisions where one or more tolerance Limits are NULL?
-- Does it Skew Enough towards one side where we can get aw
SELECT src2.FeatureID,src2.MINFeaturePropID, src2.LowerToleranceLimit, src2.UpperToleranceLimit, src4.FeatureID, src4.MAXFeaturePropID, src4.LowerToleranceLimit, src4.UpperToleranceLimit
FROM (SELECT src.*, fp2.FeatureID, fp2.LowerToleranceLimit, fp2.UpperToleranceLimit 
			FROM(SELECT fp.FeatureID[MINFeatureID], MIN(fp.FeaturePropID)[MINFeaturePropID]
				FROM MeasurLink7.dbo.FeatureProperties fp 
				GROUP BY fp.FeatureID) src
			INNER JOIN MeasurLink7.dbo.FeatureProperties fp2 ON src.MINFeatureID = fp2.FeatureID AND src.MINFeaturePropID = fp2.FeaturePropID)src2
INNER JOIN (SELECT src3.*, fp3.FeatureID, fp3.LowerToleranceLimit, fp3.UpperToleranceLimit 
			FROM(SELECT fp.FeatureID[MAXFeatureID], MAX(fp.FeaturePropID)[MAXFeaturePropID]
				FROM MeasurLink7.dbo.FeatureProperties fp 
				GROUP BY fp.FeatureID) src3
			INNER JOIN MeasurLink7.dbo.FeatureProperties fp3 ON src3.MAXFeatureID = fp3.FeatureID AND src3.MAXFeaturePropID = fp3.FeaturePropID)src4 ON src4.MAXFeatureID = src2.MINFeatureID
WHERE (src2.LowerToleranceLimit <> src4.LowerToleranceLimit OR src2.UpperToleranceLimit <> src4.UpperToleranceLimit OR
		src2.LowerToleranceLimit IS NULL OR src4.UpperToleranceLimit IS NULL)
		AND src2.MINFeaturePropID <> src4.MAXFeaturePropID
		
-- ^^^^^^^^   Plug in the Results from above, we have 6 features, 12 revisions that have discrepencies
		
SELECT pf3.FolderName, pf2.FolderName, pf.FolderName,p.PartName, f.FeatureID, f.FeatureName, fp.FeatureID, fp.FeaturePropID, fp.FeaturePropName, fp.LowerToleranceLimit, fp.UpperToleranceLimit 
FROM MeasurLink7.dbo.Feature f 
LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fp ON f.FeatureID  = fp.FeatureID 
LEFT OUTER JOIN MeasurLink7.dbo.Part p ON f.PartID = p.PartID 
LEFT OUTER JOIN MeasurLink7.dbo.PartFolder pf ON p.PartFolderID = pf.PartFolderID 
LEFT OUTER JOIN MeasurLink7.dbo.PartFolder pf2 ON pf.ParentID = pf2.PartFolderID 
LEFT OUTER JOIN MeasurLink7.dbo.PartFolder pf3 ON pf2.ParentID = pf3.PartFolderID 
WHERE f.FeatureID IN ('3662','4209','4489','4817','4819','4823','4827','4831','5363','5389','5461','5462','5464','5465','5466','5469',
'5474','5475','5499','5503','5542','5547','5636','5657','5658','5959','5960','6002','6013','6014','6015','6136','6182','6183','6222','6223',
'6299','6302','6346','6397','6418','6421','6422','6425','6426','6429','6430','6433','6434','6501','6505','6533','6534','6535','6536','6562',
'6757','6901','6902','6998','7102','7103','7105','7107','7108','7109','7111','7113','7117','7142','7147','7174','7175','7206','7207','7432',
'7575','7736','7764','7769','7773','7939','7942','8337','8338','8339','8340','8341','8407','8408','8409','8418','8474','8477','8483','8484',
'8485','8486','8487','8505','8705','8706','12746','18239','18489','18573','19061','21266')
AND pf.FolderName NOT LIKE '%Inactive%' AND pf2.FolderName NOT LIKE '%Inactive%' AND pf3.FolderName NOT LIKE '%Inactive%'
		
